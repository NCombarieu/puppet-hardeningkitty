<#
.SYNOPSIS
    Check Nagios/Shinken : lit le dernier rapport HardeningKitty et renvoie un état + perfdata.

.DESCRIPTION
    Ne lance PAS HardeningKitty (trop long pour un timeout NSClient++) : il lit
    reports\latest.csv produit par Invoke-HardeningKittyAudit.ps1.

    Codes retour : 0 OK, 1 WARNING, 2 CRITICAL, 3 UNKNOWN (rapport absent, trop vieux ou illisible).

.PARAMETER Warn
    Score de conformité (%) en dessous duquel on passe WARNING.

.PARAMETER Crit
    Score de conformité (%) en dessous duquel on passe CRITICAL.

.PARAMETER CritHigh
    Nombre d'échecs de sévérité High à partir duquel on passe CRITICAL (-1 = désactivé).

.PARAMETER MaxAgeHours
    Âge maximum du rapport avant UNKNOWN.

.EXAMPLE
    .\check_hardeningkitty.ps1 -Warn 85 -Crit 70 -CritHigh 10
#>
[CmdletBinding()]
param(
    [string]$ReportPath  = 'C:\ProgramData\HardeningKitty\reports\latest.csv',
    [double]$Warn        = 80,
    [double]$Crit        = 60,
    [int]$CritHigh       = -1,
    [int]$MaxAgeHours    = 48
)

$ci = [System.Globalization.CultureInfo]::InvariantCulture

function Exit-Nagios([int]$code, [string]$text) {
    Write-Output $text
    exit $code
}

try {
    if (-not (Test-Path -LiteralPath $ReportPath)) {
        Exit-Nagios 3 "UNKNOWN - rapport HardeningKitty introuvable : $ReportPath"
    }

    $file  = Get-Item -LiteralPath $ReportPath
    $age   = (Get-Date) - $file.LastWriteTime
    if ($age.TotalHours -gt $MaxAgeHours) {
        Exit-Nagios 3 ("UNKNOWN - rapport HardeningKitty trop ancien ({0:N0} h > {1} h)" -f $age.TotalHours, $MaxAgeHours)
    }

    # séparateur , ou ;
    $first = Get-Content -LiteralPath $ReportPath -TotalCount 1
    $delim = ','
    if ($first -and ($first.Split(';').Count -gt $first.Split(',').Count)) { $delim = ';' }
    $rows = @(Import-Csv -LiteralPath $ReportPath -Delimiter $delim)
    if ($rows.Count -eq 0) { Exit-Nagios 3 'UNKNOWN - rapport HardeningKitty vide' }

    $cols = $rows[0].PSObject.Properties.Name
    if ($cols -notcontains 'TestResult') {
        Exit-Nagios 3 "UNKNOWN - colonne 'TestResult' absente du rapport (colonnes : $($cols -join ','))"
    }

    $passed = 0; $failed = 0; $high = 0; $medium = 0; $low = 0
    $topHigh = New-Object System.Collections.Generic.List[string]

    foreach ($r in $rows) {
        $t = ([string]$r.TestResult).Trim()
        if ($t -match '^(passed|pass|ok)$') { $passed++; continue }

        $isFail = $false
        if ($t -match '^(failed|fail|low|medium|high|critical)$') { $isFail = $true }
        if (-not $isFail) { continue }
        $failed++

        # sévérité du constat : SeverityFinding > TestResult > Severity
        $sev = ''
        if ($cols -contains 'SeverityFinding') { $sev = ([string]$r.SeverityFinding).Trim() }
        if ($sev -eq '' -or $sev -match '^(none|passed)$') {
            if ($t -match '^(low|medium|high|critical)$') { $sev = $t }
            elseif ($cols -contains 'Severity') { $sev = ([string]$r.Severity).Trim() }
        }

        switch -Regex ($sev.ToLower()) {
            '^(high|critical)$' {
                $high++
                if ($topHigh.Count -lt 5) { $topHigh.Add(("{0} {1}" -f $r.ID, $r.Name).Trim()) }
            }
            '^medium$' { $medium++ }
            '^low$'    { $low++ }
        }
    }

    $evaluated = $passed + $failed
    if ($evaluated -eq 0) { Exit-Nagios 3 'UNKNOWN - aucun test evalue dans le rapport' }

    $score = [math]::Round(100.0 * $passed / $evaluated, 1)

    $code = 0; $label = 'OK'
    if ($score -lt $Warn) { $code = 1; $label = 'WARNING' }
    if ($score -lt $Crit) { $code = 2; $label = 'CRITICAL' }
    if ($CritHigh -ge 0 -and $high -ge $CritHigh) { $code = 2; $label = 'CRITICAL' }

    $text = "{0} - HardeningKitty conformite {1}% ({2} reussis, {3} echoues dont {4} High, {5} Medium, {6} Low)" -f `
        $label, $score.ToString('0.0', $ci), $passed, $failed, $high, $medium, $low

    $perf = "score={0}%;{1};{2};0;100 passed={3} failed={4} high={5} medium={6} low={7} age_h={8}" -f `
        $score.ToString('0.0', $ci), $Warn.ToString($ci), $Crit.ToString($ci), $passed, $failed, $high, $medium, $low, [math]::Round($age.TotalHours)

    $out = "$text | $perf"
    if ($topHigh.Count -gt 0) {
        $out += "`nPrincipaux echecs High :`n" + (($topHigh | ForEach-Object { " - $_" }) -join "`n")
    }

    Exit-Nagios $code $out
}
catch {
    Exit-Nagios 3 ("UNKNOWN - erreur check_hardeningkitty : {0}" -f $_.Exception.Message)
}
