<#
.SYNOPSIS
    Lance un audit HardeningKitty (lecture seule) et produit un rapport CSV horodaté.

.DESCRIPTION
    - Mode Audit uniquement : ne modifie JAMAIS la configuration de la machine.
    - Écrit reports\hk_<HOST>_<date>.csv et recopie le dernier résultat dans reports\latest.csv
      (c'est ce fichier que lit check_hardeningkitty.ps1).
    - Garde les N derniers rapports, supprime les plus anciens.
    - Prévu pour tourner en SYSTEM via une tâche planifiée (voir la classe Puppet).

.PARAMETER InstallDir
    Racine de l'installation. Attendu : <InstallDir>\module\HardeningKitty.psm1 et <InstallDir>\module\lists\

.PARAMETER FindingList
    Nom du CSV de la finding list à utiliser (dans module\lists), ou chemin complet.

.PARAMETER KeepReports
    Nombre de rapports horodatés conservés.
#>
[CmdletBinding()]
param(
    [string]$InstallDir  = 'C:\ProgramData\HardeningKitty',
    [string]$FindingList = 'auto',
    [string]$ListFamily  = 'cis',
    [string]$FallbackList = 'finding_list_0x6d69636b_machine.csv',
    [int]$KeepReports    = 14
)

$ErrorActionPreference = 'Stop'

$reportDir  = Join-Path $InstallDir 'reports'

# Module HardeningKitty : dans InstallDir\module (installé), sinon à côté de ce script (.\HardeningKitty\, toolkit non installé)
$moduleDir = Join-Path $InstallDir 'module'
if (-not (Test-Path -LiteralPath (Join-Path $moduleDir 'HardeningKitty.psm1'))) {
    $local = Join-Path $PSScriptRoot 'HardeningKitty'
    if (Test-Path -LiteralPath (Join-Path $local 'HardeningKitty.psm1')) {
        $moduleDir = $local
        Write-Host "HardeningKitty non installé dans $InstallDir : utilisation de $moduleDir (rapports dans $reportDir)"
    } else {
        throw "HardeningKitty introuvable ni dans $moduleDir ni dans $local. Lance d'abord run-manual.ps1 (ou passe -InstallDir)."
    }
}
$modulePath = Join-Path $moduleDir 'HardeningKitty.psm1'
$listsDir   = Join-Path $moduleDir 'lists'

# --- choix de la finding list ------------------------------------------------
# 'auto' : détecte l'OS (client / serveur + version) et cherche dans lists\ une liste *machine*
# de la famille demandée (cis, msft_security_baseline, ...). Sinon : nom de fichier ou chemin complet.
# Exemples avec les listes HardeningKitty 0.9.4 :
#   Server 2022 (build 20348), cis  -> finding_list_cis_microsoft_windows_server_2022_22h2_4.0.0_machine.csv
#   Server 2025 (26100), msft, DC   -> finding_list_msft_security_baseline_windows_server_2025_2602_dc_machine.csv
#   Windows 11 (>= 22000), cis      -> finding_list_cis_microsoft_windows_11_enterprise_24h2_machine.csv
#   Windows 10, msft                -> finding_list_msft_security_baseline_windows_10_22h2_machine.csv
function Resolve-AutoList([string]$Dir, [string]$Family, [string]$Fallback) {
    $os    = Get-CimInstance -ClassName Win32_OperatingSystem
    $build = [int]$os.BuildNumber
    if ($os.ProductType -eq 1) {
        if ($build -ge 22000) { $token = 'windows_11' } else { $token = 'windows_10' }
    } else {
        switch ($build) {
            14393   { $token = 'server_2016' }
            17763   { $token = 'server_2019' }
            20348   { $token = 'server_2022' }
            default { if ($build -ge 26100) { $token = 'server_2025' } else { $token = '' } }
        }
    }

    if ($token) {
        $pattern = "finding_list_${Family}_*${token}*_machine.csv"
        $hits = @(Get-ChildItem -LiteralPath $Dir -Filter $pattern -File -ErrorAction SilentlyContinue)
        # Baselines Microsoft serveur : variantes _dc_ / _member_ selon le rôle (ProductType 2 = contrôleur de domaine)
        if ($hits | Where-Object { $_.Name -match '_(dc|member)_' }) {
            if ($os.ProductType -eq 2) { $hits = @($hits | Where-Object { $_.Name -match '_dc_' }) }
            else                       { $hits = @($hits | Where-Object { $_.Name -notmatch '_dc_' }) }
        }
        # plusieurs versions (21h2_1.0.0, 22h2_4.0.0...) : la dernière par ordre alphabétique
        $hit = $hits | Sort-Object Name | Select-Object -Last 1
        if ($hit) { return $hit.FullName }
        Write-Warning "Aucune liste '$pattern' (build $build) : repli sur $Fallback"
    } else {
        Write-Warning "OS non reconnu (build $build, ProductType $($os.ProductType)) : repli sur $Fallback"
    }
    return (Join-Path $Dir $Fallback)
}

if ($FindingList -eq 'auto') {
    $listPath = Resolve-AutoList -Dir $listsDir -Family $ListFamily -Fallback $FallbackList
} elseif ([System.IO.Path]::IsPathRooted($FindingList)) {
    $listPath = $FindingList
} else {
    # nom simple : d'abord custom_lists\ (listes maison), puis module\lists\ (listes fournies)
    $customPath = Join-Path (Join-Path $InstallDir 'custom_lists') $FindingList
    if (Test-Path -LiteralPath $customPath) {
        $listPath = $customPath
    } else {
        $listPath = Join-Path $listsDir $FindingList
    }
}

if (-not (Test-Path -LiteralPath $listPath))   { throw "Finding list introuvable : $listPath" }
New-Item -ItemType Directory -Path $reportDir -Force | Out-Null

$stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$report = Join-Path $reportDir ("hk_{0}_{1}.csv" -f $env:COMPUTERNAME, $stamp)

Import-Module $modulePath -Force
Invoke-HardeningKitty -Mode Audit -FileFindingList $listPath -Report -ReportFile $report -SkipMachineInformation | Out-Null

if (-not (Test-Path -LiteralPath $report)) { throw "Aucun rapport produit : $report" }

# latest.csv n'est remplacé qu'une fois le nouveau rapport complet
Copy-Item -LiteralPath $report -Destination (Join-Path $reportDir 'latest.csv') -Force
# trace de la liste réellement utilisée (utile pour interpréter le score)
Set-Content -LiteralPath (Join-Path $reportDir 'latest.list.txt') -Value (Split-Path $listPath -Leaf) -Encoding ASCII

# rotation
Get-ChildItem -LiteralPath $reportDir -Filter 'hk_*.csv' |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip $KeepReports |
    Remove-Item -Force

Write-Output "Audit termine : $report"
