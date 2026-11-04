<#
.SYNOPSIS
    Installation « à la main » de HardeningKitty (sans Puppet). À lancer en administrateur depuis audit\.

.DESCRIPTION
    Copie le module HardeningKitty et le script d'audit dans InstallDir.
    Par défaut, AUCUNE tâche planifiée n'est créée : utiliser -Schedule pour en ajouter une.
    Prérequis : la release HardeningKitty décompressée dans .\HardeningKitty\ (psd1, psm1, lists\).

.PARAMETER InstallDir
    Racine d'installation (défaut C:\ProgramData\HardeningKitty).

.PARAMETER Schedule
    Crée (ou remplace) la tâche planifiée quotidienne « HardeningKitty audit » en SYSTEM.

.PARAMETER AuditTime
    Heure de l'audit quotidien, format HH:mm (défaut 03:00). Ignoré sans -Schedule.

.PARAMETER FindingList / ListFamily
    Passés tels quels à Invoke-HardeningKittyAudit.ps1 (défaut : auto / cis).

.PARAMETER RunNow
    Lance un premier audit tout de suite, à la fin de l'installation.

.EXAMPLE
    .\run-manual.ps1                                  # installe seulement
    .\run-manual.ps1 -Schedule -AuditTime 02:30       # installe + tâche planifiée à 02:30
    .\run-manual.ps1 -RunNow                          # installe + premier audit immédiat
#>
[CmdletBinding()]
param(
    [string]$InstallDir  = 'C:\ProgramData\HardeningKitty',
    [switch]$Schedule,
    [ValidatePattern('^\d{2}:\d{2}$')]
    [string]$AuditTime   = '03:00',
    [string]$FindingList = 'auto',
    [string]$ListFamily  = 'cis',
    [switch]$RunNow
)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path .\HardeningKitty\HardeningKitty.psm1)) {
    throw "Release HardeningKitty introuvable : décompresse-la dans $PWD\HardeningKitty\ (psd1, psm1, lists\)."
}
if (-not (Test-Path .\Invoke-HardeningKittyAudit.ps1)) { throw "Invoke-HardeningKittyAudit.ps1 introuvable dans $PWD" }

foreach ($d in 'module', 'scripts', 'reports', 'custom_lists') {
    New-Item -ItemType Directory -Path (Join-Path $InstallDir $d) -Force | Out-Null
}
Copy-Item .\HardeningKitty\* (Join-Path $InstallDir 'module') -Recurse -Force
Copy-Item .\Invoke-HardeningKittyAudit.ps1 (Join-Path $InstallDir 'scripts') -Force
Get-ChildItem -LiteralPath $InstallDir -Recurse -File | Unblock-File -ErrorAction SilentlyContinue
Write-Host "Installé dans $InstallDir"

$auditScript = Join-Path $InstallDir 'scripts\Invoke-HardeningKittyAudit.ps1'
$auditArgs   = "-InstallDir `"$InstallDir`" -FindingList `"$FindingList`" -ListFamily `"$ListFamily`""

if ($Schedule) {
    $ps = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $tr = "`"$ps`" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$auditScript`" $auditArgs"
    schtasks /Create /F /TN "HardeningKitty audit" /SC DAILY /ST $AuditTime /RU SYSTEM /RL HIGHEST /TR $tr | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "schtasks a échoué (code $LASTEXITCODE)" }
    Write-Host "Tâche planifiée 'HardeningKitty audit' créée, tous les jours à $AuditTime (SYSTEM)."
    Write-Host "Lancer à la demande : schtasks /Run /TN `"HardeningKitty audit`""
} else {
    Write-Host "Aucune tâche planifiée (ajouter -Schedule [-AuditTime HH:mm] pour en créer une)."
    Write-Host "Audit manuel : powershell -ExecutionPolicy Bypass -File `"$auditScript`" $auditArgs"
}

if ($RunNow) {
    Write-Host "Premier audit en cours (quelques minutes)..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $auditScript -InstallDir $InstallDir -FindingList $FindingList -ListFamily $ListFamily
    Write-Host "Rapport : $(Join-Path $InstallDir 'reports\latest.csv')"
}
