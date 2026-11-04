<#
.SYNOPSIS
    Transforme un ou plusieurs rapports CSV HardeningKitty en pages HTML autonomes.

.DESCRIPTION
    - Accepte un fichier CSV, un dossier, un joker (.\hardeningkitty_report_*.csv) ou plusieurs
      chemins séparés par des virgules.
    - UNE seule page HTML (CSS + JS inclus, aucune dépendance, fonctionne hors ligne) qui embarque
      un ou plusieurs rapports. Avec un seul rapport : vue détaillée. Avec plusieurs : sélecteur de
      machine, tableau du parc, vue « Par test » (quels tests échouent sur combien de machines).
    - Barre de triage cliquable, filtres combinables, recherche, tri, lignes dépliables, export CSV
      des lignes filtrées, copie des IDs en échec, état des filtres dans l'URL, thème clair/sombre,
      raccourcis clavier (? pour l'aide).
    - Le nom de machine est déduit du nom de fichier (hardeningkitty_report_<HOST>_<date>.csv,
      hk_<HOST>_<date>.csv) ou, pour latest.csv, du dossier parent.
    - Compatible Windows PowerShell 5.1 et PowerShell 7+.

.PARAMETER Path
    Fichier CSV, dossier, joker, ou liste de chemins séparés par des virgules.

.PARAMETER OutFile
    Fichier HTML de sortie. Par défaut : <rapport>.html pour un seul rapport,
    hardeningkitty_parc.html à côté du premier CSV pour plusieurs.

.PARAMETER Recurse
    Avec un dossier : cherche aussi les CSV dans les sous-dossiers (ex. un dossier par machine
    contenant latest.csv).

.PARAMETER Title
    Titre affiché en haut de page. Par défaut : "Rapport HardeningKitty".

.PARAMETER Open
    Ouvre la page générée dans le navigateur par défaut.

.EXAMPLE
    .\ConvertTo-HardeningKittyHtml.ps1 -Path .\hardeningkitty_report_*.csv -Open

.EXAMPLE
    .\ConvertTo-HardeningKittyHtml.ps1 -Path \\partage\hk -Recurse -OutFile C:\inetpub\hk\parc.html -Title "Parc CNCV"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [string]$OutFile,

    [switch]$Recurse,

    [string]$Title = 'Rapport HardeningKitty',

    [switch]$Open
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- utilitaires

function Esc([object]$s) {
    if ($null -eq $s) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$s)
}

function Get-Col($row, [string[]]$names) {
    foreach ($n in $names) {
        $p = $row.PSObject.Properties | Where-Object { $_.Name -ieq $n } | Select-Object -First 1
        if ($p) { return [string]$p.Value }
    }
    return ''
}

function Import-HkCsv([string]$file) {
    # détecte le séparateur (, ou ;) sur la première ligne
    $first = Get-Content -LiteralPath $file -TotalCount 1 -Encoding UTF8
    $delim = ','
    if ($first -and ($first.Split(';').Count -gt $first.Split(',').Count)) { $delim = ';' }
    return @(Import-Csv -LiteralPath $file -Delimiter $delim -Encoding UTF8)
}

function Get-Status([string]$testResult) {
    $t = $testResult.Trim()
    if ($t -match '^(passed|pass|ok)$')           { return 'Passed' }
    if ($t -match '^(failed|fail)$')              { return 'Failed' }
    if ($t -match '^(low|medium|high|critical)$') { return 'Failed' }
    return 'Info'
}

function Get-SevRank([string]$s) {
    switch -Regex ($s.Trim().ToLower()) {
        '^critical$' { return 4 }
        '^high$'     { return 3 }
        '^medium$'   { return 2 }
        '^low$'      { return 1 }
        default      { return 0 }
    }
}

function Get-SevLabel([string]$s) {
    $s = $s.Trim()
    if ($s -eq '' -or $s -match '^(none|passed)$') { return 'None' }
    return (Get-Culture).TextInfo.ToTitleCase($s.ToLower())
}

# ---------------------------------------------------------------- modèle

function ConvertTo-HkModel($rows) {
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rows) {
        $testResult = Get-Col $r @('TestResult', 'Test Result')
        $status     = Get-Status $testResult
        $baseSev    = Get-SevLabel (Get-Col $r @('Severity'))
        $findingSev = Get-SevLabel (Get-Col $r @('SeverityFinding', 'Severity Finding'))

        # sévérité du constat : SeverityFinding > TestResult (Low/Medium/High) > Severity du test
        $sev = 'None'
        if ($status -eq 'Failed') {
            if ($findingSev -ne 'None') { $sev = $findingSev }
            elseif ((Get-SevRank $testResult) -gt 0) { $sev = Get-SevLabel $testResult }
            else { $sev = $baseSev }
        }

        $items.Add([pscustomobject]@{
            id   = Get-Col $r @('ID')
            cat  = Get-Col $r @('Category')
            name = Get-Col $r @('Name')
            meth = Get-Col $r @('Method')
            def  = Get-Col $r @('DefaultValue', 'Default Value')
            res  = Get-Col $r @('Result')
            rec  = Get-Col $r @('Recommended')
            st   = $status
            sev  = $sev
            rank = Get-SevRank $sev
        })
    }
    return $items
}

function Get-Score($items) {
    $p = @($items | Where-Object { $_.st -eq 'Passed' }).Count
    $f = @($items | Where-Object { $_.st -eq 'Failed' }).Count
    if (($p + $f) -eq 0) { return 0 }
    return [math]::Round(100 * $p / ($p + $f), 1)
}

# ---------------------------------------------------------------- HTML

$Css = @'
:root{
  --bg:#eef1f5;--panel:#ffffff;--ink:#141c28;--ink-2:#4a5565;--ink-3:#7c8797;--line:#d7dde6;--line-2:#e9edf2;
  --accent:#2f5d9e;--accent-ink:#ffffff;--focus:#2f5d9e;
  --pass:#2a7f55;--low:#3270c7;--med:#c47a12;--high:#b8322a;--crit:#6e1a14;--info:#8a94a3;
  --pass-bg:#e3f2ea;--low-bg:#e4edfb;--med-bg:#fbefdc;--high-bg:#f9e3e1;--crit-bg:#f1dada;--info-bg:#eceff3;
  --mono:"Cascadia Mono","Cascadia Code",Consolas,"Courier New",monospace;
  --sans:"Segoe UI Variable Text","Segoe UI",system-ui,-apple-system,Roboto,sans-serif;
  --r:6px;--shadow:0 1px 2px rgba(20,28,40,.06);
}
:root[data-theme="dark"]{
  --bg:#0f141c;--panel:#161d28;--ink:#e6ebf2;--ink-2:#aab4c3;--ink-3:#76818f;--line:#2a3442;--line-2:#1f2733;
  --accent:#6b9be0;--accent-ink:#0f141c;--focus:#8db4ee;
  --pass:#4fbf85;--low:#6aa0ea;--med:#e8a43c;--high:#ea6a5f;--crit:#c9574c;--info:#7c8797;
  --pass-bg:#15301f;--low-bg:#172a45;--med-bg:#3a2a10;--high-bg:#3e1c19;--crit-bg:#331413;--info-bg:#202834;
  --shadow:none;
}
@media(prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --bg:#0f141c;--panel:#161d28;--ink:#e6ebf2;--ink-2:#aab4c3;--ink-3:#76818f;--line:#2a3442;--line-2:#1f2733;
  --accent:#6b9be0;--accent-ink:#0f141c;--focus:#8db4ee;
  --pass:#4fbf85;--low:#6aa0ea;--med:#e8a43c;--high:#ea6a5f;--crit:#c9574c;--info:#7c8797;
  --pass-bg:#15301f;--low-bg:#172a45;--med-bg:#3a2a10;--high-bg:#3e1c19;--crit-bg:#331413;--info-bg:#202834;
  --shadow:none;
}}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
@media(prefers-reduced-motion:reduce){html{scroll-behavior:auto}*{transition:none!important;animation:none!important}}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 var(--sans)}
button,input,select{font:inherit;color:inherit}
button{cursor:pointer}
:focus-visible{outline:2px solid var(--focus);outline-offset:2px}
a{color:var(--accent)}
.wrap{max-width:1440px;margin:0 auto;padding:0 24px}
.btn{border:1px solid var(--line);background:var(--panel);border-radius:var(--r);padding:6px 11px;font-size:13px;color:var(--ink-2);display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.btn:hover{border-color:var(--ink-3);color:var(--ink)}
.btn.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}
.btn.on{border-color:var(--ink);color:var(--ink);background:var(--line-2)}
.btn svg{width:14px;height:14px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
select.sel{border:1px solid var(--line);background:var(--panel);border-radius:var(--r);padding:6px 30px 6px 10px;font-size:13px;appearance:none;-webkit-appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%237c8797' stroke-width='2'%3E%3Cpath d='m6 9 6 6 6-6'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 8px center;background-size:14px;max-width:320px}
kbd{font:11.5px var(--mono);border:1px solid var(--line);border-bottom-width:2px;border-radius:4px;padding:0 5px;background:var(--panel);color:var(--ink-2)}

/* ---- en-tête ---- */
.top{border-bottom:1px solid var(--line);background:var(--panel)}
.top .wrap{display:flex;align-items:center;gap:14px;flex-wrap:wrap;padding-top:12px;padding-bottom:12px}
.top h1{margin:0;font-size:18px;font-weight:650;letter-spacing:-.01em;white-space:nowrap}
.top .meta{color:var(--ink-3);font-size:12.5px;flex:1 1 220px;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.top .meta b{color:var(--ink-2);font-weight:500}
.top .right{margin-left:auto;display:flex;gap:6px;align-items:center}
.ibtn{border:1px solid var(--line);background:var(--panel);border-radius:var(--r);width:32px;height:32px;display:inline-flex;align-items:center;justify-content:center;color:var(--ink-2);padding:0}
.ibtn:hover{border-color:var(--ink-3);color:var(--ink)}
.ibtn.on{background:var(--line-2);border-color:var(--ink);color:var(--ink)}
.ibtn svg{width:16px;height:16px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.scope{display:flex;align-items:center;gap:8px}
.scope label{font-size:13px;color:var(--ink-3)}

/* ---- triage (hero) ---- */
.triage{padding:18px 0 4px}
.stats{display:flex;align-items:stretch;gap:8px;flex-wrap:wrap;margin-bottom:10px}
.stat{border:1px solid var(--line);background:var(--panel);border-radius:var(--r);padding:8px 12px;text-align:left;display:flex;flex-direction:column;gap:1px;min-width:104px;border-top:3px solid var(--sc,var(--line))}
.stat:hover{border-color:var(--ink-3);border-top-color:var(--sc,var(--ink-3))}
.stat.on{background:var(--line-2);border-color:var(--ink);border-top-color:var(--sc,var(--ink))}
.stat .l{font-size:12px;color:var(--ink-2);display:flex;align-items:center;gap:6px}
.stat .l i{width:8px;height:8px;border-radius:2px;background:var(--sc)}
.stat .v{font-size:22px;font-weight:650;letter-spacing:-.02em;line-height:1.15;font-variant-numeric:tabular-nums}
.stat .p{font-size:11.5px;color:var(--ink-3);font-variant-numeric:tabular-nums}
.stat.score{min-width:150px;border-top-color:var(--ink);cursor:default}
.stat.score .v{font-size:30px}
.stat.score .l abbr{text-decoration:underline dotted;text-underline-offset:2px;cursor:help}
.lead{margin:0 0 10px;color:var(--ink-2);max-width:90ch;font-size:13.5px}
.bar{display:flex;height:26px;border-radius:var(--r);overflow:hidden;border:1px solid var(--line);background:var(--panel)}
.bar button{border:0;padding:0;height:100%;position:relative;transition:filter .15s;min-width:8px}
.bar button:hover{filter:brightness(1.08)}
.bar button.on{box-shadow:inset 0 0 0 3px var(--panel)}
.bar button span{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-weight:600;font-size:13px;color:#fff;text-shadow:0 1px 1px rgba(0,0,0,.25);white-space:nowrap;overflow:hidden}
.bar .s-Passed{background:var(--pass)}.bar .s-Low{background:var(--low)}.bar .s-Medium{background:var(--med)}
.bar .s-High{background:var(--high)}.bar .s-Critical{background:var(--crit)}.bar .s-Info{background:var(--info)}

/* ---- parc : tableau des machines ---- */
.fleet{margin-top:16px;border:1px solid var(--line);border-radius:var(--r);background:var(--panel);box-shadow:var(--shadow)}
.fleet h2{margin:0;padding:11px 14px;font-size:13.5px;font-weight:600;border-bottom:1px solid var(--line-2);display:flex;align-items:center;gap:10px}
.fleet h2 span{color:var(--ink-3);font-weight:400}
.fleet table{border:0;box-shadow:none;margin:0;border-radius:0}
.fleet thead th{position:static;cursor:pointer}
.fleet tbody tr{cursor:pointer}
.fleet tbody tr:hover td{background:var(--line-2)}
.fleet tbody tr.on td{background:var(--line-2);font-weight:600}
.fleet td.num{font-variant-numeric:tabular-nums;text-align:right;white-space:nowrap}
.fleet td.h{color:var(--high);font-weight:600}
.fleet .sbar{display:flex;height:8px;border-radius:4px;overflow:hidden;background:var(--line-2);min-width:140px}
.fleet .sbar i{display:block;height:100%}
.fleet .host{font-weight:600}
.fleet .host small{display:block;font-weight:400;color:var(--ink-3);font-size:12px}

/* ---- corps : latéral + tableau ---- */
.body{display:grid;grid-template-columns:250px minmax(0,1fr);gap:20px;align-items:start;padding:14px 0 60px}
.side{position:sticky;top:12px;max-height:calc(100vh - 24px);overflow:auto;border:1px solid var(--line);border-radius:var(--r);background:var(--panel);box-shadow:var(--shadow)}
.side h2{margin:0;padding:11px 14px;font-size:13px;font-weight:600;color:var(--ink-2);border-bottom:1px solid var(--line-2);display:flex;justify-content:space-between;position:sticky;top:0;background:var(--panel);z-index:1}
.side h2 button{border:0;background:none;color:var(--accent);font-size:12.5px;padding:0}
.cat{display:grid;width:100%;text-align:left;border:0;background:none;padding:8px 14px;border-left:3px solid transparent;grid-template-columns:1fr auto;gap:3px 10px;align-items:center}
.cat:hover{background:var(--line-2)}
.cat.on{border-left-color:var(--accent);background:var(--line-2)}
.cat .n{font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.cat .c{font-size:12px;color:var(--ink-3);font-variant-numeric:tabular-nums}
.cat .c b{color:var(--high);font-weight:600}
.cat .mini{grid-column:1/-1;height:4px;border-radius:2px;display:flex;overflow:hidden;background:var(--line-2)}
.cat .mini i{display:block;height:100%}
.main{min-width:0}

/* ---- barre d'outils collante ---- */
.tools{position:sticky;top:0;z-index:5;background:var(--bg);padding:10px 0;margin-bottom:-1px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;border-bottom:1px solid var(--line)}
.search{flex:1 1 240px;position:relative}
.search input{width:100%;padding:7px 32px 7px 32px;border:1px solid var(--line);border-radius:var(--r);background:var(--panel)}
.search svg{position:absolute;left:10px;top:9px;width:15px;height:15px;stroke:var(--ink-3);fill:none;stroke-width:1.8;pointer-events:none}
.search .x{position:absolute;right:6px;top:5px;border:0;background:none;color:var(--ink-3);padding:2px 6px;border-radius:4px;display:none}
.search .x.show{display:block}
.tools .btn.toggle.on{background:var(--high-bg);border-color:var(--high);color:var(--high)}
.tools .grp{display:inline-flex;gap:4px}
.seg{display:inline-flex;border:1px solid var(--line);border-radius:var(--r);overflow:hidden}
.seg button{border:0;background:var(--panel);padding:6px 11px;font-size:13px;color:var(--ink-2)}
.seg button+button{border-left:1px solid var(--line)}
.seg button.on{background:var(--ink);color:var(--bg)}
.count{color:var(--ink-3);font-size:13px;margin-left:auto;font-variant-numeric:tabular-nums;white-space:nowrap}
.active{display:flex;gap:6px;flex-wrap:wrap;align-items:center;padding:8px 0 0;font-size:13px;color:var(--ink-2)}
.active:empty{display:none}
.tag{border:1px solid var(--line);background:var(--panel);border-radius:4px;padding:2px 8px;display:inline-flex;gap:6px;align-items:center}
.tag button{border:0;background:none;color:var(--ink-3);padding:0 2px;line-height:1}
.tag button:hover{color:var(--high)}

/* ---- tableau ---- */
table{width:100%;border-collapse:separate;border-spacing:0;background:var(--panel);border:1px solid var(--line);border-radius:var(--r);box-shadow:var(--shadow);margin-top:12px}
thead th{position:sticky;top:var(--tools-h,57px);z-index:4;background:var(--panel);box-shadow:0 1px 0 var(--line);text-align:left;font-size:12.5px;font-weight:600;color:var(--ink-2);padding:9px 10px;border-bottom:1px solid var(--line);cursor:pointer;user-select:none;white-space:nowrap}
thead th:hover{color:var(--ink)}
thead th .d{opacity:.35;font-size:10px;margin-left:4px}
thead th.asc .d,thead th.desc .d{opacity:1}
tbody td{padding:8px 10px;border-bottom:1px solid var(--line-2);vertical-align:top}
tbody tr:last-child td{border-bottom:0}
tbody tr.r{cursor:pointer;scroll-margin-top:calc(var(--tools-h,57px) + 44px)}
tbody tr.r:hover td{background:var(--line-2)}
tbody tr.r.open td{background:var(--line-2);border-bottom-color:transparent}
tbody tr.r td:first-child{border-left:3px solid transparent}
tr.r.s-High td:first-child{border-left-color:var(--high)}
tr.r.s-Critical td:first-child{border-left-color:var(--crit)}
tr.r.s-Medium td:first-child{border-left-color:var(--med)}
tr.r.s-Low td:first-child{border-left-color:var(--low)}
tr.r.s-Passed td:first-child{border-left-color:var(--pass)}
td.id{font-family:var(--mono);font-size:12.5px;color:var(--ink-2);white-space:nowrap}
td.nm{max-width:520px}
td.nm .c{display:block;font-size:12px;color:var(--ink-3)}
td.hn{white-space:nowrap;color:var(--ink-2)}
td.v{font-family:var(--mono);font-size:12px;max-width:240px;word-break:break-word;color:var(--ink-2)}
td.v.bad{color:var(--high)}
td.num{font-variant-numeric:tabular-nums;white-space:nowrap}
.pill{display:inline-flex;align-items:center;gap:6px;border-radius:999px;padding:2px 9px;font-size:12px;font-weight:600;white-space:nowrap}
.pill i{width:7px;height:7px;border-radius:50%;display:inline-block;background:currentColor}
.pill.Passed{color:var(--pass);background:var(--pass-bg)}.pill.Low{color:var(--low);background:var(--low-bg)}
.pill.Medium{color:var(--med);background:var(--med-bg)}.pill.High{color:var(--high);background:var(--high-bg)}
.pill.Critical{color:var(--crit);background:var(--crit-bg)}.pill.Info{color:var(--info);background:var(--info-bg)}
.pill.None{color:var(--ink-3);background:var(--info-bg)}
.hbar{display:inline-flex;height:8px;width:120px;border-radius:4px;overflow:hidden;background:var(--line-2);vertical-align:middle;margin-right:8px}
.hbar i{display:block;height:100%;background:var(--high)}
tr.x td{padding:0 10px 14px 23px;background:var(--line-2);border-bottom:1px solid var(--line)}
.det{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px 24px;padding-top:4px}
.det .k{font-size:12px;color:var(--ink-3);margin-bottom:2px}
.det .vv{font-family:var(--mono);font-size:12.5px;white-space:pre-wrap;word-break:break-word;background:var(--panel);border:1px solid var(--line);border-radius:4px;padding:6px 8px}
.det .vv.bad{border-color:var(--high);color:var(--high)}
.det .vv.good{border-color:var(--pass)}
.det .full{grid-column:1/-1}
.det .acts{grid-column:1/-1;display:flex;gap:8px;flex-wrap:wrap}
.hosts{display:flex;flex-wrap:wrap;gap:6px}
.hosts button{border:1px solid var(--line);background:var(--panel);border-radius:4px;padding:3px 9px;font-size:12.5px;font-family:var(--mono)}
.hosts button.ok{border-color:var(--pass);color:var(--pass)}
.hosts button.ko{border-color:var(--high);color:var(--high)}
.hosts button:hover{filter:brightness(.95)}
.empty{padding:48px 20px;text-align:center;color:var(--ink-3);background:var(--panel);border:1px dashed var(--line);border-radius:var(--r);margin-top:12px}
.empty b{color:var(--ink);display:block;font-weight:600;margin-bottom:4px}

/* ---- toast, aide, haut de page ---- */
.toast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%) translateY(10px);background:var(--ink);color:var(--bg);padding:9px 14px;border-radius:var(--r);font-size:13px;opacity:0;pointer-events:none;transition:opacity .15s,transform .15s;z-index:20}
.toast.show{opacity:1;transform:translateX(-50%)}
.up{position:fixed;right:20px;bottom:20px;opacity:0;pointer-events:none;transition:opacity .15s}
.up.show{opacity:1;pointer-events:auto}
dialog{border:1px solid var(--line);border-radius:8px;background:var(--panel);color:var(--ink);padding:20px 24px;max-width:480px}
dialog p{margin:0 0 12px;color:var(--ink-2);font-size:13px}
dialog::backdrop{background:rgba(10,14,20,.5)}
dialog h3{margin:0 0 12px;font-size:15px}
dialog table{box-shadow:none;margin:0;border:0}
dialog td{padding:5px 8px 5px 0;border:0}
dialog .close{margin-top:14px}

@media(max-width:900px){
  .body{grid-template-columns:1fr}
  .side{position:static;max-height:260px}
  thead th{position:static}
  td.nm{max-width:none}
  .top .right{margin-left:0}
  .stat{min-width:0;flex:1 1 90px}
  .hide-m{display:none}
}
@media(max-width:600px){.wrap{padding:0 16px}}
@media print{
  .tools,.side,.top .right,.up,.toast,.scope{display:none!important}
  .body{grid-template-columns:1fr}
  body{background:#fff;color:#000}
  thead th{position:static}
  tr.x{display:none}
  tbody tr.r{page-break-inside:avoid}
  table,.fleet,.side{box-shadow:none}
}
'@

$Js = @'
(function(){
"use strict";
var D=JSON.parse(document.getElementById('data').textContent);
var HOSTS=D.hosts, ROWS=[], META=D.meta;
var MULTI=HOSTS.length>1;
var SEV_ORDER=['Critical','High','Medium','Low','Passed','Info'];
var SEV_LABEL={Critical:'Critique',High:'High',Medium:'Medium',Low:'Low',Passed:'Réussis',Info:'Non évalués'};
var SEV_VAR={Passed:'pass',Low:'low',Medium:'med',High:'high',Critical:'crit',Info:'info'};

/* --- données à plat --- */
HOSTS.forEach(function(h,hi){
  h.i=hi; h.passed=0; h.failed=0; h.high=0; h.med=0; h.low=0; h.info=0;
  h.rows.forEach(function(r){
    r.h=hi; r._k=r.st==='Passed'?'Passed':(r.st==='Info'?'Info':r.sev); r._i=ROWS.length;
    r._t=(h.name+' '+r.id+' '+r.cat+' '+r.name+' '+r.res+' '+r.rec+' '+r.meth).toLowerCase();
    if(r.st==='Passed')h.passed++; else if(r.st==='Failed'){h.failed++; if(r.rank>=3)h.high++; else if(r.rank===2)h.med++; else h.low++;} else h.info++;
    ROWS.push(r);
  });
  h.score=(h.passed+h.failed)?Math.round(1000*h.passed/(h.passed+h.failed))/10:0;
});

/* --- état --- */
var S={host:-1,q:'',sev:{},cat:'',sort:'sev',dir:'desc',open:{},mode:'rows',fsort:'score',fdir:'asc'};
if(!MULTI) S.host=0;

function readHash(){
  try{
    var h=location.hash.replace(/^#/,''); if(!h) return;
    var p=new URLSearchParams(h);
    if(MULTI&&p.has('h')){ var hn=p.get('h'); var f=HOSTS.filter(function(x){return x.name===hn;})[0]; S.host=f?f.i:-1; }
    S.q=p.get('q')||''; S.cat=p.get('cat')||''; S.sort=p.get('s')||'sev'; S.dir=p.get('d')||'desc';
    S.mode=p.get('m')==='tests'?'tests':'rows';
    S.sev={}; (p.get('sev')||'').split(',').filter(Boolean).forEach(function(k){S.sev[k]=1;});
  }catch(e){}
}
function writeHash(){
  var p=new URLSearchParams();
  if(MULTI&&S.host>=0) p.set('h',HOSTS[S.host].name);
  if(S.q) p.set('q',S.q); if(S.cat) p.set('cat',S.cat);
  var sv=Object.keys(S.sev); if(sv.length) p.set('sev',sv.join(','));
  if(S.sort!=='sev') p.set('s',S.sort); if(S.dir!=='desc') p.set('d',S.dir);
  if(S.mode==='tests') p.set('m','tests');
  var s=p.toString();
  try{ history.replaceState(null,'',s?'#'+s:location.pathname+location.search); }catch(e){}
}

/* --- helpers --- */
function $(s,el){ return (el||document).querySelector(s); }
function $$(s,el){ return [].slice.call((el||document).querySelectorAll(s)); }
function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];}); }
function pct(n,d){ return d?Math.round(1000*n/d)/10:0; }
function fmt(n){ return String(n).replace('.',','); }
function toast(msg){ var t=$('#toast'); t.textContent=msg; t.classList.add('show'); clearTimeout(toast._h); toast._h=setTimeout(function(){t.classList.remove('show');},1800); }
function copy(txt,msg){
  var done=function(){ toast(msg||'Copié'); };
  function fallback(){ var ta=document.createElement('textarea'); ta.value=txt; document.body.appendChild(ta); ta.select(); try{document.execCommand('copy');done();}catch(e){toast('Copie impossible');} document.body.removeChild(ta); }
  if(navigator.clipboard&&navigator.clipboard.writeText){ navigator.clipboard.writeText(txt).then(done,fallback); } else fallback();
}
function scopeRows(){ return S.host<0?ROWS:HOSTS[S.host].rows; }
function scopeLabel(){ return S.host<0?'Tout le parc':HOSTS[S.host].name; }

/* --- filtrage --- */
function visible(){
  var q=S.q.trim().toLowerCase(), sv=Object.keys(S.sev);
  return scopeRows().filter(function(r){
    if(sv.length && !S.sev[r._k]) return false;
    if(S.cat && r.cat!==S.cat) return false;
    if(q && r._t.indexOf(q)<0) return false;
    return true;
  });
}
var SORTERS={
  id:function(a,b){ return (parseInt(a.id,10)||0)-(parseInt(b.id,10)||0) || a.id.localeCompare(b.id); },
  host:function(a,b){ return HOSTS[a.h].name.localeCompare(HOSTS[b.h].name) || SORTERS.sev(b,a); },
  name:function(a,b){ return a.name.localeCompare(b.name); },
  sev:function(a,b){ var sa=a.st==='Failed'?a.rank+10:(a.st==='Info'?5:0), sb=b.st==='Failed'?b.rank+10:(b.st==='Info'?5:0); return sa-sb || SORTERS.id(a,b); },
  res:function(a,b){ return a.res.localeCompare(b.res); },
  rec:function(a,b){ return a.rec.localeCompare(b.rec); },
  /* mode tests */
  nfail:function(a,b){ return a.nfail-b.nfail || SORTERS.sev(a,b); }
};
function sorted(list){
  var f=SORTERS[S.sort]||SORTERS.sev, d=S.dir==='asc'?1:-1;
  return list.slice().sort(function(a,b){ return d*f(a,b); });
}

/* --- regroupement par test (parc) --- */
function groupTests(list){
  var m={};
  list.forEach(function(r){
    var g=m[r.id]||(m[r.id]={id:r.id,name:r.name,cat:r.cat,rec:r.rec,meth:r.meth,def:r.def,rank:0,sev:'None',st:'Passed',nfail:0,npass:0,ninfo:0,hosts:[],_k:'Passed',_i:'t'+r.id});
    g.hosts.push(r);
    if(r.st==='Failed'){ g.nfail++; if(r.rank>g.rank){g.rank=r.rank;g.sev=r.sev;} }
    else if(r.st==='Passed') g.npass++; else g.ninfo++;
  });
  var out=Object.keys(m).map(function(k){ var g=m[k]; g.st=g.nfail?'Failed':(g.npass?'Passed':'Info'); g._k=g.st==='Failed'?g.sev:(g.st==='Passed'?'Passed':'Info'); g.res=g.nfail+'/'+(g.nfail+g.npass); return g; });
  return out;
}

/* --- rendu : en-tête / portée --- */
function renderScope(){
  if(!MULTI) return;
  var sel=$('#scope'); sel.innerHTML='<option value="-1">Tout le parc ('+HOSTS.length+' machines)</option>'+
    HOSTS.slice().sort(function(a,b){return a.name.localeCompare(b.name);}).map(function(h){ return '<option value="'+h.i+'">'+esc(h.name)+' — '+fmt(h.score)+' %</option>'; }).join('');
  sel.value=String(S.host);
  $('#fleet').style.display=S.host<0?'':'none';
  $('#modeseg').style.display=S.host<0?'':'none';
  if(S.host>=0 && S.mode==='tests') S.mode='rows';
  var h=S.host<0?null:HOSTS[S.host];
  $('#meta').title=h?(h.file+(h.list?' — liste '+h.list:'')):'';
  $('#meta').innerHTML=h?'<b>'+esc(h.file)+'</b>':HOSTS.length+' rapports &nbsp; '+esc(META.gen);
}

/* --- rendu : triage --- */
function renderTriage(){
  var rows=scopeRows(), tot=rows.length, c={}; SEV_ORDER.forEach(function(k){c[k]=0;});
  var passed=0,failed=0,high=0;
  rows.forEach(function(r){ c[r._k]++; if(r.st==='Passed')passed++; else if(r.st==='Failed'){failed++; if(r.rank>=3)high++;} });
  var score=pct(passed,passed+failed);
  var st=$('#stats'); st.innerHTML='';
  var sc=document.createElement('div'); sc.className='stat score';
  sc.innerHTML='<span class="l"><abbr title="Réussis / (réussis + échoués). Les tests non évalués ne comptent pas.">Conformité</abbr></span><span class="v">'+fmt(score)+' %</span><span class="p">'+passed+' / '+(passed+failed)+' tests évalués</span>';
  st.appendChild(sc);
  var failKeys=['Critical','High','Medium','Low'].filter(function(k){return c[k];});
  var allFailOn=failKeys.length&&failKeys.every(function(k){return S.sev[k];})&&!S.sev.Passed&&!S.sev.Info;
  var fb=document.createElement('button'); fb.className='stat'+(allFailOn?' on':''); fb.style.setProperty('--sc','var(--high)');
  fb.title='Afficher seulement les tests échoués'; fb.innerHTML='<span class="l">Échoués</span><span class="v">'+failed+'</span><span class="p">'+fmt(pct(failed,tot))+' %</span>';
  fb.onclick=function(){ if(allFailOn){S.sev={};} else { S.sev={}; failKeys.forEach(function(k){S.sev[k]=1;}); } update(); };
  st.appendChild(fb);
  SEV_ORDER.forEach(function(k){
    if(!c[k]) return;
    var b=document.createElement('button'); b.className='stat'+(S.sev[k]?' on':''); b.style.setProperty('--sc','var(--'+SEV_VAR[k]+')');
    b.title='Filtrer : '+SEV_LABEL[k]; b.innerHTML='<span class="l"><i></i>'+SEV_LABEL[k]+'</span><span class="v">'+c[k]+'</span><span class="p">'+fmt(pct(c[k],tot))+' %</span>';
    b.onclick=function(){ toggleSev(k); }; st.appendChild(b);
  });
  var lead;
  if(S.host<0&&MULTI){
    var worst=HOSTS.slice().sort(function(a,b){return a.score-b.score;})[0];
    var nh=HOSTS.filter(function(h){return h.high>0;}).length;
    lead=HOSTS.length+' machines, '+failed+' constats non conformes au total'+(high?', dont '+high+' de sévérité haute sur '+nh+' machine'+(nh>1?'s':''):'')+'. Machine la moins conforme : '+worst.name+' ('+fmt(worst.score)+' %). Vue « Par test » : ce qui échoue partout.';
  } else if(!failed) lead='Tous les tests évalués sont conformes.';
  else lead=failed+' tests ne sont pas conformes'+(high?', dont '+high+' de sévérité haute : commence par ceux-là.':', aucun de sévérité haute.');
  $('#lead').textContent=lead;
  var bar=$('#bar'); bar.innerHTML='';
  SEV_ORDER.forEach(function(k){
    if(!c[k]) return;
    var b=document.createElement('button'); b.className='s-'+k+(S.sev[k]?' on':''); b.style.flex=c[k]+' 1 0';
    b.title=SEV_LABEL[k]+' : '+c[k]+' ('+fmt(pct(c[k],tot))+' %)'; b.setAttribute('aria-label',b.title);
    b.innerHTML='<span>'+(pct(c[k],tot)>=6?c[k]:'')+'</span>'; b.onclick=function(){ toggleSev(k); }; bar.appendChild(b);
  });
}

/* --- rendu : tableau du parc --- */
var FSORT={name:function(a,b){return a.name.localeCompare(b.name);},score:function(a,b){return a.score-b.score;},high:function(a,b){return a.high-b.high;},med:function(a,b){return a.med-b.med;},low:function(a,b){return a.low-b.low;},failed:function(a,b){return a.failed-b.failed;},n:function(a,b){return a.rows.length-b.rows.length;}};
function renderFleet(){
  if(!MULTI||S.host>=0) return;
  var tb=$('#ftb'); tb.innerHTML='';
  $$('#fleet thead th').forEach(function(th){ th.classList.remove('asc','desc'); if(th.dataset.k===S.fsort) th.classList.add(S.fdir); });
  var d=S.fdir==='asc'?1:-1, f=FSORT[S.fsort]||FSORT.score;
  HOSTS.slice().sort(function(a,b){ return d*f(a,b) || a.name.localeCompare(b.name); }).forEach(function(h){
    var tr=document.createElement('tr'); tr.tabIndex=0; tr.title='Voir le détail de '+h.name;
    var n=h.rows.length;
    tr.innerHTML='<td class="host">'+esc(h.name)+'<small>'+esc(h.list||h.file)+'</small></td>'+
      '<td><div class="sbar"><i style="background:var(--pass);width:'+pct(h.passed,n)+'%"></i><i style="background:var(--high);width:'+pct(h.high,n)+'%"></i><i style="background:var(--med);width:'+pct(h.med,n)+'%"></i><i style="background:var(--low);width:'+pct(h.low,n)+'%"></i></div></td>'+
      '<td class="num"><b>'+fmt(h.score)+' %</b></td><td class="num'+(h.high?' h':'')+'">'+h.high+'</td><td class="num">'+h.med+'</td><td class="num">'+h.low+'</td><td class="num">'+h.failed+'</td><td class="num hide-m">'+n+'</td>';
    tr.onclick=function(){ S.host=h.i; S.open={}; update(); window.scrollTo({top:0}); };
    tr.onkeydown=function(e){ if(e.key==='Enter'){ tr.onclick(); } };
    tb.appendChild(tr);
  });
}

/* --- rendu : latéral --- */
function renderCats(){
  var m={}; scopeRows().forEach(function(r){ var g=m[r.cat]||(m[r.cat]={n:0,f:0,h:0,p:0}); g.n++; if(r.st==='Failed'){g.f++; if(r.rank>=3)g.h++;} else if(r.st==='Passed')g.p++; });
  var cats=Object.keys(m).sort(function(a,b){ return m[b].f-m[a].f || a.localeCompare(b); });
  var box=$('#cats'); box.innerHTML='';
  cats.forEach(function(cn){
    var g=m[cn], b=document.createElement('button'); b.className='cat'+(S.cat===cn?' on':''); b.title=cn;
    b.innerHTML='<span class="n">'+esc(cn||'(sans catégorie)')+'</span><span class="c">'+(g.f?'<b>'+g.f+'</b> / ':'')+g.n+'</span>'+
      '<span class="mini"><i style="background:var(--pass);width:'+pct(g.p,g.n)+'%"></i><i style="background:var(--high);width:'+pct(g.h,g.n)+'%"></i><i style="background:var(--med);width:'+pct(g.f-g.h,g.n)+'%"></i></span>';
    b.onclick=function(){ S.cat=(S.cat===cn)?'':cn; update(); };
    box.appendChild(b);
  });
}

function toggleSev(k){ if(S.sev[k]) delete S.sev[k]; else S.sev[k]=1; update(); }

/* --- rendu : filtres actifs --- */
function renderActive(){
  var a=$('#active'); a.innerHTML='';
  function tag(label,fn){ var t=document.createElement('span'); t.className='tag'; t.innerHTML=esc(label)+' <button aria-label="Retirer ce filtre">×</button>'; t.querySelector('button').onclick=fn; a.appendChild(t); }
  if(S.cat) tag('Catégorie : '+S.cat,function(){S.cat='';update();});
  if(S.q) tag('Recherche : '+S.q,function(){S.q='';$('#q').value='';update();});
  Object.keys(S.sev).forEach(function(k){ tag(SEV_LABEL[k],function(){delete S.sev[k];update();}); });
  if(a.children.length){ var c=document.createElement('button'); c.className='btn'; c.textContent='Tout effacer'; c.onclick=clearAll; a.appendChild(c); }
}
function clearAll(){ S.q='';S.cat='';S.sev={};$('#q').value='';$('#qx').classList.remove('show');update(); }

/* --- rendu : tableau principal --- */
var LAST=[];
function renderTable(){
  var tests=(S.mode==='tests'&&S.host<0&&MULTI);
  var list=visible();
  if(tests){ list=groupTests(list); if(!SORTERS[S.sort]||S.sort==='host') S.sort='nfail'; if(S.sort==='sev') S.sort='nfail'; }
  else if(S.sort==='nfail') S.sort='sev';
  LAST=sorted(list);
  var hd=$('#thead'); var fleetCol=(MULTI&&S.host<0&&!tests);
  hd.innerHTML='<tr>'+(fleetCol?'<th data-k="host">Machine<span class="d">▲</span></th>':'')+
    '<th data-k="id">ID<span class="d">▲</span></th><th data-k="name">Test<span class="d">▲</span></th>'+
    (tests?'<th data-k="nfail">Machines en échec<span class="d">▲</span></th><th data-k="sev">Sévérité max<span class="d">▲</span></th>':'<th data-k="sev">Résultat<span class="d">▲</span></th><th data-k="res" class="hide-m">Valeur actuelle<span class="d">▲</span></th>')+
    '<th data-k="rec" class="hide-m">Recommandée<span class="d">▲</span></th></tr>';
  $$('th',hd).forEach(function(th){ if(th.dataset.k===S.sort) th.classList.add(S.dir); th.onclick=function(){ var k=th.dataset.k; if(S.sort===k) S.dir=S.dir==='asc'?'desc':'asc'; else { S.sort=k; S.dir=(k==='sev'||k==='nfail')?'desc':'asc'; } update(); }; });
  var tb=$('#tb'); tb.innerHTML='';
  $('#count').textContent=LAST.length+' / '+(tests?groupTests(scopeRows()).length+' tests':scopeRows().length+' lignes');
  $('#empty').style.display=LAST.length?'none':'block';
  $('#t').style.display=LAST.length?'':'none';
  var frag=document.createDocumentFragment(), ncol=fleetCol?6:5;
  LAST.forEach(function(r){
    var tr=document.createElement('tr'); tr.className='r s-'+r._k+(S.open[r._i]?' open':''); tr.tabIndex=0; tr.dataset.i=r._i;
    tr.setAttribute('aria-expanded',S.open[r._i]?'true':'false');
    var bad=r.st==='Failed';
    var html=(fleetCol?'<td class="hn">'+esc(HOSTS[r.h].name)+'</td>':'')+
      '<td class="id">'+esc(r.id)+'</td>'+
      '<td class="nm">'+esc(r.name)+(S.cat?'':'<span class="c">'+esc(r.cat)+'</span>')+'</td>';
    if(tests){
      var tot=r.nfail+r.npass;
      html+='<td class="num"><span class="hbar"><i style="width:'+pct(r.nfail,tot)+'%"></i></span>'+r.nfail+' / '+tot+'</td>'+
        '<td><span class="pill '+r._k+'"><i></i>'+(bad?r.sev:SEV_LABEL[r._k])+'</span></td>';
    } else {
      html+='<td><span class="pill '+r._k+'"><i></i>'+(bad?r.sev:SEV_LABEL[r._k])+'</span></td>'+
        '<td class="v'+(bad?' bad':'')+' hide-m">'+esc(r.res||'—')+'</td>';
    }
    html+='<td class="v hide-m">'+esc(r.rec||'—')+'</td>';
    tr.innerHTML=html; frag.appendChild(tr);
    if(S.open[r._i]) frag.appendChild(tests?detailTest(r,ncol):detail(r,ncol));
  });
  tb.appendChild(frag);
}
function detail(r,ncol){
  var tr=document.createElement('tr'); tr.className='x'; var bad=r.st==='Failed';
  tr.innerHTML='<td colspan="'+ncol+'"><div class="det">'+
    '<div><div class="k">Valeur actuelle</div><div class="vv'+(bad?' bad':' good')+'">'+esc(r.res||'(vide)')+'</div></div>'+
    '<div><div class="k">Valeur recommandée</div><div class="vv">'+esc(r.rec||'(vide)')+'</div></div>'+
    (r.def?'<div><div class="k">Valeur par défaut Windows</div><div class="vv">'+esc(r.def)+'</div></div>':'')+
    (r.meth?'<div><div class="k">Méthode de lecture</div><div class="vv">'+esc(r.meth)+'</div></div>':'')+
    '<div><div class="k">Catégorie</div><div class="vv">'+esc(r.cat||'—')+'</div></div>'+
    '<div class="acts"><button class="btn" data-a="copy">Copier ID + nom</button><button class="btn" data-a="copyrec">Copier la valeur recommandée</button><button class="btn" data-a="cat">Voir la catégorie</button>'+
    (MULTI?'<button class="btn" data-a="test">Ce test sur tout le parc</button>':'')+'</div></div></td>';
  tr.querySelector('[data-a=copy]').onclick=function(e){ e.stopPropagation(); copy(r.id+' '+r.name); };
  tr.querySelector('[data-a=copyrec]').onclick=function(e){ e.stopPropagation(); copy(r.rec); };
  tr.querySelector('[data-a=cat]').onclick=function(e){ e.stopPropagation(); S.cat=r.cat; update(); };
  if(MULTI) tr.querySelector('[data-a=test]').onclick=function(e){ e.stopPropagation(); S.host=-1; S.mode='tests'; S.cat=''; S.sev={}; S.q=r.id; $('#q').value=r.id; $('#qx').classList.add('show'); S.open={}; S.open['t'+r.id]=1; update(); };
  return tr;
}
function detailTest(g,ncol){
  var tr=document.createElement('tr'); tr.className='x';
  var hosts=g.hosts.slice().sort(function(a,b){ return (b.st==='Failed')-(a.st==='Failed') || HOSTS[a.h].name.localeCompare(HOSTS[b.h].name); });
  tr.innerHTML='<td colspan="'+ncol+'"><div class="det">'+
    '<div class="full"><div class="k">Machines ('+g.nfail+' en échec, '+g.npass+' conformes'+(g.ninfo?', '+g.ninfo+' non évaluées':'')+')</div><div class="hosts">'+
    hosts.map(function(r){ return '<button class="'+(r.st==='Failed'?'ko':(r.st==='Passed'?'ok':''))+'" data-h="'+r.h+'" title="'+esc(r.st==='Failed'?'Valeur : '+(r.res||'(vide)'):r.st)+'">'+esc(HOSTS[r.h].name)+'</button>'; }).join('')+'</div></div>'+
    '<div><div class="k">Valeur recommandée</div><div class="vv">'+esc(g.rec||'(vide)')+'</div></div>'+
    (g.def?'<div><div class="k">Valeur par défaut Windows</div><div class="vv">'+esc(g.def)+'</div></div>':'')+
    '<div class="acts"><button class="btn" data-a="copy">Copier ID + nom</button><button class="btn" data-a="hosts">Copier les machines en échec</button></div></div></td>';
  $$('.hosts button',tr).forEach(function(b){ b.onclick=function(e){ e.stopPropagation(); S.host=+b.dataset.h; S.mode='rows'; S.q=g.id; $('#q').value=g.id; $('#qx').classList.add('show'); S.open={}; update(); window.scrollTo({top:0}); }; });
  tr.querySelector('[data-a=copy]').onclick=function(e){ e.stopPropagation(); copy(g.id+' '+g.name); };
  tr.querySelector('[data-a=hosts]').onclick=function(e){ e.stopPropagation(); var l=g.hosts.filter(function(r){return r.st==='Failed';}).map(function(r){return HOSTS[r.h].name;}); copy(l.join('\n'),l.length+' machines copiées'); };
  return tr;
}
function toggleRow(i){ if(S.open[i]) delete S.open[i]; else S.open[i]=1; renderTable(); }

/* --- export --- */
function csvOf(list){
  var tests=(S.mode==='tests'&&S.host<0&&MULTI);
  var q=function(v){ v=String(v==null?'':v); return /[";\n\r]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; };
  var lines;
  if(tests){ lines=['ID;Category;Name;HostsFailed;HostsPassed;MaxSeverity;Recommended;FailedHosts'];
    list.forEach(function(g){ lines.push([g.id,g.cat,g.name,g.nfail,g.npass,g.sev,g.rec,g.hosts.filter(function(r){return r.st==='Failed';}).map(function(r){return HOSTS[r.h].name;}).join(',')].map(q).join(';')); }); }
  else { lines=['Host;ID;Category;Name;Status;Severity;Result;Recommended;DefaultValue'];
    list.forEach(function(r){ lines.push([HOSTS[r.h].name,r.id,r.cat,r.name,r.st,r.sev,r.res,r.rec,r.def].map(q).join(';')); }); }
  return '﻿'+lines.join('\r\n');
}
function download(name,content,type){
  var a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([content],{type:type})); a.download=name; document.body.appendChild(a); a.click();
  setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); },500);
}

/* --- thème --- */
function applyTheme(t){
  if(t==='auto') document.documentElement.removeAttribute('data-theme'); else document.documentElement.setAttribute('data-theme',t);
  try{ localStorage.setItem('hk-theme',t); }catch(e){}
  $('#theme').title={auto:'Thème : auto (cliquer pour clair)',light:'Thème : clair (cliquer pour sombre)',dark:'Thème : sombre (cliquer pour auto)'}[t]; $('#theme').dataset.t=t;
  $('#theme').innerHTML={auto:'<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8"/><path d="M12 4a8 8 0 0 1 0 16z" fill="currentColor" stroke="none"/></svg>',light:'<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M2 12h2m16 0h2M4.9 4.9l1.4 1.4m11.4 11.4 1.4 1.4M4.9 19.1l1.4-1.4m11.4-11.4 1.4-1.4"/></svg>',dark:'<svg viewBox="0 0 24 24"><path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z"/></svg>'}[t];
}

/* --- orchestration --- */
function update(){ renderScope(); renderTriage(); renderFleet(); renderCats(); renderActive(); renderMode(); renderTable(); renderOnlyFail(); writeHash(); }
function failKeys(){ var c={}; scopeRows().forEach(function(r){c[r._k]=1;}); return ['Critical','High','Medium','Low'].filter(function(k){return c[k];}); }
function onlyFailOn(){ var fk=failKeys(); return fk.length&&fk.every(function(k){return S.sev[k];})&&!S.sev.Passed&&!S.sev.Info; }
function renderOnlyFail(){ $('#onlyfail').classList.toggle('on',!!onlyFailOn()); }
function renderMode(){ if(!MULTI) return; $$('#modeseg button').forEach(function(b){ b.classList.toggle('on',b.dataset.m===S.mode); }); $('#side').style.display=(S.mode==='tests')?'none':''; $('.body').style.gridTemplateColumns=(S.mode==='tests')?'1fr':''; }

document.addEventListener('DOMContentLoaded',function(){
  readHash();
  var t='auto'; try{ t=localStorage.getItem('hk-theme')||'auto'; }catch(e){}
  applyTheme(t);
  $('#theme').onclick=function(){ applyTheme({auto:'light',light:'dark',dark:'auto'}[this.dataset.t]); };
  $('#onlyfail').onclick=function(){ var on=onlyFailOn(); S.sev={}; if(!on) failKeys().forEach(function(k){S.sev[k]=1;}); update(); };
  if(MULTI){
    $('#scope').onchange=function(){ S.host=+this.value; S.open={}; update(); };
    $$('#modeseg button').forEach(function(b){ b.onclick=function(){ S.mode=b.dataset.m; S.open={}; update(); }; });
    $$('#fleet thead th').forEach(function(th){ th.onclick=function(){ var k=th.dataset.k; if(S.fsort===k) S.fdir=S.fdir==='asc'?'desc':'asc'; else { S.fsort=k; S.fdir=(k==='name')?'asc':(k==='score'?'asc':'desc'); } renderFleet(); }; });
  }
  $('#q').value=S.q; $('#qx').classList.toggle('show',!!S.q);
  $('#q').addEventListener('input',function(){ S.q=this.value; $('#qx').classList.toggle('show',!!this.value); update(); });
  $('#qx').onclick=function(){ $('#q').value=''; S.q=''; this.classList.remove('show'); update(); $('#q').focus(); };
  $('#catclear').onclick=function(){ S.cat=''; update(); };
  $('#tb').addEventListener('click',function(e){ var tr=e.target.closest('tr.r'); if(tr) toggleRow(tr.dataset.i); });
  $('#tb').addEventListener('keydown',function(e){ var tr=e.target.closest('tr.r'); if(tr&&(e.key==='Enter'||e.key===' ')){ e.preventDefault(); toggleRow(tr.dataset.i); } });
  $('#exp').onclick=function(){ download((META.base||'hardeningkitty')+'_'+scopeLabel().replace(/\W+/g,'_')+'.csv',csvOf(LAST),'text/csv;charset=utf-8'); toast(LAST.length+' lignes exportées'); };
  $('#copyids').onclick=function(){ var ids={}; LAST.forEach(function(r){ if(r.st==='Failed') ids[r.id]=1; }); var l=Object.keys(ids); if(!l.length){toast('Aucun échec dans la sélection');return;} copy(l.join(', '),l.length+' IDs copiés'); };
  $('#share').onclick=function(){ copy(location.href,'Lien avec les filtres copié'); };
  $('#print').onclick=function(){ window.print(); };
  $('#expand').onclick=function(){ var all=LAST.length&&LAST.every(function(r){return S.open[r._i];}); S.open={}; if(!all) LAST.forEach(function(r){S.open[r._i]=1;}); renderTable(); };
  $('#help').onclick=function(){ $('#dlg').showModal(); };
  $('#dlgclose').onclick=function(){ $('#dlg').close(); };
  $('#up').onclick=function(){ window.scrollTo({top:0}); };
  window.addEventListener('scroll',function(){ $('#up').classList.toggle('show',window.scrollY>600); },{passive:true});
  window.addEventListener('hashchange',function(){ readHash(); $('#q').value=S.q; update(); });
  document.addEventListener('keydown',function(e){
    if(e.target.tagName==='INPUT'||e.target.tagName==='SELECT'||e.target.tagName==='TEXTAREA'){ if(e.key==='Escape'){ e.target.blur(); if(e.target.id==='q'&&S.q) $('#qx').click(); } return; }
    if(e.key==='/'){ e.preventDefault(); $('#q').focus(); $('#q').select(); }
    else if(e.key==='Escape'){ if($('#dlg').open) $('#dlg').close(); else clearAll(); }
    else if(e.key==='?'){ $('#dlg').open?$('#dlg').close():$('#dlg').showModal(); }
    else if(e.key==='f'){ $('#onlyfail').click(); }
    else if(e.key==='e'){ $('#exp').click(); }
    else if(e.key==='t'&&MULTI){ S.mode=S.mode==='tests'?'rows':'tests'; S.host=-1; update(); }
    else if(e.key==='p'&&MULTI){ S.host=-1; S.open={}; update(); window.scrollTo({top:0}); }
  });
  var tools=$('.tools');
  function measure(){ document.documentElement.style.setProperty('--tools-h',tools.offsetHeight+'px'); }
  if(window.ResizeObserver) new ResizeObserver(measure).observe(tools); else window.addEventListener('resize',measure);
  update(); measure();
});
})();
'@

function ConvertTo-JsonSafe($obj) {
    $json = $obj | ConvertTo-Json -Depth 6 -Compress
    # un "</script>" dans une valeur casserait la page
    return $json.Replace('</', '<\/').Replace('<!--', '<\!--')
}

function Get-HostLabel([System.IO.FileInfo]$f) {
    # hardeningkitty_report_SRV01_20261104.csv -> SRV01 ; hk_SRV01_20261104_030000.csv -> SRV01 ; latest.csv -> nom du dossier parent
    # Formats : hardeningkitty_report_SRV01_finding_list_xxx_machine(.csv)-20261104-181431.csv (nom par défaut de HardeningKitty),
    #           hk_SRV01_20261104_030000.csv (Invoke-HardeningKittyAudit.ps1), latest.csv -> nom du dossier parent
    $b = $f.BaseName
    $b = $b -replace '^(hardeningkitty_report|hk|report|rapport)[_-]', ''
    $b = $b -replace '-\d{8}-\d{6}$', ''
    $b = $b -replace '[_-]finding_list_.*$', ''
    $b = $b -replace '[_-]\d{8}([_-]\d{4,6})?$', ''
    $b = $b -replace '[_-]\d{4}-\d{2}-\d{2}([_T-]\d{2}[-:h]?\d{2}([-:]?\d{2})?)?$', ''
    if ($b -eq '' -or $b -ieq 'latest') { $b = $f.Directory.Name }
    return $b
}

function Get-ListFromName([System.IO.FileInfo]$f) {
    if ($f.BaseName -match '(finding_list_.+?)(\.csv)?(-\d{8}-\d{6})?$') { return $Matches[1] + '.csv' }
    return ''
}

function New-HkPage([object[]]$hosts, [string]$pageTitle, [string]$baseName) {
    $multi = $hosts.Count -gt 1
    $gen   = Get-Date -Format 'dd/MM/yyyy HH:mm'

    $data = ConvertTo-JsonSafe @{
        meta  = @{ title = $pageTitle; base = $baseName; gen = $gen }
        hosts = @($hosts | ForEach-Object { @{ name = $_.Name; file = $_.File; list = $_.List; rows = @($_.Rows | Select-Object id, cat, name, meth, def, res, rec, st, sev, rank) } })
    }

    $icoSearch = '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>'
    $icoDl     = '<svg viewBox="0 0 24 24"><path d="M12 4v11m0 0 4-4m-4 4-4-4M4 19h16"/></svg>'
    $icoCopy   = '<svg viewBox="0 0 24 24"><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V6a2 2 0 0 1 2-2h9"/></svg>'
    $icoLink   = '<svg viewBox="0 0 24 24"><path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/></svg>'

    $singleMeta = ''; $singleTitle = ''
    if (-not $multi) {
        $h = $hosts[0]
        $singleMeta = "<b>$(Esc $h.Name)</b> &nbsp; $gen"
        $singleTitle = "Rapport $($h.File)"; if ($h.List) { $singleTitle += " — liste $($h.List)" }
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width,initial-scale=1">')
    [void]$sb.AppendLine("<title>$(Esc $pageTitle)</title><style>$Css</style></head><body>")

    # en-tête
    [void]$sb.AppendLine('<header class="top"><div class="wrap">')
    [void]$sb.AppendLine("<h1>$(Esc $pageTitle)</h1>")
    if ($multi) {
        [void]$sb.AppendLine('<div class="scope"><label for="scope">Machine</label><select class="sel" id="scope"></select></div>')
    }
    [void]$sb.AppendLine("<div class=""meta"" id=""meta"" title=""$(Esc $singleTitle)"">$singleMeta</div>")
    [void]$sb.AppendLine('<div class="right"><button class="ibtn" id="theme" data-t="auto" aria-label="Changer de thème"></button><button class="ibtn" id="help" title="Aide et raccourcis (?)" aria-label="Aide"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M9.5 9.5a2.5 2.5 0 1 1 3.5 2.3c-.7.3-1 .8-1 1.7M12 17h.01"/></svg></button></div>')
    [void]$sb.AppendLine('</div></header>')

    # triage
    [void]$sb.AppendLine('<section class="triage"><div class="wrap">')
    [void]$sb.AppendLine('<div class="stats" id="stats" role="group" aria-label="Résumé et filtres par résultat"></div>')
    [void]$sb.AppendLine('<p class="lead" id="lead"></p>')
    [void]$sb.AppendLine('<div class="bar" id="bar" role="group" aria-label="Répartition des résultats"></div>')
    if ($multi) {
        [void]$sb.AppendLine('<div class="fleet" id="fleet"><h2>Machines <span>cliquer une ligne pour ouvrir son rapport</span></h2><table><thead><tr><th data-k="name">Machine<span class="d">▲</span></th><th>Répartition</th><th data-k="score">Conformité<span class="d">▲</span></th><th data-k="high">High<span class="d">▲</span></th><th data-k="med">Medium<span class="d">▲</span></th><th data-k="low">Low<span class="d">▲</span></th><th data-k="failed">Échoués<span class="d">▲</span></th><th data-k="n" class="hide-m">Tests<span class="d">▲</span></th></tr></thead><tbody id="ftb"></tbody></table></div>')
    }
    [void]$sb.AppendLine('</div></section>')

    # corps
    [void]$sb.AppendLine('<div class="wrap body">')
    [void]$sb.AppendLine('<aside class="side" id="side"><h2>Catégories <button id="catclear">toutes</button></h2><div id="cats"></div></aside>')
    [void]$sb.AppendLine('<div class="main">')
    [void]$sb.AppendLine('<div class="tools">')
    if ($multi) {
        [void]$sb.AppendLine('<div class="seg" id="modeseg" role="group" aria-label="Mode d''affichage"><button data-m="rows">Par ligne</button><button data-m="tests">Par test</button></div>')
    }
    [void]$sb.AppendLine("<div class=""search"">$icoSearch<input id=""q"" type=""search"" placeholder=""Rechercher…"" aria-label=""Rechercher""><button class=""x"" id=""qx"" aria-label=""Effacer la recherche"">×</button></div>")
    [void]$sb.AppendLine('<button class="btn toggle" id="onlyfail" title="Afficher seulement les tests échoués (f)"><svg viewBox="0 0 24 24"><path d="M12 9v4m0 4h.01M10.3 3.9 2.4 17.6A2 2 0 0 0 4.1 21h15.8a2 2 0 0 0 1.7-3.4L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg>Échoués seuls</button>')
    [void]$sb.AppendLine('<span class="count" id="count"></span>')
    [void]$sb.AppendLine('<div class="grp">')
    [void]$sb.AppendLine('<button class="ibtn" id="expand" title="Déplier ou replier toutes les lignes visibles" aria-label="Déplier tout"><svg viewBox="0 0 24 24"><path d="m7 9 5-5 5 5M7 15l5 5 5-5"/></svg></button>')
    [void]$sb.AppendLine("<button class=""ibtn"" id=""copyids"" title=""Copier les IDs des tests échoués visibles"" aria-label=""Copier les IDs échoués"">$icoCopy</button>")
    [void]$sb.AppendLine("<button class=""ibtn"" id=""share"" title=""Copier un lien qui conserve les filtres"" aria-label=""Copier le lien"">$icoLink</button>")
    [void]$sb.AppendLine('<button class="ibtn" id="print" title="Imprimer" aria-label="Imprimer"><svg viewBox="0 0 24 24"><path d="M6 9V3h12v6M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="7"/></svg></button>')
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine("<button class=""btn primary"" id=""exp"" title=""Exporter les lignes visibles en CSV (e)"">$icoDl Exporter CSV</button>")
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('<div class="active" id="active"></div>')
    [void]$sb.AppendLine('<table id="t"><thead id="thead"></thead><tbody id="tb"></tbody></table>')
    [void]$sb.AppendLine('<div class="empty" id="empty" style="display:none"><b>Aucun test ne correspond</b>Retire un filtre ou modifie la recherche.</div>')
    [void]$sb.AppendLine('</div></div>')

    # aide, toast, haut de page
    $help = '<tr><td><kbd>/</kbd></td><td>Aller à la recherche</td></tr><tr><td><kbd>f</kbd></td><td>Afficher seulement les échoués</td></tr><tr><td><kbd>e</kbd></td><td>Exporter la sélection en CSV</td></tr><tr><td><kbd>Entrée</kbd></td><td>Déplier la ligne sélectionnée</td></tr><tr><td><kbd>Échap</kbd></td><td>Effacer tous les filtres</td></tr>'
    if ($multi) { $help += '<tr><td><kbd>p</kbd></td><td>Revenir à tout le parc</td></tr><tr><td><kbd>t</kbd></td><td>Basculer la vue par test</td></tr>' }
    $help += '<tr><td><kbd>?</kbd></td><td>Cette aide</td></tr>'
    [void]$sb.AppendLine("<dialog id=""dlg""><h3>Lecture du rapport</h3><p>Conformité = réussis / (réussis + échoués). Les tests non évalués ne comptent pas. Les compteurs et les segments de la barre filtrent le tableau ; cliquer une ligne l'ouvre.</p><h3>Raccourcis clavier</h3><table>$help</table><button class=""btn close"" id=""dlgclose"">Fermer</button></dialog>")
    [void]$sb.AppendLine('<div class="toast" id="toast" role="status" aria-live="polite"></div>')
    [void]$sb.AppendLine('<button class="btn up" id="up" aria-label="Revenir en haut">↑ Haut</button>')

    [void]$sb.AppendLine("<script id=""data"" type=""application/json"">$data</script>")
    [void]$sb.AppendLine("<script>$Js</script></body></html>")
    return $sb.ToString()
}

# ---------------------------------------------------------------- main

# Accepte un fichier, un dossier, un joker (.\hardeningkitty_report_*.csv) ou plusieurs chemins
$files = @()
foreach ($p in ($Path -split ',')) {
    $p = $p.Trim(); if (-not $p) { continue }
    $resolved = @(Resolve-Path -Path $p -ErrorAction SilentlyContinue)
    if ($resolved.Count -eq 0) { throw "Chemin introuvable : $p" }
    foreach ($r in $resolved) {
        $it = Get-Item -LiteralPath $r.ProviderPath
        if ($it.PSIsContainer) {
            $files += @(Get-ChildItem -LiteralPath $it.FullName -Filter *.csv -File -Recurse:$Recurse)
        } else {
            $files += $it
        }
    }
}
$files = @($files | Sort-Object FullName -Unique)
if ($files.Count -eq 0) { throw "Aucun fichier CSV trouvé dans $Path" }

$hosts = @()
foreach ($f in $files) {
    $rows = Import-HkCsv $f.FullName
    if ($rows.Count -eq 0) { Write-Warning "$($f.Name) : vide, ignoré."; continue }
    if ($rows[0].PSObject.Properties.Name -notcontains 'ID') {
        Write-Warning "$($f.Name) : colonne 'ID' absente, ce n'est sans doute pas un rapport HardeningKitty. Ignoré."; continue
    }
    $items = @(ConvertTo-HkModel $rows)
    $label = Get-HostLabel $f
    # évite deux machines du même nom (deux rapports du même hôte) : suffixe
    $n = 2; $base = $label
    while ($hosts | Where-Object { $_.Name -eq $label }) { $label = "$base ($n)"; $n++ }

    # liste utilisée, si l'audit l'a notée à côté (latest.list.txt ou <base>.list.txt)
    $list = Get-ListFromName $f
    foreach ($cand in @((Join-Path $f.DirectoryName ($f.BaseName + '.list.txt')), (Join-Path $f.DirectoryName 'latest.list.txt'))) {
        if (Test-Path -LiteralPath $cand) { $list = (Get-Content -LiteralPath $cand -TotalCount 1).Trim(); break }
    }

    $hosts += [pscustomobject]@{ Name = $label; File = $f.Name; List = $list; Rows = $items; Score = Get-Score $items }
    Write-Host ("  {0,-30} {1,5} tests  {2,6} %  ({3})" -f $label, $items.Count, (Get-Score $items), $f.Name)
}
if ($hosts.Count -eq 0) { throw 'Aucun rapport exploitable.' }

if (-not $OutFile) {
    $dir = $files[0].DirectoryName
    if ($hosts.Count -eq 1) { $OutFile = Join-Path $dir ($files[0].BaseName + '.html') }
    else { $OutFile = Join-Path $dir 'hardeningkitty_parc.html' }
}
$outDir = Split-Path $OutFile -Parent
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$html = New-HkPage -hosts $hosts -pageTitle $Title -baseName ([System.IO.Path]::GetFileNameWithoutExtension($OutFile))
[System.IO.File]::WriteAllText($OutFile, $html, (New-Object System.Text.UTF8Encoding $false))
Write-Host "OK  $($hosts.Count) rapport(s) -> $OutFile"

if ($Open) { Start-Process $OutFile }
