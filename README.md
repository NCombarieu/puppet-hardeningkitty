# HardeningKitty toolkit

Audit de durcissement Windows avec [HardeningKitty](https://github.com/scipag/HardeningKitty) (scip AG),
intégré à Puppet et NSClient++, avec un rapport HTML autonome pour une machine ou tout le parc.

Tout est en **mode Audit** : rien ne modifie la configuration des machines.

```
hardeningkitty-toolkit/
├── report/
│   └── ConvertTo-HardeningKittyHtml.ps1   CSV HardeningKitty -> page HTML (1 ou N rapports)
├── audit/
│   ├── Invoke-HardeningKittyAudit.ps1     audit planifié, choix auto de la liste, rotation
│   ├── check_hardeningkitty.ps1           check Nagios/Shinken pour NSClient++
│   ├── run-manual.ps1                     installation sans Puppet (tâche planifiée en option)
│   └── HardeningKitty/                    HardeningKitty 0.9.4 (psd1, psm1, lists/, LICENSE MIT)
└── puppet/hardeningkitty/                 module Puppet Windows
    ├── manifests/init.pp
    └── files/
        ├── Invoke-HardeningKittyAudit.ps1
        ├── check_hardeningkitty.ps1
        └── HardeningKitty/                la même copie de HardeningKitty, déployée par Puppet
```

HardeningKitty est inclus (dépôt scipag/HardeningKitty, branche master au 21/07/2026, module 0.9.4,
132 finding lists, licence MIT). Pour le mettre à jour : remplacer le contenu des deux dossiers
`HardeningKitty/` par une release plus récente.

## Prérequis

- Windows avec PowerShell 5.1 (ou 7+). Le rapport HTML se génère aussi sous Linux avec `pwsh`.
- Pour Puppet : le module `puppetlabs-scheduled_task`.
- Les scripts ont été écrits et relus mais **pas exécutés sous PowerShell** : fais un premier passage
  sur une VM de test et remonte ce qui coince.

## 0. Avant tout : débloquer les scripts

Windows marque les fichiers extraits d'un zip téléchargé comme « venant d'Internet », et PowerShell
refuse alors de les exécuter (« n'est pas signé numériquement »). Deux solutions :

```powershell
# a) une fois pour toutes sur le dossier extrait (Unblock-File ne descend pas tout seul dans les sous-dossiers)
Get-ChildItem C:\chemin\hardeningkitty-toolkit -Recurse | Unblock-File

# b) ou seulement pour la session PowerShell courante
Set-ExecutionPolicy -Scope Process Bypass
```

Plus simple encore : clic droit sur le zip > Propriétés > « Débloquer » **avant** de l'extraire.

## 1. Essai rapide sur un poste

PowerShell **en administrateur**, dans `audit\HardeningKitty\` :

```powershell
Set-ExecutionPolicy -Scope Process Bypass
Get-ChildItem -Recurse | Unblock-File
Import-Module .\HardeningKitty.psm1
Invoke-HardeningKitty -Mode Audit -Log -Report
..\..\report\ConvertTo-HardeningKittyHtml.ps1 -Path .\hardeningkitty_report_*.csv -Open
```

## 2. Rapport HTML : `report/ConvertTo-HardeningKittyHtml.ps1`

Une seule page HTML, sans dépendance, utilisable hors ligne.

```powershell
# un rapport -> <rapport>.html à côté du CSV
.\ConvertTo-HardeningKittyHtml.ps1 -Path .\hardeningkitty_report_SRV01_20261104.csv -Open

# tout un parc : un dossier de CSV, ou un dossier par machine contenant latest.csv
.\ConvertTo-HardeningKittyHtml.ps1 -Path \\partage\hk -Recurse -OutFile C:\web\parc.html -Title "Parc CNCV"
```

| Paramètre  | Rôle |
|------------|------|
| `-Path`    | fichier, dossier, joker (`*.csv`) ou plusieurs chemins séparés par des virgules |
| `-Recurse` | avec un dossier, descend dans les sous-dossiers |
| `-OutFile` | fichier HTML de sortie (défaut : `<rapport>.html`, ou `hardeningkitty_parc.html`) |
| `-Title`   | titre de la page |
| `-Open`    | ouvre la page dans le navigateur |

Le nom de machine est pris dans le nom du fichier (`hardeningkitty_report_<HOST>_<date>.csv`,
`hk_<HOST>_<date>.csv`) ou, pour `latest.csv`, dans le nom du dossier parent.

Dans la page : compteurs cliquables (conformité, échoués, High/Medium/Low, réussis) qui filtrent
le tableau, barre de répartition, catégories, recherche, lignes dépliables (valeur actuelle,
recommandée, valeur par défaut Windows), export CSV de la sélection, copie des IDs échoués, lien
partageable (les filtres sont dans l'URL), thème clair/sombre. Avec plusieurs rapports : sélecteur
de machine, tableau du parc, et vue **Par test** (quels tests échouent sur combien de machines).
`?` affiche l'aide et les raccourcis.

Conformité = réussis / (réussis + échoués). C'est un pourcentage, différent du score 0-6 de HardeningKitty.

## 3. Audit planifié : `audit/Invoke-HardeningKittyAudit.ps1`

Lance l'audit, écrit `reports\hk_<HOST>_<date>.csv`, recopie dans `reports\latest.csv`,
note la liste utilisée dans `reports\latest.list.txt`, garde les N derniers rapports.

```powershell
.\Invoke-HardeningKittyAudit.ps1 -InstallDir C:\ProgramData\HardeningKitty -FindingList auto -ListFamily cis
```

| Paramètre       | Rôle |
|-----------------|------|
| `-InstallDir`   | racine : `module\` (HardeningKitty), `custom_lists\`, `reports\` |
| `-FindingList`  | `auto` (défaut), un nom de fichier (cherché dans `custom_lists\` puis `module\lists\`) ou un chemin complet |
| `-ListFamily`   | référentiel pour `auto` : `cis` (défaut), `msft_security_baseline`, … |
| `-FallbackList` | liste de repli si rien ne correspond (défaut : `finding_list_0x6d69636b_machine.csv`) |
| `-KeepReports`  | rapports conservés (défaut 14) |

`auto` (fonction `Resolve-AutoList` du script) détecte l'OS par son build et cherche
`finding_list_<famille>_*<os>*_machine.csv`, en prenant la version la plus récente, et la variante
`_dc_` ou `_member_` selon que la machine est contrôleur de domaine ou non :

| Build   | OS          | `cis`                                            | `msft_security_baseline`                        |
|---------|-------------|--------------------------------------------------|-------------------------------------------------|
| 14393   | Server 2016 | `..._server_2016_1607_3.0.0_machine`             | aucune → repli                                  |
| 17763   | Server 2019 | `..._server_2019_1809_3.0.0_machine`             | aucune → repli                                  |
| 20348   | Server 2022 | `..._server_2022_22h2_4.0.0_machine`             | `..._server_2022_21h2_{member,dc}_machine`      |
| ≥ 26100 | Server 2025 | `..._server_2025_24h2_1.0.0_machine`             | `..._server_2025_2602_{member,dc}_machine`      |
| ≥ 22000 | Windows 11  | `..._windows_11_enterprise_24h2_machine`         | `..._windows_11_25h2_machine`                   |
| autre   | Windows 10  | `..._windows_10_enterprise_22h2_3.0.0_machine`   | `..._windows_10_22h2_machine`                   |

Sans correspondance : avertissement et repli sur `-FallbackList`. La liste réellement utilisée est
écrite dans `reports\latest.list.txt`.

Le script peut tourner **sans installation**, directement depuis `audit\` : s'il ne trouve pas
HardeningKitty dans `InstallDir\module`, il utilise `audit\HardeningKitty\` et écrit quand même les
rapports dans `InstallDir\reports` (créé au besoin).

```powershell
cd audit
.\Invoke-HardeningKittyAudit.ps1            # audit immédiat, rapport dans C:\ProgramData\HardeningKitty\reports\
.\check_hardeningkitty.ps1                 # puis le check lit latest.csv
```

### Installation sans Puppet

Dans `audit\` (HardeningKitty est déjà dans `audit\HardeningKitty\`) :

```powershell
.\run-manual.ps1                               # installe dans C:\ProgramData\HardeningKitty, sans tâche planifiée
.\run-manual.ps1 -RunNow                       # idem + premier audit immédiat
.\run-manual.ps1 -Schedule -AuditTime 02:30    # + tâche planifiée quotidienne à 02:30 (SYSTEM)
```

| Paramètre      | Défaut | Rôle |
|----------------|--------|------|
| `-InstallDir`  | `C:\ProgramData\HardeningKitty` | racine d'installation |
| `-Schedule`    | non    | crée la tâche planifiée « HardeningKitty audit » |
| `-AuditTime`   | 03:00  | heure de la tâche (HH:mm), avec `-Schedule` |
| `-FindingList` | auto   | liste passée au script d'audit |
| `-ListFamily`  | cis    | référentiel pour `auto` |
| `-RunNow`      | non    | lance un audit à la fin de l'installation |

Sans `-Schedule`, l'audit se lance à la main :
`powershell -ExecutionPolicy Bypass -File C:\ProgramData\HardeningKitty\scripts\Invoke-HardeningKittyAudit.ps1`

### Collecte pour la vue parc

Depuis chaque machine, copier `reports\latest.csv` vers `\\partage\hk\<HOST>\latest.csv`
(à ajouter dans la tâche planifiée ou dans le script), puis sur la machine qui publie :

```powershell
.\ConvertTo-HardeningKittyHtml.ps1 -Path \\partage\hk -Recurse -OutFile C:\web\parc.html
```

## 4. Check NSClient++ : `audit/check_hardeningkitty.ps1`

Lit `latest.csv` (ne lance pas l'audit, trop long pour un check). Codes retour Nagios 0/1/2/3,
perfdata `score`, `passed`, `failed`, `high`, `medium`, `low`, `age_h`.

| Paramètre      | Défaut | Rôle |
|----------------|--------|------|
| `-Warn`        | 80     | WARNING si conformité < |
| `-Crit`        | 60     | CRITICAL si conformité < |
| `-CritHigh`    | -1     | CRITICAL si échecs High ≥ (désactivé par défaut) |
| `-MaxAgeHours` | 48     | UNKNOWN si le rapport est plus vieux (tâche cassée) |
| `-ReportPath`  | `C:\ProgramData\HardeningKitty\reports\latest.csv` | |

`nsclient.ini` :

```ini
[/settings/external scripts/wrapped scripts]
check_hardeningkitty = check_hardeningkitty.ps1 -Warn 85 -Crit 70 -CritHigh 10
```

Ajuste les seuils après un premier relevé : sur un Windows en français, HardeningKitty (prévu pour
`en-US`) produit des faux positifs qui tirent le score vers le bas.

## 5. Module Puppet : `puppet/hardeningkitty`

1. Installer `puppetlabs-scheduled_task`.
2. Classer les nœuds Windows avec `include hardeningkitty` et régler en Hiera :

```yaml
hardeningkitty::install_dir: 'C:/ProgramData/HardeningKitty'
hardeningkitty::finding_list: auto          # ou 'cncv_serveurs_machine.csv'
hardeningkitty::list_family: cis
hardeningkitty::audit_time: '03:30'
hardeningkitty::keep_reports: 14
hardeningkitty::nsclient_scripts_dir: 'C:/Program Files/NSClient++/scripts'   # dépose le check
hardeningkitty::custom_lists_source: 'puppet:///modules/profile/hardeningkitty_lists'  # listes maison (optionnel)
```

La classe copie le module HardeningKitty (pas besoin d'Internet sur la cible), dépose le script
d'audit, crée la tâche planifiée quotidienne en SYSTEM, et dépose le check dans NSClient++ si
`nsclient_scripts_dir` est renseigné. `puppet parser validate manifests/init.pp` avant de pousser.

Après le premier run, le check renvoie UNKNOWN jusqu'au premier audit :
`schtasks /Run /TN "HardeningKitty audit"`.

## Précautions

- Jamais de `-Mode HailMary` sans snapshot : il applique le durcissement et peut couper le RDP.
- Les listes CIS/MSFT sont longues : compter quelques minutes par audit.
- Les rapports contiennent des valeurs de configuration système : ne pas les publier hors du LAN.
