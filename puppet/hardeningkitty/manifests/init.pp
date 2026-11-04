# @summary Déploie HardeningKitty (audit seul) sur une machine Windows.
#
# - Copie le module HardeningKitty (psd1, psm1, lists/) depuis files/HardeningKitty/ du module Puppet
#   (version « vendorée » : pas besoin d'accès Internet sur la machine cible).
# - Dépose le script d'audit, planifie un audit quotidien en SYSTEM (lecture seule, mode Audit).
# - Dépose, si demandé, le check Nagios dans le dossier scripts de NSClient++.
#
# Prérequis : module puppetlabs-scheduled_task (type scheduled_task) sur le Puppet Windows.
#
# Arborescence attendue du module :
#   hardeningkitty/
#     manifests/init.pp
#     files/
#       Invoke-HardeningKittyAudit.ps1
#       check_hardeningkitty.ps1
#       HardeningKitty/            <- contenu de la release : HardeningKitty.psd1, .psm1, lists/
#
# @param install_dir          Racine d'installation (slashs normaux, Puppet Windows les accepte).
# @param finding_list         Nom du CSV dans lists/ utilisé pour l'audit.
# @param schedule_audit       Planifier l'audit quotidien.
# @param audit_time           Heure de l'audit quotidien (HH:MM).
# @param keep_reports         Nombre de rapports conservés.
# @param nsclient_scripts_dir Dossier scripts de NSClient++ ; undef = ne pas déposer le check.
class hardeningkitty (
  String           $install_dir          = 'C:/ProgramData/HardeningKitty',
  String           $finding_list         = 'auto',
  String           $list_family          = 'cis',
  Boolean          $schedule_audit       = true,
  String           $audit_time           = '03:00',
  Integer          $keep_reports         = 14,
  Optional[String] $nsclient_scripts_dir = undef,
  Optional[String] $custom_lists_source  = undef,
) {

  $install_dir_win = regsubst($install_dir, '/', '\\', 'G')
  $ps_exe          = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'

  # --- arborescence ---------------------------------------------------------
  file { [$install_dir, "${install_dir}/scripts", "${install_dir}/reports"]:
    ensure => directory,
  }

  # --- listes maison (optionnel) --------------------------------------------
  # Ex. : custom_lists_source => 'puppet:///modules/profile/hardeningkitty_lists'
  # Un FindingList donné par simple nom est cherché ici avant module/lists/.
  if $custom_lists_source {
    file { "${install_dir}/custom_lists":
      ensure             => directory,
      source             => $custom_lists_source,
      recurse            => true,
      purge              => true,
      source_permissions => ignore,
      require            => File[$install_dir],
    }
  } else {
    file { "${install_dir}/custom_lists":
      ensure  => directory,
      require => File[$install_dir],
    }
  }

  # --- module HardeningKitty (vendoré) --------------------------------------
  file { "${install_dir}/module":
    ensure             => directory,
    source             => 'puppet:///modules/hardeningkitty/HardeningKitty',
    recurse            => true,
    purge              => false,
    source_permissions => ignore,
    require            => File[$install_dir],
  }

  # --- script d'audit -------------------------------------------------------
  file { "${install_dir}/scripts/Invoke-HardeningKittyAudit.ps1":
    ensure             => file,
    source             => 'puppet:///modules/hardeningkitty/Invoke-HardeningKittyAudit.ps1',
    source_permissions => ignore,
    require            => File["${install_dir}/scripts"],
  }

  # --- tâche planifiée : audit quotidien en SYSTEM ---------------------------
  if $schedule_audit {
    $audit_args = join([
      '-NoProfile -NonInteractive -ExecutionPolicy Bypass',
      "-File \"${install_dir_win}\\scripts\\Invoke-HardeningKittyAudit.ps1\"",
      "-InstallDir \"${install_dir_win}\"",
      "-FindingList \"${finding_list}\"",
      "-ListFamily \"${list_family}\"",
      "-KeepReports ${keep_reports}",
    ], ' ')

    scheduled_task { 'HardeningKitty audit':
      ensure    => present,
      enabled   => true,
      command   => $ps_exe,
      arguments => $audit_args,
      user      => 'SYSTEM',
      trigger   => {
        schedule   => 'daily',
        start_time => $audit_time,
      },
      require   => [
        File["${install_dir}/module"],
        File["${install_dir}/scripts/Invoke-HardeningKittyAudit.ps1"],
        File["${install_dir}/reports"],
      ],
    }
  }

  # --- check Nagios pour NSClient++ (optionnel) ------------------------------
  if $nsclient_scripts_dir {
    file { "${nsclient_scripts_dir}/check_hardeningkitty.ps1":
      ensure             => file,
      source             => 'puppet:///modules/hardeningkitty/check_hardeningkitty.ps1',
      source_permissions => ignore,
    }
  }
}
