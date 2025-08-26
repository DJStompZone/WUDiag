<#
.SYNOPSIS
  Diagnose and repair common Windows Update issues (DISM, SFC, optional WU reset, log collection, and scan trigger).

.DESCRIPTION
  Runs the classic integrity triad (DISM Check/Scan/RestoreHealth + SFC), with optional Windows Update component reset
  (stopping services, renaming SoftwareDistribution and Catroot2, re-registering core WU COM DLLs), optional log collection
  (CBS.log, merged WindowsUpdate.log from ETL via Get-WindowsUpdateLog, Setup logs), and an optional USO scan trigger.

  Includes checks for pending reboot state, Insider channel basics, and returns a clear summary code.

.PARAMETER ResetWU
  Perform Windows Update component reset:
  - Stop: wuauserv, bits, cryptsvc, msiserver
  - Rename %SystemRoot%\SoftwareDistribution and %SystemRoot%\System32\catroot2
  - Optionally re-register common WU COM DLLs

.PARAMETER ReRegisterCOM
  When used with -ResetWU, re-register a curated set of WU-related DLLs (quiet regsvr32). Safe but slow.

.PARAMETER CollectLogs
  Collect and export update-related logs to a timestamped folder on the Desktop:
  - CBS.log
  - Get-WindowsUpdateLog (merges ETL to a single .log)
  - Setup logs and a quick system info snapshot

.PARAMETER TriggerScan
  Attempt to trigger Windows Update scan/download/install via UsoClient (best-effort).

.PARAMETER NoDISM
  Skip DISM operations (for quick SFC-only runs).

.PARAMETER NoSFC
  Skip SFC (for quick DISM-only runs).

.PARAMETER RebootIfNeeded
  If a pending reboot is detected (or DISM/SFC request it), prompt and reboot automatically (with -ForceReboot to skip prompt).

.PARAMETER ForceReboot
  Used with -RebootIfNeeded to reboot without prompting.

.PARAMETER Source
  Optional path or WIM source for DISM /RestoreHealth, e.g. "WIM:D:\sources\install.wim:1". If omitted, uses Windows Update.

.EXAMPLE
  .\WUDiag.ps1 -ResetWU -CollectLogs -TriggerScan -RebootIfNeeded -Verbose

.NOTES
  Author: DJ Stomp <85457381+DJStompZone@users.noreply.github.com>
  License: MIT
  Repo: https://github.com/djstompzone/WUDiag

  Run as Administrator. Tested on Windows 10/11 (incl. Insider). This script aims to be safe and idempotent.

#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [switch]$ResetWU,
  [switch]$ReRegisterCOM,
  [switch]$CollectLogs,
  [switch]$TriggerScan,
  [switch]$NoDISM,
  [switch]$NoSFC,
  [switch]$RebootIfNeeded,
  [switch]$ForceReboot,
  [string]$Source
)

function Show-Banner {
   $base64 = 'G1szODs1OzE5Nm3ilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilogKG1szODs1OzE5Nm3ilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilogKG1szODs1OzE5Nm3ilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilojilogbWzM4OzU7MjA4beKWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWkxtbMzg7NTsxOTZt4paI4paI4paI4paI4paI4paI4paI4paI4paI4paI4paI4paI4paI4paI4paI4paI4paI4paIChtbMzg7NTsxOTZt4paI4paI4paI4paIG1swbSAgICAgICAbWzM4OzU7MTk2beKWiOKWiBtbMzg7NTsyMDht4paTG1swbSAgICAgICAgG1szODs1OzIwOG3ilpPilpPilpMbWzBtICAgICAgG1szODs1OzIwOG3ilpPilpMbWzBtICAgICAgICAbWzM4OzU7MjA4beKWk+KWkxtbMG0gICAgICAbWzM4OzU7MjA4beKWk+KWk+KWkxtbMG0gIBtbMzg7NTsyMDht4paT4paT4paTG1swbSAgG1szODs1OzE5Nm3ilojilogbWzBtICAgICAgG1szODs1OzE5Nm3ilojilojilojilojilogKG1szODs1OzE5Nm3ilojilojilojilogbWzBtICAbWzM4OzU7MTk2beKWiOKWiOKWiOKWiBtbMG0gIBtbMzg7NTsyMDht4paT4paT4paT4paT4paT4paTG1swbSAgG1szODs1OzIwOG3ilpPilpPilpPilpMbWzBtICAbWzM4OzU7MjA4beKWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWk+KWkxtbMG0gIBtbMzg7NTsyMDht4paT4paT4paT4paTG1swbSAgG1szODs1OzIwOG3ilpPilpPilpPilpMbWzBtICAbWzM4OzU7MjA4beKWkxtbMG0gICAbWzM4OzU7MjA4beKWk+KWkxtbMG0gICAbWzM4OzU7MjA4beKWkxtbMG0gIBtbMzg7NTsxOTZt4paI4paI4paI4paIG1swbSAgG1szODs1OzE5Nm3ilojilojilojilogKG1szODs1OzE5Nm3ilojilojilojilogbWzBtICAbWzM4OzU7MTk2beKWiOKWiBtbMzg7NTsyMDht4paT4paTG1swbSAgG1szODs1OzIwOG3ilpPilpPilpPilpPilpPilpMbWzBtICAbWzM4OzU7MjA4beKWk+KWk+KWk+KWk+KWkxtbMG0gICAgICAbWzM4OzU7MjA4beKWk+KWk+KWk+KWk+KWkxtbMG0gIBtbMzg7NTsyMDht4paT4paT4paT4paTG1swbSAgG1szODs1OzIwOG3ilpPilpPilpPilpMbWzBtICAbWzM4OzU7MjA4beKWkxtbMG0gICAgICAgIBtbMzg7NTsyMDht4paTG1swbSAgICAgICAbWzM4OzU7MTk2beKWiOKWiOKWiOKWiOKWiAobWzM4OzU7MTk2beKWiOKWiOKWiOKWiBtbMG0gIBtbMzg7NTsyMDht4paT4paT4paT4paTG1swbSAgG1szODs1OzIwOG3ilpMbWzBtICAbWzM4OzU7MjA4beKWk+KWk+KWkxtbMG0gIBtbMzg7NTsyMDht4paT4paTG1szODs1OzIxNG3ilpLilpLilpLilpLilpLilpLilpLilpIbWzBtICAbWzM4OzU7MjE0beKWkuKWkuKWkuKWkhtbMG0gIBtbMzg7NTsyMTRt4paS4paS4paS4paSG1swbSAgG1szODs1OzIxNG3ilpLilpLilpLilpIbWzBtICAbWzM4OzU7MjA4beKWkxtbMG0gIBtbMzg7NTsyMDht4paTG1swbSAgG1szODs1OzIwOG3ilpMbWzBtICAbWzM4OzU7MjA4beKWkxtbMG0gIBtbMzg7NTsyMDht4paT4paT4paT4paTG1szODs1OzE5Nm3ilojilojilojilojilojilogKG1szODs1OzE5Nm3ilojilojilojilogbWzBtICAgICAgIBtbMzg7NTsyMDht4paT4paT4paTG1swbSAgICAgG1szODs1OzIxNG3ilpLilpLilpLilpLilpLilpIbWzBtICAgICAgG1szODs1OzIxNG3ilpLilpLilpLilpLilpIbWzBtICAbWzM4OzU7MjE0beKWkuKWkuKWkuKWkuKWkhtbMG0gICAgICAbWzM4OzU7MjE0beKWkuKWkhtbMG0gIBtbMzg7NTsyMTRt4paS4paS4paSG1szODs1OzIwOG3ilpMbWzBtICAbWzM4OzU7MjA4beKWkxtbMG0gIBtbMzg7NTsyMDht4paT4paT4paT4paT4paT4paTG1szODs1OzE5Nm3ilojilojilojilogKG1szODs1OzE5Nm3ilojilogbWzM4OzU7MjA4beKWk+KWk+KWk+KWk+KWk+KWkxtbMzg7NTsyMTRt4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paSG1szODs1OzIwOG3ilpPilpPilpPilpPilpPilpMbWzM4OzU7MTk2beKWiOKWiAobWzM4OzU7MTk2beKWiBtbMzg7NTsyMDht4paT4paT4paT4paTG1szODs1OzIxNG3ilpLilpLilpLilpLilpLilpLilpLilpLilpLilpLilpLilpLilpLilpLilpIbWzM4OzU7MjI2beKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkRtbMzg7NTsyMTRt4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paS4paSG1szODs1OzIwOG3ilpPilpPilpPilpMbWzM4OzU7MTk2beKWiAobWzM4OzU7MjA4beKWk+KWk+KWkxtbMzg7NTsyMTRt4paS4paS4paS4paS4paS4paS4paSG1szODs1OzIyNm3ilpHilpHilpHilpHilpHilpHilpHilpHilpHilpHilpHilpHilpHilpHilpHilpHilpHilpEbWzBtIMKpREobWzBtIFN0b21wG1swbSAyMDI1G1swbSAbWzM4OzU7MjI2beKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkRtbMzg7NTsyMTRt4paS4paS4paS4paS4paS4paS4paSG1szODs1OzIwOG3ilpPilpPilpMKG1szODs1OzIwOG3ilpPilpMbWzM4OzU7MjE0beKWkuKWkuKWkhtbMzg7NTsyMjZt4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paRG1swbSAiTm8bWzBtIFJpZ2h0cxtbMG0gUmVzZXJ2ZWQiG1swbSAbWzM4OzU7MjI2beKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkeKWkRtbMzg7NTsyMTRt4paS4paS4paSG1szODs1OzIwOG3ilpPilpMKG1szODs1OzIwOG3ilpMbWzM4OzU7MjE0beKWkuKWkhtbMzg7NTsyMjZt4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paRG1szODs1OzIxNG3ilpLilpIbWzM4OzU7MjA4beKWkwobWzM4OzU7MjE0beKWkuKWkhtbMzg7NTsyMjZt4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paR4paRG1szODs1OzIxNG3ilpLilpIKG1swbQo='
   $BANNER = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64))
   Write-Host $BANNER
}

function Assert-Admin {
  $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
             ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin) {
    throw "This script must be run as Administrator. Right-click PowerShell and 'Run as administrator'."
  }
}

function New-LogContext {
  $ts = Get-Date -Format "yyyyMMdd-HHmmss"
  $global:WUDoctor = [ordered]@{
    StartTime     = Get-Date
    Timestamp     = $ts
    WorkRoot      = Join-Path $env:TEMP "wu-doctor-$ts"
    DesktopRoot   = Join-Path ([Environment]::GetFolderPath('Desktop')) "WU-Doctor-$ts"
    Transcript    = Join-Path ([Environment]::GetFolderPath('Desktop')) "WU-Doctor-$ts\transcript.log"
    Summary       = Join-Path ([Environment]::GetFolderPath('Desktop')) "WU-Doctor-$ts\summary.txt"
    Actions       = @()
  }
  New-Item -ItemType Directory -Force -Path $WUDoctor.WorkRoot, $WUDoctor.DesktopRoot | Out-Null
  try { Stop-Transcript | Out-Null } catch { }
  Start-Transcript -Path $WUDoctor.Transcript -Append | Out-Null
}

function Stop-LogContext {
  Stop-Transcript | Out-Null
}

function Add-Action([string]$Text) {
  $WUDoctor.Actions += ("[{0}] {1}" -f (Get-Date -Format s), $Text)
  Write-Verbose $Text
}

function Test-PendingReboot {
  $pending = $false
  $keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
  )
  foreach ($k in $keys) {
    if (Test-Path $k) { $pending = $true }
  }
  return $pending
}

function Get-InsiderInfo {
  $path = 'HKLM:\SOFTWARE\Microsoft\WindowsSelfHost\Applicability'
  if (Test-Path $path) {
    $o = Get-ItemProperty -Path $path
    [pscustomobject]@{
      BranchName      = $o.BranchName
      Ring            = $o.Ring
      ContentType     = $o.ContentType
      FlightingOwner  = $o.FlightingOwnerGUID
      InsiderFound    = $true
    }
  } else {
    [pscustomobject]@{ InsiderFound=$false }
  }
}

function Invoke-DISM {
  param(
    [switch]$CheckHealth,
    [switch]$ScanHealth,
    [switch]$RestoreHealth,
    [string]$Source
  )

  $base = "DISM /Online /Cleanup-Image"
  if ($CheckHealth)  { & cmd /c "$base /CheckHealth";  if ($LASTEXITCODE) { throw "DISM /CheckHealth failed with code $LASTEXITCODE" } }
  if ($ScanHealth)   { & cmd /c "$base /ScanHealth";   if ($LASTEXITCODE) { throw "DISM /ScanHealth failed with code $LASTEXITCODE" } }
  if ($RestoreHealth){
    $cmd = "$base /RestoreHealth"
    if ($Source) { $cmd += " /Source:$Source /LimitAccess" }
    & cmd /c $cmd
    if ($LASTEXITCODE) { throw "DISM /RestoreHealth failed with code $LASTEXITCODE" }
  }
}

function Invoke-SFC {
  & cmd /c "sfc /scannow"
  if ($LASTEXITCODE) { throw "SFC /scannow failed with code $LASTEXITCODE" }
}

function Reset-WUComponents {
  Write-Verbose "Stopping services: wuauserv, bits, cryptsvc, msiserver"
  foreach($svc in 'wuauserv','bits','cryptsvc','msiserver'){
    try { Stop-Service -Name $svc -Force -ErrorAction Stop } catch { }
  }

  $sd = Join-Path $env:SystemRoot 'SoftwareDistribution'
  $cr = Join-Path $env:SystemRoot 'System32\catroot2'
  $sd_bak = "$sd.bak.$($WUDoctor.Timestamp)"
  $cr_bak = "$cr.bak.$($WUDoctor.Timestamp)"

  if (Test-Path $sd) { Rename-Item -Path $sd -NewName (Split-Path $sd_bak -Leaf) -ErrorAction SilentlyContinue }
  if (Test-Path $cr) { Rename-Item -Path $cr -NewName (Split-Path $cr_bak -Leaf) -ErrorAction SilentlyContinue }

  if ($ReRegisterCOM) {
    Write-Verbose "Re-registering core WU DLLs (quiet)."
    $dlls = @(
      'atl.dll','urlmon.dll','mshtml.dll','shdocvw.dll','browseui.dll','jscript.dll','vbscript.dll','scrrun.dll','msxml.dll',
      'msxml3.dll','msxml6.dll','wuapi.dll','wuaueng.dll','wuaueng1.dll','wucltui.dll','wups.dll','wups2.dll','wuweb.dll',
      'qmgr.dll','qmgrprxy.dll','wucltux.dll','muweb.dll'
    )
    foreach ($dll in $dlls) {
      $path = Join-Path $env:SystemRoot "System32\$dll"
      if (Test-Path $path) {
        & regsvr32.exe /s $path | Out-Null
      }
    }
  }

  Write-Verbose "Starting services."
  foreach($svc in 'msiserver','cryptsvc','bits','wuauserv'){
    try { Start-Service -Name $svc -ErrorAction Stop } catch { }
  }
}

function Invoke-USOScan {
  $uso = Join-Path $env:SystemRoot "System32\UsoClient.exe"
  if (Test-Path $uso) {
    Write-Verbose "Triggering USO scan/download/install (best-effort)."
    try { & $uso StartScan | Out-Null } catch { }
    Start-Sleep -Seconds 2
    try { & $uso StartDownload | Out-Null } catch { }
    Start-Sleep -Seconds 2
    try { & $uso StartInstall | Out-Null } catch { }
  } else {
    Write-Warning "UsoClient not found; skipping scan trigger."
  }
}

function Collect-WULogs {
  $out = $WUDoctor.DesktopRoot
  $si  = Join-Path $out 'systeminfo.txt'
  $cbs = Join-Path $env:windir 'Logs\CBS\CBS.log'
  $setup = Join-Path $env:windir 'Panther'
  $wuMerged = Join-Path $out 'WindowsUpdate.log'

  Write-Verbose "Collecting logs to $out"
  systeminfo.exe | Out-File -Encoding utf8 $si

  if (Test-Path $cbs) { Copy-Item $cbs -Destination (Join-Path $out 'CBS.log') -Force }
  if (Test-Path $setup) { Copy-Item $setup -Destination (Join-Path $out 'Panther') -Recurse -Force -ErrorAction SilentlyContinue }

  try {
    Get-WindowsUpdateLog -LogPath $wuMerged | Out-Null
  } catch {
    Write-Warning "Get-WindowsUpdateLog failed: $_"
  }

  $ins = Get-InsiderInfo
  $ins | ConvertTo-Json -Depth 3 | Out-File -Encoding utf8 (Join-Path $out 'insider-info.json')
}

function Write-Summary {
  $pending = Test-PendingReboot
  $lines = @()
  $lines += "Windows Update Doctor — $(Get-Date -Format u)"
  $lines += "WorkRoot: $($WUDoctor.WorkRoot)"
  $lines += "Output  : $($WUDoctor.DesktopRoot)"
  $lines += "Pending Reboot: $pending"
  $lines += ""
  $lines += "Actions:"
  $lines += $WUDoctor.Actions
  $lines += ""
  $lines += "Next steps:"
  if ($pending) {
    $lines += "- A reboot is pending. Reboot, then try Windows Update again."
  } else {
    $lines += "- If updates still fail, review WindowsUpdate.log and CBS.log in the output folder."
  }
  $lines | Out-File -Encoding utf8 $WUDoctor.Summary
  Write-Host ($lines -join [Environment]::NewLine)
}

# —— Main ——
try {
  Show-Banner
  Sleep 5
  Assert-Admin
  New-LogContext
  Add-Action "Starting Windows Update diagnostics"

  $pendingBefore = Test-PendingReboot
  if ($pendingBefore) { Add-Action "Pending reboot detected BEFORE repairs." }

  if (-not $NoDISM) {
    Add-Action "Running DISM: CheckHealth"
    Invoke-DISM -CheckHealth
    Add-Action "Running DISM: ScanHealth"
    Invoke-DISM -ScanHealth
    Add-Action "Running DISM: RestoreHealth"
    Invoke-DISM -RestoreHealth -Source $Source
  } else {
    Add-Action "Skipping DISM per -NoDISM"
  }

  if (-not $NoSFC) {
    Add-Action "Running SFC /scannow"
    Invoke-SFC
  } else {
    Add-Action "Skipping SFC per -NoSFC"
  }

  if ($ResetWU) {
    Add-Action "Resetting Windows Update components"
    Reset-WUComponents
  }

  if ($CollectLogs) {
    Add-Action "Collecting logs"
    Collect-WULogs
  }

  if ($TriggerScan) {
    Add-Action "Triggering USO scan"
    Invoke-USOScan
  }

  $pendingAfter = Test-PendingReboot
  if ($pendingAfter) {
    Add-Action "Pending reboot detected AFTER repairs."
    if ($RebootIfNeeded) {
      if ($ForceReboot -or $PSCmdlet.ShouldProcess("Reboot","Restart-Computer now")) {
        Add-Action "Rebooting system immediately."
        Write-Summary
        Stop-LogContext
        Restart-Computer -Force
        return
      }
    }
  }

  Write-Summary
}
catch {
  Write-Error $_
  if ($WUDoctor -and $WUDoctor.DesktopRoot) {
    "ERROR: $($_.Exception.Message)" | Out-File -Append -Encoding utf8 $WUDoctor.Summary
  }
}
finally {
  try { Stop-LogContext } catch { }
}
