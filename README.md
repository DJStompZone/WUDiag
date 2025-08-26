# WUDiag

**Windows Update Diagnostic (WUDiag)** is a PowerShell diagnostic and repair tool for fixing failed Windows Updates (including Insider builds).
It automates the classic **DISM / SFC integrity triad**, with optional Windows Update reset, log collection, Insider info snapshot, and on-demand update scans.

## ✨ Features

* **DISM + SFC runs**: `/CheckHealth`, `/ScanHealth`, `/RestoreHealth`, then `sfc /scannow`
* **Optional Windows Update component reset** (with optional COM DLL re-register)
* **Collect logs** (CBS.log, merged WindowsUpdate.log, Panther setup logs, Insider channel info)
* **Detect pending reboot state** before and after repairs
* **Trigger USO scan** (`UsoClient StartScan/Download/Install`)
* **Safe, idempotent service handling** (BITS, WUAUSERV, CryptSvc, MSI)
* **Verbose transcript + summary report** in a timestamped Desktop folder

## 📦 Installation

Clone or copy the script to a location under your profile, e.g.:

```powershell
mkdir "$Env:USERPROFILE\powershell"
cd "$Env:USERPROFILE\powershell"
iwr -useb https://example.com/WUDiag.ps1 -o WUDiag.ps1
```

## 🚀 Usage

Run PowerShell as **Administrator**.
Example basic run (DISM + SFC):

```powershell
& "$Env:USERPROFILE\powershell\WUDiag.ps1"
```

### Common Options

```powershell
# Run full DISM + SFC + reset Windows Update components + collect logs
& "$Env:USERPROFILE\powershell\WUDiag.ps1" -ResetWU -CollectLogs

# Run with log collection + trigger scan + reboot if needed
& "$Env:USERPROFILE\powershell\WUDiag.ps1" -CollectLogs -TriggerScan -RebootIfNeeded

# Run everything, append output to a logfile
& "$Env:USERPROFILE\powershell\WUDiag.ps1" -ResetWU -CollectLogs -TriggerScan -RebootIfNeeded -Verbose 2>&1 |
  Tee-Object -FilePath "$Env:USERPROFILE\powershell\WUDiag.log" -Append
```

### Parameters

| Parameter         | Description                                                                                |
| ----------------- | ------------------------------------------------------------------------------------------ |
| `-ResetWU`        | Reset Windows Update components (stops services, renames SoftwareDistribution + Catroot2). |
| `-ReRegisterCOM`  | (Optional) When used with `-ResetWU`, re-register core WU-related DLLs.                    |
| `-CollectLogs`    | Collect CBS.log, merged WindowsUpdate.log, Panther logs, Insider info, and systeminfo.     |
| `-TriggerScan`    | Trigger update scan/download/install using UsoClient.                                      |
| `-NoDISM`         | Skip DISM (for SFC-only runs).                                                             |
| `-NoSFC`          | Skip SFC (for DISM-only runs).                                                             |
| `-RebootIfNeeded` | If reboot required, prompt and reboot automatically.                                       |
| `-ForceReboot`    | Used with `-RebootIfNeeded` to reboot immediately without prompt.                          |
| `-Source`         | Provide DISM source (e.g. `"WIM:D:\sources\install.wim:1"`) instead of WU.                 |

## 📝 Output

* Creates a timestamped folder on your **Desktop**:
  `WU-Doctor-YYYYMMDD-HHMMSS`
* Inside:

  * `transcript.log` — full verbose run log
  * `summary.txt` — concise status + next steps
  * `CBS.log`, `WindowsUpdate.log`, setup logs (if collected)
  * `insider-info.json` — Insider branch/ring info

## ⚡ Requirements

* Windows 10 or 11 (including Insider builds)
* Run as **Administrator**
* Internet access (unless you use `-Source` for DISM repair)

## ⚠️ Disclaimer

This script makes direct changes to Windows Update components. It’s designed to be safe and idempotent, but **use at your own risk**. Always review logs before opening bug reports with Microsoft.

## 📄 License

MIT — see LICENSE for details.
Author: **DJ Stomp** ([85457381+DJStompZone@users.noreply.github.com](mailto:85457381+DJStompZone@users.noreply.github.com))
Repo: [github.com/djstompzone/WUDiag](https://github.com/djstompzone/WUDiag)
