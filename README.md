# 🛠 Winget Silent Upgrade Script

A PowerShell script for **automating silent package upgrades on Windows 10/11 workstations** using [WinGet](https://learn.microsoft.com/en-us/windows/package-manager/).  
Designed for IT admins, developers, and power users who want **hands-off, reliable updates** without UI popups or interruptions.

---

## ✨ Features

- 🔄 **Two-pass upgrade strategy** – runs a bulk upgrade, then retries package-by-package if needed.  
- 🛡 **Single-instance guard** – prevents overlapping runs.  
- 📋 **Robust logging with rotation** – customizable log path & size, automatically rotates old logs.  
- ⏱ **Timeout enforcement** – kills long-running WinGet commands to avoid hangs.  
- 🛠 **Smart source handling** – heals the `winget` source and removes `msstore` to reduce noise.  
- 📌 **Optional pinned package upgrades** – include pinned packages if specified.  
- ⚡ **Silent execution** – no UI popups or prompts.

---

## ⚙️ Parameters

| Parameter          | Default                 | Description |
|--------------------|-------------------------|-------------|
| `LogPath`          | `C:\Utils\winget.txt`  | Path to the log file. |
| `MaxLogSizeMB`     | `10`                   | Maximum log size before rotation (in MB). |
| `WingetTimeoutSec` | `1800` (30 min)        | Timeout per WinGet command. |
| `IncludePinned`    | `False`                | Whether to include pinned packages in upgrades. |

---

## 🚀 Usage

### Run with defaults
```powershell
.\winget-upgrade.ps1
