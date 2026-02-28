# LazyTransfer v2.5

**The lazy way to move your stuff to a new PC.**

LazyTransfer is a portable Windows tool that scans your old computer for installed programs and user files, then reinstalls everything on a new machine with one click. No cloud accounts, no subscriptions — just drop it on a USB drive and go.

---

## Features

### Program Migration
- **Scan** — Reads installed programs from the Windows registry. Automatically filters out noise (Visual C++ redistributables, .NET runtimes, WebView2, Windows SDK, etc.)
- **Export** — Saves your program list as a portable JSON bundle file
- **Install** — Imports a bundle on the new PC and batch-installs programs using **winget** (preferred) or **Chocolatey** as fallback
- **Smart Matching** — Maps scanned program names to package manager IDs via fuzzy search and a built-in static map
- **10 App Categories** — Toggle entire categories on/off: Browsers, Dev Tools, Media, Communication, Gaming, Productivity, Utilities, Security, Runtimes, Other

### Files Migration
- **6 User Folders** — Documents, Downloads, Pictures, Videos, Music, Desktop
- **Browser Bookmarks** — Chrome, Firefox, and Edge (raw file copy)
- **Two Modes** — ZIP archive for storage, or direct copy via robocopy (multithreaded)
- **Progress Tracking** — Async folder size calculation, real-time progress bar with ETA
- **Migration Manifest** — JSON audit trail of everything that was copied

### Remote Transfer (v2.4)
- **HTTP Server Mode** — Start a web server on your old PC and download everything as a ZIP from the new PC's browser
- **Direct TCP Mode** — Binary protocol for fast LAN transfers with progress tracking
- **Shared Folder Mode** — Copy to/from UNC network shares via robocopy
- **Auto Firewall** — Temporary firewall rules added and removed automatically
- **LAN IP Detection** — Automatically finds your local network IP address

### Restore Bundle (v2.5)
- **Auto-Detect** — Point at any bundle folder and LazyTransfer identifies programs, files, system settings, and bookmarks
- **Selective Restore** — Pick exactly which folders, programs, system settings, and browsers to restore
- **Bookmark Import** — Restore Chrome, Firefox, and Edge bookmarks with automatic backup of existing ones
- **Phased Restore** — System settings first, then programs (needs network), then files (needs browsers)
- **Error Resilient** — Continues if one phase fails, reports everything at the end

### Quality of Life
- **Dark Theme** — Easy on the eyes, dark WinForms UI throughout
- **USB Portable** — Settings saved as `settings.json` next to the script, travels with the drive
- **Window Position Memory** — Remembers where you left the window
- **Real-Time Log** — Scrollable log panel at the bottom of every page
- **Status Bar** — Current operation displayed at a glance
- **Admin Elevation** — Batch launcher requests admin rights automatically
- **CLI Mode** — Full headless CLI for scripting and automation

---

## Requirements

- **Windows 10 or 11**
- **PowerShell 5.1+** (built into Windows)
- **Administrator privileges** (for program installation)
- **winget** or **Chocolatey** (LazyTransfer will prompt to install winget if missing)

---

## Quick Start

1. Copy the `LazyTransfer` folder to a USB drive (or anywhere you like)
2. On the **old PC**: double-click `Start_LazyTransfer.bat` → go to **Scan** → scan and export your program list
3. On the **new PC**: double-click `Start_LazyTransfer.bat` → go to **Install** → load your bundle → select programs → install
4. For files: go to the **Files** tab → pick folders and bookmarks → choose ZIP or direct copy → migrate
5. For LAN transfer: go to **Transfer** → start HTTP server on old PC → download from new PC
6. For full restore: go to **Restore** → point at bundle folder → detect → select what to restore → go

That's it. No setup, no installer, no account.

---

## Files

```
LazyTransfer/
  LazyTransfer-GUI.ps1          # Main application with GUI
  LazyTransfer-CLI.ps1          # Headless CLI interface
  Start_LazyTransfer.bat        # Launcher with admin elevation
  settings.json                 # Created on first run (portable)
  README.md                     # You're here
  modules/
    NetworkEngine.ps1            # Remote transfer module (HTTP, TCP, shared)
  build/
    Build-Exe.ps1                # EXE packaging pipeline
  profiles/
    developer.json               # Developer migration preset
    gamer.json                   # Gamer migration preset
    office.json                  # Office migration preset
    everything.json              # Full migration preset
```

---

## CLI Usage

```powershell
# Scan programs on this PC
.\LazyTransfer-CLI.ps1 -Action scan -ClientName "MyPC" -OutputFolder "D:\Backup"

# Install programs from a bundle
.\LazyTransfer-CLI.ps1 -Action install -BundlePath "D:\Backup\MyPC\programs.json"

# Migrate user files
.\LazyTransfer-CLI.ps1 -Action files -OutputFolder "D:\Backup" -Folders Documents,Pictures -Mode zip

# Start HTTP transfer server
.\LazyTransfer-CLI.ps1 -Action serve -OutputFolder "D:\Backup\MyPC" -Port 8642

# Download from transfer server
.\LazyTransfer-CLI.ps1 -Action receive -RemoteIP 192.168.1.100 -OutputFolder "D:\Received"

# Full restore from bundle
.\LazyTransfer-CLI.ps1 -Action restore -BundlePath "D:\Backup\MyPC"

# Restore, skip programs
.\LazyTransfer-CLI.ps1 -Action restore -BundlePath "D:\Backup\MyPC" -SkipPrograms

# Show system info
.\LazyTransfer-CLI.ps1 -Action info
```

---

## How It Works

### Scanning
LazyTransfer reads from three registry locations:
- `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*`
- `HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*`
- `HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*`

Programs without a display name or matching the noise filter are excluded. Results are categorized automatically based on name pattern matching.

### Installing
For each selected program, LazyTransfer:
1. Checks if it's already installed
2. Searches winget for a matching package
3. Falls back to Chocolatey if winget can't find it
4. Runs a silent install with progress reporting
5. Logs success or failure for each program

### Files Migration
User folders are copied using `robocopy /E /MT:4` (direct mode) or compressed to a timestamped ZIP archive. Browser bookmarks are located by their known paths and copied as raw files — no browser extension needed.

### Remote Transfer
Three modes for getting files from old PC to new PC over LAN:
- **HTTP Server**: Starts a lightweight web server — open the URL in any browser to download a ZIP of all files
- **Direct TCP**: Binary protocol that streams files directly with 64KB chunks for maximum speed
- **Shared Folder**: Uses robocopy to copy to/from Windows network shares (UNC paths)

### Restore
Point LazyTransfer at any bundle folder and it auto-detects what's inside:
- `programs.json` → program list for installation
- User folders (Documents, Pictures, etc.) → file restore via robocopy merge
- `Bookmarks/` → browser bookmark restore with `.bak` backup
- `SystemMigration/` → system settings restore

Restore runs in phases: system settings → programs → files + bookmarks.

---

## Limitations

- Windows only (PowerShell + WinForms)
- Programs install from public package repositories — custom/enterprise software won't be found
- Bookmarks are raw file copies; the target browser must be installed first
- HTTP transfer server requires admin for the firewall rule
- TCP transfer is unencrypted (designed for trusted LANs)

---

## Roadmap

| Version | Feature | Status |
|---------|---------|--------|
| v1.0 | Program scan, install, files migration, dark theme | Done |
| v2.4 | Remote transfer (HTTP, TCP, shared folder) | Done |
| v2.5 | Restore bundle (auto-detect, selective restore) | Done |
| v2.6 | System settings migration (WiFi, SSH, git, env vars) | Next |
| v2.7 | Migration profiles + settings page | Planned |
| v2.8 | EXE packaging + auto-updater | Planned |

---

## License

MIT

---

*Built by [Kidomigon](https://github.com/Kidomigon) with Claude Code.*
