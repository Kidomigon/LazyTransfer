# LazyTransfer v1.0

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

### Quality of Life
- **Dark Theme** — Easy on the eyes, dark WinForms UI throughout
- **USB Portable** — Settings saved as `settings.json` next to the script, travels with the drive
- **Window Position Memory** — Remembers where you left the window
- **Real-Time Log** — Scrollable log panel at the bottom of every page
- **Status Bar** — Current operation displayed at a glance
- **Admin Elevation** — Batch launcher requests admin rights automatically

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

That's it. No setup, no installer, no account.

---

## Files

```
LazyTransfer/
  LazyTransfer-GUI.ps1    # Main application (~2,200 lines)
  Start_LazyTransfer.bat   # Launcher with admin elevation
  settings.json            # Created on first run (portable)
  README.md                # You're here
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

---

## Limitations

- Windows only (PowerShell + WinForms)
- Programs install from public package repositories — custom/enterprise software won't be found
- Bookmarks are raw file copies; the target browser must be installed first
- Settings page is a stub in v1.0 (planned for v1.1)

---

## License

MIT

---

*Built by [Kidomigon](https://github.com/Kidomigon) with Claude Code.*
