<#
.SYNOPSIS
    LazyTransfer CLI v2.5 — Headless command-line interface

.DESCRIPTION
    Perform scan, install, files migration, remote transfer, and restore operations
    without the GUI. Designed for scripting and automation.

.PARAMETER Action
    The operation to perform: scan, install, files, serve, receive, restore, info

.EXAMPLE
    .\LazyTransfer-CLI.ps1 -Action scan -ClientName "MyPC" -OutputFolder "D:\Backup"
    .\LazyTransfer-CLI.ps1 -Action install -BundlePath "D:\Backup\MyPC\programs.json"
    .\LazyTransfer-CLI.ps1 -Action files -OutputFolder "D:\Backup" -Folders Documents,Pictures
    .\LazyTransfer-CLI.ps1 -Action serve -OutputFolder "D:\Backup\MyPC" -Port 8642
    .\LazyTransfer-CLI.ps1 -Action receive -RemoteIP 192.168.1.100 -OutputFolder "D:\Received"
    .\LazyTransfer-CLI.ps1 -Action restore -BundlePath "D:\Backup\MyPC" -SkipPrograms
    .\LazyTransfer-CLI.ps1 -Action info

.NOTES
    Author: You + Claude
    Requires: Windows 10/11, PowerShell 5.1+, Admin for installs
#>

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('scan', 'install', 'files', 'serve', 'receive', 'restore', 'info')]
    [string]$Action,

    [string]$ClientName,
    [string]$OutputFolder,
    [string]$BundlePath,
    [string]$RemoteIP,
    [int]$Port = 8642,
    [string[]]$Folders = @('Documents', 'Downloads', 'Pictures', 'Videos', 'Music', 'Desktop'),
    [string[]]$Browsers = @('Chrome', 'Firefox', 'Edge'),
    [string]$Mode = 'zip',
    [switch]$SkipPrograms,
    [switch]$SkipFiles,
    [switch]$SkipSystem
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

#region Minimal function stubs (avoid GUI initialization)
# We need the engine functions but NOT the GUI. Define stubs for GUI-only functions.
$script:LogTextBox = $null
$script:StatusLabel = $null
$script:ContentPanel = $null
$script:SidebarButtons = @{}
$script:OperationInProgress = $false
$script:CancelRequested = $false
$script:LastMigrationPath = $null

# Prevent Show-MainForm from running
$script:CLIMode = $true
#endregion

# Determine script directory
$script:ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if ([string]::IsNullOrWhiteSpace($script:ScriptDir)) { $script:ScriptDir = $PWD.Path }

# Load settings infrastructure
$script:SettingsFile = Join-Path $script:ScriptDir "settings.json"
$script:AppName = "LazyTransfer"
$script:AppVersion = "2.5"
$script:CurrentPage = "CLI"

#region Core Functions (inline to avoid loading entire GUI)
$script:LogPath = $null
$script:LogMessages = New-Object System.Collections.Generic.List[string]

function Initialize-Logging {
    param(
        [Parameter(Mandatory=$true)][string]$BaseFolder,
        [Parameter(Mandatory=$true)][string]$Mode
    )
    $logDir = Join-Path $BaseFolder "Logs"
    if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
    $stamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $script:LogPath = Join-Path $logDir "${stamp}_${Mode}.log"
    try { Start-Transcript -Path $script:LogPath -Append | Out-Null } catch { }
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "[$ts][$Level] $Message"
    $script:LogMessages.Add($line)
    if ($script:LogMessages.Count -gt 100) { $script:LogMessages.RemoveAt(0) }
    if ($script:LogPath) {
        try { Add-Content -Path $script:LogPath -Value $line -Encoding UTF8 } catch { }
    }

    $color = switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color
}

function Update-StatusBar { param([string]$Message); Write-Host "STATUS: $Message" -ForegroundColor Cyan }
function Update-GuiStatus { }  # No-op in CLI mode

function Format-FileSize {
    param([Parameter(Mandatory=$true)][long]$Bytes)
    if ($Bytes -ge 1TB) { return "$([Math]::Round($Bytes / 1TB, 1)) TB" }
    elseif ($Bytes -ge 1GB) { return "$([Math]::Round($Bytes / 1GB, 1)) GB" }
    elseif ($Bytes -ge 1MB) { return "$([Math]::Round($Bytes / 1MB, 1)) MB" }
    elseif ($Bytes -ge 1KB) { return "$([Math]::Round($Bytes / 1KB, 1)) KB" }
    else { return "$Bytes B" }
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Require-AdminOrThrow {
    if (-not (Test-IsAdmin)) { throw "This operation requires Administrator privileges." }
}

function Sanitize-FileName {
    param([Parameter(Mandatory=$true)][string]$Name)
    $bad = [IO.Path]::GetInvalidFileNameChars()
    $clean = ($Name.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
    return ($clean -replace '\s+', ' ').Trim()
}

function New-ClientFolder {
    param([Parameter(Mandatory=$true)][string]$BaseFolder, [Parameter(Mandatory=$true)][string]$ClientName)
    $client = Sanitize-FileName $ClientName
    if ([string]::IsNullOrWhiteSpace($client)) { throw "Client name is required." }
    $folder = Join-Path $BaseFolder $client
    if (-not (Test-Path $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    return $folder
}

$script:UserSettings = @{ LastBundlePath=""; LastOutputFolder=""; LastClientName=""; LastFilesOutputFolder=""; FilesMigrationMode="zip"; WindowX=-1; WindowY=-1 }
function Load-UserSettings {
    if (Test-Path $script:SettingsFile) {
        try {
            $loaded = Get-Content $script:SettingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($loaded.LastBundlePath) { $script:UserSettings.LastBundlePath = $loaded.LastBundlePath }
            if ($loaded.LastOutputFolder) { $script:UserSettings.LastOutputFolder = $loaded.LastOutputFolder }
            if ($loaded.LastClientName) { $script:UserSettings.LastClientName = $loaded.LastClientName }
        } catch { }
    }
}
function Save-UserSettings {
    try { $script:UserSettings | ConvertTo-Json | Set-Content -Path $script:SettingsFile -Encoding UTF8 -Force } catch { }
}
function Update-Setting { param([string]$Key, $Value); if ($script:UserSettings.ContainsKey($Key)) { $script:UserSettings[$Key] = $Value; Save-UserSettings } }
Load-UserSettings
#endregion

#region Noise Filter
$script:NoiseNameRegexes = @(
    '(?i)\bMicrosoft Edge WebView2 Runtime\b', '(?i)\bWebView2 Runtime\b',
    '(?i)\bMicrosoft Visual C\+\+.*Redistributable\b', '(?i)\bVisual C\+\+.*Redistributable\b',
    '(?i)\bMicrosoft Windows Desktop Runtime\b', '(?i)\bWindows Desktop Runtime\b',
    '(?i)\bMicrosoft \.NET Runtime\b', '(?i)\b\.NET Runtime\b',
    '(?i)\bWindows Software Development Kit\b', '(?i)\bWindows SDK\b',
    '(?i)\bWindows Driver Kit\b', '(?i)\bWDK\b', '(?i)\bMicrosoft Update Health Tools\b'
)
function Test-IsNoiseAppName { param([Parameter(Mandatory=$true)][string]$DisplayName); foreach ($rx in $script:NoiseNameRegexes) { if ($DisplayName -match $rx) { return $true } }; return $false }
#endregion

#region Load Engine Modules
# Load the functions from the GUI file selectively via dot-sourcing a wrapper
# We source the core engine functions by defining them inline or loading from the GUI

# Source the NetworkEngine module
$modulesDir = Join-Path $script:ScriptDir "modules"
if (Test-Path (Join-Path $modulesDir "NetworkEngine.ps1")) {
    . (Join-Path $modulesDir "NetworkEngine.ps1")
}

# Extract engine regions from GUI script by matching #region/#endregion pairs.
# This avoids loading the entire GUI and is resilient to region ordering.
$guiScript = Join-Path $script:ScriptDir "LazyTransfer-GUI.ps1"
if (Test-Path $guiScript) {
    $content = Get-Content $guiScript -Raw -Encoding UTF8
    $engineRegions = @(
        'App Categories',
        'Files Migration Helpers',
        'Program Inventory',
        'Files Migration Engine',
        'Install Engine',
        'Restore Engine'
    )
    foreach ($regionName in $engineRegions) {
        $startTag = "#region $regionName"
        $endTag = "#endregion $regionName"
        $startIdx = $content.IndexOf($startTag)
        $endIdx = $content.IndexOf($endTag)
        if ($startIdx -ge 0 -and $endIdx -gt $startIdx) {
            $regionCode = $content.Substring($startIdx, $endIdx + $endTag.Length - $startIdx)
            try {
                # Use dot-sourcing (not .Invoke()) to define functions in the script scope
                . ([ScriptBlock]::Create($regionCode))
            } catch {
                Write-Host "Warning: Could not load region '$regionName': $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}
#endregion

# CLI progress callback (prints to console)
$cliProgress = {
    param($name, $current, $total, $status)
    $pct = if ($total -gt 0) { [Math]::Round(($current / $total) * 100) } else { 0 }
    Write-Host "`r[$pct%] $name - $status" -NoNewline
}

$cliFilesProgress = {
    param($folderName, $bytesCopied, $totalBytesAll, $currentStep, $totalSteps, $status)
    $pct = if ($totalBytesAll -gt 0) { [Math]::Round(($bytesCopied / $totalBytesAll) * 100) } else { [Math]::Round(($currentStep / $totalSteps) * 100) }
    Write-Host "`r[$pct%] $folderName - $status" -NoNewline
}

# Main action switch
Write-Host ""
Write-Host "LazyTransfer CLI v$($script:AppVersion)" -ForegroundColor Cyan
Write-Host "=" * 40

switch ($Action) {
    'scan' {
        if (-not $ClientName) { $ClientName = $env:COMPUTERNAME }
        if (-not $OutputFolder) { $OutputFolder = $script:ScriptDir }
        Write-Host "Scanning programs on $ClientName..."
        $folder = Build-ScanBundle -ClientName $ClientName -OutputFolder $OutputFolder
        Write-Host "`nScan complete! Saved to: $folder" -ForegroundColor Green
    }

    'install' {
        if (-not $BundlePath) { throw "Specify -BundlePath (path to programs.json)" }
        Require-AdminOrThrow
        $bundle = Get-Content $BundlePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $names = @($bundle.Programs | ForEach-Object { $_.DisplayName } | Where-Object { -not (Test-IsNoiseAppName -DisplayName $_) })
        Write-Host "Installing $($names.Count) programs..."
        $result = Install-ProgramsFromBundle -BundleJsonPath $BundlePath -SelectedDisplayNames $names -AutoInstallWinget -AutoInstallChocolatey -OnProgress $cliProgress
        Write-Host ""
        Write-Host "`nInstall complete!" -ForegroundColor Green
        $installed = @($result.Results | Where-Object { $_.Status -eq 'Installed' }).Count
        $failed = @($result.Results | Where-Object { $_.Status -eq 'Failed' }).Count
        $manual = @($result.Results | Where-Object { $_.Status -eq 'Manual' }).Count
        Write-Host "  Installed: $installed | Failed: $failed | Manual: $manual"
    }

    'files' {
        if (-not $OutputFolder) { throw "Specify -OutputFolder" }
        $clientName = if ($ClientName) { $ClientName } else { $env:COMPUTERNAME }
        Write-Host "Migrating files: $($Folders -join ', ')..."
        if ($Browsers) { Write-Host "Browsers: $($Browsers -join ', ')" }
        $result = Start-FilesMigration -ClientName $clientName -OutputFolder $OutputFolder -SelectedFolders $Folders -Mode $Mode -SelectedBrowsers $Browsers -OnProgress $cliFilesProgress
        Write-Host ""
        if ($result.Success) {
            Write-Host "`nMigration complete! Saved to: $($result.Path)" -ForegroundColor Green
        } else {
            Write-Host "`nMigration complete with errors:" -ForegroundColor Yellow
            $result.Errors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        }
        Write-Host "  Size: $(Format-FileSize $result.TotalBytes) | Files: $($result.TotalFiles) | Folders: $($result.FoldersCopied)"
    }

    'serve' {
        if (-not $OutputFolder) { throw "Specify -OutputFolder (folder to serve)" }
        $localIP = Get-PrimaryLocalIP
        Write-Host "Starting HTTP server on http://${localIP}:${Port}/" -ForegroundColor Cyan
        Write-Host "Press Ctrl+C to stop"
        Write-Host ""
        $serverProgress = {
            param($phase, $current, $total, $status)
            Write-Host "  $status"
        }
        try {
            Start-TransferServer -SourceFolder $OutputFolder -Port $Port -OnProgress $serverProgress
        } finally {
            Stop-TransferServer
            Write-Host "`nServer stopped." -ForegroundColor Yellow
        }
    }

    'receive' {
        if (-not $RemoteIP) { throw "Specify -RemoteIP (sender's IP address)" }
        if (-not $OutputFolder) { throw "Specify -OutputFolder (where to save files)" }
        Write-Host "Downloading from http://${RemoteIP}:${Port}/..."
        $recvProgress = {
            param($phase, $current, $total, $status)
            Write-Host "  $status"
        }
        $url = "http://${RemoteIP}:${Port}/"
        $result = Get-TransferFromServer -ServerUrl $url -OutputFolder $OutputFolder -OnProgress $recvProgress
        if ($result.Success) {
            Write-Host "`nDownload complete! Saved to: $($result.Path)" -ForegroundColor Green
        } else {
            Write-Host "`nDownload failed:" -ForegroundColor Red
            $result.Errors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        }
    }

    'restore' {
        if (-not $BundlePath) { throw "Specify -BundlePath (path to bundle folder)" }
        Write-Host "Detecting bundle contents..."
        $bc = Find-BundleContents -Path $BundlePath

        Write-Host "Found:" -ForegroundColor Cyan
        if ($bc.Programs.Found) { Write-Host "  Programs: $($bc.Programs.Count) apps" }
        if ($bc.Files.Found) { Write-Host "  Files: $($bc.Files.Folders.Count) folders ($(Format-FileSize $bc.Files.TotalSize))" }
        if ($bc.System.Found) { Write-Host "  System: $($bc.System.Items -join ', ')" }
        if ($bc.Bookmarks.Found) { Write-Host "  Bookmarks: $($bc.Bookmarks.Browsers -join ', ')" }
        if ($bc.Meta.SourcePC) { Write-Host "  Source PC: $($bc.Meta.SourcePC)" }
        Write-Host ""

        $selections = @{}
        if ($bc.Programs.Found -and -not $SkipPrograms) { $selections['Programs'] = $true }
        if ($bc.Files.Found -and -not $SkipFiles) {
            $selections['Files'] = @($bc.Files.Folders | ForEach-Object { $_.Name })
        }
        if ($bc.System.Found -and -not $SkipSystem) { $selections['System'] = $bc.System.Items }
        if ($bc.Bookmarks.Found -and -not $SkipFiles) { $selections['Bookmarks'] = $bc.Bookmarks.Browsers }

        if ($selections.Count -eq 0) {
            Write-Host "Nothing to restore (all phases skipped or not found)." -ForegroundColor Yellow
            return
        }

        Write-Host "Restoring..." -ForegroundColor Cyan
        $phaseCallback = { param($phase, $msg); Write-Host "`n--- Phase: $phase ---" -ForegroundColor Cyan; Write-Host "  $msg" }
        $progressCallback = { param($name, $current, $total, $status); Write-Host "  $name - $status" }

        $result = Start-FullRestore -BundlePath $BundlePath -BundleContents $bc -Selections $selections -OnPhaseChange $phaseCallback -OnProgress $progressCallback

        Write-Host "`n=== Restore Summary ===" -ForegroundColor Green
        if ($result.Files) { Write-Host "  Files: $($result.Files.FoldersRestored) folders restored" }
        if ($result.Programs) {
            $inst = @($result.Programs.Results | Where-Object { $_.Status -eq 'Installed' }).Count
            Write-Host "  Programs: $inst installed"
        }
        if ($result.Errors.Count -gt 0) {
            Write-Host "  Errors:" -ForegroundColor Yellow
            $result.Errors | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
        }
    }

    'info' {
        Write-Host "LazyTransfer v$($script:AppVersion)" -ForegroundColor Cyan
        Write-Host "  Computer: $env:COMPUTERNAME"
        Write-Host "  User: $env:USERNAME"
        Write-Host "  Admin: $(Test-IsAdmin)"
        Write-Host "  PowerShell: $($PSVersionTable.PSVersion)"
        Write-Host "  Script Dir: $script:ScriptDir"
        try {
            $winget = & winget --version 2>$null
            Write-Host "  Winget: $winget"
        } catch { Write-Host "  Winget: Not installed" }
        try {
            $choco = & choco --version 2>$null
            Write-Host "  Chocolatey: $choco"
        } catch { Write-Host "  Chocolatey: Not installed" }
        $localIP = Get-PrimaryLocalIP
        Write-Host "  Local IP: $localIP"
    }
}
