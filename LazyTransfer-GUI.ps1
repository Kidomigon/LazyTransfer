<#
.SYNOPSIS
    LazyTransfer v1.0 — Program Migration Tool for Windows 10/11

.DESCRIPTION
    Scan programs from old PC, install on new PC with one click.
    Built for USB portability.

.NOTES
    Author: You + Claude
    Requires: Windows 10/11, PowerShell 5.1+, Admin for installs
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load assemblies first
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

#region Custom Controls
Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

public class DwmHelper {
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    public static void EnableDarkTitleBar(IntPtr handle) {
        try {
            int value = 1;
            DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
        } catch { }
    }
}

public class ModernButton : Button {
    private Color _normalColor = Color.FromArgb(42, 44, 52);
    private Color _hoverColor  = Color.FromArgb(52, 55, 65);
    private Color _pressColor  = Color.FromArgb(35, 37, 44);
    private Color _currentColor;
    private int _radius = 8;

    public ModernButton() {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        _currentColor = _normalColor;
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        ForeColor = Color.FromArgb(235, 235, 242);
        Cursor = Cursors.Hand;
        Font = new Font("Segoe UI", 10f);
    }

    public Color NormalColor {
        get { return _normalColor; }
        set { _normalColor = value; _currentColor = value; Invalidate(); }
    }
    public Color HoverColor {
        get { return _hoverColor; }
        set { _hoverColor = value; Invalidate(); }
    }
    public Color PressColor {
        get { return _pressColor; }
        set { _pressColor = value; Invalidate(); }
    }
    public int Radius {
        get { return _radius; }
        set { _radius = value; Invalidate(); }
    }

    protected override void OnMouseEnter(EventArgs e) {
        base.OnMouseEnter(e);
        _currentColor = _hoverColor;
        Invalidate();
    }
    protected override void OnMouseLeave(EventArgs e) {
        base.OnMouseLeave(e);
        _currentColor = _normalColor;
        Invalidate();
    }
    protected override void OnMouseDown(MouseEventArgs e) {
        base.OnMouseDown(e);
        _currentColor = _pressColor;
        Invalidate();
    }
    protected override void OnMouseUp(MouseEventArgs e) {
        base.OnMouseUp(e);
        _currentColor = ClientRectangle.Contains(e.Location) ? _hoverColor : _normalColor;
        Invalidate();
    }
    protected override void OnEnabledChanged(EventArgs e) {
        base.OnEnabledChanged(e);
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e) {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        var r = ClientRectangle;
        r.Inflate(-1, -1);
        int d = _radius * 2;
        if (d > r.Height) d = r.Height;

        using (var path = new GraphicsPath()) {
            path.AddArc(r.Left, r.Top, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();

            var fillColor = Enabled ? _currentColor : Color.FromArgb(35, 37, 42);
            using (var brush = new SolidBrush(fillColor))
                g.FillPath(brush, path);
        }

        var textColor = Enabled ? ForeColor : Color.FromArgb(90, 93, 105);
        using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
        using (var textBrush = new SolidBrush(textColor))
            g.DrawString(Text, Font, textBrush, ClientRectangle, sf);
    }
}

public class ModernProgressBar : Control {
    private int _value = 0;
    private int _maximum = 100;
    private Color _barColor = Color.FromArgb(52, 211, 153);
    private Color _barEndColor = Color.Empty;
    private Color _trackColor = Color.FromArgb(34, 36, 42);
    private int _radius = 5;

    public ModernProgressBar() {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        Size = new Size(200, 14);
        DoubleBuffered = true;
    }

    public int Value {
        get { return _value; }
        set { _value = Math.Max(0, Math.Min(value, _maximum)); Invalidate(); }
    }
    public int Maximum {
        get { return _maximum; }
        set { _maximum = Math.Max(1, value); Invalidate(); }
    }
    public Color BarColor {
        get { return _barColor; }
        set { _barColor = value; Invalidate(); }
    }
    public Color BarEndColor {
        get { return _barEndColor; }
        set { _barEndColor = value; Invalidate(); }
    }
    public Color TrackColor {
        get { return _trackColor; }
        set { _trackColor = value; Invalidate(); }
    }
    public int Radius {
        get { return _radius; }
        set { _radius = value; Invalidate(); }
    }
    public ProgressBarStyle Style { get; set; }

    protected override void OnPaint(PaintEventArgs e) {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;

        var r = ClientRectangle;
        int d = Math.Min(_radius * 2, r.Height);

        // Track
        using (var trackPath = MakeRoundedPath(r, d))
        using (var trackBrush = new SolidBrush(_trackColor))
            g.FillPath(trackBrush, trackPath);

        // Fill
        double pct = _maximum > 0 ? (double)_value / _maximum : 0;
        int fillWidth = (int)(r.Width * pct);
        if (fillWidth > d && pct > 0) {
            var fillRect = new Rectangle(r.Left, r.Top, fillWidth, r.Height);
            using (var fillPath = MakeRoundedPath(fillRect, d)) {
                Color endColor = _barEndColor.IsEmpty ? _barColor : _barEndColor;
                using (var fillBrush = new LinearGradientBrush(fillRect, _barColor, endColor, LinearGradientMode.Horizontal))
                    g.FillPath(fillBrush, fillPath);
            }
        }
    }

    private GraphicsPath MakeRoundedPath(Rectangle r, int d) {
        var path = new GraphicsPath();
        path.AddArc(r.Left, r.Top, d, d, 180, 90);
        path.AddArc(r.Right - d, r.Top, d, d, 270, 90);
        path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        path.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }
}

public class ModernCard : Panel {
    private int _radius = 10;
    private Color _borderColor = Color.FromArgb(48, 50, 58);
    private Color _fillColor = Color.FromArgb(34, 36, 42);
    private string _headerText = "";
    private Color _headerColor = Color.FromArgb(130, 133, 145);
    private Font _headerFont;

    public ModernCard() {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                 ControlStyles.SupportsTransparentColor, true);
        BackColor = Color.Transparent;
        _headerFont = new Font("Segoe UI", 10f, FontStyle.Bold);
        Padding = new Padding(12, 32, 12, 12);
    }

    public int Radius {
        get { return _radius; }
        set { _radius = value; Invalidate(); }
    }
    public Color BorderColor {
        get { return _borderColor; }
        set { _borderColor = value; Invalidate(); }
    }
    public Color FillColor {
        get { return _fillColor; }
        set { _fillColor = value; Invalidate(); }
    }
    public string HeaderText {
        get { return _headerText; }
        set { _headerText = value; Invalidate(); }
    }
    public Color HeaderColor {
        get { return _headerColor; }
        set { _headerColor = value; Invalidate(); }
    }

    protected override void OnPaint(PaintEventArgs e) {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        var r = ClientRectangle;
        r.Inflate(-1, -1);
        int d = _radius * 2;

        using (var path = new GraphicsPath()) {
            path.AddArc(r.Left, r.Top, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();

            using (var fill = new SolidBrush(_fillColor))
                g.FillPath(fill, path);
            using (var pen = new Pen(_borderColor, 1f))
                g.DrawPath(pen, path);
        }

        if (!string.IsNullOrEmpty(_headerText)) {
            using (var brush = new SolidBrush(_headerColor))
                g.DrawString(_headerText, _headerFont, brush, 14, 10);
        }
    }

    protected override void OnPaintBackground(PaintEventArgs e) {
        // Skip default background painting
    }
}
"@ -ErrorAction SilentlyContinue
#endregion Custom Controls

#region Icon
function Get-AppIcon {
    # Use Windows built-in icon (Application icon from shell32.dll index 2)
    # Alternative indices: 3=folder, 15=computer, 43=star, 144=users, 208=package
    try {
        Add-Type -TypeDefinition @"
        using System;
        using System.Drawing;
        using System.Runtime.InteropServices;
        public class IconExtractor {
            [DllImport("shell32.dll", CharSet = CharSet.Auto)]
            public static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool DestroyIcon(IntPtr hIcon);
        }
"@ -ErrorAction SilentlyContinue
        $iconHandle = [IconExtractor]::ExtractIcon([IntPtr]::Zero, "shell32.dll", 13)
        if ($iconHandle -ne [IntPtr]::Zero) {
            # Clone the icon so we can release the native handle
            $tempIcon = [System.Drawing.Icon]::FromHandle($iconHandle)
            $clonedIcon = [System.Drawing.Icon]$tempIcon.Clone()
            $tempIcon.Dispose()
            [IconExtractor]::DestroyIcon($iconHandle) | Out-Null
            return $clonedIcon
        }
    } catch { }
    return [System.Drawing.SystemIcons]::Application
}
#endregion Icon


#region Global Settings
$script:AppName = "LazyTransfer"
$script:AppVersion = "2.5"
$script:CurrentPage = "Scan"

$script:NoiseNameRegexes = @(
    '(?i)\bMicrosoft Edge WebView2 Runtime\b',
    '(?i)\bWebView2 Runtime\b',
    '(?i)\bMicrosoft Visual C\+\+.*Redistributable\b',
    '(?i)\bVisual C\+\+.*Redistributable\b',
    '(?i)\bMicrosoft Windows Desktop Runtime\b',
    '(?i)\bWindows Desktop Runtime\b',
    '(?i)\bMicrosoft \.NET Runtime\b',
    '(?i)\b\.NET Runtime\b',
    '(?i)\bWindows Software Development Kit\b',
    '(?i)\bWindows SDK\b',
    '(?i)\bWindows Driver Kit\b',
    '(?i)\bWDK\b',
    '(?i)\bMicrosoft Update Health Tools\b'
)

function Test-IsNoiseAppName {
    param([Parameter(Mandatory=$true)][string]$DisplayName)
    foreach ($rx in $script:NoiseNameRegexes) {
        if ($DisplayName -match $rx) { return $true }
    }
    return $false
}
#endregion Global Settings

#region User Settings (USB Portable)
# Settings are saved next to the script so they travel with the USB drive
$script:ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if ([string]::IsNullOrWhiteSpace($script:ScriptDir)) { $script:ScriptDir = $PWD.Path }
$script:SettingsFile = Join-Path $script:ScriptDir "settings.json"
$script:ModulesDir = Join-Path $script:ScriptDir "modules"

# Load modules
if (Test-Path (Join-Path $script:ModulesDir "NetworkEngine.ps1")) {
    . (Join-Path $script:ModulesDir "NetworkEngine.ps1")
}

$script:UserSettings = @{
    LastBundlePath = ""
    LastOutputFolder = ""
    LastClientName = ""
    LastFilesOutputFolder = ""
    FilesMigrationMode = "zip"
    WindowX = -1
    WindowY = -1
}

function Load-UserSettings {
    if (Test-Path $script:SettingsFile) {
        try {
            $loaded = Get-Content $script:SettingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($loaded.LastBundlePath) { $script:UserSettings.LastBundlePath = $loaded.LastBundlePath }
            if ($loaded.LastOutputFolder) { $script:UserSettings.LastOutputFolder = $loaded.LastOutputFolder }
            if ($loaded.LastClientName) { $script:UserSettings.LastClientName = $loaded.LastClientName }
            if ($loaded.LastFilesOutputFolder) { $script:UserSettings.LastFilesOutputFolder = $loaded.LastFilesOutputFolder }
            if ($loaded.FilesMigrationMode) { $script:UserSettings.FilesMigrationMode = $loaded.FilesMigrationMode }
            if ($loaded.WindowX -ne $null) { $script:UserSettings.WindowX = $loaded.WindowX }
            if ($loaded.WindowY -ne $null) { $script:UserSettings.WindowY = $loaded.WindowY }
        } catch {
            Write-Host "Warning: Could not load settings file. Using defaults."
        }
    }
}

function Save-UserSettings {
    try {
        $script:UserSettings | ConvertTo-Json | Set-Content -Path $script:SettingsFile -Encoding UTF8 -Force
    } catch {
        Write-Host "Warning: Could not save settings to $($script:SettingsFile)"
    }
}

function Update-Setting {
    param([string]$Key, $Value)
    if ($script:UserSettings.ContainsKey($Key)) {
        $script:UserSettings[$Key] = $Value
        Save-UserSettings
    }
}

# Load settings on script start
Load-UserSettings
#endregion User Settings

#region App Categories
$script:AppCategories = @{
    'Browsers' = @{
        Icon = '[WWW]'
        Color = 'AccentBlue'
        Patterns = @(
            '(?i)\b(Chrome|Firefox|Edge|Opera|Brave|Vivaldi|Safari|Waterfox|Tor Browser|Chromium)\b',
            '(?i)\bGoogle Chrome\b',
            '(?i)\bMozilla Firefox\b',
            '(?i)\bMicrosoft Edge\b'
        )
    }
    'DevTools' = @{
        Icon = '[DEV]'
        Color = 'AccentGreen'
        Patterns = @(
            '(?i)\b(Visual Studio|VS Code|VSCode|Code - OSS|JetBrains|IntelliJ|PyCharm|WebStorm|PhpStorm|Rider|CLion|DataGrip|GoLand)\b',
            '(?i)\b(Git|GitHub|GitLab|Sourcetree|Fork|GitKraken)\b',
            '(?i)\b(Node|npm|yarn|Python|Ruby|Java|JDK|JRE|Go|Rust|Golang)\b',
            '(?i)\b(Docker|Kubernetes|Podman|Vagrant|VirtualBox|VMware)\b',
            '(?i)\b(Postman|Insomnia|Fiddler|Wireshark)\b',
            '(?i)\b(Sublime Text|Atom|Notepad\+\+|Vim|Neovim|Emacs)\b',
            '(?i)\b(PowerShell|Terminal|Windows Terminal|iTerm|Hyper)\b',
            '(?i)\b(MySQL|PostgreSQL|MongoDB|Redis|SQL Server|HeidiSQL|DBeaver|pgAdmin)\b',
            '(?i)\b(Android Studio|Xcode|Flutter|React Native)\b',
            '(?i)\b(FileZilla|WinSCP|PuTTY|mRemoteNG)\b'
        )
    }
    'Media' = @{
        Icon = '[MED]'
        Color = 'AccentOrange'
        Patterns = @(
            '(?i)\b(VLC|Media Player|MPC-HC|MPC-BE|PotPlayer|KMPlayer|GOM Player|MPV)\b',
            '(?i)\b(Spotify|iTunes|Apple Music|Amazon Music|Deezer|Tidal|Foobar2000|AIMP|Winamp)\b',
            '(?i)\b(Audacity|Adobe Audition|FL Studio|Ableton|Reaper|GarageBand)\b',
            '(?i)\b(OBS|OBS Studio|Streamlabs|XSplit|Bandicam|Camtasia|ScreenPal)\b',
            '(?i)\b(Plex|Kodi|Jellyfin|Emby)\b',
            '(?i)\b(HandBrake|FFmpeg|DaVinci Resolve|Premiere|Final Cut|Filmora|Kdenlive)\b',
            '(?i)\b(YouTube|Netflix|Twitch)\b'
        )
    }
    'Graphics' = @{
        Icon = '[GFX]'
        Color = 'AccentOrange'
        Patterns = @(
            '(?i)\b(Photoshop|GIMP|Paint\.NET|Krita|Affinity Photo|Pixelmator)\b',
            '(?i)\b(Illustrator|Inkscape|Affinity Designer|CorelDRAW|Canva)\b',
            '(?i)\b(Figma|Sketch|Adobe XD|Lunacy|Penpot)\b',
            '(?i)\b(Blender|Maya|3ds Max|Cinema 4D|ZBrush|SketchUp)\b',
            '(?i)\b(Lightroom|Capture One|darktable|RawTherapee)\b',
            '(?i)\b(IrfanView|XnView|FastStone|ShareX|Greenshot|Snagit)\b'
        )
    }
    'Office' = @{
        Icon = '[DOC]'
        Color = 'AccentBlue'
        Patterns = @(
            '(?i)\b(Microsoft Office|Office 365|Microsoft 365|Word|Excel|PowerPoint|Outlook|OneNote|Access)\b',
            '(?i)\b(LibreOffice|OpenOffice|WPS Office|FreeOffice|OnlyOffice)\b',
            '(?i)\b(Adobe Acrobat|PDF|Foxit|Sumatra|Nitro PDF)\b',
            '(?i)\b(Notion|Obsidian|Evernote|Joplin|Standard Notes)\b',
            '(?i)\b(Google Docs|Google Sheets|Google Slides)\b'
        )
    }
    'Communication' = @{
        Icon = '[MSG]'
        Color = 'AccentGreen'
        Patterns = @(
            '(?i)\b(Discord|Slack|Microsoft Teams|Zoom|Skype|WebEx|Google Meet)\b',
            '(?i)\b(Telegram|WhatsApp|Signal|Messenger|Viber|WeChat|Line)\b',
            '(?i)\b(Thunderbird|Mailspring|eM Client|Mailbird)\b',
            '(?i)\b(Loom|Krisp|Around)\b'
        )
    }
    'Utilities' = @{
        Icon = '[UTL]'
        Color = 'AccentGray'
        Patterns = @(
            '(?i)\b(7-Zip|WinRAR|WinZip|PeaZip|Bandizip)\b',
            '(?i)\b(CCleaner|BleachBit|Wise|IObit|Glary)\b',
            '(?i)\b(Everything|Listary|Wox|PowerToys|AutoHotkey)\b',
            '(?i)\b(Rufus|Etcher|Ventoy|UNetbootin)\b',
            '(?i)\b(TreeSize|WizTree|WinDirStat|SpaceSniffer)\b',
            '(?i)\b(Recuva|TestDisk|PhotoRec|Disk Drill)\b',
            '(?i)\b(HWiNFO|CPU-Z|GPU-Z|CrystalDiskInfo|Speccy)\b',
            '(?i)\b(TeamViewer|AnyDesk|Parsec|RustDesk)\b',
            '(?i)\b(f\.lux|Flux|LightBulb|Night Light)\b',
            '(?i)\b(Revo Uninstaller|Geek Uninstaller|Bulk Crap)\b'
        )
    }
    'Gaming' = @{
        Icon = '[GAM]'
        Color = 'AccentRed'
        Patterns = @(
            '(?i)\b(Steam|Epic Games|GOG Galaxy|Origin|EA App|Ubisoft Connect|Battle\.net|Blizzard)\b',
            '(?i)\b(Xbox|Game Pass|PlayStation|GeForce NOW|Moonlight)\b',
            '(?i)\b(MSI Afterburner|RTSS|RivaTuner|FRAPS|Razer|Logitech G|Corsair iCUE)\b',
            '(?i)\b(Discord|Overwolf|Medal\.tv|Plays\.tv)\b',
            '(?i)\b(RetroArch|Dolphin|PCSX2|RPCS3|Yuzu|Ryujinx|MAME|PPSSPP)\b'
        )
    }
    'Security' = @{
        Icon = '[SEC]'
        Color = 'AccentRed'
        Patterns = @(
            '(?i)\b(Antivirus|Anti-Virus|Malwarebytes|Norton|McAfee|Kaspersky|Avast|AVG|Bitdefender|ESET|Avira|Sophos|Trend Micro)\b',
            '(?i)\b(Windows Defender|Windows Security|Microsoft Defender)\b',
            '(?i)\b(NordVPN|ExpressVPN|Surfshark|ProtonVPN|Mullvad|Private Internet|CyberGhost|Windscribe)\b',
            '(?i)\b(1Password|LastPass|Bitwarden|Dashlane|KeePass|Enpass)\b',
            '(?i)\b(VeraCrypt|Cryptomator|Boxcryptor)\b',
            '(?i)\b(Firewall|GlassWire|TinyWall|Simplewall)\b'
        )
    }
    'Cloud' = @{
        Icon = '[CLD]'
        Color = 'AccentBlue'
        Patterns = @(
            '(?i)\b(Dropbox|Google Drive|OneDrive|iCloud|Box|pCloud|MEGA|Sync\.com)\b',
            '(?i)\b(Nextcloud|ownCloud|Syncthing|Resilio)\b',
            '(?i)\b(AWS|Azure|Google Cloud|Heroku|DigitalOcean)\b'
        )
    }
}

function Get-AppCategory {
    param([Parameter(Mandatory=$true)][string]$DisplayName)

    foreach ($category in $script:AppCategories.Keys) {
        foreach ($pattern in $script:AppCategories[$category].Patterns) {
            if ($DisplayName -match $pattern) {
                return $category
            }
        }
    }
    return 'Other'
}

function Get-CategoryInfo {
    param([Parameter(Mandatory=$true)][string]$Category)

    if ($script:AppCategories.ContainsKey($Category)) {
        return $script:AppCategories[$Category]
    }
    return @{ Icon = '[???]'; Color = 'AccentGray'; Patterns = @() }
}
#endregion App Categories

#region Theme Colors
$script:Colors = @{
    WindowBg        = [System.Drawing.Color]::FromArgb(18, 18, 22)
    SidebarBg       = [System.Drawing.Color]::FromArgb(14, 14, 18)
    ContentBg       = [System.Drawing.Color]::FromArgb(26, 28, 32)
    CardBg          = [System.Drawing.Color]::FromArgb(34, 36, 42)
    TextPrimary     = [System.Drawing.Color]::FromArgb(235, 235, 242)
    TextSecondary   = [System.Drawing.Color]::FromArgb(130, 133, 145)
    TextDark        = [System.Drawing.Color]::FromArgb(18, 18, 22)
    AccentBlue      = [System.Drawing.Color]::FromArgb(56, 152, 255)
    AccentGreen     = [System.Drawing.Color]::FromArgb(52, 211, 153)
    AccentOrange    = [System.Drawing.Color]::FromArgb(251, 146, 60)
    AccentGray      = [System.Drawing.Color]::FromArgb(90, 93, 105)
    AccentRed       = [System.Drawing.Color]::FromArgb(248, 113, 113)
    AccentPink      = [System.Drawing.Color]::FromArgb(244, 114, 182)
    AccentLightBlue = [System.Drawing.Color]::FromArgb(96, 165, 250)
    ButtonBg        = [System.Drawing.Color]::FromArgb(42, 44, 52)
    ButtonHover     = [System.Drawing.Color]::FromArgb(52, 55, 65)
    InputBg         = [System.Drawing.Color]::FromArgb(22, 24, 28)
    InputBorder     = [System.Drawing.Color]::FromArgb(52, 55, 65)
    BorderColor     = [System.Drawing.Color]::FromArgb(48, 50, 58)
    SidebarHover    = [System.Drawing.Color]::FromArgb(28, 30, 36)
    SidebarActive   = [System.Drawing.Color]::FromArgb(32, 34, 40)
}
#endregion Theme Colors

#region Logging
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

    if ($null -ne $script:LogTextBox -and -not $script:LogTextBox.IsDisposed) {
        try {
            # Trim log if it gets too long (prevent silent truncation at 32K default)
            if ($script:LogTextBox.TextLength -gt 30000) {
                $script:LogTextBox.Text = $script:LogTextBox.Text.Substring($script:LogTextBox.TextLength - 20000)
            }
            $script:LogTextBox.AppendText($line + "`r`n")
            $script:LogTextBox.SelectionStart = $script:LogTextBox.TextLength
            $script:LogTextBox.ScrollToCaret()
        } catch { }
    }
    
    Update-GuiStatus
}

function Update-StatusBar {
    param([string]$Message)
    if ($null -ne $script:StatusLabel -and -not $script:StatusLabel.IsDisposed) { $script:StatusLabel.Text = $Message }
    Update-GuiStatus
}
#endregion Logging

#region Helpers
function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Require-AdminOrThrow {
    if (-not (Test-IsAdmin)) {
        throw "Install requires Administrator. Right-click and Run as Administrator."
    }
}

function Sanitize-FileName {
    param([Parameter(Mandatory=$true)][string]$Name)
    $bad = [IO.Path]::GetInvalidFileNameChars()
    $clean = ($Name.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
    return ($clean -replace '\s+', ' ').Trim()
}

function New-ClientFolder {
    param(
        [Parameter(Mandatory=$true)][string]$BaseFolder,
        [Parameter(Mandatory=$true)][string]$ClientName
    )
    $client = Sanitize-FileName $ClientName
    if ([string]::IsNullOrWhiteSpace($client)) { throw "Client name is required." }
    $folder = Join-Path $BaseFolder $client
    if (-not (Test-Path $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    return $folder
}

function Open-WebSearch {
    param([Parameter(Mandatory=$true)][string]$Query)
    $q = [Uri]::EscapeDataString($Query)
    $url = "https://www.bing.com/search?q=$q+download"
    try { Start-Process $url | Out-Null } catch { Write-Log "Could not open browser: $($_.Exception.Message)" -Level 'WARN' }
}

function Update-GuiStatus {
    [System.Windows.Forms.Application]::DoEvents()
}
#endregion Helpers

#region Files Migration Helpers

function Expand-ZipSafe {
    <#
    .SYNOPSIS
        Extract a ZIP archive with path traversal (ZIP slip) protection.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$ZipPath,
        [Parameter(Mandatory=$true)][string]$DestinationFolder
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $resolvedDest = [System.IO.Path]::GetFullPath($DestinationFolder)
    if (-not (Test-Path $resolvedDest)) { New-Item -Path $resolvedDest -ItemType Directory -Force | Out-Null }
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        foreach ($entry in $archive.Entries) {
            if ([string]::IsNullOrWhiteSpace($entry.FullName)) { continue }
            $entryDest = [System.IO.Path]::GetFullPath((Join-Path $DestinationFolder $entry.FullName))
            if (-not $entryDest.StartsWith($resolvedDest)) {
                Write-Log "ZIP slip blocked: $($entry.FullName)" -Level 'ERROR'
                continue
            }
            if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) {
                if (-not (Test-Path $entryDest)) { New-Item -Path $entryDest -ItemType Directory -Force | Out-Null }
            } else {
                $entryDir = Split-Path -Parent $entryDest
                if (-not (Test-Path $entryDir)) { New-Item -Path $entryDir -ItemType Directory -Force | Out-Null }
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $entryDest, $true)
            }
        }
    } finally {
        $archive.Dispose()
    }
}

function Format-FileSize {
    param([Parameter(Mandatory=$true)][long]$Bytes)
    if ($Bytes -ge 1TB) { return "$([Math]::Round($Bytes / 1TB, 1)) TB" }
    elseif ($Bytes -ge 1GB) { return "$([Math]::Round($Bytes / 1GB, 1)) GB" }
    elseif ($Bytes -ge 1MB) { return "$([Math]::Round($Bytes / 1MB, 1)) MB" }
    elseif ($Bytes -ge 1KB) { return "$([Math]::Round($Bytes / 1KB, 1)) KB" }
    else { return "$Bytes B" }
}

function Get-UserFolderPath {
    param([Parameter(Mandatory=$true)][string]$FolderName)
    switch ($FolderName) {
        'Documents' { return [Environment]::GetFolderPath('MyDocuments') }
        'Pictures'  { return [Environment]::GetFolderPath('MyPictures') }
        'Videos'    { return [Environment]::GetFolderPath('MyVideos') }
        'Music'     { return [Environment]::GetFolderPath('MyMusic') }
        'Desktop'   { return [Environment]::GetFolderPath('Desktop') }
        'Downloads' { return Join-Path $env:USERPROFILE 'Downloads' }
        default     { return $null }
    }
}

function Get-FolderSizes {
    param(
        [Parameter(Mandatory=$true)][string[]]$FolderNames,
        [scriptblock]$OnProgress
    )
    $results = @{}
    foreach ($name in $FolderNames) {
        $path = Get-UserFolderPath -FolderName $name
        $sizeBytes = 0
        $fileCount = 0
        if ($path -and (Test-Path -LiteralPath $path)) {
            try {
                $measure = Get-ChildItem -LiteralPath $path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
                if ($measure.Sum) { $sizeBytes = $measure.Sum }
                $fileCount = $measure.Count
            } catch { }
        }
        $results[$name] = @{ SizeBytes = $sizeBytes; FileCount = $fileCount }
        if ($OnProgress) { & $OnProgress $name $sizeBytes $fileCount }
        Update-GuiStatus
    }
    return $results
}

function Test-BrowserInstalled {
    param([Parameter(Mandatory=$true)][string]$BrowserName)
    switch ($BrowserName) {
        'Chrome' {
            return (Test-Path "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe")
        }
        'Firefox' {
            return ((Test-Path "$env:PROGRAMFILES\Mozilla Firefox\firefox.exe") -or
                    (Test-Path "${env:PROGRAMFILES(x86)}\Mozilla Firefox\firefox.exe"))
        }
        'Edge' {
            return ((Test-Path "${env:PROGRAMFILES(x86)}\Microsoft\Edge\Application\msedge.exe") -or
                    (Test-Path "$env:PROGRAMFILES\Microsoft\Edge\Application\msedge.exe"))
        }
        default { return $false }
    }
}

function Export-BrowserBookmarks {
    param(
        [Parameter(Mandatory=$true)][string]$TargetFolder,
        [Parameter(Mandatory=$true)][string[]]$Browsers
    )
    $exported = [System.Collections.Generic.List[string]]::new()
    $bookmarksDir = Join-Path $TargetFolder "Bookmarks"
    if (-not (Test-Path $bookmarksDir)) { New-Item -Path $bookmarksDir -ItemType Directory -Force | Out-Null }

    $chromiumBrowsers = @{
        'Chrome' = @{ Source = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"; BackupName = "Chrome_Bookmarks.json" }
        'Edge'   = @{ Source = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Bookmarks"; BackupName = "Edge_Bookmarks.json" }
    }

    foreach ($browser in $Browsers) {
        try {
            if ($chromiumBrowsers.ContainsKey($browser)) {
                $info = $chromiumBrowsers[$browser]
                if (Test-Path $info.Source) {
                    $dest = Join-Path $bookmarksDir $info.BackupName
                    Copy-Item -Path $info.Source -Destination $dest -Force
                    $exported.Add($browser)
                    Write-Log "Exported $browser bookmarks"
                }
            }
            elseif ($browser -eq 'Firefox') {
                $profileDir = Get-ChildItem "$env:APPDATA\Mozilla\Firefox\Profiles" -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -like '*.default*' } | Select-Object -First 1
                if ($profileDir) {
                    $ffDir = Join-Path $bookmarksDir "Firefox"
                    if (-not (Test-Path $ffDir)) { New-Item -Path $ffDir -ItemType Directory -Force | Out-Null }
                    $places = Join-Path $profileDir.FullName "places.sqlite"
                    if (Test-Path $places) {
                        Copy-Item -Path $places -Destination $ffDir -Force
                        # Copy WAL and SHM files for SQLite integrity
                        foreach ($ext in @('-wal', '-shm')) {
                            $walFile = "${places}${ext}"
                            if (Test-Path $walFile) { Copy-Item -Path $walFile -Destination $ffDir -Force }
                        }
                    }
                    $backups = Join-Path $profileDir.FullName "bookmarkbackups"
                    if (Test-Path $backups) { Copy-Item -Path $backups -Destination $ffDir -Recurse -Force }
                    $exported.Add("Firefox")
                    Write-Log "Exported Firefox bookmarks"
                }
            }
        } catch {
            Write-Log "Failed to export $browser bookmarks: $($_.Exception.Message)" -Level 'WARN'
        }
    }
    return $exported
}
#endregion Files Migration Helpers

#region Program Inventory
function Get-InstalledPrograms {
    $uninstallRoots = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $rawItems = [System.Collections.Generic.List[object]]::new()
    foreach ($root in $uninstallRoots) {
        if (Test-Path $root) {
            try {
                foreach ($item in (Get-ItemProperty $root -ErrorAction SilentlyContinue)) {
                    $rawItems.Add($item)
                }
            } catch { }
        }
    }

    $out = New-Object System.Collections.Generic.List[object]

    foreach ($it in $rawItems) {
        if (-not $it) { continue }

        $pDisplayName = $it.PSObject.Properties['DisplayName']
        if (-not $pDisplayName) { continue }
        $displayName = [string]$pDisplayName.Value
        if ([string]::IsNullOrWhiteSpace($displayName)) { continue }

        if ($displayName -match '^(Security Update|Update for|Hotfix|Service Pack)\b') { continue }
        if ($displayName -match '\bKB\d{6,7}\b') { continue }
        if (Test-IsNoiseAppName -DisplayName $displayName) { continue }

        $pSystemComponent = $it.PSObject.Properties['SystemComponent']
        if ($pSystemComponent -and ($pSystemComponent.Value -eq 1)) { continue }

        $pParentKeyName = $it.PSObject.Properties['ParentKeyName']
        if ($pParentKeyName -and $pParentKeyName.Value) { continue }

        $displayVersion = $null; $publisher = $null; $installDate = $null

        $p = $it.PSObject.Properties['DisplayVersion']; if ($p) { $displayVersion = [string]$p.Value }
        $p = $it.PSObject.Properties['Publisher']; if ($p) { $publisher = [string]$p.Value }
        $p = $it.PSObject.Properties['InstallDate']; if ($p) { $installDate = [string]$p.Value }

        $psPath = $null
        $p = $it.PSObject.Properties['PSPath']; if ($p) { $psPath = [string]$p.Value }
        $scope = if ($psPath -like '*HKCU:*') { 'CurrentUser' } else { 'LocalMachine' }

        $out.Add([PSCustomObject]@{
            DisplayName    = $displayName.Trim()
            DisplayVersion = $displayVersion
            Publisher      = $publisher
            InstallDate    = $installDate
            Scope          = $scope
        }) | Out-Null
    }

    return $out | Sort-Object DisplayName, DisplayVersion -Unique
}

$script:InstalledProgramsCache = $null

function Reset-InstalledProgramsCache {
    $script:InstalledProgramsCache = $null
}

function Test-ProgramInstalled {
    param([Parameter(Mandatory=$true)][string]$DisplayName)
    if (-not $script:InstalledProgramsCache) {
        $script:InstalledProgramsCache = Get-InstalledPrograms
    }
    $searchName = $DisplayName.ToLowerInvariant()
    foreach ($p in $script:InstalledProgramsCache) {
        $progName = $p.DisplayName.ToLowerInvariant()
        if ($progName -eq $searchName -or $progName.Contains($searchName) -or $searchName.Contains($progName)) {
            return $true
        }
    }
    return $false
}

function Build-ScanBundle {
    param(
        [Parameter(Mandatory=$true)][string]$ClientName,
        [Parameter(Mandatory=$true)][string]$OutputFolder
    )

    $clientFolder = New-ClientFolder -BaseFolder $OutputFolder -ClientName $ClientName
    Initialize-Logging -BaseFolder $clientFolder -Mode "SCAN"

    Write-Log "Starting scan for: $ClientName"
    Update-StatusBar "Scanning..."
    
    $programs = Get-InstalledPrograms
    Write-Log "Found $($programs.Count) programs"

    $meta = [PSCustomObject]@{
        ClientName   = $ClientName
        ScanTime     = (Get-Date).ToString("o")
        ComputerName = $env:COMPUTERNAME
        ToolVersion  = $script:AppVersion
    }

    $bundle = [PSCustomObject]@{ Meta = $meta; Programs = $programs }

    $jsonPath = Join-Path $clientFolder "programs.json"
    $csvPath = Join-Path $clientFolder "programs.csv"
    $htmlPath = Join-Path $clientFolder "report.html"

    $bundle | ConvertTo-Json -Depth 6 | Set-Content -Path $jsonPath -Encoding UTF8
    $programs | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, Scope | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Add-Type -AssemblyName System.Web
    $safeClientName = [System.Web.HttpUtility]::HtmlEncode($ClientName)
    $safeScanTime = [System.Web.HttpUtility]::HtmlEncode($meta.ScanTime)
    $safeComputerName = [System.Web.HttpUtility]::HtmlEncode($meta.ComputerName)
    $htmlBuilder = New-Object System.Text.StringBuilder 4096
    [void]$htmlBuilder.Append(@"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>$safeClientName - Program List</title>
<style>
body { font-family: Segoe UI, Arial; margin: 20px; background: #1e1e1e; color: #eee; }
h2 { color: #0af; }
table { border-collapse: collapse; width: 100%; margin-top: 20px; }
th, td { border: 1px solid #444; padding: 10px; text-align: left; }
th { background: #333; color: #0af; }
tr:nth-child(even) { background: #2a2a2a; }
.meta { color: #888; margin-bottom: 10px; }
</style></head><body>
<h2>LazyTransfer - Program List</h2>
<div class="meta">Client: $safeClientName | Scanned: $safeScanTime | PC: $safeComputerName</div>
<table><tr><th>Name</th><th>Version</th><th>Publisher</th></tr>
"@)
    foreach ($p in $programs) {
        $n = [System.Web.HttpUtility]::HtmlEncode($p.DisplayName)
        $v = [System.Web.HttpUtility]::HtmlEncode($p.DisplayVersion)
        $pub = [System.Web.HttpUtility]::HtmlEncode($p.Publisher)
        [void]$htmlBuilder.AppendLine("<tr><td>$n</td><td>$v</td><td>$pub</td></tr>")
    }
    [void]$htmlBuilder.Append("</table></body></html>")
    $htmlContent = $htmlBuilder.ToString()
    $htmlContent | Set-Content -Path $htmlPath -Encoding UTF8

    Write-Log "Saved: programs.json, programs.csv, report.html"
    Update-StatusBar "Scan complete! Found $($programs.Count) programs."

    try { Stop-Transcript | Out-Null } catch { }
    return $clientFolder
}
#endregion Program Inventory

#region Files Migration Engine
function Start-FilesMigration {
    param(
        [Parameter(Mandatory=$true)][string]$ClientName,
        [Parameter(Mandatory=$true)][string]$OutputFolder,
        [Parameter(Mandatory=$true)][string[]]$SelectedFolders,
        [Parameter(Mandatory=$true)][string]$Mode,
        [string[]]$SelectedBrowsers = @(),
        [hashtable]$PrecomputedSizes = $null,
        [scriptblock]$OnProgress
    )

    $stamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $client = Sanitize-FileName $ClientName
    if ([string]::IsNullOrWhiteSpace($client)) { $client = "Migration" }
    $migrationFolder = Join-Path $OutputFolder "LazyTransfer-Files-${client}-${stamp}"
    New-Item -Path $migrationFolder -ItemType Directory -Force | Out-Null

    Initialize-Logging -BaseFolder $migrationFolder -Mode "FILES"
    Write-Log "Starting files migration for: $ClientName"
    Write-Log "Mode: $Mode | Folders: $($SelectedFolders -join ', ')"

    # Use pre-computed sizes if available, otherwise calculate
    # If user started migration mid-scan, some folders may be missing — measure those live
    $totalBytes = [long]0
    $folderSizes = @{}
    $folderFileCounts = @{}
    foreach ($name in $SelectedFolders) {
        if ($PrecomputedSizes -and $PrecomputedSizes.ContainsKey($name)) {
            $folderSizes[$name] = [long]$PrecomputedSizes[$name]
            $totalBytes += $folderSizes[$name]
        } else {
            $path = Get-UserFolderPath -FolderName $name
            if ($path -and (Test-Path -LiteralPath $path)) {
                try {
                    $measure = Get-ChildItem -LiteralPath $path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
                    $size = if ($measure.Sum) { [long]$measure.Sum } else { 0 }
                    $folderSizes[$name] = $size
                    $folderFileCounts[$name] = $measure.Count
                    $totalBytes += $size
                } catch { $folderSizes[$name] = 0 }
            } else { $folderSizes[$name] = 0 }
            Update-GuiStatus
        }
    }

    # Check available disk space (skip for UNC paths)
    if ($OutputFolder -notlike '\\*') {
        try {
            $destDrive = Split-Path -Qualifier $OutputFolder
            if ($destDrive) {
                $driveInfo = Get-PSDrive -Name ($destDrive.TrimEnd(':')) -ErrorAction SilentlyContinue
                $spaceMultiplier = if ($Mode -eq "zip") { 2.2 } else { 1.1 }
                if ($driveInfo -and $driveInfo.Free -lt ($totalBytes * $spaceMultiplier)) {
                    throw "Insufficient disk space. Need $(Format-FileSize ([long]($totalBytes * $spaceMultiplier))), have $(Format-FileSize $driveInfo.Free)."
                }
            }
        } catch [System.Management.Automation.RuntimeException] { throw }
        catch { }
    }

    $errors = [System.Collections.Generic.List[string]]::new()
    $copiedBytes = [long]0
    $copiedFiles = 0
    $totalFolders = $SelectedFolders.Count + $(if ($SelectedBrowsers.Count -gt 0) { 1 } else { 0 })
    $currentFolder = 0

    # Copy each selected folder
    foreach ($name in $SelectedFolders) {
        if ($script:CancelRequested) {
            Write-Log "File migration cancelled by user" -Level 'WARN'
            $errors.Add("Cancelled by user")
            break
        }
        $currentFolder++
        $srcPath = Get-UserFolderPath -FolderName $name
        if (-not $srcPath -or -not (Test-Path $srcPath)) {
            Write-Log "Skipping $name - folder not found" -Level 'WARN'
            $errors.Add("Folder not found: $name")
            continue
        }

        $destPath = Join-Path $migrationFolder $name
        Write-Log "Copying $name... ($(Format-FileSize $folderSizes[$name]))"
        if ($OnProgress) { & $OnProgress $name $copiedBytes $totalBytes $currentFolder $totalFolders "Copying..." }

        try {
            # Use robocopy for robust file copying
            $roboArgs = @($srcPath, $destPath, '/E', '/COPY:DAT', '/R:1', '/W:1', '/NP', '/MT:4', '/NFL', '/NDL')
            $roboProcess = Start-Process -FilePath "robocopy" -ArgumentList $roboArgs -Wait -PassThru -WindowStyle Hidden
            # Robocopy exit codes 0-7 indicate success (various levels of files copied/skipped)
            if ($roboProcess.ExitCode -le 7) {
                $copiedBytes += $folderSizes[$name]
                # Use pre-computed file count when available, avoids re-scanning
                if ($folderFileCounts.ContainsKey($name)) {
                    $copiedFiles += $folderFileCounts[$name]
                } else {
                    try {
                        $srcMeasure = Get-ChildItem -LiteralPath $srcPath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object
                        $copiedFiles += $srcMeasure.Count
                    } catch { }
                }
                Write-Log "[OK] $name copied successfully" -Level 'SUCCESS'
            } else {
                Write-Log "[X] $name copy had errors (robocopy exit: $($roboProcess.ExitCode))" -Level 'ERROR'
                $errors.Add("Robocopy error for $name (exit code $($roboProcess.ExitCode))")
                # Count only actually copied bytes for accurate progress
                try {
                    $partialMeasure = Get-ChildItem -LiteralPath $destPath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
                    $copiedBytes += if ($partialMeasure.Sum) { [long]$partialMeasure.Sum } else { 0 }
                    $copiedFiles += $partialMeasure.Count
                } catch { }
            }
        } catch {
            Write-Log "[X] Failed to copy $name`: $($_.Exception.Message)" -Level 'ERROR'
            $errors.Add("Failed: $name - $($_.Exception.Message)")
        }

        if ($OnProgress) { & $OnProgress $name $copiedBytes $totalBytes $currentFolder $totalFolders "Done" }
        Update-GuiStatus
    }

    # Export browser bookmarks (skip if cancelled)
    $exportedBrowsers = @()
    if (-not $script:CancelRequested -and $SelectedBrowsers.Count -gt 0) {
        $currentFolder++
        if ($OnProgress) { & $OnProgress "Browser Bookmarks" $copiedBytes $totalBytes $currentFolder $totalFolders "Exporting..." }
        $exportedBrowsers = Export-BrowserBookmarks -TargetFolder $migrationFolder -Browsers $SelectedBrowsers
        if ($OnProgress) { & $OnProgress "Browser Bookmarks" $copiedBytes $totalBytes $currentFolder $totalFolders "Done" }
    }

    # Write migration manifest (always write — records what was actually copied, including cancellation)
    $manifest = [PSCustomObject]@{
        ClientName     = $ClientName
        MigrationDate  = (Get-Date).ToString("o")
        SourceComputer = $env:COMPUTERNAME
        ToolVersion    = $script:AppVersion
        Mode           = $Mode
        Cancelled      = [bool]$script:CancelRequested
        TotalSizeBytes = $totalBytes
        TotalFiles     = $copiedFiles
        Folders        = $SelectedFolders | ForEach-Object {
            [PSCustomObject]@{
                Name      = $_
                SizeBytes = $folderSizes[$_]
            }
        }
        Bookmarks      = $exportedBrowsers
        Errors         = $errors
    }
    $manifest | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $migrationFolder "migration-manifest.json") -Encoding UTF8 -ErrorAction SilentlyContinue
    Write-Log "Migration manifest saved"

    # ZIP mode: compress and remove temp folder (skip if cancelled)
    $finalPath = $migrationFolder
    if (-not $script:CancelRequested -and $Mode -eq "zip") {
        Write-Log "Compressing to ZIP archive..."
        # Stop transcript BEFORE zipping — the log file is inside the migration folder
        # and Compress-Archive can't read a file locked by Start-Transcript
        try { Stop-Transcript | Out-Null } catch { }
        if ($OnProgress) { & $OnProgress "Compressing ZIP" $copiedBytes $totalBytes $totalFolders $totalFolders "Compressing..." }
        Update-GuiStatus

        $zipPath = "${migrationFolder}.zip"
        try {
            Compress-Archive -Path "$migrationFolder\*" -DestinationPath $zipPath -Force
            Remove-Item -Path $migrationFolder -Recurse -Force
            $finalPath = $zipPath
        } catch {
            Write-Log "ZIP compression failed, keeping uncompressed folder: $($_.Exception.Message)" -Level 'WARN'
            $errors.Add("ZIP failed: $($_.Exception.Message)")
            # Clean up partial ZIP to avoid confusion and save disk space
            if (Test-Path $zipPath) { Remove-Item $zipPath -Force -ErrorAction SilentlyContinue }
        }

        if ($OnProgress) { & $OnProgress "Complete" $totalBytes $totalBytes $totalFolders $totalFolders "Done" }
    } else {
        try { Stop-Transcript | Out-Null } catch { }
    }

    return [PSCustomObject]@{
        Success        = ($errors.Count -eq 0)
        Path           = $finalPath
        TotalBytes     = $totalBytes
        TotalFiles     = $copiedFiles
        FoldersCopied  = $SelectedFolders.Count
        BookmarksExported = $exportedBrowsers
        Errors         = $errors
    }
}
#endregion Files Migration Engine

#region Install Engine
$script:AppMap = @(
    @{ Pattern='^Google Chrome$'; WingetId='Google.Chrome'; ChocoId='googlechrome' },
    @{ Pattern='^Mozilla Firefox'; WingetId='Mozilla.Firefox'; ChocoId='firefox' },
    @{ Pattern='^Microsoft Edge$'; WingetId='Microsoft.Edge'; ChocoId=$null },
    @{ Pattern='^7-Zip'; WingetId='7zip.7zip'; ChocoId='7zip' },
    @{ Pattern='^VLC media player'; WingetId='VideoLAN.VLC'; ChocoId='vlc' },
    @{ Pattern='^Notepad\+\+'; WingetId='Notepad++.Notepad++'; ChocoId='notepadplusplus' },
    @{ Pattern='^Zoom'; WingetId='Zoom.Zoom'; ChocoId='zoom' },
    @{ Pattern='^TeamViewer'; WingetId='TeamViewer.TeamViewer'; ChocoId='teamviewer' },
    @{ Pattern='^AnyDesk'; WingetId='AnyDeskSoftwareGmbH.AnyDesk'; ChocoId='anydesk' },
    @{ Pattern='^Adobe Creative Cloud'; WingetId='Adobe.CreativeCloud'; ChocoId=$null },
    @{ Pattern='^Microsoft 365'; WingetId='Microsoft.Office'; ChocoId=$null },
    @{ Pattern='^Microsoft Office'; WingetId='Microsoft.Office'; ChocoId=$null },
    @{ Pattern='^Steam'; WingetId='Valve.Steam'; ChocoId='steam-client' },
    @{ Pattern='^Discord'; WingetId='Discord.Discord'; ChocoId='discord' },
    @{ Pattern='^Git$|^Git\s'; WingetId='Git.Git'; ChocoId='git' },
    @{ Pattern='^Visual Studio Code'; WingetId='Microsoft.VisualStudioCode'; ChocoId='vscode' },
    @{ Pattern='^ShareX'; WingetId='ShareX.ShareX'; ChocoId='sharex' },
    @{ Pattern='^SumatraPDF'; WingetId='SumatraPDF.SumatraPDF'; ChocoId='sumatrapdf' },
    @{ Pattern='^Sumatra PDF'; WingetId='SumatraPDF.SumatraPDF'; ChocoId='sumatrapdf' },
    @{ Pattern='^Everything'; WingetId='voidtools.Everything'; ChocoId='everything' },
    @{ Pattern='^Audacity'; WingetId='Audacity.Audacity'; ChocoId='audacity' },
    @{ Pattern='^Paint\.NET'; WingetId='dotPDN.PaintDotNet'; ChocoId='paint.net' }
)

function Resolve-AppFromStaticMap {
    param([Parameter(Mandatory=$true)][string]$DisplayName)
    foreach ($m in $script:AppMap) {
        if ($DisplayName -match $m.Pattern) {
            return [PSCustomObject]@{
                DisplayName = $DisplayName
                WingetId = $m.WingetId
                ChocoId = $m.ChocoId
                Method = 'StaticMap'
            }
        }
    }
    return $null
}

function Test-WingetPresent {
    try { $null = & winget --version 2>$null; return $true } catch { return $false }
}

function Search-WingetPackage {
    param([Parameter(Mandatory=$true)][string]$DisplayName)
    
    if (-not (Test-WingetPresent)) { return $null }
    
    $searchTerm = $DisplayName -replace '\s*\(.*\)\s*$', ''
    $searchTerm = $searchTerm -replace '\s+', ' '
    $searchTerm = $searchTerm.Trim()
    $words = $searchTerm -split '\s+'
    if ($words.Count -gt 3) { $searchTerm = $words[0..2] -join ' ' }
    
    Write-Log "Searching winget: $searchTerm"
    Update-GuiStatus
    
    try {
        $output = & winget search $searchTerm --source winget --accept-source-agreements 2>&1 | Out-String
        Update-GuiStatus
        
        $lines = $output -split "`n" | Where-Object { $_ -match '\S' }
        $dataStarted = $false
        $results = [System.Collections.Generic.List[object]]::new()

        foreach ($line in $lines) {
            if ($line -match '^-+') { $dataStarted = $true; continue }
            if (-not $dataStarted) { continue }
            if ($line -match '^(.+?)\s{2,}(\S+\.\S+)\s') {
                $results.Add([PSCustomObject]@{ Name = $matches[1].Trim(); Id = $matches[2].Trim() })
            }
        }
        
        if ($results.Count -eq 0) { return $null }
        
        $bestMatch = $null
        $bestScore = 0
        foreach ($r in $results) {
            $score = 0
            $nameLower = $r.Name.ToLowerInvariant()
            $searchLower = $DisplayName.ToLowerInvariant()
            
            if ($nameLower -eq $searchLower) { $score = 100 }
            elseif ($nameLower.StartsWith($searchLower) -or $searchLower.StartsWith($nameLower)) { $score = 80 }
            elseif ($nameLower.Contains($searchLower) -or $searchLower.Contains($nameLower)) { $score = 60 }
            elseif (($nameLower -split '\s+')[0] -eq ($searchLower -split '\s+')[0]) { $score = 50 }
            else { $score = 30 }
            
            if ($score -gt $bestScore) { $bestScore = $score; $bestMatch = $r }
        }
        
        if ($bestMatch -and $bestScore -ge 50) {
            Write-Log "Found: $($bestMatch.Id)"
            return [PSCustomObject]@{
                DisplayName = $DisplayName
                WingetId = $bestMatch.Id
                ChocoId = $null
                Method = 'WingetSearch'
            }
        }
        return $null
    } catch {
        Write-Log "Search error: $($_.Exception.Message)"
        return $null
    }
}

function Resolve-AppInstallTarget {
    param([Parameter(Mandatory=$true)][string]$DisplayName)
    $static = Resolve-AppFromStaticMap -DisplayName $DisplayName
    if ($static) { return $static }
    $dynamic = Search-WingetPackage -DisplayName $DisplayName
    if ($dynamic) { return $dynamic }
    return [PSCustomObject]@{ DisplayName = $DisplayName; WingetId = $null; ChocoId = $null; Method = 'Unmapped' }
}

function Ensure-Winget {
    param([switch]$AutoInstall)
    if (Test-WingetPresent) { return $true }
    if (-not $AutoInstall) { return $false }
    
    Write-Log "Installing winget..."
    try {
        $progressPreference = 'SilentlyContinue'
        $release = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest" -UseBasicParsing
        $msixUrl = $release.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1 -ExpandProperty browser_download_url
        if ($msixUrl) {
            $tempFile = Join-Path $env:TEMP "AppInstaller.msixbundle"
            Invoke-WebRequest -Uri $msixUrl -OutFile $tempFile -UseBasicParsing
            Add-AppxPackage -Path $tempFile -ErrorAction Stop
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
            if (Test-WingetPresent) { Write-Log "Winget installed!"; return $true }
        }
        return $false
    } catch { Write-Log "Winget install failed"; return $false }
}

function Ensure-Chocolatey {
    param([switch]$AutoInstall)
    if (Get-Command choco -ErrorAction SilentlyContinue) { return $true }
    if (-not $AutoInstall) { return $false }
    
    Write-Log "Installing Chocolatey..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))" | Out-Null
        Start-Sleep -Seconds 2
        return [bool](Get-Command choco -ErrorAction SilentlyContinue)
    } catch { Write-Log "Choco install failed"; return $false }
}

function Invoke-WingetInstall {
    param([Parameter(Mandatory=$true)][string]$WingetId)
    $argList = @("install", "--id", $WingetId, "--source", "winget", "--silent", "--accept-source-agreements", "--accept-package-agreements")
    Write-Log "winget install $WingetId"
    $p = Start-Process -FilePath "winget" -ArgumentList $argList -PassThru -WindowStyle Hidden
    $timeout = 300000  # 5 minutes
    if (-not $p.WaitForExit($timeout)) {
        Write-Log "winget install timed out after 5 minutes, killing process" -Level 'WARN'
        try { $p.Kill() } catch { }
        return -1
    }
    return $p.ExitCode
}

function Invoke-ChocoInstall {
    param([Parameter(Mandatory=$true)][string]$ChocoId)
    Write-Log "choco install $ChocoId"
    $p = Start-Process -FilePath "choco" -ArgumentList @("install", $ChocoId, "-y", "--no-progress") -PassThru -WindowStyle Hidden
    $timeout = 300000  # 5 minutes
    if (-not $p.WaitForExit($timeout)) {
        Write-Log "choco install timed out after 5 minutes, killing process" -Level 'WARN'
        try { $p.Kill() } catch { }
        return -1
    }
    return $p.ExitCode
}

function Install-ProgramsFromBundle {
    param(
        [Parameter(Mandatory=$true)][string]$BundleJsonPath,
        [Parameter(Mandatory=$true)][string[]]$SelectedDisplayNames,
        [switch]$AutoInstallChocolatey,
        [switch]$AutoInstallWinget,
        [scriptblock]$OnProgress
    )

    Require-AdminOrThrow
    
    $sessionFolder = Split-Path -Parent $BundleJsonPath
    Initialize-Logging -BaseFolder $sessionFolder -Mode "INSTALL"

    $bundle = Get-Content $BundleJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $bundle.Programs) { throw "Invalid bundle file." }

    # Filter out null/empty display names to prevent parameter binding errors
    $SelectedDisplayNames = @($SelectedDisplayNames | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace($_) })
    if ($SelectedDisplayNames.Count -eq 0) {
        Write-Log "No valid programs selected for installation" -Level 'WARN'
        return @()
    }

    $wingetOk = Ensure-Winget -AutoInstall:$AutoInstallWinget
    Write-Log "Winget: $(if ($wingetOk) { 'Ready' } else { 'Not available' })"
    
    $chocoOk = Ensure-Chocolatey -AutoInstall:$AutoInstallChocolatey
    Write-Log "Chocolatey: $(if ($chocoOk) { 'Ready' } else { 'Not available' })"

    # Reset the installed programs cache so verification reads fresh data
    Reset-InstalledProgramsCache

    $results = [System.Collections.Generic.List[object]]::new()
    $manualQueue = [System.Collections.Generic.List[string]]::new()
    $failedQueue = [System.Collections.Generic.List[string]]::new()
    $stillMissing = [System.Collections.Generic.List[string]]::new()
    
    $total = $SelectedDisplayNames.Count
    $current = 0

    foreach ($name in $SelectedDisplayNames) {
        if ($script:CancelRequested) {
            Write-Log "Installation cancelled by user" -Level 'WARN'
            break
        }
        $current++
        if ($OnProgress) { & $OnProgress $name $current $total "Processing..." }

        if (Test-IsNoiseAppName -DisplayName $name) {
            $results.Add([PSCustomObject]@{ DisplayName=$name; Status="Skipped"; Method="Noise"; Verified=$false })
            continue
        }

        $target = Resolve-AppInstallTarget -DisplayName $name
        Update-GuiStatus

        $status = "Skipped"
        $method = $target.Method
        $verified = $false

        try {
            $installed = $false

            # Try winget first
            if ($target.WingetId -and $wingetOk) {
                if ($OnProgress) { & $OnProgress $name $current $total "Installing (winget)..." }
                $exit = Invoke-WingetInstall -WingetId $target.WingetId
                # 0=success, -1978335189=no applicable update, -1978335188=already installed
                if ($exit -eq 0 -or $exit -eq -1978335189 -or $exit -eq -1978335188) {
                    $status = "Installed"
                    $method = "Winget"
                    $installed = $true
                    Start-Sleep -Milliseconds 500
                    Reset-InstalledProgramsCache
                    $verified = Test-ProgramInstalled -DisplayName $name
                    if (-not $verified) { $stillMissing.Add($name) }
                }
            }

            # Fallback to Chocolatey if winget failed or unavailable
            if (-not $installed -and $target.ChocoId -and $chocoOk) {
                if ($OnProgress) { & $OnProgress $name $current $total "Installing (choco)..." }
                $exit = Invoke-ChocoInstall -ChocoId $target.ChocoId
                # 0=success, 1641=reboot initiated, 3010=reboot required
                if ($exit -eq 0 -or $exit -eq 1641 -or $exit -eq 3010) {
                    $status = "Installed"
                    $method = "Chocolatey"
                    $installed = $true
                    Start-Sleep -Milliseconds 500
                    Reset-InstalledProgramsCache
                    $verified = Test-ProgramInstalled -DisplayName $name
                    if (-not $verified) { $stillMissing.Add($name) }
                }
            }

            if (-not $installed) {
                if ($target.WingetId -or $target.ChocoId) {
                    $status = "Failed"
                    $failedQueue.Add($name)
                } else {
                    $status = "Manual"
                    $method = "Not Found"
                    $manualQueue.Add($name)
                    $stillMissing.Add($name)
                }
            }
        } catch {
            $status = "Failed"
            $failedQueue.Add($name)
            Write-Log "Error: $name - $($_.Exception.Message)"
        }

        $results.Add([PSCustomObject]@{
            DisplayName = $name
            Status = $status
            Method = $method
            Verified = $verified
        })

        if ($OnProgress) { & $OnProgress $name $current $total $status }
        Update-GuiStatus
    }

    $stamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $results | Export-Csv -Path (Join-Path $sessionFolder "install_results_$stamp.csv") -NoTypeInformation -Encoding UTF8
    $results | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $sessionFolder "install_results_$stamp.json") -Encoding UTF8

    try { Stop-Transcript | Out-Null } catch { }

    return [PSCustomObject]@{
        Results = $results
        ManualQueue = $manualQueue
        FailedQueue = $failedQueue
        StillMissing = $stillMissing
    }
}
#endregion Install Engine

#region Restore Engine

function Import-BrowserBookmarks {
    param(
        [Parameter(Mandatory=$true)][string]$SourceFolder,
        [string[]]$Browsers = @('Chrome', 'Firefox', 'Edge')
    )
    $results = [System.Collections.Generic.List[object]]::new()
    $bookmarksDir = Join-Path $SourceFolder "Bookmarks"
    if (-not (Test-Path $bookmarksDir)) {
        Write-Log "No Bookmarks subfolder found in source" -Level 'WARN'
        return $results
    }

    foreach ($browser in $Browsers) {
        try {
            if (-not (Test-BrowserInstalled -BrowserName $browser)) {
                Write-Log "$browser not installed, skipping bookmark restore" -Level 'WARN'
                $results.Add([PSCustomObject]@{ Browser = $browser; Status = "Skipped - not installed" })
                continue
            }

            $chromiumBrowsers = @{
                'Chrome' = @{ BackupName = "Chrome_Bookmarks.json"; Dest = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks" }
                'Edge'   = @{ BackupName = "Edge_Bookmarks.json"; Dest = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Bookmarks" }
            }

            if ($chromiumBrowsers.ContainsKey($browser)) {
                $info = $chromiumBrowsers[$browser]
                $src = Join-Path $bookmarksDir $info.BackupName
                $dest = $info.Dest
                if (Test-Path $src) {
                    $destDir = Split-Path -Parent $dest
                    if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force | Out-Null }
                    if (Test-Path $dest) {
                        $bakName = "${dest}.bak.$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                        if (-not (Test-Path "${dest}.bak")) { $bakName = "${dest}.bak" }
                        Copy-Item -Path $dest -Destination $bakName -Force
                    }
                    Copy-Item -Path $src -Destination $dest -Force
                    $results.Add([PSCustomObject]@{ Browser = $browser; Status = "Restored" })
                    Write-Log "$browser bookmarks restored" -Level 'SUCCESS'
                } else {
                    $results.Add([PSCustomObject]@{ Browser = $browser; Status = "No backup found" })
                }
            }
            elseif ($browser -eq 'Firefox') {
                    $ffSrc = Join-Path $bookmarksDir "Firefox"
                    if (Test-Path $ffSrc) {
                        $profileDir = Get-ChildItem "$env:APPDATA\Mozilla\Firefox\Profiles" -Directory -ErrorAction SilentlyContinue |
                            Where-Object { $_.Name -like '*.default*' } | Select-Object -First 1
                        if ($profileDir) {
                            $places = Join-Path $ffSrc "places.sqlite"
                            if (Test-Path $places) {
                                $destPlaces = Join-Path $profileDir.FullName "places.sqlite"
                                if (Test-Path $destPlaces) {
                                    $bakName = "${destPlaces}.bak.$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                                    if (-not (Test-Path "${destPlaces}.bak")) { $bakName = "${destPlaces}.bak" }
                                    Copy-Item -Path $destPlaces -Destination $bakName -Force
                                }
                                Copy-Item -Path $places -Destination $profileDir.FullName -Force
                                foreach ($ext in @('-wal', '-shm')) {
                                    $walFile = "${places}${ext}"
                                    if (Test-Path $walFile) { Copy-Item -Path $walFile -Destination $profileDir.FullName -Force }
                                }
                            }
                            $backups = Join-Path $ffSrc "bookmarkbackups"
                            if (Test-Path $backups) { Copy-Item -Path $backups -Destination $profileDir.FullName -Recurse -Force }
                            $results.Add([PSCustomObject]@{ Browser = "Firefox"; Status = "Restored" })
                            Write-Log "Firefox bookmarks restored" -Level 'SUCCESS'
                        } else {
                            $results.Add([PSCustomObject]@{ Browser = "Firefox"; Status = "No profile found" })
                        }
                    } else {
                        $results.Add([PSCustomObject]@{ Browser = "Firefox"; Status = "No backup found" })
                    }
            }
        } catch {
            Write-Log "Failed to restore $browser bookmarks: $($_.Exception.Message)" -Level 'ERROR'
            $results.Add([PSCustomObject]@{ Browser = $browser; Status = "Error: $($_.Exception.Message)" })
        }
    }
    return $results
}

function Restore-FilesMigration {
    param(
        [Parameter(Mandatory=$true)][string]$BundlePath,
        [string[]]$SelectedFolders = @(),
        [string[]]$SelectedBrowsers = @(),
        [scriptblock]$OnProgress
    )

    $tempExtract = $null
    $sourcePath = $BundlePath

    # If ZIP, extract to temp first
    if ($BundlePath -like '*.zip' -and (Test-Path $BundlePath)) {
        $tempExtract = Join-Path $env:TEMP "LazyTransfer-Restore-$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        Write-Log "Extracting ZIP bundle..."
        if ($OnProgress) { & $OnProgress "Extracting" 0 0 0 0 "Extracting ZIP..." }
        Expand-ZipSafe -ZipPath $BundlePath -DestinationFolder $tempExtract
        $sourcePath = $tempExtract
        Write-Log "Extracted to temp folder" -Level 'SUCCESS'
    }

    $errors = [System.Collections.Generic.List[string]]::new()
    $totalBytesRestored = [long]0
    $foldersRestored = 0
    $bookmarksRestored = @()

    try {
        # Read manifest
        $manifest = $null
        $manifestPath = Join-Path $sourcePath "migration-manifest.json"
        if (Test-Path $manifestPath) {
            $manifest = Get-Content $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        }

        $totalSteps = $SelectedFolders.Count + $(if ($SelectedBrowsers.Count -gt 0) { 1 } else { 0 })
        $currentStep = 0

        # Restore each folder
        foreach ($name in $SelectedFolders) {
            if ($script:CancelRequested) { break }
            $currentStep++
            $srcPath = Join-Path $sourcePath $name
            if (-not (Test-Path $srcPath)) {
                Write-Log "Folder not found in bundle: $name" -Level 'WARN'
                $errors.Add("Folder not in bundle: $name")
                continue
            }

            $destPath = Get-UserFolderPath -FolderName $name
            if (-not $destPath) {
                $errors.Add("Unknown folder: $name")
                continue
            }

            Write-Log "Restoring $name to $destPath..."
            if ($OnProgress) { & $OnProgress $name 0 0 $currentStep $totalSteps "Restoring..." }

            try {
                # robocopy merge mode (no /PURGE)
                $roboArgs = @($srcPath, $destPath, '/E', '/COPY:DAT', '/R:1', '/W:1', '/NP', '/MT:4', '/NFL', '/NDL')
                $roboProcess = Start-Process -FilePath "robocopy" -ArgumentList $roboArgs -Wait -PassThru -WindowStyle Hidden
                if ($roboProcess.ExitCode -le 7) {
                    $foldersRestored++
                    try {
                        $measure = Get-ChildItem -LiteralPath $srcPath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
                        if ($measure.Sum) { $totalBytesRestored += [long]$measure.Sum }
                    } catch { }
                    Write-Log "[OK] $name restored" -Level 'SUCCESS'
                } else {
                    $errors.Add("Robocopy error for $name (exit $($roboProcess.ExitCode))")
                    Write-Log "[X] $name restore had errors" -Level 'ERROR'
                }
            } catch {
                $errors.Add("Failed: $name - $($_.Exception.Message)")
                Write-Log "[X] Failed to restore $name" -Level 'ERROR'
            }
            if ($OnProgress) { & $OnProgress $name $totalBytesRestored 0 $currentStep $totalSteps "Done" }
            Update-GuiStatus
        }

        # Restore bookmarks
        if ($SelectedBrowsers.Count -gt 0) {
            $currentStep++
            if ($OnProgress) { & $OnProgress "Bookmarks" 0 0 $currentStep $totalSteps "Restoring bookmarks..." }
            $bookmarksRestored = Import-BrowserBookmarks -SourceFolder $sourcePath -Browsers $SelectedBrowsers
            if ($OnProgress) { & $OnProgress "Bookmarks" 0 0 $currentStep $totalSteps "Done" }
        }
    } finally {
        # Always cleanup temp extraction, even on error
        if ($tempExtract -and (Test-Path $tempExtract)) {
            try { Remove-Item $tempExtract -Recurse -Force } catch { }
        }
    }

    return [PSCustomObject]@{
        Success          = ($errors.Count -eq 0)
        TotalBytes       = $totalBytesRestored
        FoldersRestored  = $foldersRestored
        BookmarksRestored = $bookmarksRestored
        Errors           = $errors
    }
}

function Find-BundleContents {
    param([Parameter(Mandatory=$true)][string]$Path)

    $result = [PSCustomObject]@{
        Programs = [PSCustomObject]@{ Found = $false; Path = $null; Count = 0 }
        Files    = [PSCustomObject]@{ Found = $false; Path = $null; Folders = @(); TotalSize = [long]0 }
        System   = [PSCustomObject]@{ Found = $false; Path = $null; Items = @() }
        Bookmarks = [PSCustomObject]@{ Found = $false; Path = $null; Browsers = @() }
        Meta     = [PSCustomObject]@{ SourcePC = ""; Date = ""; ToolVersion = "" }
    }

    # Look for programs.json
    $programsJson = Join-Path $Path "programs.json"
    if (Test-Path $programsJson) {
        try {
            $bundle = Get-Content $programsJson -Raw -Encoding UTF8 | ConvertFrom-Json
            $result.Programs = [PSCustomObject]@{
                Found = $true
                Path  = $programsJson
                Count = @($bundle.Programs).Count
            }
            if ($bundle.Meta) {
                $result.Meta = [PSCustomObject]@{
                    SourcePC    = if ($bundle.Meta.ComputerName) { $bundle.Meta.ComputerName } else { "" }
                    Date        = if ($bundle.Meta.ScanTime) { $bundle.Meta.ScanTime } else { "" }
                    ToolVersion = if ($bundle.Meta.ToolVersion) { $bundle.Meta.ToolVersion } else { "" }
                }
            }
        } catch { }
    }

    # Look for SystemMigration/ folder
    $systemDir = Join-Path $Path "SystemMigration"
    if (Test-Path $systemDir) {
        $items = @()
        if (Test-Path (Join-Path $systemDir "wifi-profiles")) { $items += "WiFi" }
        if (Test-Path (Join-Path $systemDir "ssh")) { $items += "SSH" }
        if (Test-Path (Join-Path $systemDir "git-config")) { $items += "Git" }
        if (Test-Path (Join-Path $systemDir "env-vars.json")) { $items += "Environment Variables" }
        $result.System = [PSCustomObject]@{
            Found = $true
            Path  = $systemDir
            Items = $items
        }
    }

    # Look for LazyTransfer-Files-* folder or ZIP
    $filesFolders = Get-ChildItem -Path $Path -Directory -Filter "LazyTransfer-Files-*" -ErrorAction SilentlyContinue
    $filesZips = Get-ChildItem -Path $Path -File -Filter "LazyTransfer-Files-*.zip" -ErrorAction SilentlyContinue

    $filesSource = $null
    if ($filesFolders) {
        $filesSource = $filesFolders | Select-Object -First 1
    }
    # Also check if the path itself IS a files migration folder (contains migration-manifest.json)
    $manifestPath = Join-Path $Path "migration-manifest.json"
    if (Test-Path $manifestPath) {
        $filesSource = Get-Item $Path
    }

    if ($filesSource -or $filesZips) {
        $scanPath = if ($filesSource) { $filesSource.FullName } else { $null }
        $folders = [System.Collections.Generic.List[object]]::new()
        $totalSize = [long]0
        $manifest = $null

        # Read manifest first — use stored sizes when available
        if ($scanPath) {
            $mf = Join-Path $scanPath "migration-manifest.json"
            if (Test-Path $mf) {
                try {
                    $manifest = Get-Content $mf -Raw -Encoding UTF8 | ConvertFrom-Json
                    if ($manifest.SourceComputer -and -not $result.Meta.SourcePC) {
                        $result.Meta = [PSCustomObject]@{
                            SourcePC    = $manifest.SourceComputer
                            Date        = if ($manifest.MigrationDate) { $manifest.MigrationDate } else { "" }
                            ToolVersion = if ($manifest.ToolVersion) { $manifest.ToolVersion } else { "" }
                        }
                    }
                } catch { $manifest = $null }
            }
        }

        # Build folder list: prefer manifest sizes, fall back to recursive scan
        if ($scanPath) {
            $manifestSizes = @{}
            if ($manifest -and $manifest.Folders) {
                foreach ($mfFolder in $manifest.Folders) {
                    if ($mfFolder.Name -and $mfFolder.SizeBytes) { $manifestSizes[$mfFolder.Name] = [long]$mfFolder.SizeBytes }
                }
            }

            $folderNames = @('Documents', 'Downloads', 'Pictures', 'Videos', 'Music', 'Desktop')
            foreach ($fn in $folderNames) {
                $fp = Join-Path $scanPath $fn
                if (Test-Path $fp) {
                    if ($manifestSizes.ContainsKey($fn)) {
                        $size = $manifestSizes[$fn]
                    } else {
                        $size = 0
                        try {
                            $measure = Get-ChildItem -LiteralPath $fp -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
                            if ($measure.Sum) { $size = [long]$measure.Sum }
                        } catch { }
                    }
                    $folders.Add([PSCustomObject]@{ Name = $fn; SizeBytes = $size })
                    $totalSize += $size
                }
            }
        }

        $result.Files = [PSCustomObject]@{
            Found     = $true
            Path      = if ($filesSource) { $filesSource.FullName } elseif ($filesZips) { $filesZips[0].FullName } else { $null }
            Folders   = $folders
            TotalSize = $totalSize
        }
    }

    # Look for Bookmarks/ subfolder
    # Check both in root and in any files migration subfolder
    $bookmarksDir = Join-Path $Path "Bookmarks"
    if (-not (Test-Path $bookmarksDir) -and $filesSource) {
        $bookmarksDir = Join-Path $filesSource.FullName "Bookmarks"
    }
    if (Test-Path $bookmarksDir) {
        $browsers = [System.Collections.Generic.List[string]]::new()
        if (Test-Path (Join-Path $bookmarksDir "Chrome_Bookmarks.json")) { $browsers.Add("Chrome") }
        if (Test-Path (Join-Path $bookmarksDir "Edge_Bookmarks.json")) { $browsers.Add("Edge") }
        if (Test-Path (Join-Path $bookmarksDir "Firefox")) { $browsers.Add("Firefox") }
        $result.Bookmarks = [PSCustomObject]@{
            Found    = $true
            Path     = $bookmarksDir
            Browsers = $browsers
        }
    }

    return $result
}

function Start-FullRestore {
    param(
        [Parameter(Mandatory=$true)][string]$BundlePath,
        [PSCustomObject]$BundleContents = $null,
        [hashtable]$Selections = @{},
        [scriptblock]$OnPhaseChange,
        [scriptblock]$OnProgress
    )

    if (-not $BundleContents) {
        $BundleContents = Find-BundleContents -Path $BundlePath
    }

    $results = [PSCustomObject]@{
        System   = $null
        Programs = $null
        Files    = $null
        Errors   = @()
    }

    # Phase 1: System settings
    if ($Selections.ContainsKey('System') -and $Selections.System -and $BundleContents.System.Found) {
        if ($OnPhaseChange) { & $OnPhaseChange "System" "Restoring system settings..." }
        try {
            if (Get-Command Start-SystemImport -ErrorAction SilentlyContinue) {
                $results.System = Start-SystemImport -BundlePath $BundlePath -SelectedItems $Selections.System -OnProgress $OnProgress
            } else {
                Write-Log "SystemEngine not loaded, skipping system restore" -Level 'WARN'
                $results.Errors += "SystemEngine not available"
            }
        } catch {
            Write-Log "System restore failed: $($_.Exception.Message)" -Level 'ERROR'
            $results.Errors += "System: $($_.Exception.Message)"
        }
    }

    # Phase 2: Programs (slowest, needs network)
    if ($Selections.ContainsKey('Programs') -and $Selections.Programs -eq $true -and $BundleContents.Programs.Found) {
        if ($OnPhaseChange) { & $OnPhaseChange "Programs" "Installing programs..." }
        try {
            $programsJson = $BundleContents.Programs.Path
            $bundle = Get-Content $programsJson -Raw -Encoding UTF8 | ConvertFrom-Json
            $selectedNames = @($bundle.Programs | ForEach-Object { $_.DisplayName })
            $results.Programs = Install-ProgramsFromBundle -BundleJsonPath $programsJson -SelectedDisplayNames $selectedNames -AutoInstallWinget -AutoInstallChocolatey -OnProgress $OnProgress
        } catch {
            Write-Log "Program install failed: $($_.Exception.Message)" -Level 'ERROR'
            $results.Errors += "Programs: $($_.Exception.Message)"
        }
    }

    # Phase 3: Files + Bookmarks (needs browsers installed)
    $hasFiles = $Selections.ContainsKey('Files') -and $Selections.Files -and $BundleContents.Files.Found
    $hasBookmarks = $Selections.ContainsKey('Bookmarks') -and $Selections.Bookmarks -and $BundleContents.Bookmarks.Found
    if ($hasFiles -or $hasBookmarks) {
        if ($OnPhaseChange) { & $OnPhaseChange "Files" "Restoring files and bookmarks..." }
        try {
            $fileFolders = if ($Selections.Files -is [array]) { $Selections.Files } else { @() }
            $fileBrowsers = if ($Selections.Bookmarks -is [array]) { $Selections.Bookmarks } else { @() }
            $filesPath = if ($BundleContents.Files.Path) { $BundleContents.Files.Path } else { $BundlePath }
            $results.Files = Restore-FilesMigration -BundlePath $filesPath -SelectedFolders $fileFolders -SelectedBrowsers $fileBrowsers -OnProgress $OnProgress
        } catch {
            Write-Log "Files restore failed: $($_.Exception.Message)" -Level 'ERROR'
            $results.Errors += "Files: $($_.Exception.Message)"
        }
    }

    return $results
}
#endregion Restore Engine

#region GUI

$script:LogTextBox = $null
$script:StatusLabel = $null
$script:ContentPanel = $null
$script:SidebarButtons = @{}
$script:OperationInProgress = $false
$script:CancelRequested = $false
$script:LastMigrationPath = $null

function New-StyledButton {
    param(
        [string]$Text,
        [int]$Width = 150,
        [int]$Height = 35,
        [System.Drawing.Color]$BackColor,
        [System.Drawing.Color]$ForeColor
    )
    $btn = New-Object ModernButton
    $btn.Text = $Text
    $btn.Size = New-Object System.Drawing.Size($Width, $Height)
    $btn.NormalColor = $BackColor
    # Compute lighter hover color
    $hR = [Math]::Min(255, $BackColor.R + 15)
    $hG = [Math]::Min(255, $BackColor.G + 15)
    $hB = [Math]::Min(255, $BackColor.B + 15)
    $btn.HoverColor = [System.Drawing.Color]::FromArgb($hR, $hG, $hB)
    $pR = [Math]::Max(0, $BackColor.R - 10)
    $pG = [Math]::Max(0, $BackColor.G - 10)
    $pB = [Math]::Max(0, $BackColor.B - 10)
    $btn.PressColor = [System.Drawing.Color]::FromArgb($pR, $pG, $pB)
    $btn.ForeColor = $ForeColor
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    return $btn
}

function New-SidebarButton {
    param(
        [string]$Text,
        [string]$Page,
        [System.Drawing.Color]$AccentColor
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "      $Text"
    $btn.Size = New-Object System.Drawing.Size(200, 44)
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.FlatAppearance.BorderSize = 0
    $btn.BackColor = [System.Drawing.Color]::Transparent
    $btn.ForeColor = $script:Colors.TextSecondary
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $btn.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.Tag = @{ Page = $Page; Accent = $AccentColor; IsActive = $false }

    # Paint accent bar on left edge when active
    $btn.Add_Paint({
        param($s, $e)
        if ($s.Tag.IsActive) {
            $g = $e.Graphics
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            try {
                $accentBrush = New-Object System.Drawing.SolidBrush($s.Tag.Accent)
                # 3px rounded accent bar, vertically centered
                $barHeight = [int]($s.Height * 0.6)
                $barY = ($s.Height - $barHeight) / 2
                $path = New-Object System.Drawing.Drawing2D.GraphicsPath
                $path.AddArc(0, $barY, 3, 3, 180, 90)
                $path.AddArc(0, $barY + $barHeight - 3, 3, 3, 90, 90)
                $path.CloseFigure()
                $g.FillRectangle($accentBrush, 0, $barY, 3, $barHeight)
                $accentBrush.Dispose()
                $path.Dispose()
            } catch { }
        }
    })

    $btn.Add_MouseEnter({
        if (-not $this.Tag.IsActive) {
            $this.BackColor = $script:Colors.SidebarHover
        }
    })
    $btn.Add_MouseLeave({
        if ($this.Tag.IsActive) {
            $this.BackColor = $script:Colors.SidebarActive
        } else {
            $this.BackColor = [System.Drawing.Color]::Transparent
        }
    })

    return $btn
}

function Set-ActiveSidebarButton {
    param([string]$Page)
    $script:CurrentPage = $Page
    foreach ($key in $script:SidebarButtons.Keys) {
        $btn = $script:SidebarButtons[$key]
        $isActive = ($key -eq $Page)
        $btn.Tag.IsActive = $isActive
        if ($isActive) {
            $btn.BackColor = $script:Colors.SidebarActive
            $btn.ForeColor = $btn.Tag.Accent
            $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        } else {
            $btn.BackColor = [System.Drawing.Color]::Transparent
            $btn.ForeColor = $script:Colors.TextSecondary
            $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
        }
        $btn.Invalidate()
    }
}

function Clear-ContentPanel {
    if ($script:ContentPanel) {
        foreach ($ctrl in @($script:ContentPanel.Controls)) {
            try { $ctrl.Dispose() } catch { }
        }
        $script:ContentPanel.Controls.Clear()
    }
}

function Show-ScanPage {
    if ($script:OperationInProgress) { return }
    Clear-ContentPanel
    Set-ActiveSidebarButton -Page "Scan"
    Update-StatusBar "Ready to scan"
    
    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Scan Programs"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $script:Colors.AccentBlue
    $title.Location = New-Object System.Drawing.Point(20, 15)
    $title.AutoSize = $true
    
    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Scan installed programs on this PC and save to a bundle file."
    $subtitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitle.ForeColor = $script:Colors.TextSecondary
    $subtitle.Location = New-Object System.Drawing.Point(20, 50)
    $subtitle.AutoSize = $true
    
    $lblClient = New-Object System.Windows.Forms.Label
    $lblClient.Text = "Client / PC Name:"
    $lblClient.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblClient.ForeColor = $script:Colors.TextPrimary
    $lblClient.Location = New-Object System.Drawing.Point(20, 100)
    $lblClient.AutoSize = $true
    
    $txtClient = New-Object System.Windows.Forms.TextBox
    $txtClient.Location = New-Object System.Drawing.Point(20, 125)
    $txtClient.Size = New-Object System.Drawing.Size(350, 28)
    $txtClient.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $txtClient.BackColor = $script:Colors.InputBg
    $txtClient.ForeColor = $script:Colors.TextPrimary
    $txtClient.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    # Pre-fill from saved settings
    if ($script:UserSettings.LastClientName) { $txtClient.Text = $script:UserSettings.LastClientName }

    $lblOut = New-Object System.Windows.Forms.Label
    $lblOut.Text = "Save Location:"
    $lblOut.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblOut.ForeColor = $script:Colors.TextPrimary
    $lblOut.Location = New-Object System.Drawing.Point(20, 170)
    $lblOut.AutoSize = $true
    
    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(20, 195)
    $txtOut.Size = New-Object System.Drawing.Size(580, 28)
    $txtOut.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $txtOut.BackColor = $script:Colors.InputBg
    $txtOut.ForeColor = $script:Colors.TextPrimary
    $txtOut.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    # Pre-fill from saved settings, or default to script directory
    if ($script:UserSettings.LastOutputFolder) {
        $txtOut.Text = $script:UserSettings.LastOutputFolder
    } else {
        $txtOut.Text = $script:ScriptDir
    }

    $btnBrowse = New-StyledButton -Text "Browse..." -Width 100 -Height 28 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnBrowse.Location = New-Object System.Drawing.Point(610, 195)
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Choose where to save the scan bundle"
        $dlg.ShowNewFolderButton = $true
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtOut.Text = $dlg.SelectedPath
        }
        $dlg.Dispose()
    }.GetNewClosure())
    
    $btnScan = New-StyledButton -Text "Scan + Save" -Width 200 -Height 45 -BackColor $script:Colors.AccentBlue -ForeColor $script:Colors.TextPrimary
    $btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $btnScan.Location = New-Object System.Drawing.Point(20, 250)
    $btnScan.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($txtClient.Text)) { throw "Enter a client/PC name." }
            if ([string]::IsNullOrWhiteSpace($txtOut.Text)) { throw "Choose a save location." }
            
            $btnScan.Enabled = $false
            $btnScan.Text = "Scanning..."
            $script:OperationInProgress = $true
            Update-GuiStatus

            $folder = Build-ScanBundle -ClientName $txtClient.Text -OutputFolder $txtOut.Text

            # Save settings for next time
            Update-Setting -Key 'LastClientName' -Value $txtClient.Text
            Update-Setting -Key 'LastOutputFolder' -Value $txtOut.Text

            [System.Windows.Forms.MessageBox]::Show("Scan complete! Saved to: $folder", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            Write-Log "Error: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $script:OperationInProgress = $false
            $btnScan.Enabled = $true
            $btnScan.Text = "Scan + Save"
        }
    }.GetNewClosure())
    
    $script:ContentPanel.Controls.AddRange(@($title, $subtitle, $lblClient, $txtClient, $lblOut, $txtOut, $btnBrowse, $btnScan))
}

function Show-InstallPage {
    if ($script:OperationInProgress) { return }
    Clear-ContentPanel
    Set-ActiveSidebarButton -Page "Install"
    Update-StatusBar "Ready to install"
    
    $script:InstallAllPrograms = @()
    $script:InstallManualQueue = @()
    $script:InstallFailedQueue = @()
    $script:InstallCheckStates = @{}
    
    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Install Programs"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $script:Colors.AccentGreen
    $title.Location = New-Object System.Drawing.Point(20, 15)
    $title.AutoSize = $true
    
    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Load a scan bundle and install programs on this PC."
    $subtitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitle.ForeColor = $script:Colors.TextSecondary
    $subtitle.Location = New-Object System.Drawing.Point(20, 50)
    $subtitle.AutoSize = $true
    
    $lblBundle = New-Object System.Windows.Forms.Label
    $lblBundle.Text = "Bundle File (programs.json):"
    $lblBundle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblBundle.ForeColor = $script:Colors.TextPrimary
    $lblBundle.Location = New-Object System.Drawing.Point(20, 90)
    $lblBundle.AutoSize = $true
    
    $txtBundle = New-Object System.Windows.Forms.TextBox
    $txtBundle.Location = New-Object System.Drawing.Point(20, 115)
    $txtBundle.Size = New-Object System.Drawing.Size(540, 28)
    $txtBundle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtBundle.BackColor = $script:Colors.InputBg
    $txtBundle.ForeColor = $script:Colors.TextPrimary
    $txtBundle.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    # Pre-fill from saved settings
    if ($script:UserSettings.LastBundlePath -and (Test-Path $script:UserSettings.LastBundlePath)) {
        $txtBundle.Text = $script:UserSettings.LastBundlePath
    }

    $btnBrowseBundle = New-StyledButton -Text "Browse" -Width 80 -Height 28 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnBrowseBundle.Location = New-Object System.Drawing.Point(565, 115)
    
    $btnLoad = New-StyledButton -Text "Load" -Width 70 -Height 28 -BackColor $script:Colors.AccentGreen -ForeColor $script:Colors.TextPrimary
    $btnLoad.Location = New-Object System.Drawing.Point(650, 115)
    
    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "Filter:"
    $lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblFilter.ForeColor = $script:Colors.TextSecondary
    $lblFilter.Location = New-Object System.Drawing.Point(20, 152)
    $lblFilter.AutoSize = $true
    
    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = New-Object System.Drawing.Point(60, 150)
    $txtFilter.Size = New-Object System.Drawing.Size(140, 24)
    $txtFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtFilter.BackColor = $script:Colors.InputBg
    $txtFilter.ForeColor = $script:Colors.TextPrimary
    $txtFilter.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Category filter panel
    $categoryPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $categoryPanel.Location = New-Object System.Drawing.Point(210, 148)
    $categoryPanel.Size = New-Object System.Drawing.Size(550, 28)
    $categoryPanel.BackColor = $script:Colors.ContentBg
    $categoryPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $categoryPanel.WrapContents = $false
    $categoryPanel.AutoScroll = $true

    # Store category buttons for later updates
    $script:CategoryButtons = @{}
    $script:CategoryCounts = @{}

    # Create category buttons
    $categoryList = @('Browsers', 'DevTools', 'Media', 'Graphics', 'Office', 'Communication', 'Utilities', 'Gaming', 'Security', 'Other')
    foreach ($cat in $categoryList) {
        $catInfo = Get-CategoryInfo -Category $cat
        $catBtn = New-Object ModernButton
        $catBtn.Text = "$cat (0)"
        $catBtn.Font = New-Object System.Drawing.Font("Segoe UI", 7)
        $catBtn.Size = New-Object System.Drawing.Size(75, 24)
        $catBtn.NormalColor = $script:Colors.ButtonBg
        $catBtn.HoverColor = $script:Colors.ButtonHover
        $catBtn.ForeColor = $script:Colors.TextSecondary
        $catBtn.Tag = $cat
        $catBtn.Radius = 4
        $catBtn.Margin = New-Object System.Windows.Forms.Padding(2, 0, 2, 0)
        $script:CategoryButtons[$cat] = $catBtn
        $script:CategoryCounts[$cat] = 0
        $categoryPanel.Controls.Add($catBtn)
    }

    $listPrograms = New-Object System.Windows.Forms.CheckedListBox
    $listPrograms.Location = New-Object System.Drawing.Point(20, 180)
    $listPrograms.Size = New-Object System.Drawing.Size(400, 280)
    $listPrograms.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $listPrograms.BackColor = $script:Colors.InputBg
    $listPrograms.ForeColor = $script:Colors.TextPrimary
    $listPrograms.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $listPrograms.CheckOnClick = $true

    # Track user check state changes
    $listPrograms.Add_ItemCheck({
        param($sender, $e)
        $itemName = [string]$sender.Items[$e.Index]
        $script:InstallCheckStates[$itemName] = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
    }.GetNewClosure())

    # Wire up category button click handlers
    foreach ($cat in $categoryList) {
        $btn = $script:CategoryButtons[$cat]
        $btn.Add_Click({
            param($sender, $e)
            $clickedCat = $sender.Tag
            # Toggle all apps in this category
            for ($i = 0; $i -lt $listPrograms.Items.Count; $i++) {
                $appName = [string]$listPrograms.Items[$i]
                $app = $script:InstallAllPrograms | Where-Object { $_.Name -eq $appName } | Select-Object -First 1
                if ($app -and $app.Category -eq $clickedCat) {
                    $currentState = $listPrograms.GetItemChecked($i)
                    $listPrograms.SetItemChecked($i, -not $currentState)
                }
            }
            # Visual feedback - briefly highlight the button
            $sender.NormalColor = $script:Colors.AccentGreen
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 200
            $timer.Add_Tick({
                try { if (-not $sender.IsDisposed) { $sender.NormalColor = $script:Colors.ButtonBg; $sender.Invalidate() } } catch { }
                $timer.Stop()
                $timer.Dispose()
            }.GetNewClosure())
            $timer.Start()
        }.GetNewClosure())
    }

    $lblResults = New-Object System.Windows.Forms.Label
    $lblResults.Text = "Results:"
    $lblResults.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblResults.ForeColor = $script:Colors.TextPrimary
    $lblResults.Location = New-Object System.Drawing.Point(440, 152)
    $lblResults.AutoSize = $true
    
    $txtResults = New-Object System.Windows.Forms.TextBox
    $txtResults.Location = New-Object System.Drawing.Point(440, 180)
    $txtResults.Size = New-Object System.Drawing.Size(320, 220)
    $txtResults.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtResults.BackColor = $script:Colors.InputBg
    $txtResults.ForeColor = $script:Colors.TextPrimary
    $txtResults.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtResults.Multiline = $true
    $txtResults.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtResults.ReadOnly = $true
    
    # Enhanced Progress Panel
    $progressPanel = New-Object System.Windows.Forms.Panel
    $progressPanel.Location = New-Object System.Drawing.Point(440, 405)
    $progressPanel.Size = New-Object System.Drawing.Size(320, 55)
    $progressPanel.BackColor = $script:Colors.ContentBg

    $lblPercent = New-Object System.Windows.Forms.Label
    $lblPercent.Text = "0%"
    $lblPercent.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblPercent.ForeColor = $script:Colors.AccentGreen
    $lblPercent.Location = New-Object System.Drawing.Point(0, 0)
    $lblPercent.Size = New-Object System.Drawing.Size(50, 25)

    $progress = New-Object ModernProgressBar
    $progress.Location = New-Object System.Drawing.Point(55, 5)
    $progress.Size = New-Object System.Drawing.Size(265, 14)
    $progress.BarColor = $script:Colors.AccentGreen
    $progress.TrackColor = $script:Colors.CardBg

    $lblCurrentApp = New-Object System.Windows.Forms.Label
    $lblCurrentApp.Text = ""
    $lblCurrentApp.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblCurrentApp.ForeColor = $script:Colors.TextPrimary
    $lblCurrentApp.Location = New-Object System.Drawing.Point(0, 28)
    $lblCurrentApp.Size = New-Object System.Drawing.Size(200, 18)

    $lblEta = New-Object System.Windows.Forms.Label
    $lblEta.Text = ""
    $lblEta.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblEta.ForeColor = $script:Colors.TextSecondary
    $lblEta.Location = New-Object System.Drawing.Point(200, 28)
    $lblEta.Size = New-Object System.Drawing.Size(120, 18)
    $lblEta.TextAlign = [System.Drawing.ContentAlignment]::TopRight

    $progressPanel.Controls.AddRange(@($lblPercent, $progress, $lblCurrentApp, $lblEta))

    # Keep lblProgress for compatibility but hide it
    $lblProgress = New-Object System.Windows.Forms.Label
    $lblProgress.Text = ""
    $lblProgress.Visible = $false
    
    $btnSelectAll = New-StyledButton -Text "Select All" -Width 95 -Height 30 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnSelectAll.Location = New-Object System.Drawing.Point(20, 468)
    $btnSelectAll.Enabled = $false
    
    $btnSelectNone = New-StyledButton -Text "Select None" -Width 95 -Height 30 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnSelectNone.Location = New-Object System.Drawing.Point(120, 468)
    $btnSelectNone.Enabled = $false
    
    $btnInstall = New-StyledButton -Text "Install Selected" -Width 170 -Height 35 -BackColor $script:Colors.AccentGreen -ForeColor $script:Colors.TextPrimary
    $btnInstall.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnInstall.Location = New-Object System.Drawing.Point(230, 465)
    $btnInstall.Enabled = $false
    
    $btnCancel = New-StyledButton -Text "Cancel" -Width 80 -Height 35 -BackColor $script:Colors.AccentRed -ForeColor $script:Colors.TextPrimary
    $btnCancel.Location = New-Object System.Drawing.Point(405, 465)
    $btnCancel.Visible = $false
    $btnCancel.Add_Click({
        $script:CancelRequested = $true
        $btnCancel.Enabled = $false
        $btnCancel.Text = "Stopping..."
    }.GetNewClosure())

    $btnRetry = New-StyledButton -Text "Retry Failed" -Width 100 -Height 30 -BackColor $script:Colors.AccentOrange -ForeColor $script:Colors.TextDark
    $btnRetry.Location = New-Object System.Drawing.Point(500, 465)
    $btnRetry.Enabled = $false

    $btnManual = New-StyledButton -Text "Open Manual" -Width 100 -Height 30 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnManual.Location = New-Object System.Drawing.Point(605, 465)
    $btnManual.Enabled = $false
    
    $btnBrowseBundle.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        $dlg.Title = "Select programs.json"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtBundle.Text = $dlg.FileName
        }
        $dlg.Dispose()
    }.GetNewClosure())
    
    $btnLoad.Add_Click({
        try {
            $listPrograms.Items.Clear()
            $txtResults.Clear()
            $progress.Value = 0
            $lblPercent.Text = "0%"
            $lblPercent.ForeColor = $script:Colors.AccentGreen
            $lblCurrentApp.Text = ""
            $lblEta.Text = ""
            $script:InstallManualQueue = @()
            $script:InstallFailedQueue = @()
            
            if ([string]::IsNullOrWhiteSpace($txtBundle.Text)) { throw "Select a bundle file first." }
            if (-not (Test-Path $txtBundle.Text)) { throw "File not found." }
            
            $bundle = Get-Content $txtBundle.Text -Raw -Encoding UTF8 | ConvertFrom-Json
            if (-not $bundle.Programs) { throw "Invalid bundle." }
            
            $script:InstallAllPrograms = $bundle.Programs | ForEach-Object {
                $cat = Get-AppCategory -DisplayName $_.DisplayName
                [PSCustomObject]@{ Name = $_.DisplayName; Publisher = $_.Publisher; Category = $cat }
            } | Sort-Object Name

            # Reset category counts
            foreach ($cat in $script:CategoryCounts.Keys) { $script:CategoryCounts[$cat] = 0 }

            # Initialize check states
            $script:InstallCheckStates = @{}
            foreach ($p in $script:InstallAllPrograms) {
                $checked = -not (Test-IsNoiseAppName -DisplayName $p.Name)
                $script:InstallCheckStates[$p.Name] = $checked
                [void]$listPrograms.Items.Add($p.Name, $checked)
                # Count categories
                if ($script:CategoryCounts.ContainsKey($p.Category)) {
                    $script:CategoryCounts[$p.Category]++
                } else {
                    $script:CategoryCounts['Other']++
                }
            }

            # Update category button labels
            foreach ($cat in $script:CategoryButtons.Keys) {
                $count = $script:CategoryCounts[$cat]
                $script:CategoryButtons[$cat].Text = "$cat ($count)"
                if ($count -gt 0) {
                    $script:CategoryButtons[$cat].ForeColor = $script:Colors.TextPrimary
                } else {
                    $script:CategoryButtons[$cat].ForeColor = $script:Colors.TextSecondary
                }
            }

            Write-Log "Loaded $($script:InstallAllPrograms.Count) programs"
            Update-StatusBar "Loaded $($script:InstallAllPrograms.Count) programs"

            # Save bundle path for next time
            Update-Setting -Key 'LastBundlePath' -Value $txtBundle.Text
            
            $btnSelectAll.Enabled = $true
            $btnSelectNone.Enabled = $true
            $btnInstall.Enabled = $true
        } catch {
            Write-Log "Error: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }.GetNewClosure())
    
    $txtFilter.Add_TextChanged({
        $filter = $txtFilter.Text.ToLowerInvariant()
        $listPrograms.Items.Clear()
        foreach ($p in $script:InstallAllPrograms) {
            $searchText = "$($p.Name) $($p.Publisher)".ToLowerInvariant()
            if ([string]::IsNullOrWhiteSpace($filter) -or $searchText.Contains($filter)) {
                $checked = if ($script:InstallCheckStates.ContainsKey($p.Name)) { $script:InstallCheckStates[$p.Name] } else { -not (Test-IsNoiseAppName -DisplayName $p.Name) }
                [void]$listPrograms.Items.Add($p.Name, $checked)
            }
        }
    }.GetNewClosure())
    
    $btnSelectAll.Add_Click({
        for ($i = 0; $i -lt $listPrograms.Items.Count; $i++) {
            if (-not (Test-IsNoiseAppName -DisplayName ([string]$listPrograms.Items[$i]))) {
                $listPrograms.SetItemChecked($i, $true)
            }
        }
    }.GetNewClosure())
    
    $btnSelectNone.Add_Click({
        for ($i = 0; $i -lt $listPrograms.Items.Count; $i++) {
            $listPrograms.SetItemChecked($i, $false)
        }
    }.GetNewClosure())
    
    $btnInstall.Add_Click({
        try {
            $selected = @()
            for ($i = 0; $i -lt $listPrograms.Items.Count; $i++) {
                if ($listPrograms.GetItemChecked($i)) {
                    $selected += [string]$listPrograms.Items[$i]
                }
            }
            if ($selected.Count -eq 0) { throw "Select at least one program." }
            
            $btnInstall.Enabled = $false
            $btnRetry.Enabled = $false
            $btnManual.Enabled = $false
            $script:OperationInProgress = $true
            $script:CancelRequested = $false
            $btnCancel.Visible = $true
            $btnCancel.Enabled = $true
            $btnCancel.Text = "Cancel"
            $txtResults.Clear()
            $progress.Maximum = $selected.Count
            $progress.Value = 0
            $lblPercent.Text = "0%"
            $lblCurrentApp.Text = ""
            $lblEta.Text = ""

            # Timing variables for ETA
            $script:InstallStartTime = Get-Date
            $script:InstallAppTimes = @()
            $script:LastAppStart = Get-Date

            $progressCallback = {
                param($appName, $current, $total, $status)

                # Calculate percentage
                $pct = [Math]::Round(($current / $total) * 100)
                $lblPercent.Text = "$pct%"
                $lblPercent.ForeColor = if ($status -eq 'Failed') { $script:Colors.AccentRed } `
                                        elseif ($status -eq 'Installed') { $script:Colors.AccentGreen } `
                                        else { $script:Colors.AccentOrange }

                # Update progress bar
                $progress.Value = [Math]::Min($current, $progress.Maximum)

                # Current app with status icon
                $statusIcon = switch ($status) {
                    'Processing...' { '...' }
                    'Installing...' { '>>>' }
                    'Installed' { '[OK]' }
                    'Failed' { '[X]' }
                    'Manual' { '[?]' }
                    'Skipped' { '[-]' }
                    default { '...' }
                }
                $lblCurrentApp.Text = "$statusIcon $appName"

                # Track app completion times for ETA
                if ($status -in @('Installed', 'Failed', 'Manual', 'Skipped')) {
                    $appTime = (Get-Date) - $script:LastAppStart
                    $script:InstallAppTimes += $appTime.TotalSeconds
                    $script:LastAppStart = Get-Date
                }

                # Calculate ETA
                $elapsed = (Get-Date) - $script:InstallStartTime
                $elapsedStr = "{0:mm}:{0:ss}" -f $elapsed
                if ($current -gt 0 -and $script:InstallAppTimes.Count -gt 0) {
                    $avgTime = ($script:InstallAppTimes | Measure-Object -Average).Average
                    $remaining = $total - $current
                    $etaSec = [Math]::Round($avgTime * $remaining)
                    $etaSpan = [TimeSpan]::FromSeconds($etaSec)
                    $lblEta.Text = "$elapsedStr | ~$("{0:mm}:{0:ss}" -f $etaSpan) left"
                } else {
                    $lblEta.Text = "$elapsedStr | calculating..."
                }

                # Update results log in real-time
                if ($status -in @('Installed', 'Failed', 'Manual', 'Skipped')) {
                    $txtResults.AppendText("$statusIcon $appName`r`n")
                    $txtResults.SelectionStart = $txtResults.TextLength
                    $txtResults.ScrollToCaret()
                }

                Update-GuiStatus
            }

            $session = Install-ProgramsFromBundle -BundleJsonPath $txtBundle.Text -SelectedDisplayNames $selected -AutoInstallWinget -AutoInstallChocolatey -OnProgress $progressCallback
            
            $results = @($session.Results)
            $script:InstallManualQueue = @($session.ManualQueue)
            $script:InstallFailedQueue = @($session.FailedQueue)
            
            $installed = @($results | Where-Object { $_.Status -eq 'Installed' }).Count
            $verified = @($results | Where-Object { $_.Verified -eq $true }).Count
            $failed = @($results | Where-Object { $_.Status -eq 'Failed' }).Count
            $manual = @($results | Where-Object { $_.Status -eq 'Manual' }).Count
            
            # Final summary
            $totalTime = (Get-Date) - $script:InstallStartTime
            $totalTimeStr = "{0:mm}:{0:ss}" -f $totalTime

            $txtResults.AppendText("`r`n=== COMPLETE ===" + "`r`n")
            $txtResults.AppendText("Installed: $installed ($verified verified)" + "`r`n")
            $txtResults.AppendText("Failed: $failed" + "`r`n")
            $txtResults.AppendText("Manual: $manual" + "`r`n")
            $txtResults.AppendText("Time: $totalTimeStr" + "`r`n")

            # Update progress panel to show completion
            $lblPercent.Text = "100%"
            $lblPercent.ForeColor = if ($failed -eq 0) { $script:Colors.AccentGreen } else { $script:Colors.AccentOrange }
            $lblCurrentApp.Text = "Done!"
            $lblEta.Text = "Total: $totalTimeStr"
            Update-StatusBar "Install complete: $installed installed, $failed failed, $manual manual"
            
            if ($script:InstallFailedQueue.Count -gt 0) { $btnRetry.Enabled = $true }
            if ($script:InstallManualQueue.Count -gt 0) { $btnManual.Enabled = $true }
            
        } catch {
            Write-Log "Error: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $script:OperationInProgress = $false
            $script:CancelRequested = $false
            $btnInstall.Enabled = $true
            $btnCancel.Visible = $false
        }
    }.GetNewClosure())

    $btnRetry.Add_Click({
        for ($i = 0; $i -lt $listPrograms.Items.Count; $i++) {
            $name = [string]$listPrograms.Items[$i]
            $listPrograms.SetItemChecked($i, ($script:InstallFailedQueue -contains $name))
        }
        $script:InstallFailedQueue = @()
        $btnRetry.Enabled = $false
        $btnInstall.PerformClick()
    }.GetNewClosure())
    
    $btnManual.Add_Click({
        $count = $script:InstallManualQueue.Count
        if ($count -gt 10) {
            $confirm = [System.Windows.Forms.MessageBox]::Show("Open $count browser tabs for manual downloads?", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        foreach ($q in $script:InstallManualQueue) {
            Open-WebSearch -Query $q
        }
        Write-Log "Opened $count manual searches"
        $script:InstallManualQueue = @()
        $btnManual.Enabled = $false
    }.GetNewClosure())
    
    $script:ContentPanel.Controls.AddRange(@($title, $subtitle, $lblBundle, $txtBundle, $btnBrowseBundle, $btnLoad, $lblFilter, $txtFilter, $categoryPanel, $listPrograms, $lblResults, $txtResults, $progressPanel, $btnSelectAll, $btnSelectNone, $btnInstall, $btnCancel, $btnRetry, $btnManual))
}

function Show-FilesPage {
    if ($script:OperationInProgress) { return }
    Clear-ContentPanel
    Set-ActiveSidebarButton -Page "Files"
    Update-StatusBar "Ready for files migration"

    # Store folder sizes for recalculation
    $script:FilesFolderSizes = @{}

    # --- Title ---
    $title = New-Object System.Windows.Forms.Label
    $title.Text = "User Files Migration"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $script:Colors.AccentOrange
    $title.Location = New-Object System.Drawing.Point(20, 15)
    $title.AutoSize = $true

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Backup user files and browser bookmarks for transfer"
    $subtitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitle.ForeColor = $script:Colors.TextSecondary
    $subtitle.Location = New-Object System.Drawing.Point(20, 50)
    $subtitle.AutoSize = $true

    # --- LEFT COLUMN: Folder Selection ---
    $lblFolders = New-Object System.Windows.Forms.Label
    $lblFolders.Text = "SELECT FOLDERS"
    $lblFolders.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblFolders.ForeColor = $script:Colors.TextSecondary
    $lblFolders.Location = New-Object System.Drawing.Point(20, 85)
    $lblFolders.AutoSize = $true

    $listFolders = New-Object System.Windows.Forms.CheckedListBox
    $listFolders.Location = New-Object System.Drawing.Point(20, 108)
    $listFolders.Size = New-Object System.Drawing.Size(380, 170)
    $listFolders.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $listFolders.BackColor = $script:Colors.InputBg
    $listFolders.ForeColor = $script:Colors.TextPrimary
    $listFolders.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $listFolders.CheckOnClick = $true

    $folderNames = @('Documents', 'Downloads', 'Pictures', 'Videos', 'Music', 'Desktop')
    foreach ($name in $folderNames) {
        [void]$listFolders.Items.Add("$name (calculating...)", $true)
    }

    $lblTotalSize = New-Object System.Windows.Forms.Label
    $lblTotalSize.Text = "Total Selected: calculating..."
    $lblTotalSize.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblTotalSize.ForeColor = $script:Colors.AccentOrange
    $lblTotalSize.Location = New-Object System.Drawing.Point(20, 285)
    $lblTotalSize.Size = New-Object System.Drawing.Size(380, 20)

    # --- RIGHT COLUMN: Browser Bookmarks ---
    $lblBrowsers = New-Object System.Windows.Forms.Label
    $lblBrowsers.Text = "BROWSER BOOKMARKS"
    $lblBrowsers.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblBrowsers.ForeColor = $script:Colors.TextSecondary
    $lblBrowsers.Location = New-Object System.Drawing.Point(420, 85)
    $lblBrowsers.AutoSize = $true

    $chkChrome = New-Object System.Windows.Forms.CheckBox
    $chkChrome.Text = "Chrome"
    $chkChrome.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $chkChrome.ForeColor = $script:Colors.TextPrimary
    $chkChrome.BackColor = $script:Colors.ContentBg
    $chkChrome.Location = New-Object System.Drawing.Point(420, 108)
    $chkChrome.AutoSize = $true

    $chkFirefox = New-Object System.Windows.Forms.CheckBox
    $chkFirefox.Text = "Firefox"
    $chkFirefox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $chkFirefox.ForeColor = $script:Colors.TextPrimary
    $chkFirefox.BackColor = $script:Colors.ContentBg
    $chkFirefox.Location = New-Object System.Drawing.Point(420, 133)
    $chkFirefox.AutoSize = $true

    $chkEdge = New-Object System.Windows.Forms.CheckBox
    $chkEdge.Text = "Edge"
    $chkEdge.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $chkEdge.ForeColor = $script:Colors.TextPrimary
    $chkEdge.BackColor = $script:Colors.ContentBg
    $chkEdge.Location = New-Object System.Drawing.Point(420, 158)
    $chkEdge.AutoSize = $true

    # Detect installed browsers and enable/disable accordingly
    if (-not (Test-BrowserInstalled -BrowserName 'Chrome')) {
        $chkChrome.Enabled = $false; $chkChrome.ForeColor = $script:Colors.TextSecondary; $chkChrome.Text = "Chrome (not found)"
    } else { $chkChrome.Checked = $true }

    if (-not (Test-BrowserInstalled -BrowserName 'Firefox')) {
        $chkFirefox.Enabled = $false; $chkFirefox.ForeColor = $script:Colors.TextSecondary; $chkFirefox.Text = "Firefox (not found)"
    } else { $chkFirefox.Checked = $true }

    if (-not (Test-BrowserInstalled -BrowserName 'Edge')) {
        $chkEdge.Enabled = $false; $chkEdge.ForeColor = $script:Colors.TextSecondary; $chkEdge.Text = "Edge (not found)"
    } else { $chkEdge.Checked = $true }

    # --- RIGHT COLUMN: Output Mode ---
    $lblMode = New-Object System.Windows.Forms.Label
    $lblMode.Text = "OUTPUT MODE"
    $lblMode.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblMode.ForeColor = $script:Colors.TextSecondary
    $lblMode.Location = New-Object System.Drawing.Point(420, 200)
    $lblMode.AutoSize = $true

    $rdoZip = New-Object System.Windows.Forms.RadioButton
    $rdoZip.Text = "ZIP Archive (recommended)"
    $rdoZip.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $rdoZip.ForeColor = $script:Colors.TextPrimary
    $rdoZip.BackColor = $script:Colors.ContentBg
    $rdoZip.Location = New-Object System.Drawing.Point(420, 223)
    $rdoZip.AutoSize = $true

    $rdoDirect = New-Object System.Windows.Forms.RadioButton
    $rdoDirect.Text = "Direct Copy (faster)"
    $rdoDirect.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $rdoDirect.ForeColor = $script:Colors.TextPrimary
    $rdoDirect.BackColor = $script:Colors.ContentBg
    $rdoDirect.Location = New-Object System.Drawing.Point(420, 248)
    $rdoDirect.AutoSize = $true

    # Pre-fill mode from settings
    if ($script:UserSettings.FilesMigrationMode -eq "direct") { $rdoDirect.Checked = $true } else { $rdoZip.Checked = $true }

    $rdoZip.Add_CheckedChanged({
        if ($rdoZip.Checked) { Update-Setting -Key 'FilesMigrationMode' -Value 'zip' }
    }.GetNewClosure())
    $rdoDirect.Add_CheckedChanged({
        if ($rdoDirect.Checked) { Update-Setting -Key 'FilesMigrationMode' -Value 'direct' }
    }.GetNewClosure())

    # --- FULL WIDTH: Save Location ---
    $lblSave = New-Object System.Windows.Forms.Label
    $lblSave.Text = "SAVE LOCATION"
    $lblSave.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblSave.ForeColor = $script:Colors.TextSecondary
    $lblSave.Location = New-Object System.Drawing.Point(20, 320)
    $lblSave.AutoSize = $true

    $txtSaveLocation = New-Object System.Windows.Forms.TextBox
    $txtSaveLocation.Location = New-Object System.Drawing.Point(20, 343)
    $txtSaveLocation.Size = New-Object System.Drawing.Size(580, 28)
    $txtSaveLocation.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $txtSaveLocation.BackColor = $script:Colors.InputBg
    $txtSaveLocation.ForeColor = $script:Colors.TextPrimary
    $txtSaveLocation.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    if ($script:UserSettings.LastFilesOutputFolder) {
        $txtSaveLocation.Text = $script:UserSettings.LastFilesOutputFolder
    } else {
        $txtSaveLocation.Text = $script:ScriptDir
    }

    $btnBrowseSave = New-StyledButton -Text "Browse..." -Width 100 -Height 28 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnBrowseSave.Location = New-Object System.Drawing.Point(610, 343)
    $btnBrowseSave.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Choose where to save migrated files"
        $dlg.ShowNewFolderButton = $true
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtSaveLocation.Text = $dlg.SelectedPath
            Update-Setting -Key 'LastFilesOutputFolder' -Value $dlg.SelectedPath
        }
        $dlg.Dispose()
    }.GetNewClosure())

    # --- Start Migration Button ---
    $btnStartMigration = New-StyledButton -Text "Start Migration" -Width 220 -Height 45 -BackColor $script:Colors.AccentOrange -ForeColor $script:Colors.TextDark
    $btnStartMigration.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $btnStartMigration.Location = New-Object System.Drawing.Point(20, 390)

    $btnCancelMigration = New-StyledButton -Text "Cancel" -Width 100 -Height 45 -BackColor $script:Colors.AccentRed -ForeColor $script:Colors.TextPrimary
    $btnCancelMigration.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnCancelMigration.Location = New-Object System.Drawing.Point(250, 390)
    $btnCancelMigration.Visible = $false
    $btnCancelMigration.Add_Click({
        $script:CancelRequested = $true
        $btnCancelMigration.Enabled = $false
        $btnCancelMigration.Text = "Stopping..."
    }.GetNewClosure())

    # --- Progress Panel (hidden initially) ---
    $progressPanel = New-Object System.Windows.Forms.Panel
    $progressPanel.Location = New-Object System.Drawing.Point(20, 445)
    $progressPanel.Size = New-Object System.Drawing.Size(720, 55)
    $progressPanel.BackColor = $script:Colors.ContentBg
    $progressPanel.Visible = $false

    $lblFilesPercent = New-Object System.Windows.Forms.Label
    $lblFilesPercent.Text = "0%"
    $lblFilesPercent.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblFilesPercent.ForeColor = $script:Colors.AccentOrange
    $lblFilesPercent.Location = New-Object System.Drawing.Point(0, 0)
    $lblFilesPercent.Size = New-Object System.Drawing.Size(55, 25)

    $filesProgress = New-Object ModernProgressBar
    $filesProgress.Location = New-Object System.Drawing.Point(60, 5)
    $filesProgress.Size = New-Object System.Drawing.Size(400, 14)
    $filesProgress.BarColor = $script:Colors.AccentOrange
    $filesProgress.TrackColor = $script:Colors.CardBg

    $lblCurrentFolder = New-Object System.Windows.Forms.Label
    $lblCurrentFolder.Text = ""
    $lblCurrentFolder.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblCurrentFolder.ForeColor = $script:Colors.TextPrimary
    $lblCurrentFolder.Location = New-Object System.Drawing.Point(0, 28)
    $lblCurrentFolder.Size = New-Object System.Drawing.Size(350, 18)

    $lblFilesEta = New-Object System.Windows.Forms.Label
    $lblFilesEta.Text = ""
    $lblFilesEta.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblFilesEta.ForeColor = $script:Colors.TextSecondary
    $lblFilesEta.Location = New-Object System.Drawing.Point(470, 5)
    $lblFilesEta.Size = New-Object System.Drawing.Size(250, 18)
    $lblFilesEta.TextAlign = [System.Drawing.ContentAlignment]::TopRight

    $progressPanel.Controls.AddRange(@($lblFilesPercent, $filesProgress, $lblCurrentFolder, $lblFilesEta))

    # --- Results Panel (hidden initially) ---
    $resultsPanel = New-Object System.Windows.Forms.Panel
    $resultsPanel.Location = New-Object System.Drawing.Point(20, 445)
    $resultsPanel.Size = New-Object System.Drawing.Size(720, 60)
    $resultsPanel.BackColor = $script:Colors.ContentBg
    $resultsPanel.Visible = $false

    $lblResultText = New-Object System.Windows.Forms.Label
    $lblResultText.Text = ""
    $lblResultText.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblResultText.ForeColor = $script:Colors.AccentGreen
    $lblResultText.Location = New-Object System.Drawing.Point(0, 0)
    $lblResultText.Size = New-Object System.Drawing.Size(500, 25)

    $lblResultDetails = New-Object System.Windows.Forms.Label
    $lblResultDetails.Text = ""
    $lblResultDetails.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblResultDetails.ForeColor = $script:Colors.TextSecondary
    $lblResultDetails.Location = New-Object System.Drawing.Point(0, 28)
    $lblResultDetails.Size = New-Object System.Drawing.Size(500, 18)

    $btnOpenFolder = New-StyledButton -Text "Open Folder" -Width 110 -Height 35 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnOpenFolder.Location = New-Object System.Drawing.Point(610, 8)
    $btnOpenFolder.Visible = $false
    $btnOpenFolder.Add_Click({
        if ($script:LastMigrationPath) {
            $openPath = if (Test-Path $script:LastMigrationPath -PathType Leaf) {
                Split-Path -Parent $script:LastMigrationPath
            } else { $script:LastMigrationPath }
            try { Start-Process "explorer.exe" -ArgumentList $openPath } catch { Write-Log "Could not open folder: $($_.Exception.Message)" -Level 'WARN' }
        }
    }.GetNewClosure())

    $resultsPanel.Controls.AddRange(@($lblResultText, $lblResultDetails, $btnOpenFolder))

    # --- ItemCheck event: recalculate total size ---
    $listFolders.Add_ItemCheck({
        param($sender, $e)
        $totalSelected = [long]0
        for ($i = 0; $i -lt $listFolders.Items.Count; $i++) {
            $isChecked = if ($i -eq $e.Index) { $e.NewValue -eq [System.Windows.Forms.CheckState]::Checked } else { $listFolders.GetItemChecked($i) }
            if ($isChecked -and $script:FilesFolderSizes.ContainsKey($folderNames[$i])) {
                $totalSelected += $script:FilesFolderSizes[$folderNames[$i]]
            }
        }
        $lblTotalSize.Text = "Total Selected: $(Format-FileSize $totalSelected)"
    }.GetNewClosure())

    # --- Start Migration Click ---
    $btnStartMigration.Add_Click({
        try {
            # Gather selected folders
            $selected = @()
            for ($i = 0; $i -lt $listFolders.Items.Count; $i++) {
                if ($listFolders.GetItemChecked($i)) { $selected += $folderNames[$i] }
            }
            if ($selected.Count -eq 0) { throw "Select at least one folder." }
            if ([string]::IsNullOrWhiteSpace($txtSaveLocation.Text)) { throw "Choose a save location." }

            # Gather selected browsers
            $browsers = @()
            if ($chkChrome.Checked -and $chkChrome.Enabled) { $browsers += 'Chrome' }
            if ($chkFirefox.Checked -and $chkFirefox.Enabled) { $browsers += 'Firefox' }
            if ($chkEdge.Checked -and $chkEdge.Enabled) { $browsers += 'Edge' }

            $mode = if ($rdoZip.Checked) { "zip" } else { "direct" }

            # Confirmation dialog
            $totalSize = [long]0
            foreach ($f in $selected) { if ($script:FilesFolderSizes.ContainsKey($f)) { $totalSize += $script:FilesFolderSizes[$f] } }
            $sizeStr = Format-FileSize $totalSize
            $confirm = [System.Windows.Forms.MessageBox]::Show("Migrate $($selected.Count) folder(s) ($sizeStr) to $($txtSaveLocation.Text)?`n`nMode: $mode", "Confirm Migration", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

            # Save settings
            Update-Setting -Key 'LastFilesOutputFolder' -Value $txtSaveLocation.Text

            # Disable controls during migration
            $script:OperationInProgress = $true
            $script:CancelRequested = $false
            $btnStartMigration.Enabled = $false
            $btnStartMigration.Text = "Migrating..."
            $btnCancelMigration.Visible = $true
            $btnCancelMigration.Enabled = $true
            $btnCancelMigration.Text = "Cancel"
            $listFolders.Enabled = $false
            $resultsPanel.Visible = $false
            $btnOpenFolder.Visible = $false
            $progressPanel.Visible = $true
            $filesProgress.Value = 0
            $lblFilesPercent.Text = "0%"
            $lblCurrentFolder.Text = ""
            $lblFilesEta.Text = ""
            Update-GuiStatus

            $script:MigrationStartTime = Get-Date

            # Progress callback
            $progressCallback = {
                param($folderName, $bytesCopied, $totalBytesAll, $currentStep, $totalSteps, $status)

                $pct = if ($totalBytesAll -gt 0) { [Math]::Min([Math]::Round(($bytesCopied / $totalBytesAll) * 100), 100) } else { [Math]::Round(($currentStep / $totalSteps) * 100) }
                $lblFilesPercent.Text = "$pct%"
                $filesProgress.Maximum = 100
                $filesProgress.Value = [Math]::Min($pct, 100)

                $statusIcon = if ($status -eq 'Done') { '[OK]' } elseif ($status -eq 'Compressing...') { '[>>]' } else { '...' }
                $lblCurrentFolder.Text = "$statusIcon $folderName"

                $elapsed = (Get-Date) - $script:MigrationStartTime
                $elapsedStr = "{0:mm}:{0:ss}" -f $elapsed
                if ($pct -gt 0 -and $pct -lt 100) {
                    $etaSec = [Math]::Round(($elapsed.TotalSeconds / $pct) * (100 - $pct))
                    $etaSpan = [TimeSpan]::FromSeconds($etaSec)
                    $lblFilesEta.Text = "$elapsedStr elapsed | ~$("{0:mm}:{0:ss}" -f $etaSpan) left"
                } else {
                    $lblFilesEta.Text = "$elapsedStr elapsed"
                }

                Update-GuiStatus
            }

            $clientName = $env:COMPUTERNAME
            $result = Start-FilesMigration -ClientName $clientName -OutputFolder $txtSaveLocation.Text -SelectedFolders $selected -Mode $mode -SelectedBrowsers $browsers -PrecomputedSizes $script:FilesFolderSizes -OnProgress $progressCallback

            # Show results
            $progressPanel.Visible = $false
            $resultsPanel.Visible = $true
            $script:LastMigrationPath = $result.Path

            $totalTime = (Get-Date) - $script:MigrationStartTime
            $totalTimeStr = "{0:mm}:{0:ss}" -f $totalTime

            if ($result.Success) {
                $lblResultText.Text = "Migration Complete!"
                $lblResultText.ForeColor = $script:Colors.AccentGreen
            } else {
                $lblResultText.Text = "Migration Complete (with $($result.Errors.Count) warning(s))"
                $lblResultText.ForeColor = $script:Colors.AccentOrange
            }

            $details = "$($result.FoldersCopied) folders | $(Format-FileSize $result.TotalBytes) | $totalTimeStr"
            if ($result.BookmarksExported.Count -gt 0) { $details += " | Bookmarks: $($result.BookmarksExported -join ', ')" }
            $lblResultDetails.Text = $details
            $btnOpenFolder.Visible = $true

            Update-StatusBar "Migration complete: $(Format-FileSize $result.TotalBytes) in $totalTimeStr"

        } catch {
            Write-Log "Error: $($_.Exception.Message)" -Level 'ERROR'
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Migration Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $progressPanel.Visible = $false
        } finally {
            $script:OperationInProgress = $false
            $script:CancelRequested = $false
            $btnStartMigration.Enabled = $true
            $btnStartMigration.Text = "Start Migration"
            $btnCancelMigration.Visible = $false
            $listFolders.Enabled = $true
        }
    }.GetNewClosure())

    # Add all controls
    $script:ContentPanel.Controls.AddRange(@(
        $title, $subtitle,
        $lblFolders, $listFolders, $lblTotalSize,
        $lblBrowsers, $chkChrome, $chkFirefox, $chkEdge,
        $lblMode, $rdoZip, $rdoDirect,
        $lblSave, $txtSaveLocation, $btnBrowseSave,
        $btnStartMigration, $btnCancelMigration,
        $progressPanel, $resultsPanel
    ))

    # Folder size calculation (updates UI incrementally via DoEvents)
    $sizeCallback = {
        param($folderName, $sizeBytes, $fileCount)
        $script:FilesFolderSizes[$folderName] = $sizeBytes
        $idx = [Array]::IndexOf($folderNames, $folderName)
        if ($idx -ge 0) {
            $listFolders.Items[$idx] = "$folderName ($(Format-FileSize $sizeBytes))"
            $listFolders.SetItemChecked($idx, $true)
        }
        # Update total
        $totalSelected = [long]0
        for ($i = 0; $i -lt $listFolders.Items.Count; $i++) {
            if ($listFolders.GetItemChecked($i) -and $script:FilesFolderSizes.ContainsKey($folderNames[$i])) {
                $totalSelected += $script:FilesFolderSizes[$folderNames[$i]]
            }
        }
        $lblTotalSize.Text = "Total Selected: $(Format-FileSize $totalSelected)"
    }.GetNewClosure()

    # Run size calculation (updates UI incrementally via Update-GuiStatus)
    $null = Get-FolderSizes -FolderNames $folderNames -OnProgress $sizeCallback
}

function Show-TransferPage {
    if ($script:OperationInProgress) { return }
    Clear-ContentPanel
    Set-ActiveSidebarButton -Page "Transfer"
    Update-StatusBar "Remote Transfer"

    $accentColor = $script:Colors.AccentPink

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Remote Transfer"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $accentColor
    $title.Location = New-Object System.Drawing.Point(20, 15)
    $title.AutoSize = $true

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Transfer files between PCs over your local network"
    $subtitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitle.ForeColor = $script:Colors.TextSecondary
    $subtitle.Location = New-Object System.Drawing.Point(20, 50)
    $subtitle.AutoSize = $true

    # --- CARD 1: Send (HTTP Server) ---
    $card1 = New-Object ModernCard
    $card1.HeaderText = "Send (HTTP Server)"
    $card1.HeaderColor = $accentColor
    $card1.FillColor = $script:Colors.CardBg
    $card1.BorderColor = $script:Colors.BorderColor
    $card1.Location = New-Object System.Drawing.Point(20, 80)
    $card1.Size = New-Object System.Drawing.Size(360, 130)

    $txtServerFolder = New-Object System.Windows.Forms.TextBox
    $txtServerFolder.Location = New-Object System.Drawing.Point(10, 25)
    $txtServerFolder.Size = New-Object System.Drawing.Size(250, 24)
    $txtServerFolder.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtServerFolder.BackColor = $script:Colors.InputBg
    $txtServerFolder.ForeColor = $script:Colors.TextPrimary
    $txtServerFolder.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $btnBrowseServer = New-StyledButton -Text "..." -Width 40 -Height 24 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnBrowseServer.Location = New-Object System.Drawing.Point(265, 25)
    $btnBrowseServer.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select folder to serve"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtServerFolder.Text = $dlg.SelectedPath }
        $dlg.Dispose()
    }.GetNewClosure())

    $lblServerUrl = New-Object System.Windows.Forms.Label
    $lblServerUrl.Text = "URL will appear here"
    $lblServerUrl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblServerUrl.ForeColor = $script:Colors.TextSecondary
    $lblServerUrl.Location = New-Object System.Drawing.Point(10, 58)
    $lblServerUrl.Size = New-Object System.Drawing.Size(340, 20)

    $btnStartServer = New-StyledButton -Text "Start Server" -Width 120 -Height 30 -BackColor $accentColor -ForeColor $script:Colors.TextPrimary
    $btnStartServer.Location = New-Object System.Drawing.Point(10, 85)

    $btnStopServer = New-StyledButton -Text "Stop" -Width 80 -Height 30 -BackColor $script:Colors.AccentRed -ForeColor $script:Colors.TextPrimary
    $btnStopServer.Location = New-Object System.Drawing.Point(140, 85)
    $btnStopServer.Visible = $false

    $card1.Controls.AddRange(@($txtServerFolder, $btnBrowseServer, $lblServerUrl, $btnStartServer, $btnStopServer))

    # --- CARD 2: Send (Direct TCP) ---
    $card2 = New-Object ModernCard
    $card2.HeaderText = "Send (Direct TCP)"
    $card2.HeaderColor = $accentColor
    $card2.FillColor = $script:Colors.CardBg
    $card2.BorderColor = $script:Colors.BorderColor
    $card2.Location = New-Object System.Drawing.Point(400, 80)
    $card2.Size = New-Object System.Drawing.Size(360, 130)

    $txtTcpFolder = New-Object System.Windows.Forms.TextBox
    $txtTcpFolder.Location = New-Object System.Drawing.Point(10, 25)
    $txtTcpFolder.Size = New-Object System.Drawing.Size(250, 24)
    $txtTcpFolder.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtTcpFolder.BackColor = $script:Colors.InputBg
    $txtTcpFolder.ForeColor = $script:Colors.TextPrimary
    $txtTcpFolder.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $btnBrowseTcp = New-StyledButton -Text "..." -Width 40 -Height 24 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnBrowseTcp.Location = New-Object System.Drawing.Point(265, 25)
    $btnBrowseTcp.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select folder to send"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtTcpFolder.Text = $dlg.SelectedPath }
        $dlg.Dispose()
    }.GetNewClosure())

    $lblTcpInfo = New-Object System.Windows.Forms.Label
    $lblTcpInfo.Text = "Receiver must connect to your IP"
    $lblTcpInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblTcpInfo.ForeColor = $script:Colors.TextSecondary
    $lblTcpInfo.Location = New-Object System.Drawing.Point(10, 58)
    $lblTcpInfo.Size = New-Object System.Drawing.Size(340, 20)

    $btnStartTcp = New-StyledButton -Text "Start Sender" -Width 120 -Height 30 -BackColor $accentColor -ForeColor $script:Colors.TextPrimary
    $btnStartTcp.Location = New-Object System.Drawing.Point(10, 85)

    $card2.Controls.AddRange(@($txtTcpFolder, $btnBrowseTcp, $lblTcpInfo, $btnStartTcp))

    # --- CARD 3: Receive ---
    $card3 = New-Object ModernCard
    $card3.HeaderText = "Receive"
    $card3.HeaderColor = $accentColor
    $card3.FillColor = $script:Colors.CardBg
    $card3.BorderColor = $script:Colors.BorderColor
    $card3.Location = New-Object System.Drawing.Point(20, 220)
    $card3.Size = New-Object System.Drawing.Size(360, 130)

    $lblIP = New-Object System.Windows.Forms.Label
    $lblIP.Text = "Sender IP:"
    $lblIP.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblIP.ForeColor = $script:Colors.TextPrimary
    $lblIP.Location = New-Object System.Drawing.Point(10, 25)
    $lblIP.AutoSize = $true

    $txtRemoteIP = New-Object System.Windows.Forms.TextBox
    $txtRemoteIP.Location = New-Object System.Drawing.Point(80, 23)
    $txtRemoteIP.Size = New-Object System.Drawing.Size(150, 24)
    $txtRemoteIP.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtRemoteIP.BackColor = $script:Colors.InputBg
    $txtRemoteIP.ForeColor = $script:Colors.TextPrimary
    $txtRemoteIP.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $lblSaveRecv = New-Object System.Windows.Forms.Label
    $lblSaveRecv.Text = "Save to:"
    $lblSaveRecv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSaveRecv.ForeColor = $script:Colors.TextPrimary
    $lblSaveRecv.Location = New-Object System.Drawing.Point(10, 55)
    $lblSaveRecv.AutoSize = $true

    $txtRecvFolder = New-Object System.Windows.Forms.TextBox
    $txtRecvFolder.Location = New-Object System.Drawing.Point(80, 53)
    $txtRecvFolder.Size = New-Object System.Drawing.Size(200, 24)
    $txtRecvFolder.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtRecvFolder.BackColor = $script:Colors.InputBg
    $txtRecvFolder.ForeColor = $script:Colors.TextPrimary
    $txtRecvFolder.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $btnBrowseRecv = New-StyledButton -Text "..." -Width 40 -Height 24 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnBrowseRecv.Location = New-Object System.Drawing.Point(285, 53)
    $btnBrowseRecv.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Choose where to save received files"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtRecvFolder.Text = $dlg.SelectedPath }
        $dlg.Dispose()
    }.GetNewClosure())

    $btnDownloadHttp = New-StyledButton -Text "Download (HTTP)" -Width 140 -Height 30 -BackColor $accentColor -ForeColor $script:Colors.TextPrimary
    $btnDownloadHttp.Location = New-Object System.Drawing.Point(10, 88)

    $btnReceiveTcp = New-StyledButton -Text "Receive (TCP)" -Width 130 -Height 30 -BackColor $accentColor -ForeColor $script:Colors.TextPrimary
    $btnReceiveTcp.Location = New-Object System.Drawing.Point(160, 88)

    $card3.Controls.AddRange(@($lblIP, $txtRemoteIP, $lblSaveRecv, $txtRecvFolder, $btnBrowseRecv, $btnDownloadHttp, $btnReceiveTcp))

    # --- CARD 4: Shared Folder ---
    $card4 = New-Object ModernCard
    $card4.HeaderText = "Shared Folder"
    $card4.HeaderColor = $accentColor
    $card4.FillColor = $script:Colors.CardBg
    $card4.BorderColor = $script:Colors.BorderColor
    $card4.Location = New-Object System.Drawing.Point(400, 220)
    $card4.Size = New-Object System.Drawing.Size(360, 130)

    $lblShareSrc = New-Object System.Windows.Forms.Label
    $lblShareSrc.Text = "Source:"
    $lblShareSrc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblShareSrc.ForeColor = $script:Colors.TextPrimary
    $lblShareSrc.Location = New-Object System.Drawing.Point(10, 25)
    $lblShareSrc.AutoSize = $true

    $txtShareSrc = New-Object System.Windows.Forms.TextBox
    $txtShareSrc.Location = New-Object System.Drawing.Point(70, 23)
    $txtShareSrc.Size = New-Object System.Drawing.Size(210, 24)
    $txtShareSrc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtShareSrc.BackColor = $script:Colors.InputBg
    $txtShareSrc.ForeColor = $script:Colors.TextPrimary
    $txtShareSrc.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $btnBrowseShareSrc = New-StyledButton -Text "..." -Width 40 -Height 24 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnBrowseShareSrc.Location = New-Object System.Drawing.Point(285, 23)
    $btnBrowseShareSrc.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select source folder"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtShareSrc.Text = $dlg.SelectedPath }
        $dlg.Dispose()
    }.GetNewClosure())

    $lblShareDest = New-Object System.Windows.Forms.Label
    $lblShareDest.Text = "UNC Dest:"
    $lblShareDest.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblShareDest.ForeColor = $script:Colors.TextPrimary
    $lblShareDest.Location = New-Object System.Drawing.Point(10, 55)
    $lblShareDest.AutoSize = $true

    $txtShareDest = New-Object System.Windows.Forms.TextBox
    $txtShareDest.Location = New-Object System.Drawing.Point(70, 53)
    $txtShareDest.Size = New-Object System.Drawing.Size(210, 24)
    $txtShareDest.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtShareDest.BackColor = $script:Colors.InputBg
    $txtShareDest.ForeColor = $script:Colors.TextPrimary
    $txtShareDest.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtShareDest.Text = "\\PC-NAME\Share"

    $btnBrowseShareDest = New-StyledButton -Text "..." -Width 40 -Height 24 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnBrowseShareDest.Location = New-Object System.Drawing.Point(285, 53)
    $btnBrowseShareDest.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select UNC destination"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtShareDest.Text = $dlg.SelectedPath }
        $dlg.Dispose()
    }.GetNewClosure())

    $btnCopyShare = New-StyledButton -Text "Copy" -Width 120 -Height 30 -BackColor $accentColor -ForeColor $script:Colors.TextPrimary
    $btnCopyShare.Location = New-Object System.Drawing.Point(10, 88)

    $card4.Controls.AddRange(@($lblShareSrc, $txtShareSrc, $btnBrowseShareSrc, $lblShareDest, $txtShareDest, $btnBrowseShareDest, $btnCopyShare))

    # --- Status Panel ---
    $statusGroup = New-Object ModernCard
    $statusGroup.HeaderText = "Transfer Status"
    $statusGroup.HeaderColor = $script:Colors.TextSecondary
    $statusGroup.FillColor = $script:Colors.CardBg
    $statusGroup.BorderColor = $script:Colors.BorderColor
    $statusGroup.Location = New-Object System.Drawing.Point(20, 360)
    $statusGroup.Size = New-Object System.Drawing.Size(740, 150)

    $txtTransferLog = New-Object System.Windows.Forms.TextBox
    $txtTransferLog.Location = New-Object System.Drawing.Point(10, 22)
    $txtTransferLog.Size = New-Object System.Drawing.Size(610, 80)
    $txtTransferLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtTransferLog.BackColor = $script:Colors.InputBg
    $txtTransferLog.ForeColor = $script:Colors.TextSecondary
    $txtTransferLog.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtTransferLog.Multiline = $true
    $txtTransferLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtTransferLog.ReadOnly = $true

    $transferProgress = New-Object ModernProgressBar
    $transferProgress.Location = New-Object System.Drawing.Point(10, 110)
    $transferProgress.Size = New-Object System.Drawing.Size(610, 14)
    $transferProgress.BarColor = $script:Colors.AccentPink
    $transferProgress.TrackColor = $script:Colors.CardBg

    $btnStopTransfer = New-StyledButton -Text "Stop" -Width 90 -Height 35 -BackColor $script:Colors.AccentRed -ForeColor $script:Colors.TextPrimary
    $btnStopTransfer.Location = New-Object System.Drawing.Point(635, 22)
    $btnStopTransfer.Enabled = $false

    $statusGroup.Controls.AddRange(@($txtTransferLog, $transferProgress, $btnStopTransfer))

    # --- Helper: log to transfer panel ---
    $logTransfer = {
        param([string]$Message)
        $txtTransferLog.AppendText("$Message`r`n")
        $txtTransferLog.SelectionStart = $txtTransferLog.TextLength
        $txtTransferLog.ScrollToCaret()
        Update-GuiStatus
    }.GetNewClosure()

    # --- Event Handlers ---
    $btnStopTransfer.Add_Click({
        $script:CancelRequested = $true
        Stop-TransferServer
        $btnStopTransfer.Enabled = $false
        & $logTransfer "Transfer stopped by user."
    }.GetNewClosure())

    $btnStartServer.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($txtServerFolder.Text)) { throw "Select a folder to serve." }
            if (-not (Test-Path $txtServerFolder.Text)) { throw "Folder not found." }

            $script:OperationInProgress = $true
            $script:CancelRequested = $false
            $btnStartServer.Enabled = $false
            $btnStopServer.Visible = $true
            $btnStopTransfer.Enabled = $true
            $txtTransferLog.Clear()
            & $logTransfer "Starting HTTP server..."

            $serverProgress = {
                param($phase, $current, $total, $status)
                & $logTransfer $status
                if ($total -gt 0) {
                    $transferProgress.Maximum = $total
                    $transferProgress.Value = [Math]::Min($current, $total)
                }
            }.GetNewClosure()

            $localIP = Get-PrimaryLocalIP
            $url = "http://${localIP}:8642/"
            $lblServerUrl.Text = $url
            $lblServerUrl.ForeColor = $script:Colors.AccentGreen
            & $logTransfer "Server running at $url"
            & $logTransfer "Open this URL on the other PC to download files."

            Start-TransferServer -SourceFolder $txtServerFolder.Text -Port 8642 -OnProgress $serverProgress

        } catch {
            & $logTransfer "Error: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $script:OperationInProgress = $false
            $btnStartServer.Enabled = $true
            $btnStopServer.Visible = $false
            $btnStopTransfer.Enabled = $false
            $lblServerUrl.Text = "Server stopped"
            $lblServerUrl.ForeColor = $script:Colors.TextSecondary
        }
    }.GetNewClosure())

    $btnStopServer.Add_Click({
        Stop-TransferServer
        $btnStopServer.Visible = $false
        $btnStartServer.Enabled = $true
        $lblServerUrl.Text = "Server stopped"
        $lblServerUrl.ForeColor = $script:Colors.TextSecondary
    }.GetNewClosure())

    $btnStartTcp.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($txtTcpFolder.Text)) { throw "Select a folder to send." }
            if (-not (Test-Path $txtTcpFolder.Text)) { throw "Folder not found." }

            $script:OperationInProgress = $true
            $script:CancelRequested = $false
            $btnStartTcp.Enabled = $false
            $btnStopTransfer.Enabled = $true
            $txtTransferLog.Clear()
            & $logTransfer "Starting TCP sender..."

            $tcpProgress = {
                param($phase, $current, $total, $status)
                & $logTransfer $status
                if ($total -gt 0) {
                    $transferProgress.Maximum = 100
                    $transferProgress.Value = [Math]::Min([Math]::Round(($current / $total) * 100), 100)
                }
            }.GetNewClosure()

            $result = Start-DirectSend -SourceFolder $txtTcpFolder.Text -Port 8643 -OnProgress $tcpProgress
            if ($result.Success) {
                & $logTransfer "Send complete: $($result.FilesSent) files, $(Format-FileSize $result.BytesSent)"
            } else {
                & $logTransfer "Send failed: $($result.Errors -join '; ')"
            }
        } catch {
            & $logTransfer "Error: $($_.Exception.Message)"
        } finally {
            $script:OperationInProgress = $false
            $btnStartTcp.Enabled = $true
            $btnStopTransfer.Enabled = $false
        }
    }.GetNewClosure())

    $btnDownloadHttp.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($txtRemoteIP.Text)) { throw "Enter the sender's IP address." }
            if ([string]::IsNullOrWhiteSpace($txtRecvFolder.Text)) { throw "Select a save folder." }

            $script:OperationInProgress = $true
            $btnDownloadHttp.Enabled = $false
            $btnStopTransfer.Enabled = $true
            $txtTransferLog.Clear()
            & $logTransfer "Downloading from HTTP server..."

            $httpRecvProgress = {
                param($phase, $current, $total, $status)
                & $logTransfer $status
            }.GetNewClosure()

            $url = "http://$($txtRemoteIP.Text):8642/"
            $result = Get-TransferFromServer -ServerUrl $url -OutputFolder $txtRecvFolder.Text -OnProgress $httpRecvProgress
            if ($result.Success) {
                & $logTransfer "Download complete! Saved to: $($result.Path)"
            } else {
                & $logTransfer "Download failed: $($result.Errors -join '; ')"
            }
        } catch {
            & $logTransfer "Error: $($_.Exception.Message)"
        } finally {
            $script:OperationInProgress = $false
            $btnDownloadHttp.Enabled = $true
            $btnStopTransfer.Enabled = $false
        }
    }.GetNewClosure())

    $btnReceiveTcp.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($txtRemoteIP.Text)) { throw "Enter the sender's IP address." }
            if ([string]::IsNullOrWhiteSpace($txtRecvFolder.Text)) { throw "Select a save folder." }

            $script:OperationInProgress = $true
            $script:CancelRequested = $false
            $btnReceiveTcp.Enabled = $false
            $btnStopTransfer.Enabled = $true
            $txtTransferLog.Clear()
            & $logTransfer "Connecting to TCP sender..."

            $tcpRecvProgress = {
                param($phase, $current, $total, $status)
                & $logTransfer $status
                if ($total -gt 0) {
                    $transferProgress.Maximum = 100
                    $transferProgress.Value = [Math]::Min([Math]::Round(($current / $total) * 100), 100)
                }
            }.GetNewClosure()

            $result = Start-DirectReceive -RemoteIP $txtRemoteIP.Text -OutputFolder $txtRecvFolder.Text -Port 8643 -OnProgress $tcpRecvProgress
            if ($result.Success) {
                & $logTransfer "Receive complete: $($result.FilesReceived) files"
            } else {
                & $logTransfer "Receive had errors: $($result.Errors -join '; ')"
            }
        } catch {
            & $logTransfer "Error: $($_.Exception.Message)"
        } finally {
            $script:OperationInProgress = $false
            $btnReceiveTcp.Enabled = $true
            $btnStopTransfer.Enabled = $false
        }
    }.GetNewClosure())

    $btnCopyShare.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($txtShareSrc.Text)) { throw "Enter source folder." }
            if ([string]::IsNullOrWhiteSpace($txtShareDest.Text)) { throw "Enter UNC destination." }

            $script:OperationInProgress = $true
            $btnCopyShare.Enabled = $false
            $txtTransferLog.Clear()
            & $logTransfer "Copying to shared folder..."

            $shareProgress = {
                param($phase, $current, $total, $status)
                & $logTransfer $status
                if ($total -gt 0) {
                    $transferProgress.Maximum = 100
                    $transferProgress.Value = [Math]::Min([Math]::Round(($current / $total) * 100), 100)
                }
            }.GetNewClosure()

            $result = Export-ToSharedFolder -SourceFolder $txtShareSrc.Text -DestinationUNC $txtShareDest.Text -OnProgress $shareProgress
            if ($result.Success) {
                & $logTransfer "Copy complete: $($result.TotalFiles) files, $(Format-FileSize $result.TotalBytes)"
            } else {
                & $logTransfer "Copy failed: $($result.Errors -join '; ')"
            }
        } catch {
            & $logTransfer "Error: $($_.Exception.Message)"
        } finally {
            $script:OperationInProgress = $false
            $btnCopyShare.Enabled = $true
        }
    }.GetNewClosure())

    $script:ContentPanel.Controls.AddRange(@($title, $subtitle, $card1, $card2, $card3, $card4, $statusGroup))
}

function Show-RestorePage {
    if ($script:OperationInProgress) { return }
    Clear-ContentPanel
    Set-ActiveSidebarButton -Page "Restore"
    Update-StatusBar "Restore Bundle"

    $accentColor = $script:Colors.AccentLightBlue

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Restore Bundle"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $accentColor
    $title.Location = New-Object System.Drawing.Point(20, 15)
    $title.AutoSize = $true

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Detect and restore a migration bundle to this PC"
    $subtitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitle.ForeColor = $script:Colors.TextSecondary
    $subtitle.Location = New-Object System.Drawing.Point(20, 50)
    $subtitle.AutoSize = $true

    # --- Browse + Detect ---
    $txtBundlePath = New-Object System.Windows.Forms.TextBox
    $txtBundlePath.Location = New-Object System.Drawing.Point(20, 80)
    $txtBundlePath.Size = New-Object System.Drawing.Size(530, 28)
    $txtBundlePath.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $txtBundlePath.BackColor = $script:Colors.InputBg
    $txtBundlePath.ForeColor = $script:Colors.TextPrimary
    $txtBundlePath.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $btnBrowseBundle = New-StyledButton -Text "Browse" -Width 80 -Height 28 -BackColor $script:Colors.ButtonBg -ForeColor $script:Colors.TextPrimary
    $btnBrowseBundle.Location = New-Object System.Drawing.Point(560, 80)
    $btnBrowseBundle.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select bundle folder"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtBundlePath.Text = $dlg.SelectedPath }
        $dlg.Dispose()
    }.GetNewClosure())

    $btnDetect = New-StyledButton -Text "Detect" -Width 80 -Height 28 -BackColor $accentColor -ForeColor $script:Colors.TextPrimary
    $btnDetect.Location = New-Object System.Drawing.Point(650, 80)

    # --- Detection Results Panel (hidden until detect) ---
    $resultsPanel = New-Object System.Windows.Forms.Panel
    $resultsPanel.Location = New-Object System.Drawing.Point(20, 120)
    $resultsPanel.Size = New-Object System.Drawing.Size(740, 300)
    $resultsPanel.BackColor = $script:Colors.ContentBg
    $resultsPanel.Visible = $false

    # Source PC info
    $lblSourceInfo = New-Object System.Windows.Forms.Label
    $lblSourceInfo.Text = ""
    $lblSourceInfo.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSourceInfo.ForeColor = $script:Colors.TextPrimary
    $lblSourceInfo.Location = New-Object System.Drawing.Point(0, 0)
    $lblSourceInfo.Size = New-Object System.Drawing.Size(740, 25)

    # Programs card
    $chkPrograms = New-Object System.Windows.Forms.CheckBox
    $chkPrograms.Text = "Programs"
    $chkPrograms.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $chkPrograms.ForeColor = $script:Colors.AccentGreen
    $chkPrograms.BackColor = $script:Colors.ContentBg
    $chkPrograms.Location = New-Object System.Drawing.Point(0, 35)
    $chkPrograms.AutoSize = $true
    $chkPrograms.Enabled = $false

    $lblProgramsInfo = New-Object System.Windows.Forms.Label
    $lblProgramsInfo.Text = ""
    $lblProgramsInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblProgramsInfo.ForeColor = $script:Colors.TextSecondary
    $lblProgramsInfo.Location = New-Object System.Drawing.Point(20, 58)
    $lblProgramsInfo.Size = New-Object System.Drawing.Size(350, 18)

    # Files card - checkboxes per folder
    $lblFilesHeader = New-Object System.Windows.Forms.Label
    $lblFilesHeader.Text = "FILES"
    $lblFilesHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblFilesHeader.ForeColor = $script:Colors.AccentOrange
    $lblFilesHeader.Location = New-Object System.Drawing.Point(0, 85)
    $lblFilesHeader.AutoSize = $true

    $listFileFolders = New-Object System.Windows.Forms.CheckedListBox
    $listFileFolders.Location = New-Object System.Drawing.Point(0, 105)
    $listFileFolders.Size = New-Object System.Drawing.Size(350, 85)
    $listFileFolders.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $listFileFolders.BackColor = $script:Colors.InputBg
    $listFileFolders.ForeColor = $script:Colors.TextPrimary
    $listFileFolders.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $listFileFolders.CheckOnClick = $true
    $listFileFolders.Enabled = $false

    # System card
    $lblSystemHeader = New-Object System.Windows.Forms.Label
    $lblSystemHeader.Text = "SYSTEM SETTINGS"
    $lblSystemHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblSystemHeader.ForeColor = $accentColor
    $lblSystemHeader.Location = New-Object System.Drawing.Point(380, 85)
    $lblSystemHeader.AutoSize = $true

    $listSystemItems = New-Object System.Windows.Forms.CheckedListBox
    $listSystemItems.Location = New-Object System.Drawing.Point(380, 105)
    $listSystemItems.Size = New-Object System.Drawing.Size(350, 85)
    $listSystemItems.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $listSystemItems.BackColor = $script:Colors.InputBg
    $listSystemItems.ForeColor = $script:Colors.TextPrimary
    $listSystemItems.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $listSystemItems.CheckOnClick = $true
    $listSystemItems.Enabled = $false

    # Bookmarks card
    $lblBookmarksHeader = New-Object System.Windows.Forms.Label
    $lblBookmarksHeader.Text = "BOOKMARKS (close browsers before restoring)"
    $lblBookmarksHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblBookmarksHeader.ForeColor = $script:Colors.AccentOrange
    $lblBookmarksHeader.Location = New-Object System.Drawing.Point(0, 200)
    $lblBookmarksHeader.AutoSize = $true

    $chkRestoreChrome = New-Object System.Windows.Forms.CheckBox
    $chkRestoreChrome.Text = "Chrome"
    $chkRestoreChrome.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkRestoreChrome.ForeColor = $script:Colors.TextPrimary
    $chkRestoreChrome.BackColor = $script:Colors.ContentBg
    $chkRestoreChrome.Location = New-Object System.Drawing.Point(0, 222)
    $chkRestoreChrome.AutoSize = $true
    $chkRestoreChrome.Enabled = $false

    $chkRestoreFirefox = New-Object System.Windows.Forms.CheckBox
    $chkRestoreFirefox.Text = "Firefox"
    $chkRestoreFirefox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkRestoreFirefox.ForeColor = $script:Colors.TextPrimary
    $chkRestoreFirefox.BackColor = $script:Colors.ContentBg
    $chkRestoreFirefox.Location = New-Object System.Drawing.Point(100, 222)
    $chkRestoreFirefox.AutoSize = $true
    $chkRestoreFirefox.Enabled = $false

    $chkRestoreEdge = New-Object System.Windows.Forms.CheckBox
    $chkRestoreEdge.Text = "Edge"
    $chkRestoreEdge.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkRestoreEdge.ForeColor = $script:Colors.TextPrimary
    $chkRestoreEdge.BackColor = $script:Colors.ContentBg
    $chkRestoreEdge.Location = New-Object System.Drawing.Point(200, 222)
    $chkRestoreEdge.AutoSize = $true
    $chkRestoreEdge.Enabled = $false

    # Restore button
    $btnRestore = New-StyledButton -Text "Restore Selected" -Width 200 -Height 40 -BackColor $accentColor -ForeColor $script:Colors.TextDark
    $btnRestore.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnRestore.Location = New-Object System.Drawing.Point(0, 255)
    $btnRestore.Enabled = $false

    $resultsPanel.Controls.AddRange(@(
        $lblSourceInfo, $chkPrograms, $lblProgramsInfo,
        $lblFilesHeader, $listFileFolders,
        $lblSystemHeader, $listSystemItems,
        $lblBookmarksHeader, $chkRestoreChrome, $chkRestoreFirefox, $chkRestoreEdge,
        $btnRestore
    ))

    # --- Progress area ---
    $restoreProgressPanel = New-Object System.Windows.Forms.Panel
    $restoreProgressPanel.Location = New-Object System.Drawing.Point(20, 430)
    $restoreProgressPanel.Size = New-Object System.Drawing.Size(740, 80)
    $restoreProgressPanel.BackColor = $script:Colors.ContentBg
    $restoreProgressPanel.Visible = $false

    $lblPhase = New-Object System.Windows.Forms.Label
    $lblPhase.Text = ""
    $lblPhase.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblPhase.ForeColor = $accentColor
    $lblPhase.Location = New-Object System.Drawing.Point(0, 0)
    $lblPhase.Size = New-Object System.Drawing.Size(740, 20)

    $restoreBar = New-Object ModernProgressBar
    $restoreBar.Location = New-Object System.Drawing.Point(0, 25)
    $restoreBar.Size = New-Object System.Drawing.Size(740, 14)
    $restoreBar.BarColor = $script:Colors.AccentLightBlue
    $restoreBar.TrackColor = $script:Colors.CardBg

    $lblRestoreStatus = New-Object System.Windows.Forms.Label
    $lblRestoreStatus.Text = ""
    $lblRestoreStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblRestoreStatus.ForeColor = $script:Colors.TextSecondary
    $lblRestoreStatus.Location = New-Object System.Drawing.Point(0, 50)
    $lblRestoreStatus.Size = New-Object System.Drawing.Size(740, 18)

    $restoreProgressPanel.Controls.AddRange(@($lblPhase, $restoreBar, $lblRestoreStatus))

    # Store bundle contents for restore
    $script:RestoreBundleContents = $null

    # --- Detect handler ---
    $btnDetect.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($txtBundlePath.Text)) { throw "Select a bundle path." }
            if (-not (Test-Path $txtBundlePath.Text)) { throw "Path not found." }

            $btnDetect.Enabled = $false
            $btnDetect.Text = "..."
            Update-GuiStatus

            $script:RestoreBundleContents = Find-BundleContents -Path $txtBundlePath.Text

            $bc = $script:RestoreBundleContents
            $resultsPanel.Visible = $true
            $restoreProgressPanel.Visible = $false

            # Source info
            $infoText = ""
            if ($bc.Meta.SourcePC) { $infoText += "Source PC: $($bc.Meta.SourcePC)" }
            if ($bc.Meta.Date) {
                try { $d = [DateTime]::Parse($bc.Meta.Date); $infoText += " | Date: $($d.ToString('yyyy-MM-dd HH:mm'))" } catch { $infoText += " | Date: $($bc.Meta.Date)" }
            }
            $lblSourceInfo.Text = $infoText

            # Programs
            if ($bc.Programs.Found) {
                $chkPrograms.Enabled = $true
                $chkPrograms.Checked = $true
                $chkPrograms.Text = "Programs ($($bc.Programs.Count) found)"
                $lblProgramsInfo.Text = "Will install via winget/chocolatey"
            } else {
                $chkPrograms.Enabled = $false
                $chkPrograms.Checked = $false
                $chkPrograms.Text = "Programs (not found)"
                $lblProgramsInfo.Text = ""
            }

            # Files
            $listFileFolders.Items.Clear()
            if ($bc.Files.Found -and $bc.Files.Folders.Count -gt 0) {
                $listFileFolders.Enabled = $true
                foreach ($f in $bc.Files.Folders) {
                    $sizeStr = Format-FileSize $f.SizeBytes
                    [void]$listFileFolders.Items.Add("$($f.Name) ($sizeStr)", $true)
                }
            } else {
                $listFileFolders.Enabled = $false
            }

            # System
            $listSystemItems.Items.Clear()
            if ($bc.System.Found -and $bc.System.Items.Count -gt 0) {
                $listSystemItems.Enabled = $true
                foreach ($item in $bc.System.Items) {
                    [void]$listSystemItems.Items.Add($item, $true)
                }
            } else {
                $listSystemItems.Enabled = $false
            }

            # Bookmarks
            $chkRestoreChrome.Enabled = $false; $chkRestoreChrome.Checked = $false
            $chkRestoreFirefox.Enabled = $false; $chkRestoreFirefox.Checked = $false
            $chkRestoreEdge.Enabled = $false; $chkRestoreEdge.Checked = $false
            if ($bc.Bookmarks.Found) {
                foreach ($b in $bc.Bookmarks.Browsers) {
                    switch ($b) {
                        'Chrome' { $chkRestoreChrome.Enabled = $true; $chkRestoreChrome.Checked = $true }
                        'Firefox' { $chkRestoreFirefox.Enabled = $true; $chkRestoreFirefox.Checked = $true }
                        'Edge' { $chkRestoreEdge.Enabled = $true; $chkRestoreEdge.Checked = $true }
                    }
                }
            }

            # Enable restore button if anything found
            $hasAnything = $bc.Programs.Found -or $bc.Files.Found -or $bc.System.Found -or $bc.Bookmarks.Found
            $btnRestore.Enabled = $hasAnything

            Write-Log "Bundle detected: Programs=$($bc.Programs.Found) Files=$($bc.Files.Found) System=$($bc.System.Found) Bookmarks=$($bc.Bookmarks.Found)"
            Update-StatusBar "Bundle detected - ready to restore"

        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $btnDetect.Enabled = $true
            $btnDetect.Text = "Detect"
        }
    }.GetNewClosure())

    # --- Restore handler ---
    $btnRestore.Add_Click({
        try {
            $bc = $script:RestoreBundleContents
            if (-not $bc) { throw "Run Detect first." }

            # Build selections
            $selections = @{}

            if ($chkPrograms.Checked -and $chkPrograms.Enabled) {
                $selections['Programs'] = $true
            }

            $selectedFolders = @()
            if ($bc.Files.Found) {
                for ($i = 0; $i -lt $listFileFolders.Items.Count; $i++) {
                    if ($listFileFolders.GetItemChecked($i)) {
                        $folderName = $bc.Files.Folders[$i].Name
                        $selectedFolders += $folderName
                    }
                }
            }
            if ($selectedFolders.Count -gt 0) { $selections['Files'] = $selectedFolders }

            $selectedSystem = @()
            for ($i = 0; $i -lt $listSystemItems.Items.Count; $i++) {
                if ($listSystemItems.GetItemChecked($i)) {
                    $selectedSystem += [string]$listSystemItems.Items[$i]
                }
            }
            if ($selectedSystem.Count -gt 0) { $selections['System'] = $selectedSystem }

            $selectedBrowsers = @()
            if ($chkRestoreChrome.Checked -and $chkRestoreChrome.Enabled) { $selectedBrowsers += 'Chrome' }
            if ($chkRestoreFirefox.Checked -and $chkRestoreFirefox.Enabled) { $selectedBrowsers += 'Firefox' }
            if ($chkRestoreEdge.Checked -and $chkRestoreEdge.Enabled) { $selectedBrowsers += 'Edge' }
            if ($selectedBrowsers.Count -gt 0) { $selections['Bookmarks'] = $selectedBrowsers }

            if ($selections.Count -eq 0) { throw "Nothing selected to restore." }

            # Confirmation
            $msg = "Restore the following?`n"
            if ($selections.ContainsKey('Programs')) { $msg += "`n- Programs ($($bc.Programs.Count) apps)" }
            if ($selections.ContainsKey('Files')) { $msg += "`n- Files ($($selectedFolders.Count) folders)" }
            if ($selections.ContainsKey('System')) { $msg += "`n- System ($($selectedSystem.Count) items)" }
            if ($selections.ContainsKey('Bookmarks')) { $msg += "`n- Bookmarks ($($selectedBrowsers -join ', '))" }

            $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm Restore", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

            $script:OperationInProgress = $true
            $script:CancelRequested = $false
            $btnRestore.Enabled = $false
            $btnRestore.Text = "Restoring..."
            $restoreProgressPanel.Visible = $true

            $phaseCallback = {
                param($phase, $message)
                $lblPhase.Text = "$phase - $message"
                Update-GuiStatus
            }.GetNewClosure()

            $progressCallback = {
                param($name, $current, $total, $status)
                $lblRestoreStatus.Text = "$name - $status"
                Update-GuiStatus
            }.GetNewClosure()

            $result = Start-FullRestore -BundlePath $txtBundlePath.Text -BundleContents $bc -Selections $selections -OnPhaseChange $phaseCallback -OnProgress $progressCallback

            # Show results
            $restoreBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
            $restoreBar.Value = 100

            $summary = "Restore complete!"
            if ($result.Errors.Count -gt 0) {
                $summary += " ($($result.Errors.Count) error(s))"
                $lblPhase.ForeColor = $script:Colors.AccentOrange
            } else {
                $lblPhase.ForeColor = $script:Colors.AccentGreen
            }
            $lblPhase.Text = $summary

            $details = @()
            if ($result.Files) { $details += "Files: $($result.Files.FoldersRestored) folders restored" }
            if ($result.Programs) { $details += "Programs: $(@($result.Programs.Results | Where-Object { $_.Status -eq 'Installed' }).Count) installed" }
            $lblRestoreStatus.Text = $details -join " | "

            Update-StatusBar $summary
            Write-Log $summary -Level 'SUCCESS'

        } catch {
            Write-Log "Restore error: $($_.Exception.Message)" -Level 'ERROR'
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $script:OperationInProgress = $false
            $btnRestore.Enabled = $true
            $btnRestore.Text = "Restore Selected"
        }
    }.GetNewClosure())

    $script:ContentPanel.Controls.AddRange(@($title, $subtitle, $txtBundlePath, $btnBrowseBundle, $btnDetect, $resultsPanel, $restoreProgressPanel))
}

function Show-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($script:AppName) v$($script:AppVersion)"
    $form.Size = New-Object System.Drawing.Size(1024, 700)
    $form.BackColor = $script:Colors.WindowBg
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false

    # Enable double buffering to eliminate flicker
    $form.DoubleBuffered = $true
    try {
        $form.GetType().GetMethod('SetStyle', [System.Reflection.BindingFlags]'NonPublic,Instance').Invoke($form, @(
            [System.Windows.Forms.ControlStyles]::OptimizedDoubleBuffer -bor
            [System.Windows.Forms.ControlStyles]::AllPaintingInWmPaint -bor
            [System.Windows.Forms.ControlStyles]::UserPaint, $true))
    } catch { }

    # Restore window position from settings, or center if first run
    if ($script:UserSettings.WindowX -ge 0 -and $script:UserSettings.WindowY -ge 0) {
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
        $form.Location = New-Object System.Drawing.Point($script:UserSettings.WindowX, $script:UserSettings.WindowY)
    } else {
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    }

    # Save window position on close
    $form.Add_FormClosing({
        if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
            Update-Setting -Key 'WindowX' -Value $form.Location.X
            Update-Setting -Key 'WindowY' -Value $form.Location.Y
        }
    })

    # Enable dark title bar on Windows 10/11
    $form.Add_HandleCreated({
        try { [DwmHelper]::EnableDarkTitleBar($form.Handle) } catch { }
    })

    # Set app icon (Windows built-in)
    $appIcon = Get-AppIcon
    if ($appIcon) {
        $form.Icon = $appIcon
    }

    $sidebar = New-Object System.Windows.Forms.Panel
    $sidebar.Location = New-Object System.Drawing.Point(0, 0)
    $sidebar.Size = New-Object System.Drawing.Size(200, 700)
    $sidebar.BackColor = $script:Colors.SidebarBg

    # Sidebar gradient background
    $sidebar.Add_Paint({
        param($s, $e)
        $g = $e.Graphics
        $r = $s.ClientRectangle
        try {
            $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                $r,
                [System.Drawing.Color]::FromArgb(16, 16, 22),
                [System.Drawing.Color]::FromArgb(12, 12, 16),
                [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
            $g.FillRectangle($brush, $r)
            $brush.Dispose()
        } catch { }
    })

    $lblAppName = New-Object System.Windows.Forms.Label
    $lblAppName.Text = "LazyTransfer"
    $lblAppName.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblAppName.ForeColor = $script:Colors.TextPrimary
    $lblAppName.Location = New-Object System.Drawing.Point(15, 20)
    $lblAppName.AutoSize = $true
    $lblAppName.BackColor = [System.Drawing.Color]::Transparent
    $sidebar.Controls.Add($lblAppName)

    # Separator line between title and nav
    $separator = New-Object System.Windows.Forms.Panel
    $separator.Location = New-Object System.Drawing.Point(15, 55)
    $separator.Size = New-Object System.Drawing.Size(170, 1)
    $separator.BackColor = $script:Colors.BorderColor
    $sidebar.Controls.Add($separator)

    $btnScan = New-SidebarButton -Text "Scan" -Page "Scan" -AccentColor $script:Colors.AccentBlue
    $btnScan.Location = New-Object System.Drawing.Point(0, 70)
    $btnScan.Add_Click({ Show-ScanPage })
    $script:SidebarButtons["Scan"] = $btnScan

    $btnInstall = New-SidebarButton -Text "Install" -Page "Install" -AccentColor $script:Colors.AccentGreen
    $btnInstall.Location = New-Object System.Drawing.Point(0, 118)
    $btnInstall.Add_Click({ Show-InstallPage })
    $script:SidebarButtons["Install"] = $btnInstall

    $btnFiles = New-SidebarButton -Text "Files" -Page "Files" -AccentColor $script:Colors.AccentOrange
    $btnFiles.Location = New-Object System.Drawing.Point(0, 166)
    $btnFiles.Add_Click({ Show-FilesPage })
    $script:SidebarButtons["Files"] = $btnFiles

    $btnTransfer = New-SidebarButton -Text "Transfer" -Page "Transfer" -AccentColor $script:Colors.AccentPink
    $btnTransfer.Location = New-Object System.Drawing.Point(0, 214)
    $btnTransfer.Add_Click({ Show-TransferPage })
    $script:SidebarButtons["Transfer"] = $btnTransfer

    $btnRestore = New-SidebarButton -Text "Restore" -Page "Restore" -AccentColor $script:Colors.AccentLightBlue
    $btnRestore.Location = New-Object System.Drawing.Point(0, 262)
    $btnRestore.Add_Click({ Show-RestorePage })
    $script:SidebarButtons["Restore"] = $btnRestore

    $sidebar.Controls.AddRange(@($btnScan, $btnInstall, $btnFiles, $btnTransfer, $btnRestore))

    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.Text = "v$($script:AppVersion)"
    $lblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblVersion.ForeColor = $script:Colors.TextSecondary
    $lblVersion.Location = New-Object System.Drawing.Point(15, 620)
    $lblVersion.AutoSize = $true
    $lblVersion.BackColor = [System.Drawing.Color]::Transparent
    $sidebar.Controls.Add($lblVersion)

    $script:ContentPanel = New-Object System.Windows.Forms.Panel
    $script:ContentPanel.Location = New-Object System.Drawing.Point(200, 0)
    $script:ContentPanel.Size = New-Object System.Drawing.Size(808, 550)
    $script:ContentPanel.BackColor = $script:Colors.ContentBg
    $script:ContentPanel.AutoScroll = $true

    $logPanel = New-Object System.Windows.Forms.Panel
    $logPanel.Location = New-Object System.Drawing.Point(200, 550)
    $logPanel.Size = New-Object System.Drawing.Size(808, 80)
    $logPanel.BackColor = $script:Colors.SidebarBg

    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = "Log:"
    $lblLog.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblLog.ForeColor = $script:Colors.TextSecondary
    $lblLog.Location = New-Object System.Drawing.Point(10, 5)
    $lblLog.AutoSize = $true
    $logPanel.Controls.Add($lblLog)

    $script:LogTextBox = New-Object System.Windows.Forms.TextBox
    $script:LogTextBox.Location = New-Object System.Drawing.Point(10, 25)
    $script:LogTextBox.Size = New-Object System.Drawing.Size(785, 45)
    $script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 8)
    $script:LogTextBox.BackColor = $script:Colors.InputBg
    $script:LogTextBox.ForeColor = $script:Colors.TextSecondary
    $script:LogTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $script:LogTextBox.Multiline = $true
    $script:LogTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $script:LogTextBox.ReadOnly = $true
    $script:LogTextBox.MaxLength = 0  # Unlimited (we trim in Write-Log)
    $logPanel.Controls.Add($script:LogTextBox)

    $statusBar = New-Object System.Windows.Forms.Panel
    $statusBar.Location = New-Object System.Drawing.Point(200, 630)
    $statusBar.Size = New-Object System.Drawing.Size(808, 30)
    $statusBar.BackColor = $script:Colors.WindowBg

    $script:StatusLabel = New-Object System.Windows.Forms.Label
    $script:StatusLabel.Text = "Ready"
    $script:StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:StatusLabel.ForeColor = $script:Colors.TextSecondary
    $script:StatusLabel.Location = New-Object System.Drawing.Point(10, 5)
    $script:StatusLabel.AutoSize = $true
    $statusBar.Controls.Add($script:StatusLabel)

    $form.Controls.AddRange(@($sidebar, $script:ContentPanel, $logPanel, $statusBar))

    Show-ScanPage
    Write-Log "$($script:AppName) started"

    [void]$form.ShowDialog()
}
#endregion GUI

try {
    Show-MainForm
} catch {
    [System.Windows.Forms.MessageBox]::Show("Fatal error: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
