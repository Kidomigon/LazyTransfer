<#
.SYNOPSIS
    Edge-case test suite for LazyTransfer v2.5

.DESCRIPTION
    Exercises edge cases across Format-FileSize, Find-BundleContents,
    Get-AppCategory, NetworkEngine helpers, CLI region extraction, and
    New-TransferPin.  Uses plain if/throw assertions so Pester is not required.

.NOTES
    Run from any PowerShell 5.1+ session.  Does NOT require admin privileges.
    Creates temporary files/folders under $env:TEMP and cleans up on exit.
#>

#region Bootstrap — load project functions without launching the GUI
$ErrorActionPreference = 'Stop'

$script:TestDir    = Split-Path -Parent $MyInvocation.MyCommand.Definition
$script:ProjectDir = Split-Path -Parent $script:TestDir

# ---- Minimal stubs so sourced code never touches a real GUI ----
$script:LogTextBox        = $null
$script:StatusLabel       = $null
$script:ContentPanel      = $null
$script:SidebarButtons    = @{}
$script:OperationInProgress = $false
$script:CancelRequested   = $false
$script:LastMigrationPath = $null
$script:CLIMode           = $true
$script:LogPath           = $null
$script:LogMessages       = New-Object System.Collections.Generic.List[string]
$script:SettingsFile      = Join-Path $env:TEMP "lazytransfer-test-settings.json"
$script:AppName           = "LazyTransfer"
$script:AppVersion        = "2.5"
$script:CurrentPage       = "Test"
$script:InstalledProgramsCache = $null

# Stub for Update-GuiStatus (GUI no-op)
function Update-GuiStatus { }

# Stub for Write-Log so sourced modules can call it
function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    # Silent during tests — uncomment the line below for debug output
    # Write-Host "[$Level] $Message"
}

function Update-StatusBar { param([string]$Message) }

# ---- Source the engine regions from the GUI file (same technique the CLI uses) ----
$guiScript = Join-Path $script:ProjectDir "LazyTransfer-GUI.ps1"
if (-not (Test-Path $guiScript)) {
    throw "Cannot find LazyTransfer-GUI.ps1 at $guiScript"
}

$guiContent = Get-Content $guiScript -Raw -Encoding UTF8

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
    $endTag   = "#endregion $regionName"
    $startIdx = $guiContent.IndexOf($startTag)
    $endIdx   = $guiContent.IndexOf($endTag)
    if ($startIdx -ge 0 -and $endIdx -gt $startIdx) {
        $regionCode = $guiContent.Substring($startIdx, $endIdx + $endTag.Length - $startIdx)
        try {
            . ([ScriptBlock]::Create($regionCode))
        } catch {
            Write-Host "Warning: Could not load region '$regionName': $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# ---- Source the NetworkEngine module ----
$networkEngine = Join-Path $script:ProjectDir "modules" "NetworkEngine.ps1"
if (-not (Test-Path $networkEngine)) {
    throw "Cannot find NetworkEngine.ps1 at $networkEngine"
}
. $networkEngine
#endregion Bootstrap

# ====================================================================
#  Test harness
# ====================================================================
$script:PassCount = 0
$script:FailCount = 0
$script:TestErrors = @()

function Assert-True {
    param(
        [Parameter(Mandatory=$true)][bool]$Condition,
        [Parameter(Mandatory=$true)][string]$TestName
    )
    if ($Condition) {
        Write-Host "  [PASS] $TestName" -ForegroundColor Green
        $script:PassCount++
    } else {
        Write-Host "  [FAIL] $TestName" -ForegroundColor Red
        $script:FailCount++
        $script:TestErrors += $TestName
    }
}

function Assert-Equal {
    param(
        [Parameter(Mandatory=$true)]$Expected,
        [Parameter(Mandatory=$true)]$Actual,
        [Parameter(Mandatory=$true)][string]$TestName
    )
    if ($Expected -eq $Actual) {
        Write-Host "  [PASS] $TestName" -ForegroundColor Green
        $script:PassCount++
    } else {
        Write-Host "  [FAIL] $TestName  (expected='$Expected'  actual='$Actual')" -ForegroundColor Red
        $script:FailCount++
        $script:TestErrors += $TestName
    }
}

function Assert-Throws {
    param(
        [Parameter(Mandatory=$true)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory=$true)][string]$TestName
    )
    $threw = $false
    try {
        & $ScriptBlock
    } catch {
        $threw = $true
    }
    if ($threw) {
        Write-Host "  [PASS] $TestName" -ForegroundColor Green
        $script:PassCount++
    } else {
        Write-Host "  [FAIL] $TestName  (expected an exception but none was thrown)" -ForegroundColor Red
        $script:FailCount++
        $script:TestErrors += $TestName
    }
}

function Assert-NoThrow {
    param(
        [Parameter(Mandatory=$true)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory=$true)][string]$TestName
    )
    $threw = $false
    $errMsg = ""
    try {
        & $ScriptBlock
    } catch {
        $threw = $true
        $errMsg = $_.Exception.Message
    }
    if (-not $threw) {
        Write-Host "  [PASS] $TestName" -ForegroundColor Green
        $script:PassCount++
    } else {
        Write-Host "  [FAIL] $TestName  (unexpected exception: $errMsg)" -ForegroundColor Red
        $script:FailCount++
        $script:TestErrors += $TestName
    }
}

function Assert-Match {
    param(
        [Parameter(Mandatory=$true)][string]$Pattern,
        [Parameter(Mandatory=$true)][string]$Actual,
        [Parameter(Mandatory=$true)][string]$TestName
    )
    if ($Actual -match $Pattern) {
        Write-Host "  [PASS] $TestName" -ForegroundColor Green
        $script:PassCount++
    } else {
        Write-Host "  [FAIL] $TestName  (value='$Actual' did not match pattern='$Pattern')" -ForegroundColor Red
        $script:FailCount++
        $script:TestErrors += $TestName
    }
}

# ====================================================================
#  Temporary folder management
# ====================================================================
$script:TempRoot = Join-Path $env:TEMP "LazyTransfer-EdgeCaseTests-$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -Path $script:TempRoot -ItemType Directory -Force | Out-Null

function New-TestFolder {
    param([string]$Name)
    $path = Join-Path $script:TempRoot $Name
    New-Item -Path $path -ItemType Directory -Force | Out-Null
    return $path
}

# ====================================================================
#  1. Format-FileSize edge cases
# ====================================================================
Write-Host "`n=== 1. Format-FileSize Edge Cases ===" -ForegroundColor Cyan

Assert-Equal "0 B" (Format-FileSize -Bytes 0) "Format-FileSize: 0 bytes returns '0 B'"

# Negative values — the function parameter is [long], negatives are valid longs.
# The function has no guard; it should fall through to the else branch.
$negResult = Format-FileSize -Bytes (-100)
Assert-Equal "-100 B" $negResult "Format-FileSize: negative value (-100) returns '-100 B'"

# Exactly 1 KB boundary
Assert-Equal "1 KB" (Format-FileSize -Bytes 1024) "Format-FileSize: exactly 1024 bytes returns '1 KB'"

# One byte below 1 KB
Assert-Equal "1023 B" (Format-FileSize -Bytes 1023) "Format-FileSize: 1023 bytes returns '1023 B'"

# Exactly 1 MB
Assert-Equal "1 MB" (Format-FileSize -Bytes 1048576) "Format-FileSize: exactly 1 MB returns '1 MB'"

# Exactly 1 GB
Assert-Equal "1 GB" (Format-FileSize -Bytes 1073741824) "Format-FileSize: exactly 1 GB returns '1 GB'"

# Exactly 1 TB
$oneTB = [long]1024 * 1024 * 1024 * 1024
Assert-Equal "1 TB" (Format-FileSize -Bytes $oneTB) "Format-FileSize: exactly 1 TB returns '1 TB'"

# Very large value (5 TB)
$fiveTB = [long]5 * 1024 * 1024 * 1024 * 1024
Assert-Equal "5 TB" (Format-FileSize -Bytes $fiveTB) "Format-FileSize: 5 TB returns '5 TB'"

# $null input — [long] parameter will coerce $null to 0
Assert-Equal "0 B" (Format-FileSize -Bytes $null) "Format-FileSize: `$null coerces to 0 and returns '0 B'"

# Fractional KB display (1536 bytes = 1.5 KB)
Assert-Equal "1.5 KB" (Format-FileSize -Bytes 1536) "Format-FileSize: 1536 bytes returns '1.5 KB'"

# ====================================================================
#  2. Find-BundleContents edge cases
# ====================================================================
Write-Host "`n=== 2. Find-BundleContents Edge Cases ===" -ForegroundColor Cyan

# 2a. Empty folder
$emptyFolder = New-TestFolder "empty-bundle"
$bc = Find-BundleContents -Path $emptyFolder
Assert-True ($bc.Programs.Found -eq $false) "Find-BundleContents: empty folder — Programs.Found is false"
Assert-True ($bc.Files.Found -eq $false)    "Find-BundleContents: empty folder — Files.Found is false"
Assert-True ($bc.System.Found -eq $false)   "Find-BundleContents: empty folder — System.Found is false"
Assert-True ($bc.Bookmarks.Found -eq $false) "Find-BundleContents: empty folder — Bookmarks.Found is false"
Assert-True ($bc.Meta.SourcePC -eq "")       "Find-BundleContents: empty folder — Meta.SourcePC is empty"

# 2b. Folder that does not exist
$nonExistent = Join-Path $script:TempRoot "this-folder-does-not-exist"
$threw = $false
try {
    $bc2 = Find-BundleContents -Path $nonExistent
    # The function uses Get-ChildItem with -ErrorAction SilentlyContinue internally,
    # and Test-Path returns $false.  It should return an empty bundle, not throw.
    $threw = $false
} catch {
    $threw = $true
}
# Whether it throws or returns empty, we document the behaviour
if ($threw) {
    Write-Host "  [PASS] Find-BundleContents: non-existent folder throws (acceptable)" -ForegroundColor Green
    $script:PassCount++
} else {
    Assert-True ($bc2.Programs.Found -eq $false) "Find-BundleContents: non-existent folder — Programs.Found is false"
}

# 2c. Corrupted / empty programs.json
$corruptedFolder = New-TestFolder "corrupted-bundle"
Set-Content -Path (Join-Path $corruptedFolder "programs.json") -Value "THIS IS NOT JSON" -Encoding UTF8
$bc3 = Find-BundleContents -Path $corruptedFolder
Assert-True ($bc3.Programs.Found -eq $false) "Find-BundleContents: corrupted programs.json — Programs.Found is false"

# Empty programs.json file
$emptyJsonFolder = New-TestFolder "emptyjson-bundle"
Set-Content -Path (Join-Path $emptyJsonFolder "programs.json") -Value "" -Encoding UTF8
$bc3b = Find-BundleContents -Path $emptyJsonFolder
Assert-True ($bc3b.Programs.Found -eq $false) "Find-BundleContents: empty programs.json — Programs.Found is false"

# 2d. Folder with valid manifest but no actual user-folders on disk
$manifestOnlyFolder = New-TestFolder "manifest-only"
$manifestJson = @{
    SourceComputer = "TESTPC"
    MigrationDate  = "2025-01-01T00:00:00"
    ToolVersion    = "2.5"
    Folders = @(
        @{ Name = "Documents"; SizeBytes = 1000 }
        @{ Name = "Pictures";  SizeBytes = 2000 }
    )
} | ConvertTo-Json -Depth 5
Set-Content -Path (Join-Path $manifestOnlyFolder "migration-manifest.json") -Value $manifestJson -Encoding UTF8
$bc4 = Find-BundleContents -Path $manifestOnlyFolder
Assert-True ($bc4.Files.Found -eq $true) "Find-BundleContents: manifest present — Files.Found is true"
Assert-True ($bc4.Files.Folders.Count -eq 0) "Find-BundleContents: manifest but no folders on disk — Folders count is 0"
Assert-Equal "TESTPC" $bc4.Meta.SourcePC "Find-BundleContents: manifest SourcePC read correctly"

# 2e. Path with spaces and unicode characters
$unicodeFolder = New-TestFolder "bundle with spaces and unicode"
$progJson = @{
    Programs = @(
        @{ DisplayName = "TestApp"; DisplayVersion = "1.0" }
    )
    Meta = @{
        ComputerName = "PC"
        ScanTime     = "2025-06-01"
        ToolVersion  = "2.5"
    }
} | ConvertTo-Json -Depth 5
Set-Content -Path (Join-Path $unicodeFolder "programs.json") -Value $progJson -Encoding UTF8
$bc5 = Find-BundleContents -Path $unicodeFolder
Assert-True ($bc5.Programs.Found -eq $true) "Find-BundleContents: path with spaces — Programs.Found is true"
Assert-Equal 1 $bc5.Programs.Count "Find-BundleContents: path with spaces — Count is 1"

# 2f. Valid bundle with bookmarks
$fullBundle = New-TestFolder "full-bundle"
$bookmarksDir = Join-Path $fullBundle "Bookmarks"
New-Item -Path $bookmarksDir -ItemType Directory -Force | Out-Null
Set-Content -Path (Join-Path $bookmarksDir "Chrome_Bookmarks.json") -Value '{"roots":{}}' -Encoding UTF8
Set-Content -Path (Join-Path $bookmarksDir "Edge_Bookmarks.json") -Value '{"roots":{}}' -Encoding UTF8
New-Item -Path (Join-Path $bookmarksDir "Firefox") -ItemType Directory -Force | Out-Null
$bc6 = Find-BundleContents -Path $fullBundle
Assert-True ($bc6.Bookmarks.Found -eq $true) "Find-BundleContents: Bookmarks directory detected"
Assert-True ($bc6.Bookmarks.Browsers.Count -eq 3) "Find-BundleContents: all three browsers detected (Chrome, Edge, Firefox)"

# 2g. Bundle with SystemMigration folder
$sysBundle = New-TestFolder "sys-bundle"
$sysMigDir = Join-Path $sysBundle "SystemMigration"
New-Item -Path $sysMigDir -ItemType Directory -Force | Out-Null
New-Item -Path (Join-Path $sysMigDir "wifi-profiles") -ItemType Directory -Force | Out-Null
New-Item -Path (Join-Path $sysMigDir "ssh") -ItemType Directory -Force | Out-Null
Set-Content -Path (Join-Path $sysMigDir "env-vars.json") -Value '{}' -Encoding UTF8
$bc7 = Find-BundleContents -Path $sysBundle
Assert-True ($bc7.System.Found -eq $true) "Find-BundleContents: SystemMigration detected"
Assert-True ($bc7.System.Items -contains "WiFi") "Find-BundleContents: WiFi detected in system items"
Assert-True ($bc7.System.Items -contains "SSH") "Find-BundleContents: SSH detected in system items"
Assert-True ($bc7.System.Items -contains "Environment Variables") "Find-BundleContents: Env Vars detected in system items"

# 2h. Bundle with actual files migration subfolder containing real files
$filesBundle = New-TestFolder "files-bundle"
$filesMigDir = Join-Path $filesBundle "LazyTransfer-Files-TestPC"
New-Item -Path $filesMigDir -ItemType Directory -Force | Out-Null
$docsDir = Join-Path $filesMigDir "Documents"
New-Item -Path $docsDir -ItemType Directory -Force | Out-Null
Set-Content -Path (Join-Path $docsDir "readme.txt") -Value "hello world" -Encoding UTF8
$bc8 = Find-BundleContents -Path $filesBundle
Assert-True ($bc8.Files.Found -eq $true) "Find-BundleContents: LazyTransfer-Files-* subfolder detected"
Assert-True ($bc8.Files.Folders.Count -ge 1) "Find-BundleContents: at least one folder found inside files migration"

# ====================================================================
#  3. Get-AppCategory edge cases
# ====================================================================
Write-Host "`n=== 3. Get-AppCategory Edge Cases ===" -ForegroundColor Cyan

# Empty string — mandatory [string] parameter; an empty string is still a valid string
# The regex patterns should not match, so it returns 'Other'
Assert-Equal "Other" (Get-AppCategory -DisplayName " ") "Get-AppCategory: whitespace-only name returns 'Other'"

# Known category match
Assert-Equal "Browsers" (Get-AppCategory -DisplayName "Google Chrome") "Get-AppCategory: 'Google Chrome' returns 'Browsers'"
Assert-Equal "DevTools" (Get-AppCategory -DisplayName "Visual Studio Code") "Get-AppCategory: 'Visual Studio Code' returns 'DevTools'"
Assert-Equal "Media" (Get-AppCategory -DisplayName "VLC media player") "Get-AppCategory: 'VLC media player' returns 'Media'"
Assert-Equal "Gaming" (Get-AppCategory -DisplayName "Steam") "Get-AppCategory: 'Steam' returns 'Gaming'"

# App name with special characters
Assert-Equal "Other" (Get-AppCategory -DisplayName "My@#%App!") "Get-AppCategory: special characters returns 'Other'"

# Very long app name (no category match)
$longName = "A" * 5000
Assert-Equal "Other" (Get-AppCategory -DisplayName $longName) "Get-AppCategory: very long name (5000 chars) returns 'Other'"

# Very long app name that contains a keyword buried inside
$longWithKeyword = ("X" * 2000) + " Chrome " + ("X" * 2000)
Assert-Equal "Browsers" (Get-AppCategory -DisplayName $longWithKeyword) "Get-AppCategory: long name with embedded 'Chrome' returns 'Browsers'"

# $null — [string] mandatory parameter would normally throw a binding error.
# We test this by wrapping in a try/catch.
Assert-Throws { Get-AppCategory -DisplayName $null } "Get-AppCategory: `$null name throws parameter binding error"

# Case insensitivity — "FIREFOX" should still match
Assert-Equal "Browsers" (Get-AppCategory -DisplayName "FIREFOX") "Get-AppCategory: 'FIREFOX' (all caps) returns 'Browsers'"

# App name that looks like multiple categories — first match wins
$multiMatch = "Chrome Steam Visual Studio"
$multiResult = Get-AppCategory -DisplayName $multiMatch
Assert-True ($multiResult -ne "Other") "Get-AppCategory: multi-category name does not return 'Other'"

# ====================================================================
#  4. Network helper edge cases
# ====================================================================
Write-Host "`n=== 4. Network Helper Edge Cases ===" -ForegroundColor Cyan

# 4a. Test-PortAvailable with port 0 — OS may or may not allow this
# On most systems, binding to port 0 asks the OS to assign an ephemeral port, so it should succeed
$port0Result = Test-PortAvailable -Port 0
Assert-True ($port0Result -is [bool]) "Test-PortAvailable: port 0 returns a boolean"

# 4b. Test-PortAvailable with port 99999 — out of valid TCP range, TcpListener will throw
$port99999Result = Test-PortAvailable -Port 99999
Assert-True ($port99999Result -eq $false) "Test-PortAvailable: port 99999 returns false (out of range)"

# 4c. Test-PortAvailable with a typical available high port
# We try 58921 — unlikely to be in use
$highPortResult = Test-PortAvailable -Port 58921
Assert-True ($highPortResult -is [bool]) "Test-PortAvailable: port 58921 returns a boolean"

# 4d. Get-LocalIPAddresses returns an array
$ips = Get-LocalIPAddresses
Assert-True ($ips -is [array] -or $ips -is [System.Collections.IEnumerable] -or $null -eq $ips) "Get-LocalIPAddresses: returns array or null (no crash)"

# If we got results, verify structure
if ($ips -and $ips.Count -gt 0) {
    $firstIP = $ips[0]
    Assert-True ($null -ne $firstIP.IP) "Get-LocalIPAddresses: first result has IP property"
    Assert-True ($null -ne $firstIP.Interface) "Get-LocalIPAddresses: first result has Interface property"
    Assert-True ($null -ne $firstIP.Type) "Get-LocalIPAddresses: first result has Type property"
    Assert-Match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' $firstIP.IP "Get-LocalIPAddresses: IP looks like an IPv4 address"
} else {
    Write-Host "  [PASS] Get-LocalIPAddresses: returned empty (no active interfaces) — acceptable" -ForegroundColor Green
    $script:PassCount++
}

# 4e. Add-FirewallRule with invalid name characters
Assert-Throws { Add-FirewallRule -Name "bad name with spaces!" -Port 8080 } "Add-FirewallRule: name with spaces/special chars throws"
Assert-Throws { Add-FirewallRule -Name "rule;injection" -Port 8080 } "Add-FirewallRule: name with semicolon throws"
Assert-Throws { Add-FirewallRule -Name "" -Port 8080 } "Add-FirewallRule: empty name throws"

# Valid name format should NOT throw on validation (the actual netsh call may fail
# without admin, but the validation check itself should pass)
$fwValidationPassed = $false
try {
    Add-FirewallRule -Name "Valid-Rule-Name-123" -Port 8080
    $fwValidationPassed = $true
} catch {
    # If it threw because of netsh (not validation), that is acceptable
    if ($_.Exception.Message -like "*Invalid firewall*") {
        $fwValidationPassed = $false
    } else {
        $fwValidationPassed = $true  # validation passed, netsh failed (no admin)
    }
}
Assert-True $fwValidationPassed "Add-FirewallRule: valid name 'Valid-Rule-Name-123' passes name validation"

# 4f. Port validation boundaries for Add-FirewallRule
Assert-Throws { Add-FirewallRule -Name "TestRule" -Port 1023 } "Add-FirewallRule: port 1023 throws (below 1024)"
Assert-Throws { Add-FirewallRule -Name "TestRule" -Port 0 } "Add-FirewallRule: port 0 throws (below 1024)"
Assert-Throws { Add-FirewallRule -Name "TestRule" -Port 65536 } "Add-FirewallRule: port 65536 throws (above 65535)"
Assert-Throws { Add-FirewallRule -Name "TestRule" -Port -1 } "Add-FirewallRule: port -1 throws (negative)"

# Port 1024 should pass validation (may fail on netsh without admin, but validation passes)
$port1024Passed = $false
try {
    Add-FirewallRule -Name "TestRule" -Port 1024
    $port1024Passed = $true
} catch {
    if ($_.Exception.Message -like "*Port must be*") {
        $port1024Passed = $false
    } else {
        $port1024Passed = $true  # validation passed, netsh failed
    }
}
Assert-True $port1024Passed "Add-FirewallRule: port 1024 passes port validation"

# Port 65535 should pass validation
$port65535Passed = $false
try {
    Add-FirewallRule -Name "TestRule" -Port 65535
    $port65535Passed = $true
} catch {
    if ($_.Exception.Message -like "*Port must be*") {
        $port65535Passed = $false
    } else {
        $port65535Passed = $true  # validation passed, netsh failed
    }
}
Assert-True $port65535Passed "Add-FirewallRule: port 65535 passes port validation"

# 4g. Add-FirewallRule with invalid protocol
Assert-Throws { Add-FirewallRule -Name "TestRule" -Port 8080 -Protocol "ICMP" } "Add-FirewallRule: protocol 'ICMP' throws"
Assert-Throws { Add-FirewallRule -Name "TestRule" -Port 8080 -Protocol "" } "Add-FirewallRule: empty protocol throws"

# 4h. Remove-FirewallRule with invalid name
Assert-Throws { Remove-FirewallRule -Name "bad name!" } "Remove-FirewallRule: invalid name throws"

# ====================================================================
#  5. CLI region extraction edge cases
# ====================================================================
Write-Host "`n=== 5. CLI Region Extraction Edge Cases ===" -ForegroundColor Cyan

# 5a. Missing region — when a region name is not found, startIdx will be -1
$testContent = @"
#region KnownRegion
Write-Host 'hello'
#endregion KnownRegion
"@

$missingRegionName = "NonExistentRegion"
$startTag = "#region $missingRegionName"
$endTag   = "#endregion $missingRegionName"
$startIdx = $testContent.IndexOf($startTag)
$endIdx   = $testContent.IndexOf($endTag)
Assert-True ($startIdx -eq -1) "CLI region extraction: missing region — startIdx is -1"
Assert-True (-not ($startIdx -ge 0 -and $endIdx -gt $startIdx)) "CLI region extraction: missing region — guard condition prevents extraction"

# 5b. Region with no matching endregion tag
$noEndContent = @"
#region OrphanRegion
Write-Host 'orphan code'
some more code here
"@

$orphanStart = "#region OrphanRegion"
$orphanEnd   = "#endregion OrphanRegion"
$oStartIdx = $noEndContent.IndexOf($orphanStart)
$oEndIdx   = $noEndContent.IndexOf($orphanEnd)
Assert-True ($oStartIdx -ge 0) "CLI region extraction: orphan region — start tag found"
Assert-True ($oEndIdx -eq -1) "CLI region extraction: orphan region — end tag NOT found"
Assert-True (-not ($oStartIdx -ge 0 -and $oEndIdx -gt $oStartIdx)) "CLI region extraction: orphan region — guard condition prevents extraction"

# 5c. Region where endregion appears BEFORE region (reversed order)
$reversedContent = @"
#endregion Backwards
Some code
#region Backwards
"@
$bStartIdx = $reversedContent.IndexOf("#region Backwards")
$bEndIdx   = $reversedContent.IndexOf("#endregion Backwards")
Assert-True ($bEndIdx -lt $bStartIdx) "CLI region extraction: reversed tags — endIdx < startIdx"
Assert-True (-not ($bStartIdx -ge 0 -and $bEndIdx -gt $bStartIdx)) "CLI region extraction: reversed tags — guard condition prevents extraction"

# 5d. Verify that all expected regions actually exist in the real GUI file
$guiFileContent = Get-Content $guiScript -Raw -Encoding UTF8
$expectedRegions = @(
    'App Categories',
    'Files Migration Helpers',
    'Program Inventory',
    'Files Migration Engine',
    'Install Engine',
    'Restore Engine'
)
foreach ($rn in $expectedRegions) {
    $sIdx = $guiFileContent.IndexOf("#region $rn")
    $eIdx = $guiFileContent.IndexOf("#endregion $rn")
    Assert-True ($sIdx -ge 0 -and $eIdx -gt $sIdx) "CLI region extraction: GUI file contains matched '#region $rn' / '#endregion $rn'"
}

# ====================================================================
#  6. New-TransferPin edge cases
# ====================================================================
Write-Host "`n=== 6. New-TransferPin Edge Cases ===" -ForegroundColor Cyan

# 6a. Returns a 6-character string
$pin1 = New-TransferPin
Assert-True ($pin1 -is [string]) "New-TransferPin: returns a string"
Assert-True ($pin1.Length -eq 6) "New-TransferPin: string length is 6"
Assert-Match '^\d{6}$' $pin1 "New-TransferPin: matches 6-digit pattern"

# 6b. PIN is within expected numeric range (100000–999998, since Maximum is exclusive)
$pinInt = [int]$pin1
Assert-True ($pinInt -ge 100000 -and $pinInt -le 999998) "New-TransferPin: numeric value in range [100000, 999998]"

# 6c. Multiple calls should produce different values (randomness check)
# Generate 20 PINs and verify we have at least 2 distinct values
$pins = @()
for ($i = 0; $i -lt 20; $i++) {
    $pins += New-TransferPin
}
$uniquePins = $pins | Select-Object -Unique
Assert-True ($uniquePins.Count -ge 2) "New-TransferPin: 20 calls produce at least 2 distinct PINs (got $($uniquePins.Count))"

# 6d. Every generated PIN should be a valid 6-digit string
$allValid = $true
foreach ($p in $pins) {
    if ($p -notmatch '^\d{6}$') { $allValid = $false; break }
}
Assert-True $allValid "New-TransferPin: all 20 generated PINs match 6-digit pattern"

# 6e. No PIN should start with 0 (since minimum is 100000)
$noneStartWithZero = $true
foreach ($p in $pins) {
    if ($p.StartsWith('0')) { $noneStartWithZero = $false; break }
}
Assert-True $noneStartWithZero "New-TransferPin: no PIN starts with 0"

# ====================================================================
#  Cleanup
# ====================================================================
Write-Host "`n=== Cleanup ===" -ForegroundColor Cyan
try {
    Remove-Item -Path $script:TempRoot -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "  Temp folder cleaned up: $($script:TempRoot)" -ForegroundColor Gray
} catch {
    Write-Host "  Warning: could not fully clean temp folder: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Remove test settings file if created
if (Test-Path $script:SettingsFile) {
    Remove-Item -Path $script:SettingsFile -Force -ErrorAction SilentlyContinue
}

# ====================================================================
#  Summary
# ====================================================================
Write-Host "`n============================================" -ForegroundColor White
Write-Host "  Test Results" -ForegroundColor White
Write-Host "============================================" -ForegroundColor White
Write-Host "  PASSED: $($script:PassCount)" -ForegroundColor Green
Write-Host "  FAILED: $($script:FailCount)" -ForegroundColor $(if ($script:FailCount -gt 0) { 'Red' } else { 'Green' })
if ($script:TestErrors.Count -gt 0) {
    Write-Host "`n  Failed tests:" -ForegroundColor Red
    foreach ($err in $script:TestErrors) {
        Write-Host "    - $err" -ForegroundColor Red
    }
}
Write-Host "============================================`n" -ForegroundColor White

# Exit with non-zero code if any test failed
if ($script:FailCount -gt 0) { exit 1 } else { exit 0 }
