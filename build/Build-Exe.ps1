<#
.SYNOPSIS
    Build-Exe.ps1 — Package LazyTransfer into a standalone EXE

.DESCRIPTION
    Uses ps2exe or similar to compile LazyTransfer into a portable executable.
    Bundles all modules into a single output.

.NOTES
    Requires: ps2exe module (Install-Module ps2exe)
#>

param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "..\dist"),
    [switch]$NoCompress
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSScriptRoot  # LazyTransfer root

Write-Host "LazyTransfer Build Pipeline" -ForegroundColor Cyan
Write-Host "=" * 40

# Module list — all modules to bundle
$modules = @(
    "modules\NetworkEngine.ps1"
)

$mainScript = "LazyTransfer-GUI.ps1"
$cliScript = "LazyTransfer-CLI.ps1"

# Verify all files exist
Write-Host "`nVerifying files..."
$allFiles = @($mainScript, $cliScript) + $modules
foreach ($file in $allFiles) {
    $fullPath = Join-Path $scriptRoot $file
    if (-not (Test-Path $fullPath)) {
        Write-Host "  [MISSING] $file" -ForegroundColor Red
    } else {
        $size = (Get-Item $fullPath).Length
        Write-Host "  [OK] $file ($([Math]::Round($size / 1KB, 1)) KB)" -ForegroundColor Green
    }
}

# Create output directory
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# Check for ps2exe
$hasPs2Exe = Get-Module -ListAvailable -Name ps2exe
if (-not $hasPs2Exe) {
    Write-Host "`nps2exe not found. Install with: Install-Module ps2exe" -ForegroundColor Yellow
    Write-Host "Falling back to folder-based distribution..." -ForegroundColor Yellow

    # Copy all files to dist folder
    $distFolder = Join-Path $OutputPath "LazyTransfer"
    if (Test-Path $distFolder) { Remove-Item $distFolder -Recurse -Force }
    New-Item -Path $distFolder -ItemType Directory -Force | Out-Null

    # Copy main files
    Copy-Item (Join-Path $scriptRoot $mainScript) -Destination $distFolder -Force
    Copy-Item (Join-Path $scriptRoot $cliScript) -Destination $distFolder -Force
    Copy-Item (Join-Path $scriptRoot "Start_LazyTransfer.bat") -Destination $distFolder -Force
    Copy-Item (Join-Path $scriptRoot "README.md") -Destination $distFolder -Force

    # Copy modules
    $modulesDir = Join-Path $distFolder "modules"
    New-Item -Path $modulesDir -ItemType Directory -Force | Out-Null
    foreach ($mod in $modules) {
        $src = Join-Path $scriptRoot $mod
        if (Test-Path $src) {
            Copy-Item $src -Destination $modulesDir -Force
        }
    }

    # Copy profiles
    $profilesSrc = Join-Path $scriptRoot "profiles"
    if (Test-Path $profilesSrc) {
        Copy-Item $profilesSrc -Destination $distFolder -Recurse -Force
    }

    # Create ZIP
    if (-not $NoCompress) {
        $zipPath = Join-Path $OutputPath "LazyTransfer.zip"
        if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
        Compress-Archive -Path "$distFolder\*" -DestinationPath $zipPath -Force
        Write-Host "`nZIP created: $zipPath" -ForegroundColor Green
        $zipSize = (Get-Item $zipPath).Length
        Write-Host "  Size: $([Math]::Round($zipSize / 1MB, 2)) MB"
    }

    Write-Host "`nDist folder: $distFolder" -ForegroundColor Green
    return
}

# Build EXE with ps2exe
Write-Host "`nBuilding EXE..."
$exePath = Join-Path $OutputPath "LazyTransfer.exe"

# Create a merged script that includes all modules inline
$mergedPath = Join-Path $env:TEMP "LazyTransfer-Merged.ps1"
$mergedContent = New-Object System.Text.StringBuilder

# Add module contents first
foreach ($mod in $modules) {
    $modPath = Join-Path $scriptRoot $mod
    if (Test-Path $modPath) {
        [void]$mergedContent.AppendLine("# === Module: $mod ===")
        [void]$mergedContent.AppendLine((Get-Content $modPath -Raw -Encoding UTF8))
        [void]$mergedContent.AppendLine("")
    }
}

# Add main script (skip dot-source lines for modules since they're already inline)
$mainContent = Get-Content (Join-Path $scriptRoot $mainScript) -Raw -Encoding UTF8
$mainContent = $mainContent -replace '(?m)^\s*\.\s+\(Join-Path.*NetworkEngine\.ps1.*$', '# (module already merged)'
[void]$mergedContent.AppendLine($mainContent)

Set-Content -Path $mergedPath -Value $mergedContent.ToString() -Encoding UTF8

try {
    Invoke-ps2exe -InputFile $mergedPath -OutputFile $exePath `
        -Title "LazyTransfer" `
        -Description "PC Migration Tool" `
        -Version "2.5.0.0" `
        -Company "LazyTransfer" `
        -Product "LazyTransfer" `
        -Copyright "(c) 2025" `
        -requireAdmin `
        -noConsole

    Write-Host "`nEXE built: $exePath" -ForegroundColor Green
    $exeSize = (Get-Item $exePath).Length
    Write-Host "  Size: $([Math]::Round($exeSize / 1MB, 2)) MB"
} catch {
    Write-Host "EXE build failed: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    Remove-Item $mergedPath -Force -ErrorAction SilentlyContinue
}

Write-Host "`nBuild complete!" -ForegroundColor Green
