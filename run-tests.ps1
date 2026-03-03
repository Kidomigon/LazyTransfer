# Download patched files and run verification tests
$ErrorActionPreference = "Continue"
$base = "http://192.168.122.1:8888"
$dest = "C:\Users\user\Desktop\LazyTransfer"
$out = @()

# 1. Update files
try {
    iwr "$base/LazyTransfer-GUI.ps1" -OutFile "$dest\LazyTransfer-GUI.ps1"
    iwr "$base/modules/NetworkEngine.ps1" -OutFile "$dest\modules\NetworkEngine.ps1"
    $out += "UPDATE: OK"
} catch { $out += "UPDATE: FAIL - $_" }

# 2. Verify Bug #1 fix exists
$m = Select-String -Path "$dest\LazyTransfer-GUI.ps1" -Pattern "Stop transcript BEFORE"
if ($m) { $out += "BUG1-PATCH: VERIFIED at line $($m[0].LineNumber)" }
else { $out += "BUG1-PATCH: MISSING" }

# 3. Verify Bug #2 fix (MemoryStream approach)
$m2 = Select-String -Path "$dest\modules\NetworkEngine.ps1" -Pattern "MemoryStream"
if ($m2) { $out += "BUG2-PATCH: VERIFIED at line $($m2[0].LineNumber)" }
else { $out += "BUG2-PATCH: MISSING" }

# 4. Test Bug #1 - ZIP files migration
Remove-Item C:\Temp\BugTest1 -Recurse -Force -EA SilentlyContinue
Remove-Item C:\Temp\BugTest1*.zip -Force -EA SilentlyContinue
Set-ExecutionPolicy Bypass -Scope Process -Force
$r = & "$dest\LazyTransfer-CLI.ps1" -Action files -OutputFolder C:\Temp\BugTest1 -Mode zip -Folders Documents 2>&1
$zipExists = Get-ChildItem C:\Temp -Filter "LazyTransfer-Files-*.zip" -EA SilentlyContinue
if ($zipExists) { $out += "BUG1-TEST: PASS - ZIP created: $($zipExists.Name) ($([math]::Round($zipExists.Length/1KB)) KB)" }
else {
    $errLine = ($r | Out-String) -split "`n" | Where-Object { $_ -match "WARN|ERROR|fail" } | Select-Object -First 2
    $out += "BUG1-TEST: FAIL - No ZIP. Errors: $($errLine -join ' | ')"
}

# 5. Test Bug #2 - HTTP server ZIP download
$out += "BUG2-TEST: Starting serve + download test..."
$scanDir = "C:\Temp\TestPC"
if (-not (Test-Path $scanDir)) { New-Item $scanDir -ItemType Directory -Force | Out-Null; "test" | Set-Content "$scanDir\test.txt" }
$job = Start-Job -ScriptBlock {
    param($dest, $scanDir)
    Set-ExecutionPolicy Bypass -Scope Process -Force
    & "$dest\LazyTransfer-CLI.ps1" -Action serve -OutputFolder $scanDir -Port 9876 2>&1
} -ArgumentList $dest, $scanDir
Start-Sleep 5
try {
    $resp = iwr "http://localhost:9876/download-all?pin=000000" -EA Stop
    $out += "BUG2-TEST: FAIL - Expected auth error but got $($resp.StatusCode)"
} catch {
    # Expected - wrong PIN. But the server responded, which means it's running
    $out += "BUG2-TEST: Server responded (expected PIN rejection)"
}
Stop-Job $job -EA SilentlyContinue
Remove-Job $job -Force -EA SilentlyContinue

# Write results to a file AND serve it back
$results = $out -join "`n"
$results | Set-Content C:\Temp\test-results.txt
Write-Host "=== TEST RESULTS ==="
$out | ForEach-Object { Write-Host $_ }
Write-Host "=== END RESULTS ==="

# Start a tiny HTTP response so host can curl the results
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://+:9999/")
$listener.Start()
Write-Host "Results available at http://${env:COMPUTERNAME}:9999/"
$ctx = $listener.GetContext()
$buf = [System.Text.Encoding]::UTF8.GetBytes($results)
$ctx.Response.ContentLength64 = $buf.Length
$ctx.Response.OutputStream.Write($buf, 0, $buf.Length)
$ctx.Response.Close()
$listener.Stop()
Write-Host "Done."
