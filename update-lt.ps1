$base = "http://192.168.122.1:8888"
$dest = "C:\Users\user\Desktop\LazyTransfer"
Invoke-WebRequest "$base/LazyTransfer-GUI.ps1" -OutFile "$dest\LazyTransfer-GUI.ps1"
Invoke-WebRequest "$base/modules/NetworkEngine.ps1" -OutFile "$dest\modules\NetworkEngine.ps1"
Write-Host "PATCHED: GUI + NetworkEngine updated"
# Verify the patch
$match = Select-String -Path "$dest\LazyTransfer-GUI.ps1" -Pattern "Stop transcript BEFORE"
if ($match) { Write-Host "VERIFIED: Bug #1 fix present at line $($match[0].LineNumber)" }
else { Write-Host "ERROR: Bug #1 fix NOT found" }
