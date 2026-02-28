<#
.SYNOPSIS
    NetworkEngine.ps1 — Remote transfer module for LazyTransfer v2.4

.DESCRIPTION
    Three transfer modes: HTTP server, Direct TCP, Shared Folder.
    Provides LAN-based file transfer between PCs.
#>

#region Network Helpers

function Get-LocalIPAddresses {
    $ips = @()
    try {
        $adapters = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
            Where-Object { $_.OperationalStatus -eq 'Up' -and $_.NetworkInterfaceType -ne 'Loopback' }
        foreach ($adapter in $adapters) {
            $props = $adapter.GetIPProperties()
            foreach ($addr in $props.UnicastAddresses) {
                if ($addr.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
                    $ips += [PSCustomObject]@{
                        IP        = $addr.Address.ToString()
                        Interface = $adapter.Name
                        Type      = $adapter.NetworkInterfaceType.ToString()
                    }
                }
            }
        }
    } catch {
        Write-Log "Failed to enumerate network interfaces: $($_.Exception.Message)" -Level 'WARN'
    }
    return $ips
}

function Get-PrimaryLocalIP {
    $ips = Get-LocalIPAddresses
    # Prefer non-virtual, non-169.254 addresses
    $preferred = $ips | Where-Object { $_.IP -notlike '169.254.*' -and $_.IP -notlike '127.*' }
    if ($preferred) { return $preferred[0].IP }
    if ($ips) { return $ips[0].IP }
    return "127.0.0.1"
}

function Test-PortAvailable {
    param([Parameter(Mandatory=$true)][int]$Port)
    $listener = $null
    try {
        $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
        $listener.Start()
        return $true
    } catch {
        return $false
    } finally {
        if ($listener) { try { $listener.Stop() } catch { } }
    }
}

function Add-FirewallRule {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][int]$Port,
        [string]$Protocol = "TCP"
    )
    # Validate inputs to prevent netsh command injection
    if ($Name -notmatch '^[a-zA-Z0-9\-]+$') { throw "Invalid firewall rule name: $Name" }
    if ($Port -lt 1024 -or $Port -gt 65535) { throw "Port must be between 1024 and 65535" }
    if ($Protocol -notin @('TCP', 'UDP')) { throw "Invalid protocol: $Protocol" }
    try {
        $null = & netsh advfirewall firewall add rule name="$Name" dir=in action=allow protocol=$Protocol localport=$Port 2>&1
        Write-Log "Firewall rule added: $Name (port $Port)" -Level 'INFO'
        return $true
    } catch {
        Write-Log "Failed to add firewall rule: $($_.Exception.Message)" -Level 'WARN'
        return $false
    }
}

function Remove-FirewallRule {
    param([Parameter(Mandatory=$true)][string]$Name)
    if ($Name -notmatch '^[a-zA-Z0-9\-]+$') { throw "Invalid firewall rule name: $Name" }
    try {
        $null = & netsh advfirewall firewall delete rule name="$Name" 2>&1
        Write-Log "Firewall rule removed: $Name" -Level 'INFO'
        return $true
    } catch {
        Write-Log "Failed to remove firewall rule: $($_.Exception.Message)" -Level 'WARN'
        return $false
    }
}
#endregion Network Helpers

#region HTTP Server Mode

$script:HttpListener = $null
$script:HttpServerRunning = $false
$script:HttpServerFolder = $null
$script:HttpFirewallRuleName = $null
$script:HttpServedFiles = $null
$script:HttpLandingPageHtml = $null
$script:TransferPin = $null

function New-TransferPin {
    # Generate a random 6-digit PIN for transfer authentication
    return (Get-Random -Minimum 100000 -Maximum 999999).ToString()
}

function Start-TransferServer {
    param(
        [Parameter(Mandatory=$true)][string]$SourceFolder,
        [int]$Port = 8642,
        [scriptblock]$OnProgress
    )

    if (-not (Test-Path $SourceFolder)) { throw "Source folder not found: $SourceFolder" }
    if ($Port -lt 1024 -or $Port -gt 65535) { throw "Port must be between 1024 and 65535" }
    if (-not (Test-PortAvailable -Port $Port)) { throw "Port $Port is already in use." }

    $script:HttpServerFolder = $SourceFolder
    $script:HttpServerRunning = $true

    # Generate pairing PIN for access control
    $script:TransferPin = New-TransferPin
    Write-Log "Transfer PIN: $($script:TransferPin) — share this with the receiving PC" -Level 'SUCCESS'

    $script:HttpFirewallRuleName = "LazyTransfer-HTTP-$Port"
    Add-FirewallRule -Name $script:HttpFirewallRuleName -Port $Port | Out-Null

    # Cache file list and landing page HTML
    $script:HttpServedFiles = Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue
    $script:HttpLandingPageHtml = Get-TransferLandingPage -SourceFolder $SourceFolder -Port $Port

    $prefix = "http://+:$Port/"
    $script:HttpListener = New-Object System.Net.HttpListener
    $script:HttpListener.Prefixes.Add($prefix)

    try {
        $script:HttpListener.Start()
    } catch {
        Remove-FirewallRule -Name $script:HttpFirewallRuleName | Out-Null
        $script:HttpServerRunning = $false
        $script:HttpFirewallRuleName = $null
        throw "Failed to start HTTP listener: $($_.Exception.Message)"
    }

    $localIP = Get-PrimaryLocalIP
    $serverUrl = "http://${localIP}:${Port}/"
    Write-Log "HTTP Transfer Server started at $serverUrl" -Level 'SUCCESS'
    Write-Log "Serving folder: $SourceFolder"

    if ($OnProgress) { & $OnProgress "Server" 0 0 "Listening on $serverUrl" }

    # Serve requests — wrapped in try/finally to ensure firewall cleanup on crash/exit
    try {
        while ($script:HttpServerRunning -and $script:HttpListener.IsListening) {
            try {
                $contextTask = $script:HttpListener.GetContextAsync()
                while (-not $contextTask.IsCompleted) {
                    if (-not $script:HttpServerRunning) { break }
                    Update-GuiStatus
                    Start-Sleep -Milliseconds 100
                }
                if (-not $script:HttpServerRunning) { break }

                $context = $contextTask.Result
                $request = $context.Request
                $response = $context.Response
                $path = $request.Url.LocalPath

                Write-Log "HTTP request: $($request.HttpMethod) $path"

                # Verify PIN authentication (passed as ?pin= query parameter)
                $queryPin = $request.QueryString["pin"]
                if ($script:TransferPin -and $queryPin -ne $script:TransferPin) {
                    $response.StatusCode = 403
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes("Access denied. Append ?pin=YOUR_PIN to the URL.")
                    $response.ContentType = "text/plain"
                    $response.ContentLength64 = $buffer.Length
                    $response.OutputStream.Write($buffer, 0, $buffer.Length)
                    $response.OutputStream.Close()
                    Write-Log "Rejected request — invalid PIN from $($request.RemoteEndPoint)" -Level 'WARN'
                    continue
                }

                if ($path -eq '/' -or $path -eq '/index.html') {
                    # Serve cached landing page
                    $html = if ($script:HttpLandingPageHtml) { $script:HttpLandingPageHtml } else { Get-TransferLandingPage -SourceFolder $SourceFolder -Port $Port }
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                    $response.ContentType = "text/html; charset=utf-8"
                    $response.ContentLength64 = $buffer.Length
                    $response.OutputStream.Write($buffer, 0, $buffer.Length)
                    $response.OutputStream.Close()
                }
                elseif ($path -eq '/download-all') {
                    # Stream ZIP on-the-fly
                    Write-Log "Client downloading all files as ZIP..."
                    if ($OnProgress) { & $OnProgress "Download" 0 0 "Client downloading ZIP..." }

                    $response.ContentType = "application/zip"
                    $response.Headers.Add("Content-Disposition", "attachment; filename=`"LazyTransfer-Bundle.zip`"")
                    $response.SendChunked = $true

                    try {
                        Add-Type -AssemblyName System.IO.Compression

                        $archive = New-Object System.IO.Compression.ZipArchive($response.OutputStream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
                        $files = Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue
                        $totalFiles = $files.Count
                        $currentFile = 0

                        foreach ($file in $files) {
                            if (-not $script:HttpServerRunning) { break }
                            $currentFile++
                            $relativePath = $file.FullName.Substring($SourceFolder.Length).TrimStart('\', '/')
                            $entry = $archive.CreateEntry($relativePath, [System.IO.Compression.CompressionLevel]::Fastest)
                            $entryStream = $entry.Open()
                            try {
                                $fileStream = [System.IO.File]::OpenRead($file.FullName)
                                try {
                                    $fileStream.CopyTo($entryStream)
                                } finally {
                                    $fileStream.Close()
                                }
                            } finally {
                                $entryStream.Close()
                            }

                            if ($OnProgress -and ($currentFile % 10 -eq 0)) {
                                & $OnProgress "Download" $currentFile $totalFiles "Streaming file $currentFile of $totalFiles"
                                Update-GuiStatus
                            }
                        }

                        $archive.Dispose()
                        Write-Log "ZIP download complete ($totalFiles files)" -Level 'SUCCESS'
                        if ($OnProgress) { & $OnProgress "Download" $totalFiles $totalFiles "Download complete" }
                    } catch {
                        Write-Log "ZIP streaming error: $($_.Exception.Message)" -Level 'ERROR'
                    }

                    try { $response.OutputStream.Close() } catch { }
                }
                else {
                    # 404
                    $response.StatusCode = 404
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes("Not Found")
                    $response.OutputStream.Write($buffer, 0, $buffer.Length)
                    $response.OutputStream.Close()
                }
            } catch [System.ObjectDisposedException] {
                break
            } catch {
                if ($script:HttpServerRunning) {
                    Write-Log "HTTP server error: $($_.Exception.Message)" -Level 'WARN'
                }
            }
        }
    } finally {
        Stop-TransferServer
    }

    return $serverUrl
}

function Get-TransferLandingPage {
    param(
        [Parameter(Mandatory=$true)][string]$SourceFolder,
        [int]$Port = 8642
    )

    Add-Type -AssemblyName System.Web
    $files = Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue
    $totalSize = ($files | Measure-Object -Property Length -Sum).Sum
    if (-not $totalSize) { $totalSize = 0 }
    $totalSizeStr = Format-FileSize $totalSize

    $fileListHtml = New-Object System.Text.StringBuilder
    foreach ($file in ($files | Sort-Object FullName)) {
        $relativePath = $file.FullName.Substring($SourceFolder.Length).TrimStart('\', '/')
        $safeName = [System.Web.HttpUtility]::HtmlEncode($relativePath)
        $sizeStr = Format-FileSize $file.Length
        [void]$fileListHtml.AppendLine("<tr><td>$safeName</td><td>$sizeStr</td></tr>")
    }

    $safeFolder = [System.Web.HttpUtility]::HtmlEncode((Split-Path $SourceFolder -Leaf))

    return @"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>LazyTransfer - File Transfer</title>
<style>
body { font-family: Segoe UI, Arial; margin: 0; padding: 20px; background: #1e1e1e; color: #eee; }
h1 { color: #e94560; margin-bottom: 5px; }
.info { color: #888; margin-bottom: 20px; }
.download-btn {
    display: inline-block; padding: 15px 30px; background: #e94560; color: white;
    text-decoration: none; border-radius: 8px; font-size: 16px; font-weight: bold;
    margin: 15px 0;
}
.download-btn:hover { background: #c73a52; }
table { border-collapse: collapse; width: 100%; margin-top: 15px; }
th, td { border: 1px solid #444; padding: 8px 12px; text-align: left; }
th { background: #333; color: #e94560; }
tr:nth-child(even) { background: #2a2a2a; }
.stats { color: #aaa; font-size: 14px; }
</style></head><body>
<h1>LazyTransfer</h1>
<p class="info">Serving: $safeFolder</p>
<a href="/download-all" class="download-btn">Download Everything (ZIP)</a>
<p class="stats">$($files.Count) files | $totalSizeStr total</p>
<table><tr><th>File</th><th>Size</th></tr>
$($fileListHtml.ToString())
</table>
</body></html>
"@
}

function Stop-TransferServer {
    $script:HttpServerRunning = $false
    if ($script:HttpListener) {
        try {
            $script:HttpListener.Stop()
            $script:HttpListener.Close()
        } catch { }
        $script:HttpListener = $null
    }
    if ($script:HttpFirewallRuleName) {
        Remove-FirewallRule -Name $script:HttpFirewallRuleName | Out-Null
        $script:HttpFirewallRuleName = $null
    }
    $script:HttpServedFiles = $null
    $script:HttpLandingPageHtml = $null
    Write-Log "HTTP Transfer Server stopped" -Level 'INFO'
}

function Get-TransferFromServer {
    param(
        [Parameter(Mandatory=$true)][string]$ServerUrl,
        [Parameter(Mandatory=$true)][string]$OutputFolder,
        [string]$Pin,
        [scriptblock]$OnProgress
    )

    if (-not (Test-Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    $pinQuery = if ($Pin) { "?pin=$Pin" } else { "" }
    $downloadUrl = $ServerUrl.TrimEnd('/') + "/download-all$pinQuery"
    $zipPath = Join-Path $OutputFolder "LazyTransfer-Bundle.zip"

    Write-Log "Downloading from $downloadUrl..."
    if ($OnProgress) { & $OnProgress "Download" 0 0 "Connecting to server..." }

    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($downloadUrl, $zipPath)
        Write-Log "Download complete: $zipPath" -Level 'SUCCESS'

        if ($OnProgress) { & $OnProgress "Extract" 0 0 "Extracting ZIP..." }
        $extractFolder = Join-Path $OutputFolder "LazyTransfer-Bundle"
        if (Test-Path $extractFolder) { Remove-Item $extractFolder -Recurse -Force }

        # Safe ZIP extraction — prevent ZIP slip (path traversal in archive entries)
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $resolvedExtract = [System.IO.Path]::GetFullPath($extractFolder)
        $archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
        try {
            foreach ($entry in $archive.Entries) {
                if ([string]::IsNullOrWhiteSpace($entry.FullName)) { continue }
                $destPath = [System.IO.Path]::GetFullPath((Join-Path $extractFolder $entry.FullName))
                if (-not $destPath.StartsWith($resolvedExtract)) {
                    Write-Log "ZIP slip blocked: $($entry.FullName)" -Level 'ERROR'
                    continue
                }
                if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) {
                    if (-not (Test-Path $destPath)) { New-Item -Path $destPath -ItemType Directory -Force | Out-Null }
                } else {
                    $destDir = Split-Path -Parent $destPath
                    if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force | Out-Null }
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destPath, $true)
                }
            }
        } finally {
            $archive.Dispose()
        }

        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
        Write-Log "Extracted to: $extractFolder" -Level 'SUCCESS'
        if ($OnProgress) { & $OnProgress "Complete" 1 1 "Transfer complete" }

        return [PSCustomObject]@{
            Success = $true
            Path    = $extractFolder
            Errors  = @()
        }
    } catch {
        Write-Log "Download failed: $($_.Exception.Message)" -Level 'ERROR'
        # Clean up partial ZIP on failure
        if (Test-Path $zipPath) { Remove-Item $zipPath -Force -ErrorAction SilentlyContinue }
        if ($OnProgress) { & $OnProgress "Error" 0 0 $_.Exception.Message }
        return [PSCustomObject]@{
            Success = $false
            Path    = $null
            Errors  = @($_.Exception.Message)
        }
    } finally {
        if ($webClient) { $webClient.Dispose() }
    }
}
#endregion HTTP Server Mode

#region Direct TCP Mode

function Start-DirectSend {
    param(
        [Parameter(Mandatory=$true)][string]$SourceFolder,
        [int]$Port = 8643,
        [scriptblock]$OnProgress
    )

    if (-not (Test-Path $SourceFolder)) { throw "Source folder not found: $SourceFolder" }
    if ($Port -lt 1024 -or $Port -gt 65535) { throw "Port must be between 1024 and 65535" }
    if (-not (Test-PortAvailable -Port $Port)) { throw "Port $Port is already in use." }

    $fwRuleName = "LazyTransfer-TCP-$Port"
    Add-FirewallRule -Name $fwRuleName -Port $Port | Out-Null

    $files = Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue
    $fileCount = $files.Count
    $totalSize = [long]($files | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum)

    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Any, $Port)
    $listener.Start()

    # Generate pairing PIN for TCP authentication
    $tcpPin = New-TransferPin
    Write-Log "TCP Transfer PIN: $tcpPin — share this with the receiving PC" -Level 'SUCCESS'

    $localIP = Get-PrimaryLocalIP
    Write-Log "TCP Sender waiting for connection on ${localIP}:${Port}" -Level 'INFO'
    Write-Log "Files: $fileCount | Total: $(Format-FileSize $totalSize)"
    if ($OnProgress) { & $OnProgress "Waiting" 0 $totalSize "PIN: $tcpPin | Waiting on ${localIP}:${Port}..." }

    try {
        # Wait for connection with cancellation support
        $acceptTask = $listener.AcceptTcpClientAsync()
        while (-not $acceptTask.IsCompleted) {
            if ($script:CancelRequested) {
                $listener.Stop()
                Remove-FirewallRule -Name $fwRuleName | Out-Null
                return [PSCustomObject]@{ Success = $false; Errors = @("Cancelled") }
            }
            Update-GuiStatus
            Start-Sleep -Milliseconds 100
        }

        $client = $acceptTask.Result
        $stream = $client.GetStream()
        Write-Log "Receiver connected from $($client.Client.RemoteEndPoint)" -Level 'SUCCESS'

        # PIN authentication: send PIN, receiver must echo it back
        $writer = New-Object System.IO.BinaryWriter($stream)
        $reader = New-Object System.IO.BinaryReader($stream)
        $pinBytes = [System.Text.Encoding]::UTF8.GetBytes($tcpPin)
        $writer.Write([int32]$pinBytes.Length)
        $writer.Write($pinBytes)
        $writer.Flush()
        # Read PIN response from receiver
        $responseLen = $reader.ReadInt32()
        if ($responseLen -lt 1 -or $responseLen -gt 64) {
            throw "Invalid PIN response"
        }
        $responseBytes = $reader.ReadBytes($responseLen)
        $responsePin = [System.Text.Encoding]::UTF8.GetString($responseBytes)
        if ($responsePin -ne $tcpPin) {
            Write-Log "TCP PIN mismatch — access denied" -Level 'ERROR'
            throw "PIN mismatch — receiver sent wrong PIN"
        }
        Write-Log "TCP PIN verified" -Level 'SUCCESS'

        if ($OnProgress) { & $OnProgress "Connected" 0 $totalSize "Sending file list..." }

        # Binary protocol: file count (int32), total size (int64)
        $writer.Write([int32]$fileCount)
        $writer.Write([int64]$totalSize)

        $sentBytes = [long]0
        $currentFile = 0
        $buffer = New-Object byte[] 65536

        foreach ($file in $files) {
            if ($script:CancelRequested) { break }
            $currentFile++

            $relativePath = $file.FullName.Substring($SourceFolder.Length).TrimStart('\', '/')
            $nameBytes = [System.Text.Encoding]::UTF8.GetBytes($relativePath)

            # Per-file: name length (int32), name (UTF8), file size (int64), data
            $writer.Write([int32]$nameBytes.Length)
            $writer.Write($nameBytes)
            $writer.Write([int64]$file.Length)

            $fileStream = [System.IO.File]::OpenRead($file.FullName)
            try {
                $bytesRead = 0
                while (($bytesRead = $fileStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $stream.Write($buffer, 0, $bytesRead)
                    $sentBytes += $bytesRead
                    if ($OnProgress -and ($currentFile % 5 -eq 0 -or $bytesRead -lt $buffer.Length)) {
                        & $OnProgress "Sending" $sentBytes $totalSize "File $currentFile/$fileCount`: $relativePath"
                        Update-GuiStatus
                    }
                }
            } finally {
                $fileStream.Close()
            }
        }

        $writer.Flush()
        Write-Log "TCP send complete: $fileCount files, $(Format-FileSize $sentBytes)" -Level 'SUCCESS'
        if ($OnProgress) { & $OnProgress "Complete" $sentBytes $totalSize "Send complete" }

        return [PSCustomObject]@{
            Success    = $true
            FilesSent  = $fileCount
            BytesSent  = $sentBytes
            Errors     = @()
        }
    } catch {
        Write-Log "TCP send error: $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Errors = @($_.Exception.Message) }
    } finally {
        if ($writer) { try { $writer.Close() } catch { } }
        if ($stream) { try { $stream.Close() } catch { } }
        if ($client) { try { $client.Close() } catch { } }
        $listener.Stop()
        Remove-FirewallRule -Name $fwRuleName | Out-Null
    }
}

function Start-DirectReceive {
    param(
        [Parameter(Mandatory=$true)][string]$RemoteIP,
        [Parameter(Mandatory=$true)][string]$OutputFolder,
        [int]$Port = 8643,
        [string]$Pin,
        [scriptblock]$OnProgress
    )

    if ($Port -lt 1024 -or $Port -gt 65535) { throw "Port must be between 1024 and 65535" }
    # Validate IP address format
    $ipAddr = $null
    if (-not [System.Net.IPAddress]::TryParse($RemoteIP, [ref]$ipAddr)) {
        throw "Invalid IP address: $RemoteIP"
    }
    if (-not (Test-Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    Write-Log "Connecting to TCP sender at ${RemoteIP}:${Port}..."
    if ($OnProgress) { & $OnProgress "Connecting" 0 0 "Connecting to ${RemoteIP}:${Port}..." }

    $client = $null
    $stream = $null
    $reader = $null

    try {
        $client = New-Object System.Net.Sockets.TcpClient
        # Async connect with 10-second timeout to avoid indefinite GUI freeze
        $connectTask = $client.ConnectAsync($RemoteIP, $Port)
        if (-not $connectTask.Wait(10000)) {
            $client.Close()
            throw "Connection timed out after 10 seconds — check that the sender is running"
        }
        if ($connectTask.IsFaulted) { throw $connectTask.Exception.InnerException }
        $stream = $client.GetStream()
        $reader = New-Object System.IO.BinaryReader($stream)
        $writer = New-Object System.IO.BinaryWriter($stream)

        Write-Log "Connected to sender" -Level 'SUCCESS'

        # PIN authentication: read sender's PIN, echo back user-provided PIN
        $pinLen = $reader.ReadInt32()
        if ($pinLen -lt 1 -or $pinLen -gt 64) { throw "Invalid PIN exchange" }
        $pinBytes = $reader.ReadBytes($pinLen)
        $senderPin = [System.Text.Encoding]::UTF8.GetString($pinBytes)
        # Send back the PIN the user provided
        $userPin = if ($Pin) { $Pin } else { $senderPin }
        $userPinBytes = [System.Text.Encoding]::UTF8.GetBytes($userPin)
        $writer.Write([int32]$userPinBytes.Length)
        $writer.Write($userPinBytes)
        $writer.Flush()
        if ($userPin -ne $senderPin) {
            throw "PIN mismatch — check the PIN displayed on the sender"
        }
        Write-Log "PIN verified" -Level 'SUCCESS'

        # Read header with bounds validation
        $fileCount = $reader.ReadInt32()
        $totalSize = $reader.ReadInt64()
        if ($fileCount -lt 0 -or $fileCount -gt 500000) {
            throw "Invalid file count from sender: $fileCount"
        }
        if ($totalSize -lt 0 -or $totalSize -gt 10TB) {
            throw "Invalid total size from sender: $(Format-FileSize $totalSize)"
        }
        Write-Log "Receiving: $fileCount files, $(Format-FileSize $totalSize)"
        if ($OnProgress) { & $OnProgress "Receiving" 0 $totalSize "Receiving $fileCount files..." }

        $receivedBytes = [long]0
        $receivedFiles = 0
        $errors = @()
        $buffer = New-Object byte[] 65536

        for ($i = 0; $i -lt $fileCount; $i++) {
            if ($script:CancelRequested) { break }

            # Read file metadata
            $nameLen = $reader.ReadInt32()
            if ($nameLen -lt 1 -or $nameLen -gt 4096) {
                throw "Invalid file name length: $nameLen"
            }
            $nameBytes = $reader.ReadBytes($nameLen)
            $relativePath = [System.Text.Encoding]::UTF8.GetString($nameBytes)
            $fileSize = $reader.ReadInt64()
            if ($fileSize -lt 0 -or $fileSize -gt 100GB) {
                throw "Invalid file size: $fileSize"
            }

            # Prevent path traversal — ensure destPath stays within OutputFolder
            $destPath = [System.IO.Path]::GetFullPath((Join-Path $OutputFolder $relativePath))
            $resolvedOutput = [System.IO.Path]::GetFullPath($OutputFolder)
            if (-not $destPath.StartsWith($resolvedOutput)) {
                Write-Log "Path traversal blocked: $relativePath" -Level 'ERROR'
                $errors += "Blocked path traversal: $relativePath"
                # Drain this file's bytes from the stream so we stay in sync
                $remaining = $fileSize
                while ($remaining -gt 0) {
                    $toRead = [Math]::Min($remaining, $buffer.Length)
                    $bytesRead = $stream.Read($buffer, 0, $toRead)
                    if ($bytesRead -eq 0) { break }
                    $remaining -= $bytesRead
                }
                continue
            }

            $destDir = Split-Path -Parent $destPath
            if (-not (Test-Path $destDir)) {
                New-Item -Path $destDir -ItemType Directory -Force | Out-Null
            }

            $fileComplete = $false
            try {
                $fileStream = [System.IO.File]::Create($destPath)
                try {
                    $remaining = $fileSize
                    while ($remaining -gt 0) {
                        $toRead = [Math]::Min($remaining, $buffer.Length)
                        $bytesRead = $stream.Read($buffer, 0, $toRead)
                        if ($bytesRead -eq 0) { throw "Connection closed prematurely" }
                        $fileStream.Write($buffer, 0, $bytesRead)
                        $remaining -= $bytesRead
                        $receivedBytes += $bytesRead
                    }
                    $fileComplete = $true
                } finally {
                    $fileStream.Close()
                    # Delete partial files on failure to prevent truncated data
                    if (-not $fileComplete -and (Test-Path $destPath)) {
                        Remove-Item $destPath -Force -ErrorAction SilentlyContinue
                    }
                }
                $receivedFiles++
            } catch {
                $errors += "Failed: $relativePath - $($_.Exception.Message)"
                Write-Log "Error receiving $relativePath`: $($_.Exception.Message)" -Level 'WARN'
                # Drain remaining bytes to keep stream protocol in sync
                if ($remaining -gt 0) {
                    try {
                        while ($remaining -gt 0) {
                            $toRead = [Math]::Min($remaining, $buffer.Length)
                            $bytesRead = $stream.Read($buffer, 0, $toRead)
                            if ($bytesRead -eq 0) { throw "Connection lost" }
                            $remaining -= $bytesRead
                        }
                    } catch {
                        Write-Log "Stream desynchronized — aborting transfer" -Level 'ERROR'
                        break
                    }
                }
            }

            if ($OnProgress -and (($i + 1) % 5 -eq 0 -or $i -eq $fileCount - 1)) {
                & $OnProgress "Receiving" $receivedBytes $totalSize "File $($i + 1)/$fileCount`: $relativePath"
                Update-GuiStatus
            }
        }

        Write-Log "TCP receive complete: $receivedFiles files, $(Format-FileSize $receivedBytes)" -Level 'SUCCESS'
        if ($OnProgress) { & $OnProgress "Complete" $receivedBytes $totalSize "Receive complete" }

        return [PSCustomObject]@{
            Success       = ($errors.Count -eq 0)
            FilesReceived = $receivedFiles
            BytesReceived = $receivedBytes
            Path          = $OutputFolder
            Errors        = $errors
        }
    } catch {
        Write-Log "TCP receive error: $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{
            Success = $false
            FilesReceived = 0
            BytesReceived = 0
            Path = $OutputFolder
            Errors = @($_.Exception.Message)
        }
    } finally {
        if ($reader) { try { $reader.Close() } catch { } }
        if ($stream) { try { $stream.Close() } catch { } }
        if ($client) { try { $client.Close() } catch { } }
    }
}
#endregion Direct TCP Mode

#region Shared Folder Mode

function Export-ToSharedFolder {
    param(
        [Parameter(Mandatory=$true)][string]$SourceFolder,
        [Parameter(Mandatory=$true)][string]$DestinationUNC,
        [scriptblock]$OnProgress
    )

    if (-not (Test-Path $SourceFolder)) { throw "Source folder not found: $SourceFolder" }
    if (-not (Test-Path $DestinationUNC)) {
        try {
            New-Item -Path $DestinationUNC -ItemType Directory -Force | Out-Null
        } catch {
            throw "Cannot access destination: $DestinationUNC - $($_.Exception.Message)"
        }
    }

    $files = Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue
    $totalSize = [long]($files | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum)
    $totalFiles = $files.Count

    Write-Log "Copying $totalFiles files ($(Format-FileSize $totalSize)) to $DestinationUNC"
    if ($OnProgress) { & $OnProgress "Copying" 0 $totalSize "Starting copy to shared folder..." }

    $roboArgs = @($SourceFolder, $DestinationUNC, '/E', '/COPY:DAT', '/R:2', '/W:2', '/NP', '/MT:4', '/NFL', '/NDL')
    $roboProcess = Start-Process -FilePath "robocopy" -ArgumentList $roboArgs -Wait -PassThru -WindowStyle Hidden

    if ($roboProcess.ExitCode -le 7) {
        Write-Log "Copy to shared folder complete" -Level 'SUCCESS'
        if ($OnProgress) { & $OnProgress "Complete" $totalSize $totalSize "Copy complete" }
        return [PSCustomObject]@{
            Success    = $true
            TotalFiles = $totalFiles
            TotalBytes = $totalSize
            Errors     = @()
        }
    } else {
        Write-Log "Robocopy error (exit code $($roboProcess.ExitCode))" -Level 'ERROR'
        if ($OnProgress) { & $OnProgress "Error" 0 $totalSize "Robocopy error (exit $($roboProcess.ExitCode))" }
        return [PSCustomObject]@{
            Success    = $false
            TotalFiles = $totalFiles
            TotalBytes = $totalSize
            Errors     = @("Robocopy exit code: $($roboProcess.ExitCode)")
        }
    }
}

function Import-FromSharedFolder {
    param(
        [Parameter(Mandatory=$true)][string]$SourceUNC,
        [Parameter(Mandatory=$true)][string]$OutputFolder,
        [scriptblock]$OnProgress
    )

    if (-not (Test-Path $SourceUNC)) { throw "Source path not accessible: $SourceUNC" }
    if (-not (Test-Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    $files = Get-ChildItem -Path $SourceUNC -Recurse -File -ErrorAction SilentlyContinue
    $totalSize = [long]($files | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum)
    $totalFiles = $files.Count

    Write-Log "Importing $totalFiles files ($(Format-FileSize $totalSize)) from $SourceUNC"
    if ($OnProgress) { & $OnProgress "Copying" 0 $totalSize "Importing from shared folder..." }

    $roboArgs = @($SourceUNC, $OutputFolder, '/E', '/COPY:DAT', '/R:2', '/W:2', '/NP', '/MT:4', '/NFL', '/NDL')
    $roboProcess = Start-Process -FilePath "robocopy" -ArgumentList $roboArgs -Wait -PassThru -WindowStyle Hidden

    if ($roboProcess.ExitCode -le 7) {
        Write-Log "Import from shared folder complete" -Level 'SUCCESS'
        if ($OnProgress) { & $OnProgress "Complete" $totalSize $totalSize "Import complete" }
        return [PSCustomObject]@{
            Success    = $true
            Path       = $OutputFolder
            TotalFiles = $totalFiles
            TotalBytes = $totalSize
            Errors     = @()
        }
    } else {
        Write-Log "Robocopy error (exit code $($roboProcess.ExitCode))" -Level 'ERROR'
        return [PSCustomObject]@{
            Success    = $false
            Path       = $OutputFolder
            TotalFiles = $totalFiles
            TotalBytes = $totalSize
            Errors     = @("Robocopy exit code: $($roboProcess.ExitCode)")
        }
    }
}
#endregion Shared Folder Mode
