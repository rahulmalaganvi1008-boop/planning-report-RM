$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:8080/")
try {
    $listener.Start()
    Write-Host "Web server started! Listening on http://localhost:8080/converter.html"
} catch {
    Write-Host "Failed to start server: $_"
    exit 1
}

$workspaceDir = if ($PSScriptRoot) { $PSScriptRoot } else { "." }

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $req = $context.Request
        $res = $context.Response
        
        $urlPath = $req.Url.LocalPath
        if ($urlPath -eq "/" -or $urlPath -eq "") {
            $urlPath = "/converter.html"
        }
        
        $filePath = Join-Path $workspaceDir $urlPath.TrimStart('/')
        
        if (Test-Path $filePath -PathType Leaf) {
            $bytes = [System.IO.File]::ReadAllBytes($filePath)
            
            if ($filePath -like "*.html") {
                $res.ContentType = "text/html; charset=utf-8"
            } elseif ($filePath -like "*.js") {
                $res.ContentType = "application/javascript; charset=utf-8"
            } elseif ($filePath -like "*.css") {
                $res.ContentType = "text/css; charset=utf-8"
            } else {
                $res.ContentType = "application/octet-stream"
            }
            
            $res.ContentLength64 = $bytes.Length
            $res.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $res.StatusCode = 404
            $msg = [System.Text.Encoding]::UTF8.GetBytes("File Not Found")
            $res.ContentLength64 = $msg.Length
            $res.OutputStream.Write($msg, 0, $msg.Length)
        }
        $res.OutputStream.Close()
    } catch {
        # Catch connection aborts and keep listening
    }
}
