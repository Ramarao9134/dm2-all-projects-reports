# Simple HTTP Server for DM2 Project
$port = 8080
$path = Get-Location

Write-Host "🚀 Starting DM2 Project Management System..." -ForegroundColor Green
Write-Host "📁 Serving files from: $path" -ForegroundColor Yellow
Write-Host "🌐 Server running at: http://localhost:$port" -ForegroundColor Cyan
Write-Host "📱 Open your browser and go to: http://localhost:$port" -ForegroundColor Magenta
Write-Host ""
Write-Host "🔐 Login Credentials:" -ForegroundColor White
Write-Host "   Admin: admin@dm2.com / admin123" -ForegroundColor Gray
Write-Host "   Manager: manager@dm2.com / manager123" -ForegroundColor Gray
Write-Host "   TL: tl@dm2.com / tl123" -ForegroundColor Gray
Write-Host ""
Write-Host "Press Ctrl+C to stop the server" -ForegroundColor Red
Write-Host ""

# Create a simple HTTP server
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()

Write-Host "✅ Server started successfully!" -ForegroundColor Green

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        $localPath = $request.Url.LocalPath
        if ($localPath -eq "/") { $localPath = "/index.html" }
        
        $filePath = Join-Path $path $localPath.TrimStart('/')
        
        if (Test-Path $filePath -PathType Leaf) {
            $content = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentLength64 = $content.Length
            
            # Set content type based on file extension
            $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
            switch ($extension) {
                ".html" { $response.ContentType = "text/html; charset=utf-8" }
                ".css" { $response.ContentType = "text/css; charset=utf-8" }
                ".js" { $response.ContentType = "application/javascript; charset=utf-8" }
                ".json" { $response.ContentType = "application/json; charset=utf-8" }
                ".png" { $response.ContentType = "image/png" }
                ".jpg" { $response.ContentType = "image/jpeg" }
                ".gif" { $response.ContentType = "image/gif" }
                default { $response.ContentType = "text/plain; charset=utf-8" }
            }
            
            $response.OutputStream.Write($content, 0, $content.Length)
            Write-Host "📄 Served: $localPath" -ForegroundColor DarkGray
        } else {
            $response.StatusCode = 404
            $errorMessage = "File not found: $localPath"
            $errorBytes = [System.Text.Encoding]::UTF8.GetBytes($errorMessage)
            $response.ContentLength64 = $errorBytes.Length
            $response.ContentType = "text/plain; charset=utf-8"
            $response.OutputStream.Write($errorBytes, 0, $errorBytes.Length)
            Write-Host "❌ 404: $localPath" -ForegroundColor Red
        }
        
        $response.Close()
    }
} catch {
    Write-Host "❌ Server error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    $listener.Stop()
    Write-Host "🛑 Server stopped." -ForegroundColor Yellow
}
