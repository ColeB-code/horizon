# Simple zero-dependency static file server for local testing.
# Run from this folder: .\serve.ps1
# Serves http://localhost:5500 to match the Azure AD redirect URI.

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:5500/")
$listener.Start()
Write-Host "Serving http://localhost:5500 - press Ctrl+C to stop"

while ($listener.IsListening) {
    $context = $listener.GetContext()
    $path = $context.Request.Url.LocalPath.TrimStart("/")
    if ([string]::IsNullOrEmpty($path)) { $path = "index.html" }
    $filePath = Join-Path $PSScriptRoot $path

    if (Test-Path $filePath -PathType Leaf) {
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        $ext = [System.IO.Path]::GetExtension($filePath)
        $contentType = switch ($ext) {
            ".html" { "text/html" }
            ".css"  { "text/css" }
            ".js"   { "application/javascript" }
            ".json" { "application/json" }
            ".png"  { "image/png" }
            default { "application/octet-stream" }
        }
        $context.Response.ContentType = $contentType
        $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
        $context.Response.StatusCode = 404
    }
    $context.Response.Close()
}
