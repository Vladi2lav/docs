# Минимальный статический сервер для запуска без Python (http://localhost:8080)
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$types = @{ '.html' = 'text/html; charset=utf-8'; '.js' = 'text/javascript; charset=utf-8'; '.css' = 'text/css; charset=utf-8'; '.json' = 'application/json; charset=utf-8' }
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add('http://localhost:8080/')
$listener.Start()
Start-Process 'http://localhost:8080'
while ($listener.IsListening) {
    $ctx = $listener.GetContext()
    $rel = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath.TrimStart('/'))
    if (-not $rel) { $rel = 'index.html' }
    $path = Join-Path $root $rel
    if ((Test-Path $path -PathType Leaf) -and $path.StartsWith($root)) {
        $ext = [IO.Path]::GetExtension($path)
        $ctx.Response.Headers.Add('Cache-Control', 'no-store')
        $ctx.Response.ContentType = $(if ($types[$ext]) { $types[$ext] } else { 'application/octet-stream' })
        $bytes = [IO.File]::ReadAllBytes($path)
        $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else { $ctx.Response.StatusCode = 404 }
    $ctx.Response.Close()
}
