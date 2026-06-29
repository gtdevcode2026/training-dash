# Start dashboard + upload API (saves files under trainings/*/uploads/)
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$py = if (Test-Path ".venv\Scripts\python.exe") { ".\.venv\Scripts\python.exe" } else { "python" }

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    python -m venv .venv
    .\.venv\Scripts\pip install -r requirements.txt
    $py = ".\.venv\Scripts\python.exe"
}

& $py scripts\ensure_upload_folders.py
Write-Host ""
Write-Host "Open: http://127.0.0.1:8765/dashboard/index.html" -ForegroundColor Green
Write-Host "Press Ctrl+C to stop." -ForegroundColor DarkGray
Write-Host ""
& $py server\upload_server.py
