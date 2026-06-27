$ErrorActionPreference = "Stop"
$backend = Join-Path $PSScriptRoot "backend"
$frontend = Join-Path $PSScriptRoot "frontend"
$backendProcess = $null

try {
    Write-Host "Iniciando API en http://localhost:8000 ..." -ForegroundColor DarkCyan
    $backendProcess = Start-Process `
        -FilePath python `
        -ArgumentList @("-m", "uvicorn", "main:app", "--host", "127.0.0.1", "--port", "8000") `
        -WorkingDirectory $backend `
        -WindowStyle Hidden `
        -PassThru

    $apiReady = $false
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        if ($backendProcess.HasExited) {
            throw "La API terminó durante el arranque. Ejecuta 'python -m uvicorn main:app --port 8000' desde backend para ver el error."
        }
        try {
            $health = Invoke-RestMethod -Uri "http://127.0.0.1:8000/health" -TimeoutSec 1
            if ($health.status -eq "ok") {
                $apiReady = $true
                break
            }
        }
        catch {
            Start-Sleep -Milliseconds 250
        }
    }
    if (-not $apiReady) {
        throw "La API no respondió en http://localhost:8000."
    }

    Write-Host "DataLex 1581 disponible en http://localhost:5173" -ForegroundColor Cyan
    Write-Host "Presiona Ctrl+C para detener ambos servicios." -ForegroundColor DarkGray
    python -m http.server 5173 --bind 127.0.0.1 --directory $frontend
}
finally {
    if ($backendProcess -and -not $backendProcess.HasExited) {
        Stop-Process -Id $backendProcess.Id -Force
    }
}
