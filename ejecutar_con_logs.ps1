# Ejecuta el Generador de Informe Electrico y muestra logs en tiempo real
# Uso: .\ejecutar_con_logs.ps1
# O desde la carpeta del proyecto: powershell -ExecutionPolicy Bypass -File ejecutar_con_logs.ps1

$ErrorActionPreference = "Stop"
$exeDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$destDir = Join-Path $exeDir "Generador_Informe_Electrico_EXE"
$exePath = Join-Path $destDir "GeneradorInformeElectrico_NUEVO.exe"
if (-not (Test-Path $exePath)) { $exePath = Join-Path $destDir "GeneradorInformeElectrico.exe" }
$logPath = Join-Path $destDir "GeneradorInformeElectrico.log"

if (-not (Test-Path $exePath)) {
    Write-Host "No se encuentra el ejecutable en: $exePath" -ForegroundColor Red
    Write-Host "Ejecute primero crear_exe_csharp.bat para generar el .exe" -ForegroundColor Yellow
    exit 1
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Generador Informe Electrico - Modo con Logs" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Log: $logPath" -ForegroundColor Gray
Write-Host "Iniciando aplicacion..." -ForegroundColor Green
Write-Host ""

# Ejecutar el .exe (abre la ventana de la app)
Start-Process -FilePath $exePath -WorkingDirectory $destDir

# Esperar a que exista el log
$timeout = 15
$elapsed = 0
while (-not (Test-Path $logPath) -and $elapsed -lt $timeout) {
    Start-Sleep -Seconds 1
    $elapsed++
}

if (Test-Path $logPath) {
    Write-Host "Mostrando logs (Ctrl+C para salir):" -ForegroundColor Yellow
    Write-Host ""
    Get-Content $logPath -Wait -Tail 20
} else {
    Write-Host "El log se creara al usar la aplicacion." -ForegroundColor Yellow
    Write-Host "Cuando termine, abra manualmente: $logPath" -ForegroundColor Gray
    Read-Host "Presione Enter para salir"
    if (Test-Path $logPath) { notepad $logPath }
}
