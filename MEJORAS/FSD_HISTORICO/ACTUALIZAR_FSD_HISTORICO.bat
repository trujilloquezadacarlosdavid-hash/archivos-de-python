@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Generando FSD_HISTORICO.html (90 dias)...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "generar_fsd_historico.ps1"
if errorlevel 1 (
  echo ERROR generando el reporte.
  pause
  exit /b 1
)
echo Abriendo dashboard...
start "" "FSD_HISTORICO.html"
