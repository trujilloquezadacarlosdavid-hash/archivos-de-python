@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Generando SALUD_AGENTES.html...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "generar_salud.ps1"
if errorlevel 1 (
  echo ERROR generando el dashboard.
  pause
  exit /b 1
)
echo Abriendo dashboard...
start "" "SALUD_AGENTES.html"
