@echo off
REM Lanza el monitor de rate-limit en background (ventana oculta)
REM Solo mantiene UNA instancia corriendo

setlocal
set "SCRIPT=%~dp0monitor_ratelimit.ps1"

REM Chequear si ya hay un monitor corriendo
tasklist /FI "IMAGENAME eq powershell.exe" /FO CSV /V 2>NUL | findstr /I /C:"monitor_ratelimit" >NUL
if %ERRORLEVEL% EQU 0 (
    echo Monitor ya esta corriendo.
    pause
    exit /b 0
)

echo Lanzando monitor en background...
start "" /B powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT%"
echo OK. Log: %LOCALAPPDATA%\opencode_failover.log
timeout /t 3 >NUL
exit /b 0
