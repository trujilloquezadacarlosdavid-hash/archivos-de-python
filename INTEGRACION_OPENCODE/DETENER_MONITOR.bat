@echo off
REM Detiene el monitor de rate-limit

setlocal EnableDelayedExpansion
set "found=0"

for /f "tokens=2 delims=," %%p in ('tasklist /FI "IMAGENAME eq powershell.exe" /FO CSV /NH /V 2^>NUL ^| findstr /I /C:"monitor_ratelimit"') do (
    set "pid=%%~p"
    echo Terminando PID !pid!...
    taskkill /PID !pid! /F >NUL 2>&1
    set "found=1"
)

if "%found%"=="0" (
    echo No hay monitor corriendo.
) else (
    echo Monitor detenido.
)
pause
exit /b 0
