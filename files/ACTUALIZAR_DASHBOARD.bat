@echo off
chcp 65001 >nul
title Actualizando Dashboard Me Llega S.A.
color 0A

echo.
echo  ================================================
echo   ME LLEGA S.A. - ACTUALIZADOR DE DASHBOARD
echo  ================================================
echo.
echo  Leyendo archivo Excel...
echo  Por favor espera...
echo.

py -3 "%~dp0ACTUALIZAR_DASHBOARD.py"

if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo  ERROR: Algo salió mal. Verifica que:
    echo  - Python 3 esté instalado
    echo  - El archivo Excel esté en esta carpeta
    echo  - El archivo Excel NO esté abierto en Excel
    echo.
    pause
    exit /b 1
)

echo.
echo  Abriendo dashboard en el navegador...
echo.
start "" "%~dp0dashboard_equipo.html"

echo  ================================================
echo   Dashboard actualizado exitosamente!
echo  ================================================
echo.
timeout /t 3 >nul
