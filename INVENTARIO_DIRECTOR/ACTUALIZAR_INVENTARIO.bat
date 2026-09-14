@echo off
chcp 65001 >nul
title Inventario Director - Actualizando desde SQL
color 0A

echo.
echo  ====================================================
echo   ME LLEGA S.A. - INVENTARIO DIRECTOR
echo   Extraccion directa desde Korber WMS (SQL)
echo  ====================================================
echo.
echo  Conectando a 172.16.5.43 / IN_E ...
echo  Por favor espera (~30 segundos)...
echo.

py -3 "%~dp0EXTRAER_INVENTARIO.py"

if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo  ============================================
    echo   ERROR: No se pudo extraer los datos.
    echo.
    echo   Verifica que:
    echo   - Estes conectado a la red / VPN
    echo   - Python 3 este instalado
    echo   - pyodbc y openpyxl esten instalados
    echo  ============================================
    echo.
    pause
    exit /b 1
)

echo.
echo  Abriendo dashboard...
start "" "%~dp0Dashboard_Inventario.html"

echo.
echo  ====================================================
echo   Datos actualizados exitosamente!
echo  ====================================================
echo.
timeout /t 3 >nul
