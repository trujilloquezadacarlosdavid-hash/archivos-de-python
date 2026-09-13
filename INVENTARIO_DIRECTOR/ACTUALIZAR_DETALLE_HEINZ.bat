@echo off
chcp 65001 >nul
title Detalle Kraft Heinz - Actualizando desde SQL
color 0A

echo.
echo  ====================================================
echo   ME LLEGA S.A. - DETALLE KRAFT HEINZ
echo   Vencimiento y Planeacion a nivel LP
echo  ====================================================
echo.
echo  Conectando a Korber (172.16.5.43 / IN_E)...
echo  Por favor espera (~20 segundos)...
echo.

py -3 "%~dp0EXTRAER_DETALLE_HEINZ.py"

if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo  ============================================
    echo   ERROR: No se pudo extraer los datos.
    echo   Verifica la red/VPN y que Python este instalado.
    echo  ============================================
    echo.
    pause
    exit /b 1
)

echo.
echo  Abriendo detalle...
start "" "%~dp0DETALLE_HEINZ.html"

echo.
echo  ====================================================
echo   Detalle actualizado exitosamente!
echo  ====================================================
timeout /t 3 >nul
