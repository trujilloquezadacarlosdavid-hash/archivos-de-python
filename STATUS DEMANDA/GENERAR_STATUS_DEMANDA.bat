@echo off
chcp 65001 >nul
title STATUS DEMANDA - Me Llega S.A.
color 0B

echo.
echo  ================================================
echo   ME LLEGA S.A. - STATUS DE INGRESO (DEMANDA)
echo  ================================================
echo.
echo  Leyendo "Archivo Mario Perez avance descargas.xlsx"...
echo  Calculando estatus por cita...
echo.

py -3 "%~dp0generar_status_demanda.py"

if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo  ERROR: Algo salio mal. Verifica que:
    echo  - Python 3 este instalado ^(py -3^)
    echo  - Exista "Archivo Mario Perez avance descargas.xlsx"
    echo    en el Escritorio o en esta carpeta
    echo.
    pause
    exit /b 1
)

echo.
echo  ================================================
echo   Status generado. Abriendo tablero...
echo  ================================================
echo.
timeout /t 2 >nul
