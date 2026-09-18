@echo off
title CONTEO RUTAS 1111 - Actualizando desde SQL
cd /d "%~dp0"
echo.
echo ============================================================
echo   CONTEO RUTAS PREPARADAS - BODEGA 1111
echo   Extrayendo desde 172.16.5.43/IN_E ...
echo ============================================================
echo.
py EXTRAER_CONTEO.py
if errorlevel 1 (
  echo.
  echo *** ERROR al extraer ***
  pause
  exit /b 1
)
echo.
echo Abriendo dashboard ...
start "" "Conteo_Rutas_1111.html"
timeout /t 2 >nul
