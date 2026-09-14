@echo off
title Dashboard de Cortes
cd /d "%~dp0"

echo.
echo  ============================================
echo    Dashboard de Cortes - Iniciando...
echo  ============================================
echo.
echo  [OK] Iniciando servidor...
echo.
echo  ============================================
echo    Abre tu navegador en:
echo    http://localhost:8050
echo.
echo    Para cerrar: Ctrl+C
echo  ============================================
echo.

py servidor.py

echo.
echo  Servidor cerrado.
pause
