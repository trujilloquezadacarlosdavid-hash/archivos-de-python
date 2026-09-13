@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ══════════════════════════════════════════════════
echo   CONTROL DE BLOQUES FEFO - Bodega 1110
echo ══════════════════════════════════════════════════
python control_bloques_fefo.py
if errorlevel 1 (
    echo.
    echo ERROR: Revisa la configuracion del script.
    pause
)
