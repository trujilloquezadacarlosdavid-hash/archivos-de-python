@echo off
title Diagnostico Dashboard
echo.
echo  ============================================
echo    DIAGNOSTICO DEL SISTEMA
echo  ============================================
echo.

echo [1] Verificando Python...
python --version 2>&1
if %errorlevel% neq 0 (
    echo     ERROR: python no encontrado
    py --version 2>&1
)

echo.
echo [2] Verificando pandas...
python -c "import pandas; print('pandas OK - version', pandas.__version__)" 2>&1

echo.
echo [3] Verificando openpyxl...
python -c "import openpyxl; print('openpyxl OK - version', openpyxl.__version__)" 2>&1

echo.
echo [4] Verificando archivos del sistema...
if exist "servidor.py" (echo     servidor.py: OK) else (echo     servidor.py: NO ENCONTRADO)
if exist "dashboard_cortes.html" (echo     dashboard_cortes.html: OK) else (echo     dashboard_cortes.html: NO ENCONTRADO)

echo.
echo [5] Instalando dependencias faltantes...
python -m pip install pandas openpyxl --quiet 2>&1
echo     Instalacion completada.

echo.
echo  ============================================
echo    Si todo dice OK, ejecuta INICIAR.bat
echo  ============================================
echo.
pause
