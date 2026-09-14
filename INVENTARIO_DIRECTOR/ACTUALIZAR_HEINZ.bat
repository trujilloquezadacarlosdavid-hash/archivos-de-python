@echo off
chcp 65001 >nul
title S^&OP Kraft Heinz - Actualizando desde SQL
color 0A

echo.
echo  ========================================================
echo   ME LLEGA S.A. - S^&OP KRAFT HEINZ
echo   Infografia (director) + Detalle (operacion)
echo   Extraccion en vivo desde Korber WMS
echo  ========================================================
echo.
echo  Conectando a 172.16.5.43 / IN_E ...
echo  Por favor espera (~40 segundos)...
echo.

echo  [1/2] Generando infografia del director...
py -3 "%~dp0EXTRAER_SOP_HEINZ.py"
if %ERRORLEVEL% NEQ 0 goto ERROR

echo.
echo  [2/2] Generando detalle a nivel LP...
py -3 "%~dp0EXTRAER_DETALLE_HEINZ.py"
if %ERRORLEVEL% NEQ 0 goto ERROR

rem Mantener sincronizada la copia en Proyecto Heinz
set "PH=%USERPROFILE%\OneDrive - Me Llega, S.A\Escritorio\TRUJILLO.C\Proyecto Heinz"
if exist "%PH%" (
  copy /Y "%~dp0INFOGRAFIA_SOP_DIRECTOR.html" "%PH%\" >nul
  copy /Y "%~dp0DETALLE_HEINZ.html" "%PH%\" >nul
)

echo.
echo  Abriendo tableros...
start "" "%~dp0INFOGRAFIA_SOP_DIRECTOR.html"
start "" "%~dp0DETALLE_HEINZ.html"

echo.
echo  ========================================================
echo   Actualizado exitosamente! Ambos tableros cuadran.
echo  ========================================================
ping -n 4 127.0.0.1 >nul
exit /b 0

:ERROR
color 0C
echo.
echo  ============================================
echo   ERROR: No se pudo extraer los datos.
echo   Verifica que:
echo   - Estes conectado a la red / VPN
echo   - Python 3 este instalado (py -3)
echo   - pyodbc este instalado
echo  ============================================
echo.
pause
exit /b 1
