@echo off
rem Refresco SILENCIOSO (sin abrir navegador) - para la tarea programada por hora
rem Regenera INFOGRAFIA_SOP_DIRECTOR.html y DETALLE_HEINZ.html con datos frescos de Korber.
chcp 65001 >nul
cd /d "%~dp0"
py -3 "%~dp0EXTRAER_SOP_HEINZ.py"     >> "%~dp0refresco_log.txt" 2>&1
py -3 "%~dp0EXTRAER_DETALLE_HEINZ.py" >> "%~dp0refresco_log.txt" 2>&1
rem Mantener sincronizada la copia en Proyecto Heinz
set "PH=%USERPROFILE%\OneDrive - Me Llega, S.A\Escritorio\TRUJILLO.C\Proyecto Heinz"
if exist "%PH%" (
  copy /Y "%~dp0INFOGRAFIA_SOP_DIRECTOR.html" "%PH%\" >nul
  copy /Y "%~dp0DETALLE_HEINZ.html" "%PH%\" >nul
)
echo [%date% %time%] Refresco completado >> "%~dp0refresco_log.txt"
exit /b 0
