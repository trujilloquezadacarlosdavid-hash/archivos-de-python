@echo off
chcp 65001 >nul
title Registrar tarea automatica - Dashboard Picking

echo.
echo  =====================================================
echo    ME LLEGA S.A.  -  AGENTE PICKING (UNA VEZ AL DIA)
echo  =====================================================
echo.
echo  Registrando tarea que arranca al iniciar sesion.
echo  Espera en segundo plano a que llegue el correo de
echo  Planner, lo procesa UNA vez y se cierra.
echo  NO repite cada X minutos.
echo.

schtasks /create ^
  /tn "MeLlega - Vigilar Correo Planner Picking" ^
  /tr "powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"%~dp0VIGILAR_CORREO_PLANNER.ps1\"" ^
  /sc ONLOGON ^
  /ru "%USERNAME%" ^
  /f

if %ERRORLEVEL% EQU 0 (
  echo.
  echo  LISTO. La tarea quedo registrada.
  echo.
  echo  A partir de manana:
  echo    - Al encender tu PC arranca en segundo plano.
  echo    - Espera silenciosamente a que llegue el Estatus de Picking.
  echo    - Cuando lo detecta, procesa todo y abre el dashboard.
  echo    - Te avisa por globo + correo.
  echo    - Despues se cierra. Un solo analisis al dia.
  echo    - Outlook debe estar abierto para que funcione.
  echo.
  echo  Para cancelar: ejecuta CANCELAR TAREA AUTOMATICA.bat
) else (
  echo.
  echo  ERROR al registrar la tarea. Intenta correr este .bat
  echo  como Administrador (clic derecho - Ejecutar como admin^).
)

echo.
pause
