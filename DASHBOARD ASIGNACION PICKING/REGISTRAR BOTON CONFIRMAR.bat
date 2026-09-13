@echo off
chcp 65001 >nul
title Registrar botones del dashboard de Picking
echo ==============================================================
echo   Registrando los botones del dashboard de Picking:
echo.
echo    1) "Confirmar planeacion 100%%"  -> respuesta a todos en Outlook
echo    2) "Generar correo (Outlook)"    -> abre el correo en el
echo                                        OUTLOOK DE ESCRITORIO (no web)
echo.
echo   Solo hay que correrlo UNA VEZ.
echo ==============================================================
echo.

set "PS1=%~dp0CONFIRMAR_PLANEACION.ps1"
set "PS2=%~dp0ABRIR_CORREO_OUTLOOK.ps1"

reg add "HKCU\Software\Classes\confirmarplan" /ve /d "URL:Confirmar Planeacion Picking" /f >nul
reg add "HKCU\Software\Classes\confirmarplan" /v "URL Protocol" /d "" /f >nul
reg add "HKCU\Software\Classes\confirmarplan\shell\open\command" /ve /d "powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"%PS1%\" \"%%1\"" /f >nul

reg add "HKCU\Software\Classes\pickingmail" /ve /d "URL:Abrir correo Picking en Outlook" /f >nul
reg add "HKCU\Software\Classes\pickingmail" /v "URL Protocol" /d "" /f >nul
reg add "HKCU\Software\Classes\pickingmail\shell\open\command" /ve /d "powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"%PS2%\" \"%%1\"" /f >nul

if %errorlevel%==0 (
  echo   LISTO. Los dos botones ya quedaron registrados.
  echo   - El boton verde responde a todos en Outlook.
  echo   - El boton "Generar correo" abrira el borrador en el
  echo     Outlook de escritorio con la imagen y el adjunto.
) else (
  echo   Hubo un problema al registrar. Corre este .bat de nuevo.
)
echo.
pause
