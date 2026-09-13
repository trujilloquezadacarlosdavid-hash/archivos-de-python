@echo off
chcp 65001 >nul
title Buscar pedidos de correo - Asignacion Picking - Me Llega S.A.
color 0A

echo.
echo  ====================================================
echo    ME LLEGA S.A.  -  PEDIDOS POR CORREO (Planeacion)
echo  ====================================================
echo.
echo  Leyendo los correos de hoy en Outlook y buscando los
echo  numeros de orden pedidos (TR / ST / OV)...
echo  (Outlook debe estar ABIERTO. Si sale aviso de Outlook
echo   pidiendo permiso, acepta "Permitir".)
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0BUSCAR_CORREOS_PEDIDOS.ps1" %*

if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo  ERROR: No se pudieron leer los correos.
    echo  Verifica que Outlook este abierto y acepta el permiso.
    echo.
    pause
    exit /b 1
)

echo.
echo  Abriendo el dashboard...
start "" "%~dp0Dashboard Asignacion Picking.html"

echo.
echo  ====================================================
echo    Listo. En el dashboard arrastra el Excel del dia;
echo    en "Verificar pedidos por correo" ya apareceran los
echo    pedidos de hoy y te marcara lo que NO se agrego.
echo  ====================================================
timeout /t 5 >nul
