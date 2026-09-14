@echo off
REM Wrapper para ejecutar INSTALAR.ps1 sin cambiar politicas globales
powershell.exe -ExecutionPolicy Bypass -File "%~dp0INSTALAR.ps1"
pause
