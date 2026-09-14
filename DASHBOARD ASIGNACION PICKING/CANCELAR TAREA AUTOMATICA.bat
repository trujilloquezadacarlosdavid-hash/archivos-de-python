@echo off
chcp 65001 >nul
title Cancelar tarea automatica - Dashboard Picking

echo.
echo  Cancelando la tarea de vigilancia automatica...
echo.

schtasks /delete /tn "MeLlega - Vigilar Correo Planner Picking" /f

if %ERRORLEVEL% EQU 0 (
  echo  Tarea cancelada. El agente ya no vigilara la bandeja.
) else (
  echo  La tarea no estaba registrada (o ya fue cancelada).
)

echo.
pause
