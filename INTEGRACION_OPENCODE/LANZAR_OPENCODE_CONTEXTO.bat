@echo off
REM ============================================================
REM  LANZAR_OPENCODE_CONTEXTO.bat
REM  Abre OpenCode con la carpeta actual como workspace
REM  para continuar el trabajo cuando Claude Code se queda sin tokens
REM ============================================================

setlocal
set "OPENCODE=%LOCALAPPDATA%\Programs\@opencode-aidesktop\OpenCode.exe"

if not exist "%OPENCODE%" (
    echo [ERROR] No se encuentra OpenCode en:
    echo   %OPENCODE%
    pause
    exit /b 1
)

REM Si se pasa un argumento, es la carpeta a abrir. Si no, usa la carpeta actual.
if "%~1"=="" (
    set "WORKDIR=%CD%"
) else (
    set "WORKDIR=%~1"
)

echo Lanzando OpenCode en: %WORKDIR%
start "" "%OPENCODE%" "%WORKDIR%"
endlocal
exit /b 0
