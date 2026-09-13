@echo off
REM Lee la key desde el portapapeles y la guarda como GROQ_API_KEY
REM Requisito: tener la key copiada (Ctrl+C) desde console.groq.com/keys

echo.
echo ============================================
echo   Guardar GROQ_API_KEY desde portapapeles
echo ============================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$k = (Get-Clipboard).Trim();" ^
  "if ([string]::IsNullOrWhiteSpace($k)) { Write-Host 'ERROR: portapapeles vacio' -ForegroundColor Red; exit 1 };" ^
  "if ($k -notmatch '^gsk_') { Write-Host ('ERROR: no parece key de Groq (empieza con: ' + $k.Substring(0,[Math]::Min(10,$k.Length)) + '...)') -ForegroundColor Red; exit 1 };" ^
  "if ($k.Length -lt 40) { Write-Host ('ERROR: key muy corta (' + $k.Length + ' chars)') -ForegroundColor Red; exit 1 };" ^
  "[Environment]::SetEnvironmentVariable('GROQ_API_KEY', $k, 'User');" ^
  "Write-Host ('OK - Key guardada. Longitud: ' + $k.Length + ' chars. Prefijo: ' + $k.Substring(0,8) + '... Sufijo: ...' + $k.Substring($k.Length-4)) -ForegroundColor Green"

echo.
echo Presiona una tecla para cerrar...
pause >NUL
