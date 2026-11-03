# Setear el PAT de GitHub como variable de entorno User
# Lee el token directamente del portapapeles (no hay que pegarlo manualmente)

Write-Host ""
Write-Host "============================================"
Write-Host " SETEAR TOKEN GITHUB (desde portapapeles)"
Write-Host "============================================"
Write-Host ""

$token = (Get-Clipboard).Trim()

if ([string]::IsNullOrWhiteSpace($token)) {
    Write-Host "ERROR: tu portapapeles esta vacio." -ForegroundColor Red
    Write-Host ""
    Write-Host "Pasos:"
    Write-Host "  1) Andate a github.com/settings/tokens"
    Write-Host "  2) Click al icono de copiar al lado del token"
    Write-Host "  3) Volve a correr este script"
    Write-Host ""
    pause
    exit 1
}

if (-not $token.StartsWith("ghp_")) {
    Write-Host "ERROR: el portapapeles NO tiene un token de GitHub." -ForegroundColor Red
    Write-Host "Tiene esto: $($token.Substring(0, [Math]::Min(30, $token.Length)))..."
    Write-Host ""
    Write-Host "Un token empieza con 'ghp_'. Copialo bien y volve a correr."
    Write-Host ""
    pause
    exit 1
}

[Environment]::SetEnvironmentVariable("GITHUB_PAT", $token, "User")

$verif = [Environment]::GetEnvironmentVariable("GITHUB_PAT", "User")
if ($verif -eq $token) {
    Write-Host "============================================" -ForegroundColor Green
    Write-Host " OK - Token guardado como variable User" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Primeros caracteres: $($token.Substring(0,10))..."
    Write-Host "Longitud total: $($token.Length) chars"
    Write-Host ""
    Write-Host "LIMPIANDO portapapeles por seguridad..."
    Set-Clipboard -Value ""
    Write-Host "OK portapapeles limpio."
    Write-Host ""
    Write-Host "Listo. Podes cerrar esta ventana."
} else {
    Write-Host "ERROR: no se pudo guardar. Probá como administrador." -ForegroundColor Red
}

Write-Host ""
pause
