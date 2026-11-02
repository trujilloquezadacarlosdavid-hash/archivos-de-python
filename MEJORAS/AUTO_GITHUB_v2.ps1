# AUTO_GITHUB_v2.ps1
# Version limpia: lee el PAT de la variable de entorno GITHUB_PAT (no hardcoded).
# Ver MEJORAS\INSTRUCCIONES_PAT.md para el setup inicial.

Set-Location "C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON"

$pat = [Environment]::GetEnvironmentVariable("GITHUB_PAT", "User")
if ([string]::IsNullOrWhiteSpace($pat)) {
    Write-Host "ERROR: variable de entorno GITHUB_PAT no configurada."
    Write-Host "Ver MEJORAS\INSTRUCCIONES_PAT.md para configurarla."
    exit 1
}

$cambios = git status --porcelain
if ($cambios) {
    $fecha = Get-Date -Format "yyyy-MM-dd HH:mm"
    git add -A
    git commit -m "Auto-sync $fecha"
    $remoteUrl = "https://trujilloquezadacarlosdavid-hash:$pat@github.com/trujilloquezadacarlosdavid-hash/archivos-de-python.git"
    git push $remoteUrl main 2>&1
}
