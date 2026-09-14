Set-Location "C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON"

$cambios = git status --porcelain
if ($cambios) {
    $fecha = Get-Date -Format "yyyy-MM-dd HH:mm"
    git add -A
    git commit -m "Auto-sync $fecha"
    git push https://trujilloquezadacarlosdavid-hash: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@github.com/trujilloquezadacarlosdavid-hash/archivos-de-python.git main 2>&1
}
