# generar_index.ps1
# Escanea los dashboards HTML de Carlos y genera un index navegable.

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outHtml = Join-Path $scriptDir "index.html"

# ============================================================
# Catalogo de dashboards (ruta + grupo + descripcion + icon)
# Si una ruta no existe se marca como no disponible pero igual se lista.
# ============================================================
$dashboards = @(
    # INVENTARIO
    @{ ruta='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON\INVENTARIO_DIRECTOR\INFOGRAFIA_SOP_DIRECTOR.html'; nombre='Infografia SOP Director'; grupo='Inventario'; icono='&#x1F4CA;'; desc='Vista ejecutiva S&OP Kraft Heinz para direccion'},
    @{ ruta='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON\INVENTARIO_DIRECTOR\DETALLE_HEINZ.html'; nombre='Detalle Heinz nivel LP'; grupo='Inventario'; icono='&#x1F4D1;'; desc='Detalle a nivel LP de inventario Heinz'},

    # RECEPCIONES
    @{ ruta='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\DASHBOARD RECEPCIONES\index.html'; nombre='Dashboard Recepciones'; grupo='Recepciones'; icono='&#x1F69A;'; desc='Dashboard operativo de recepciones (SQL Korber en vivo)'},
    @{ ruta='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\Notas de envio Mondelez\_AGENTE\CORREOS FSD ENVIADOS'; nombre='FSD enviados (carpeta)'; grupo='Recepciones'; icono='&#x1F4E7;'; desc='Historial de FSDs generados (HTMLs por fecha)'},

    # NC MONDELEZ
    @{ ruta='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\RECLAMOS MONDELEZ 2025\DASHBOARD NC MONDELEZ.png'; nombre='Dashboard NC Mondelez'; grupo='NC Mondelez'; icono='&#x26A0;&#xFE0F;'; desc='NC Mondelez - vista grafica (PNG)'},

    # OPERACION DIARIA
    @{ ruta='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON\MEJORAS\SALUD_AGENTES\SALUD_AGENTES.html'; nombre='Salud de agentes'; grupo='Operacion'; icono='&#x2764;&#xFE0F;'; desc='Estado de los 20+ agentes y tareas programadas'},
    @{ ruta='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON\MEJORAS\FSD_HISTORICO\FSD_HISTORICO.html'; nombre='FSD Mondelez - 90 dias'; grupo='Operacion'; icono='&#x1F4C8;'; desc='Historico cobertura FSD ultimos 90 dias'}
)

# ============================================================
# Generar
# ============================================================
$timestamp = Get-Date -Format 'dd-MMM-yyyy HH:mm'

$dashboards | ForEach-Object {
    if (Test-Path $_.ruta) {
        $fi = Get-Item $_.ruta
        if ($fi.PSIsContainer) {
            $_.ultima = 'carpeta'
            $_.tamanio = (Get-ChildItem $_.ruta -Recurse -ErrorAction SilentlyContinue | Measure-Object).Count.ToString() + ' items'
        } else {
            $_.ultima  = $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
            $hrs = [math]::Round(((Get-Date) - $fi.LastWriteTime).TotalHours, 0)
            $_.frescura = if ($hrs -lt 24) { 'ok' } elseif ($hrs -lt 72) { 'warn' } else { 'err' }
            $_.tamanio  = "$([math]::Round($fi.Length / 1KB, 0)) KB"
        }
        $_.disponible = $true
    } else {
        $_.disponible = $false
        $_.ultima = 'NO ENCONTRADO'
        $_.frescura = 'err'
    }
}

# Agrupar por grupo
$grupos = $dashboards | Group-Object -Property grupo

$tarjetas = foreach ($g in $grupos) {
    $cards = foreach ($d in $g.Group) {
        $enlace = if ($d.disponible) {
            "file:///$($d.ruta -replace '\\','/')"
        } else { '#' }
        $disabledClass = if ($d.disponible) { '' } else { 'disabled' }
        $frescura = if ($d.frescura) { $d.frescura } else { 'ok' }
        $tamanio = if ($d.tamanio) { $d.tamanio } else { '' }
@"
      <a href="$enlace" class="card $disabledClass $frescura" target="_blank">
        <div class="icon">$($d.icono)</div>
        <div class="body">
          <div class="titulo">$($d.nombre)</div>
          <div class="descripcion">$($d.desc)</div>
          <div class="meta">Ultima: $($d.ultima) &middot; $tamanio</div>
        </div>
      </a>
"@
    }
@"
    <section class="grupo">
      <h2>$($g.Name) <span class="count">($($g.Count))</span></h2>
      <div class="cards">
$($cards -join "`n")
      </div>
    </section>
"@
}

$html = @"
<!DOCTYPE html>
<html lang="es-GT">
<head>
<meta charset="UTF-8">
<title>Dashboards - Bodega 1110</title>
<style>
:root {
  --arium-blue: #003A70;
  --arium-gold: #B8860B;
  --ok: #2E8B57;
  --warn: #F4C430;
  --err: #C8102E;
}
* { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', sans-serif; }
body { background: #F8F9FA; color: #333; padding: 1rem; }
header { background: var(--arium-blue); color: white; padding: 1.5rem; border-radius: 6px; margin-bottom: 1.5rem; }
h1 { font-size: 1.5rem; }
.timestamp { font-size: 0.85rem; opacity: 0.85; margin-top: 0.3rem; }
section.grupo { margin-bottom: 2rem; }
h2 { color: var(--arium-blue); font-size: 1.2rem; margin-bottom: 1rem; border-bottom: 2px solid var(--arium-gold); padding-bottom: 0.4rem; }
.count { color: #999; font-size: 0.9rem; font-weight: normal; }
.cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 1rem; }
a.card { display: flex; background: white; padding: 1rem; border-radius: 6px; text-decoration: none; color: #333; border-left: 5px solid var(--arium-blue); transition: transform 0.1s, box-shadow 0.1s; }
a.card:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
a.card.disabled { opacity: 0.4; pointer-events: none; border-left-color: #ccc; }
a.card.ok { border-left-color: var(--ok); }
a.card.warn { border-left-color: var(--warn); }
a.card.err { border-left-color: var(--err); }
.icon { font-size: 2rem; margin-right: 1rem; align-self: center; }
.body { flex: 1; }
.titulo { font-weight: 700; color: var(--arium-blue); margin-bottom: 0.3rem; }
.descripcion { font-size: 0.85rem; color: #666; margin-bottom: 0.4rem; }
.meta { font-size: 0.75rem; color: #999; font-style: italic; }
</style>
</head>
<body>
<header>
  <h1>&#x1F4C2; Portal de Dashboards - Bodega 1110</h1>
  <div class="timestamp">Generado: $timestamp &middot; $($dashboards.Count) dashboards registrados</div>
</header>

$($tarjetas -join "`n")

</body>
</html>
"@

$html | Out-File -FilePath $outHtml -Encoding UTF8
Write-Host "Generado: $outHtml"
