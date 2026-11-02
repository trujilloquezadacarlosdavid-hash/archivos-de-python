# generar_fsd_historico.ps1
# Genera un heatmap HTML de ultimos 90 dias de cobertura FSD Mondelez.
# Combina SQL (sp_DataRecepciones) con los archivos FSD_MDLZ_*.html del folder de envios.

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outHtml   = Join-Path $scriptDir "FSD_HISTORICO.html"
$dirFSD    = "C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\Notas de envio Mondelez\_AGENTE\CORREOS FSD ENVIADOS"

$diasAtras = 90
$inicio = (Get-Date).Date.AddDays(-$diasAtras)
$fin    = (Get-Date).Date

Write-Host "Analizando desde $($inicio.ToString('yyyy-MM-dd')) hasta $($fin.ToString('yyyy-MM-dd'))..."

# ============================================================
# 1) Query SQL: citas Mondelez finalizadas en el rango
# ============================================================
$cs = "Server=172.16.5.43;Database=IN_E;Integrated Security=True;TrustServerCertificate=True"
$cn = New-Object System.Data.SqlClient.SqlConnection($cs); $cn.Open()
$cmd = $cn.CreateCommand()
$cmd.CommandText = "EXEC [LOGISTICA].[sp_DataRecepciones] @BodegaCodigo = '1110'"
$cmd.CommandTimeout = 120
$da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
$dt = New-Object System.Data.DataTable
[void]$da.Fill($dt); $cn.Close()

# Agrupar por cita
$citas = @{}
foreach ($r in $dt.Rows) {
    if ($r['Proveedor'] -notmatch 'MONDELEZ') { continue }
    if ($r['FechaCita'] -eq [DBNull]::Value) { continue }
    try { $f = [DateTime]::Parse($r['FechaCita'].ToString()).Date } catch { continue }
    if ($f -lt $inicio -or $f -gt $fin) { continue }
    $id = "$($r['IdCita'])"
    if (-not $citas.ContainsKey($id)) {
        $citas[$id] = @{ Fecha=$f; Comp="$($r['CitaCompletada'])"; Cancel="$($r['Cancelada'])" }
    }
}
$citasFinalizadas = @($citas.Values | Where-Object { $_.Comp -eq 'SI' })

# Agrupar por dia
$citasPorDia = @{}
foreach ($c in $citasFinalizadas) {
    $k = $c.Fecha.ToString('yyyy-MM-dd')
    if (-not $citasPorDia.ContainsKey($k)) { $citasPorDia[$k] = 0 }
    $citasPorDia[$k]++
}

# ============================================================
# 2) Leer archivos FSD_MDLZ_*.html
# ============================================================
$fsdPorDia = @{}
if (Test-Path $dirFSD) {
    Get-ChildItem $dirFSD -Filter "FSD_MDLZ_*.html" | ForEach-Object {
        if ($_.Name -match 'FSD_MDLZ_(\d{4}-\d{2}-\d{2})\.html') {
            try {
                $d = [DateTime]::ParseExact($matches[1], 'yyyy-MM-dd', $null)
                if ($d -ge $inicio -and $d -le $fin) {
                    $fsdPorDia[$matches[1]] = $true
                }
            } catch {}
        }
    }
}

# ============================================================
# 3) Construir dataset dia por dia
# ============================================================
$datos = @()
$cur = $inicio
while ($cur -le $fin) {
    $k = $cur.ToString('yyyy-MM-dd')
    $nCitas = if ($citasPorDia.ContainsKey($k)) { $citasPorDia[$k] } else { 0 }
    $hayFSD = $fsdPorDia.ContainsKey($k)

    $estado = if ($nCitas -eq 0) {
        'sin_cita'
    } elseif ($hayFSD) {
        'enviado'
    } else {
        'brecha'
    }

    $datos += [PSCustomObject]@{
        fecha = $k
        dow = [int]$cur.DayOfWeek  # 0=Dom, 6=Sab
        dia = $cur.Day
        mes = $cur.Month
        anio = $cur.Year
        semana = [math]::Floor(([datetime]$k - $inicio).Days / 7)
        citas = $nCitas
        fsd = $hayFSD
        estado = $estado
    }
    $cur = $cur.AddDays(1)
}

# ============================================================
# 4) KPIs
# ============================================================
$totalConCitas = @($datos | Where-Object { $_.citas -gt 0 }).Count
$totalEnviados = @($datos | Where-Object { $_.estado -eq 'enviado' }).Count
$totalBrechas  = @($datos | Where-Object { $_.estado -eq 'brecha' }).Count
$totalCitas    = ($datos | Measure-Object -Property citas -Sum).Sum
$citasConFSD   = ($datos | Where-Object { $_.estado -eq 'enviado' } | Measure-Object -Property citas -Sum).Sum
$citasSinFSD   = ($datos | Where-Object { $_.estado -eq 'brecha' } | Measure-Object -Property citas -Sum).Sum
if (-not $totalCitas)  { $totalCitas = 0 }
if (-not $citasConFSD) { $citasConFSD = 0 }
if (-not $citasSinFSD) { $citasSinFSD = 0 }
$pctCobertura  = if ($totalConCitas -gt 0) { [math]::Round(($totalEnviados / $totalConCitas) * 100, 1) } else { 0 }

# Analisis por dia de la semana
$semanales = @{}
0..6 | ForEach-Object { $semanales[$_] = @{ conCita=0; brecha=0 } }
foreach ($d in ($datos | Where-Object { $_.citas -gt 0 })) {
    $semanales[$d.dow].conCita++
    if ($d.estado -eq 'brecha') { $semanales[$d.dow].brecha++ }
}
$diasSem = @('Dom','Lun','Mar','Mie','Jue','Vie','Sab')

# ============================================================
# 5) Lista detallada de brechas
# ============================================================
$brechas = $datos | Where-Object { $_.estado -eq 'brecha' } | Sort-Object -Property fecha
$brechasHtml = foreach ($b in $brechas) {
    $fobj = [datetime]$b.fecha
    "<li><strong>$($b.fecha)</strong> ($($diasSem[$b.dow])) &mdash; $($b.citas) cita(s) sin FSD</li>"
}

# ============================================================
# 6) Heatmap: filas = semanas, columnas = dias de la semana
# ============================================================
$maxSemana = ($datos | Measure-Object -Property semana -Maximum).Maximum
$heatmapHtml = ""
for ($s = 0; $s -le $maxSemana; $s++) {
    $heatmapHtml += "<tr>"
    for ($dow = 0; $dow -le 6; $dow++) {
        $cell = $datos | Where-Object { $_.semana -eq $s -and $_.dow -eq $dow } | Select-Object -First 1
        if ($cell) {
            $cls = $cell.estado
            $tip = "$($cell.fecha): $($cell.citas) cita(s), FSD=$($cell.fsd)"
            $heatmapHtml += "<td class='cell $cls' title='$tip'>$($cell.dia)</td>"
        } else {
            $heatmapHtml += "<td class='cell empty'></td>"
        }
    }
    $heatmapHtml += "</tr>`n"
}

# ============================================================
# 7) HTML final
# ============================================================
$timestamp = Get-Date -Format 'dd-MMM-yyyy HH:mm'
$html = @"
<!DOCTYPE html>
<html lang="es-GT">
<head>
<meta charset="UTF-8">
<title>FSD Mondelez - Historico 90 dias</title>
<style>
:root {
  --arium-blue: #003A70;
  --arium-gold: #B8860B;
  --ok: #2E8B57;
  --warn: #F4C430;
  --err: #C8102E;
  --sin: #E0E0E0;
}
* { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', sans-serif; }
body { background: #F8F9FA; color: #333; padding: 1rem; }
header { background: var(--arium-blue); color: white; padding: 1rem 1.5rem; border-radius: 6px; }
h1 { font-size: 1.4rem; }
.timestamp { font-size: 0.85rem; opacity: 0.85; margin-top: 0.3rem; }
section { background: white; padding: 1rem 1.5rem; border-radius: 6px; margin-top: 1rem; }
h2 { color: var(--arium-blue); font-size: 1.1rem; margin-bottom: 0.6rem; border-bottom: 2px solid var(--arium-gold); padding-bottom: 0.3rem; }
.kpi-row { display: flex; gap: 1rem; }
.kpi { flex: 1; background: #FAFAFA; padding: 1rem; border-left: 5px solid var(--arium-blue); border-radius: 4px; }
.kpi-num { font-size: 2.2rem; font-weight: 700; color: var(--arium-blue); }
.kpi.ok .kpi-num { color: var(--ok); } .kpi.err .kpi-num { color: var(--err); }
.kpi.ok { border-color: var(--ok); } .kpi.err { border-color: var(--err); }
.kpi-lbl { font-size: 0.85rem; color: #666; text-transform: uppercase; margin-top: 0.3rem; }
.heatmap { border-collapse: separate; border-spacing: 3px; margin: 0 auto; }
.heatmap th { color: #666; font-size: 0.75rem; padding: 0.2rem; }
.cell { width: 32px; height: 32px; text-align: center; font-size: 0.75rem; border-radius: 4px; cursor: default; }
.cell.enviado { background: var(--ok); color: white; }
.cell.brecha { background: var(--err); color: white; font-weight: 700; }
.cell.sin_cita { background: var(--sin); color: #999; }
.cell.empty { background: transparent; }
.leyenda { text-align: center; margin-top: 1rem; font-size: 0.85rem; }
.leyenda span { display: inline-block; width: 14px; height: 14px; vertical-align: middle; margin: 0 4px 2px 10px; border-radius: 3px; }
ul { margin-left: 1.5rem; margin-top: 0.5rem; }
li { margin-bottom: 0.3rem; font-size: 0.9rem; }
table.weekly { width: 100%; border-collapse: collapse; margin-top: 0.5rem; }
table.weekly th, table.weekly td { padding: 0.4rem; text-align: center; border-bottom: 1px solid #eee; font-size: 0.85rem; }
table.weekly th { background: var(--arium-blue); color: white; }
table.weekly td.alto { background: #FDF2F2; color: var(--err); font-weight: 700; }
</style>
</head>
<body>
<header>
  <h1>FSD Mondelez - Historico ultimos $diasAtras dias</h1>
  <div class="timestamp">Generado: $timestamp &middot; Rango: $($inicio.ToString('yyyy-MM-dd')) a $($fin.ToString('yyyy-MM-dd'))</div>
</header>

<section>
  <h2>Resumen</h2>
  <div class="kpi-row">
    <div class="kpi"><div class="kpi-num">$totalCitas</div><div class="kpi-lbl">Citas Mondelez</div></div>
    <div class="kpi"><div class="kpi-num">$totalConCitas</div><div class="kpi-lbl">Dias con cita</div></div>
    <div class="kpi ok"><div class="kpi-num">$totalEnviados</div><div class="kpi-lbl">Dias con FSD enviado</div></div>
    <div class="kpi err"><div class="kpi-num">$totalBrechas</div><div class="kpi-lbl">Brechas (dias sin FSD)</div></div>
    <div class="kpi"><div class="kpi-num">$pctCobertura%</div><div class="kpi-lbl">Cobertura (dias)</div></div>
  </div>
</section>

<section>
  <h2>Heatmap diario</h2>
  <table class="heatmap">
    <thead><tr><th>Dom</th><th>Lun</th><th>Mar</th><th>Mie</th><th>Jue</th><th>Vie</th><th>Sab</th></tr></thead>
    <tbody>
$heatmapHtml
    </tbody>
  </table>
  <div class="leyenda">
    <span style="background:var(--ok)"></span> Enviado
    <span style="background:var(--err)"></span> Brecha
    <span style="background:var(--sin)"></span> Sin cita
  </div>
</section>

<section>
  <h2>Brechas por dia de la semana</h2>
  <table class="weekly">
    <thead><tr><th>Dia</th><th>Dias con cita</th><th>Brechas</th><th>% brechas</th></tr></thead>
    <tbody>
"@
foreach ($dow in 1..6 + @(0)) {
    $info = $semanales[$dow]
    $pct = if ($info.conCita -gt 0) { [math]::Round(($info.brecha / $info.conCita) * 100, 0) } else { 0 }
    $cls = if ($pct -gt 20) { 'alto' } else { '' }
    $html += "<tr><td>$($diasSem[$dow])</td><td>$($info.conCita)</td><td>$($info.brecha)</td><td class='$cls'>$pct%</td></tr>`n"
}
$html += @"
    </tbody>
  </table>
</section>

<section>
  <h2>Lista de brechas ($totalBrechas dias, $citasSinFSD citas)</h2>
  <ul>
$($brechasHtml -join "`n")
  </ul>
</section>

</body>
</html>
"@

$html | Out-File -FilePath $outHtml -Encoding UTF8
Write-Host "Generado: $outHtml"
Write-Host "Resumen: $totalEnviados/$totalConCitas dias con FSD ($pctCobertura%), $totalBrechas brechas"
