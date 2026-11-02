# generar_salud.ps1
# Lee los logs y tareas programadas de Carlos, genera SALUD_AGENTES.html.
# Corrida manual: .\generar_salud.ps1
# Programable c/30min via tarea (ver INSTALAR_SALUD.bat)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outHtml   = Join-Path $scriptDir "SALUD_AGENTES.html"

# ============================================================
# 1) Catalogo de agentes (nombre, descripcion, log, tarea programada esperada)
# ============================================================
$catalogo = @(
    @{ nombre='Agente FSD Mondelez'; descripcion='Envia FSD + Fill Rate a Mondelez (diario)'; log='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\Notas de envio Mondelez\_AGENTE\_log.txt'; tarea='FSD Mondelez Auto'; frecuencia='diaria' },
    @{ nombre='Agente FSD Casa'; descripcion='FSD para casas no-Mondelez'; log='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\Notas de envio Mondelez\_AGENTE\_log.txt'; tarea='FSD Mondelez Auto'; frecuencia='diaria' },
    @{ nombre='Auditoria FSD Diaria'; descripcion='Compara calendario vs FSDs enviados del dia'; log='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\Notas de envio Mondelez\_AGENTE\_audit_fsd.log'; tarea='Auditoria FSD Diaria'; frecuencia='diaria 12:15 PM' },
    @{ nombre='Monitor FSD'; descripcion='Monitorea estado FSD c/15min'; log='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\Notas de envio Mondelez\_AGENTE\_monitor_fsd.log'; tarea='Monitor FSD Diario'; frecuencia='diaria 3:30 AM' },
    @{ nombre='Archivar Mondelez (David)'; descripcion='Archiva PDFs Reporte de descarga MONDELEZ'; log=''; tarea='Archivar Reportes Mondelez'; frecuencia='c/10min' },
    @{ nombre='Archivar Notas Envio Mondelez'; descripcion='Archiva Notas de envio Mondelez'; log=''; tarea='Archivar Notas Envio Mondelez'; frecuencia='c/15min 7AM-6PM' },
    @{ nombre='Carta Aceptacion'; descripcion='Borrador diario carta aceptacion'; log=''; tarea='Agente Carta Aceptacion'; frecuencia='diaria 6PM' },
    @{ nombre='Dashboard Recepciones (SQL)'; descripcion='Actualiza dashboard desde SQL Korber'; log=''; tarea='Dashboard Recepciones - Actualizar desde SQL'; frecuencia='manual + servicio' },
    @{ nombre='MTTO Montacargas Lector'; descripcion='Lee grupo WA MTTO + envia correo diario'; log='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\MTTO Montacargas\first_run.log'; tarea='MTTO Montacargas - Lector'; frecuencia='c/hora 6AM-6PM lun-sab' },
    @{ nombre='Monitor Fechas 1110'; descripcion='Revisa recepciones y alerta WA si fecha incorrecta'; log=''; tarea='Monitor Fechas 1110'; frecuencia='c/15min 3AM-12PM' },
    @{ nombre='Detector Vencidos Pasillos'; descripcion='Detecta LPs vencidos en pasillos (3 corridas)'; log=''; tarea='Detector Vencidos Pasillos 0400'; frecuencia='4,7,11 AM lun-sab' },
    @{ nombre='Detector Exceso LPs'; descripcion='Detecta posiciones B-F con exceso de LPs'; log=''; tarea='detector-exceso-lps'; frecuencia='c/15min 3AM-12PM' },
    @{ nombre='Monitor Tiempos Muertos'; descripcion='Detecta montacarguistas inactivos >10min'; log='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\TRUJILLO.C\PROYECTO MEDICION EMPLEADOS\_tiempos_muertos_log.txt'; tarea='Monitor Tiempos Muertos - Montacarguistas'; frecuencia='c/10min' },
    @{ nombre='Castro Status WA'; descripcion='Status descargas al grupo WA'; log=''; tarea='Castro STATUS - Respaldo 0305'; frecuencia='multiples horarios' },
    @{ nombre='Reporte Cierre Ops SAP vs KORBER'; descripcion='Cuadre diario SAP vs Korber'; log=''; tarea='Reporte Cierre Ops - SAP vs KORBER'; frecuencia='diaria 3AM' },
    @{ nombre='Avance Mensual Equipo'; descripcion='Mensaje avance a Recepciones 2026 sabados'; log=''; tarea='Avance Mensual Equipo - Sabados'; frecuencia='sabado 3AM' },
    @{ nombre='AutoGitHub Python'; descripcion='Sync del repo cada 2h'; log=''; tarea='AutoGitHub_Python'; frecuencia='c/2h' },
    @{ nombre='Dashboard Heinz (SOP + Detalle)'; descripcion='Dashboards de inventario para Director'; log='C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON\INVENTARIO_DIRECTOR\refresco_log.txt'; tarea=''; frecuencia='manual (ACTUALIZAR_HEINZ.bat)' }
)

# ============================================================
# 2) Leer tareas programadas y mapear por nombre
# ============================================================
$tareas = @{}
try {
    # schtasks /FO CSV genera CSV MALFORMADO (comillas sin escapar).
    # Usamos /FO LIST que es mas robusto: cada campo en su propia linea.
    $raw = & "schtasks.exe" "/Query" "/FO" "LIST" "/V" 2>$null
    if (-not $raw) { throw "schtasks sin output" }

    # Helper para buscar valor por patron ASCII (evita problemas de encoding con
    # caracteres como "ó", "ú", "á" en los nombres de campo).
    function Get-CampoByPattern {
        param($campos, [string]$pattern)
        foreach ($k in $campos.Keys) {
            if ($k -match $pattern) { return $campos[$k] }
        }
        return ''
    }

    $bloques = ($raw -join "`n") -split "(?:`r?`n){2,}"
    foreach ($b in $bloques) {
        $campos = @{}
        foreach ($linea in ($b -split "`r?`n")) {
            if ($linea -match '^\s*([^:]+?):\s*(.*)$') {
                $k = $matches[1].Trim()
                $v = $matches[2].Trim()
                if (-not $campos.ContainsKey($k)) { $campos[$k] = $v }
            }
        }
        $nombreRaw = Get-CampoByPattern $campos 'Nombre de tarea'
        $nombre = ($nombreRaw -replace '^\\','').Trim()
        if ($nombre -eq '' -or $nombre -match '^Microsoft') { continue }

        $datos = @{
            proxima     = Get-CampoByPattern $campos 'Hora pr.xima ejecuci.n'
            ultima      = Get-CampoByPattern $campos '.ltimo tiempo de ejecuci.n'
            resultado   = Get-CampoByPattern $campos '.ltimo resultado'
            estado      = Get-CampoByPattern $campos 'Estado de tarea programada'
            horaInicio  = Get-CampoByPattern $campos 'Hora de inicio'
            fechaInicio = Get-CampoByPattern $campos 'Fecha de inicio'
        }

        if ($tareas.ContainsKey($nombre)) {
            foreach ($k in @('proxima','ultima','resultado','estado')) {
                if ($datos[$k] -ne '' -and $datos[$k] -ne 'N/A' -and ($tareas[$nombre][$k] -eq '' -or $tareas[$nombre][$k] -eq 'N/A')) {
                    $tareas[$nombre][$k] = $datos[$k]
                }
            }
            $tareas[$nombre].disparadores++
            if ($datos.horaInicio -ne '' -and $tareas[$nombre].horasInicio -notcontains $datos.horaInicio) {
                $tareas[$nombre].horasInicio += $datos.horaInicio
            }
        } else {
            $datos.disparadores = 1
            $datos.horasInicio = @($datos.horaInicio)
            $tareas[$nombre] = $datos
        }
    }
} catch {
    Write-Host "ERROR leyendo schtasks: $($_.Exception.Message)"
}

# ============================================================
# 3) Evaluar salud de cada agente
# ============================================================
function Get-SaludAgente {
    param($agente, $tareas)

    $info = @{
        nombre = $agente.nombre
        descripcion = $agente.descripcion
        frecuencia = $agente.frecuencia
        proxima = ''
        ultima = ''
        resultado = ''
        estado = 'desconocido'  # ok | warn | error | desconocido
        detalle = ''
        logSize = 0
        logMtime = ''
    }

    # Datos de schtasks
    if ($agente.tarea -and $tareas.ContainsKey($agente.tarea)) {
        $t = $tareas[$agente.tarea]
        $info.proxima = $t.proxima
        $info.ultima  = $t.ultima
        $info.resultado = $t.resultado

        if ($t.resultado -eq '0') {
            $info.estado = 'ok'
            $info.detalle = 'Ultima corrida exitosa'
        } elseif ($t.resultado -match '^-21\d{8}') {
            $info.estado = 'error'
            $info.detalle = "Error $($t.resultado) (codigo Windows)"
        } elseif ($t.resultado -eq '1' -or $t.resultado -eq '267011' -or $t.resultado -eq '267014') {
            $info.estado = 'warn'
            $info.detalle = "Codigo $($t.resultado) (no fatal, revisar)"
        } elseif ($t.resultado) {
            $info.estado = 'warn'
            $info.detalle = "Codigo $($t.resultado)"
        }
    } elseif ($agente.tarea) {
        $info.estado = 'warn'
        $info.detalle = "Tarea '$($agente.tarea)' no encontrada en schtasks"
    }

    # Info del log
    if ($agente.log -and (Test-Path $agente.log)) {
        $fi = Get-Item $agente.log
        $info.logSize  = [math]::Round($fi.Length / 1KB, 1)
        $info.logMtime = $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm')

        # Si log no se tocó en >24h y la tarea es diaria/más frecuente, bajar a warn
        $hrsSinActividad = ((Get-Date) - $fi.LastWriteTime).TotalHours
        if ($hrsSinActividad -gt 48 -and $info.estado -eq 'ok') {
            $info.estado = 'warn'
            $info.detalle += " | Log sin updates hace $([math]::Round($hrsSinActividad,0))h"
        }
    } elseif ($agente.log) {
        if ($info.estado -eq 'ok') { $info.estado = 'warn' }
        $info.detalle += " | Log esperado NO existe: $($agente.log)"
    }

    # Forzamos retorno como PSCustomObject (si retornamos hashtable PowerShell
    # lo enumera como KeyValuePair en pipeline, rompiendo el filtrado).
    return [PSCustomObject]$info
}

$resultados = @(foreach ($a in $catalogo) { Get-SaludAgente -agente $a -tareas $tareas })

# ============================================================
# 4) Detectar duplicados de tareas (bug conocido Castro)
# Un duplicado real = varios disparadores (bloques) con diferentes horas de inicio
# para la misma tarea (schtasks los agrupa como un sola entry si son el mismo trigger).
# Aqui lo detectamos como tareas con >1 disparador con horas distintas.
# ============================================================
$duplicadas = @{}
foreach ($nombre in $tareas.Keys) {
    $t = $tareas[$nombre]
    $horasUnicas = @($t.horasInicio | Where-Object { $_ -ne '' -and $_ -ne 'N/A' } | Select-Object -Unique)
    if ($horasUnicas.Count -gt 1) {
        $duplicadas[$nombre] = $horasUnicas.Count
    }
}
$dupList = $duplicadas.GetEnumerator() | Sort-Object -Property Value -Descending

# ============================================================
# 5) Construir HTML
# ============================================================
$ok    = @($resultados | Where-Object { $_.estado -eq 'ok' }).Count
$warn  = @($resultados | Where-Object { $_.estado -eq 'warn' }).Count
$err   = @($resultados | Where-Object { $_.estado -eq 'error' }).Count
$deskn = @($resultados | Where-Object { $_.estado -eq 'desconocido' }).Count
$total = @($resultados).Count

$timestamp = Get-Date -Format 'dd-MMM-yyyy HH:mm'

$filasHtml = foreach ($r in $resultados) {
    $cls = switch ($r.estado) {
        'ok' { 'ok' }
        'warn' { 'warn' }
        'error' { 'err' }
        default { 'unk' }
    }
    $icono = switch ($r.estado) {
        'ok' { '&#x2705;' }
        'warn' { '&#x26A0;' }
        'error' { '&#x274C;' }
        default { '&#x2753;' }
    }
    $logInfo = if ($r.logMtime) { "$($r.logMtime) ($($r.logSize) KB)" } else { '-' }
    @"
        <tr class="$cls">
          <td class="estado">$icono</td>
          <td><strong>$($r.nombre)</strong><br><span class="desc">$($r.descripcion)</span></td>
          <td>$($r.frecuencia)</td>
          <td>$($r.ultima)</td>
          <td>$($r.proxima)</td>
          <td>$logInfo</td>
          <td class="det">$($r.detalle)</td>
        </tr>
"@
}

$dupHtml = if ($dupList.Count -gt 0) {
    $items = foreach ($d in $dupList) { "<li><strong>$($d.Key)</strong> &times;$($d.Value)</li>" }
    @"
    <section class="dup">
      <h2>&#x26A0; Tareas DUPLICADAS detectadas ($($dupList.Count))</h2>
      <p>Patron conocido (Castro ya pasó por esto). Eliminar los duplicados con <code>schtasks /Delete /TN</code>.</p>
      <ul>$($items -join "`n")</ul>
    </section>
"@
} else { '' }

$html = @"
<!DOCTYPE html>
<html lang="es-GT">
<head>
<meta charset="UTF-8">
<title>Salud Agentes - Bodega 1110</title>
<meta http-equiv="refresh" content="300">
<style>
:root {
  --arium-blue: #003A70;
  --arium-gold: #B8860B;
  --ok: #2E8B57;
  --warn: #F4C430;
  --err: #C8102E;
  --unk: #999;
  --bg: #F8F9FA;
}
* { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', sans-serif; }
body { background: var(--bg); color: #333; padding: 1rem; }
header { background: var(--arium-blue); color: white; padding: 1rem 1.5rem; border-radius: 6px; }
h1 { font-size: 1.4rem; }
.timestamp { font-size: 0.85rem; opacity: 0.85; margin-top: 0.3rem; }
.kpi-row { display: flex; gap: 0.8rem; margin: 1rem 0; }
.kpi { flex: 1; background: white; padding: 0.8rem 1rem; border-radius: 6px; border-left: 5px solid var(--arium-blue); }
.kpi.ok { border-color: var(--ok); } .kpi.warn { border-color: var(--warn); }
.kpi.err { border-color: var(--err); } .kpi.unk { border-color: var(--unk); }
.kpi-num { font-size: 2rem; font-weight: 700; }
.kpi.ok .kpi-num { color: var(--ok); } .kpi.warn .kpi-num { color: var(--warn); }
.kpi.err .kpi-num { color: var(--err); } .kpi.unk .kpi-num { color: var(--unk); }
.kpi-lbl { font-size: 0.85rem; text-transform: uppercase; color: #666; }
table { width: 100%; border-collapse: collapse; background: white; margin-top: 1rem; border-radius: 6px; overflow: hidden; }
th { background: var(--arium-blue); color: white; padding: 0.6rem; text-align: left; font-size: 0.85rem; }
td { padding: 0.6rem; border-bottom: 1px solid #eee; font-size: 0.85rem; vertical-align: top; }
tr:nth-child(even) td { background: #FAFAFA; }
tr.ok { border-left: 5px solid var(--ok); }
tr.warn { border-left: 5px solid var(--warn); background: #FFFBF0 !important; }
tr.err { border-left: 5px solid var(--err); background: #FDF2F2 !important; }
tr.unk { border-left: 5px solid var(--unk); }
.estado { text-align: center; font-size: 1.2rem; width: 40px; }
.desc { color: #666; font-size: 0.8rem; }
.det { font-size: 0.78rem; color: #555; font-style: italic; }
.dup { background: #FFF3CD; padding: 1rem 1.5rem; border-radius: 6px; margin: 1rem 0; border-left: 5px solid var(--warn); }
.dup h2 { color: var(--arium-blue); margin-bottom: 0.5rem; }
.dup ul { margin-left: 1.5rem; margin-top: 0.5rem; }
footer { text-align: center; margin-top: 2rem; font-size: 0.8rem; color: #999; }
</style>
</head>
<body>
<header>
  <h1>Salud de Agentes - Bodega 1110</h1>
  <div class="timestamp">Generado: $timestamp &middot; PC: $env:COMPUTERNAME &middot; Autorrefresca c/5min</div>
</header>

<section class="kpi-row">
  <div class="kpi ok"><div class="kpi-num">$ok</div><div class="kpi-lbl">OK</div></div>
  <div class="kpi warn"><div class="kpi-num">$warn</div><div class="kpi-lbl">Advertencia</div></div>
  <div class="kpi err"><div class="kpi-num">$err</div><div class="kpi-lbl">Error</div></div>
  <div class="kpi unk"><div class="kpi-num">$deskn</div><div class="kpi-lbl">Desconocido</div></div>
</section>

$dupHtml

<table>
  <thead>
    <tr>
      <th></th>
      <th>Agente</th>
      <th>Frecuencia</th>
      <th>Ultima corrida</th>
      <th>Proxima</th>
      <th>Log</th>
      <th>Detalle</th>
    </tr>
  </thead>
  <tbody>
$($filasHtml -join "`n")
  </tbody>
</table>

<footer>$total agentes monitoreados &middot; Fuente: schtasks + logs locales &middot; Carlos Trujillo - Jefe Bodega 1110</footer>
</body>
</html>
"@

$html | Out-File -FilePath $outHtml -Encoding UTF8
Write-Host "Generado: $outHtml"
Write-Host "Resumen: $ok OK, $warn WARN, $err ERROR, $deskn desconocidos"
if ($dupList.Count -gt 0) {
    Write-Host "ATENCION: $($dupList.Count) tarea(s) duplicada(s) detectada(s)"
}
