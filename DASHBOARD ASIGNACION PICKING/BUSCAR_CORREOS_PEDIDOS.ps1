# =====================================================================
#  BUSCAR_CORREOS_PEDIDOS.ps1  ·  Me Llega S.A.
#  Lee los correos de planeacion del dia en Outlook (Bandeja de
#  entrada), extrae los numeros de orden pedidos (TR.../ST.../OVG...)
#  y genera "pedidos_data.js" que el Dashboard de Asignacion de
#  Picking auto-carga para marcar lo que NO se agrego.
#
#  Mismo metodo que GENERAR_CALENDARIO.ps1 (Outlook COM).
#  Outlook debe estar ABIERTO.
#
#  Uso:
#    .\BUSCAR_CORREOS_PEDIDOS.ps1                 -> correos de HOY
#    .\BUSCAR_CORREOS_PEDIDOS.ps1 -Dias 2        -> hoy y ayer
#    .\BUSCAR_CORREOS_PEDIDOS.ps1 -Desde 2026-06-16
# =====================================================================
param(
  [string]$Desde = '',          # YYYY-MM-DD (vacio = hoy)
  [int]$Dias = 7,               # cuantos dias hacia atras incluir (7 = semana completa, cubre requerimientos anticipados)
  # Palabras clave ESPECIFICAS de correos de planeacion de picking. No agregar palabras genericas.
  [string]$Claves = 'programacion|programaci|carga frutal|cargas frutal|traslado maq|traslado mac|maquila|picking|planeacion|planeaci|apoya|apoyo|considerar|agregar|incluir|tomar en cuenta|favor agregar|favor incluir',
  [string]$DominioInterno = 'ariumcorp.com',   # solo leer correos de este dominio
  [string]$Salida = ''          # ruta del .js (vacio = junto a este script)
)

$ErrorActionPreference = 'Stop'
if (-not $Salida) { $Salida = Join-Path $PSScriptRoot 'pedidos_data.js' }

$from = if ($Desde) { [datetime]::ParseExact($Desde,'yyyy-MM-dd',$null).Date } else { (Get-Date).Date.AddDays(-([math]::Max(0,$Dias-1))) }
$hoy  = (Get-Date).Date

function Esc([string]$s) {
  if ($null -eq $s) { return '' }
  $s = $s -replace '\\','\\'; $s = $s -replace '"','\"'
  $s = $s -replace "`r",'\r'; $s = $s -replace "`n",'\n'; $s = $s -replace "`t",'\t'
  return $s
}

# ---- Conectar con Outlook (mismo patron que el calendario) ----
Write-Host "Conectando con Outlook..."
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace('MAPI')
$inbox = $ns.GetDefaultFolder(6)   # 6 = olFolderInbox

Write-Host ("Leyendo correos del {0:yyyy-MM-dd} a hoy..." -f $from)

$items = $inbox.Items
$items.Sort('[ReceivedTime]', $true)   # mas nuevos primero

$reOrden  = [regex]'(?i)\b(OVG|OV|TR|ST|RTC|DV)\s*-?\s*(\d{6,})\b'
$reG      = [regex]'(?i)\bG\s*-?\s*(\d{6,})\b'
$reSerial = [regex]'\b(4700\d{6})\b'   # numero de orden suelto (sin prefijo, ej. en tablas con encabezado ST/TR)

# ---- Extraer la FECHA ESPECIFICA pareada con una orden en una tabla de correo ----
# Busca el renglon que contiene el numero de orden y extrae la fecha del mismo renglon o adyacentes
function Get-OrderDate([string]$text, [string]$orden) {
  $digits = ($orden -replace '\D','')
  if (-not $digits) { return '' }
  $lines = @($text -split '[\r\n]+' | Where-Object { $_.Trim() -ne '' })
  $reFecha = '(?i)\b(\d{1,2})\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)(?:\s+(?:de\s+)?(\d{4}))?\b'
  $reDia   = '(?i)\b(lunes|martes|mi[eé]rcoles|jueves|viernes|s[aá]bado|domingo)\b'
  for ($idx = 0; $idx -lt $lines.Count; $idx++) {
    if ($lines[$idx] -match $digits) {
      $w = ($lines[[Math]::Max(0,$idx-2)..[Math]::Min($lines.Count-1,$idx+2)]) -join ' '
      $mF = [regex]::Match($w, $reFecha)
      $mD = [regex]::Match($w, $reDia)
      if ($mF.Success) { return ($(if($mD.Success){$mD.Value+', '}else{''}) + $mF.Value) }
      if ($mD.Success) { return $mD.Value }
    }
  }
  return ''
}

# ---- Extraer texto alrededor de un numero de orden en el cuerpo del correo ----
function Get-ContextSnippet([string]$text, [string]$orden, [int]$chars = 180) {
  $digits = ($orden -replace '\D', '')
  if (-not $digits) { return '' }
  $idx = $text.IndexOf($digits)
  if ($idx -lt 0) { return '' }
  $start = [Math]::Max(0, $idx - $chars)
  $len   = [Math]::Min($text.Length - $start, $chars * 2 + $orden.Length)
  $s = ($text.Substring($start, $len) -replace '[\r\n\t]+', ' ').Trim() -replace ' {2,}', ' '
  if ($s.Length -gt 300) { $s = $s.Substring(0, 297) + '...' }
  return $s
}

# ---- Detectar fechas/requerimientos especiales en un texto ----
function Get-Tags([string]$text) {
  $tags = @()
  if ($text -match '(?i)\b(urgente|urgencia|prioritario|prioridad)\b') { $tags += 'URGENTE' }
  $mDia = [regex]::Match($text, '(?i)\b(lunes|martes|mi[eé]rcoles|jueves|viernes|s[aá]bado|domingo)\b')
  if ($mDia.Success) { $tags += $mDia.Value.ToUpper() }
  $mMes = [regex]::Match($text, '(?i)\b(\d{1,2})\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\b')
  if ($mMes.Success) { $tags += $mMes.Value }
  $mFec = [regex]::Match($text, '\b(\d{1,2})[/\-\.](\d{1,2})(?:[/\-\.]\d{2,4})?\b')
  if ($mFec.Success -and -not $mMes.Success) { $tags += 'fecha: ' + $mFec.Value }
  if ($text -match '(?i)\b(ma[nn]ana|hoy)\b') { $tags += ($Matches[1]) }
  if ($text -match '(?i)\b(temperatura|cadena.fr[ií]o|refriger)\b') { $tags += 'TEMPERATURA' }
  if ($text -match '(?i)\b(requiere|cliente solicita|solicita|necesita|antes de las|especial)\b') { $tags += 'REQUERIMIENTO' }
  return ($tags -join ' | ')
}

# ---- Detectar si un .xlsx es el archivo de Miguel por CONTENIDO (columnas firma) ----
# Columnas: "numero de orden" + ("fuerza de venta" | "vehiculo") = Data Miguel
# Asi no importa como se llame el archivo.
function Test-EsMiguel($path) {
  $xml = ''
  try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zip = $null
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    foreach ($e in $zip.Entries) {
      if ($e.FullName -eq 'xl/sharedStrings.xml') {
        $sr = New-Object System.IO.StreamReader($e.Open())
        $xml = $sr.ReadToEnd().ToLower()
        $sr.Close(); $sr.Dispose()
        break
      }
    }
    $zip.Dispose(); $zip = $null
  } catch {
    if ($zip) { try { $zip.Dispose() } catch {} }
  }
  if (-not $xml) { return $false }
  $tieneOrden    = $xml.Contains('numero de orden')
  $tieneVehiculo = $xml.Contains('vehiculo')
  $tieneFuerza   = $xml.Contains('fuerza de venta')
  # Firma: "numero de orden" + (vehiculo O fuerza de venta)
  return ($tieneOrden -and ($tieneVehiculo -or $tieneFuerza))
}

# ---- Leer un .xlsx adjunto como ZIP/XML (sin abrir Excel) y sacar las ordenes ----
function Get-OrdersFromXlsx($path) {
  $set = New-Object System.Collections.Generic.HashSet[string]
  try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
      foreach ($e in $zip.Entries) {
        if ($e.FullName -match '(?i)^xl/(sharedStrings\.xml$|worksheets/)') {
          $sr = New-Object System.IO.StreamReader($e.Open())
          $xml = $sr.ReadToEnd(); $sr.Dispose()
          foreach ($m in $reOrden.Matches($xml))  { [void]$set.Add(($m.Groups[1].Value + $m.Groups[2].Value).ToUpper()) }
          foreach ($m in $reSerial.Matches($xml)) { [void]$set.Add($m.Groups[1].Value) }
        }
      }
    } finally { $zip.Dispose() }
  } catch {}
  return $set
}

$correos = New-Object System.Collections.ArrayList
$requerimientos = New-Object System.Collections.ArrayList   # correos con keywords pero sin numero de orden
$todasOrdenes = New-Object System.Collections.Generic.HashSet[string]
$contextos = @{}   # orden -> texto de contexto del correo que la menciono
$revisados = 0
$estatusB64 = $null
$estatusFilename = $null
$miguelB64 = $null
$miguelFilename = $null

foreach ($i in $items) {
  try {
    $rt = $i.ReceivedTime
    if ($null -eq $rt) { continue }
    if ($rt.Date -gt $hoy) { continue }
    if ($rt.Date -lt $from) { break }   # ya entramos a dias viejos -> fin
    $revisados++
    if ($revisados -gt 400) { break }   # tope de seguridad

    $subj = "$($i.Subject)"
    $body = "$($i.Body)"
    $hay  = ($subj + ' ' + $body)

    # ---- Filtro de dominio: ignorar correos externos (BAC, bancos, sistemas ajenos) ----
    # Exchange interno devuelve DN (/O=EXCHANGELABS/...) no SMTP; detectar ambos
    $senderEmail = ''
    try { $senderEmail = "$($i.SenderEmailAddress)".ToLower() } catch {}
    $senderType = ''
    try { $senderType = "$($i.SenderEmailType)".ToUpper() } catch {}
    $esInterno = ($senderEmail -like "*$DominioInterno*") -or ($senderType -eq 'EX')

    $ordenes = New-Object System.Collections.Generic.HashSet[string]
    foreach ($m in $reOrden.Matches($hay))  { [void]$ordenes.Add(($m.Groups[1].Value + $m.Groups[2].Value).ToUpper()) }
    foreach ($m in $reG.Matches($hay))      { [void]$ordenes.Add('OVG' + $m.Groups[1].Value) }
    foreach ($m in $reSerial.Matches($hay)) { [void]$ordenes.Add($m.Groups[1].Value) }

    # Saltar si no es interno Y no tiene numeros de orden reconocibles
    if (-not $esInterno -and $ordenes.Count -eq 0) { continue }

    # ---- Buscar adjunto "Estatus de Picking" (asignacion Körber de Planner) ----
    if (-not $estatusB64) {
      try {
        foreach ($att in $i.Attachments) {
          $fn = "$($att.FileName)"
          if ($fn -match '(?i)estatus.*picking.*\.(xlsx|xlsm)$') {
            $tmp = Join-Path $env:TEMP ('est_' + [guid]::NewGuid().ToString('N') + '.xlsx')
            try {
              $att.SaveAsFile($tmp)
              $bytes = [System.IO.File]::ReadAllBytes($tmp)
              $estatusB64 = [Convert]::ToBase64String($bytes)
              $estatusFilename = $fn
            } finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
            break
          }
        }
      } catch {}
    }

    # ---- Buscar adjunto de Miguel por NOMBRE o por CONTENIDO (columnas firma) ----
    if (-not $miguelB64) {
      foreach ($att in $i.Attachments) {
        try {
          $fn = "$($att.FileName)"
          if ($fn -notmatch '(?i)\.(xlsx|xlsm)$') { continue }
          if ($fn -match '(?i)estatus.*picking') { continue }
          # Match rapido por nombre (cubre GT-FO-TR, reporte de carga, planificacion frutal)
          $nombreOK = $fn -match '(?i)(reporte.*carga|GT-FO-TR|planificaci.n.*frutal)'
          $tmp = Join-Path $env:TEMP ('mig_' + [guid]::NewGuid().ToString('N') + '.xlsx')
          $att.SaveAsFile($tmp)
          $contenidoOK = $false
          if (-not $nombreOK) { $contenidoOK = Test-EsMiguel $tmp }
          if ($nombreOK -or $contenidoOK) {
            $miguelB64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($tmp))
            $miguelFilename = $fn
            Write-Host ("   [MIGUEL] Detectado: {0} (nombre={1} contenido={2})" -f $fn, $nombreOK, $contenidoOK)
          }
          try { [System.IO.File]::Delete($tmp) } catch {}
          if ($miguelB64) { break }
        } catch {
          Write-Host ("   [MIGUEL-ERR] {0}: {1}" -f "$($att.FileName)", $_.Exception.Message)
        }
      }
    }

    $matchKw = ($hay -match $Claves)
    if (-not $matchKw -and $ordenes.Count -eq 0) { continue }   # ni clave ni orden -> ignorar

    # ---- Leer adjuntos .xlsx/.xlsm de los correos de planeacion ----
    $adjInfo = ''
    if ($matchKw) {
      try {
        foreach ($att in $i.Attachments) {
          $fn = "$($att.FileName)"
          if ($fn -match '(?i)\.(xlsx|xlsm)$') {
            $tmp = Join-Path $env:TEMP ('ped_' + [guid]::NewGuid().ToString('N') + '.xlsx')
            try {
              $att.SaveAsFile($tmp)
              $got = Get-OrdersFromXlsx $tmp
              foreach ($o in $got) { [void]$ordenes.Add($o) }
              if ($got.Count -gt 0) { $adjInfo += (" [adjunto {0}: {1} orden(es)]" -f $fn, $got.Count) }
            } finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
          }
        }
      } catch {}
    }

    $sender = ''
    try { $sender = "$($i.SenderName)" } catch {}
    foreach ($o in $ordenes) {
      [void]$todasOrdenes.Add($o)
      # Guardar contexto (solo del primer correo que menciona la orden)
      if (-not $contextos.ContainsKey($o)) {
        $ordDate = Get-OrderDate $body $o   # fecha especifica del renglon de esta orden (tablas de plan)
        $snippet = Get-ContextSnippet $body $o
        $tags    = if ($snippet) { Get-Tags $snippet } else { '' }

        # Si encontramos fecha especifica (tabla planificacion semanal), reemplazar tags
        # genericos de dia/fecha por la fecha correcta de ESTA orden concretamente
        if ($ordDate) {
          $tagsArr = @($tags -split '\|') | ForEach-Object { $_.Trim() } | Where-Object {
            $_ -ne '' -and
            $_ -notmatch '(?i)^(lunes|martes|mi.rcoles|jueves|viernes|sabado|domingo)$' -and
            $_ -notmatch '(?i)^\d{1,2}\s+de\s+\w+' -and
            $_ -notmatch '(?i)^fecha:'
          }
          $tagsLimpios = $tagsArr -join ' | '
          $tags = 'PLANIFICADO: ' + $ordDate + $(if ($tagsLimpios) { ' | ' + $tagsLimpios } else { '' })
        }

        $ctx = 'De: ' + $sender + ' | ' + $subj
        if ($snippet) { $ctx += ' | ' + $snippet }
        if ($tags)    { $ctx = '[' + $tags + '] ' + $ctx }
        $contextos[$o] = $ctx
      }
    }

    # Correos con keywords pero SIN numero de orden: solo internos, pueden contener requerimientos anticipados
    if ($matchKw -and $ordenes.Count -eq 0 -and $esInterno) {
      $bodySnippet = ($body -replace '[\r\n\t]+',' ').Trim() -replace ' {2,}',' '
      if ($bodySnippet.Length -gt 500) { $bodySnippet = $bodySnippet.Substring(0, 497) + '...' }
      $tags = Get-Tags ($subj + ' ' + $bodySnippet)
      [void]$requerimientos.Add([pscustomobject]@{
        fecha  = $rt.ToString('yyyy-MM-dd HH:mm')
        de     = $sender
        asunto = $subj
        tags   = $tags
        cuerpo = $bodySnippet
      })
    }

    [void]$correos.Add([pscustomobject]@{
      fecha    = $rt.ToString('yyyy-MM-dd HH:mm')
      de       = $sender
      asunto   = $subj
      adj      = $adjInfo
      ordenes  = @($ordenes)
    })
  } catch {}
}

# ---- Construir el "texto" que el dashboard mete a la caja de Verificar ----
$sbTxt = New-Object System.Text.StringBuilder
foreach ($c in $correos) {
  [void]$sbTxt.AppendLine(("=== {0} | DE: {1} | {2}{3} ===" -f $c.fecha, $c.de, $c.asunto, $c.adj))
  if ($c.ordenes.Count -gt 0) { [void]$sbTxt.AppendLine(("Ordenes: {0}" -f ($c.ordenes -join ', '))) }
  [void]$sbTxt.AppendLine('')
}
$texto = $sbTxt.ToString()

# ---- JS de auto-carga (mismo estilo que calendario_data.js) ----
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('window.__EMBEDDED_PEDIDOS__ = {')
[void]$sb.AppendLine('  "generado": "' + (Esc (Get-Date).ToString('yyyy-MM-dd HH:mm')) + '",')
[void]$sb.AppendLine('  "fecha": "' + $hoy.ToString('yyyy-MM-dd') + '",')
[void]$sb.AppendLine('  "ordenes": [' + (($todasOrdenes | ForEach-Object { '"' + (Esc $_) + '"' }) -join ', ') + '],')
[void]$sb.AppendLine('  "texto": "' + (Esc $texto) + '",')
$reqJson = ($requerimientos | ForEach-Object { '{"fecha":"'+(Esc $_.fecha)+'","de":"'+(Esc $_.de)+'","asunto":"'+(Esc $_.asunto)+'","tags":"'+(Esc $_.tags)+'","cuerpo":"'+(Esc $_.cuerpo)+'"}' }) -join ','
[void]$sb.AppendLine('  "requerimientos": [' + $reqJson + '],')
$ctxPairs = $contextos.GetEnumerator() | ForEach-Object { '"' + (Esc $_.Key) + '":"' + (Esc $_.Value) + '"' }
[void]$sb.AppendLine('  "contextos": {' + ($ctxPairs -join ',') + '}')
[void]$sb.AppendLine('};')

$utf8bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($Salida, $sb.ToString(), $utf8bom)

# ---- estatus_data.js (adjunto "Estatus de Picking" de Planner para auto-carga en dashboard) ----
$salidaEst = Join-Path $PSScriptRoot 'estatus_data.js'
if ($estatusB64) {
  $sbEst = New-Object System.Text.StringBuilder
  [void]$sbEst.AppendLine('window.__EMBEDDED_ESTATUS__ = {')
  [void]$sbEst.AppendLine('  "generado": "' + (Esc (Get-Date).ToString('yyyy-MM-dd HH:mm')) + '",')
  [void]$sbEst.AppendLine('  "fecha": "' + $hoy.ToString('yyyy-MM-dd') + '",')
  [void]$sbEst.AppendLine('  "filename": "' + (Esc $estatusFilename) + '",')
  [void]$sbEst.AppendLine('  "b64": "' + $estatusB64 + '"')
  [void]$sbEst.AppendLine('};')
  [System.IO.File]::WriteAllText($salidaEst, $sbEst.ToString(), $utf8bom)
} else {
  [System.IO.File]::WriteAllText($salidaEst, 'window.__EMBEDDED_ESTATUS__ = null;', $utf8bom)
}

# ---- miguel_data.js (REPORTE DE CARGA de Miguel para el plan del dia) ----
$salidaMig = Join-Path $PSScriptRoot 'miguel_data.js'
if ($miguelB64) {
  $sbMig = New-Object System.Text.StringBuilder
  [void]$sbMig.AppendLine('window.__EMBEDDED_MIGUEL__ = {')
  [void]$sbMig.AppendLine('  "generado": "' + (Esc (Get-Date).ToString('yyyy-MM-dd HH:mm')) + '",')
  [void]$sbMig.AppendLine('  "fecha": "' + $hoy.ToString('yyyy-MM-dd') + '",')
  [void]$sbMig.AppendLine('  "filename": "' + (Esc $miguelFilename) + '",')
  [void]$sbMig.AppendLine('  "b64": "' + $miguelB64 + '"')
  [void]$sbMig.AppendLine('};')
  [System.IO.File]::WriteAllText($salidaMig, $sbMig.ToString(), $utf8bom)
} else {
  [System.IO.File]::WriteAllText($salidaMig, 'window.__EMBEDDED_MIGUEL__ = null;', $utf8bom)
}

Write-Host ""
Write-Host ("OK -> {0}" -f $Salida) -ForegroundColor Green
Write-Host ("   {0} correo(s) de planeacion · {1} orden(es) detectada(s)." -f $correos.Count, $todasOrdenes.Count)
if ($todasOrdenes.Count -gt 0) { Write-Host ("   Ordenes: {0}" -f (($todasOrdenes) -join ', ')) }
if ($estatusB64) {
  Write-Host ("   [OK] Estatus de Picking: {0}" -f $estatusFilename) -ForegroundColor Cyan
} else {
  Write-Host "   (Sin adjunto Estatus de Picking)" -ForegroundColor Yellow
}
if ($miguelB64) {
  Write-Host ("   [OK] Reporte de Carga Miguel: {0}" -f $miguelFilename) -ForegroundColor Cyan
} else {
  Write-Host "   (Sin adjunto Reporte de Carga de Miguel)" -ForegroundColor Yellow
}
