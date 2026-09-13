# =====================================================================
#  VIGILAR_CORREO_PLANNER.ps1  -  Me Llega S.A.
#  Corre cada hora via Task Scheduler. Revisa si llego el
#  "Estatus de Picking" de Planner. Si lo encuentra y NO se ha
#  procesado hoy, lo procesa UNA VEZ, abre el dashboard, notifica
#  y no vuelve a hacer nada hasta manana.
#  Rapido: entra, revisa, si no hay nada sale en segundos.
#  Requiere Outlook abierto.
# =====================================================================
param(
  [string]$NotificarA = "trujillo.c@ariumcorp.com"
)

$ErrorActionPreference = 'SilentlyContinue'
$base     = $PSScriptRoot
$marcador = Join-Path $base "_picking_procesado.txt"
$hoy      = (Get-Date).Date

# Si ya procesamos hoy, salir de inmediato (0 segundos de costo)
if (Test-Path $marcador) {
  $contenido = ([System.IO.File]::ReadAllText($marcador)).Trim()
  if ($contenido -eq $hoy.ToString('yyyy-MM-dd')) { exit 0 }
}

# Conectar a Outlook (si no esta abierto, salir y reintentar en 1 hora)
$ol = $null
try {
  $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch { exit 0 }

$ns    = $ol.GetNamespace('MAPI')
$inbox = $ns.GetDefaultFolder(6)
$items = $inbox.Items
$items.Sort('[ReceivedTime]', $true)

# Buscar el adjunto "Estatus de Picking" llegado HOY
$encontrado = $false
foreach ($i in $items) {
  try {
    $rt = $i.ReceivedTime
    if ($null -eq $rt -or $rt.Date -gt $hoy) { continue }
    if ($rt.Date -lt $hoy) { break }
    foreach ($att in $i.Attachments) {
      if ("$($att.FileName)" -match '(?i)estatus.*picking.*\.(xlsx|xlsm)$') {
        $encontrado = $true; break
      }
    }
    if ($encontrado) { break }
  } catch {}
}

# Si no llego, salir. La tarea vuelve a intentar en 1 hora.
if (-not $encontrado) { exit 0 }

# === ENCONTRADO === Procesar todo
try {
  $ps1 = Join-Path $base "BUSCAR_CORREOS_PEDIDOS.ps1"
  & powershell -NoProfile -ExecutionPolicy Bypass -File $ps1 -Dias 3
} catch { exit 1 }

# Marcar como procesado hoy (SOLO si realmente proceso)
[System.IO.File]::WriteAllText($marcador, $hoy.ToString('yyyy-MM-dd'), [System.Text.Encoding]::UTF8)

# Abrir el dashboard para revision
try {
  $html = Join-Path $base "Dashboard Asignacion Picking.html"
  if (Test-Path $html) { Start-Process $html }
} catch {}

# Globo de notificacion
try {
  Add-Type -AssemblyName System.Windows.Forms
  $icon = New-Object System.Windows.Forms.NotifyIcon
  $icon.Icon    = [System.Drawing.SystemIcons]::Information
  $icon.Visible = $true
  $icon.BalloonTipTitle = "Dashboard Picking listo"
  $icon.BalloonTipText  = "El Estatus de Picking de hoy ya fue procesado. Revisa el dashboard y confirma."
  $icon.ShowBalloonTip(12000)
  Start-Sleep 13
  $icon.Dispose()
} catch {}

# Correo de aviso
try {
  $mail = $ol.CreateItem(0)
  $mail.To        = $NotificarA
  $mail.Subject   = "Dashboard Picking listo - " + $hoy.ToString('dd/MM/yyyy')
  $mail.HTMLBody  = "<p>Hola,</p><p>El <b>Estatus de Picking del " + $hoy.ToString('dd/MM/yyyy') + "</b> acaba de llegar y ya fue procesado automaticamente.</p><p>Abre el dashboard para revisar y confirmar la programacion del dia.</p><p style='color:#888;font-size:12px'>-- Agente Picking / Me Llega S.A.</p>"
  $mail.Send()
} catch {}
