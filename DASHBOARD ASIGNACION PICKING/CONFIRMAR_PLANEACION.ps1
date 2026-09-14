# =====================================================================
#  CONFIRMAR_PLANEACION.ps1  -  Me Llega S.A.
#  Lo dispara el boton "Confirmar planeacion 100%" del dashboard
#  (protocolo confirmarplan:). Busca el correo de Planner de hoy con
#  el adjunto "Estatus de Picking", arma la RESPUESTA A TODOS con el
#  texto de confirmacion y la ABRE COMO BORRADOR para que la revises
#  y le des Enviar. No envia solo.
#  Requiere Outlook abierto.
# =====================================================================
param([string]$url)

$ErrorActionPreference = 'SilentlyContinue'

function Mostrar([string]$titulo, [string]$texto) {
  try {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    [System.Windows.Forms.MessageBox]::Show($texto, $titulo) | Out-Null
  } catch {}
}

# Conectar a Outlook (debe estar abierto)
$ol = $null
try {
  $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
  Mostrar 'Confirmar planeacion' 'No se pudo conectar con Outlook. Abre Outlook y vuelve a intentar.'
  exit 1
}

$ns    = $ol.GetNamespace('MAPI')
$inbox = $ns.GetDefaultFolder(6)
$items = $inbox.Items
$items.Sort('[ReceivedTime]', $true)
$hoy   = (Get-Date).Date

# Buscar el correo de hoy que trae el adjunto "Estatus de Picking"
$found = $null
foreach ($i in $items) {
  try {
    $rt = $i.ReceivedTime
    if ($null -eq $rt -or $rt.Date -gt $hoy) { continue }
    if ($rt.Date -lt $hoy) { break }
    foreach ($att in $i.Attachments) {
      if ("$($att.FileName)" -match '(?i)estatus.*picking.*\.(xlsx|xlsm)$') {
        $found = $i; break
      }
    }
    if ($found) { break }
  } catch {}
}

if (-not $found) {
  Mostrar 'Confirmar planeacion' 'No encontre el correo de Planner con el "Estatus de Picking" recibido hoy en tu Bandeja de entrada. Revisa que ya haya llegado.'
  exit 1
}

# Texto de confirmacion (HTML con entidades para acentos, .ps1 en ASCII)
$msg = @'
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1a1a1a">
<p>Buenos d&iacute;as,</p>
<p>Confirmo que la <b>planeaci&oacute;n del d&iacute;a fue auditada</b> y validada contra la asignaci&oacute;n en K&ouml;rber. Todo qued&oacute; <b>correcto y va al 100%</b>.</p>
<p>Saludos,</p>
</div>
<hr>
'@

try {
  $reply = $found.ReplyAll()
  $reply.HTMLBody = $msg + $reply.HTMLBody
  $reply.Display()   # abre el borrador; tu le das Enviar
} catch {
  Mostrar 'Confirmar planeacion' 'No se pudo crear la respuesta. Detalle: ' + $_.Exception.Message
  exit 1
}
