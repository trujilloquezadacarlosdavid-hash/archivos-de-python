# =====================================================================
#  ABRIR_CORREO_OUTLOOK.ps1  -  Me Llega S.A.
#  Lo dispara el boton "Generar correo (Outlook)" del dashboard
#  (protocolo pickingmail:). Toma el .eml recien descargado y lo abre
#  como BORRADOR en el OUTLOOK DE ESCRITORIO (no el web / no el nuevo),
#  con la imagen y el adjunto ya listos. Solo le das Enviar.
#  Requiere Outlook de escritorio instalado.
# =====================================================================
param([string]$url)

$ErrorActionPreference = 'SilentlyContinue'

function Mostrar([string]$titulo, [string]$texto) {
  try {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    [System.Windows.Forms.MessageBox]::Show($texto, $titulo) | Out-Null
  } catch {}
}

# --- Carpeta Descargas (known folder, con respaldo) ---
$dl = ''
try {
  $dl = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -ErrorAction Stop).'{374DE290-123F-4565-9164-39C4925E467B}'
} catch {}
if (-not $dl -or -not (Test-Path $dl)) { $dl = Join-Path $env:USERPROFILE 'Downloads' }

# --- Esperar el .eml recien descargado (hasta ~12s) ---
$eml = $null
for ($k = 0; $k -lt 24; $k++) {
  $eml = Get-ChildItem -Path $dl -Filter 'Reporte Picking*.eml' -ErrorAction SilentlyContinue |
         Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-2) } |
         Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($eml) { break }
  Start-Sleep -Milliseconds 500
}
if (-not $eml) {
  Mostrar 'Abrir correo en Outlook' 'No encontre el correo recien generado en la carpeta Descargas. Genera el correo con el boton del dashboard y vuelve a intentar.'
  exit 1
}

# --- Outlook de ESCRITORIO ---
$ol = $null
try { $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch {}
if (-not $ol) {
  try { $ol = New-Object -ComObject Outlook.Application } catch {
    Mostrar 'Abrir correo en Outlook' 'No se pudo abrir Outlook de escritorio. Abrelo e intenta de nuevo.'
    exit 1
  }
}

# --- Abrir el .eml como borrador editable en Outlook de escritorio ---
try {
  $item = $ol.Session.OpenSharedItem($eml.FullName)
  $item.Display($false)
} catch {
  Mostrar 'Abrir correo en Outlook' ('No se pudo abrir el correo en Outlook de escritorio. Detalle: ' + $_.Exception.Message)
  exit 1
}
