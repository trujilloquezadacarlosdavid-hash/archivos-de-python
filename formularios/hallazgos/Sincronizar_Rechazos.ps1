# =====================================================================
# SINCRONIZAR_RECHAZOS.ps1
# Descarga los hallazgos del Google Sheet y los agrega al Excel local
# =====================================================================

param(
    [string]$SheetId   = "PEGA_AQUI_TU_SHEET_ID",
    [string]$ExcelPath = "$env:USERPROFILE\OneDrive - Me Llega, S.A\Escritorio\Copia de RechazosModerno_Inplant2025 (2).xlsx"
)

$ErrorActionPreference = 'Stop'

# --- 1. Descargar CSV del Google Sheet publicado ---
Write-Host "[1/4] Descargando datos de Google Sheets..." -ForegroundColor Cyan
$csvUrl = "https://docs.google.com/spreadsheets/d/$SheetId/gviz/tq?tqx=out:csv&sheet=Respuestas"
$csvPath = "$env:TEMP\hallazgos_sync_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

try {
    Invoke-WebRequest -Uri $csvUrl -OutFile $csvPath -UseBasicParsing
} catch {
    Write-Host "ERROR: No se pudo descargar. Verifica que el Sheet este publicado." -ForegroundColor Red
    Write-Host "  Sheet ID: $SheetId" -ForegroundColor Yellow
    Write-Host "  En Google Sheets: Archivo > Compartir > Publicar en la web > Hoja 'Respuestas' > CSV" -ForegroundColor Yellow
    exit 1
}

$newRows = Import-Csv $csvPath -Encoding UTF8
$total = $newRows.Count
Write-Host "  $total registros encontrados en Google Sheet" -ForegroundColor Green

if ($total -eq 0) {
    Write-Host "No hay registros nuevos. Nada que sincronizar." -ForegroundColor Yellow
    exit 0
}

# --- 2. Abrir Excel ---
Write-Host "[2/4] Abriendo Excel..." -ForegroundColor Cyan
$staging = "$env:TEMP\RechazosSync_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
Copy-Item $ExcelPath $staging -Force

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($staging)
$ws = $wb.Worksheets.Item("DETALLERECHAZOS")
$lastRow = $ws.UsedRange.Rows.Count

# --- 3. Detectar filas ya existentes (por timestamp) ---
Write-Host "[3/4] Comparando registros existentes..." -ForegroundColor Cyan
$existing = @{}
# Leer timestamps existentes de una columna auxiliar o comparar por fecha+codigo+cantidad
# Usamos compositeKey: fecha|codigo|cantidad|motivo
for ($r = 2; $r -le $lastRow; $r++) {
    $f = $ws.Cells.Item($r, 1).Text   # Fecha
    $c = $ws.Cells.Item($r, 5).Text   # Codigo
    $q = $ws.Cells.Item($r, 6).Text   # Cantidad
    $m = $ws.Cells.Item($r, 8).Text   # Motivo
    $key = "$f|$c|$q|$m"
    $existing[$key] = $true
}

# --- 4. Agregar filas nuevas ---
Write-Host "[4/4] Agregando registros nuevos..." -ForegroundColor Cyan
$added = 0
foreach ($row in $newRows) {
    $fecha = $row.'FECHA ENTREGA'
    if (-not $fecha) { continue }

    # Convertir fecha de yyyy-mm-dd a dd/mm/yyyy si es necesario
    if ($fecha -match '^\d{4}-\d{2}-\d{2}$') {
        $parts = $fecha -split '-'
        $fechaExcel = "$($parts[2])/$($parts[1])/$($parts[0])"
    } else {
        $fechaExcel = $fecha
    }

    $codigo  = $row.'CODIGO ARIUM'
    $cant    = $row.'CANTIDAD'
    $motivo  = $row.'MOTIVO'
    $key     = "$fechaExcel|$codigo|$cant|$motivo"

    if ($existing.ContainsKey($key)) { continue }

    $lastRow++
    $ws.Cells.Item($lastRow, 1).Value2  = $fechaExcel              # FECHA ENTREGA
    $ws.Cells.Item($lastRow, 2).Value2  = [int]($row.'MES')        # MES
    $ws.Cells.Item($lastRow, 3).Value2  = [int]($row.'AÑO')        # AÑO
    $ws.Cells.Item($lastRow, 4).Value2  = $row.'CITA'              # CITA
    $ws.Cells.Item($lastRow, 5).Value2  = $codigo                  # CODIGO ARIUM
    $ws.Cells.Item($lastRow, 6).Value2  = [int]$cant               # CANTIDAD
    $ws.Cells.Item($lastRow, 7).Value2  = $row.'UM RECHAZADA'      # UM RECHAZADA
    $ws.Cells.Item($lastRow, 8).Value2  = $motivo                  # MOTIVO
    $ws.Cells.Item($lastRow, 9).Value2  = [double]($row.'MONTO UM')# MONTO UM
    $ws.Cells.Item($lastRow, 10).Value2 = $row.'DESCRIPCION'       # DESCRIPCION
    $ws.Cells.Item($lastRow, 11).Value2 = $row.'UM MD'             # UM MD
    $ws.Cells.Item($lastRow, 12).Value2 = [int]($row.'UXC')        # UXC
    $ws.Cells.Item($lastRow, 13).Value2 = [double]($row.'MONTO TOTAL') # MONTO TOTAL
    $ws.Cells.Item($lastRow, 14).Value2 = $row.'OC / OV'           # OC / OV
    $ws.Cells.Item($lastRow, 15).Value2 = $row.'CLIENTE'           # CLIENTE
    $ws.Cells.Item($lastRow, 16).Value2 = $row.'CASA'              # CASA
    $ws.Cells.Item($lastRow, 17).Value2 = $row.'RESPONSABILIDAD'   # RESPONSABILIDAD

    $existing[$key] = $true
    $added++
}

$wb.Save()
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

# Copiar de vuelta al original
Copy-Item $staging $ExcelPath -Force
Remove-Item $staging -Force -ErrorAction SilentlyContinue
Remove-Item $csvPath -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=== LISTO ===" -ForegroundColor Green
Write-Host "  $added registros nuevos agregados al Excel" -ForegroundColor Green
Write-Host "  $($total - $added) ya existian (duplicados ignorados)" -ForegroundColor Yellow
