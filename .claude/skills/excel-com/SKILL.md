---
name: excel-com
description: Reglas y gotchas para automatizar Excel vía COM desde PowerShell o Python (openpyxl también cubierto). Usar SIEMPRE que haya que leer/escribir/formatear un .xlsx o .xlsm en los flujos de Carlos (reclamos Mondelez, cuadres SAP, matrices de descarga, conteo inventario, dashboard que exporta a Excel, formulario inplant). Cubre bug cache .Value2, staging + Copy-Item para archivos abiertos, formato DESPUÉS de datos, liberar COM, sin $null traps, y cuándo usar openpyxl en vez de COM.
---

# Excel COM - Gotchas y patrones

Carlos trabaja con muchos .xlsx/.xlsm (reclamos Mondelez, cuadres SAP, matrices descarga, conteo inventario, formulario inplant). Esta skill centraliza las trampas que ya le costaron tiempo.

## Decisión primero: COM vs openpyxl

| Caso | Usar |
|---|---|
| El Excel tiene macros (.xlsm) o VBA | **COM** (openpyxl no corre macros) |
| El Excel tiene gráficos, tablas dinámicas, formato condicional complejo | **COM** |
| Carlos tiene el archivo ABIERTO en Excel mientras corre el script | **COM + staging** (ver abajo) |
| Es lectura/escritura simple de celdas | **openpyxl** (más rápido, no requiere Excel instalado) |
| Script corre en tarea programada oculta | **openpyxl** si se puede (COM abre instancia) |

## Trampa #1 — Bug cache de .Value2

**Síntoma**: Escribir celda-a-celda cachea el tipo de la PRIMERA escritura y truena mezclando strings+números. Error típico: "Type mismatch" o escribe string donde debería ir número.

❌ **MAL** — truena:
```powershell
$ws.Cells.Item(1, 1).Value2 = "Encabezado"   # cachea string
$ws.Cells.Item(2, 1).Value2 = 1234            # truena o guarda como string
```

✅ **BIEN** — construir arreglo 2D con TODO y escribir de un golpe:
```powershell
$nr = 10; $nc = 3
$arr = New-Object 'object[,]' $nr, $nc
# Llenar el arreglo
$arr[0,0] = "SKU"
$arr[0,1] = "Qty"
$arr[0,2] = "Costo"
$arr[1,0] = "A001"
$arr[1,1] = 1234
$arr[1,2] = 12.50
# ... etc
# Escribir TODO de un solo golpe
$rng = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($nr, $nc))
$rng.Value2 = $arr
```

En Python (win32com):
```python
import pythoncom
arr = [
    ("SKU", "Qty", "Costo"),
    ("A001", 1234, 12.50),
    # ...
]
rng = ws.Range(ws.Cells(1, 1), ws.Cells(len(arr), len(arr[0])))
rng.Value = arr
```

## Trampa #2 — Formato DESPUÉS, nunca MEZCLADO con datos

❌ **MAL**:
```powershell
$ws.Cells.Item(1,1).Value2 = "Header"
$ws.Cells.Item(1,1).Font.Bold = $true
$ws.Cells.Item(1,1).Interior.Color = 0xFF0000
$ws.Cells.Item(2,1).Value2 = "A001"   # ← ya cacheó bold+color, interfiere
```

✅ **BIEN**:
```powershell
# 1) Escribir TODOS los datos primero (arreglo 2D)
$rng.Value2 = $arr

# 2) Aplicar formato DESPUÉS por rangos
$ws.Range("A1:C1").Font.Bold = $true
$ws.Range("A1:C1").Interior.Color = [int](((0 -shl 16) -bor (58 -shl 8) -bor 112))  # #003A70 Arium blue
$ws.Range("B2:B10").NumberFormat = "#,##0"
$ws.Range("C2:C10").NumberFormat = "Q #,##0.00"
```

## Trampa #3 — Archivo abierto por Carlos → usar staging

Si Carlos tiene el .xlsx abierto en Excel mientras el script corre, PowerShell no puede escribirlo directo (file locked).

**Patrón staging**:
```powershell
$origen = "C:\...\MATRIZ RECLAMOS.xlsx"
$staging = Join-Path $env:TEMP "MATRIZ_RECLAMOS_staging_$(Get-Date -Format 'yyyyMMddHHmmss').xlsx"

# 1) Copiar al staging
Copy-Item $origen $staging -Force

# 2) Trabajar sobre el staging
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($staging)
$ws = $wb.Worksheets.Item(1)
# ... escribir ...
$wb.Save()
$wb.Close($false)
$excel.Quit()

# 3) Pedirle a Carlos cerrar el original y copiar el staging encima
# O: dejar el staging con nombre reconocible y avisarle
```

NUNCA intentar `$excel.Workbooks.Open($origen)` cuando otro Excel lo tiene abierto → cuelga el script.

## Trampa #4 — Liberar COM al final

Sin esto queda `EXCEL.EXE` zombie en memoria (acumula varios al día = PC lenta):

```powershell
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable ws, wb, excel
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
```

## Trampa #5 — `$null` traps en pipeline

```powershell
$valor = $ws.Cells.Item(5,2).Value2
# Si la celda está vacía, $valor es $null
# ❌ Esto truena:
$valor + 10
# ✅ Esto no:
if ($null -ne $valor) { $valor + 10 }
# o
[int]($valor ?? 0)   # PS 7+; en 5.1 usar if
```

Comparar `$null` con `-eq`:
```powershell
# ❌ MAL (arrays con nulls se confunden)
if ($valor -eq $null) { ... }
# ✅ BIEN
if ($null -eq $valor) { ... }
```

## Trampa #6 — `.Visible = $false` + `.DisplayAlerts = $false`

Siempre al inicio, SIEMPRE:
```powershell
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false         # no abrir ventana
$excel.DisplayAlerts = $false   # no preguntar "¿quiere guardar?"
$excel.ScreenUpdating = $false  # más rápido
```

Al final, si querés devolver `ScreenUpdating`:
```powershell
$excel.ScreenUpdating = $true
```

## Trampa #7 — Range vs Cells

Para rangos grandes, usar `Range` siempre que puedas:
```powershell
# ❌ Lento (iterar celda a celda)
for ($r=1; $r -le 1000; $r++) {
  $ws.Cells.Item($r, 1).Value2 = "x"
}

# ✅ Rápido (una sola llamada COM)
$rng = $ws.Range("A1:A1000")
$arr = New-Object 'object[,]' 1000, 1
for ($r=0; $r -lt 1000; $r++) { $arr[$r, 0] = "x" }
$rng.Value2 = $arr
```

Diferencia: 20s → 0.3s para 1000 filas.

## Trampa #8 — Guardar como .xlsx (no .xls)

```powershell
# xlFileFormat: 51 = xlsx, 52 = xlsm, 56 = xls (legacy)
$wb.SaveAs($ruta, 51)   # xlsx
$wb.SaveAs($ruta, 52)   # xlsm (si tiene macros)
```

Sin especificar formato puede guardar como .xls (legacy, pierde datos > 65k filas).

## openpyxl equivalente (cuando no se necesita COM)

```python
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

wb = load_workbook("archivo.xlsx")
ws = wb.active

# Escribir
ws["A1"] = "Header"
ws.cell(row=2, column=1, value=1234)

# Formato
ws["A1"].font = Font(bold=True, color="FFFFFF")
ws["A1"].fill = PatternFill("solid", fgColor="003A70")
ws["B2"].number_format = '#,##0'

# Columna completa
for row in ws["B2:B100"]:
    for cell in row:
        cell.number_format = '#,##0'

wb.save("archivo.xlsx")
```

## Colores Arium (hex → entero BGR para COM)

Excel COM usa colores en formato BGR entero (NO RGB hex). Fórmula:
```powershell
function Hex2BGR($hex) {
    $r = [Convert]::ToInt32($hex.Substring(0,2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2,2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4,2), 16)
    return ($b -shl 16) -bor ($g -shl 8) -bor $r
}
$ariumBlue = Hex2BGR "003A70"
$ws.Range("A1:C1").Interior.Color = $ariumBlue
```

## Checklist antes de dar por terminado el script Excel

- [ ] Datos escritos como arreglo 2D de un solo golpe (NO celda a celda)
- [ ] Formato aplicado DESPUÉS de los datos, por rangos
- [ ] `$excel.Visible = $false` + `DisplayAlerts = $false` al inicio
- [ ] COM liberado al final (ReleaseComObject + GC.Collect)
- [ ] Si Carlos pudiera tener el archivo abierto → usé staging + Copy-Item
- [ ] Guardado con `SaveAs(ruta, 51)` para .xlsx
- [ ] Comparaciones `$null` con `$null -eq $x` (no `$x -eq $null`)
- [ ] Si corre en Scheduler → cumple `/tarea-silenciosa` (pythonw, hidden)
