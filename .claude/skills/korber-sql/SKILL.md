---
name: korber-sql
description: Reglas y gotchas para consultar la base SQL Server de Körber (172.16.5.43 / IN_E) desde PowerShell, Python o Excel COM. Usar SIEMPRE que haya que escribir una query, un conteo, un reporte, un dashboard o un cuadre basado en datos de Körber. Cubre connection string estándar, overflow de SUM, tipos time que rompen la lectura, joins OR que no indexan, HJ_Inventory BASE vs vista, convertidor a cajas, códigos de transacción (341 despacho) y bodega física C001.
---

# Körber SQL - Conexión y gotchas

Estas son las trampas que ya le costaron horas a Carlos. Antes de escribir cualquier query, revisá si aplica.

## Conexión estándar

```
DRIVER={ODBC Driver 17 for SQL Server};SERVER=172.16.5.43;DATABASE=IN_E;Trusted_Connection=yes;TrustServerCertificate=yes
```

En Python (pyodbc):
```python
import pyodbc
cn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=172.16.5.43;DATABASE=IN_E;"
    "Trusted_Connection=yes;TrustServerCertificate=yes"
)
```

En PowerShell (ADO.NET):
```powershell
$cs = "Server=172.16.5.43;Database=IN_E;Integrated Security=True;TrustServerCertificate=True"
$cn = New-Object System.Data.SqlClient.SqlConnection($cs)
```

En Excel VBA (ADODB):
```vba
cn.Open "Provider=MSOLEDBSQL;Server=172.16.5.43;Database=IN_E;Trusted_Connection=yes;TrustServerCertificate=yes"
```

## Trampa 1 — SUM(int) OVERFLOW

Las columnas `Qty` y afines son INT. Sumar grandes volúmenes revienta.

```sql
-- ❌ Rompe con "Arithmetic overflow"
SELECT SUM(Qty) FROM HJ_TRANSACTIONS WHERE WhsCode = '1110'

-- ✅ Forzar bigint
SELECT SUM(CAST(Qty AS bigint)) FROM HJ_TRANSACTIONS WHERE WhsCode = '1110'
```

## Trampa 2 — Columnas tipo TIME rompen GetSchemaTable

`TrxStartTime` y `TrxEndTime` son tipo `time`. Al leer con pyodbc / DataReader de .NET, revientan `GetSchemaTable()`.

```sql
-- ❌ Falla al leer
SELECT TrxDate, TrxStartTime, TrxEndTime, UserID FROM HJ_TRANSACTIONS

-- ✅ Usar solo TrxDate, o castear
SELECT TrxDate, CAST(TrxStartTime AS varchar(8)) AS StartT, UserID FROM HJ_TRANSACTIONS
```

Si solo necesitás la fecha, bastaría con `TrxDate`.

## Trampa 3 — Join OR a HJ_TRANSACTIONS NO usa índice

```sql
-- ❌ Lento, no indexa, se cae por timeout en producción
SELECT * FROM HJ_TRANSACTIONS t
WHERE t.SourceLocID = @loc OR t.TargetLocID = @loc

-- ✅ UNION ALL o partir por índice
SELECT * FROM HJ_TRANSACTIONS WHERE SourceLocID = @loc
UNION ALL
SELECT * FROM HJ_TRANSACTIONS WHERE TargetLocID = @loc
```

## Trampa 4 — HJ_Inventory: BASE vs VISTA

- **`HJ_Inventory` (tabla BASE)** — incluye producto sin matrícula (nivel picking)
- **`V_HJ_Inventory_1110` (vista)** — OMITE producto sin matrícula (filtrada para slotting)

Para slotting / posiciones vacías → usar la BASE, porque una posición de pick con producto sin matrícula igual está OCUPADA aunque la vista no lo muestre.

## Trampa 5 — Cajas vs unidades

Los reportes externos piden cajas (CJ), pero Qty está en UND. Convertir vía `HJ_MasterData`:

```sql
SELECT 
    m.SKU,
    SUM(CAST(t.Qty AS bigint)) AS UND,
    SUM(CAST(t.Qty AS bigint)) / NULLIF(uom.ConversionFactor, 0) AS CJ
FROM HJ_TRANSACTIONS t
JOIN HJ_MasterData m ON m.SKU = t.SKU
JOIN HJ_MasterDataUom uom ON uom.SKU = m.SKU AND uom.Uom = 'CJ'
WHERE t.TranType = 341
GROUP BY m.SKU, uom.ConversionFactor
```

## Códigos de transacción clave

| TranType | Qué es |
|---|---|
| 341 | Despacho / salida (USAR ESTE, no 322) |
| 301 | Pick FEFO (fuente de pick) |
| ... | preguntar antes de asumir |

## Bodega física

- `C001` = bodega física, ~8,349 posiciones (~62% ocupación)
- `location_id`: `PISO-B*` / `CB` = bloque, `RAMPA` = staging, `P##C##[A-E]` donde A=piso, E=altura
- `HJ_Location.StorageStrategy`: A / B / C
- Capacidad real del bloque **NO está en el maestro SQL** — está en código de Carlos (hardcoded)
- Producto sin matrícula a nivel picking = posición ocupada

## Casas, status y filtros frecuentes

- `WhsCode LIKE '1110%'` para bodega 1110
- `status='A'` activos, `status='H'` hold/dañado
- OC/ST **4700 = casa importada DUCA**, NO es traslado interno
- "Heinz" en reportes excluye Philips, Listerine, Durex, Edgewell (otras casas Kraft/aliadas)
- Centinela de fecha: `F_Expira = 1900-01-01` significa "sin fecha", NO vencido
- Pick rule real = procedure `usp_PkLstCJQtyFEFO`

## Cuadre SAP vs Körber

- Fuente correcta = `CIERRE_INVENTARIO_STATUS` (no las vistas regulares)
- "Diferencias abismales" en el diario suelen ser pares +/- por status (mismas unidades en distinto estado SAP vs Körber que se cancelan)

## Reporte diario bodega 1110

- Grupo WA destino: **Recepciones 2026** (código `DVx3epM5Gx00kZvxBESxuv`)
- Reglas de envío: ver skill `/wa-send`

## Antes de entregar números

Aplicar skill `/alcance` SIEMPRE: Tarjeta de Alcance primero, OK del usuario, ejecución después.
