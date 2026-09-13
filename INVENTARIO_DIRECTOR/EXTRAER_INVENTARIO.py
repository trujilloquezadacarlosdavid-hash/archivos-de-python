"""
INVENTARIO DIRECTOR - Extractor SQL
Conecta a Korber WMS (172.16.5.43/IN_E) y genera:
  1) inventario_data.js  (inventario completo embebido)
  2) INVENTARIO_COMPLETO.xlsx (exportable)
"""
import pyodbc, json, sys, os, datetime, base64, io, struct

sys.stdout.reconfigure(encoding='utf-8')
DIR = os.path.dirname(os.path.abspath(__file__))

SQL_SERVER = "172.16.5.43"
SQL_DB = "IN_E"
SQL_DRIVER = "ODBC Driver 17 for SQL Server"
BODEGA = "1110A001"
BODEGA_COD = "1110"
YEAR = datetime.date.today().year

def connect():
    return pyodbc.connect(
        f"DRIVER={{{SQL_DRIVER}}};SERVER={SQL_SERVER};DATABASE={SQL_DB};"
        "Trusted_Connection=yes;TrustServerCertificate=yes",
        timeout=30
    )

def query(conn, sql, params=None):
    cur = conn.cursor()
    if params:
        cur.execute(sql, params)
    else:
        cur.execute(sql)
    cols = [c[0] for c in cur.description]
    rows = []
    for r in cur.fetchall():
        rows.append(dict(zip(cols, [v for v in r])))
    return rows

def serial(obj):
    if isinstance(obj, (datetime.date, datetime.datetime)):
        return obj.isoformat()
    if isinstance(obj, (float,)) and obj != obj:
        return None
    return str(obj)

print("Conectando a SQL...")
conn = connect()

# ── 1) INVENTARIO COMPLETO ──────────────────────────────────────────
print("Extrayendo inventario completo...")
inv_rows = query(conn, """
SELECT
    i.Location,
    SUBSTRING(i.Location, 1,
        CASE WHEN PATINDEX('%[0-9]%', i.Location) > 0
             THEN PATINDEX('%[0-9]%', i.Location) + 1
             ELSE LEN(i.Location) END
    ) AS Pasillo,
    i.ItemCode,
    i.ItemName,
    i.Qty,
    i.AvailableQty,
    i.BlockedQty,
    i.UnavailableQty,
    i.Status,
    i.BlockedCode,
    i.LotNumber,
    i.ExpirationDate,
    i.FifoDate,
    i.License,
    i.UM,
    m.Class AS Categoria,
    i.PostedDatetime
FROM LOGISTICA.HJ_Inventory i
LEFT JOIN LOGISTICA.HJ_MasterData m
    ON i.Warehouse = m.WhsCode AND i.ItemCode = m.ItemCode
WHERE i.Warehouse = ?
    AND i.Qty > 0
ORDER BY i.Location, i.ItemCode
""", [BODEGA])
print(f"  {len(inv_rows)} registros de inventario")

# Calcular pasillo correctamente (P01, P02... CB, PI, etc.)
for r in inv_rows:
    loc = r['Location'] or ''
    if loc.startswith('P') and len(loc) >= 3 and loc[1:3].isdigit():
        r['Pasillo'] = loc[:3]
    elif loc.startswith('CB'):
        r['Pasillo'] = 'CB'
    elif loc.startswith('PI'):
        r['Pasillo'] = 'PI'
    elif loc.startswith('SC'):
        r['Pasillo'] = 'SC'
    elif loc.startswith('PU'):
        r['Pasillo'] = 'PU'
    elif loc.startswith('HH'):
        r['Pasillo'] = 'HH'
    elif loc.startswith('RA'):
        r['Pasillo'] = 'RA'
    elif loc.startswith('RE'):
        r['Pasillo'] = 'RE'
    elif loc.startswith('RC'):
        r['Pasillo'] = 'RC'
    elif loc.startswith('UN'):
        r['Pasillo'] = 'UN'
    elif 'DANADO' in loc or 'DANO' in loc:
        r['Pasillo'] = 'DANADO'
    elif 'SCRAP' in loc:
        r['Pasillo'] = 'SCRAP'
    elif 'DIF' in loc:
        r['Pasillo'] = 'DIF'
    elif 'DEV' in loc:
        r['Pasillo'] = 'DEV'
    elif 'CD-' in loc:
        r['Pasillo'] = 'CD'
    else:
        r['Pasillo'] = loc[:3] if len(loc) >= 3 else loc

# ── 2) RECEPCIONES KRAFT HEINZ 2026 ────────────────────────────────
print("Extrayendo recepciones Kraft Heinz...")
kh_rows = query(conn, """
SELECT
    e.cita_id AS Cita,
    e.fecha_cita AS FechaCita,
    e.proveedor_nombre_co AS Proveedor,
    e.horario_cita AS HorarioCita,
    e.rampa_num AS Rampa,
    e.nombre_piloto AS Piloto,
    e.placas_unidad AS Placa,
    e.cita_completada AS Completada,
    e.hora_inicio AS HoraInicio,
    e.hora_fin AS HoraFin,
    e.f_finalizacion AS Finalizacion,
    d.orden_compra AS OC,
    d.codigo_sku AS SKU,
    d.descripcion_sku AS Producto,
    d.cajas_enviadas AS CajasEnviadas,
    d.uxc AS UxC,
    m.Class AS Categoria
FROM LOGISTICA.ENCABEZADO_RECEPCION_20 e
JOIN LOGISTICA.DETALLE_RECEPCION_20 d ON e.cita_id = d.cita_id
LEFT JOIN LOGISTICA.HJ_MasterData m ON m.WhsCode = e.bodega_codigo AND m.ItemCode = d.codigo_sku
WHERE e.bodega_codigo = ?
    AND e.fecha_cita >= ?
    AND e.enc_canceled = 'N'
    AND d.det_canceled = 'N'
    AND (e.proveedor_nombre_co LIKE '%KRAFT%' OR e.proveedor_nombre_co LIKE '%HEINZ%')
ORDER BY e.fecha_cita DESC, e.cita_id, d.codigo_sku
""", [BODEGA_COD, f'{YEAR}-01-01'])
print(f"  {len(kh_rows)} lineas Kraft Heinz")

# ── 3) PASILLO 34 (DANADO) DETALLE ─────────────────────────────────
print("Extrayendo pasillo 34 (dañado)...")
p34_rows = query(conn, """
SELECT
    i.Location,
    i.ItemCode,
    i.ItemName,
    i.Qty,
    i.Status,
    i.BlockedCode,
    i.LotNumber,
    i.ExpirationDate,
    i.License,
    m.Class AS Categoria
FROM LOGISTICA.HJ_Inventory i
LEFT JOIN LOGISTICA.HJ_MasterData m
    ON i.Warehouse = m.WhsCode AND i.ItemCode = m.ItemCode
WHERE i.Warehouse = ?
    AND i.Location LIKE 'P34%'
    AND i.Qty > 0
ORDER BY i.Location, i.ItemCode
""", [BODEGA])
print(f"  {len(p34_rows)} registros en P34")

# ── 4) CIERRE SAP vs KORBER ────────────────────────────────────────
print("Extrayendo cierre inventario SAP vs Korber...")
cierre_rows = query(conn, """
SELECT SKU, Nombre_SKU, UM_SAP, Cant_SAP, UM_HJ, Cant_HJ,
       Status, Diferencia_Inventarios, Tipo_Diferencia, ExecDate
FROM LOGISTICA.CIERRE_INVENTARIO_STATUS
WHERE Bodega = '1110A001'
    AND ExecDate = (SELECT MAX(ExecDate) FROM LOGISTICA.CIERRE_INVENTARIO_STATUS WHERE Bodega='1110A001')
ORDER BY SKU, Status
""")
print(f"  {len(cierre_rows)} registros cierre")

# ── 5) RESUMEN POR PASILLO ─────────────────────────────────────────
pasillo_summary = {}
for r in inv_rows:
    p = r['Pasillo']
    if p not in pasillo_summary:
        pasillo_summary[p] = {'Pasillo': p, 'Unidades': 0, 'SKUs': set(), 'LPs': 0}
    pasillo_summary[p]['Unidades'] += r['Qty'] or 0
    pasillo_summary[p]['SKUs'].add(r['ItemCode'])
    pasillo_summary[p]['LPs'] += 1
for v in pasillo_summary.values():
    v['SKUs'] = len(v['SKUs'])
pasillo_list = sorted(pasillo_summary.values(), key=lambda x: x['Pasillo'])

# ── 6) RESUMEN POR CASA ────────────────────────────────────────────
cat_summary = {}
for r in inv_rows:
    c = r['Categoria'] or 'SIN CATEGORIA'
    if c not in cat_summary:
        cat_summary[c] = {'Categoria': c, 'Unidades': 0, 'SKUs': set(), 'LPs': 0}
    cat_summary[c]['Unidades'] += r['Qty'] or 0
    cat_summary[c]['SKUs'].add(r['ItemCode'])
    cat_summary[c]['LPs'] += 1
for v in cat_summary.values():
    v['SKUs'] = len(v['SKUs'])
cat_list = sorted(cat_summary.values(), key=lambda x: -x['Unidades'])

# ── 7) ANALISIS DE DANADO ─────────────────────────────────────────
print("Extrayendo analisis de dañado...")

# 7b) Flujo P34 mensual (antes del SP para evitar connection busy)
flujo_in = query(conn, """
SELECT FORMAT(TrxDate, 'yyyy-MM') as Mes, SUM(TrxQty) as Units
FROM LOGISTICA.HJ_TRANSACTIONS
WHERE WhsCode=? AND TargetLocID LIKE 'P34%' AND TrxDate >= ?
GROUP BY FORMAT(TrxDate, 'yyyy-MM') ORDER BY Mes
""", [BODEGA, f'{YEAR}-01-01'])

flujo_out = query(conn, """
SELECT FORMAT(TrxDate, 'yyyy-MM') as Mes, SUM(TrxQty) as Units
FROM LOGISTICA.HJ_TRANSACTIONS
WHERE WhsCode=? AND SourceLocID LIKE 'P34%' AND TrxDate >= ?
GROUP BY FORMAT(TrxDate, 'yyyy-MM') ORDER BY Mes
""", [BODEGA, f'{YEAR}-01-01'])

# 7c) P34 vencimiento
p34_exp = query(conn, """
SELECT
    CASE
        WHEN ExpirationDate < GETDATE() THEN 'VENCIDO'
        WHEN ExpirationDate < DATEADD(MONTH, 1, GETDATE()) THEN 'VENCE_30D'
        WHEN ExpirationDate < DATEADD(MONTH, 3, GETDATE()) THEN 'VENCE_90D'
        ELSE 'VIGENTE'
    END as Estado,
    SUM(Qty) as Units, COUNT(DISTINCT ItemCode) as SKUs
FROM LOGISTICA.HJ_Inventory
WHERE Warehouse=? AND Location LIKE 'P34%' AND Qty > 0
GROUP BY CASE
    WHEN ExpirationDate < GETDATE() THEN 'VENCIDO'
    WHEN ExpirationDate < DATEADD(MONTH, 1, GETDATE()) THEN 'VENCE_30D'
    WHEN ExpirationDate < DATEADD(MONTH, 3, GETDATE()) THEN 'VENCE_90D'
    ELSE 'VIGENTE'
END
""", [BODEGA])

# 7d) P34 por Status/BlockedCode (tipos de dano)
p34_tipos = query(conn, """
SELECT Status, BlockedCode,
    SUM(Qty) as Units, COUNT(*) as LPs, COUNT(DISTINCT ItemCode) as SKUs
FROM LOGISTICA.HJ_Inventory
WHERE Warehouse=? AND Location LIKE 'P34%' AND Qty > 0
GROUP BY Status, BlockedCode ORDER BY Units DESC
""", [BODEGA])

# 7e) P34 top SKUs
p34_top_skus = query(conn, """
SELECT TOP 15 i.ItemCode, i.ItemName, SUM(i.Qty) as Units, COUNT(*) as LPs,
       m.Class as Categoria, i.Status
FROM LOGISTICA.HJ_Inventory i
LEFT JOIN LOGISTICA.HJ_MasterData m ON i.Warehouse=m.WhsCode AND i.ItemCode=m.ItemCode
WHERE i.Warehouse=? AND i.Location LIKE 'P34%' AND i.Qty > 0
GROUP BY i.ItemCode, i.ItemName, m.Class, i.Status
ORDER BY Units DESC
""", [BODEGA])

print(f"  Flujo P34: {len(flujo_in)} meses ingreso, {len(flujo_out)} meses salida")
print(f"  P34 vencimiento: {len(p34_exp)} categorias")

# 8) VENCIMIENTO INVENTARIO GENERAL ─────────────────────────────────
print("Extrayendo vencimiento inventario general...")
venc_general = query(conn, """
SELECT
    CASE
        WHEN ExpirationDate IS NULL OR ExpirationDate <= '1900-01-02' THEN 'SIN_FECHA'
        WHEN ExpirationDate < GETDATE() THEN 'VENCIDO'
        WHEN ExpirationDate < DATEADD(MONTH, 1, GETDATE()) THEN 'VENCE_30D'
        WHEN ExpirationDate < DATEADD(MONTH, 3, GETDATE()) THEN 'VENCE_90D'
        ELSE 'VIGENTE'
    END as Estado,
    SUM(Qty) as Cajas, COUNT(DISTINCT ItemCode) as SKUs, COUNT(*) as LPs
FROM LOGISTICA.HJ_Inventory
WHERE Warehouse=? AND Qty > 0
GROUP BY CASE
    WHEN ExpirationDate IS NULL OR ExpirationDate <= '1900-01-02' THEN 'SIN_FECHA'
    WHEN ExpirationDate < GETDATE() THEN 'VENCIDO'
    WHEN ExpirationDate < DATEADD(MONTH, 1, GETDATE()) THEN 'VENCE_30D'
    WHEN ExpirationDate < DATEADD(MONTH, 3, GETDATE()) THEN 'VENCE_90D'
    ELSE 'VIGENTE'
END
""", [BODEGA])

venc_top = query(conn, """
SELECT TOP 20 i.ItemCode, i.ItemName, SUM(i.Qty) as Cajas,
       COUNT(*) as LPs, m.Class as Categoria,
       MIN(i.ExpirationDate) as FechaMin, MAX(i.ExpirationDate) as FechaMax
FROM LOGISTICA.HJ_Inventory i
LEFT JOIN LOGISTICA.HJ_MasterData m ON i.Warehouse=m.WhsCode AND i.ItemCode=m.ItemCode
WHERE i.Warehouse=? AND i.Qty > 0
  AND i.ExpirationDate IS NOT NULL AND i.ExpirationDate > '1900-01-02'
  AND i.ExpirationDate < GETDATE()
GROUP BY i.ItemCode, i.ItemName, m.Class
ORDER BY Cajas DESC
""", [BODEGA])
print(f"  Vencimiento general: {len(venc_general)} estados, {len(venc_top)} SKUs vencidos")

conn.close()

# 7a) Dañado recibido from SP (conexion separada para evitar cursor busy)
print("  Procesando SP recepciones para dañado...")
conn2 = connect()
cur_sp = conn2.cursor()
cur_sp.execute("EXEC LOGISTICA.sp_DataRecepciones @BodegaCodigo=?", [BODEGA_COD])
sp_cols = [c[0] for c in cur_sp.description]
sp_idx = {c: i for i, c in enumerate(sp_cols)}

dmg_mes = {}
dmg_prov = {}
dmg_sku = {}
dmg_total = 0.0
dmg_recibido = 0.0

for row in cur_sp.fetchall():
    fecha = row[sp_idx['FechaCita']]
    if not fecha or str(fecha)[:4] != str(YEAR):
        continue
    dan = float(row[sp_idx['Dañado']] or 0)
    rec = float(row[sp_idx['KorberRecibido']] or 0)
    prov = row[sp_idx['Proveedor']] or 'SIN PROVEEDOR'
    sku = row[sp_idx['SKU']] or ''
    desc = row[sp_idx['DescripcionSKU']] or ''
    cita = row[sp_idx['IdCita']]
    mes = str(fecha)[:7]

    if prov not in dmg_prov:
        dmg_prov[prov] = {'d': 0, 't': 0}
    dmg_prov[prov]['t'] += rec

    if mes not in dmg_mes:
        dmg_mes[mes] = {'d': 0, 'r': 0, 'ci': set()}
    dmg_mes[mes]['r'] += rec
    dmg_recibido += rec

    if dan > 0:
        dmg_total += dan
        dmg_mes[mes]['d'] += dan
        dmg_mes[mes]['ci'].add(cita)
        dmg_prov[prov]['d'] += dan
        if sku not in dmg_sku:
            dmg_sku[sku] = {'d': 0, 'ds': desc, 'ci': 0}
        dmg_sku[sku]['d'] += dan
        dmg_sku[sku]['ci'] += 1

conn2.close()
if dmg_recibido:
    print(f"  Dañado: {dmg_total:,.0f} cajas / {dmg_recibido:,.0f} recibidas ({dmg_total/dmg_recibido*100:.2f}%)")

# ── GENERAR JS ──────────────────────────────────────────────────────
print("Generando inventario_data.js...")
now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

totals = {
    'totalUnidades': sum(r['Qty'] or 0 for r in inv_rows),
    'totalSKUs': len(set(r['ItemCode'] for r in inv_rows)),
    'totalLPs': len(inv_rows),
    'totalPasillos': len(pasillo_summary),
    'p34Unidades': sum(r['Qty'] or 0 for r in p34_rows),
    'p34SKUs': len(set(r['ItemCode'] for r in p34_rows)),
    'khCitas': len(set(r['Cita'] for r in kh_rows)),
    'khLineas': len(kh_rows),
    'khCajasEnviadas': sum(r['CajasEnviadas'] or 0 for r in kh_rows),
    'fechaActualizacion': now,
}

data = {
    'totals': totals,
    'inventario': inv_rows,
    'kraftHeinz': kh_rows,
    'pasillo34': p34_rows,
    'cierre': cierre_rows,
    'resumenPasillo': pasillo_list,
    'resumenCategoria': cat_list,
}

def to_num(v):
    try: return float(v) if v else 0
    except: return 0

def compact_inv(rows):
    """Reduce payload: short keys, drop nulls, trim names"""
    out = []
    for r in rows:
        o = {
            'p': r['Pasillo'],
            'l': r['Location'],
            'c': r['ItemCode'],
            'n': (r['ItemName'] or '')[:50],
            'q': to_num(r['Qty']),
            's': r['Status'],
            'ca': r['Categoria'] or '',
        }
        if r['AvailableQty']: o['av'] = to_num(r['AvailableQty'])
        if r['BlockedQty']: o['bq'] = to_num(r['BlockedQty'])
        if r['BlockedCode']: o['bc'] = r['BlockedCode']
        if r['ExpirationDate']:
            ed = r['ExpirationDate']
            o['ex'] = ed.isoformat()[:10] if hasattr(ed, 'isoformat') else str(ed)[:10]
        if r['License']: o['lp'] = r['License']
        if r['LotNumber']: o['lt'] = r['LotNumber']
        out.append(o)
    return out

def compact_kh(rows):
    out = []
    for r in rows:
        o = {
            'ci': r['Cita'],
            'f': serial(r['FechaCita'])[:10] if r['FechaCita'] else '',
            'ra': r['Rampa'] or '',
            'co': r['Completada'] or 'N',
            'oc': r['OC'] or '',
            'sk': r['SKU'] or '',
            'pr': (r['Producto'] or '')[:50],
            'ce': float(r['CajasEnviadas'] or 0),
            'ux': int(r['UxC'] or 0),
            'pi': (r['Piloto'] or '')[:25],
            'pl': r['Placa'] or '',
            'ca': r['Categoria'] or '',
        }
        out.append(o)
    return out

def compact_p34(rows):
    out = []
    for r in rows:
        o = {
            'l': r['Location'],
            'c': r['ItemCode'],
            'n': (r['ItemName'] or '')[:50],
            'q': to_num(r['Qty']),
            's': r['Status'],
            'ca': r['Categoria'] or '',
        }
        if r['BlockedCode']: o['bc'] = r['BlockedCode']
        if r['LotNumber']: o['lt'] = r['LotNumber']
        if r['ExpirationDate']:
            ed = r['ExpirationDate']
            o['ex'] = ed.isoformat()[:10] if hasattr(ed, 'isoformat') else str(ed)[:10]
        if r['License']: o['lp'] = r['License']
        out.append(o)
    return out

# ── Build analisis de danado compact data ──
analisis_por_mes = []
for mes in sorted(dmg_mes.keys()):
    d = dmg_mes[mes]
    pct = d['d'] / d['r'] * 100 if d['r'] else 0
    analisis_por_mes.append({'m': mes, 'd': round(d['d']), 'r': round(d['r']), 'p': round(pct, 2), 'ci': len(d['ci'])})

analisis_top_prov = []
for prov, d in sorted(dmg_prov.items(), key=lambda x: -x[1]['d'])[:15]:
    if d['d'] > 0:
        pct = d['d'] / d['t'] * 100 if d['t'] else 0
        analisis_top_prov.append({'pr': prov[:40], 'd': round(d['d']), 't': round(d['t']), 'p': round(pct, 2)})

analisis_top_sku = []
for sku, d in sorted(dmg_sku.items(), key=lambda x: -x[1]['d'])[:15]:
    analisis_top_sku.append({'sk': sku, 'ds': d['ds'][:40], 'd': round(d['d']), 'ci': d['ci']})

fi_map = {r['Mes']: to_num(r['Units']) for r in flujo_in}
fo_map = {r['Mes']: to_num(r['Units']) for r in flujo_out}
all_meses_flujo = sorted(set(list(fi_map.keys()) + list(fo_map.keys())))
analisis_flujo = [{'m': m, 'i': round(fi_map.get(m, 0)), 'o': round(fo_map.get(m, 0))} for m in all_meses_flujo]

analisis_exp = [{'e': r['Estado'], 'u': round(to_num(r['Units'])), 'sk': r['SKUs']} for r in p34_exp]

und_A = sum(to_num(r['Units']) for r in p34_tipos if r['Status'] == 'A')
und_H = sum(to_num(r['Units']) for r in p34_tipos if r['Status'] == 'H')

analisis_bc = [{'s': r['Status'], 'bc': r['BlockedCode'] or '', 'u': round(to_num(r['Units'])), 'lp': r['LPs'], 'sk': r['SKUs']} for r in p34_tipos]

analisis_p34_top = []
for r in p34_top_skus:
    analisis_p34_top.append({
        'c': r['ItemCode'], 'n': (r['ItemName'] or '')[:40], 'u': round(to_num(r['Units'])),
        'lp': r['LPs'], 'ca': r['Categoria'] or '', 's': r['Status']
    })

venc_gen_data = [{'e': r['Estado'], 'u': round(to_num(r['Cajas'])), 'sk': r['SKUs'], 'lp': r['LPs']} for r in venc_general]
venc_top_data = []
for r in venc_top:
    o = {'c': r['ItemCode'], 'n': (r['ItemName'] or '')[:40], 'u': round(to_num(r['Cajas'])),
         'lp': r['LPs'], 'ca': r['Categoria'] or ''}
    if r['FechaMin']:
        o['fm'] = r['FechaMin'].isoformat()[:10] if hasattr(r['FechaMin'], 'isoformat') else str(r['FechaMin'])[:10]
    if r['FechaMax']:
        o['fx'] = r['FechaMax'].isoformat()[:10] if hasattr(r['FechaMax'], 'isoformat') else str(r['FechaMax'])[:10]
    venc_top_data.append(o)

analisis_data = {
    'porMes': analisis_por_mes,
    'topProv': analisis_top_prov,
    'topSKU': analisis_top_sku,
    'flujo': analisis_flujo,
    'exp': analisis_exp,
    'tipos': {'corr': round(und_A), 'noR': round(und_H)},
    'bc': analisis_bc,
    'topP34': analisis_p34_top,
    'tot': {
        'td': round(dmg_total),
        'tr': round(dmg_recibido),
        'pct': round(dmg_total / dmg_recibido * 100, 2) if dmg_recibido else 0
    },
    'vencGen': venc_gen_data,
    'vencTop': venc_top_data,
}

totals['totalDanadoCajas'] = round(dmg_total)
totals['totalRecibidoCajas'] = round(dmg_recibido)
totals['pctDanado'] = round(dmg_total / dmg_recibido * 100, 2) if dmg_recibido else 0
venc_dict = {r['e']: r['u'] for r in venc_gen_data}
totals['totalVencido'] = venc_dict.get('VENCIDO', 0)
totals['totalVence30'] = venc_dict.get('VENCE_30D', 0)
totals['totalVence90'] = venc_dict.get('VENCE_90D', 0)

data_compact = {
    'totals': totals,
    'inventario': compact_inv(inv_rows),
    'kraftHeinz': compact_kh(kh_rows),
    'pasillo34': compact_p34(p34_rows),
    'resumenPasillo': pasillo_list,
    'resumenCategoria': cat_list,
    'analisisDanado': analisis_data,
}

js_path = os.path.join(DIR, 'inventario_data.js')
with open(js_path, 'w', encoding='utf-8') as f:
    f.write('// Auto-generated ' + now + '\n')
    f.write('const INVENTARIO_DATA = ')
    json.dump(data_compact, f, ensure_ascii=False, default=serial, separators=(',',':'))
    f.write(';\n')

size_mb = os.path.getsize(js_path) / 1024 / 1024
print(f"  inventario_data.js: {size_mb:.1f} MB")

# ── GENERAR EXCEL ───────────────────────────────────────────────────
print("Generando INVENTARIO_COMPLETO.xlsx...")
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()

    hdr_font = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    hdr_fill = PatternFill(start_color='0D2A6B', end_color='0D2A6B', fill_type='solid')
    hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin = Side(style='thin', color='D0D0D0')
    cell_border = Border(bottom=thin)

    def write_sheet(ws, title, headers, rows_data, col_widths=None):
        ws.title = title
        ws.freeze_panes = 'A2'
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=ci, value=h)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = hdr_align
        for ri, row in enumerate(rows_data, 2):
            for ci, val in enumerate(row, 1):
                c = ws.cell(row=ri, column=ci, value=val)
                c.border = cell_border
                if isinstance(val, (int, float)) and val is not None:
                    c.number_format = '#,##0'
        if col_widths:
            for ci, w in enumerate(col_widths, 1):
                ws.column_dimensions[get_column_letter(ci)].width = w
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows_data)+1}"

    # Hoja 1: Inventario
    inv_headers = ['Pasillo','Ubicacion','Codigo','Producto','Cantidad','Disponible','Bloqueado',
                   'No Disponible','Status','Cod Bloqueo','Lote','Vencimiento','FIFO','LP','UM','Categoria']
    inv_data = []
    for r in inv_rows:
        inv_data.append([
            r['Pasillo'], r['Location'], r['ItemCode'], r['ItemName'],
            r['Qty'], r['AvailableQty'], r['BlockedQty'], r['UnavailableQty'],
            r['Status'], r['BlockedCode'], r['LotNumber'],
            r['ExpirationDate'].isoformat() if r['ExpirationDate'] else None,
            r['FifoDate'].isoformat() if r['FifoDate'] else None,
            r['License'], r['UM'], r['Categoria']
        ])
    ws1 = wb.active
    write_sheet(ws1, 'Inventario', inv_headers, inv_data,
                [10,18,16,45,12,12,12,12,8,12,15,12,12,20,6,10])

    # Hoja 2: Kraft Heinz
    kh_headers = ['Cita','Fecha','Proveedor','Horario','Rampa','Piloto','Placa',
                  'Completada','OC','SKU','Producto','Cajas Enviadas','UxC']
    kh_data = []
    for r in kh_rows:
        kh_data.append([
            r['Cita'],
            r['FechaCita'].isoformat() if r['FechaCita'] else None,
            r['Proveedor'], str(r['HorarioCita'] or ''), r['Rampa'],
            r['Piloto'], r['Placa'], r['Completada'],
            r['OC'], r['SKU'], r['Producto'],
            r['CajasEnviadas'], r['UxC']
        ])
    ws2 = wb.create_sheet()
    write_sheet(ws2, 'Kraft Heinz', kh_headers, kh_data,
                [10,12,20,10,10,25,12,10,15,16,45,14,8])

    # Hoja 3: Pasillo 34
    p34_headers = ['Ubicacion','Codigo','Producto','Cantidad','Status','Cod Bloqueo','Lote','Vencimiento','LP','Categoria']
    p34_data = []
    for r in p34_rows:
        p34_data.append([
            r['Location'], r['ItemCode'], r['ItemName'], r['Qty'],
            r['Status'], r['BlockedCode'], r['LotNumber'],
            r['ExpirationDate'].isoformat() if r['ExpirationDate'] else None,
            r['License'], r['Categoria']
        ])
    ws3 = wb.create_sheet()
    write_sheet(ws3, 'Pasillo 34 Dañado', p34_headers, p34_data,
                [18,16,45,12,8,12,15,12,20,10])

    # Hoja 4: Resumen Pasillo
    ws4 = wb.create_sheet()
    write_sheet(ws4, 'Resumen Pasillo', ['Pasillo','Unidades','SKUs','LPs'],
                [[r['Pasillo'], r['Unidades'], r['SKUs'], r['LPs']] for r in pasillo_list],
                [12,15,10,10])

    # Hoja 5: Resumen Categoria
    ws5 = wb.create_sheet()
    write_sheet(ws5, 'Resumen Categoria', ['Categoria','Unidades','SKUs','LPs'],
                [[r['Categoria'], r['Unidades'], r['SKUs'], r['LPs']] for r in cat_list],
                [15,15,10,10])

    xlsx_path = os.path.join(DIR, 'INVENTARIO_COMPLETO.xlsx')
    wb.save(xlsx_path)
    print(f"  INVENTARIO_COMPLETO.xlsx: {os.path.getsize(xlsx_path)/1024:.0f} KB")

except ImportError:
    print("  ADVERTENCIA: openpyxl no instalado, no se genero Excel")
    print("  Instalar con: pip install openpyxl")

print("\n=== RESUMEN ===")
print(f"  Inventario: {totals['totalUnidades']:,.0f} unidades / {totals['totalSKUs']} SKUs / {totals['totalLPs']} LPs")
print(f"  Pasillo 34: {totals['p34Unidades']:,.0f} unidades / {totals['p34SKUs']} SKUs")
print(f"  Kraft Heinz: {totals['khCitas']} citas / {totals['khLineas']} lineas / {totals['khCajasEnviadas']:,.0f} cajas")
print(f"  Actualizado: {now}")
print("\nListo.")
