"""
CONTROL DE BLOQUES FEFO - Bodega 1110
Sistema de N bloques por SKU con cola de prioridad automatica.
Cada posicion PISO-Bx es una ubicacion en Korber con varios SKUs.
El sistema auto-descubre todos los SKUs, agrupa por posicion,
ordena por fecha de vencimiento y asigna prioridades de despacho:
  - Prioridad #1 (verde) = despachar primero (fecha mas vieja)
  - Prioridad #2..N-1 (azul) = en espera, despachar en orden
  - Ultimo con producto (naranja) = recepcion de producto nuevo
  - Vacios (gris) = disponible para proxima descarga
Genera dashboard HTML + rotulos imprimibles A4.
"""
import pyodbc, json, sys, os, datetime
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')
DIR = os.path.dirname(os.path.abspath(__file__))

# ══════════════════════════════════════════════════════════════════════
# CONFIGURACION
# ══════════════════════════════════════════════════════════════════════
SQL_SERVER = "172.16.5.43"
SQL_DB     = "IN_E"
SQL_DRIVER = "ODBC Driver 17 for SQL Server"
BODEGA     = "1110A001"

# Las 4 posiciones del area de bloques a piso
POSICIONES_PISO = ['PISO-B1', 'PISO-B2', 'PISO-B3', 'PISO-B4']

# ══════════════════════════════════════════════════════════════════════

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
    return [dict(zip(cols, list(r))) for r in cur.fetchall()]

def serial(obj):
    if isinstance(obj, (datetime.date, datetime.datetime)):
        return obj.isoformat()
    if isinstance(obj, float) and obj != obj:
        return None
    return str(obj)

def safe_date(d):
    if isinstance(d, (datetime.date, datetime.datetime)) and d > datetime.date(2000, 1, 1):
        return d
    return None


def fetch_inventory(conn):
    ph = ",".join(["?"] * len(POSICIONES_PISO))
    return query(conn, f"""
        SELECT i.Location, i.ItemCode, i.ItemName,
               i.Qty, i.ExpirationDate, i.FifoDate,
               i.License, i.LotNumber, i.Status,
               m.Class AS Casa
        FROM LOGISTICA.HJ_Inventory i
        LEFT JOIN LOGISTICA.HJ_MasterData m
            ON i.Warehouse = m.WhsCode AND i.ItemCode = m.ItemCode
        WHERE i.Warehouse = ?
          AND i.Qty > 0
          AND i.Location IN ({ph})
        ORDER BY i.ItemCode, i.ExpirationDate
    """, [BODEGA] + POSICIONES_PISO)


def block_stats(inv_list):
    if not inv_list:
        return {'qty': 0, 'fecha_min': None, 'fecha_max': None, 'lps': 0, 'detalle': []}
    total_qty = sum(r['Qty'] or 0 for r in inv_list)
    fechas = [safe_date(r['ExpirationDate']) for r in inv_list if safe_date(r.get('ExpirationDate'))]
    lps = len(set(r.get('License', '') for r in inv_list if r.get('License')))
    detalle = [{
        'location': r['Location'], 'qty': r['Qty'],
        'expDate': r['ExpirationDate'].isoformat() if isinstance(r.get('ExpirationDate'), (datetime.date, datetime.datetime)) else '',
        'lote': r.get('LotNumber', ''), 'lp': r.get('License', ''), 'status': r.get('Status', '')
    } for r in inv_list]
    return {
        'qty': total_qty,
        'fecha_min': min(fechas).isoformat() if fechas else None,
        'fecha_max': max(fechas).isoformat() if fechas else None,
        'lps': lps, 'detalle': detalle
    }


def build_sku_blocks(rows):
    """Auto-descubre SKUs y arma bloques por posicion."""
    # Agrupar por SKU -> posicion
    by_sku = defaultdict(lambda: defaultdict(list))
    sku_info = {}
    for r in rows:
        sku = r['ItemCode']
        loc = r['Location']
        by_sku[sku][loc].append(r)
        if sku not in sku_info:
            sku_info[sku] = {'nombre': r.get('ItemName', ''), 'casa': r.get('Casa', '')}

    results = []
    for sku in sorted(by_sku.keys()):
        locs = by_sku[sku]
        info = sku_info[sku]

        # Cada posicion donde este SKU tiene inventario = un bloque
        bloques = []
        for loc in POSICIONES_PISO:
            inv = locs.get(loc, [])
            stats = block_stats(inv)
            bloques.append({
                'id': loc.replace('PISO-', ''),  # B1, B2, B3, B4
                'posicion': loc,
                'posiciones': [loc],
                'stats': stats,
            })

        # Ordenar: bloques con inventario por fecha mas vieja, vacios al final
        con_inv = [b for b in bloques if b['stats']['qty'] > 0]
        sin_inv = [b for b in bloques if b['stats']['qty'] == 0]
        con_inv.sort(key=lambda b: b['stats']['fecha_min'] or '9999')
        ordered = con_inv + sin_inv

        # Asignar roles por prioridad
        for i, b in enumerate(ordered):
            if b['stats']['qty'] == 0:
                b['rol'] = 'DISPONIBLE'
                b['prioridad'] = None
            elif i == len(con_inv) - 1 and len(con_inv) > 1:
                # Ultimo con inventario = recepcion (tiene la fecha mas nueva)
                b['rol'] = 'RECEPCION'
                b['prioridad'] = len(con_inv)
            elif len(con_inv) == 1:
                # Solo un bloque con inventario = despacho
                b['rol'] = 'DESPACHO'
                b['prioridad'] = 1
            else:
                b['rol'] = 'DESPACHO'
                b['prioridad'] = i + 1

        # Si todos vacios, marcar el primero como recepcion
        if all(b['stats']['qty'] == 0 for b in ordered) and ordered:
            ordered[0]['rol'] = 'RECEPCION'

        # Solo incluir SKUs que tengan inventario en al menos 1 piso
        total_qty = sum(b['stats']['qty'] for b in ordered)
        if total_qty == 0:
            continue

        hay_vacio = any(b['stats']['qty'] == 0 for b in ordered)
        n_con = len(con_inv)

        results.append({
            'sku': sku,
            'nombre': info['nombre'],
            'casa': info['casa'],
            'total_bloques': len(POSICIONES_PISO),
            'con_producto': n_con,
            'bloques': ordered,
            'necesita_espacio': not hay_vacio and n_con == len(POSICIONES_PISO),
        })

    return results


def generate_html(sku_data, timestamp):
    data_json = json.dumps({
        'skus': sku_data,
        'posiciones': POSICIONES_PISO,
        'timestamp': timestamp,
    }, default=serial, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Control Bloques FEFO - Bodega 1110</title>
<style>
:root {{
  --navy: #00306C; --blue: #2f86e0;
  --bg: #0a1628; --card: #111d33; --card-border: #1a2d4a;
  --text: #e8ecf1; --text-dim: #8899aa;
  --green: #22c55e; --green-bg: #0a2e1a; --green-border: #166534;
  --sky: #38bdf8; --sky-bg: #0a1e2e; --sky-border: #0369a1;
  --orange: #f59e0b; --orange-bg: #2e1f0a; --orange-border: #92400e;
  --red: #ef4444; --red-bg: #2e0a0a; --red-border: #991b1b;
  --gray: #6b7280; --gray-bg: #1a1d24; --gray-border: #374151;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }}

.header {{
  background: linear-gradient(135deg, var(--navy) 0%, #001a3a 100%);
  padding: 20px 32px; display: flex; align-items: center; gap: 16px;
  border-bottom: 2px solid var(--blue);
}}
.header-logo {{
  width: 48px; height: 48px; border-radius: 12px;
  background: linear-gradient(135deg, var(--navy), var(--blue));
  display: flex; align-items: center; justify-content: center;
  font-size: 28px; font-weight: 800; color: #fff;
}}
.header h1 {{ font-size: 20px; font-weight: 700; letter-spacing: 1px; }}
.header .sub {{ font-size: 12px; color: var(--text-dim); }}
.header .ts {{ font-size: 11px; color: var(--text-dim); margin-left: auto; }}

.tabs {{ display: flex; background: var(--navy); border-bottom: 1px solid var(--card-border); padding: 0 32px; }}
.tab {{ padding: 12px 24px; cursor: pointer; font-size: 13px; font-weight: 600; color: var(--text-dim); border-bottom: 3px solid transparent; transition: 0.2s; }}
.tab:hover {{ color: var(--text); }}
.tab.active {{ color: var(--blue); border-bottom-color: var(--blue); }}

.content {{ padding: 24px 32px; max-width: 1500px; margin: 0 auto; }}
.section {{ display: none; }}
.section.active {{ display: block; }}

.kpis {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 14px; margin-bottom: 24px; }}
.kpi {{ background: var(--card); border: 1px solid var(--card-border); border-radius: 12px; padding: 18px; text-align: center; }}
.kpi .value {{ font-size: 34px; font-weight: 800; }}
.kpi .label {{ font-size: 11px; color: var(--text-dim); margin-top: 4px; text-transform: uppercase; letter-spacing: 1px; }}
.kpi.green .value {{ color: var(--green); }}
.kpi.orange .value {{ color: var(--orange); }}
.kpi.red .value {{ color: var(--red); }}
.kpi.blue .value {{ color: var(--blue); }}

/* SKU Card */
.sku-card {{ background: var(--card); border: 1px solid var(--card-border); border-radius: 12px; margin-bottom: 14px; overflow: hidden; }}
.sku-header {{ padding: 14px 20px; display: flex; align-items: center; gap: 12px; border-bottom: 1px solid var(--card-border); cursor: pointer; flex-wrap: wrap; }}
.sku-header:hover {{ background: rgba(47,134,224,0.05); }}
.sku-code {{ font-size: 13px; font-weight: 700; font-family: 'Consolas', monospace; color: var(--blue); min-width: 120px; }}
.sku-name {{ font-size: 13px; flex: 1; min-width: 150px; }}
.sku-casa {{ font-size: 11px; color: var(--text-dim); background: var(--navy); padding: 2px 10px; border-radius: 99px; }}
.sku-piso-dots {{ display: flex; gap: 4px; }}
.piso-dot {{
  width: 28px; height: 20px; border-radius: 4px; font-size: 9px; font-weight: 700;
  display: flex; align-items: center; justify-content: center;
}}
.piso-dot.g {{ background: var(--green); color: #fff; }}
.piso-dot.b {{ background: var(--sky); color: #000; }}
.piso-dot.o {{ background: var(--orange); color: #000; }}
.piso-dot.gr {{ background: var(--gray-bg); color: var(--gray); border: 1px dashed var(--gray-border); }}
.sku-badge {{ font-size: 11px; font-weight: 700; padding: 3px 10px; border-radius: 6px; }}
.sku-badge.ok {{ background: var(--green-bg); color: var(--green); border: 1px solid var(--green-border); }}
.sku-badge.warn {{ background: var(--orange-bg); color: var(--orange); border: 1px solid var(--orange-border); }}
.sku-badge.full {{ background: var(--red-bg); color: var(--red); border: 1px solid var(--red-border); animation: pulse 1.5s infinite; }}
@keyframes pulse {{ 0%,100%{{opacity:1}} 50%{{opacity:0.6}} }}

.sku-body {{ display: none; padding: 16px 20px; }}
.sku-body.open {{ display: block; }}

/* Block Grid */
.blocks-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }}
@media (max-width: 900px) {{ .blocks-grid {{ grid-template-columns: repeat(2, 1fr); }} }}

.blk {{ border-radius: 10px; padding: 14px; position: relative; }}
.blk.despacho-1 {{ background: var(--green-bg); border: 2px solid var(--green-border); }}
.blk.despacho-n {{ background: var(--sky-bg); border: 2px solid var(--sky-border); }}
.blk.recepcion {{ background: var(--orange-bg); border: 2px solid var(--orange-border); }}
.blk.disponible {{ background: var(--gray-bg); border: 2px dashed var(--gray-border); }}

.blk-role {{ font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px; }}
.blk.despacho-1 .blk-role {{ color: var(--green); }}
.blk.despacho-n .blk-role {{ color: var(--sky); }}
.blk.recepcion .blk-role {{ color: var(--orange); }}
.blk.disponible .blk-role {{ color: var(--gray); }}

.blk-prio {{
  position: absolute; top: 10px; right: 10px;
  width: 28px; height: 28px; border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  font-size: 13px; font-weight: 900;
}}
.blk.despacho-1 .blk-prio {{ background: var(--green); color: #fff; }}
.blk.despacho-n .blk-prio {{ background: var(--sky); color: #000; }}
.blk.recepcion .blk-prio {{ background: var(--orange); color: #000; }}

.blk-id {{ font-size: 14px; font-weight: 700; color: var(--text); margin-bottom: 8px; }}
.blk-stat {{ display: flex; justify-content: space-between; padding: 3px 0; font-size: 12px; border-bottom: 1px solid rgba(255,255,255,0.05); }}
.blk-stat .lbl {{ color: var(--text-dim); }}
.blk-stat .val {{ font-weight: 600; font-family: 'Consolas', monospace; }}

.blk-detail {{ margin-top: 8px; max-height: 200px; overflow-y: auto; }}
.blk-detail table {{ width: 100%; border-collapse: collapse; font-size: 10px; }}
.blk-detail th {{ text-align: left; color: var(--text-dim); padding: 3px 4px; border-bottom: 1px solid var(--card-border); position: sticky; top: 0; background: inherit; }}
.blk-detail td {{ padding: 3px 4px; border-bottom: 1px solid rgba(255,255,255,0.03); font-family: 'Consolas', monospace; }}

/* Legend */
.legend {{ display: flex; gap: 18px; margin-bottom: 20px; flex-wrap: wrap; font-size: 12px; }}
.legend-item {{ display: flex; align-items: center; gap: 6px; }}
.legend-dot {{ width: 14px; height: 14px; border-radius: 4px; }}
.legend-dot.g {{ background: var(--green); }}
.legend-dot.b {{ background: var(--sky); }}
.legend-dot.o {{ background: var(--orange); }}
.legend-dot.gr {{ background: var(--gray); border: 1px dashed var(--gray-border); }}

/* Filter */
.filter-bar {{ display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; align-items: center; }}
.filter-bar input, .filter-bar select {{
  background: var(--card); border: 1px solid var(--card-border); border-radius: 8px;
  padding: 8px 14px; color: var(--text); font-size: 13px; outline: none;
}}
.filter-bar input:focus {{ border-color: var(--blue); }}
.btn {{ padding: 10px 24px; border: none; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 700; }}
.btn-primary {{ background: var(--blue); color: #fff; }}
.btn-primary:hover {{ background: #1a6fc4; }}
.empty-state {{ text-align: center; padding: 60px 20px; color: var(--text-dim); }}

/* ═══ PRINT ═══ */
.print-section {{ display: none; }}
.sign-page {{
  width: 277mm; height: 190mm;
  page-break-after: always; padding: 12mm;
  border: 3px solid #000; display: flex; flex-direction: column;
  font-family: 'Arial Black', 'Segoe UI', sans-serif;
  background: #fff; color: #000;
}}
.sign-page:last-child {{ page-break-after: avoid; }}
.sign-role-bar {{
  padding: 14px 20px; text-align: center;
  font-size: 36px; font-weight: 900; letter-spacing: 3px;
  border-radius: 12px; margin-bottom: 8px;
}}
.sign-role-bar.despacho {{ background: #22c55e; color: #fff; }}
.sign-role-bar.espera {{ background: #38bdf8; color: #000; }}
.sign-role-bar.recepcion {{ background: #f59e0b; color: #000; }}
.sign-prio-circle {{
  width: 65px; height: 65px; border-radius: 50%; border: 4px solid #000;
  display: flex; align-items: center; justify-content: center;
  font-size: 36px; font-weight: 900; margin: 0 auto 6px;
}}
.sign-prio-circle.despacho {{ background: #22c55e; color: #fff; }}
.sign-prio-circle.espera {{ background: #38bdf8; color: #000; }}
.sign-prio-circle.recepcion {{ background: #f59e0b; color: #000; }}
.sign-sku-row {{ display: flex; align-items: baseline; gap: 14px; margin-bottom: 4px; padding: 0 8px; }}
.sign-sku {{ font-size: 24px; font-weight: 900; font-family: 'Consolas', monospace; }}
.sign-name {{ font-size: 18px; font-weight: 700; }}
.sign-block-id {{
  font-size: 30px; font-weight: 900; text-align: center; padding: 6px; margin: 4px 0;
  border: 4px solid #000; border-radius: 12px; background: #f0f0f0;
}}
.sign-dates {{
  display: flex; gap: 14px; justify-content: center; padding: 8px; margin: 6px 0;
  border: 2px dashed #666; border-radius: 8px; font-size: 17px;
}}
.sign-dates .dt-label {{ color: #666; }}
.sign-dates .dt-value {{ font-weight: 900; font-family: 'Consolas', monospace; }}
.sign-warning {{
  flex: 1; display: flex; align-items: center; justify-content: center;
  text-align: center; font-size: 20px; font-weight: 800; padding: 10px; margin-top: 6px;
  border-radius: 12px; background: #fee2e2; border: 3px solid #ef4444; color: #991b1b;
}}
.sign-footer {{
  display: flex; justify-content: space-between; font-size: 10px; color: #999;
  margin-top: auto; padding-top: 6px; border-top: 1px solid #ddd;
}}
.protocol-page {{
  width: 277mm; min-height: 190mm; page-break-after: always; padding: 15mm;
  font-family: 'Segoe UI', sans-serif; background: #fff; color: #000;
  font-size: 14px; line-height: 1.6;
}}
.protocol-page h2 {{ font-size: 22px; border-bottom: 3px solid #00306C; padding-bottom: 8px; margin-bottom: 16px; color: #00306C; }}
.protocol-page h3 {{ font-size: 16px; margin-top: 16px; color: #00306C; }}
.protocol-page .rule {{ background: #fef3c7; border-left: 4px solid #f59e0b; padding: 10px 14px; margin: 10px 0; border-radius: 4px; }}
.protocol-page .danger {{ background: #fee2e2; border-left: 4px solid #ef4444; padding: 10px 14px; margin: 10px 0; border-radius: 4px; font-weight: 700; }}
.protocol-page .ok {{ background: #dcfce7; border-left: 4px solid #22c55e; padding: 10px 14px; margin: 10px 0; border-radius: 4px; }}

@media print {{
  body {{ background: #fff !important; }}
  .header, .tabs, .content {{ display: none !important; }}
  .print-section {{ display: block !important; }}
  @page {{ size: A4 landscape; margin: 5mm; }}
}}
.preview-signs {{ display: flex; flex-direction: column; gap: 24px; align-items: center; }}
.preview-signs .sign-page, .preview-signs .protocol-page {{
  transform: scale(0.5); transform-origin: top center;
  box-shadow: 0 8px 32px rgba(0,0,0,0.5); margin-bottom: -140px;
}}
</style>
</head>
<body>

<div class="header">
  <div class="header-logo">A</div>
  <div>
    <h1>CONTROL DE BLOQUES FEFO</h1>
    <div class="sub">BODEGA 1110 &middot; PISO-B1 / B2 / B3 / B4 &middot; COLA DE PRIORIDAD</div>
  </div>
  <div class="ts" id="hdr-ts"></div>
</div>
<div class="tabs">
  <div class="tab active" data-tab="dashboard">Dashboard</div>
  <div class="tab" data-tab="signs">Rotulos</div>
  <div class="tab" data-tab="protocol">Protocolo</div>
</div>
<div class="content">
  <div class="section active" id="sec-dashboard"></div>
  <div class="section" id="sec-signs"></div>
  <div class="section" id="sec-protocol">
    <div style="background:var(--card);border:1px solid var(--card-border);border-radius:12px;padding:24px;max-width:900px;line-height:1.8;font-size:14px;">
      <h2 style="color:var(--blue);margin-bottom:16px;font-size:20px;">PROTOCOLO - BLOQUES FEFO (PISO B1-B4)</h2>

      <h3 style="color:var(--green);margin-top:20px;">REGLA PRINCIPAL</h3>
      <div style="background:var(--green-bg);border:1px solid var(--green-border);padding:14px;border-radius:8px;margin:8px 0;font-size:15px;font-weight:700;">
        Cada codigo puede estar en varios PISOS. Siempre despachar del PISO con prioridad #1 (verde = fecha mas vieja). Producto nuevo SOLO al PISO naranja o gris.
      </div>

      <h3 style="color:var(--blue);margin-top:20px;">COLORES</h3>
      <div style="margin:12px 0;">
        <div style="display:flex;gap:10px;align-items:center;margin:8px 0;">
          <span style="background:var(--green);width:24px;height:24px;border-radius:6px;display:inline-block;"></span>
          <span><strong>VERDE #1</strong> &mdash; despachar de aqui PRIMERO</span>
        </div>
        <div style="display:flex;gap:10px;align-items:center;margin:8px 0;">
          <span style="background:var(--sky);width:24px;height:24px;border-radius:6px;display:inline-block;"></span>
          <span><strong>AZUL #2, #3</strong> &mdash; en espera, despachar cuando el anterior se vacie</span>
        </div>
        <div style="display:flex;gap:10px;align-items:center;margin:8px 0;">
          <span style="background:var(--orange);width:24px;height:24px;border-radius:6px;display:inline-block;"></span>
          <span><strong>NARANJA</strong> &mdash; recepcion, producto nuevo VA AQUI</span>
        </div>
        <div style="display:flex;gap:10px;align-items:center;margin:8px 0;">
          <span style="background:var(--gray);width:24px;height:24px;border-radius:6px;border:1px dashed var(--gray-border);display:inline-block;"></span>
          <span><strong>GRIS</strong> &mdash; disponible para proxima descarga</span>
        </div>
      </div>

      <h3 style="color:var(--blue);margin-top:20px;">FLUJO</h3>
      <div style="margin:12px 0;">
        <div style="display:flex;gap:8px;margin:8px 0;"><span style="background:var(--blue);color:#fff;width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:800;flex-shrink:0;">1</span><span>Descarga nueva &rarr; colocar en el PISO NARANJA o un PISO GRIS</span></div>
        <div style="display:flex;gap:8px;margin:8px 0;"><span style="background:var(--blue);color:#fff;width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:800;flex-shrink:0;">2</span><span>Picking/despacho &rarr; tomar del PISO VERDE (#1)</span></div>
        <div style="display:flex;gap:8px;margin:8px 0;"><span style="background:var(--blue);color:#fff;width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:800;flex-shrink:0;">3</span><span>PISO #1 se vacia &rarr; #2 sube a #1, #3 a #2, etc.</span></div>
        <div style="display:flex;gap:8px;margin:8px 0;"><span style="background:var(--blue);color:#fff;width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:800;flex-shrink:0;">4</span><span>Reimprimir rotulos con las nuevas prioridades</span></div>
      </div>

      <h3 style="color:var(--red);margin-top:20px;">PROHIBIDO</h3>
      <div style="background:var(--red-bg);border:1px solid var(--red-border);padding:14px;border-radius:8px;margin:8px 0;">
        <div style="margin:4px 0;">&#x26D4; Despachar de un PISO #2+ mientras el #1 tenga existencia de ese codigo</div>
        <div style="margin:4px 0;">&#x26D4; Colocar producto nuevo en un PISO verde o azul</div>
        <div style="margin:4px 0;">&#x26D4; Mezclar fechas distintas del mismo codigo en un solo PISO</div>
      </div>
    </div>
  </div>
</div>

<div class="print-section" id="print-area"></div>

<script>
const DATA = {data_json};

document.querySelectorAll('.tab').forEach(t => {{
  t.addEventListener('click', () => {{
    document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.section').forEach(x => x.classList.remove('active'));
    t.classList.add('active');
    document.getElementById('sec-' + t.dataset.tab).classList.add('active');
  }});
}});

function fmtDate(d) {{
  if (!d) return '-';
  const dt = new Date(d); if (isNaN(dt)) return d;
  const M = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];
  return dt.getDate().toString().padStart(2,'0')+'/'+M[dt.getMonth()]+'/'+dt.getFullYear();
}}
function fmtQty(q) {{ return (q||0).toLocaleString('es-GT'); }}
function daysTo(d) {{
  if (!d) return null;
  const dt = new Date(d); if (isNaN(dt)) return null;
  const t = new Date(); t.setHours(0,0,0,0); dt.setHours(0,0,0,0);
  return Math.round((dt-t)/864e5);
}}

document.getElementById('hdr-ts').textContent = 'Generado: ' + fmtDate(DATA.timestamp);

function renderDashboard() {{
  const el = document.getElementById('sec-dashboard');
  const skus = DATA.skus;
  if (!skus.length) {{
    el.innerHTML = '<div class="empty-state"><h3>Sin inventario en PISO-B1 a B4</h3></div>';
    return;
  }}

  const totalSku = skus.length;
  const totalBloques = skus.reduce((s,sk) => s + sk.con_producto, 0);
  const totalQty = skus.reduce((s,sk) => s + sk.bloques.reduce((ss,b) => ss + (b.stats.qty||0), 0), 0);
  const sinEspacio = skus.filter(sk => sk.necesita_espacio).length;
  const enVarios = skus.filter(sk => sk.con_producto > 1).length;

  el.innerHTML = `<div class="kpis">
    <div class="kpi blue"><div class="value">${{totalSku}}</div><div class="label">SKUs en Pisos</div></div>
    <div class="kpi blue"><div class="value">${{totalBloques}}</div><div class="label">Bloques Activos</div></div>
    <div class="kpi blue"><div class="value">${{fmtQty(totalQty)}}</div><div class="label">Unidades</div></div>
    <div class="kpi orange"><div class="value">${{enVarios}}</div><div class="label">En 2+ Pisos</div></div>
    <div class="kpi ${{sinEspacio?'red':'green'}}"><div class="value">${{sinEspacio}}</div><div class="label">4 Pisos Llenos</div></div>
  </div>
  <div class="legend">
    <div class="legend-item"><div class="legend-dot g"></div> #1 Despachar</div>
    <div class="legend-item"><div class="legend-dot b"></div> #2+ En espera</div>
    <div class="legend-item"><div class="legend-dot o"></div> Recepcion</div>
    <div class="legend-item"><div class="legend-dot gr"></div> Disponible</div>
  </div>
  <div class="filter-bar">
    <input type="text" id="flt" placeholder="Buscar SKU o producto..." style="width:280px;">
    <select id="flt-casa"><option value="">Todas las casas</option></select>
    <select id="flt-estado">
      <option value="">Todos</option>
      <option value="multi">En 2+ pisos</option>
      <option value="full">4 pisos llenos</option>
    </select>
    <div style="flex:1"></div>
    <button class="btn btn-primary" onclick="window.print()">&#x1F5A8; Imprimir</button>
  </div>
  <div id="sku-list"></div>`;

  // Populate casa filter
  const casas = [...new Set(skus.map(s => s.casa).filter(Boolean))].sort();
  const sel = document.getElementById('flt-casa');
  casas.forEach(c => {{ const o = document.createElement('option'); o.value = c; o.textContent = c; sel.appendChild(o); }});

  renderSkuList(skus);
  const doFilter = () => {{
    const q = document.getElementById('flt').value.toLowerCase();
    const casa = document.getElementById('flt-casa').value;
    const estado = document.getElementById('flt-estado').value;
    renderSkuList(skus.filter(s => {{
      if (q && !s.sku.toLowerCase().includes(q) && !(s.nombre||'').toLowerCase().includes(q)) return false;
      if (casa && s.casa !== casa) return false;
      if (estado === 'multi' && s.con_producto < 2) return false;
      if (estado === 'full' && !s.necesita_espacio) return false;
      return true;
    }}));
  }};
  document.getElementById('flt').addEventListener('input', doFilter);
  document.getElementById('flt-casa').addEventListener('change', doFilter);
  document.getElementById('flt-estado').addEventListener('change', doFilter);
}}

function renderSkuList(skus) {{
  document.getElementById('sku-list').innerHTML = skus.map(sk => {{
    // Mini dots for header
    const dots = sk.bloques.map(b => {{
      let cls = 'gr';
      if (b.rol === 'DESPACHO' && b.prioridad === 1) cls = 'g';
      else if (b.rol === 'DESPACHO') cls = 'b';
      else if (b.rol === 'RECEPCION') cls = 'o';
      return '<span class="piso-dot '+cls+'" title="'+b.posicion+' - '+b.rol+'">'+b.id+'</span>';
    }}).join('');

    let badgeCls = 'ok', badgeText = sk.con_producto + ' piso' + (sk.con_producto !== 1 ? 's' : '');
    if (sk.necesita_espacio) {{ badgeCls = 'full'; badgeText = '4/4 LLENO'; }}
    else if (sk.con_producto > 1) {{ badgeCls = 'warn'; badgeText = sk.con_producto + ' pisos'; }}

    const blocksHtml = sk.bloques.map(b => {{
      let cls, roleText;
      if (b.rol === 'DESPACHO' && b.prioridad === 1) {{ cls = 'despacho-1'; roleText = 'DESPACHO #1'; }}
      else if (b.rol === 'DESPACHO') {{ cls = 'despacho-n'; roleText = 'DESPACHO #' + b.prioridad; }}
      else if (b.rol === 'RECEPCION') {{ cls = 'recepcion'; roleText = 'RECEPCION'; }}
      else {{ cls = 'disponible'; roleText = 'DISPONIBLE'; }}

      const s = b.stats;
      const dias = daysTo(s.fecha_min);
      let diasT = '-';
      if (dias !== null) diasT = dias < 0 ? '<span style="color:var(--red)">VENCIDO '+Math.abs(dias)+'d</span>' : dias+'d';

      let det = '';
      if (s.detalle && s.detalle.length) {{
        det = '<div class="blk-detail"><table><tr><th>Cant</th><th>Vence</th><th>Lote</th><th>LP</th></tr>' +
          s.detalle.map(d => '<tr><td>'+fmtQty(d.qty)+'</td><td>'+fmtDate(d.expDate)+'</td><td>'+(d.lote||'-')+'</td><td>'+(d.lp||'-')+'</td></tr>').join('') +
          '</table></div>';
      }}

      const prioHtml = b.prioridad ? '<div class="blk-prio">#'+b.prioridad+'</div>' : '';

      return '<div class="blk '+cls+'">' + prioHtml +
        '<div class="blk-role">'+roleText+'</div>' +
        '<div class="blk-id">'+b.posicion+'</div>' +
        '<div class="blk-stat"><span class="lbl">Cantidad</span><span class="val">'+fmtQty(s.qty)+'</span></div>' +
        '<div class="blk-stat"><span class="lbl">Vence min</span><span class="val">'+fmtDate(s.fecha_min)+'</span></div>' +
        '<div class="blk-stat"><span class="lbl">Vence max</span><span class="val">'+fmtDate(s.fecha_max)+'</span></div>' +
        '<div class="blk-stat"><span class="lbl">Dias</span><span class="val">'+diasT+'</span></div>' +
        '<div class="blk-stat"><span class="lbl">LPs</span><span class="val">'+s.lps+'</span></div>' +
        det + '</div>';
    }}).join('');

    return '<div class="sku-card"><div class="sku-header" onclick="this.nextElementSibling.classList.toggle(\'open\')">' +
      '<span class="sku-code">'+sk.sku+'</span>' +
      '<span class="sku-name">'+(sk.nombre||'')+'</span>' +
      '<span class="sku-casa">'+(sk.casa||'')+'</span>' +
      '<span class="sku-piso-dots">'+dots+'</span>' +
      '<span class="sku-badge '+badgeCls+'">'+badgeText+'</span>' +
      '</div><div class="sku-body"><div class="blocks-grid">'+blocksHtml+'</div></div></div>';
  }}).join('');
}}

function renderSigns() {{
  const el = document.getElementById('sec-signs');
  const pa = document.getElementById('print-area');
  if (!DATA.skus.length) {{
    el.innerHTML = '<div class="empty-state"><h3>Sin datos para rotulos</h3></div>';
    return;
  }}

  const ts = DATA.timestamp;
  let signs = '';
  let signCount = 0;

  DATA.skus.forEach(sk => {{
    sk.bloques.forEach(b => {{
      if (b.rol === 'DISPONIBLE') return;
      signCount++;

      let roleCls, roleText, warningText, prioCls;
      if (b.rol === 'DESPACHO' && b.prioridad === 1) {{
        roleCls = 'despacho'; prioCls = 'despacho';
        roleText = 'DESPACHAR PRIMERO - PRIORIDAD #1';
        warningText = 'NO COLOCAR PRODUCTO NUEVO EN ESTE PISO';
      }} else if (b.rol === 'DESPACHO') {{
        roleCls = 'espera'; prioCls = 'espera';
        roleText = 'EN ESPERA - PRIORIDAD #' + b.prioridad;
        warningText = 'DESPACHAR CUANDO #' + (b.prioridad-1) + ' SE VACIE - NO COLOCAR PRODUCTO NUEVO';
      }} else {{
        roleCls = 'recepcion'; prioCls = 'recepcion';
        roleText = 'RECEPCION - PRODUCTO NUEVO AQUI';
        warningText = 'NO DESPACHAR HASTA QUE LOS PISOS ANTERIORES SE VACIEN';
      }}

      signs += '<div class="sign-page">' +
        '<div class="sign-role-bar '+roleCls+'">'+roleText+'</div>' +
        '<div class="sign-prio-circle '+prioCls+'">'+(b.prioridad||'R')+'</div>' +
        '<div class="sign-sku-row"><span class="sign-sku">'+sk.sku+'</span><span class="sign-name">'+(sk.nombre||'')+'</span></div>' +
        '<div class="sign-block-id">'+b.posicion+'</div>' +
        '<div class="sign-dates">' +
          '<div><span class="dt-label">Vence desde: </span><span class="dt-value">'+fmtDate(b.stats.fecha_min)+'</span></div>' +
          '<div><span class="dt-label">Hasta: </span><span class="dt-value">'+fmtDate(b.stats.fecha_max)+'</span></div>' +
          '<div><span class="dt-label">Cantidad: </span><span class="dt-value">'+fmtQty(b.stats.qty)+' und</span></div>' +
        '</div>' +
        '<div class="sign-warning">'+warningText+'</div>' +
        '<div class="sign-footer"><span>Control Bloques FEFO - Bodega 1110</span><span>'+fmtDate(ts)+'</span><span>ARIUM</span></div>' +
        '</div>';
    }});
  }});

  const prot = `<div class="protocol-page">
    <h2>PROTOCOLO - COLA DE PRIORIDAD FEFO (PISO B1-B4)</h2>
    <p>Cada codigo puede estar en <strong>1, 2, 3 o 4 pisos</strong>. El sistema ordena automaticamente por fecha de vencimiento.</p>
    <h3>ROTULOS POR COLOR</h3>
    <div class="ok"><strong style="color:#166534;">VERDE #1</strong> = Despachar PRIMERO (fecha mas vieja)</div>
    <div class="rule"><strong style="color:#0369a1;">AZUL #2, #3</strong> = En espera. Despachar cuando el anterior se vacie.</div>
    <div class="rule"><strong style="color:#92400e;">NARANJA</strong> = Recepcion. Colocar producto nuevo SOLO aqui.</div>
    <h3>CUANDO UN PISO SE VACIA</h3>
    <div class="ok">Las prioridades se recorren: #2 sube a #1, #3 a #2, etc. El piso vacio queda disponible. Reimprimir rotulos.</div>
    <h3>DESCARGA NUEVA</h3>
    <div class="rule">Colocar en el piso NARANJA. Si no hay, usar un piso GRIS (disponible). Ese piso recibe la prioridad mas baja.</div>
    <h3 style="color:#ef4444;">PROHIBIDO</h3>
    <div class="danger">
    - NO despachar de un piso #2+ mientras #1 tenga existencia de ese codigo<br>
    - NO colocar producto nuevo en pisos verdes o azules<br>
    - NO mezclar fechas distintas del mismo codigo en un solo piso
    </div>
    <div style="margin-top:20px;text-align:center;font-size:12px;color:#999;">
    Control Bloques FEFO &middot; Bodega 1110 &middot; ARIUM CORPORACION &middot; ${{fmtDate(ts)}}
    </div>
  </div>`;

  el.innerHTML = '<div style="margin-bottom:20px;display:flex;gap:12px;align-items:center;">' +
    '<button class="btn btn-primary" onclick="window.print()">&#x1F5A8; Imprimir Todo</button>' +
    '<span style="color:var(--text-dim);font-size:13px;">'+signCount+' rotulos + 1 protocolo</span></div>' +
    '<div class="preview-signs">' + signs + prot + '</div>';
  pa.innerHTML = signs + prot;
}}

renderDashboard();
renderSigns();
</script>
</body>
</html>"""
    return html


# ══════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    print("=" * 60)
    print("  CONTROL DE BLOQUES FEFO - Bodega 1110")
    print("  Posiciones: " + ", ".join(POSICIONES_PISO))
    print("=" * 60)

    conn = connect()
    print("Conectado a SQL OK")

    timestamp = datetime.datetime.now().isoformat()

    print("\nConsultando inventario...")
    rows = fetch_inventory(conn)
    print(f"  {len(rows)} registros en PISO-B1 a B4")

    sku_data = build_sku_blocks(rows)
    print(f"  {len(sku_data)} SKUs con inventario\n")

    for sk in sku_data:
        pisos = [f"{b['id']}={b['rol']}{'#'+str(b['prioridad']) if b['prioridad'] else ''}({b['stats']['qty']:.0f})"
                 for b in sk['bloques'] if b['stats']['qty'] > 0]
        warn = " *** 4/4 LLENO" if sk['necesita_espacio'] else ""
        print(f"  {sk['sku']} {sk['nombre'][:30]:30s} [{' | '.join(pisos)}]{warn}")

    conn.close()

    html = generate_html(sku_data, timestamp)
    out_path = os.path.join(DIR, "CONTROL_BLOQUES_FEFO.html")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\nDashboard: {out_path}")
    os.startfile(out_path)
