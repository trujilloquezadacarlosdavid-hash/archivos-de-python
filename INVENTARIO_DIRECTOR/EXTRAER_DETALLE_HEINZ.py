# -*- coding: utf-8 -*-
"""
DETALLE KRAFT HEINZ - Extractor SQL -> HTML interactivo
Genera DETALLE_HEINZ.html con el detalle a nivel LP de todo el producto Kraft Heinz:
ubicacion, SKU, descripcion, status, LP, cajas, fecha PEPS, fecha expiracion,
dias para vencer y grupo de vencimiento (VENCIDO / <=30 / 31-60 / 61-90 / >90).
Conversion a cajas: Qty / ConversionFactor(Uom='CJ') de HJ_MasterData (factor oficial Korber).
Excluye Philips y Listerine (otras casas).
"""
import pyodbc, json, os, datetime, sys
sys.stdout.reconfigure(encoding='utf-8')

DIR = os.path.dirname(os.path.abspath(__file__))
SQL = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=172.16.5.43;DATABASE=IN_E;Trusted_Connection=yes;TrustServerCertificate=yes"
BODEGA = '1110A001'
EXCLUIR = ('A0000004439', 'A0010000291')  # Philips LED, Listerine (otras casas)

print("Conectando a SQL...")
conn = pyodbc.connect(SQL, timeout=30)
cur = conn.cursor()

# SKUs Kraft Heinz (por proveedor), menos Philips/Listerine
cur.execute("""
SELECT DISTINCT d.codigo_sku
FROM LOGISTICA.ENCABEZADO_RECEPCION_20 e
JOIN LOGISTICA.DETALLE_RECEPCION_20 d ON e.cita_id=d.cita_id
WHERE e.bodega_codigo='1110' AND e.proveedor_nombre_co='KRAFT HEINZ' AND d.det_canceled='N'
""")
skus = [r[0] for r in cur.fetchall() if r[0] not in EXCLUIR]
ph = ",".join("?"*len(skus))
print(f"SKUs Kraft Heinz: {len(skus)}")

# Factor de caja por SKU
cur.execute(f"SELECT ItemCode, ConversionFactor FROM LOGISTICA.HJ_MasterData WHERE WhsCode='{BODEGA}' AND Uom='CJ' AND ItemCode IN ({ph})", skus)
cf = {r[0]: float(r[1]) for r in cur.fetchall() if r[1]}
for s in skus: cf.setdefault(s, 1.0)

# Costo por unidad base (Quetzales) = V_MD_DATA_SAP_OITM.STOCK_PRICE (costo real SAP)
cur.execute(f"SELECT ItemCode, MAX(STOCK_PRICE) FROM LOGISTICA.V_MD_DATA_SAP_OITM WHERE ItemCode IN ({ph}) GROUP BY ItemCode", skus)
price = {r[0]: float(r[1]) for r in cur.fetchall() if r[1]}
for s in skus: price.setdefault(s, 0.0)

# Detalle a nivel LP
print("Extrayendo detalle a nivel LP...")
cur.execute(f"""
SELECT i.Location, i.ItemCode, i.ItemName, i.Status, i.BlockedCode, i.License,
       i.Qty, i.FifoDate, i.ExpirationDate, i.LotNumber
FROM LOGISTICA.HJ_Inventory i
WHERE i.Warehouse='{BODEGA}' AND i.Qty>0 AND i.ItemCode IN ({ph})
ORDER BY i.ExpirationDate, i.ItemCode
""", skus)
cols = [c[0] for c in cur.description]
hoy = datetime.date.today()

BLK = {  # nombres legibles de bloqueos
    'PLNER':'Planeacion','BDVE':'Bodega Vencido','BDCDV':'CD Vencido','BDCU':'Cuarentena',
    'BDQA':'QA','BDSC':'Scrap','BDCDS':'CD Scrap','BDUN1':'Unidades','LIPQA':'QA Lipton'
}

def grupo(dv, tiene_fecha):
    if not tiene_fecha: return ('SINF','Sin fecha')
    if dv < 0: return ('VEN','Vencido')
    if dv <= 30: return ('G30','<=30 dias')
    if dv <= 60: return ('G60','31-60 dias')
    if dv <= 90: return ('G90','61-90 dias')
    return ('VIG','Vigente (>90d)')

rows = []
for r in cur.fetchall():
    d = dict(zip(cols, r))
    factor = cf.get(d['ItemCode'], 1.0) or 1.0
    cajas_x = d['Qty']/factor if factor else float(d['Qty'])  # sin redondear (para totales)
    cajas = round(cajas_x, 1)  # redondeado (para mostrar por LP)
    valor_q = d['Qty'] * price.get(d['ItemCode'], 0.0)  # Quetzales (Qty en unidad base)
    exp = d['ExpirationDate']
    tiene = bool(exp) and str(exp)[:4] > '1900'
    expd = exp.date() if hasattr(exp, 'date') else (exp if isinstance(exp, datetime.date) else None)
    dv = (expd - hoy).days if (tiene and expd) else None
    g, gnom = grupo(dv if dv is not None else 99999, tiene)
    fifo = d['FifoDate']
    rows.append({
        'u': d['Location'] or '',
        'c': d['ItemCode'],
        'n': (d['ItemName'] or '')[:45],
        's': d['Status'] or '',
        'b': d['BlockedCode'] or '',
        'bn': BLK.get(d['BlockedCode'] or '', d['BlockedCode'] or ''),
        'lp': str(d['License']) if d['License'] else '',
        'q': cajas, 'qx': cajas_x, 'val': round(valor_q),
        'peps': (fifo.isoformat()[:10] if hasattr(fifo,'isoformat') else (str(fifo)[:10] if fifo else '')),
        'exp': (expd.isoformat() if expd else ''),
        'dv': dv if dv is not None else '',
        'g': g, 'gn': gnom,
        'lt': d['LotNumber'] or '',
    })

conn.close()
print(f"  {len(rows)} LPs Heinz")

# Resumen por grupo (cajas SIN redondear por LP -> cuadra con la infografia)
resumen = {}
for r in rows:
    g = r['g']
    resumen.setdefault(g, {'cajas':0,'lps':0,'skus':set(),'val':0})
    resumen[g]['cajas'] += r['qx']; resumen[g]['lps'] += 1; resumen[g]['skus'].add(r['c']); resumen[g]['val'] += r['val']
for g in resumen: resumen[g]['skus'] = len(resumen[g]['skus'])

# Planeacion (PLNER)
plner = [r for r in rows if r['b']=='PLNER']
plner_cajas = round(sum(r['qx'] for r in plner))
plner_val = round(sum(r['val'] for r in plner))
plner_lps = len(plner)
plner_skus = len(set(r['c'] for r in plner))

tot_cajas = round(sum(r['qx'] for r in rows))
tot_val = round(sum(r['val'] for r in rows))
now = datetime.datetime.now()
for r in rows: del r['qx']  # quitar antes de serializar (solo se usaba para totales)

def gc(g): return round(resumen.get(g,{}).get('cajas',0))
def gl(g): return resumen.get(g,{}).get('lps',0)
def gv(g): return round(resumen.get(g,{}).get('val',0))
def Qm(v):
    if v>=1e6: return f"Q {v/1e6:.2f}M"
    if v>=1e3: return f"Q {v/1e3:.0f}K"
    return f"Q {v:,.0f}"

meses=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
fecha_txt = f"{now.day} de {meses[now.month-1]} de {now.year}, {now.strftime('%H:%M')}"

data_json = json.dumps(rows, ensure_ascii=False, separators=(',',':'))

html = f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Detalle Kraft Heinz - Vencimiento y Planeacion</title>
<style>
:root{{--azul:#0d2a6b;--azul2:#1a3d8f;--dorado:#f5b400;--verde:#0ca30c;--rojo:#d03b3b;--naranja:#eb6834;--warn:#fab219;--gris1:#f8f9fa;--gris2:#e9ecef;--gris3:#6c757d;--grisL:#e1e0d9;--texto:#0b0b0b;--texto2:#52514e;--muted:#898781;--surface:#fcfcfb;--bg:#f0f2f5;}}
*{{margin:0;padding:0;box-sizing:border-box;}}
body{{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--texto);font-size:13px;}}
.hdr{{background:linear-gradient(135deg,var(--azul),var(--azul2));color:#fff;padding:14px 28px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;}}
.hdr h1{{font-size:20px;font-weight:800;}} .hdr h1 span{{color:var(--dorado);}}
.hdr .sub{{font-size:11px;opacity:.7;margin-top:2px;}}
.hdr-r{{text-align:right;font-size:11px;opacity:.85;}}
.wrap{{max-width:1550px;margin:0 auto;padding:16px 20px;}}
.kpis{{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:14px;}}
.kp{{background:var(--surface);border:1px solid var(--grisL);border-radius:8px;padding:10px 12px;border-left:4px solid var(--azul);cursor:pointer;transition:all .15s;}}
.kp:hover{{box-shadow:0 2px 8px rgba(0,0,0,.1);transform:translateY(-1px);}}
.kp.active{{box-shadow:0 0 0 2px var(--azul);}}
.kp-lb{{font-size:10px;font-weight:600;text-transform:uppercase;color:var(--muted);}}
.kp-v{{font-size:20px;font-weight:800;}} .kp-s{{font-size:10px;color:var(--texto2);}}
.panel{{background:var(--surface);border:1px solid var(--grisL);border-radius:10px;padding:14px 18px;margin-bottom:14px;}}
.panel-t{{font-size:13px;font-weight:700;color:var(--azul);text-transform:uppercase;letter-spacing:.3px;margin-bottom:10px;border-bottom:2px solid var(--dorado);padding-bottom:5px;display:inline-block;}}
.plan-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px;}}
.plan-box{{background:#fff9e6;border:1px solid var(--dorado);border-radius:8px;padding:12px 14px;}}
.plan-hero{{display:flex;gap:20px;align-items:center;}}
.plan-num{{font-size:34px;font-weight:800;color:#8B6914;line-height:1;}}
.acc{{font-size:12px;line-height:1.7;}} .acc li{{margin-left:16px;margin-bottom:3px;}}
.acc .tag{{display:inline-block;padding:1px 6px;border-radius:3px;font-size:9px;font-weight:700;color:#fff;margin-right:5px;}}
.filters{{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:10px;background:var(--surface);border:1px solid var(--grisL);border-radius:8px;padding:10px 14px;}}
.filters label{{font-size:11px;font-weight:600;color:var(--texto2);margin-right:4px;}}
select,input{{padding:6px 10px;border:1px solid var(--grisL);border-radius:6px;font-size:12px;font-family:inherit;}}
input[type=text]{{min-width:220px;}}
.btn{{padding:6px 14px;background:var(--azul);color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;}}
.btn:hover{{background:var(--azul2);}}
.tbl-wrap{{background:var(--surface);border:1px solid var(--grisL);border-radius:10px;overflow:auto;max-height:calc(100vh - 120px);}}
table{{width:100%;border-collapse:separate;border-spacing:0;font-size:12px;}}
th{{position:sticky;top:0;background:var(--azul);color:#fff;padding:8px 10px;text-align:left;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap;z-index:2;}}
th:hover{{background:var(--azul2);}} th .ar{{opacity:.5;font-size:9px;}}
td{{padding:6px 10px;border-bottom:1px solid var(--gris2);white-space:nowrap;}}
tr:hover td{{background:#f0f6ff;}}
.badge{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;color:#fff;}}
.b-VEN{{background:var(--rojo);}} .b-G30{{background:var(--naranja);}} .b-G60{{background:var(--warn);color:#5a4400;}} .b-G90{{background:#c9a227;color:#fff;}} .b-VIG{{background:var(--verde);}} .b-SINF{{background:var(--gris3);}}
.st-H{{color:var(--rojo);font-weight:700;}} .st-A{{color:var(--verde);font-weight:700;}}
.plner-row td{{background:#fff8e1;}}
.num{{text-align:right;font-variant-numeric:tabular-nums;}}
.cnt{{font-size:11px;color:var(--texto2);margin:8px 0;}}
</style></head>
<body>
<div class="hdr">
  <div><h1>Detalle <span>Kraft Heinz</span> — Vencimiento y Planeaci&oacute;n</h1>
  <div class="sub">Inventario a nivel LP · Bodega 1110A001 · CD El Frutal · Me Llega, S.A.</div></div>
  <div class="hdr-r"><div><b>{tot_cajas:,} cajas · {Qm(tot_val)}</b> · {len(rows):,} LPs · {len(skus)} SKUs</div>
  <div>Corte: {fecha_txt}</div><div style="opacity:.6">K&ouml;rber WMS · cajas (factor oficial) · Q = costo x Qty</div></div>
</div>
<div class="wrap">

  <div class="kpis">
    <div class="kp" style="border-left-color:var(--naranja)" onclick="filtrarPlner()"><div class="kp-lb">Planeaci&oacute;n (PLNER)</div><div class="kp-v" style="color:var(--naranja)">{plner_cajas:,}</div><div class="kp-s">{plner_lps} LPs · <b>{Qm(plner_val)}</b></div></div>
    <div class="kp" style="border-left-color:var(--rojo)" onclick="filtrarG('VEN')"><div class="kp-lb">Vencido</div><div class="kp-v" style="color:var(--rojo)">{gc('VEN'):,}</div><div class="kp-s">{gl('VEN')} LPs · <b>{Qm(gv('VEN'))}</b></div></div>
    <div class="kp" style="border-left-color:var(--naranja)" onclick="filtrarG('G30')"><div class="kp-lb">&le;30 d&iacute;as</div><div class="kp-v" style="color:var(--naranja)">{gc('G30'):,}</div><div class="kp-s">{gl('G30')} LPs · <b>{Qm(gv('G30'))}</b></div></div>
    <div class="kp" style="border-left-color:var(--warn)" onclick="filtrarG('G60')"><div class="kp-lb">31-60 d&iacute;as</div><div class="kp-v" style="color:#a67c00">{gc('G60'):,}</div><div class="kp-s">{gl('G60')} LPs · <b>{Qm(gv('G60'))}</b></div></div>
    <div class="kp" style="border-left-color:#c9a227" onclick="filtrarG('G90')"><div class="kp-lb">61-90 d&iacute;as</div><div class="kp-v" style="color:#a67c00">{gc('G90'):,}</div><div class="kp-s">{gl('G90')} LPs · <b>{Qm(gv('G90'))}</b></div></div>
    <div class="kp" style="border-left-color:var(--verde)" onclick="filtrarG('VIG')"><div class="kp-lb">Vigente (&gt;90d)</div><div class="kp-v" style="color:var(--verde)">{gc('VIG'):,}</div><div class="kp-s">{gl('VIG')} LPs · <b>{Qm(gv('VIG'))}</b></div></div>
  </div>

  <div class="panel">
    <div class="panel-t">Producto en Planeaci&oacute;n — Plan de Acci&oacute;n para dar salida</div>
    <div class="plan-grid">
      <div class="plan-box"><div class="plan-hero">
        <div><div class="plan-num">{plner_cajas:,}</div><div style="font-size:11px;color:#8B6914;font-weight:600">cajas · {Qm(plner_val)} en PLNER</div></div>
        <div style="font-size:12px;color:#5a4400;line-height:1.6"><b>{plner_skus}</b> SKUs · <b>{plner_lps}</b> LPs activas<br>Producto separado por corrugado da&ntilde;ado,<br>destino: <b>CDDs</b>. Independiente del pasillo.</div>
      </div></div>
      <div class="plan-box" style="background:#eef4ff;border-color:var(--azul)">
        <div style="font-size:12px;font-weight:700;color:var(--azul);margin-bottom:6px">Planes sugeridos (por urgencia)</div>
        <ul class="acc">
          <li><span class="tag" style="background:var(--rojo)">VENCIDO</span>Dar de baja / merma; tramitar destrucci&oacute;n o devoluci&oacute;n a Heinz.</li>
          <li><span class="tag" style="background:var(--naranja)">&le;30d</span>Despacho inmediato prioritario; liquidaci&oacute;n con comercial.</li>
          <li><span class="tag" style="background:var(--warn);color:#5a4400">31-60d</span>Empujar rotaci&oacute;n PEPS; ofertar a Mayoreo / CEDIS.</li>
          <li><span class="tag" style="background:var(--azul)">PLNER</span>Coordinar salida a CDDs; re-empaque si el corrugado lo permite.</li>
        </ul>
      </div>
    </div>
  </div>

  <div class="filters">
    <div><label>Grupo</label><select id="fGrupo" onchange="render()">
      <option value="">Todos</option><option value="VEN">Vencido</option><option value="G30">&le;30 dias</option>
      <option value="G60">31-60 dias</option><option value="G90">61-90 dias</option><option value="VIG">Vigente &gt;90d</option><option value="SINF">Sin fecha</option>
    </select></div>
    <div><label>Status</label><select id="fStatus" onchange="render()"><option value="">Todos</option><option value="A">A (Disponible)</option><option value="H">H (Bloqueado)</option></select></div>
    <div><label>Bloqueo</label><select id="fBloq" onchange="render()"><option value="">Todos</option><option value="PLNER">Planeacion (PLNER)</option></select></div>
    <div><label>Buscar</label><input type="text" id="fBuscar" placeholder="SKU, descripcion, ubicacion, LP..." oninput="render()"></div>
    <button class="btn" onclick="limpiar()">Limpiar</button>
    <button class="btn" onclick="exportCSV()">Exportar CSV</button>
    <span class="cnt" id="contador" style="margin-left:auto"></span>
  </div>

  <div class="tbl-wrap"><table id="tbl">
    <thead><tr>
      <th onclick="ordenar('u')">Ubicaci&oacute;n<span class="ar"> &#8645;</span></th>
      <th onclick="ordenar('c')">SKU<span class="ar"> &#8645;</span></th>
      <th onclick="ordenar('n')">Descripci&oacute;n<span class="ar"> &#8645;</span></th>
      <th onclick="ordenar('s')">St<span class="ar"> &#8645;</span></th>
      <th onclick="ordenar('bn')">Bloqueo<span class="ar"> &#8645;</span></th>
      <th onclick="ordenar('lp')">LP<span class="ar"> &#8645;</span></th>
      <th class="num" onclick="ordenar('q')">Cajas<span class="ar"> &#8645;</span></th>
      <th class="num" onclick="ordenar('val')">Valor Q<span class="ar"> &#8645;</span></th>
      <th onclick="ordenar('peps')">PEPS<span class="ar"> &#8645;</span></th>
      <th onclick="ordenar('exp')">Expira<span class="ar"> &#8645;</span></th>
      <th class="num" onclick="ordenar('dv')">D&iacute;as<span class="ar"> &#8645;</span></th>
      <th onclick="ordenar('g')">Grupo<span class="ar"> &#8645;</span></th>
    </tr></thead>
    <tbody id="tbody"></tbody>
  </table></div>
</div>

<script>
const DATA = {data_json};
let orden = {{col:'exp', asc:true}};

function pasa(r){{
  const g=document.getElementById('fGrupo').value;
  const s=document.getElementById('fStatus').value;
  const b=document.getElementById('fBloq').value;
  const q=document.getElementById('fBuscar').value.toLowerCase().trim();
  if(g && r.g!==g) return false;
  if(s && r.s!==s) return false;
  if(b && r.b!==b) return false;
  if(q){{
    const hay=(r.c+' '+r.n+' '+r.u+' '+r.lp+' '+r.lt).toLowerCase();
    if(!hay.includes(q)) return false;
  }}
  return true;
}}
function ordenar(col){{ if(orden.col===col) orden.asc=!orden.asc; else {{orden.col=col; orden.asc=true;}} render(); }}
function render(){{
  let f = DATA.filter(pasa);
  const c=orden.col;
  f.sort((a,b)=>{{ let x=a[c],y=b[c]; if(x===''&&y!=='')return 1; if(y===''&&x!=='')return -1;
    if(typeof x==='number'&&typeof y==='number') return orden.asc?x-y:y-x;
    x=(''+x).toLowerCase(); y=(''+y).toLowerCase(); return orden.asc?(x<y?-1:x>y?1:0):(x>y?-1:x<y?1:0); }});
  const tb=document.getElementById('tbody');
  let cajas=0, valor=0; const html=[];
  for(const r of f){{
    cajas+=r.q; valor+=r.val;
    html.push('<tr'+(r.b==='PLNER'?' class="plner-row"':'')+'>'+
      '<td>'+r.u+'</td><td>'+r.c+'</td><td>'+r.n+'</td>'+
      '<td class="st-'+r.s+'">'+r.s+'</td><td>'+(r.bn||'-')+'</td><td>'+(r.lp||'-')+'</td>'+
      '<td class="num">'+r.q.toLocaleString()+'</td><td class="num">Q '+r.val.toLocaleString()+'</td>'+
      '<td>'+(r.peps||'-')+'</td><td>'+(r.exp||'-')+'</td>'+
      '<td class="num">'+(r.dv===''?'-':r.dv)+'</td>'+
      '<td><span class="badge b-'+r.g+'">'+r.gn+'</span></td></tr>');
  }}
  tb.innerHTML=html.join('');
  document.getElementById('contador').textContent=f.length.toLocaleString()+' LPs · '+Math.round(cajas).toLocaleString()+' cajas · Q '+Math.round(valor).toLocaleString();
}}
function filtrarG(g){{ limpiarSolo(); document.getElementById('fGrupo').value=g; render(); }}
function filtrarPlner(){{ limpiarSolo(); document.getElementById('fBloq').value='PLNER'; render(); }}
function limpiarSolo(){{ document.getElementById('fGrupo').value=''; document.getElementById('fStatus').value=''; document.getElementById('fBloq').value=''; document.getElementById('fBuscar').value=''; }}
function limpiar(){{ limpiarSolo(); render(); }}
function exportCSV(){{
  const f=DATA.filter(pasa);
  let csv='Ubicacion,SKU,Descripcion,Status,Bloqueo,LP,Cajas,ValorQ,PEPS,Expira,Dias,Grupo\\n';
  for(const r of f){{ csv+=[r.u,r.c,'"'+r.n+'"',r.s,r.bn,r.lp,r.q,r.val,r.peps,r.exp,r.dv,r.gn].join(',')+'\\n'; }}
  const a=document.createElement('a'); a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(csv);
  a.download='detalle_heinz.csv'; a.click();
}}
render();
</script>
</body></html>"""

out = os.path.join(DIR, 'DETALLE_HEINZ.html')
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"Generado: {out}  ({os.path.getsize(out)/1024:.0f} KB)")
print(f"  Planeacion: {plner_cajas:,} cajas / {plner_skus} SKUs / {plner_lps} LPs")
print(f"  Vencido: {gc('VEN'):,} | <=30d: {gc('G30'):,} | 31-60d: {gc('G60'):,} | 61-90d: {gc('G90'):,} | Vigente: {gc('VIG'):,}")
print("Listo.")
