# -*- coding: utf-8 -*-
"""
INFOGRAFIA S&OP KRAFT HEINZ - Extractor SQL -> HTML
Genera INFOGRAFIA_SOP_DIRECTOR.html con datos EN VIVO de Korber.
- Solo producto Kraft Heinz (proveedor KRAFT HEINZ), excluye Philips y Listerine (otras casas).
- Todo en CAJAS: Qty / ConversionFactor(Uom='CJ') de HJ_MasterData (factor oficial Korber).
- Despacho = TranType 341 (Shipping Detail = lo realmente embarcado).
"""
import pyodbc, os, datetime, math, sys
sys.stdout.reconfigure(encoding='utf-8')

DIR = os.path.dirname(os.path.abspath(__file__))
SQL = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=172.16.5.43;DATABASE=IN_E;Trusted_Connection=yes;TrustServerCertificate=yes"
BODEGA = '1110A001'
EXCLUIR = ('A0000004439', 'A0010000291')  # Philips LED, Listerine
YEAR = 2026

print("Conectando a SQL...")
conn = pyodbc.connect(SQL, timeout=30)
cur = conn.cursor()

cur.execute("""SELECT DISTINCT d.codigo_sku FROM LOGISTICA.ENCABEZADO_RECEPCION_20 e
JOIN LOGISTICA.DETALLE_RECEPCION_20 d ON e.cita_id=d.cita_id
WHERE e.bodega_codigo='1110' AND e.proveedor_nombre_co='KRAFT HEINZ' AND d.det_canceled='N'""")
skus = [r[0] for r in cur.fetchall() if r[0] not in EXCLUIR]
ph = ",".join("?"*len(skus))
print(f"SKUs Kraft Heinz: {len(skus)}")

cur.execute(f"SELECT ItemCode,ConversionFactor FROM LOGISTICA.HJ_MasterData WHERE WhsCode='{BODEGA}' AND Uom='CJ' AND ItemCode IN ({ph})", skus)
cf = {r[0]: float(r[1]) for r in cur.fetchall() if r[1]}
for s in skus: cf.setdefault(s, 1.0)

def cajas(where, extra=None, table='inv'):
    if table == 'inv':
        sql = f"SELECT ItemCode,CAST(SUM(CAST(Qty AS BIGINT))AS BIGINT) FROM LOGISTICA.HJ_Inventory WHERE Warehouse='{BODEGA}' AND Qty>0 AND ItemCode IN ({ph}) AND {where} GROUP BY ItemCode"
    else:
        sql = f"SELECT ItemCode,CAST(SUM(CAST(TrxQty AS BIGINT))AS BIGINT) FROM LOGISTICA.HJ_TRANSACTIONS WHERE WhsCode='{BODEGA}' AND ItemCode IN ({ph}) AND {where} GROUP BY ItemCode"
    cur.execute(sql, list(skus)+(extra or []))
    return round(sum((r[1]/cf.get(r[0],1.0) if cf.get(r[0]) else r[1]) for r in cur.fetchall()))

print("Calculando metricas...")
stock = cajas("1=1")
dispo = cajas("Status='A'")
bloq = cajas("Status='H'")
plner = cajas("BlockedCode='PLNER'")
# Vencimiento por DIAS EXACTOS y solo FECHA (identico al detalle para que cuadren)
VALID = "ExpirationDate IS NOT NULL AND ExpirationDate>'1900-01-02'"
DV = "DATEDIFF(DAY,CAST(GETDATE() AS DATE),CAST(ExpirationDate AS DATE))"
vencido = cajas(f"{VALID} AND {DV}<0")
v30 = cajas(f"{VALID} AND {DV} BETWEEN 0 AND 30")
v90 = cajas(f"{VALID} AND {DV} BETWEEN 31 AND 90")
vigente = cajas(f"{VALID} AND {DV}>90")
sinf = cajas("(ExpirationDate IS NULL OR ExpirationDate<='1900-01-02')")
desp = cajas("TranType=341 AND TrxDate>=?", [f'{YEAR}-01-01'], table='trx')

# ── VALOR EN QUETZALES (costo por unidad base = V_MD_DATA_SAP_OITM.STOCK_PRICE, costo real SAP) ──
cur.execute(f"SELECT ItemCode,MAX(STOCK_PRICE) FROM LOGISTICA.V_MD_DATA_SAP_OITM WHERE ItemCode IN ({ph}) GROUP BY ItemCode", skus)
price = {r[0]: float(r[1]) for r in cur.fetchall() if r[1]}
for s in skus: price.setdefault(s, 0.0)
def valorQ(where, extra=None):
    cur.execute(f"SELECT ItemCode,CAST(SUM(CAST(Qty AS BIGINT))AS BIGINT) FROM LOGISTICA.HJ_Inventory WHERE Warehouse='{BODEGA}' AND Qty>0 AND ItemCode IN ({ph}) AND {where} GROUP BY ItemCode", list(skus)+(extra or []))
    return round(sum(r[1]*price.get(r[0],0.0) for r in cur.fetchall()))
val_total = valorQ("1=1")
val_plner = valorQ("BlockedCode='PLNER'")
val_vencido = valorQ(f"{VALID} AND {DV}<0")
val_riesgo = valorQ(f"{VALID} AND {DV}<=90")
# Composicion 100% por estado/disposicion (cajas + valor Q)
c_disp_c, c_disp_v = dispo, valorQ("Status='A'")
c_plan_c, c_plan_v = plner, val_plner
c_venc_c, c_venc_v = cajas("BlockedCode IN ('BDVE','BDCDV')"), valorQ("BlockedCode IN ('BDVE','BDCDV')")
c_cuar_c, c_cuar_v = cajas("BlockedCode='BDCU'"), valorQ("BlockedCode='BDCU'")
c_otros_c = round(stock - c_disp_c - c_plan_c - c_venc_c - c_cuar_c)
c_otros_v = round(val_total - c_disp_v - c_plan_v - c_venc_v - c_cuar_v)
# Desglose fino para el detalle de bloqueos
c_bdve_c, c_bdve_v = cajas("BlockedCode='BDVE'"), valorQ("BlockedCode='BDVE'")
c_bdcdv_c, c_bdcdv_v = cajas("BlockedCode='BDCDV'"), valorQ("BlockedCode='BDCDV'")
comp_c = bloq                    # todo lo bloqueado (H) = comprometido
comp_v = valorQ("Status='H'")

cur.execute(f"SELECT COUNT(DISTINCT ItemCode),COUNT(*) FROM LOGISTICA.HJ_Inventory WHERE Warehouse='{BODEGA}' AND Qty>0 AND ItemCode IN ({ph})", skus)
r = cur.fetchone(); stock_skus, stock_lps = r[0], r[1]
cur.execute(f"SELECT COUNT(DISTINCT ItemCode),COUNT(*) FROM LOGISTICA.HJ_Inventory WHERE Warehouse='{BODEGA}' AND Qty>0 AND BlockedCode='PLNER' AND ItemCode IN ({ph})", skus)
r = cur.fetchone(); plner_skus, plner_lps = r[0], r[1]
cur.execute(f"SELECT COUNT(DISTINCT ItemCode) FROM LOGISTICA.HJ_Inventory WHERE Warehouse='{BODEGA}' AND Qty>0 AND ItemCode IN ({ph}) AND ExpirationDate>'1900-01-02' AND CAST(ExpirationDate AS DATE)<CAST(GETDATE() AS DATE)", skus)
vencido_skus = cur.fetchone()[0]

cur.execute(f"""SELECT CAST(SUM(CAST(d.cajas_enviadas AS BIGINT))AS BIGINT),COUNT(DISTINCT e.cita_id),COUNT(DISTINCT d.codigo_sku)
FROM LOGISTICA.ENCABEZADO_RECEPCION_20 e JOIN LOGISTICA.DETALLE_RECEPCION_20 d ON e.cita_id=d.cita_id
WHERE e.bodega_codigo='1110' AND e.proveedor_nombre_co='KRAFT HEINZ' AND e.fecha_cita>=? AND e.enc_canceled='N' AND d.det_canceled='N' AND d.codigo_sku IN ({ph})""", [f'{YEAR}-01-01']+skus)
r = cur.fetchone(); recibido, citas, recibido_skus = r[0], r[1], r[2]  # 'recibido' = enviado por Heinz (cajas_enviadas)
# Recibido REAL en Korber (korber_recibido) vs enviado = fill rate de recepcion
# MISMA lista de SKUs que el resto del deck (excluye Philips/Listerine) para no mezclar poblaciones
cur.execute(f"""SELECT SUM(CAST(korber_recibido AS FLOAT)) FROM LOGISTICA.HISTORICO_ASIGNACION_RECEPCION
WHERE bodega_codigo='1110' AND proveedor_nombre_co='KRAFT HEINZ' AND fecha_cita>=? AND det_canceled='N' AND enc_canceled='N'
AND codigo_sku IN ({ph})""", [f'{YEAR}-01-01']+skus)
recibido_real = round(cur.fetchone()[0] or 0)
fill_recep = round(recibido_real/recibido*100,1) if recibido else 0
faltante_recep = recibido - recibido_real

MESES = ['2026-01','2026-02','2026-03','2026-04','2026-05','2026-06','2026-07','2026-08','2026-09']
MLBL = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep']
rec_mes = {m: 0 for m in MESES}
cur.execute(f"""SELECT FORMAT(e.fecha_cita,'yyyy-MM'),CAST(SUM(CAST(d.cajas_enviadas AS BIGINT))AS BIGINT)
FROM LOGISTICA.ENCABEZADO_RECEPCION_20 e JOIN LOGISTICA.DETALLE_RECEPCION_20 d ON e.cita_id=d.cita_id
WHERE e.bodega_codigo='1110' AND e.proveedor_nombre_co='KRAFT HEINZ' AND e.fecha_cita>=? AND e.enc_canceled='N' AND d.det_canceled='N' AND d.codigo_sku IN ({ph})
GROUP BY FORMAT(e.fecha_cita,'yyyy-MM')""", [f'{YEAR}-01-01']+skus)
for r in cur.fetchall():
    if r[0] in rec_mes: rec_mes[r[0]] = r[1]
desp_mes = {m: cajas("TranType=341 AND FORMAT(TrxDate,'yyyy-MM')=?", [m], table='trx') for m in MESES}

canales = {c: cajas("TranType=341 AND TrxDate>=? AND Channel=?", [f'{YEAR}-01-01', c], table='trx') for c in ['Moderno','CEDIS','Mayoreo','Top Stores']}
canal_total = sum(canales.values()) or 1

def limpia(n):
    n = (n or '').upper().replace('HEINZ ','').replace('KRAFT ','').strip()
    return n.title()[:26]

cur.execute(f"""SELECT ItemCode,MAX(ItemName),CAST(SUM(CAST(Qty AS BIGINT))AS BIGINT)
FROM LOGISTICA.HJ_Inventory WHERE Warehouse='{BODEGA}' AND Qty>0 AND ItemCode IN ({ph}) GROUP BY ItemCode""", skus)
tops = sorted([{'n':limpia(r[1]),'cajas':round(r[2]/cf.get(r[0],1.0))} for r in cur.fetchall()], key=lambda x:-x['cajas'])[:4]

cur.execute(f"""SELECT ItemCode,MAX(ItemName),CAST(SUM(CAST(Qty AS BIGINT))AS BIGINT),MIN(ExpirationDate)
FROM LOGISTICA.HJ_Inventory WHERE Warehouse='{BODEGA}' AND Qty>0 AND ItemCode IN ({ph})
AND ExpirationDate>'1900-01-02' AND CAST(ExpirationDate AS DATE)<CAST(GETDATE() AS DATE) GROUP BY ItemCode""", skus)
tv = sorted([{'n':limpia(r[1]),'cajas':round(r[2]/cf.get(r[0],1.0)),'f':str(r[3])[:7]} for r in cur.fetchall()], key=lambda x:-x['cajas'])[:4]

conn.close()

# Derivados
dias = (datetime.date.today() - datetime.date(YEAR,1,1)).days
desp_diario = round(desp/dias) if dias else 0
dias_inv = round(stock/desp_diario) if desp_diario else 0
riesgo = vencido + v30 + v90
pct = lambda x: round(x/stock*100,1) if stock else 0
now = datetime.datetime.now()
meses_es = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
fecha_txt = f"{now.day} de {meses_es[now.month-1]} de {now.year}"
mes_actual = meses_es[now.month-1]

def K(v): return f"{v/1000:.1f}K"
def Qm(v):  # Quetzales abreviado
    if v>=1e6: return f"Q {v/1e6:.2f}M"
    if v>=1e3: return f"Q {v/1e3:.0f}K"
    return f"Q {v:,.0f}"
def pctv(v): return round(v/val_total*100,1) if val_total else 0  # % del valor total

# ---- SVG chart ----
def svg_chart():
    vals = list(rec_mes.values()) + list(desp_mes.values())
    mx = max(vals) if vals else 1
    nice = max(16000, math.ceil(mx/16000)*16000)
    scale = 180/nice
    p = ['<line x1="42" y1="190" x2="474" y2="190" style="stroke:var(--grisL);stroke-width:1"/>',
         '<line x1="42" y1="10" x2="42" y2="190" style="stroke:var(--grisL);stroke-width:1"/>',
         '<text x="38" y="194" text-anchor="end" style="font-size:9px;fill:var(--muted)">0</text>']
    for k in range(1,5):
        val = nice*k/4; y = 190 - val*scale
        p.append(f'<line x1="42" y1="{y:.1f}" x2="474" y2="{y:.1f}" style="stroke:var(--grisL);stroke-width:.5"/>')
        p.append(f'<text x="38" y="{y+4:.1f}" text-anchor="end" style="font-size:9px;fill:var(--muted)">{val/1000:.0f}K</text>')
    keys = list(rec_mes.keys())
    for i,mk in enumerate(keys):
        bx = 46 + i*48
        rv = rec_mes[mk]; dv = desp_mes[mk]
        rh = rv*scale; dh = dv*scale
        partial = (i == len(keys)-1)
        op = '.5' if partial else '.85'
        p.append(f'<rect x="{bx}" y="{190-rh:.1f}" width="18" height="{rh:.1f}" rx="2" fill="#1baf7a" opacity="{op}"/>')
        p.append(f'<rect x="{bx+20}" y="{190-dh:.1f}" width="18" height="{dh:.1f}" rx="2" fill="#eb6834" opacity="{op}"/>')
        p.append(f'<text x="{bx+9}" y="{190-rh-3:.1f}" text-anchor="middle" style="font-size:7px;fill:var(--texto)">{K(rv)}</text>')
        p.append(f'<text x="{bx+29}" y="{190-dh-3:.1f}" text-anchor="middle" style="font-size:7px;fill:var(--texto);font-weight:600">{K(dv)}</text>')
        lbl = MLBL[i] + ('*' if partial else '')
        style = ';opacity:.5' if partial else ''
        p.append(f'<text x="{bx+18}" y="204" text-anchor="middle" style="font-size:9px;fill:var(--muted){style}">{lbl}</text>')
    p.append('<rect x="150" y="213" width="8" height="5" rx="1" fill="#1baf7a"/><text x="162" y="218" style="font-size:8px;fill:var(--muted)">Recibido</text>')
    p.append('<rect x="230" y="213" width="8" height="5" rx="1" fill="#eb6834"/><text x="242" y="218" style="font-size:8px;fill:var(--muted)">Despachado</text>')
    p.append('<text x="474" y="218" text-anchor="end" style="font-size:7px;fill:var(--muted)">*Sep parcial</text>')
    return '\n'.join(p)

top_stock_html = ''.join(f'<tr><td>{t["n"]}</td><td style="text-align:right">{t["cajas"]:,}</td></tr>' for t in tops)
tv_html = ''
for i,t in enumerate(tv):
    cls = ' class="hl"' if i==0 else ''
    col = ';color:var(--rojo)' if t['f'] < '2025' else ''
    tv_html += f'<tr{cls}><td>{t["n"]}</td><td style="text-align:right">{t["cajas"]:,}</td><td style="{col.lstrip(";")}">{t["f"]}</td></tr>'

ch = {c: (canales[c], round(canales[c]/canal_total*100,1)) for c in canales}

TPL = open(os.path.join(DIR, '_plantilla_sop.html'), encoding='utf-8').read()
rep = {
 '@@PERIODO@@': f'enero – {mes_actual} {YEAR} (YTD)',
 '@@CORTE@@': fecha_txt,
 '@@RECIBIDO@@': f'{recibido:,}', '@@CITAS@@': str(citas), '@@REC_SKUS@@': str(recibido_skus),
 '@@RECIBIDO_REAL@@': f'{recibido_real:,}', '@@FILL_RECEP@@': f'{fill_recep}', '@@FALTANTE_RECEP@@': f'{faltante_recep:,}',
 '@@STOCK@@': f'{stock:,}', '@@STOCK_SKUS@@': str(stock_skus), '@@STOCK_LPS@@': f'{stock_lps:,}',
 '@@DESP@@': f'{desp:,}', '@@RIESGO@@': f'{riesgo:,}',
 '@@DISPO@@': f'{dispo:,}', '@@DISPO_PCT@@': f'{pct(dispo)}', '@@BLOQ@@': f'{bloq:,}', '@@BLOQ_PCT@@': f'{pct(bloq)}',
 '@@DESP_DIARIO@@': f'{desp_diario:,}', '@@DIAS_INV@@': str(dias_inv),
 '@@PLNER@@': f'{plner:,}', '@@PLNER_SKUS@@': str(plner_skus), '@@PLNER_LPS@@': str(plner_lps),
 '@@VENCIDO@@': f'{vencido:,}', '@@VENCIDO_PCT@@': f'{pct(vencido)}', '@@VENCIDO_SKUS@@': str(vencido_skus),
 '@@VIGENTE@@': f'{vigente:,}', '@@V90@@': f'{v90:,}', '@@V30@@': f'{v30:,}', '@@SINF@@': f'{sinf:,}',
 '@@P_VIG@@': f'{pct(vigente)}', '@@P_VEN@@': f'{pct(vencido)}', '@@P_V90@@': f'{pct(v90)}', '@@P_V30@@': f'{pct(v30)}', '@@P_SINF@@': f'{pct(sinf)}',
 '@@RIESGO_PCT@@': f'{pct(riesgo)}', '@@RIESGO_DIAS@@': str(round(riesgo/desp_diario) if desp_diario else 0),
 '@@CH_MOD@@': f'{ch["Moderno"][0]:,}', '@@CH_MOD_P@@': f'{ch["Moderno"][1]}', '@@CH_MOD_W@@': f'{ch["Moderno"][1]}',
 '@@CH_CED@@': f'{ch["CEDIS"][0]:,}', '@@CH_CED_P@@': f'{ch["CEDIS"][1]}', '@@CH_CED_W@@': f'{ch["CEDIS"][1]}',
 '@@CH_MAY@@': f'{ch["Mayoreo"][0]:,}', '@@CH_MAY_P@@': f'{ch["Mayoreo"][1]}', '@@CH_MAY_W@@': f'{ch["Mayoreo"][1]}',
 '@@CH_TOP_P@@': f'{ch["Top Stores"][1]}', '@@CH_TOP_W@@': f'{ch["Top Stores"][1]}',
 '@@SVG@@': svg_chart(),
 '@@TOP_STOCK@@': top_stock_html, '@@TOP_VENC@@': tv_html,
 # ── Valor en Quetzales ──
 '@@VAL_TOTAL@@': Qm(val_total), '@@RIESGO_Q@@': Qm(val_riesgo),
 '@@PLNER_Q@@': Qm(val_plner), '@@VENCIDO_Q@@': Qm(val_vencido),
 # ── Composicion 100% (cajas% para ancho de barra, + cajas/%/Q por bucket) ──
 '@@CD_W@@': f'{pct(c_disp_c)}', '@@CD_C@@': f'{c_disp_c:,}', '@@CD_P@@': f'{pct(c_disp_c)}', '@@CD_Q@@': Qm(c_disp_v), '@@CD_QP@@': f'{pctv(c_disp_v)}',
 '@@CP_W@@': f'{pct(c_plan_c)}', '@@CP_C@@': f'{c_plan_c:,}', '@@CP_P@@': f'{pct(c_plan_c)}', '@@CP_Q@@': Qm(c_plan_v), '@@CP_QP@@': f'{pctv(c_plan_v)}',
 '@@CV_W@@': f'{pct(c_venc_c)}', '@@CV_C@@': f'{c_venc_c:,}', '@@CV_P@@': f'{pct(c_venc_c)}', '@@CV_Q@@': Qm(c_venc_v), '@@CV_QP@@': f'{pctv(c_venc_v)}',
 '@@CC_W@@': f'{pct(c_cuar_c)}', '@@CC_C@@': f'{c_cuar_c:,}', '@@CC_P@@': f'{pct(c_cuar_c)}', '@@CC_Q@@': Qm(c_cuar_v),
 '@@CO_W@@': f'{pct(c_otros_c)}', '@@CO_C@@': f'{c_otros_c:,}', '@@CO_P@@': f'{pct(c_otros_c)}', '@@CO_Q@@': Qm(c_otros_v),
 # Comprometido total + desglose fino
 '@@COMP_C@@': f'{comp_c:,}', '@@COMP_P@@': f'{pct(comp_c)}', '@@COMP_Q@@': Qm(comp_v),
 '@@BDVE_C@@': f'{c_bdve_c:,}', '@@BDVE_P@@': f'{pct(c_bdve_c)}', '@@BDVE_Q@@': Qm(c_bdve_v),
 '@@BDCDV_C@@': f'{c_bdcdv_c:,}', '@@BDCDV_P@@': f'{pct(c_bdcdv_c)}', '@@BDCDV_Q@@': Qm(c_bdcdv_v),
}
html = TPL
for k,v in rep.items(): html = html.replace(k, v)
out = os.path.join(DIR, 'INFOGRAFIA_SOP_DIRECTOR.html')
with open(out, 'w', encoding='utf-8') as f: f.write(html)
print(f"Generado: {out} ({os.path.getsize(out)/1024:.0f} KB)")
print(f"  Stock {stock:,} | Despacho {desp:,} | Recibido {recibido:,} | Riesgo {riesgo:,} ({pct(riesgo)}%) | PLNER {plner:,} | Rotacion {dias_inv}d")
print("Listo.")
