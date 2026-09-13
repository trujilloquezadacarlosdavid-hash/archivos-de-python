"""
VALIDAR_OC_FORMULARIOS.py
Busca correos de formulario de entrega en Outlook (hoy),
extrae los Excel, y cruza cada código contra V_OC_CAJAS_DISPONIBLES
para verificar si hay OC disponible para recibir mañana.
"""
import win32com.client
import pyodbc
import openpyxl
import os, sys, tempfile, re
from datetime import datetime, timedelta
from collections import defaultdict

# ── CONFIG ──────────────────────────────────────────────────
SQL_SERVER  = "172.16.5.43"
SQL_DB      = "IN_E"
SQL_DRIVER  = "ODBC Driver 17 for SQL Server"
BODEGA_CD   = "1110"

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
HTML_OUTPUT = os.path.join(SCRIPT_DIR, "VALIDACION_OC_FORMULARIOS.html")

# ── SQL ─────────────────────────────────────────────────────
def conectar_sql():
    conn_str = (
        f"DRIVER={{{SQL_DRIVER}}};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={SQL_DB};"
        "Trusted_Connection=yes"
    )
    return pyodbc.connect(conn_str, timeout=15)

def obtener_disponibilidad(conn, oc_numbers):
    """Obtiene disponibilidad para una lista de OC (busca en todas las bodegas)."""
    if not oc_numbers:
        return {}

    cursor = conn.cursor()
    placeholders = ",".join(["?"] * len(oc_numbers))
    query = f"""
        SELECT OC, Codigo, Cant_OC_CJ, Cant_Recep_CJ, Cajas_Disponibles, Fecha_OC, Bodega
        FROM LOGISTICA.V_OC_CAJAS_DISPONIBLES
        WHERE OC IN ({placeholders})
    """
    cursor.execute(query, oc_numbers)
    rows = cursor.fetchall()

    disp = {}
    for r in rows:
        key = (r.OC, r.Codigo)
        existing = disp.get(key)
        if existing and existing["cajas_disponibles"] >= float(r.Cajas_Disponibles or 0):
            continue
        disp[key] = {
            "oc": r.OC,
            "codigo": r.Codigo,
            "cant_oc_cj": float(r.Cant_OC_CJ or 0),
            "cant_recep_cj": float(r.Cant_Recep_CJ or 0),
            "cajas_disponibles": float(r.Cajas_Disponibles or 0),
            "fecha_oc": r.Fecha_OC,
            "bodega": str(r.Bodega).strip() if r.Bodega else "",
        }
    return disp

def buscar_disponibilidad_alternativa(conn, codigos):
    """Para códigos no encontrados en su OC, busca disponibilidad en cualquier OC."""
    if not codigos:
        return {}
    cursor = conn.cursor()
    placeholders = ",".join(["?"] * len(codigos))
    query = f"""
        SELECT OC, Codigo, Cant_OC_CJ, Cant_Recep_CJ, Cajas_Disponibles, Fecha_OC
        FROM LOGISTICA.V_OC_CAJAS_DISPONIBLES
        WHERE Codigo IN ({placeholders})
          AND Cajas_Disponibles > 0
        ORDER BY Codigo, Cajas_Disponibles DESC
    """
    cursor.execute(query, codigos)
    alt = defaultdict(list)
    for r in cursor.fetchall():
        alt[r.Codigo].append({
            "oc": r.OC,
            "cajas_disponibles": float(r.Cajas_Disponibles or 0),
            "cant_oc_cj": float(r.Cant_OC_CJ or 0),
            "fecha_oc": r.Fecha_OC,
        })
    return alt


# ── OUTLOOK ─────────────────────────────────────────────────
def _es_correo_formulario(subj):
    """Detecta si el asunto corresponde a un formulario/DEM de entrega."""
    s = subj.upper()
    if "FORMULARIO" in s:
        return True
    if "PROCESO IMPORTACIONES" in s:
        return True
    if "DESPACHO" in s:
        return True
    if re.search(r"RECEPCI[OÓ]N\b", s) and "OC" in s:
        return True
    if "DEM" in s and ("ENTREGA" in s or "FRUTAL" in s or "ESCUINTLA" in s):
        return True
    if "ENTREGA" in s and any(kw in s for kw in ["GUATEMALA", "FRUTAL", "ESCUINTLA", "P0", "P08", "P07", "P09", "P10", "P11", "P12"]):
        return True
    return False


def buscar_formularios_hoy():
    """Busca correos con formularios/DEM de entrega recibidos hoy."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    hoy = datetime.now().strftime("%m/%d/%Y")
    manana_str = (datetime.now() + timedelta(days=1)).strftime("%m/%d/%Y")

    filtro = (
        f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{hoy}' "
        f"AND \"urn:schemas:httpmail:datereceived\" < '{manana_str}' "
        f"AND \"urn:schemas:httpmail:hasattachment\" = 1"
    )
    items = inbox.Items.Restrict(filtro)

    formularios = []
    seen_keys = set()

    for mail in items:
        try:
            subj = mail.Subject or ""
            sender = mail.SenderName or ""
        except:
            continue

        if not _es_correo_formulario(subj):
            continue

        xlsx_atts = []
        try:
            for i in range(1, mail.Attachments.Count + 1):
                att = mail.Attachments.Item(i)
                fn = att.FileName.lower()
                if fn.endswith(".xlsx") and ("formulario" in fn or "dem" in fn or "recepcion" in fn or "entrega" in fn):
                    xlsx_atts.append(att)
                elif fn.endswith(".xlsx") and not any(skip in fn for skip in ["reporte de carga", "planificacion", "estatus", "inventario", "cortes", "status"]):
                    xlsx_atts.append(att)
        except:
            continue

        if not xlsx_atts:
            continue

        oc_match = re.search(r"(?:Orden|OC)[:\s]*(\d{10})", subj, re.IGNORECASE)
        oc_subject = oc_match.group(1) if oc_match else None

        for xlsx_att in xlsx_atts:
            dedup_key = xlsx_att.FileName.upper()
            if dedup_key in seen_keys:
                continue
            seen_keys.add(dedup_key)

            formularios.append({
                "subject": subj,
                "sender": sender,
                "received": mail.ReceivedTime,
                "oc_subject": oc_subject,
                "fecha_entrega": None,
                "attachment": xlsx_att,
                "filename": xlsx_att.FileName,
            })

    return formularios


def parsear_formulario_excel(att):
    """Parsea un Excel de formulario y retorna las líneas de detalle."""
    tmp_path = os.path.join(tempfile.gettempdir(), att.FileName)
    att.SaveAsFile(tmp_path)

    wb = openpyxl.load_workbook(tmp_path, data_only=True)
    ws = wb["Formulario"] if "Formulario" in wb.sheetnames else wb.active

    proveedor_code = ws["C6"].value or ""
    proveedor_nombre = ws["D6"].value or ws["D7"].value or ""
    bodega_code = str(ws["C3"].value or "")
    bodega_nombre = ws["D3"].value or ""

    fecha_cita = None
    raw_fecha = ws["H13"].value
    if raw_fecha:
        if isinstance(raw_fecha, datetime):
            fecha_cita = raw_fecha.date()
        else:
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
                try:
                    fecha_cita = datetime.strptime(str(raw_fecha).strip()[:10], fmt).date()
                    break
                except:
                    pass

    if not fecha_cita:
        date_match = re.search(r"(\d{2})[.\-](\d{2})[.\-](\d{4})", tmp_path)
        if date_match:
            try:
                fecha_cita = datetime(
                    int(date_match.group(3)),
                    int(date_match.group(2)),
                    int(date_match.group(1))
                ).date()
            except:
                pass
        if not fecha_cita:
            date_match2 = re.search(r"(\d{2})[.\-](\d{2})[.\-](\d{2,4})", tmp_path)
            if date_match2:
                y = int(date_match2.group(3))
                if y < 100:
                    y += 2000
                try:
                    fecha_cita = datetime(y, int(date_match2.group(2)), int(date_match2.group(1))).date()
                except:
                    pass

    header_row = None
    col_cajas = 8
    for row_idx in range(1, 25):
        cell_val = ws.cell(row=row_idx, column=3).value
        if cell_val and "CODIGO" in str(cell_val).upper():
            header_row = row_idx
            h_val = str(ws.cell(row=row_idx, column=8).value or "").upper()
            if "CAJA" not in h_val:
                col_cajas = 6
            break

    if header_row is None:
        header_row = 20

    lineas = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        oc_val = ws.cell(row=row_idx, column=2).value
        codigo = ws.cell(row=row_idx, column=3).value
        desc = ws.cell(row=row_idx, column=4).value
        uxc = ws.cell(row=row_idx, column=5).value
        cant_env = ws.cell(row=row_idx, column=6).value
        uom = ws.cell(row=row_idx, column=7).value
        cajas_env = ws.cell(row=row_idx, column=col_cajas).value

        if not codigo or not oc_val:
            continue

        oc_str = str(int(oc_val) if isinstance(oc_val, float) else oc_val).strip()
        codigo_str = str(codigo).strip()

        lineas.append({
            "oc_raw": oc_str,
            "codigo": codigo_str,
            "descripcion": str(desc or "").strip(),
            "uxc": uxc,
            "cantidad_enviada": cant_env,
            "uom": str(uom or ""),
            "cajas_enviadas": float(cajas_env) if cajas_env else 0,
        })

    wb.close()
    try:
        os.remove(tmp_path)
    except:
        pass

    return {
        "proveedor_code": proveedor_code,
        "proveedor_nombre": proveedor_nombre,
        "bodega_code": bodega_code,
        "bodega_nombre": bodega_nombre,
        "fecha_cita": fecha_cita,
        "lineas": lineas,
    }


# ── CRUCE ───────────────────────────────────────────────────
def oc_sin_prefijo(oc):
    s = str(oc)
    if s.startswith("OC") or s.startswith("ST"):
        return s[2:]
    return s

def normalizar_oc(oc_raw):
    """Agrega prefijo OC o ST si es necesario para coincidir con la vista."""
    s = str(oc_raw).strip()
    if s.startswith("OC") or s.startswith("ST"):
        return s
    if s.startswith("44"):
        return f"OC{s}"
    if s.startswith("47"):
        return f"ST{s}"
    if s.startswith("41"):
        return s
    return s

def cruzar_formulario(parsed, conn):
    """Cruza las líneas del formulario contra V_OC_CAJAS_DISPONIBLES.
    Busca disponibilidad por código en cualquier OC, no solo la del Excel."""
    codigos_unicos = list(set(lin["codigo"] for lin in parsed["lineas"]))

    cursor = conn.cursor()
    placeholders = ",".join(["?"] * len(codigos_unicos))
    cursor.execute(f"""
        SELECT OC, Codigo, Cant_OC_CJ, Cant_Recep_CJ, Cajas_Disponibles
        FROM LOGISTICA.V_OC_CAJAS_DISPONIBLES
        WHERE Codigo IN ({placeholders})
          AND Cajas_Disponibles > 0
        ORDER BY Codigo, Cajas_Disponibles DESC
    """, codigos_unicos)

    disp_por_codigo = defaultdict(list)
    for r in cursor.fetchall():
        disp_por_codigo[r.Codigo].append({
            "oc": r.OC,
            "cant_oc_cj": float(r.Cant_OC_CJ or 0),
            "cant_recep_cj": float(r.Cant_Recep_CJ or 0),
            "cajas_disponibles": float(r.Cajas_Disponibles or 0),
        })

    resultados = []
    for lin in parsed["lineas"]:
        oc_norm = normalizar_oc(lin["oc_raw"])
        codigo = lin["codigo"]
        cajas_env = lin["cajas_enviadas"]
        opciones = disp_por_codigo.get(codigo, [])

        total_disp = sum(o["cajas_disponibles"] for o in opciones)

        exacta = [o for o in opciones if o["oc"] == oc_norm or o["oc"] == lin["oc_raw"]]

        if exacta:
            mejor = exacta[0]
            disp_exacta = mejor["cajas_disponibles"]
            if disp_exacta >= cajas_env:
                status = "OK"
            elif disp_exacta > 0:
                status = "PARCIAL"
            else:
                status = "SIN_DISPONIBILIDAD"
            resultados.append({
                **lin,
                "oc_norm": oc_norm,
                "cajas_disponibles": disp_exacta,
                "cant_oc_sql": mejor["cant_oc_cj"],
                "cant_recep_sql": mejor["cant_recep_cj"],
                "status": status,
                "oc_alternativa": None,
            })
        elif opciones:
            mejor = opciones[0]
            disp_mejor = mejor["cajas_disponibles"]
            if disp_mejor >= cajas_env:
                status = "EN_OTRA_OC"
            elif disp_mejor > 0:
                status = "PARCIAL"
            else:
                status = "SIN_DISPONIBILIDAD"
            resultados.append({
                **lin,
                "oc_norm": oc_norm,
                "cajas_disponibles": 0 if status == "EN_OTRA_OC" else disp_mejor,
                "cant_oc_sql": mejor["cant_oc_cj"],
                "cant_recep_sql": mejor["cant_recep_cj"],
                "status": status,
                "oc_alternativa": opciones[:3],
            })
        else:
            resultados.append({
                **lin,
                "oc_norm": oc_norm,
                "cajas_disponibles": 0,
                "cant_oc_sql": 0,
                "cant_recep_sql": 0,
                "status": "SIN_DISPONIBILIDAD",
                "oc_alternativa": None,
            })

    # ── Validación inversa: otros productos en las mismas OCs con 0 disponibilidad ──
    # Solo revisar OCs que aparecen pocas veces (<=3) en el DEM para evitar ruido
    from collections import Counter
    conteo_ocs = Counter(normalizar_oc(lin["oc_raw"]) for lin in parsed["lineas"])
    ocs_secundarias = [oc for oc, n in conteo_ocs.items() if n <= 3]
    codigos_ya = set(lin["codigo"] for lin in parsed["lineas"])

    if ocs_secundarias:
        ph = ",".join(["?"] * len(ocs_secundarias))
        cursor.execute(f"""
            SELECT OC, Codigo, Cant_OC_CJ, Cant_Recep_CJ, Cajas_Disponibles
            FROM LOGISTICA.V_OC_CAJAS_DISPONIBLES
            WHERE OC IN ({ph})
        """, ocs_secundarias)
        for r in cursor.fetchall():
            if r.Codigo in codigos_ya:
                continue
            disp = float(r.Cajas_Disponibles or 0)
            cant_oc = float(r.Cant_OC_CJ or 0)
            cant_recep = float(r.Cant_Recep_CJ or 0)
            if disp <= 0 and cant_oc > 0:
                resultados.append({
                    "oc_raw": r.OC,
                    "codigo": r.Codigo,
                    "descripcion": "",
                    "uxc": None,
                    "cantidad_enviada": None,
                    "uom": "",
                    "cajas_enviadas": 0,
                    "oc_norm": r.OC,
                    "cajas_disponibles": 0,
                    "cant_oc_sql": cant_oc,
                    "cant_recep_sql": cant_recep,
                    "status": "OC_AGOTADA",
                    "oc_alternativa": None,
                })
                codigos_ya.add(r.Codigo)

    return resultados


# ── HTML ────────────────────────────────────────────────────
def generar_html(formularios_resultado, timestamp):
    """Genera el dashboard HTML con los resultados."""
    manana = (datetime.now() + timedelta(days=1)).strftime("%d/%m/%Y")
    hoy_str = datetime.now().strftime("%d/%m/%Y %H:%M")

    cards_html = ""

    for fr in formularios_resultado:
        info = fr["info"]
        parsed = fr["parsed"]
        resultados = fr["resultados"]

        lineas_dem = [r for r in resultados if r["status"] != "OC_AGOTADA"]
        lineas_agotadas = [r for r in resultados if r["status"] == "OC_AGOTADA"]

        total = len(lineas_dem)
        ok = sum(1 for r in lineas_dem if r["status"] == "OK")
        parcial = sum(1 for r in lineas_dem if r["status"] == "PARCIAL")
        en_otra = sum(1 for r in lineas_dem if r["status"] == "EN_OTRA_OC")
        sin_disp = sum(1 for r in lineas_dem if r["status"] == "SIN_DISPONIBILIDAD")
        no_encontrado = sum(1 for r in lineas_dem if r["status"] == "NO_ENCONTRADO")
        agotadas = len(lineas_agotadas)

        cumple = ok
        no_cumple = sin_disp + no_encontrado + en_otra

        if no_cumple > 0 or agotadas > 0:
            card_border = "#e74c3c"
            card_icon = "&#10060;"
            card_label = "NO CUMPLE"
        elif parcial > 0:
            card_border = "#f39c12"
            card_icon = "&#9888;"
            card_label = "PARCIAL"
        else:
            card_border = "#27ae60"
            card_icon = "&#10004;"
            card_label = "CUMPLE"

        oc_display = oc_sin_prefijo(resultados[0]["oc_norm"]) if resultados else "?"
        fecha_entrega = info.get("fecha_entrega", "")
        fecha_str = fecha_entrega.strftime("%d/%m/%Y") if fecha_entrega else manana

        rows_html = ""
        for r in lineas_dem:
            if r["status"] == "OK":
                cls = "row-ok"
                icon = "&#10004;"
                status_display = "CUMPLE"
            elif r["status"] == "PARCIAL":
                cls = "row-warn"
                icon = "&#9888;"
                status_display = "PARCIAL"
            else:
                cls = "row-bad"
                icon = "&#10060;"
                status_display = "NO CUMPLE"

            oc_linea = oc_sin_prefijo(r["oc_norm"])

            nota = ""
            if r["status"] == "EN_OTRA_OC" and r.get("oc_alternativa"):
                alt = r["oc_alternativa"][0]
                nota = f"Disponible en OC {oc_sin_prefijo(alt['oc'])} ({alt['cajas_disponibles']:,.0f} cajas)"

            rows_html += f"""
            <tr class="{cls}">
                <td class="status-icon">{icon}</td>
                <td>{oc_linea}</td>
                <td>{r['codigo']}</td>
                <td class="desc">{r['descripcion']}</td>
                <td class="num">{r['cajas_enviadas']:,.0f}</td>
                <td class="num">{r['cajas_disponibles']:,.0f}</td>
                <td class="status-text">{status_display}</td>
                <td class="nota">{nota}</td>
            </tr>"""

        agotadas_html = ""
        if lineas_agotadas:
            agotadas_html = """
            <tr><td colspan="8" style="padding:10px 12px;font-weight:600;font-size:12px;color:#c0392b;background:#fde8e8;border-top:2px solid #e74c3c;">
                &#9888; Otros productos en esta OC con capacidad agotada en SAP
            </td></tr>"""
            for r in lineas_agotadas:
                oc_linea = oc_sin_prefijo(r["oc_norm"])
                desc = r["descripcion"] if r["descripcion"] else f"(Pedido: {r['cant_oc_sql']:,.0f} / Recibido: {r['cant_recep_sql']:,.0f})"
                agotadas_html += f"""
            <tr class="row-bad">
                <td class="status-icon">&#128683;</td>
                <td>{oc_linea}</td>
                <td>{r['codigo']}</td>
                <td class="desc">{desc}</td>
                <td class="num">-</td>
                <td class="num">0</td>
                <td class="status-text">AGOTADA</td>
                <td class="nota"></td>
            </tr>"""

        total_cajas_form = sum(r["cajas_enviadas"] for r in lineas_dem)
        total_cajas_disp = sum(r["cajas_disponibles"] for r in lineas_dem)

        cards_html += f"""
        <div class="card" style="border-left: 5px solid {card_border}">
            <div class="card-header">
                <div class="card-title">
                    <span class="card-icon">{card_icon}</span>
                    <span>{parsed['proveedor_nombre']}</span>
                    <span class="card-badge" style="background:{card_border}">{card_label}</span>
                </div>
                <div class="card-meta">
                    OC: <b>{oc_display}</b> &nbsp;|&nbsp;
                    Entrega: <b>{fecha_str}</b> &nbsp;|&nbsp;
                    Bodega: <b>{parsed['bodega_code']} {parsed['bodega_nombre']}</b> &nbsp;|&nbsp;
                    De: {info['sender']}
                </div>
            </div>
            <div class="kpis">
                <div class="kpi"><span class="kpi-val">{total}</span><span class="kpi-lbl">SKUs</span></div>
                <div class="kpi kpi-ok"><span class="kpi-val">{cumple}</span><span class="kpi-lbl">Cumple</span></div>
                <div class="kpi kpi-warn"><span class="kpi-val">{parcial}</span><span class="kpi-lbl">Parcial</span></div>
                <div class="kpi kpi-bad"><span class="kpi-val">{no_cumple}</span><span class="kpi-lbl">No Cumple</span></div>
                <div class="kpi kpi-bad"><span class="kpi-val">{agotadas}</span><span class="kpi-lbl">Agotada</span></div>
                <div class="kpi"><span class="kpi-val">{total_cajas_form:,.0f}</span><span class="kpi-lbl">Cajas Formulario</span></div>
                <div class="kpi"><span class="kpi-val">{total_cajas_disp:,.0f}</span><span class="kpi-lbl">Cajas Disponibles</span></div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th></th>
                        <th>OC</th>
                        <th>Codigo</th>
                        <th>Descripcion</th>
                        <th>Cajas Env.</th>
                        <th>Cajas Disp.</th>
                        <th>Status</th>
                        <th>Nota</th>
                    </tr>
                </thead>
                <tbody>{rows_html}{agotadas_html}
                </tbody>
            </table>
        </div>
        """

    if not formularios_resultado:
        cards_html = """
        <div class="empty">
            <p>No se encontraron correos de formulario de entrega hoy.</p>
            <p>Se buscan correos con asunto que contenga "FORMULARIO" + "ENTREGA"/"RECEPCION"/"DEM" y adjunto .xlsx</p>
        </div>
        """

    total_forms = len(formularios_resultado)
    total_problemas = sum(
        1 for fr in formularios_resultado
        if any(r["status"] in ("SIN_DISPONIBILIDAD", "NO_ENCONTRADO") for r in fr["resultados"])
    )

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Validacion OC vs Formularios</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background:#f0f2f5; color:#333; padding:20px; }}
.header {{ background: linear-gradient(135deg, #1a237e 0%, #283593 100%); color:#fff; padding:24px 32px; border-radius:12px; margin-bottom:24px; }}
.header h1 {{ font-size:22px; font-weight:600; }}
.header .sub {{ opacity:0.85; margin-top:6px; font-size:14px; }}
.summary {{ display:flex; gap:16px; margin-bottom:24px; }}
.summary-item {{ background:#fff; border-radius:10px; padding:16px 24px; flex:1; text-align:center; box-shadow:0 1px 3px rgba(0,0,0,0.08); }}
.summary-item .val {{ font-size:28px; font-weight:700; }}
.summary-item .lbl {{ font-size:12px; color:#888; text-transform:uppercase; margin-top:4px; }}
.card {{ background:#fff; border-radius:10px; margin-bottom:20px; box-shadow:0 1px 3px rgba(0,0,0,0.08); overflow:hidden; }}
.card-header {{ padding:16px 20px; }}
.card-title {{ display:flex; align-items:center; gap:10px; font-size:16px; font-weight:600; flex-wrap:wrap; }}
.card-icon {{ font-size:20px; }}
.card-badge {{ color:#fff; font-size:11px; padding:3px 10px; border-radius:12px; font-weight:500; }}
.card-meta {{ font-size:13px; color:#666; margin-top:6px; }}
.kpis {{ display:flex; gap:12px; padding:0 20px 16px; flex-wrap:wrap; }}
.kpi {{ background:#f8f9fa; border-radius:8px; padding:10px 16px; text-align:center; min-width:80px; }}
.kpi-val {{ font-size:20px; font-weight:700; display:block; }}
.kpi-lbl {{ font-size:11px; color:#888; text-transform:uppercase; }}
.kpi-ok .kpi-val {{ color:#27ae60; }}
.kpi-warn .kpi-val {{ color:#f39c12; }}
.kpi-info .kpi-val {{ color:#2980b9; }}
.kpi-bad .kpi-val {{ color:#e74c3c; }}
table {{ width:100%; border-collapse:collapse; font-size:13px; }}
thead {{ background:#f8f9fa; }}
th {{ text-align:left; padding:10px 12px; font-weight:600; color:#555; font-size:12px; text-transform:uppercase; border-bottom:2px solid #e0e0e0; }}
td {{ padding:8px 12px; border-bottom:1px solid #f0f0f0; }}
.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
.desc {{ max-width:280px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:#555; }}
.status-icon {{ text-align:center; font-size:16px; }}
.status-text {{ font-weight:600; font-size:12px; }}
.row-ok {{ }}
.row-ok .status-text {{ color:#27ae60; }}
.row-warn {{ background:#fff8e1; }}
.row-warn .status-text {{ color:#f39c12; }}
.row-info {{ background:#e8f4fd; }}
.row-info .status-text {{ color:#2980b9; font-size:11px; }}
.row-bad {{ background:#fde8e8; }}
.row-bad .status-text {{ color:#e74c3c; }}
.nota {{ font-size:11px; color:#2980b9; font-style:italic; max-width:220px; }}
.empty {{ text-align:center; padding:60px 20px; color:#888; }}
.empty p {{ margin:8px 0; }}
.footer {{ text-align:center; color:#aaa; font-size:12px; margin-top:24px; }}
</style>
</head>
<body>
<div class="header">
    <h1>Validacion de OC vs Formularios de Entrega</h1>
    <div class="sub">Verifica disponibilidad de Orden de Compra para las entregas de hoy y manana &mdash; Generado: {hoy_str}</div>
</div>

<div class="summary">
    <div class="summary-item">
        <div class="val">{total_forms}</div>
        <div class="lbl">Formularios</div>
    </div>
    <div class="summary-item">
        <div class="val" style="color:#27ae60">{total_forms - total_problemas}</div>
        <div class="lbl">Sin Problemas</div>
    </div>
    <div class="summary-item">
        <div class="val" style="color:#e74c3c">{total_problemas}</div>
        <div class="lbl">Con Alertas</div>
    </div>
</div>

{cards_html}

<div class="footer">
    Fuente: LOGISTICA.V_OC_CAJAS_DISPONIBLES &mdash; SQL {SQL_SERVER}/{SQL_DB}
</div>
</body>
</html>"""
    return html


# ── MAIN ────────────────────────────────────────────────────
def main():
    print("Buscando formularios de entrega en Outlook (hoy)...")
    formularios = buscar_formularios_hoy()
    print(f"  Encontrados: {len(formularios)} correos con formulario")

    if not formularios:
        print("  No se encontraron formularios hoy.")
        html = generar_html([], datetime.now())
        with open(HTML_OUTPUT, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  Reporte vacio guardado en: {HTML_OUTPUT}")
        os.startfile(HTML_OUTPUT)
        return

    print("Conectando a SQL Server...")
    conn = conectar_sql()

    resultados_todos = []

    manana_date = (datetime.now() + timedelta(days=1)).date()

    for i, form in enumerate(formularios, 1):
        print(f"\n[{i}/{len(formularios)}] {form['filename']}")
        print(f"  Asunto: {form['subject'][:80]}...")

        try:
            parsed = parsear_formulario_excel(form["attachment"])
        except Exception as e:
            print(f"  ERROR al parsear Excel: {e}")
            continue

        fecha_entrega = parsed.get("fecha_cita")
        print(f"  Proveedor: {parsed['proveedor_nombre']} ({parsed['proveedor_code']})")
        print(f"  Bodega: {parsed['bodega_code']} {parsed['bodega_nombre']}")
        print(f"  Fecha cita (Excel H13): {fecha_entrega}")
        print(f"  Lineas: {len(parsed['lineas'])}")

        if not parsed["lineas"]:
            print("  Sin lineas de detalle (no es formulario estandar), saltando...")
            continue

        if parsed["bodega_code"] != BODEGA_CD:
            print(f"  SALTANDO: bodega {parsed['bodega_code']} (solo se valida {BODEGA_CD})")
            continue

        if fecha_entrega and fecha_entrega != manana_date:
            print(f"  SALTANDO: entrega es {fecha_entrega}, no manana ({manana_date})")
            continue

        resultados = cruzar_formulario(parsed, conn)

        ok = sum(1 for r in resultados if r["status"] == "OK")
        parcial = sum(1 for r in resultados if r["status"] == "PARCIAL")
        en_otra = sum(1 for r in resultados if r["status"] == "EN_OTRA_OC")
        sin_disp = sum(1 for r in resultados if r["status"] == "SIN_DISPONIBILIDAD")
        no_enc = sum(1 for r in resultados if r["status"] == "NO_ENCONTRADO")

        print(f"  Resultado: OK={ok}  Parcial={parcial}  EnOtraOC={en_otra}  SinDisp={sin_disp}  NoEncontrado={no_enc}")

        for r in resultados:
            if r["status"] != "OK":
                extra = ""
                if r["status"] == "EN_OTRA_OC" and r.get("oc_alternativa"):
                    extra = f" -> {r['oc_alternativa'][0]['oc']}"
                print(f"    {r['status']:20s} {r['codigo']}  env={r['cajas_enviadas']:.0f}  disp={r['cajas_disponibles']:.0f}  {r['descripcion'][:40]}{extra}")

        resultados_todos.append({
            "info": {
                "subject": form["subject"],
                "sender": form["sender"],
                "received": form["received"],
                "oc_subject": form.get("oc_subject"),
                "fecha_entrega": fecha_entrega,
            },
            "parsed": parsed,
            "resultados": resultados,
        })

    conn.close()

    html = generar_html(resultados_todos, datetime.now())
    with open(HTML_OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\nReporte guardado en: {HTML_OUTPUT}")
    os.startfile(HTML_OUTPUT)

if __name__ == "__main__":
    main()
