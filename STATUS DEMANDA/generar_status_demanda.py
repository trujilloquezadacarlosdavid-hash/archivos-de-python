# -*- coding: utf-8 -*-
"""
STATUS DE INGRESO DE UNIDADES - AREA DE DEMANDA  (vista POR TIPO, sin casa)
===========================================================================
Por segregacion / auditoria, DEMANDA NO debe ver la casa/proveedor.
Solo reporta el TIPO de unidad (A/C/D) que esta pendiente.

  Tipos:  A = a granel | C = requiere cuadrilla | D = paletizada

Flujo:
  - PROGRAMACION.xlsx  -> unidades PLANEADAS por tipo para el dia (el "calendario").
  - "Archivo Mario Perez avance descargas.xlsx" (Korber) -> lo que YA esta en sistema.
  - El tipo de cada cita se deduce internamente con MAPEO_CASAS.xlsx (marca -> tipo);
    la casa/marca NUNCA sale en el reporte.

Estatus por unidad:
  - Cajas Recibido = 0                         -> PENDIENTE DE ASIGNAR
  - 0 < Recibido < Programado                  -> EN DESCARGA (con % de avance)
  - Recibido >= Programado o Cita Completada SI -> FINALIZADO
  - Cancelada = S                              -> CANCELADA
  - Planeada que aun no aparece en sistema     -> NO PRESENTADA  (solo se reporta el tipo)

Genera:
  - STATUS_DEMANDA_<fecha>.xlsx   (resumen por tipo + detalle anonimo)
  - status_demanda.html           (tablero visual por tipo)

Archivos manuales (persisten entre corridas):
  - PROGRAMACION.xlsx   Tipo (A/C/D) | Cantidad | Hora/Nota   <- el calendario del dia
  - MAPEO_CASAS.xlsx    Marca detectada | Casa (interno) | Tipo (A/C/D)  <- clasificacion privada
"""

import os, sys, io, re, json, ctypes, datetime, html, collections

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: falta openpyxl. Instalalo con:  py -3 -m pip install openpyxl")
    sys.exit(1)

# --------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
NOMBRE_FUENTE = "Archivo Mario Perez avance descargas.xlsx"
CANDIDATOS = [
    os.path.join(SCRIPT_DIR, NOMBRE_FUENTE),
    os.path.join(os.path.dirname(SCRIPT_DIR), NOMBRE_FUENTE),
    os.path.join(os.path.dirname(os.path.dirname(SCRIPT_DIR)), NOMBRE_FUENTE),
]
MAPEO_PATH = os.path.join(SCRIPT_DIR, "MAPEO_CASAS.xlsx")
PROG_PATH = os.path.join(SCRIPT_DIR, "PROGRAMACION.xlsx")

TIPOS = ["A", "C", "D"]
TIPO_DESC = {"A": "A granel", "C": "Cuadrilla", "D": "Paletizada", "": "Sin clasificar"}


# --------------------------------------------------------------------------
# Lectura tolerante a archivo bloqueado (Excel/OneDrive abierto)
# --------------------------------------------------------------------------
def leer_bytes_aunque_este_abierto(path):
    GENERIC_READ = 0x80000000
    FILE_SHARE_ALL = 0x1 | 0x2 | 0x4
    OPEN_EXISTING = 3
    FILE_ATTRIBUTE_NORMAL = 0x80
    INVALID = ctypes.c_void_p(-1).value
    CreateFileW = ctypes.windll.kernel32.CreateFileW
    CreateFileW.restype = ctypes.c_void_p
    h = CreateFileW(ctypes.c_wchar_p(path), GENERIC_READ, FILE_SHARE_ALL,
                    None, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None)
    if h == INVALID or h is None:
        with open(path, "rb") as f:
            return f.read()
    try:
        ReadFile = ctypes.windll.kernel32.ReadFile
        buf = ctypes.create_string_buffer(1024 * 1024)
        leido = ctypes.c_ulong(0)
        out = io.BytesIO()
        while True:
            ok = ReadFile(ctypes.c_void_p(h), buf, len(buf), ctypes.byref(leido), None)
            if not ok or leido.value == 0:
                break
            out.write(buf.raw[:leido.value])
        return out.getvalue()
    finally:
        ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(h))


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# --------------------------------------------------------------------------
# Cargar fuente Korber
# --------------------------------------------------------------------------
def cargar_fuente():
    ruta = next((c for c in CANDIDATOS if os.path.exists(c)), None)
    if not ruta:
        print("Aviso: no encontre '%s' (no pasa nada: la asistente cargara el Excel desde el HTML)." % NOMBRE_FUENTE)
        return None
    print("Leyendo:", ruta)
    datos = leer_bytes_aunque_este_abierto(ruta)
    import warnings
    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(io.BytesIO(datos), data_only=True, read_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=12, values_only=True):
            if row and any(str(c).strip() == "Id Cita" for c in row if c is not None):
                return ws
    return wb.worksheets[0]


def localizar_encabezado(ws):
    filas = list(ws.iter_rows(min_row=1, max_row=15, values_only=True))
    for i, row in enumerate(filas):
        if row and any(str(c).strip() == "Id Cita" for c in row if c is not None):
            cols = {str(c).strip(): j for j, c in enumerate(row) if c is not None}
            return i, cols
    raise RuntimeError("No encontre la fila de encabezados con 'Id Cita'.")


def construir_citas(ws):
    hdr_idx, cols = localizar_encabezado(ws)

    def col(*nombres):
        for n in nombres:
            if n in cols:
                return cols[n]
        return None

    c_fecha = col("Fecha Creacion", "Fecha")
    c_cita = col("Id Cita")
    c_canc = col("Cancelada")
    c_comp = col("Cita Completada")
    c_desc = col("DescripcionSKU", "Descripcion SKU", "Descripcion")
    c_env = col("Cajas Enviadas", "Cajas Enviadas ")
    c_rec = col("Cajas Recibido", "Cajas Recibido ", "Cajas Recibidas")

    citas = collections.OrderedDict()
    for r in ws.iter_rows(min_row=hdr_idx + 2, values_only=True):
        if r is None:
            continue
        cita = r[c_cita] if c_cita is not None and c_cita < len(r) else None
        if cita in (None, ""):
            continue
        c = citas.get(cita)
        if c is None:
            c = {"fecha": None, "canc": set(), "comp": set(), "env": 0.0,
                 "rec": 0.0, "marcas": collections.Counter()}
            citas[cita] = c
        if c_fecha is not None and r[c_fecha]:
            c["fecha"] = r[c_fecha]
        if c_canc is not None:
            c["canc"].add(str(r[c_canc]).strip().upper())
        if c_comp is not None:
            c["comp"].add(str(r[c_comp]).strip().upper())
        c["env"] += num(r[c_env]) if c_env is not None else 0.0
        c["rec"] += num(r[c_rec]) if c_rec is not None else 0.0
        if c_desc is not None and r[c_desc]:
            desc = str(r[c_desc]).strip().split()
            if desc:
                marca = desc[0]
                if len(desc) > 1 and len(desc[0]) <= 4:
                    marca = desc[0] + " " + desc[1]
                c["marcas"][marca.upper()] += 1
    return citas


def estatus_de(c):
    env, rec = c["env"], c["rec"]
    pct = (rec / env * 100.0) if env > 0 else (100.0 if rec > 0 else 0.0)
    pct = max(0, min(100, round(pct)))
    if c["canc"] & {"S", "Y", "SI"}:
        return "CANCELADA", "CANCELADA", pct
    if c["comp"] & {"SI", "S", "YES", "Y"}:
        return "FINALIZADO", "FINALIZADO", 100
    if rec <= 0:
        return "PENDIENTE", "PENDIENTE DE ASIGNAR", 0
    if env > 0 and rec >= env:
        return "FINALIZADO", "FINALIZADO", 100
    return "EN_DESCARGA", "EN DESCARGA", pct


# --------------------------------------------------------------------------
# Manuales: MAPEO (marca->tipo) y PROGRAMACION (planeadas por tipo)
# --------------------------------------------------------------------------
def cargar_mapeo():
    mapeo = {}
    if os.path.exists(MAPEO_PATH):
        try:
            ws = openpyxl.load_workbook(MAPEO_PATH, data_only=True).active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue
                marca = str(row[0]).strip().upper()
                tipo = str(row[2]).strip().upper()[:1] if len(row) > 2 and row[2] else ""
                if tipo not in TIPOS:
                    tipo = ""
                mapeo[marca] = tipo
        except Exception as e:
            print("Aviso: no pude leer MAPEO_CASAS.xlsx (%s)." % e)
    return mapeo


def guardar_mapeo(mapeo, casas_ref):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MAPEO"
    ws.append(["Marca detectada", "Casa (uso interno)", "Tipo (A/C/D)"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="222222")
    for marca in sorted(mapeo):
        ws.append([marca, casas_ref.get(marca, ""), mapeo[marca]])
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 14
    ws.append([])
    ws.append(["Asigna el TIPO (A/C/D) a cada marca. Solo se usa internamente para clasificar; la casa NO sale en el reporte."])
    ws.append(["A = a granel    C = requiere cuadrilla    D = paletizada"])
    try:
        wb.save(MAPEO_PATH)
    except PermissionError:
        print("Aviso: MAPEO_CASAS.xlsx esta abierto; no pude guardarlo.")


def crear_plantilla_programacion():
    if os.path.exists(PROG_PATH):
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PROGRAMACION"

    ws["A1"] = "PROGRAMACION DEL DIA  (escribe una letra por celda: A, C o D)"
    ws["A1"].font = Font(bold=True, size=12, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws.merge_cells("A1:J1")
    ws["A2"] = "Acomoda las letras como tu calendario 'Descargas Log'. El script las cuenta solas."
    ws["A2"].font = Font(italic=True, size=10)
    ws.merge_cells("A2:J2")

    # Cuadricula igual al ejemplo del jueves 18:  A x6, D x2 / C x2, D x3 / D x1
    grid = [
        ["A", "A", "A", "A", "A", "A", "D", "D"],
        ["C", "C", "D", "D", "D"],
        ["D"],
    ]
    fila_inicio = 4
    fills = {"A": "B4A7D6", "C": "A2C4C9", "D": "B6D7A8"}
    for i, fila in enumerate(grid):
        for j, letra in enumerate(fila):
            cell = ws.cell(row=fila_inicio + i, column=1 + j, value=letra)
            cell.font = Font(bold=True, size=14)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill("solid", fgColor=fills.get(letra, "FFFFFF"))
            cell.border = Border(*([Side(style="thin", color="888888")] * 4))
        ws.row_dimensions[fila_inicio + i].height = 30
    for col in range(1, 11):
        ws.column_dimensions[get_column_letter(col)].width = 6

    nota = fila_inicio + len(grid) + 1
    ws.cell(row=nota, column=1,
            value="A = a granel    C = requiere cuadrilla    D = paletizada").font = Font(bold=True)
    ws.cell(row=nota + 1, column=1,
            value="Borra/agrega letras segun la programacion real del dia. Solo se cuentan celdas con una sola letra A, C o D.")
    try:
        wb.save(PROG_PATH)
        print("Plantilla creada:", PROG_PATH)
    except PermissionError:
        pass


def cargar_programacion():
    """Cuenta unidades planeadas por tipo. Soporta:
       1) Cuadricula de letras: cada celda con una sola letra A/C/D (como el calendario).
       2) Formato estructurado: encabezado con 'Cantidad' y columnas Tipo|Cantidad.
    """
    plan = collections.OrderedDict((t, 0) for t in TIPOS)
    if not os.path.exists(PROG_PATH):
        return plan
    try:
        ws = openpyxl.load_workbook(PROG_PATH, data_only=True).active
        filas = list(ws.iter_rows(values_only=True))
    except Exception as e:
        print("Aviso: no pude leer PROGRAMACION.xlsx (%s)." % e)
        return plan

    # detectar formato estructurado (Tipo | Cantidad)
    estructurado = any(
        row and any(str(c).strip().upper() == "CANTIDAD" for c in row if c is not None)
        for row in filas[:6])

    if estructurado:
        encontrado_hdr = False
        for row in filas:
            if not row:
                continue
            if not encontrado_hdr:
                if any(str(c).strip().upper() == "CANTIDAD" for c in row if c is not None):
                    encontrado_hdr = True
                continue
            if not row[0]:
                continue
            tipo = str(row[0]).strip().upper()[:1]
            if tipo not in TIPOS:
                continue
            cant = int(num(row[1])) if len(row) > 1 and row[1] not in (None, "") else 1
            plan[tipo] += max(0, cant)
        return plan

    # cuadricula de letras: contar celdas A/C/D
    for row in filas:
        if not row:
            continue
        for c in row:
            if c is None:
                continue
            s = str(c).strip().upper()
            if not s:
                continue
            if s in TIPOS:                       # celda con una sola letra
                plan[s] += 1
            else:                                # tolerar "A A D" o "A,C,D" en una celda
                toks = [t for t in re.split(r"[\s,;/|]+", s) if t]
                if toks and all(t in TIPOS for t in toks):
                    for t in toks:
                        plan[t] += 1
                elif 1 < len(s) <= 30 and all(ch in TIPOS for ch in s):  # "AAAAAA"
                    for ch in s:
                        plan[ch] += 1
    return plan


# --------------------------------------------------------------------------
# Construir unidades en sistema + reconciliar con programacion
# --------------------------------------------------------------------------
def construir_unidades(citas, mapeo):
    unidades = []
    sin_clasificar = 0
    fecha_ref = None
    for cita, c in citas.items():
        marca = c["marcas"].most_common(1)[0][0] if c["marcas"] else ""
        tipo = mapeo.get(marca, "")
        if tipo == "":
            sin_clasificar += 1
        cod, etiqueta, pct = estatus_de(c)
        fecha = c["fecha"]
        if isinstance(fecha, datetime.datetime):
            fecha = fecha.date()
        if fecha and fecha_ref is None:
            fecha_ref = fecha
        unidades.append({
            "tipo": tipo, "prog": int(round(c["env"])), "rec": int(round(c["rec"])),
            "pct": pct, "cod": cod, "etiqueta": etiqueta, "presente": True,
        })
    return unidades, sin_clasificar, fecha_ref


def armar_status(unidades, plan):
    # resumen por tipo
    resumen = collections.OrderedDict()
    todos_tipos = list(TIPOS) + (["" ] if any(u["tipo"] == "" for u in unidades) else [])
    for t in todos_tipos:
        us = [u for u in unidades if u["tipo"] == t]
        en_sistema = len(us)
        en_desc = sum(1 for u in us if u["cod"] == "EN_DESCARGA")
        fin = sum(1 for u in us if u["cod"] == "FINALIZADO")
        pend_asig = sum(1 for u in us if u["cod"] == "PENDIENTE")
        canc = sum(1 for u in us if u["cod"] == "CANCELADA")
        planeada = plan.get(t, 0)
        no_pres = max(0, planeada - (en_sistema - canc))
        resumen[t] = {
            "planeada": planeada, "en_sistema": en_sistema, "en_descarga": en_desc,
            "finalizada": fin, "pend_asignar": pend_asig, "cancelada": canc,
            "no_presentada": no_pres,
            "pendiente_total": pend_asig + no_pres,
        }

    # detalle anonimo (incluye filas sinteticas de NO PRESENTADA)
    orden_cod = {"EN_DESCARGA": 0, "PENDIENTE": 1, "NO_PRESENTADA": 2,
                 "FINALIZADO": 3, "CANCELADA": 4}
    detalle = []
    for t in todos_tipos:
        us = [u for u in unidades if u["tipo"] == t]
        us.sort(key=lambda u: orden_cod.get(u["cod"], 9))
        detalle.extend(us)
        # agregar pendientes no presentadas como filas tipo-only
        for _ in range(resumen[t]["no_presentada"]):
            detalle.append({"tipo": t, "prog": None, "rec": None, "pct": 0,
                            "cod": "NO_PRESENTADA", "etiqueta": "NO PRESENTADA",
                            "presente": False})
    for i, d in enumerate(detalle, start=1):
        d["n"] = i
    return resumen, detalle


# --------------------------------------------------------------------------
# Colores
# --------------------------------------------------------------------------
XL = {
    "EN_DESCARGA": ("FFF2CC", "7F6000"),
    "PENDIENTE":   ("DDEBF7", "1F4E78"),
    "NO_PRESENTADA": ("F8CBAD", "843C0C"),
    "FINALIZADO":  ("C6EFCE", "006100"),
    "CANCELADA":   ("D9D9D9", "555555"),
}
HX = {
    "EN_DESCARGA": ("#fff2cc", "#7f6000", "#f1c232"),
    "PENDIENTE":   ("#ddebf7", "#1f4e78", "#3d85c6"),
    "NO_PRESENTADA": ("#f8cbad", "#843c0c", "#cc4125"),
    "FINALIZADO":  ("#c6efce", "#006100", "#38761d"),
    "CANCELADA":   ("#d9d9d9", "#555555", "#999999"),
}
TIPO_COLOR = {"A": "#b4a7d6", "C": "#a2c4c9", "D": "#b6d7a8", "": "#dddddd"}


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------
def escribir_excel(resumen, detalle, fecha_ref, glob):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "STATUS DEMANDA"
    fstr = fecha_ref.strftime("%d/%m/%Y") if fecha_ref else ""

    ws.merge_cells("A1:F1")
    ws["A1"] = "STATUS DE INGRESO DE UNIDADES - DEMANDA"
    ws["A1"].font = Font(size=15, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="111111")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    ws.merge_cells("A2:F2")
    ws["A2"] = "Fecha: %s     Generado: %s     (sin casa - solo tipo de unidad)" % (
        fstr, datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))
    ws["A2"].font = Font(size=10, italic=True)
    ws["A2"].alignment = Alignment(horizontal="center")

    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---- RESUMEN POR TIPO ----
    r = 4
    ws.cell(row=r, column=1, value="RESUMEN POR TIPO DE UNIDAD").font = Font(bold=True, size=12)
    r += 1
    hdr = ["TIPO", "DESCRIPCION", "PLANEADAS", "EN SISTEMA", "EN DESCARGA",
           "FINALIZADAS", "PEND. ASIGNAR", "NO PRESENTADAS", "PENDIENTE TOTAL"]
    for j, h in enumerate(hdr, start=1):
        cell = ws.cell(row=r, column=j, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="222222")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[r].height = 28
    r += 1
    tot = collections.Counter()
    for t, s in resumen.items():
        vals = [t or "-", TIPO_DESC.get(t, ""), s["planeada"], s["en_sistema"],
                s["en_descarga"], s["finalizada"], s["pend_asignar"],
                s["no_presentada"], s["pendiente_total"]]
        for j, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if j == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            if j == 9 and isinstance(v, int) and v > 0:
                cell.fill = PatternFill("solid", fgColor="F8CBAD")
                cell.font = Font(bold=True, color="843C0C")
        for k in ["planeada", "en_sistema", "en_descarga", "finalizada",
                  "pend_asignar", "no_presentada", "pendiente_total"]:
            tot[k] += s[k]
        r += 1
    # fila total
    totv = ["TOTAL", "", tot["planeada"], tot["en_sistema"], tot["en_descarga"],
            tot["finalizada"], tot["pend_asignar"], tot["no_presentada"], tot["pendiente_total"]]
    for j, v in enumerate(totv, start=1):
        cell = ws.cell(row=r, column=j, value=v)
        cell.border = border
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="EEEEEE")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    r += 2

    # ---- DETALLE ANONIMO ----
    ws.cell(row=r, column=1, value="DETALLE DE UNIDADES (anonimo - sin casa)").font = Font(bold=True, size=12)
    r += 1
    hdr2 = ["#", "TIPO", "CAJAS PROG.", "CAJAS RECIB.", "% AVANCE", "ESTATUS"]
    base_hdr = r
    for j, h in enumerate(hdr2, start=1):
        cell = ws.cell(row=r, column=j, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="222222")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    r += 1
    for d in detalle:
        bg, fg = XL.get(d["cod"], ("FFFFFF", "000000"))
        prog = d["prog"] if d["prog"] is not None else "-"
        rec = d["rec"] if d["rec"] is not None else "-"
        vals = [d["n"], d["tipo"] or "-", prog, rec, d["pct"] / 100.0, d["etiqueta"]]
        for j, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if j == 5:
                cell.number_format = "0%"
            if j == 6:
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.font = Font(bold=True, color=fg)
        if not d["presente"]:
            for j in range(1, 7):
                ws.cell(row=r, column=j).fill = PatternFill("solid", fgColor="FCE4D6")
            ws.cell(row=r, column=6).fill = PatternFill("solid", fgColor=XL["NO_PRESENTADA"][0])
        r += 1
    ws.freeze_panes = "A%d" % (base_hdr + 1)

    anchos = [6, 8, 13, 13, 11, 24, 16, 16, 16]
    for j, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    r += 1
    ws.cell(row=r, column=1, value="TIPOS:").font = Font(bold=True)
    ws.cell(row=r, column=2, value="A = A granel    C = Requiere cuadrilla    D = Paletizada")
    r += 1
    ws.cell(row=r, column=1, value="NOTA:").font = Font(bold=True)
    ws.cell(row=r, column=2,
            value="Reporte SIN casa por auditoria. Filas salmon = unidades planeadas que aun NO se presentan. "
                  "Edita planeadas en PROGRAMACION.xlsx y el tipo de cada marca en MAPEO_CASAS.xlsx.")

    nombre = "STATUS_DEMANDA_%s.xlsx" % (fecha_ref.strftime("%Y-%m-%d") if fecha_ref else "hoy")
    out = os.path.join(SCRIPT_DIR, nombre)
    try:
        wb.save(out)
    except PermissionError:
        out = os.path.join(SCRIPT_DIR, "STATUS_DEMANDA_nuevo.xlsx")
        wb.save(out)
    print("Excel generado:", out)
    return out


# --------------------------------------------------------------------------
# HTML
# --------------------------------------------------------------------------
def escribir_html(unidades, fecha_ref, sin_clasificar, plan, mapeo=None, sheetjs=None):
    """HTML AUTOSERVICIO Y OFFLINE: la asistente carga el Excel (lector incrustado)
    y la programacion (letras o imagen con OCR opcional) desde el propio HTML.
    Todo se calcula en el navegador; el reporte sale solo por tipo (sin casa)."""
    mapeo = mapeo or {}
    fstr = fecha_ref.strftime("%d/%m/%Y") if fecha_ref else ""
    fkey = fecha_ref.strftime("%Y-%m-%d") if fecha_ref else "hoy"
    gen = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

    units_json = json.dumps([
        {"tipo": u["tipo"], "prog": u["prog"], "rec": u["rec"],
         "pct": u["pct"], "cod": u["cod"], "etiqueta": u["etiqueta"], "marca": ""}
        for u in unidades], ensure_ascii=True)

    mapeo_json = json.dumps({k: v for k, v in mapeo.items() if v}, ensure_ascii=True)

    if sheetjs:
        sheetjs_block = "<script>" + sheetjs + "</script>"
    else:
        sheetjs_block = '<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>'

    plan_letters = "\n".join(" ".join([t] * plan.get(t, 0)) for t in TIPOS if plan.get(t, 0) > 0)

    aviso = ""
    if sin_clasificar:
        aviso = ("<div class='warn'>&#9888; %d unidad(es) en sistema sin TIPO asignado "
                 "(aparecen como tipo \"?\"). Para clasificarlas, quien genera el archivo debe "
                 "completar la columna Tipo en MAPEO_CASAS.xlsx.</div>" % sin_clasificar)

    doc = r"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>STATUS DE INGRESO - DEMANDA</title>
<style>
:root{font-family:'Segoe UI',Arial,sans-serif}
*{box-sizing:border-box}
body{margin:0;background:#f0f2f5;color:#1a1a1a}
.wrap{max-width:1100px;margin:0 auto;padding:18px}
header{background:#111;color:#fff;border-radius:12px;padding:18px 22px;display:flex;
  justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}
header h1{margin:0;font-size:20px;letter-spacing:.5px}
header .meta{font-size:13px;opacity:.85;text-align:right}
.warn{background:#fff3cd;border:1px solid #ffe08a;color:#7a5b00;border-radius:10px;
  padding:10px 14px;margin:14px 0;font-size:13px}
.editor{background:#fff;border-radius:12px;padding:14px 18px;margin:16px 0;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.editor h3{margin:0 0 4px;font-size:15px}
.editor p{margin:0 0 8px;font-size:12px;color:#666}
.editor textarea{width:100%;min-height:74px;font-family:'Consolas',monospace;font-size:18px;
  letter-spacing:3px;line-height:1.6;border:1px solid #cfd6dd;border-radius:8px;padding:8px 10px;resize:vertical}
.editrow{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-top:8px}
.count{font-size:14px;font-weight:700}
.count .a{color:#6a4fa0}.count .c{color:#3a7d86}.count .d{color:#4a8a2e}
.btn{border:none;background:#1f4e78;color:#fff;padding:8px 14px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600}
.btn.sec{background:#e6e8eb;color:#333}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin:16px 0}
.kpi{background:#fff;border-radius:12px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,.08);text-align:center}
.kpi-val{font-size:32px;font-weight:700;line-height:1}
.kpi-lbl{font-size:12px;color:#666;margin-top:6px;text-transform:uppercase;letter-spacing:.4px}
.tcards{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin:8px 0 18px}
.tcard{background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.tcard-h{display:flex;align-items:center;gap:12px;padding:12px 16px}
.tletter{font-size:30px;font-weight:800;color:#1a1a1a;background:rgba(255,255,255,.55);
  width:46px;height:46px;border-radius:10px;display:flex;align-items:center;justify-content:center}
.tdesc{font-weight:700;font-size:15px;color:#1a1a1a}
.tbody{padding:8px 16px 14px}
.trow{display:flex;justify-content:space-between;padding:5px 0;font-size:13px;border-bottom:1px solid #f0f0f0}
.trow b{font-variant-numeric:tabular-nums}
.trow.pend{border-bottom:none;margin-top:4px;color:#cc4125;font-weight:700}
.trow.pend b{font-size:18px}
table{width:100%;border-collapse:collapse;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08)}
thead th{background:#222;color:#fff;font-size:12px;text-transform:uppercase;letter-spacing:.4px;
  padding:11px 8px;text-align:center}
tbody td{padding:9px 8px;border-bottom:1px solid #eee;font-size:13px;text-align:center}
td.num{font-variant-numeric:tabular-nums;font-weight:600}
.tchip{display:inline-block;min-width:26px;padding:3px 8px;border-radius:6px;font-weight:800;color:#1a1a1a}
tr.no_presentada td{background:#fff4ee}
tr:hover td{background:#f6f9ff}
.badge{display:inline-block;padding:4px 10px;border-radius:20px;font-size:11px;font-weight:700;white-space:nowrap}
.bar{position:relative;height:18px;background:#eceef0;border-radius:9px;min-width:90px;overflow:hidden}
.bar-fill{position:absolute;left:0;top:0;height:100%;border-radius:9px}
.bar span{position:relative;font-size:11px;font-weight:700;line-height:18px;color:#333}
.dash{color:#bbb}
h2{font-size:15px;color:#333;margin:10px 2px 4px}
.legend{margin:16px 0;font-size:12px;color:#555;background:#fff;border-radius:12px;padding:12px 16px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.tag{display:inline-block;padding:2px 8px;border-radius:6px;margin:0 4px;font-weight:700}
footer{margin:18px 0;text-align:center;color:#999;font-size:12px}
.loaders{display:grid;grid-template-columns:repeat(auto-fit,minmax(290px,1fr));gap:14px;margin:16px 0}
.loadcard{background:#fff;border-radius:12px;padding:14px 18px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.loadcard h3{margin:0 0 4px;font-size:15px;display:flex;align-items:center;gap:8px}
.loadcard h3 i{background:#1f4e78;color:#fff;width:22px;height:22px;border-radius:50%;display:inline-flex;
  align-items:center;justify-content:center;font-style:normal;font-size:13px;flex:none}
.loadcard p{margin:0 0 10px;font-size:12px;color:#666}
.filebtn{display:inline-block;background:#1f4e78;color:#fff;padding:8px 14px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600}
.finfo{display:block;margin-top:8px;font-size:12px;color:#444;white-space:pre-wrap}
.padbtns{display:inline-flex;gap:6px}
.padbtns button{width:34px;height:30px;border:1px solid #cfd6dd;border-radius:6px;cursor:pointer;font-weight:800;background:#fff;font-size:14px}
.padbtns button:hover{background:#eef2f7}
#imgwrap{margin-top:10px}
#progimg{max-width:100%;max-height:300px;border:1px solid #e0e0e0;border-radius:8px;cursor:zoom-in;display:block}
#progimg.zoom{max-height:none;cursor:zoom-out}
.unknownbox{background:#fff8e1;border:1px solid #ffe08a;border-radius:10px;padding:12px 16px;margin:14px 0;font-size:13px}
.unknownbox .um{display:inline-flex;align-items:center;gap:6px;margin:5px 10px 5px 0;background:#fff;border:1px solid #eedca0;border-radius:8px;padding:4px 8px}
.unknownbox .um b{font-size:13px}
.unknownbox .um button{width:26px;height:24px;border:1px solid #cfd6dd;border-radius:5px;cursor:pointer;font-weight:800;background:#fff}
.unknownbox .um button.on{background:#1f4e78;color:#fff;border-color:#1f4e78}
.empty{background:#fff;border-radius:12px;padding:26px;text-align:center;color:#999;box-shadow:0 1px 4px rgba(0,0,0,.08);margin:8px 0}
@media print{.editor .editrow,.btn,.loaders,.filebtn,.padbtns,.unknownbox{display:none}.editor textarea{border:none}}
</style></head>
<body><div class="wrap">
<header>
  <h1>STATUS DE INGRESO DE UNIDADES &middot; DEMANDA</h1>
  <div class="meta">Fecha: <b>__FECHA__</b><br>Actualizado: __GEN__<br>Sin casa &middot; solo tipo</div>
</header>
__AVISO__

<div class="loaders">
  <div class="loadcard">
    <h3><i>1</i> Cargar Excel actualizado</h3>
    <p>Elige tu archivo de descargas (cualquier nombre). Se lee aqu&iacute; mismo en tu navegador &mdash; no se sube a ning&uacute;n lado.</p>
    <label class="filebtn">Elegir Excel&hellip; <input type="file" id="xlsfile" accept=".xlsx,.xlsm,.xls" hidden></label>
    <span id="xlsinfo" class="finfo"></span>
  </div>
  <div class="loadcard">
    <h3><i>2</i> Programaci&oacute;n desde imagen (opcional)</h3>
    <p>Sube la foto del calendario. Con internet intenta leer las letras solo; si no, queda de referencia para escribirlas.</p>
    <label class="filebtn">Elegir imagen&hellip; <input type="file" id="imgfile" accept="image/*" hidden></label>
    <button class="btn" id="ocrbtn" style="display:none;margin-left:8px">Leer letras de la imagen</button>
    <span id="imginfo" class="finfo"></span>
    <div id="imgwrap" style="display:none"><img id="progimg" alt="programacion"></div>
  </div>
</div>
<div id="unknownbox"></div>

<div class="editor">
  <h3><i style="font-style:normal">3</i>. Programaci&oacute;n del d&iacute;a (letras: A, C o D)</h3>
  <p>Escr&iacute;belas o usa los botones, acomodadas como tu calendario "Descargas Log". Los pendientes se recalculan solos.</p>
  <textarea id="prog" spellcheck="false" placeholder="Ej:  A A A A A A D D&#10;C C D D D&#10;D">__PLANLETTERS__</textarea>
  <div class="editrow">
    <span class="padbtns">
      <button onclick="addLetter('A')" title="A granel">A</button>
      <button onclick="addLetter('C')" title="Cuadrilla">C</button>
      <button onclick="addLetter('D')" title="Paletizada">D</button>
      <button onclick="delLetter()" title="Borrar ultima">&larr;</button>
    </span>
    <span class="count" id="count"></span>
    <button class="btn" onclick="window.print()">Imprimir / PDF</button>
    <button class="btn sec" onclick="document.getElementById('prog').value='';recompute()">Limpiar</button>
  </div>
</div>

<div class="kpis" id="kpis"></div>
<h2>Resumen por tipo de unidad</h2>
<div class="tcards" id="cards"></div>
<h2>Detalle de unidades (an&oacute;nimo)</h2>
<table><thead><tr>
<th>#</th><th>Tipo</th><th>Cajas Prog.</th><th>Cajas Recib.</th><th>% Avance</th><th>Estatus</th>
</tr></thead><tbody id="tbody"></tbody></table>
<div class="legend">
<b>Tipos de unidad:</b>
<span class="tag" style="background:#b4a7d6">A = A granel</span>
<span class="tag" style="background:#a2c4c9">C = Requiere cuadrilla</span>
<span class="tag" style="background:#b6d7a8">D = Paletizada</span>
&nbsp;|&nbsp; <b>PENDIENTE DE ASIGNAR</b> = unidad en sistema sin cajas recibidas.
&nbsp;|&nbsp; <b>NO PRESENTADA</b> = unidad planeada que a&uacute;n no aparece en sistema (se reporta solo el tipo).
</div>
<footer>Vista por tipo &middot; sin identificar casa (auditor&iacute;a) &middot; Me Llega, S.A. &middot; archivo aut&oacute;nomo</footer>
</div>

<script>
var UNITS = __UNITS__;
var MAPEO = __MAPEO__;
try{ var _ov = JSON.parse(localStorage.getItem("mapeoDemanda")||"{}"); for(var _k in _ov) MAPEO[_k]=_ov[_k]; }catch(e){}
var FKEY = "statusDemanda_" + "__FKEY__";
var TIPOS = ["A","C","D"];
var TDESC = {A:"A granel", C:"Cuadrilla", D:"Paletizada", "?":"Sin clasificar"};
var TCOLOR = {A:"#b4a7d6", C:"#a2c4c9", D:"#b6d7a8", "?":"#dddddd"};
var HX = {EN_DESCARGA:["#fff2cc","#7f6000","#f1c232"], PENDIENTE:["#ddebf7","#1f4e78","#3d85c6"],
          NO_PRESENTADA:["#f8cbad","#843c0c","#cc4125"], FINALIZADO:["#c6efce","#006100","#38761d"],
          CANCELADA:["#d9d9d9","#555555","#999999"]};
var ORDEN = {EN_DESCARGA:0, PENDIENTE:1, NO_PRESENTADA:2, FINALIZADO:3, CANCELADA:4};

function planFromText(txt){
  var c = {A:0,C:0,D:0};
  txt.toUpperCase().split(/[\s,;/|]+/).forEach(function(tok){
    if(!tok) return;
    if(tok.length===1 && c.hasOwnProperty(tok)) c[tok]++;
    else if(/^[ACD]+$/.test(tok)){ for(var i=0;i<tok.length;i++) c[tok[i]]++; }
  });
  return c;
}
function aggTipo(t){
  var us = UNITS.filter(function(u){ return (u.tipo||"?")===t; });
  return {
    units: us,
    en_sistema: us.length,
    en_descarga: us.filter(function(u){return u.cod==="EN_DESCARGA";}).length,
    finalizada: us.filter(function(u){return u.cod==="FINALIZADO";}).length,
    pend_asignar: us.filter(function(u){return u.cod==="PENDIENTE";}).length,
    cancelada: us.filter(function(u){return u.cod==="CANCELADA";}).length
  };
}
function esc(s){ return String(s).replace(/[&<>]/g,function(x){return {"&":"&amp;","<":"&lt;",">":"&gt;"}[x];}); }

function recompute(){
  var ta = document.getElementById("prog");
  var plan = planFromText(ta.value);
  try{ localStorage.setItem(FKEY, ta.value); }catch(e){}

  document.getElementById("count").innerHTML =
    "<span class='a'>A: "+plan.A+"</span> &nbsp; <span class='c'>C: "+plan.C+"</span> &nbsp; "+
    "<span class='d'>D: "+plan.D+"</span> &nbsp; &middot; &nbsp; Total: "+(plan.A+plan.C+plan.D);

  var tiposMostrar = TIPOS.slice();
  if(UNITS.some(function(u){ return !u.tipo; })) tiposMostrar.push("?");

  var tot = {planeada:0, en_descarga:0, finalizada:0, pend_total:0};
  var cardsHtml = "", detail = [];

  tiposMostrar.forEach(function(t){
    var a = aggTipo(t);
    var planeada = plan[t] || 0;
    var noPres = Math.max(0, planeada - (a.en_sistema - a.cancelada));
    var pendTotal = a.pend_asignar + noPres;
    tot.planeada += planeada; tot.en_descarga += a.en_descarga;
    tot.finalizada += a.finalizada; tot.pend_total += pendTotal;

    var col = TCOLOR[t] || "#ddd";
    cardsHtml +=
      "<div class='tcard'><div class='tcard-h' style='background:"+col+"'>"+
      "<span class='tletter'>"+t+"</span><span class='tdesc'>"+TDESC[t]+"</span></div>"+
      "<div class='tbody'>"+
      "<div class='trow'><span>Planeadas</span><b>"+planeada+"</b></div>"+
      "<div class='trow'><span>En descarga</span><b>"+a.en_descarga+"</b></div>"+
      "<div class='trow'><span>Finalizadas</span><b>"+a.finalizada+"</b></div>"+
      "<div class='trow pend'><span>PENDIENTES</span><b>"+pendTotal+"</b></div>"+
      "</div></div>";

    var us = a.units.slice().sort(function(x,y){ return (ORDEN[x.cod]||9)-(ORDEN[y.cod]||9); });
    us.forEach(function(u){ detail.push(u); });
    for(var k=0;k<noPres;k++)
      detail.push({tipo:t, prog:null, rec:null, pct:0, cod:"NO_PRESENTADA", etiqueta:"NO PRESENTADA"});
  });

  document.getElementById("kpis").innerHTML =
    kpi("Planeadas", tot.planeada, "#111") + kpi("En descarga", tot.en_descarga, "#f1c232") +
    kpi("Finalizadas", tot.finalizada, "#38761d") + kpi("Pendientes (faltan)", tot.pend_total, "#cc4125");
  document.getElementById("cards").innerHTML = cardsHtml;

  var rows = "";
  detail.forEach(function(d, i){
    var c = HX[d.cod] || ["#fff","#000","#999"];
    var t = d.tipo || "?";
    var prog = d.prog!=null ? d.prog.toLocaleString() : "-";
    var rec = d.rec!=null ? d.rec.toLocaleString() : "-";
    var bar = d.prog!=null ?
      "<div class='bar'><div class='bar-fill' style='width:"+d.pct+"%;background:"+c[2]+"'></div><span>"+d.pct+"%</span></div>"
      : "<span class='dash'>-</span>";
    rows += "<tr class='"+d.cod.toLowerCase()+"'><td class='num'>"+(i+1)+"</td>"+
      "<td><span class='tchip' style='background:"+(TCOLOR[t]||"#ddd")+"'>"+t+"</span></td>"+
      "<td class='num'>"+prog+"</td><td class='num'>"+rec+"</td><td>"+bar+"</td>"+
      "<td><span class='badge' style='background:"+c[0]+";color:"+c[1]+"'>"+esc(d.etiqueta)+"</span></td></tr>";
  });
  document.getElementById("tbody").innerHTML = rows;
}
function kpi(l,v,c){ return "<div class='kpi'><div class='kpi-val' style='color:"+c+"'>"+v+"</div><div class='kpi-lbl'>"+l+"</div></div>"; }

function addLetter(L){ var ta=document.getElementById("prog"); var v=ta.value; if(v && !/\s$/.test(v)) v+=" "; ta.value=v+L; recompute(); }
function delLetter(){ var ta=document.getElementById("prog"); ta.value=ta.value.replace(/\s*[ACDacd]\s*$/,""); recompute(); }

function detectMarca(desc){
  var p=String(desc||"").trim().split(/\s+/);
  if(!p[0]) return "";
  var m=p[0]; if(p.length>1 && p[0].length<=4) m=p[0]+" "+p[1];
  return m.toUpperCase();
}
function estatusDe(env,rec,canc,comp){
  var pct = env>0 ? Math.round(rec/env*100) : (rec>0?100:0);
  pct=Math.max(0,Math.min(100,pct));
  if(canc) return ["CANCELADA","CANCELADA",pct];
  if(comp) return ["FINALIZADO","FINALIZADO",100];
  if(rec<=0) return ["PENDIENTE","PENDIENTE DE ASIGNAR",0];
  if(env>0 && rec>=env) return ["FINALIZADO","FINALIZADO",100];
  return ["EN_DESCARGA","EN DESCARGA",pct];
}
function colIndex(hdr,names){ for(var i=0;i<names.length;i++){ for(var j=0;j<hdr.length;j++){ if(String(hdr[j]).trim()===names[i]) return j; } } return -1; }
function numv(x){ var n=parseFloat(x); return isNaN(n)?0:n; }

function parseExcel(buf){
  if(typeof XLSX==="undefined") throw new Error("lector de Excel no cargado");
  var wb=XLSX.read(buf,{type:"array",cellDates:true});
  var rows=null, hdrIdx=-1;
  for(var si=0; si<wb.SheetNames.length && hdrIdx<0; si++){
    var rr=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[si]],{header:1,raw:true,defval:null});
    for(var r=0;r<Math.min(15,rr.length);r++){
      if(rr[r] && rr[r].some(function(c){return String(c).trim()==="Id Cita";})){ rows=rr; hdrIdx=r; break; }
    }
  }
  if(hdrIdx<0) throw new Error("no encontre la hoja con 'Id Cita'");
  var hdr=rows[hdrIdx];
  var cF=colIndex(hdr,["Fecha Creacion","Fecha"]), cC=colIndex(hdr,["Id Cita"]),
      cX=colIndex(hdr,["Cancelada"]), cM=colIndex(hdr,["Cita Completada"]),
      cD=colIndex(hdr,["DescripcionSKU","Descripcion SKU","Descripcion"]),
      cE=colIndex(hdr,["Cajas Enviadas","Cajas Enviadas "]),
      cR=colIndex(hdr,["Cajas Recibido","Cajas Recibido ","Cajas Recibidas"]);
  var map={}, order=[], fecha=null;
  for(var i=hdrIdx+1;i<rows.length;i++){
    var row=rows[i]; if(!row) continue;
    var cita=row[cC]; if(cita==null||cita==="") continue;
    var o=map[cita]; if(!o){ o={env:0,rec:0,canc:false,comp:false,marcas:{}}; map[cita]=o; order.push(cita); }
    if(cF>=0 && row[cF]!=null && !fecha){ var fv=row[cF]; fecha=(fv instanceof Date)? fv.toLocaleDateString() : String(fv).slice(0,10); }
    if(cX>=0){ var xv=String(row[cX]).trim().toUpperCase(); if(xv==="S"||xv==="Y"||xv==="SI") o.canc=true; }
    if(cM>=0){ var mv=String(row[cM]).trim().toUpperCase(); if(mv==="SI"||mv==="S"||mv==="YES"||mv==="Y") o.comp=true; }
    if(cE>=0) o.env+=numv(row[cE]);
    if(cR>=0) o.rec+=numv(row[cR]);
    if(cD>=0 && row[cD]){ var mk=detectMarca(row[cD]); if(mk) o.marcas[mk]=(o.marcas[mk]||0)+1; }
  }
  var units=[], unknown={};
  order.forEach(function(cita){
    var o=map[cita], marca="", best=-1;
    for(var k in o.marcas){ if(o.marcas[k]>best){best=o.marcas[k];marca=k;} }
    var tipo=MAPEO[marca]||"";
    if(!tipo && marca) unknown[marca]=true;
    var st=estatusDe(o.env,o.rec,o.canc,o.comp);
    units.push({tipo:tipo,prog:Math.round(o.env),rec:Math.round(o.rec),pct:st[2],cod:st[0],etiqueta:st[1],marca:marca});
  });
  return {units:units, unknown:Object.keys(unknown), fecha:fecha};
}
function applyMapeo(){ UNITS.forEach(function(u){ if(u.marca) u.tipo=MAPEO[u.marca]||""; }); recompute(); }
function handleUnknown(list){
  var box=document.getElementById("unknownbox");
  list=(list||[]).filter(function(m){ return !MAPEO[m]; });
  if(!list.length){ box.innerHTML=""; box.className=""; return; }
  box.className="unknownbox";
  var h="<b>&#9888; Asigna el tipo a estas marcas</b> (solo una vez; queda guardado): ";
  list.forEach(function(m){
    h+="<span class='um' data-m=\""+m.replace(/"/g,"&quot;")+"\"><b>"+m+"</b>"+
       "<button onclick=\"setTipo(this,'A')\">A</button>"+
       "<button onclick=\"setTipo(this,'C')\">C</button>"+
       "<button onclick=\"setTipo(this,'D')\">D</button></span>";
  });
  box.innerHTML=h;
}
function setTipo(btn,t){
  var um=btn.parentNode, m=um.getAttribute("data-m");
  MAPEO[m]=t;
  try{ var ov=JSON.parse(localStorage.getItem("mapeoDemanda")||"{}"); ov[m]=t; localStorage.setItem("mapeoDemanda",JSON.stringify(ov)); }catch(e){}
  var bs=um.getElementsByTagName("button"); for(var i=0;i<bs.length;i++) bs[i].className="";
  btn.className="on";
  applyMapeo();
}

var imgDataURL=null;
function runOCR(){
  if(!imgDataURL) return;
  var info=document.getElementById("imginfo");
  function go(){
    info.textContent="Leyendo imagen con OCR ...";
    Tesseract.recognize(imgDataURL,"eng").then(function(res){
      var letters=(res.data.text.toUpperCase().match(/[ACD]/g)||[]);
      if(letters.length){ document.getElementById("prog").value=letters.join(" "); recompute();
        info.textContent="✔ Leí "+letters.length+" letras. REVISA y corrige (el OCR no es 100% exacto)."; }
      else info.textContent="No detecté letras claras. Escríbelas viendo la imagen.";
    }).catch(function(e){ info.textContent="OCR no disponible ("+e.message+"). Escribe las letras viendo la imagen."; });
  }
  if(typeof Tesseract!=="undefined"){ go(); return; }
  info.textContent="Cargando lector OCR (necesita internet la 1a vez) ...";
  var s=document.createElement("script");
  s.src="https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
  s.onload=go;
  s.onerror=function(){ info.textContent="Sin internet: el OCR no cargó. Escribe las letras A/C/D viendo la imagen de abajo."; };
  document.head.appendChild(s);
}
function wireFiles(){
  document.getElementById("xlsfile").addEventListener("change", function(e){
    var f=e.target.files[0]; if(!f) return;
    var info=document.getElementById("xlsinfo"); info.textContent="Leyendo "+f.name+" ...";
    var rd=new FileReader();
    rd.onload=function(){
      try{
        var res=parseExcel(new Uint8Array(rd.result));
        UNITS=res.units;
        info.textContent="✔ "+f.name+" — "+res.units.length+" unidades"+(res.fecha?(" ("+res.fecha+")"):"");
        handleUnknown(res.unknown);
        recompute();
      }catch(err){ info.textContent="✖ No pude leer el Excel: "+err.message; }
    };
    rd.readAsArrayBuffer(f);
  });
  document.getElementById("imgfile").addEventListener("change", function(e){
    var f=e.target.files[0]; if(!f) return;
    var rd=new FileReader();
    rd.onload=function(){
      imgDataURL=rd.result;
      document.getElementById("progimg").src=imgDataURL;
      document.getElementById("imgwrap").style.display="block";
      document.getElementById("ocrbtn").style.display="inline-block";
      document.getElementById("imginfo").textContent="Imagen cargada. Pulsa 'Leer letras' o escríbelas viendo la imagen.";
    };
    rd.readAsDataURL(f);
  });
  document.getElementById("progimg").addEventListener("click", function(){ this.classList.toggle("zoom"); });
  document.getElementById("ocrbtn").addEventListener("click", runOCR);
}

(function(){
  var saved=null;
  try{ saved=localStorage.getItem(FKEY); }catch(e){}
  if(saved!=null && saved.trim()!=="") document.getElementById("prog").value=saved;
  document.getElementById("prog").addEventListener("input", recompute);
  wireFiles();
  recompute();
})();
</script>
__SHEETJS__
</body></html>"""

    doc = (doc.replace("__FECHA__", html.escape(fstr))
              .replace("__GEN__", html.escape(gen))
              .replace("__AVISO__", aviso)
              .replace("__PLANLETTERS__", html.escape(plan_letters))
              .replace("__UNITS__", units_json)
              .replace("__MAPEO__", mapeo_json)
              .replace("__FKEY__", fkey))
    doc = doc.replace("__SHEETJS__", sheetjs_block)
    out = os.path.join(SCRIPT_DIR, "status_demanda.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(doc)
    print("HTML generado:", out)
    return out


# --------------------------------------------------------------------------
def cargar_sheetjs():
    """Lee el lector de Excel (SheetJS) para incrustarlo en el HTML (offline)."""
    ruta = os.path.join(SCRIPT_DIR, "lib", "xlsx.full.min.js")
    if os.path.exists(ruta):
        with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    print("Aviso: no encontre lib/xlsx.full.min.js; el boton 'Cargar Excel' usara internet (CDN).")
    return None


def main():
    ws = cargar_fuente()
    citas = construir_citas(ws) if ws is not None else collections.OrderedDict()
    print("Citas en sistema:", len(citas))

    mapeo = cargar_mapeo()
    casas_ref = {}
    nuevas = 0
    for c in citas.values():
        if c["marcas"]:
            marca = c["marcas"].most_common(1)[0][0]
            casas_ref[marca] = marca
            if marca not in mapeo:
                mapeo[marca] = ""
                nuevas += 1
    guardar_mapeo(mapeo, casas_ref)
    if nuevas:
        print("Se agregaron %d marca(s) nueva(s) a MAPEO_CASAS.xlsx (asigna su Tipo)." % nuevas)

    crear_plantilla_programacion()
    plan = cargar_programacion()

    unidades, sin_clasificar, fecha_ref = construir_unidades(citas, mapeo)

    if unidades:
        resumen, detalle = armar_status(unidades, plan)
        tot_prog = sum(u["prog"] for u in unidades)
        tot_rec = sum(u["rec"] for u in unidades)
        glob = int(round(tot_rec / tot_prog * 100)) if tot_prog else 0
        escribir_excel(resumen, detalle, fecha_ref, glob)

    sheetjs = cargar_sheetjs()
    ruta_html = escribir_html(unidades, fecha_ref, sin_clasificar, plan, mapeo, sheetjs)
    try:
        os.startfile(ruta_html)
    except Exception:
        pass

    print("\nHTML autoservicio listo: la asistente carga el Excel y la programacion desde el propio HTML.")
    if sin_clasificar:
        print("  Nota: %d unidad(es) sin tipo; se completan en MAPEO_CASAS.xlsx o dentro del HTML." % sin_clasificar)
    print("Listo.")


if __name__ == "__main__":
    main()
