"""
CONTEO RUTAS PREPARADAS - Bodega 1111
Extractor SQL Korber (172.16.5.43/IN_E) -> conteo_data.js

Fuente:
  LOGISTICA.REVISION_CARGAS_CONTEO_1   (primer conteo por auditor)
  LOGISTICA.REVISION_CARGAS_CONTEO_2   (segundo conteo / reconteo)
"""
import pyodbc, json, sys, os, datetime

sys.stdout.reconfigure(encoding='utf-8')
DIR = os.path.dirname(os.path.abspath(__file__))

SQL = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=172.16.5.43;DATABASE=IN_E;"
    "Trusted_Connection=yes;TrustServerCertificate=yes"
)
CENTRO = "1111"


def serial(o):
    if isinstance(o, (datetime.date, datetime.datetime)):
        return o.isoformat()
    if isinstance(o, float) and o != o:
        return None
    return o


def rows(cur):
    cols = [c[0] for c in cur.description]
    out = []
    for r in cur.fetchall():
        out.append({c: serial(v) for c, v in zip(cols, r)})
    return out


print("Conectando SQL...")
conn = pyodbc.connect(SQL, timeout=30)
cur = conn.cursor()

print("Extrayendo primer conteo (REVISION_CARGAS_CONTEO_1)...")
cur.execute("""
    SELECT ID, CENTRO, ALMACEN, ALMACEN_NOMBRE,
           USUARIO_CODIGO, USUARIO_NOMBRE, FECHA_CREACION,
           SKU_CODIGO, SKU_NOMBRE, SKU_BARRA,
           CAJAS, PAQUETES, UNIDADES,
           CONTEO_FISICO, CANTIDAD_SAP, DIFERENCIA,
           ESTADO_NUMERO, ESTADO_NOMBRE,
           TEXTO1, TEXTO2, TEXTO3, ID_W2W
    FROM LOGISTICA.REVISION_CARGAS_CONTEO_1
    WHERE CENTRO = ?
""", CENTRO)
c1 = rows(cur)
print(f"  {len(c1)} filas")

print("Extrayendo segundo conteo (REVISION_CARGAS_CONTEO_2)...")
cur.execute("""
    SELECT ID, CENTRO, ALMACEN, ALMACEN_NOMBRE,
           USUARIO_CODIGO, USUARIO_NOMBRE, FECHA_CREACION,
           SKU_CODIGO, SKU_NOMBRE, SKU_BARRA,
           CAJAS, PAQUETES, UNIDADES,
           CONTEO_FISICO, CANTIDAD_SAP, DIFERENCIA,
           ESTADO_NUMERO, ESTADO_NOMBRE,
           TEXTO1, TEXTO2, TEXTO3,
           CONTEO_2, CANTIDAD_SAP_2, DIFERENCIA_2,
           ID_W2W, UM_MINIMA
    FROM LOGISTICA.REVISION_CARGAS_CONTEO_2
    WHERE CENTRO = ?
""", CENTRO)
c2 = rows(cur)
print(f"  {len(c2)} filas")

conn.close()

data = {
    "centro": CENTRO,
    "generado": datetime.datetime.now().isoformat(timespec='seconds'),
    "conteo1": c1,
    "conteo2": c2,
}

out = os.path.join(DIR, "conteo_data.js")
with open(out, "w", encoding="utf-8") as f:
    f.write("window.CONTEO_DATA = ")
    json.dump(data, f, ensure_ascii=False)
    f.write(";\n")

size_kb = os.path.getsize(out) / 1024
print(f"OK -> conteo_data.js  ({size_kb:.1f} KB)")
print(f"     primer conteo: {len(c1)}   segundo conteo: {len(c2)}")
