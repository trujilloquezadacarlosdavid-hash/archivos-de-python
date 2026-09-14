#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════╗
║   SERVIDOR LOCAL — Dashboard de Cortes                       ║
║   Uso: python servidor.py                                    ║
║   Luego abre: http://localhost:8050                          ║
╚══════════════════════════════════════════════════════════════╝
"""

import http.server
import json
import os
import re
import webbrowser
import threading
import shutil
from pathlib import Path
from urllib.parse import parse_qs
import pandas as pd
from io import BytesIO

PORT = 8050
GTQ_RATE = 7.65
HISTORICO_FILE = "historico_cortes.xlsx"
DASHBOARD_TEMPLATE = "dashboard_cortes.html"

# ─────────────────────────────────────────────
# DATA PROCESSING (same logic as update_dashboard.py)
# ─────────────────────────────────────────────

def load_and_clean(file_bytes, filename):
    df = pd.read_excel(BytesIO(file_bytes), sheet_name='Sheet1')
    df = df.dropna(subset=['Número de Orden'])
    df['Monto'] = pd.to_numeric(df['Monto'], errors='coerce').fillna(0)
    df['Corte'] = pd.to_numeric(df['Corte'], errors='coerce').fillna(0)
    df['Canal'] = df['Canal'].fillna('Sin Canal')
    df['Motivo de Corte'] = df['Motivo de Corte'].fillna('Sin Motivo')
    df['Cliente'] = df['Cliente'].fillna('Sin Cliente')
    # Detect date from filename
    m = re.search(r'(\d{2})_(\d{2})_(\d{4})', filename)
    fecha = f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else pd.Timestamp.now().strftime('%Y-%m-%d')
    return df, fecha

def compute_all(df, fecha):
    R = GTQ_RATE
    monto_total = round(df['Monto'].sum(), 2)

    canal_monto = df.groupby('Canal')['Monto'].sum().round(2).to_dict()
    canal_corte = {k: int(v) for k,v in df.groupby('Canal')['Corte'].sum().items()}
    cliente_monto = df.groupby('Cliente')['Monto'].sum().round(2).sort_values(ascending=False).head(10).to_dict()
    cliente_corte = {k: int(v) for k,v in df.groupby('Cliente')['Corte'].sum().sort_values(ascending=False).head(10).items()}
    motivo_monto = df.groupby('Motivo de Corte')['Monto'].sum().round(2).to_dict()
    motivo_corte = {k: int(v) for k,v in df.groupby('Motivo de Corte')['Corte'].sum().items()}

    top_sku_monto_s = df.groupby(['Número del Artículo','Description'])['Monto'].sum().round(2).sort_values(ascending=False).head(10)
    top_sku_corte_s = df.groupby(['Número del Artículo','Description'])['Corte'].sum().sort_values(ascending=False).head(10)
    top_sku_monto = [{'sku':i[0],'desc':i[1][:35],'monto':float(v)} for i,v in top_sku_monto_s.items()]
    top_sku_corte = [{'sku':i[0],'desc':i[1][:35],'corte':int(v)} for i,v in top_sku_corte_s.items()]

    alerts = []
    dups = df[df.duplicated(subset=['Número de Orden','Número del Artículo'], keep=False)]
    if len(dups): alerts.append({'tipo':'CRÍTICA','msg':f'{len(dups)} líneas duplicadas (misma orden + SKU)','icon':'🔴'})
    top_corte_s = df.groupby('Description')['Corte'].sum()
    q95 = top_corte_s.quantile(0.95)
    for desc, val in top_corte_s[top_corte_s >= q95].items():
        alerts.append({'tipo':'ALERTA','msg':f'Corte masivo: {desc[:40]} → {int(val):,} unidades','icon':'🟡'})
    si_pct = int(round((df['Motivo de Corte']=='Sin inventario').sum()/len(df)*100))
    if si_pct > 70:
        alerts.append({'tipo':'INFO','msg':f'Sin inventario representa el {si_pct}% de los cortes','icon':'🔵'})
    if not alerts:
        alerts.append({'tipo':'OK','msg':'Sin anomalías detectadas','icon':'🟢'})

    top_canal = max(canal_monto, key=canal_monto.get) if canal_monto else ''
    top_cliente_name = list(cliente_monto.keys())[0] if cliente_monto else ''
    top_sku_u = top_sku_corte[0] if top_sku_corte else None
    insights = [
        f"Canal <b>{top_canal}</b> lidera con Q{canal_monto.get(top_canal,0):,.2f} en valor de corte",
        f"Cliente más impactado: <b>{top_cliente_name}</b> con Q{cliente_monto.get(top_cliente_name,0):,.2f}",
        f"<b>{si_pct}%</b> de los cortes son por sin inventario — revisar reposición",
    ]
    if top_sku_u:
        insights.append(f"SKU crítico: <b>{top_sku_u['desc'][:40]}</b> con {top_sku_u['corte']:,} unidades cortadas")

    # Table rows HTML
    table_rows = []
    for _, row in df.iterrows():
        motivo = row.get('Motivo de Corte','')
        bc = 'badge-danger' if motivo=='Sin inventario' else ('badge-warning' if 'Fecha' in str(motivo) else 'badge-success')
        monto = float(row.get('Monto',0))
        table_rows.append(f"""
        <tr>
          <td><span class="canal-tag">{row.get('Canal','')}</span></td>
          <td class="mono">{row.get('Número de Orden','')}</td>
          <td>{str(row.get('Cliente',''))[:30]}</td>
          <td class="mono">{row.get('Número del Artículo','')}</td>
          <td>{str(row.get('Description',''))[:35]}</td>
          <td class="text-right">{int(row.get('Corte',0)):,}</td>
          <td><span class="badge {bc}">{motivo}</span></td>
          <td class="text-right amount">Q{monto:,.2f}</td>
        </tr>""")

    # Historico update
    df_new = df.copy()
    df_new['Fecha'] = fecha
    if os.path.exists(HISTORICO_FILE):
        df_hist = pd.read_excel(HISTORICO_FILE, sheet_name='HISTORICO')
        df_hist = df_hist[df_hist['Fecha'].astype(str) != fecha]
        df_combined = pd.concat([df_hist, df_new], ignore_index=True)
    else:
        df_combined = df_new
    df_combined.to_excel(HISTORICO_FILE, sheet_name='HISTORICO', index=False)

    df_combined['Fecha'] = pd.to_datetime(df_combined['Fecha'])
    df_combined['Monto'] = pd.to_numeric(df_combined['Monto'], errors='coerce').fillna(0)
    df_combined['Corte'] = pd.to_numeric(df_combined['Corte'], errors='coerce').fillna(0)
    daily = df_combined.groupby('Fecha').agg(
        Monto=('Monto','sum'), Corte=('Corte','sum'),
        Ordenes=('Número de Orden','nunique')
    ).reset_index().sort_values('Fecha')

    historico = {
        'fechas': daily['Fecha'].dt.strftime('%Y-%m-%d').tolist(),
        'monto_diario': [round(x,2) for x in daily['Monto'].tolist()],
        'corte_diario': [int(x) for x in daily['Corte'].tolist()],
        'ordenes_diario': daily['Ordenes'].tolist(),
    }

    return {
        'fecha': fecha,
        'kpis': {
            'fecha': fecha,
            'monto_total': monto_total,
            'monto_total_gtq': round(monto_total * R, 2),
            'gtq_rate': R,
            'ordenes_totales': int(df['Número de Orden'].nunique()),
            'ticket_promedio': round(df.groupby('Número de Orden')['Monto'].sum().mean(), 2),
            'skus_afectados': int(df['Número del Artículo'].nunique()),
            'clientes_activos': int(df['Cliente'].nunique()),
            'corte_total': int(df['Corte'].sum()),
            'canal_mayor_monto': top_canal,
            'cliente_top': top_cliente_name,
            'motivo_principal': max(motivo_corte, key=motivo_corte.get) if motivo_corte else '',
            'lineas_sin_inventario': int((df['Motivo de Corte']=='Sin inventario').sum()),
            'lineas_fecha': int(df['Motivo de Corte'].str.contains('Fecha', na=False).sum()),
            'lineas_recuperado': int((df['Motivo de Corte']=='Recuperado').sum()),
        },
        'canal_monto': canal_monto,
        'canal_corte': canal_corte,
        'cliente_monto': {k:float(v) for k,v in cliente_monto.items()},
        'cliente_corte': cliente_corte,
        'top_sku_monto': top_sku_monto,
        'top_sku_corte': top_sku_corte,
        'motivo_monto': motivo_monto,
        'motivo_corte': motivo_corte,
        'historico': historico,
        'alerts': alerts,
        'insights': insights,
        'table_rows': ''.join(table_rows),
        'total_rows': len(df),
    }

def build_dashboard_html(data):
    """Inject new data into the HTML template."""
    template = open(DASHBOARD_TEMPLATE, 'r', encoding='utf-8').read()

    # Replace DATA object
    new_data_js = json.dumps({k:v for k,v in data.items() if k != 'table_rows'}, ensure_ascii=False)
    template = re.sub(r'const DATA = \{.*?\};', f'const DATA = {new_data_js};', template, flags=re.DOTALL)

    # Replace table body
    template = re.sub(
        r'<tbody id="table-body">.*?</tbody>',
        f'<tbody id="table-body">{data["table_rows"]}</tbody>',
        template, flags=re.DOTALL
    )

    # Update title
    template = template.replace(
        template[template.find('<title>'):template.find('</title>')+8],
        f'<title>Dashboard Reporte de Cortes | {data["fecha"]}</title>'
    )

    return template

# ─────────────────────────────────────────────
# HTTP SERVER
# ─────────────────────────────────────────────

class DashboardHandler(http.server.BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        print(f"  {args[0]} {args[1]}")

    def do_GET(self):
        if self.path == '/' or self.path == '/dashboard':
            self.serve_file(DASHBOARD_TEMPLATE, 'text/html; charset=utf-8')
        elif self.path == '/health':
            self.send_json({'status': 'ok'})
        else:
            self.send_error(404)

    def do_POST(self):
        if self.path == '/upload':
            self.handle_upload()
        else:
            self.send_error(404)

    def handle_upload(self):
        try:
            # Read multipart body
            content_type = self.headers.get('Content-Type', '')
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)

            # Extract boundary
            boundary_match = re.search(r'boundary=(.+)', content_type)
            if not boundary_match:
                self.send_json({'error': 'No boundary found'}, 400); return
            boundary = boundary_match.group(1).strip().encode()

            # Parse multipart
            parts = body.split(b'--' + boundary)
            file_data = None
            filename = 'archivo.xlsx'
            for part in parts:
                if b'filename=' in part:
                    fn_match = re.search(rb'filename="([^"]+)"', part)
                    if fn_match:
                        filename = fn_match.group(1).decode('utf-8', errors='replace')
                    # Body is after double CRLF
                    idx = part.find(b'\r\n\r\n')
                    if idx != -1:
                        file_data = part[idx+4:].rstrip(b'\r\n')
                    break

            if not file_data:
                self.send_json({'error': 'No se recibió archivo'}, 400); return

            print(f"\n📂 Archivo recibido: {filename} ({len(file_data):,} bytes)")

            # Process
            df, fecha = load_and_clean(file_data, filename)
            print(f"📊 Registros: {len(df)} | Fecha: {fecha}")
            data = compute_all(df, fecha)
            html = build_dashboard_html(data)

            # Save updated dashboard
            out_path = f"dashboard_cortes_{fecha}.html"
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(html)
            # Also update the main template
            with open(DASHBOARD_TEMPLATE, 'w', encoding='utf-8') as f:
                f.write(html)

            print(f"✅ Dashboard guardado: {out_path}")
            print(f"   💰 Monto: Q{data['kpis']['monto_total']:,.2f} | ✂️ Corte: {data['kpis']['corte_total']:,} uds")

            self.send_json({
                'ok': True,
                'fecha': fecha,
                'rows': data['total_rows'],
                'monto': data['kpis']['monto_total'],
                'corte': data['kpis']['corte_total'],
                'saved_as': out_path,
                'html': html,   # send back full HTML so browser reloads instantly
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_json({'error': str(e)}, 500)

    def serve_file(self, path, content_type):
        try:
            with open(path, 'rb') as f:
                data = f.read()
            self.send_response(200)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', len(data))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(data)
        except FileNotFoundError:
            self.send_error(404)

    def send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', len(body))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()


def main():
    os.chdir(Path(__file__).parent)

    print("""
╔══════════════════════════════════════════════════════════════╗
║   🚀 Dashboard de Cortes — Servidor Local                    ║
╠══════════════════════════════════════════════════════════════╣
║   URL:  http://localhost:8050                                 ║
║   Para cerrar: Ctrl + C                                      ║
╚══════════════════════════════════════════════════════════════╝
""")

    server = http.server.HTTPServer(('localhost', PORT), DashboardHandler)

    # Open browser automatically after 1 second
    def open_browser():
        import time; time.sleep(1)
        webbrowser.open(f'http://localhost:{PORT}')
    threading.Thread(target=open_browser, daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n\n🛑 Servidor cerrado.")
        server.server_close()


if __name__ == '__main__':
    main()
