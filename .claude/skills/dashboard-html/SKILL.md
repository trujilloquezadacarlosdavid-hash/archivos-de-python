---
name: dashboard-html
description: Patrón estándar para los dashboards HTML autónomos de Carlos (SOP Heinz, DETALLE Heinz, Recepciones, Picking, NC Mondelez, Conteos, Posiciones, Reabasto, Status Demanda). Usar SIEMPRE que haya que crear o modificar un dashboard HTML. Cubre estructura (extractor SQL → plantilla → .bat ACTUALIZAR), auto-refresh, timestamp visible, colores Arium, formato de número GT, no editar el HTML generado sino la plantilla/extractor.
---

# Dashboard HTML - Patrón estándar

Carlos tiene 10+ dashboards HTML autónomos (abrir con doble click, actualizar con un .bat). Esta skill estandariza el patrón para que todos se vean/comporten igual y no haya que reinventar.

## Arquitectura: 3 piezas

```
ACTUALIZAR_<nombre>.bat      ← lo que Carlos ejecuta
    │
    ├─> extractor.py / .ps1  ← lee SQL / archivo, genera JSON
    │
    └─> plantilla.html       ← template que lee el JSON e inserta datos
         (o)
         render.py           ← genera el HTML final a partir de plantilla + JSON

El HTML FINAL que se abre (ej. INFOGRAFIA_SOP_DIRECTOR.html) ES EL OUTPUT.
NUNCA editar el output a mano — editar la plantilla/extractor.
```

## Regla #1 — NO editar el HTML generado

❌ Si Carlos quiere cambiar un color/número/texto del dashboard, NO abrir el HTML generado y editarlo — el próximo `ACTUALIZAR_*.bat` lo pisa.

✅ Editar la **plantilla** (HTML template) o el **extractor** (script que genera los datos).

Ejemplo real: `INVENTARIO_DIRECTOR\INFOGRAFIA_SOP_DIRECTOR.html` se regenera cada vez. La fuente es el `extractor_*.py` + `plantilla_*.html` en el mismo folder.

## Regla #2 — Timestamp visible

Todo dashboard debe mostrar en el header:
```html
<div class="timestamp">
  Generado: 02-nov-2026 14:22 · Fuente: SQL Körber (sp_DataRecepciones)
</div>
```

Esto es parte de la skill `/verificar-efecto` — permite ver de un vistazo si el dashboard está fresco o congelado.

## Regla #3 — Auto-refresh (opcional, con cuidado)

Si el dashboard debe refrescarse solo (ej. está pegado en pantalla en bodega):
```html
<meta http-equiv="refresh" content="300">
```

300 = cada 5 minutos. Solo si el `.bat` corre en loop; si no, el refresh carga el mismo HTML viejo.

## Regla #4 — Colores Arium

Paleta estándar (de los dashboards vigentes):

| Uso | Color | Hex |
|---|---|---|
| Primario (headers, bordes) | Azul Arium | `#003A70` |
| Secundario | Dorado Arium | `#B8860B` |
| Verde (OK, enviado, finalizado) | Verde | `#2E8B57` |
| Rojo (error, faltante, dañado) | Rojo | `#C8102E` |
| Amarillo (pendiente, warning) | Amarillo | `#F4C430` |
| Fondo tarjeta | Blanco/gris claro | `#F8F9FA` |
| Texto principal | Gris oscuro | `#333333` |

Variables CSS:
```css
:root {
  --arium-blue: #003A70;
  --arium-gold: #B8860B;
  --ok: #2E8B57;
  --err: #C8102E;
  --warn: #F4C430;
  --bg: #F8F9FA;
  --text: #333333;
}
```

## Regla #5 — Formato de número GT

❌ `1,234.56` (US)
✅ `1,234` (sin decimales para unidades/cajas) o `1,234.56` (si realmente son dinero)

```javascript
// Formatear:
(1234567).toLocaleString('es-GT')   // "1,234,567"
(1234.5).toLocaleString('es-GT', {minimumFractionDigits:2}) // "1,234.50"
```

Para porcentajes: `97.3%` (siempre con 1 decimal, sin excepción).
Para Quetzales: `Q 12,345` (prefijo Q + espacio).

## Regla #6 — Responsive mínimo (si va a pantalla grande)

Si el dashboard se va a mostrar en TV/pantalla grande de bodega:
```css
@media (min-width: 1920px) {
  body { font-size: 1.5rem; }
  .big-number { font-size: 4rem; }
}
```

## Estructura mínima recomendada

```html
<!DOCTYPE html>
<html lang="es-GT">
<head>
  <meta charset="UTF-8">
  <title>Dashboard X - Bodega 1110</title>
  <!-- Opcional: <meta http-equiv="refresh" content="300"> -->
  <style>
    :root { /* paleta arium */ }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', sans-serif; }
    body { background: var(--bg); color: var(--text); padding: 1rem; }
    header { background: var(--arium-blue); color: white; padding: 1rem; }
    .timestamp { font-size: 0.85rem; opacity: 0.8; margin-top: 0.5rem; }
    .kpi-row { display: flex; gap: 1rem; margin: 1rem 0; }
    .kpi { flex: 1; background: white; padding: 1rem; border-left: 4px solid var(--arium-blue); }
    .big-number { font-size: 2.5rem; font-weight: 700; color: var(--arium-blue); }
    .ok { color: var(--ok); } .err { color: var(--err); } .warn { color: var(--warn); }
    table { width: 100%; border-collapse: collapse; background: white; margin-top: 1rem; }
    th { background: var(--arium-blue); color: white; padding: 0.5rem; text-align: left; }
    td { padding: 0.5rem; border-bottom: 1px solid #eee; }
    tr:nth-child(even) td { background: #FAFAFA; }
  </style>
</head>
<body>
  <header>
    <h1>Dashboard X - Bodega 1110</h1>
    <div class="timestamp">Generado: {{TIMESTAMP}} · Fuente: {{FUENTE}}</div>
  </header>
  <section class="kpi-row">
    <div class="kpi"><div class="big-number">{{KPI1}}</div><div>Etiqueta 1</div></div>
    <div class="kpi"><div class="big-number ok">{{KPI2}}</div><div>Etiqueta 2</div></div>
  </section>
  <table>
    <thead><tr><th>Col 1</th><th>Col 2</th></tr></thead>
    <tbody>{{FILAS}}</tbody>
  </table>
</body>
</html>
```

## .bat ACTUALIZAR estándar

```batch
@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Actualizando dashboard...
python extractor.py
if errorlevel 1 (
  echo ERROR en extractor. Revisar _error.log
  pause
  exit /b 1
)
echo Listo. Abriendo HTML...
start "" "dashboard_x.html"
```

Si se dispara desde tarea programada → cumple `/tarea-silenciosa` (envolver en .vbs, sin `pause`, sin `start`).

## Checklist antes de dar por terminado el dashboard

- [ ] Hay timestamp visible en el header
- [ ] Colores siguen paleta Arium (azul `#003A70` + dorado `#B8860B`)
- [ ] Números usan `toLocaleString('es-GT')`
- [ ] El HTML generado NO se editó a mano (plantilla + extractor son la fuente)
- [ ] Hay `.bat ACTUALIZAR_*` que regenera de un click
- [ ] Si se publica como artifact o se manda por correo, el HTML es autocontenido (sin CDNs externos de red interna)
- [ ] Si usa SQL → cumple `/korber-sql` (bigint, sin TrxStartTime, etc.)
- [ ] Si corre en Scheduler → cumple `/tarea-silenciosa`
