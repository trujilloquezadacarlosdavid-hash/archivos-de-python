---
name: reusar-primero
description: ANTES de escribir cualquier integración frágil (WhatsApp, Outlook, Körber SQL, Excel COM, Edge/CDP, Selenium, archivado de PDFs), buscar primero los scripts de Carlos que ya funcionan. Usar SIEMPRE que la tarea involucre una integración que ya haya resuelto antes. Fuerza grep + glob antes de teclear código nuevo. Cubre la lista de agentes probados (Castro, David, FSD, Notas Envío, Carta Aceptación, MTTO Montacargas, Dashboards) y el flujo de búsqueda estándar.
---

# Reusar primero - buscar antes de reescribir

Carlos ya le costó horas (incidente Castro) que un agente ignorara su script funcionando y reescribiera desde cero una integración frágil. Esta skill fuerza el chequeo previo.

## Regla

ANTES de escribir una línea de código para WA, Outlook, Körber, Excel COM, Edge/CDP, Selenium, archivado de PDFs o cualquier integración recurrente:

1. Hacer `Glob` + `Grep` buscando scripts suyos que ya hagan algo parecido
2. Leer el que más se acerca
3. Partir de esa base, no de cero

Si después de buscar NO existe nada → entonces escribir nuevo. Pero decirlo explícito: "no encontré script previo, lo escribo desde cero".

## Catálogo de agentes y scripts probados

Antes de reescribir, consultar si alguno de estos ya cubre o se acerca:

| Dominio | Agente/Script | Ruta base |
|---|---|---|
| **Envío WA** | `enviar_grupo_verificado.py` (único sender) | `...\PROYECTO MEDICION EMPLEADOS\` |
| **Status WA Castro** | 2x/día + Operaciones +4min + Asistentes 1x/día | ver memoria `agente-castro` |
| **Archivar PDFs Outlook (Mondelez)** | David (`ARCHIVAR_MONDELEZ.ps1`) | `...\Notas de envio Mondelez\_AGENTE\` |
| **Archivar Notas Envío** | `ARCHIVAR_MONDELEZ.ps1` v2 | mismo folder |
| **FSD Mondelez diario** | `AGENTE_FSD_MONDELEZ.ps1` | mismo folder |
| **FSD casas (no Mondelez)** | `AGENTE_FSD_CASA.ps1` | mismo folder |
| **Audit FSD diario** | `AUDIT_FSD_DIARIO.ps1` (compara calendario vs enviados) | mismo folder |
| **Carta Aceptación** | borrador Outlook 6PM lun-sab | ver memoria `agente-carta-aceptacion` |
| **MTTO Montacargas** | lector grupo WA + reply Outlook | `...\MTTO Montacargas\` |
| **Dashboard Recepciones** | `ACTUALIZAR_DESDE_SQL.bat` → `sp_DataRecepciones` | ver memoria `dashboard-recepciones-sql-directo` |
| **Dashboard SOP Heinz** | `ACTUALIZAR_HEINZ.bat` → `INFOGRAFIA_SOP_DIRECTOR.html` | `...\INVENTARIO_DIRECTOR\` |
| **Dashboard Picking** | HTML autónomo | ver memoria `dashboard-asignacion-picking` |
| **Monitor Rótulos Hidravida** | `MONITOR_ROTULOS_HIDRAVIDA.ps1` | ver memoria `rotulos-hidravida-filtro-bodega` |
| **Reabasto FEFO** | `ACTUALIZAR_REABASTO.bat` | `Escritorio\REABASTO BLOQUE\` |
| **Conteo Inventario Físico** | `.xlsm` con botón ACTUALIZAR (VBA ADODB) | ver memoria `conteo-inventario-fisico` |
| **Detector Exceso LPs** | tarea c/15min lun-sab | ver memoria `detector-exceso-lps` |
| **Detector Vencidos** | 3 corridas lun-sab (4,7,11 AM) | ver memoria `detector-vencidos-pasillos` |
| **Monitor Tiempos Muertos** | c/10min montacarguistas inactivos | ver memoria `monitor-tiempos-muertos` |
| **Avance Mensual Sábados** | mensaje Recepciones 2026 | ver memoria `avance-mensual-semanal-equipo` |

## Flujo de búsqueda estándar

### 1. Grep por palabra clave

```powershell
# PowerShell desde el repo
Grep -r "palabra_clave" "C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\"
```

Ejemplos de keywords por dominio:
- WA: `enviar_grupo_verificado`, `whatsapp`, `wa.me`, `selenium`
- Outlook: `Outlook.Application`, `GetDefaultFolder`, `Restrict`, `DASL`
- Körber: `172.16.5.43`, `IN_E`, `sp_DataRecepciones`, `HJ_TRANSACTIONS`
- Excel COM: `Excel.Application`, `.Value2`, `Workbooks.Open`

### 2. Glob por patrón de archivo

```
**/*envi*.py        → scripts de envío
**/AGENTE_*.ps1     → agentes PowerShell
**/MONITOR_*.ps1    → monitores
**/ARCHIVAR_*.ps1   → archivadores
**/_fsd_*.json      → configs FSD
```

### 3. Consultar MEMORY.md

El índice de memoria en `C:\Users\trujillo.c\.claude\projects\...\memory\MEMORY.md` lista TODOS los agentes y proyectos. Leerlo primero evita duplicar.

## Qué NO hacer

- ❌ Abrir editor y empezar a escribir "una versión limpia" de algo que ya existe
- ❌ Asumir "mejor hago uno mío que vaya al grano"
- ❌ Decir "voy a hacer un script más simple" sin haber leído el actual
- ❌ Importar `pywhatkit`, `twilio` u otra lib WA — ya está resuelto con `enviar_grupo_verificado.py`
- ❌ Crear tarea programada nueva sin ver si otra ya hace algo parecido (`schtasks /Query`)

## Qué SÍ hacer

- ✅ Mencionar explícito "revisé y encontré X.ps1 que hace Y, parto de ahí"
- ✅ Mencionar explícito "revisé y NO encontré script previo, lo escribo desde cero"
- ✅ Si el script existente tiene bug, repararlo IN PLACE antes que hacer "v2"
- ✅ Si realmente hay que duplicar, documentar el por qué (ej. "este no se puede modificar porque corre en PC Dorcas y no tengo acceso")

## Checklist antes de teclear código nuevo

- [ ] Hice `Grep` por palabra clave relacionada
- [ ] Hice `Glob` por patrón de nombre esperado
- [ ] Revisé MEMORY.md por si hay agente/proyecto ya resuelto
- [ ] Pregunté a Carlos si hay algo que ya resolvió esto
- [ ] Si existe → lo leí y entendí antes de modificar/extender
- [ ] Si NO existe → lo digo explícito antes de escribir desde cero
