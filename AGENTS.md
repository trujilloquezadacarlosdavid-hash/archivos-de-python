# AGENTS.md — Reglas de trabajo para Carlos Trujillo

> Este archivo lo lee OpenCode (y cualquier agente compatible) para conocer el contexto del proyecto antes de responder. Es el equivalente de CLAUDE.md.

## Quién soy

Carlos David Trujillo Quezada — Jefe de Bodega 1110 en Arium (Me Llega, S.A.), Guatemala. Automatizo tareas de bodega (SAP + Körber) con PowerShell y Python. Trabajo en Windows 11, PC `FRU-LOG-C1G4NN3`. Email: trujillo.c@ariumcorp.com.

Perfil IBS: Sanguíneo 79% / Melancólico 71% / Flemático 50% / Colérico 36%.

## Reglas duras (no se negocian)

### 1. Horario estricto de envíos WhatsApp — PROHIBIDO fuera de 3AM–1PM
- WA al equipo: solo entre **3:00 AM y 1:00 PM**
- Impresión de rótulos Hidravida: **3:00 AM a 12:00 PM**
- Si la PC enciende en la tarde, los agentes NO disparan nada
- Toda tarea con efecto visible verifica la hora ANTES de ejecutar su payload

### 2. Un solo sender WA — nunca duplicar
- Único script permitido: `C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\TRUJILLO.C\PROYECTO MEDICION EMPLEADOS\enviar_grupo_verificado.py`
- Signature: `py enviar_grupo_verificado.py --grupo "<nombre>" --texto-file <ruta.txt> <imagen.png>`
- Whitelist hardcoded (4 grupos): **Recepciones 2026**, **El Frutal - Descargas**, **Operaciones C001**, **Asistentes**
- `enviar_cdp.py` está DEPRECATED (bug sticker 24-oct-2026) — NO USAR
- `enviar_imagen_wa.py` es solo respaldo
- Antes de crear CUALQUIER tarea nueva de envío: `schtasks /Query` y verificar que no exista otra al mismo grupo

### 3. NUNCA mandar texto técnico a grupos reales
- JAMÁS enviar `"test"`, `"migracion"`, `"prueba"`, `"debug"`, `"fix"` a grupos reales
- Para probar: usar `--dry` o chat propio (nunca al grupo del equipo)
- Si el texto va a producción: redactarlo profesional desde el inicio

### 4. Auditar antes de entregar
- Después de CADA acción, verificar el efecto REAL (no exit code, no log, no "sin excepciones")
- WA: abrir el grupo y confirmar que la imagen/mensaje realmente llegó
- Cambios visuales: generar con `-Probar` y LEER el PNG
- Si no puedo verificar: decirlo claro, nunca maquillar con "listo"

### 5. Validar alcance antes de analizar — Tarjeta de Alcance
Antes de calcular CUALQUIER análisis sobre datos, entregar primero:
- **Fuente**: tabla/vista exacta
- **Filtros aplicados**: (ej. `status='A'`, TranType 341, WhsCode 1110A001, YEAR=2026)
- **Qué se EXCLUYE y por qué**: dañado/hold (status='H'), scrap, vencido, P20-P23, filas TOTALES/RESUMEN, SKUs sin costo
- **Supuestos NO verificados**: marcarlos para OK antes de calcular

### 6. Reusar lo que ya funciona
- ANTES de escribir integración frágil (WA, Outlook, Körber, Excel): leer scripts míos que YA funcionan
- Agentes probados: **Castro** (WA Selenium/CDP), **David** (archivar PDFs Outlook), **FSD Casas**, scripts de Medición Empleados
- No reescribir desde cero: parte de la versión que ya envía/funciona

### 7. FSD casas — buscar coordinador externo primero
- Casa sin dominio configurado: buscar en Outlook por placa/DPI/OC quién coordina
- NUNCA generar FSD sobre la cadena interna de Logística/Oscar Ayapan
- Patrón: DEMAGUSA→GRUMA, ICASA→MARAVILLA, forwarders que coordinan por el proveedor

### 8. Tareas programadas 100% silenciosas
Toda tarea nueva debe cumplir:
- `<Hidden>true</Hidden>` en el XML
- Python: `pythonw.exe` (nunca `python.exe`)
- PowerShell: envolver en `.vbs` → `wscript.exe "envoltorio.vbs"` con `sh.Run "powershell -WindowStyle Hidden -File ...", 0, False`
- `.bat`: delegar a `.vbs` inmediatamente (`start "" wscript.exe FOO_HIDDEN.vbs & exit /b 0`)
- Edge/Chrome embebido: `--headless=new` + `CREATE_NO_WINDOW (0x08000000)`

### 9. Conexión SQL estándar
```
DRIVER={ODBC Driver 17 for SQL Server};SERVER=172.16.5.43;DATABASE=IN_E;Trusted_Connection=yes;TrustServerCertificate=yes
```
Trampas:
- `SUM(int)` overflow → `SUM(CAST(col AS bigint))`
- `TrxStartTime`/`TrxEndTime` tipo `time` rompen GetSchemaTable → usar `TrxDate`
- Join a `HJ_TRANSACTIONS` con `WHERE Source=x OR Target=x` NO usa índice → partir por `SourceLocID` o UNION ALL

### 10. PowerShell + Excel COM — bug cache de .Value2
Escribir celda-a-celda cachea el tipo de la PRIMERA escritura y truena mezclando strings+números. FIX definitivo: construir arreglo 2D `New-Object 'object[,]' $nr,$nc` con TODO mezclado y escribir el bloque de un solo golpe (`$ws.Range($rng).Value2 = $arr`). Formato (Font.Bold, Interior.Color, NumberFormat) aplicar DESPUÉS por rangos.

## Datos y vistas de Körber

- Bodega física: `C001` (~8,349 posiciones)
- Reporte diario Bodega 1110: SIEMPRE al grupo "Recepciones 2026" (código `DVx3epM5Gx00kZvxBESxuv`)
- Inventario con matrícula: usar `HJ_Inventory` BASE (la vista omite producto sin matrícula para slotting)
- Capacidad real del bloque: NO está en el maestro SQL — está en código mío
- OC/ST 4700 = casa importada (DUCA), NO es traslado

## Carpeta de trabajo y archivos clave

- Working dir principal: `C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON`
- GitHub repo: `archivos-de-python` (usuario `trujilloquezadacarlosdavid-hash`)
- Auto-sync cada 2h (AutoGitHub_Python)
- PC redundante (failover): PC Dorcas Castro (`recepciones@`) — FAILOVER_CHECK.ps1 hace ping antes de ejecutar
- Memoria completa (reglas, agentes, decisiones): `C:\Users\trujillo.c\.claude\projects\C--Users-trujillo-c-OneDrive---Me-Llega--S-A-Escritorio-ARCHIVOS-DE-PHYTON\memory\MEMORY.md`

## Estilo de respuesta

- Español (es-GT)
- Concisa, al grano, sin preámbulos
- Nunca ejecutar operaciones destructivas sin confirmación
- Confirmar riesgos (push, delete, send) antes de actuar
- Si veo algo raro en git status (archivos sospechosos de secretos), avisar antes de commit
