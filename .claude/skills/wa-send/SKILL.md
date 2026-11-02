---
name: wa-send
description: Reglas DURAS para enviar cualquier mensaje de WhatsApp al equipo desde la PC de Carlos (bodega 1110 Arium). Usar SIEMPRE que haya que mandar WA, crear una tarea programada de envío, modificar un script de envío, o cuando el usuario pida "manda al grupo X". Cubre horario permitido, único sender, whitelist de grupos, prohibición de textos técnicos, chequeo de duplicados y verificación real del envío.
---

# Envío WhatsApp - Reglas duras

Carlos quedó mal frente a dirección por mensajes mandados fuera de horario, duplicados y con textos tipo "test". Esta skill existe para que eso NO vuelva a pasar.

## 1. Horario estricto — PROHIBIDO fuera de ventana

| Destino | Ventana permitida |
|---|---|
| WA al equipo (cualquier grupo) | **3:00 AM a 1:00 PM** |
| Impresión de rótulos Hidravida | 3:00 AM a 12:00 PM |

**Reglas de horario:**
- Si la PC enciende en la tarde, los agentes NO disparan nada
- Toda tarea con efecto visible verifica la hora ANTES de ejecutar su payload
- Agregá al script: `if not (3 <= datetime.now().hour < 13): sys.exit(0)`
- Nunca cambiar el horario sin pedir OK a Carlos

## 2. Único sender permitido

**SOLO** este script puede enviar WA:
```
C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\TRUJILLO.C\PROYECTO MEDICION EMPLEADOS\enviar_grupo_verificado.py
```

Signature:
```powershell
py "...\enviar_grupo_verificado.py" --grupo "<nombre_exacto>" --texto-file <ruta.txt> <imagen.png>
```

Flags útiles:
- `--dry` → simula sin enviar (usar SIEMPRE primero cuando el texto es nuevo)
- `--texto-file` → preferir sobre `--texto` para evitar escape de comillas

**Scripts PROHIBIDOS:**
- ❌ `enviar_cdp.py` — DEPRECATED (bug sticker 24-oct-2026)
- ❌ `enviar_imagen_wa.py` — solo respaldo manual, no agendar
- ❌ Cualquier script nuevo con Selenium/pywhatkit/CDP — NO crear

## 3. Whitelist de grupos (hardcoded)

Solo estos 4 grupos están autorizados en el sender:
1. **Recepciones 2026** (código `DVx3epM5Gx00kZvxBESxuv`)
2. **El Frutal - Descargas**
3. **Operaciones C001**
4. **Asistentes**

Grupo fuera de whitelist → el sender lo rechaza. No pedir "agregalo porfa", preguntar a Carlos primero.

## 4. NUNCA textos técnicos a grupos reales

PROHIBIDO enviar a grupos reales:
- ❌ `"test"`, `"prueba"`, `"debug"`, `"fix"`, `"migracion"`, `"ping"`
- ❌ Mensajes con placeholders tipo `{var}` sin reemplazar
- ❌ Stacktraces, exit codes, logs crudos

Para probar:
- ✅ Usar `--dry`
- ✅ Enviar al chat propio de Carlos (+502 3763 9994)
- ✅ Si el texto va a producción, redactarlo profesional desde el primer intento

## 5. Antes de crear tarea nueva — chequear duplicados

ANTES de registrar una tarea programada nueva de envío WA:
```powershell
schtasks /Query /FO LIST /V | Select-String -Pattern "enviar|WA|whatsapp|grupo" -Context 0,3
```

Si ya existe otra tarea al mismo grupo:
- ❌ No crear la segunda
- Preguntar a Carlos si reemplazar o fusionar

Patrón de duplicados que ya pasó: 22 tareas duplicadas de Castro eliminadas 01-ago. No repetir.

## 6. Verificar envío REAL (no confiar en exit code)

Después de disparar el sender:
- Exit 0 ≠ envío exitoso
- Log "sin excepciones" ≠ envío exitoso
- La única verificación válida es: abrir el grupo y confirmar que la imagen/mensaje llegó

Si no se puede verificar (tarea programada corriendo oculta): decir claro "disparé, falta confirmar", NUNCA maquillar con "listo".

## 7. Tarea programada 100% silenciosa

Si el envío va en Scheduler:
- `<Hidden>true</Hidden>` en el XML
- `pythonw.exe` (NUNCA `python.exe`)
- Si usa PowerShell: envolver en `.vbs` con `sh.Run "powershell -WindowStyle Hidden -File ...", 0, False`
- Edge/CDP embebido: `--headless=new` + `CREATE_NO_WINDOW (0x08000000)`

## Flujo correcto end-to-end

1. Validar hora actual está en 3AM-1PM
2. `schtasks /Query` → confirmar que NO hay otra tarea al mismo grupo
3. Redactar texto profesional (NUNCA "test")
4. Correr `--dry` primero
5. Ejecutar envío real
6. Abrir WA Web y verificar que llegó al grupo correcto
7. Reportar a Carlos: "enviado a X, verificado en el grupo a las HH:MM"

## Qué contestar si Carlos pide algo que viola las reglas

- "mandá a Logística ahora" (son las 3PM) → "No puedo, estamos fuera de ventana 3AM-1PM. ¿Lo agendo para mañana 7AM?"
- "crea otro script más simple" → "El único permitido es enviar_grupo_verificado.py, los demás están deprecated por bugs. ¿Modificamos ese?"
- "mandá 'test' al grupo para probar" → "A grupo real no. Lo pruebo con --dry o al chat tuyo (+502 3763 9994). ¿Cuál preferís?"
