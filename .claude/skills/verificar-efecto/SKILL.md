---
name: verificar-efecto
description: DESPUÉS de ejecutar cualquier acción con efecto real en el mundo (envío WA, correo, impresión de rótulos, actualización de dashboard, creación de tarea programada, archivo generado, commit, push, SQL que modifica), verificar el EFECTO REAL antes de reportar como "listo". Prohibido confiar en exit 0, "sin excepciones", "log sin errores". Usar siempre que la tarea cambie algo observable fuera del script.
---

# Verificar efecto real - no confiar en exit code

Esta skill fuerza la regla #4 de AGENTS.md: "Auditar antes de entregar". Carlos quedó mal frente a dirección por reportar "listo" cuando el envío/dashboard/correo nunca llegó realmente.

## Regla

ANTES de decir "listo", "enviado", "actualizado", "generado": verificar el **efecto observable**.

| Acción | Qué NO cuenta | Qué SÍ cuenta como verificación |
|---|---|---|
| Envío WA | exit 0, "sin excepciones" en log, "mensaje encolado" | Abrir el grupo en WA Web y confirmar el mensaje/imagen |
| Correo Outlook | `.Save()` sin error | Abrir Borradores (o Enviados) y leer el correo |
| Impresión rótulo | PDF generado | Confirmar en la cola de impresión o preguntar a Carlos |
| Dashboard HTML | Script sin error | Abrir el HTML y leer timestamp + número clave |
| Tarea programada nueva | `schtasks /Create` returned 0 | `schtasks /Query /TN "nombre"` + próxima ejecución visible + XML Hidden=true |
| Commit git | `git commit` sin error | `git log -1 --stat` muestra los cambios esperados |
| Push git | `git push` sin error visible | `git push` → ver que diga "To <url>" y "main -> main" |
| Archivo generado | script creó el `.xlsx`/`.pdf` | Abrirlo y leer una celda/línea esperada |
| SQL UPDATE | `rowsAffected > 0` | `SELECT` que confirma el valor nuevo |

## Flujo estándar

```
1. Ejecuto la acción
2. Capturo el output / exit code (necesario pero NO suficiente)
3. HAGO LA VERIFICACIÓN REAL (ver tabla arriba)
4. Reporto con proof: "enviado al grupo X, verificado 14:22" (NO "listo")
```

## Casos específicos ya ocurridos

### WA — exit 0 pero no llegó
`enviar_grupo_verificado.py` imprime "✓ Enviado" aunque el mensaje haya quedado atascado en la cola de WA Web. Verificación real: abrir WA Web (puerto 9224 si es el Edge dedicado, o perfil propio) y confirmar bubble.

### Cierre de turno — HTML sin envío
El script generó el HTML pero `config.json` tenía `sendEmail=false` → nunca mandó correo. Exit 0, log limpio, NADIE lo recibió. Verificar siempre en Enviados.

### Dashboard congelado
`ACTUALIZAR_HEINZ.bat` corrió sin error pero el HTML quedó cacheado en el navegador → Carlos ve datos viejos. Verificar: abrir HTML y leer el timestamp "generado el", comparar contra Get-Date.

### Tarea programada "creada" pero no corre
`schtasks /Create` returned 0, pero la tarea quedó deshabilitada por política de Windows. Verificar: `schtasks /Query /TN "nombre" /V | Select-String "Estado"` → debe decir "Listo" o "En ejecución", no "Deshabilitado".

## Qué decir cuando NO se puede verificar

Si el agente corre oculto en background y no puedo abrir el grupo WA en vivo:
- ✅ "Disparé el envío, exit 0, pero no puedo verificar que llegó. ¿Querés que abra WA Web para confirmar?"
- ❌ "Listo, enviado al grupo X"

Si la tarea programada se dispara mañana a las 7AM y hoy es 2PM:
- ✅ "Tarea creada y visible en schtasks. Próxima corrida mañana 7:00 AM, se verifica cuando corra."
- ❌ "Tarea lista y funcionando"

## Checklist antes de reportar "listo"

- [ ] ¿Ejecuté la acción? → sí
- [ ] ¿Exit 0? → sí (necesario, no suficiente)
- [ ] ¿Verifiqué el efecto observable? → sí (nombrar cuál)
- [ ] ¿El timestamp/número clave coincide con lo esperado? → sí
- [ ] En el reporte, ¿incluyo la prueba? (hora, archivo, screenshot, SELECT confirmatorio) → sí

## Reglas duras

- ❌ NUNCA reportar "listo" basado solo en exit 0
- ❌ NUNCA decir "sin errores en el log" como prueba de éxito
- ❌ NUNCA asumir que algo funcionó porque "no lanzó excepción"
- ✅ Si no puedo verificar, decirlo claro y proponer cómo verificar
- ✅ Si Carlos pregunta "¿ya se mandó?", responder con la verificación real, no con lo que mostró el script
