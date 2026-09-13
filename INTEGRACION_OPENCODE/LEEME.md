# Integracion Claude Code <-> OpenCode

Reparacion + failover automatico para OpenCode Desktop.

## Que hace este paquete

1. **AGENTS.md** — el mismo contexto de Claude Code (tu memoria, tus reglas, tus proyectos) copiado al lugar donde OpenCode lo lee al iniciar.
2. **Monitor de rate-limit** — vigila las sesiones activas de Claude Code y, cuando detecta que te quedaste sin tokens, lanza OpenCode con la carpeta del proyecto abierta y **copia tu ultimo prompt al portapapeles** para que solo tengas que pegar y continuar.
3. **Tarea programada** — el monitor arranca solo al iniciar sesion (con 1 min de retraso).
4. **Acceso directo** — "OpenCode (con contexto)" en tu Escritorio para lanzar OpenCode manualmente cuando quieras.

## Instalar

1. Doble click en `INSTALAR.bat`
2. Cuando te pida configurar la API key:
   - Ve a https://console.anthropic.com/settings/keys y genera una key nueva
   - Abre CMD y ejecuta (pegando tu key real):
     ```
     setx ANTHROPIC_API_KEY "sk-ant-api03-TU-KEY-AQUI"
     ```
   - Cierra CMD y abre OpenCode nuevamente
3. Doble click en `INICIAR_MONITOR.bat` (para arrancar el vigilante ahora, sin esperar el proximo login)

## Verificar que funciona

- **AGENTS.md aplicado:** abre OpenCode, empieza una conversacion nueva y preguntale "que sabes de mi?". Deberia responder con tu rol, jefes y proyectos activos.
- **API key OK:** en OpenCode manda un mensaje. Ya no deberia salir "Anthropic API key is missing".
- **Monitor corriendo:** ejecuta en PowerShell:
  ```
  Get-Process powershell | Where-Object { $_.CommandLine -like '*monitor_ratelimit*' }
  ```
- **Log de failover:** `%LOCALAPPDATA%\opencode_failover.log`

## Detener / desinstalar

- **Detener monitor ahora:** doble click en `DETENER_MONITOR.bat`
- **Quitar tarea de login:** en CMD:
  ```
  schtasks /Delete /TN "OpenCode Failover Monitor" /F
  ```
- **Quitar API key:** en CMD:
  ```
  setx ANTHROPIC_API_KEY ""
  ```

## Archivos del paquete

| Archivo | Que hace |
|---|---|
| `AGENTS.md` | Contexto para OpenCode (se copia a `%USERPROFILE%\.config\opencode\AGENTS.md`) |
| `INSTALAR.bat` / `INSTALAR.ps1` | Instalador (copia AGENTS.md, crea tarea, crea acceso directo) |
| `INICIAR_MONITOR.bat` | Arranca el vigilante en background |
| `DETENER_MONITOR.bat` | Detiene el vigilante |
| `monitor_ratelimit.ps1` | El vigilante en si (chequea cada 15s) |
| `LANZAR_OPENCODE_CONTEXTO.bat` | Abre OpenCode con la carpeta actual (se puede llamar suelto) |

## Precios y limites

**Recuerda:** La API key de Anthropic (`console.anthropic.com`) tiene **billing separado por token consumido** de tu suscripcion Claude Pro/Max.
- Ventaja: si tu Claude Code se queda sin tokens del ciclo, OpenCode sigue funcionando con la API key (billing aparte).
- Desventaja: pagas por token en OpenCode. Manten un ojo en el consumo en `console.anthropic.com/settings/usage`.

## Notas

- El monitor NO envia mensajes automaticos a OpenCode. Solo lo lanza y te copia el ultimo prompt al portapapeles. Tu decides que hacer.
- El monitor recuerda que sesiones ya notifico (en `%LOCALAPPDATA%\opencode_failover_state.json`) para no relanzar OpenCode 10 veces por el mismo incidente.
- Solo detecta rate-limit REAL de API (`rate_limit_error`, `usage limit reached`, HTTP 429, `5-hour limit`). No dispara por mera mencion del texto en conversacion.
