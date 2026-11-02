---
name: failover-dorcas
description: Patrón de redundancia para agentes/tareas de Carlos que deben seguir corriendo cuando su PC (FRU-LOG-C1G4NN3) está apagada. Usar SIEMPRE que haya que crear una tarea programada crítica, replicar un agente a la PC Dorcas Castro (recepciones@), o decidir si una tarea nueva necesita failover. Cubre el patrón FAILOVER_CHECK.ps1 (ping antes de ejecutar), cuándo poner failover vs cuándo NO (evitar envíos duplicados), sincronización vía GitHub repo clonado, y checklist de validación en ambas PCs.
---

# Failover PC Dorcas - redundancia controlada

Carlos tiene su PC principal (`FRU-LOG-C1G4NN3`) y la PC de Dorcas Castro (buzón `recepciones@`) como respaldo. Los agentes críticos deben correr en ambas sin duplicar envíos.

## Cuándo poner failover (y cuándo NO)

### ✅ SÍ poner failover

- Correo diario al proveedor (ej. Agente FSD Mondelez — si Carlos no vino, igual tiene que salir)
- Carta Aceptación 6PM lun-sab
- Monitor Vencidos Pasillos (3 corridas lun-sab)
- Reporte diario bodega 1110

### ❌ NO poner failover

- Impresión de rótulos Hidravida (solo PC de Carlos tiene la impresora configurada)
- Dashboard que vive en una pantalla específica de la oficina
- Scripts que leen OneDrive de `trujillo.c` (la PC Dorcas tiene `recepciones@`, otro OneDrive)
- Pruebas y scripts en desarrollo

### 🤔 EVALUAR caso por caso

- Envíos WA → riesgo de duplicar si ambas PCs disparan. Default: solo PC de Carlos. Failover solo si Carlos está de vacaciones (ver abajo "modo vacaciones").
- Scripts que escriben a SQL → bajo riesgo (SQL es idempotente), pero chequear.

## Patrón FAILOVER_CHECK.ps1

Script que se ejecuta ANTES del payload real. Decide si la PC actual debe correr o no.

```powershell
# FAILOVER_CHECK.ps1
# Devuelve exit 0 si ESTA PC debe ejecutar el payload, exit 1 si debe saltarlo.
param(
    [string]$PCPrincipal = "FRU-LOG-C1G4NN3",   # PC Carlos
    [string]$PCRespaldo  = "DORCAS-PC"           # PC Dorcas (ajustar nombre real)
)

$esta = $env:COMPUTERNAME

# Si estoy en la PC principal, siempre corro
if ($esta -eq $PCPrincipal) { exit 0 }

# Si estoy en la respaldo, solo corro si la principal NO responde
if ($esta -eq $PCRespaldo) {
    $ping = Test-Connection -ComputerName $PCPrincipal -Count 2 -Quiet -ErrorAction SilentlyContinue
    if ($ping) {
        # Principal viva → saltar
        exit 1
    } else {
        # Principal apagada → tomo el relevo
        exit 0
    }
}

# PC desconocida → no correr (seguro)
exit 1
```

## Uso en el payload real

```powershell
# Al inicio del script del agente
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$scriptDir\FAILOVER_CHECK.ps1"
if ($LASTEXITCODE -ne 0) {
    # Esta PC no debe correr. Salir silencioso.
    exit 0
}

# ... payload real ...
```

En Python:
```python
import subprocess, sys, os
check = subprocess.run(
    ["powershell", "-NoProfile", "-File",
     os.path.join(os.path.dirname(__file__), "FAILOVER_CHECK.ps1")],
    capture_output=True
)
if check.returncode != 0:
    sys.exit(0)
# ... payload ...
```

## Instalación en PC Dorcas

1. **Clonar el repo** (ya configurado):
   ```powershell
   git clone https://github.com/trujilloquezadacarlosdavid-hash/archivos-de-python.git
   ```
2. **Correr `AutoGitHub_Python`** cada 2h en Dorcas también (sincroniza cambios de Carlos)
3. **Registrar las tareas programadas** correspondientes con el mismo nombre
4. **Cada payload debe empezar con `FAILOVER_CHECK.ps1`** (ya instalado)

## Modo vacaciones (opcional)

Si Carlos avisa que estará ausente varios días (vacaciones, enfermedad):
```
# Setear variable de entorno en PC Dorcas
[Environment]::SetEnvironmentVariable("MODO_VACACIONES", "1", "Machine")
```

Y modificar FAILOVER_CHECK.ps1 para que, si `$env:MODO_VACACIONES = "1"`, la PC Dorcas corra TODO (incluso envíos WA) sin importar si la principal responde.

## Trampa #1 — Reloj desincronizado

Si las dos PCs tienen hora distinta, el chequeo horario (3AM-1PM) puede permitir que una envíe algo que la otra ya envió a otra hora. Verificar:
```powershell
w32tm /query /status
```

## Trampa #2 — OneDrive diferente

PC Carlos lee de `C:\Users\trujillo.c\OneDrive - Me Llega, S.A\...`
PC Dorcas lee de `C:\Users\recepciones\OneDrive - Me Llega, S.A\...`

Scripts con rutas hardcoded de Carlos **truenan en Dorcas**. Usar siempre:
```powershell
$oneDrive = $env:OneDrive
# o
$oneDrive = [Environment]::GetFolderPath('UserProfile') + "\OneDrive - Me Llega, S.A"
```

## Trampa #3 — Acceso a buzones Outlook diferentes

Scripts que leen Outlook en Carlos ven `trujillo.c@` por default; en Dorcas ven `recepciones@`. Si el agente necesita un buzón específico, debe iterar `$ns.Folders` y buscar por nombre:
```powershell
$store = $ns.Folders | Where-Object { $_.Name -match 'trujillo|recepciones' } | Select-Object -First 1
```

## Trampa #4 — Sincronización de outputs

Si el agente escribe un log/estado (ej. `_fsd_auto.json`), las dos PCs van a competir:
- PC Carlos corre, actualiza `_fsd_auto.json`, AutoGitHub lo pushea
- PC Dorcas pullea, pero si corre antes del pull va a sobreescribir

**Solución**: estados persistentes van en una **ruta no versionada** o en una **tabla SQL**, no en el repo.

## Checklist antes de dar por lista una tarea con failover

- [ ] FAILOVER_CHECK.ps1 está en el folder del agente
- [ ] El payload invoca FAILOVER_CHECK.ps1 al inicio y respeta su exit code
- [ ] Rutas NO están hardcoded a `trujillo.c` — usan `$env:USERPROFILE` o `$env:OneDrive`
- [ ] Accesos a Outlook iteran Folders por nombre, no asumen `GetDefaultFolder`
- [ ] Estado persistente va en SQL o ruta no versionada (no en el repo)
- [ ] Registrada la MISMA tarea con el MISMO nombre en ambas PCs
- [ ] Probada en PC Carlos (debe correr) y en PC Dorcas con Carlos encendido (debe NO correr)
- [ ] Documentado qué hace la tarea en MEMORY.md
- [ ] Si envía WA → cumple `/wa-send` (sender único, verificar duplicados)
- [ ] Si corre oculta → cumple `/tarea-silenciosa`
