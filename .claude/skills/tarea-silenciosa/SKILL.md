---
name: tarea-silenciosa
description: Plantilla y reglas para crear una tarea programada de Windows 100% silenciosa (sin ventanas negras, sin flashes, sin pop-ups) en la PC de Carlos. Usar SIEMPRE que haya que agendar un script (Python, PowerShell, .bat, Edge/Chrome headless) con schtasks. Cubre Hidden en XML, pythonw, envoltorios .vbs para PowerShell, delegación de .bat a .vbs, flags headless de navegadores embebidos, y el check de horario 3AM-1PM.
---

# Tarea Programada Silenciosa

Carlos trabaja con la PC visible todo el día; una ventana negra que parpadea cada 5 min lo delata frente a dirección. Toda tarea nueva TIENE que ser invisible.

## Reglas duras

Toda tarea nueva cumple TODO esto:

1. **XML con `<Hidden>true</Hidden>`** (no basta la opción "Oculto" del Task Scheduler GUI)
2. **Python → `pythonw.exe`** (nunca `python.exe`)
3. **PowerShell → `.vbs` envoltorio** con `sh.Run "..., 0, False"`
4. **.bat → delegar a `.vbs` en la primera línea** y salir
5. **Edge/Chrome embebido → `--headless=new` + `CREATE_NO_WINDOW (0x08000000)`**
6. **Chequeo horario 3AM-1PM** al inicio del script (si manda WA / imprime rótulos)

## Plantillas

### A) Python puro (ideal)

```xml
<Actions Context="Author">
  <Exec>
    <Command>C:\Users\trujillo.c\AppData\Local\Programs\Python\Python312\pythonw.exe</Command>
    <Arguments>"C:\ruta\mi_script.py"</Arguments>
    <WorkingDirectory>C:\ruta</WorkingDirectory>
  </Exec>
</Actions>
<Settings>
  <Hidden>true</Hidden>
  ...
</Settings>
```

### B) PowerShell → .vbs envoltorio

Crear `envoltorio_hidden.vbs`:
```vbscript
Set sh = CreateObject("WScript.Shell")
sh.Run "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File ""C:\ruta\mi_script.ps1""", 0, False
```

Y la tarea apunta a:
```xml
<Command>wscript.exe</Command>
<Arguments>"C:\ruta\envoltorio_hidden.vbs"</Arguments>
```

Nunca lanzar `powershell.exe` directo desde schtasks aunque le pongás `-WindowStyle Hidden` — igual parpadea.

### C) .bat → delegar inmediato a .vbs

Primera línea del `.bat`:
```batch
@echo off
start "" wscript.exe "C:\ruta\envoltorio_hidden.vbs" & exit /b 0
```

### D) Edge/Chrome headless embebido

Python con Selenium:
```python
from selenium.webdriver.edge.options import Options
opts = Options()
opts.add_argument("--headless=new")
opts.add_argument("--disable-gpu")
opts.add_experimental_option("excludeSwitches", ["enable-logging"])

# Y al lanzar subprocess:
import subprocess
CREATE_NO_WINDOW = 0x08000000
subprocess.Popen([...], creationflags=CREATE_NO_WINDOW)
```

Si se usa Selenium con WA Web (sender), ver skill `/wa-send` — ahí hay perfiles ya configurados.

## Check de horario (al inicio del script)

Si la tarea tiene efecto visible (WA, impresión, correo a grupos), agregar:

```python
import sys
from datetime import datetime
h = datetime.now().hour
if not (3 <= h < 13):  # 3AM-1PM
    sys.exit(0)
```

En PowerShell:
```powershell
$h = (Get-Date).Hour
if ($h -lt 3 -or $h -ge 13) { exit 0 }
```

Esto cubre: PC encendida en la tarde → la tarea no dispara nada.

## Registrar la tarea

Opción A — XML + schtasks:
```powershell
schtasks /Create /TN "Mi Tarea" /XML "C:\ruta\mi_tarea.xml" /F
```

Opción B — Script PowerShell con `Register-ScheduledTask` (más legible):
```powershell
$a = New-ScheduledTaskAction -Execute "wscript.exe" -Argument '"C:\ruta\envoltorio_hidden.vbs"'
$t = New-ScheduledTaskTrigger -Daily -At 7:00AM
$s = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName "Mi Tarea" -Action $a -Trigger $t -Settings $s -User $env:USERNAME -RunLevel Highest -Force
```

## Antes de registrar — chequear duplicados

```powershell
schtasks /Query /FO LIST /V | Select-String -Pattern "Mi.Patron" -Context 0,3
```

Si ya existe una tarea similar, preguntar a Carlos antes de crear la segunda. Patrón de duplicados que ya pasó: 22 tareas duplicadas de Castro eliminadas 01-ago.

## Checklist antes de dar por terminada la tarea

- [ ] XML tiene `<Hidden>true</Hidden>`
- [ ] Ejecutable es `pythonw.exe` o `wscript.exe` (NUNCA `python.exe` o `powershell.exe` directo)
- [ ] Chequeo horario está en el script si aplica (envío WA/rótulos)
- [ ] Probada con la PC visible: NO parpadea ninguna ventana
- [ ] No duplica otra tarea existente
- [ ] FAILOVER_CHECK.ps1 contemplado si la tarea también corre en PC Dorcas Castro (ping antes de ejecutar)
