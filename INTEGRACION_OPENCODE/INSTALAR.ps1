# ============================================================
#  INSTALAR.ps1
#  Instalador de la integracion Claude Code <-> OpenCode
#  Ejecutar: click derecho -> Run with PowerShell
#           o desde CMD: powershell -ExecutionPolicy Bypass -File INSTALAR.ps1
# ============================================================

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Instalador Integracion Claude Code <-> OpenCode" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

# --- Paso 1: Copiar AGENTS.md y opencode.json a la config global de OpenCode ---
$destConfig = "$env:USERPROFILE\.config\opencode"
$destAgents = Join-Path $destConfig "AGENTS.md"
$destOpencodeJson = Join-Path $destConfig "opencode.json"
$srcAgents  = Join-Path $here "AGENTS.md"
$srcOpencodeJson = Join-Path $here "opencode.json"

if (-not (Test-Path $destConfig)) {
    New-Item -ItemType Directory -Path $destConfig -Force | Out-Null
}

if (Test-Path $destAgents) {
    $backup = "$destAgents.bak.$(Get-Date -Format 'yyyyMMddHHmmss')"
    Copy-Item $destAgents $backup
    Write-Host "[1/4] Backup del AGENTS.md previo: $backup" -ForegroundColor Yellow
}
Copy-Item -Path $srcAgents -Destination $destAgents -Force
Write-Host "[1/4] AGENTS.md copiado a $destAgents" -ForegroundColor Green

if (Test-Path $srcOpencodeJson) {
    if (Test-Path $destOpencodeJson) {
        $backupJson = "$destOpencodeJson.bak.$(Get-Date -Format 'yyyyMMddHHmmss')"
        Copy-Item $destOpencodeJson $backupJson
        Write-Host "[1/4] Backup del opencode.json previo: $backupJson" -ForegroundColor Yellow
    }
    Copy-Item -Path $srcOpencodeJson -Destination $destOpencodeJson -Force
    Write-Host "[1/4] opencode.json copiado (modelo default: Groq)" -ForegroundColor Green
}

# --- Paso 2: Chequear GROQ_API_KEY (GRATIS) ---
$groqKey = [Environment]::GetEnvironmentVariable('GROQ_API_KEY', 'User')
if ([string]::IsNullOrWhiteSpace($groqKey) -or $groqKey.Length -lt 30) {
    Write-Host ""
    Write-Host "[2/4] GROQ_API_KEY NO esta configurada (necesaria - es GRATIS)." -ForegroundColor Red
    Write-Host ""
    Write-Host "  Para configurarla (proceso GRATIS, sin tarjeta):" -ForegroundColor Yellow
    Write-Host "  1. Ve a: https://console.groq.com/keys"
    Write-Host "  2. Registrate con tu correo (o Google/GitHub)"
    Write-Host "  3. Click en 'Create API Key' -> nombra: 'OpenCode'"
    Write-Host "  4. COPIA la key completa (empieza con gsk_...)"
    Write-Host "  5. Abre una terminal NUEVA (CMD) y ejecuta:"
    Write-Host ""
    Write-Host '     setx GROQ_API_KEY "gsk_PEGAR-AQUI-TU-KEY-COMPLETA"' -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  6. Cierra CMD y reinicia OpenCode."
    Write-Host ""
    Write-Host "  IMPORTANTE: setx guarda la variable de forma permanente para tu usuario."
    Write-Host "              Necesitas abrir una terminal NUEVA para que OpenCode la lea."
} else {
    $masked = $groqKey.Substring(0, [Math]::Min(8, $groqKey.Length)) + "..." + $groqKey.Substring([Math]::Max(0, $groqKey.Length - 4))
    Write-Host "[2/4] GROQ_API_KEY ya configurada: $masked" -ForegroundColor Green
}

# Advertir si hay ANTHROPIC_API_KEY con basura del intento anterior
$anthKey = [Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY', 'User')
if (-not [string]::IsNullOrWhiteSpace($anthKey) -and ($anthKey.Length -lt 100 -or $anthKey -match '[^\x21-\x7E]')) {
    Write-Host ""
    Write-Host "  AVISO: hay un ANTHROPIC_API_KEY invalido guardado. Para limpiarlo:" -ForegroundColor Yellow
    Write-Host '     setx ANTHROPIC_API_KEY ""' -ForegroundColor Cyan
    Write-Host "  (No es urgente si vas a usar solo Groq. OpenCode ignorara Anthropic sin key valida.)"
}

# --- Paso 3: Crear tarea programada para el monitor ---
# Usar cmdlets (Get/Register/Unregister-ScheduledTask) en lugar de schtasks.exe
# porque en Windows PowerShell 5.1, schtasks con 2>$null lanza NativeCommandError.
$taskName = "OpenCode Failover Monitor"
$monitor = Join-Path $here "monitor_ratelimit.ps1"

try {
    $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
    Write-Host "[3/4] Tarea '$taskName' ya existe. Actualizando..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
} catch {
    # No existia - normal en primera instalacion
}

try {
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    $trigger.Delay = "PT1M"  # 1 min de retraso
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$monitor`""
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
    Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -Settings $settings -Principal $principal -Description "Vigila rate-limits de Claude Code y lanza OpenCode" -ErrorAction Stop | Out-Null
    Write-Host "[3/4] Tarea '$taskName' creada. Arranca al login con 1 min de retraso." -ForegroundColor Green
} catch {
    Write-Host "[3/4] ERROR creando tarea programada: $_" -ForegroundColor Red
}

# --- Paso 4: Acceso directo en Escritorio para lanzar OpenCode con contexto ---
$desktop = [Environment]::GetFolderPath("Desktop")
$lnkPath = Join-Path $desktop "OpenCode (con contexto).lnk"
$launcher = Join-Path $here "LANZAR_OPENCODE_CONTEXTO.bat"

$shell = New-Object -ComObject WScript.Shell
$lnk = $shell.CreateShortcut($lnkPath)
$lnk.TargetPath = $launcher
$lnk.WorkingDirectory = $here
$lnk.WindowStyle = 7  # minimizada
$lnk.IconLocation = "$env:LOCALAPPDATA\Programs\@opencode-aidesktop\OpenCode.exe,0"
$lnk.Description = "Lanza OpenCode con contexto del working dir actual"
$lnk.Save()
Write-Host "[4/4] Acceso directo: $lnkPath" -ForegroundColor Green

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Instalacion completa." -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Proximos pasos:" -ForegroundColor Yellow
if ([string]::IsNullOrWhiteSpace($groqKey) -or $groqKey.Length -lt 30) {
    Write-Host "  1. Consigue tu GROQ_API_KEY (ver instrucciones arriba). Es GRATIS."
    Write-Host "  2. Ejecuta INICIAR_MONITOR.bat para arrancar el vigilante ahora."
    Write-Host "  3. Reinicia OpenCode para que lea el AGENTS.md y use Groq."
} else {
    Write-Host "  1. Ejecuta INICIAR_MONITOR.bat para arrancar el vigilante ahora."
    Write-Host "  2. Reinicia OpenCode para que lea el AGENTS.md y use Groq."
}
Write-Host ""
Write-Host "El monitor arrancara solo despues del proximo login." -ForegroundColor Gray
Write-Host "Log de eventos: $env:LOCALAPPDATA\opencode_failover.log" -ForegroundColor Gray
Write-Host ""
