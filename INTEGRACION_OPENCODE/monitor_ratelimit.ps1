# ============================================================
#  monitor_ratelimit.ps1
#  Vigila las sesiones activas de Claude Code.
#  Si detecta rate-limit / usage-limit, lanza OpenCode con la
#  carpeta del proyecto y pone el ultimo prompt del usuario en
#  el portapapeles para pegarlo alli.
#
#  Uso normal (background): pwsh -File monitor_ratelimit.ps1
#  Uso debug:               pwsh -File monitor_ratelimit.ps1 -Debug
# ============================================================

param(
    [switch]$Debug,
    [int]$IntervalSec = 15
)

$ErrorActionPreference = "Continue"

$claudeProjects = Join-Path $env:USERPROFILE ".claude\projects"
$openCode       = Join-Path $env:LOCALAPPDATA "Programs\@opencode-aidesktop\OpenCode.exe"
$stateFile      = Join-Path $env:LOCALAPPDATA "opencode_failover_state.json"
$logFile        = Join-Path $env:LOCALAPPDATA "opencode_failover.log"

# Patrones para detectar rate-limit REAL en cada linea del jsonl.
# Se parsea cada linea como JSON y se aplican reglas estrictas para
# NO gatillar por mera mencion del texto en conversacion del usuario.
#
# La deteccion ocurre solo si la ENTRY completa cumple una de:
#   a) type == "system" o type == "error" y contiene rate_limit/usage_limit en toolUseResult/content
#   b) message.content contiene un item con type "tool_result" is_error=true con texto de rate limit
#   c) Existe una propiedad "isApiErrorMessage": true con texto de limit
$limitTextPatterns = @(
    'Claude usage limit reached',
    'session limit',
    'rate limit',
    'rate_limit_error',
    '5-hour limit',
    'usage_limit',
    'exceeded your',
    'quota exceeded'
)

function Write-Log([string]$msg) {
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
    Add-Content -Path $logFile -Value $line -Encoding utf8
    if ($Debug) { Write-Host $line }
}

function Load-State {
    if (Test-Path $stateFile) {
        try {
            $json = Get-Content $stateFile -Raw -Encoding utf8
            $obj = $json | ConvertFrom-Json
            $ht = @{}
            $obj.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
            return $ht
        } catch {
            return @{}
        }
    }
    return @{}
}

function Save-State($ht) {
    try {
        $obj = New-Object PSObject
        foreach ($k in $ht.Keys) { $obj | Add-Member -NotePropertyName $k -NotePropertyValue $ht[$k] }
        $obj | ConvertTo-Json | Set-Content -Path $stateFile -Encoding utf8
    } catch {
        Write-Log "ERROR guardando estado: $_"
    }
}

function Decode-ProjectPath([string]$encoded) {
    # C--Users-trujillo-c-OneDrive---Me-Llega--S-A-Escritorio-ARCHIVOS-DE-PHYTON
    # -> C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON
    # Heuristica simple: no siempre invertible sin datos. Fallback al working dir del user.
    $candidates = @(
        "$env:USERPROFILE\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON",
        "$env:USERPROFILE\Desktop",
        (Get-Location).Path
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { return $c }
    }
    return $env:USERPROFILE
}

function Get-LastUserPrompt([string]$jsonlPath) {
    try {
        $lines = Get-Content $jsonlPath -Tail 200 -Encoding utf8 -ErrorAction Stop
        # Buscar de atras hacia adelante el ultimo mensaje "role":"user" con contenido de texto
        for ($i = $lines.Length - 1; $i -ge 0; $i--) {
            $line = $lines[$i]
            if ($line -match '"role"\s*:\s*"user"') {
                try {
                    $obj = $line | ConvertFrom-Json
                    if ($obj.message -and $obj.message.content) {
                        if ($obj.message.content -is [string]) { return $obj.message.content }
                        foreach ($c in $obj.message.content) {
                            if ($c.type -eq "text" -and $c.text) { return $c.text }
                        }
                    }
                } catch { }
            }
        }
    } catch { }
    return ""
}

function Show-Toast([string]$title, [string]$message) {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $notify = New-Object System.Windows.Forms.NotifyIcon
        $notify.Icon = [System.Drawing.SystemIcons]::Information
        $notify.BalloonTipTitle = $title
        $notify.BalloonTipText = $message
        $notify.Visible = $true
        $notify.ShowBalloonTip(8000)
        Start-Sleep -Seconds 1
        $notify.Dispose()
    } catch {
        Write-Log "No se pudo mostrar toast: $_"
    }
}

function Set-Clipboard-Safe([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return }
    try {
        Set-Clipboard -Value $text -ErrorAction Stop
        Write-Log "Ultimo prompt copiado al portapapeles ($($text.Length) chars)"
    } catch {
        Write-Log "No se pudo copiar al portapapeles: $_"
    }
}

# ---------- Sanity checks ----------
if (-not (Test-Path $openCode)) {
    Write-Log "FATAL: OpenCode no existe en $openCode"
    if ($Debug) { Write-Host "FATAL: OpenCode no existe en $openCode" -ForegroundColor Red }
    exit 1
}

if (-not (Test-Path $claudeProjects)) {
    Write-Log "FATAL: Carpeta de proyectos Claude no existe: $claudeProjects"
    exit 1
}

Write-Log "Monitor iniciado. IntervalSec=$IntervalSec  Debug=$Debug"
if ($Debug) { Write-Host "Monitor iniciado. Presiona Ctrl+C para detener." -ForegroundColor Green }

# ---------- Main loop ----------
while ($true) {
    try {
        $state = Load-State

        # Sesiones activas: modificadas en los ultimos 3 min
        $active = Get-ChildItem -Path $claudeProjects -Recurse -Filter "*.jsonl" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-3) -and $_.Length -gt 0 }

        foreach ($file in $active) {
            $sessionId = $file.BaseName
            $mtime = $file.LastWriteTime.ToString("o")

            # Si ya notificamos esta mtime, skip
            if ($state[$sessionId] -eq $mtime) { continue }

            # Leer las ultimas 40 lineas
            $tail = Get-Content $file.FullName -Tail 40 -Encoding utf8 -ErrorAction SilentlyContinue
            if (-not $tail) { continue }

            # Parsear linea por linea como JSON y aplicar reglas ESTRICTAS
            $matched = $false
            $matchedReason = ""

            foreach ($line in $tail) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                $obj = $null
                try { $obj = $line | ConvertFrom-Json -ErrorAction Stop } catch { continue }

                # Regla A: isApiErrorMessage=true con texto de limite
                if ($obj.PSObject.Properties.Name -contains 'isApiErrorMessage' -and $obj.isApiErrorMessage) {
                    $textBlob = $line
                    foreach ($p in $limitTextPatterns) {
                        if ($textBlob -match [regex]::Escape($p)) {
                            $matched = $true
                            $matchedReason = "isApiErrorMessage + '$p'"
                            break
                        }
                    }
                    if ($matched) { break }
                }

                # Regla B: message.content[i].type == 'tool_result' con is_error=true y texto de limite
                if ($obj.message -and $obj.message.content) {
                    foreach ($c in $obj.message.content) {
                        if ($c.type -eq 'tool_result' -and $c.is_error -eq $true) {
                            $txt = if ($c.content -is [string]) { $c.content } else { $c | ConvertTo-Json -Compress -Depth 4 }
                            foreach ($p in $limitTextPatterns) {
                                if ($txt -match [regex]::Escape($p)) {
                                    $matched = $true
                                    $matchedReason = "tool_result.is_error + '$p'"
                                    break
                                }
                            }
                            if ($matched) { break }
                        }
                    }
                    if ($matched) { break }
                }

                # Regla C: response.error.type == 'rate_limit_error'
                if ($obj.error -and $obj.error.type -match 'rate_limit|usage_limit|overloaded') {
                    $matched = $true
                    $matchedReason = "error.type=$($obj.error.type)"
                    break
                }

                # Regla D: type == 'system' con subtype de error y mensaje de limite
                if ($obj.type -eq 'system' -and $obj.subtype -match 'error|api_error' -and $obj.message) {
                    $txt = if ($obj.message -is [string]) { $obj.message } else { $obj.message | ConvertTo-Json -Compress -Depth 4 }
                    foreach ($p in $limitTextPatterns) {
                        if ($txt -match [regex]::Escape($p)) {
                            $matched = $true
                            $matchedReason = "system.$($obj.subtype) + '$p'"
                            break
                        }
                    }
                    if ($matched) { break }
                }
            }

            if ($matched) {
                Write-Log "RATE LIMIT detectado en sesion $sessionId (razon: $matchedReason)"

                # Marcar antes de lanzar, para no duplicar
                $state[$sessionId] = $mtime
                Save-State $state

                # Guess project dir
                $projectDirName = Split-Path $file.DirectoryName -Leaf
                $workDir = Decode-ProjectPath $projectDirName
                Write-Log "Lanzando OpenCode con carpeta: $workDir"

                # Extraer ultimo prompt del user
                $lastPrompt = Get-LastUserPrompt $file.FullName
                if ($lastPrompt) {
                    Set-Clipboard-Safe $lastPrompt
                }

                # Toast
                Show-Toast "Claude Code sin tokens" "Abriendo OpenCode. Tu ultimo prompt esta en el portapapeles."

                # Lanzar OpenCode
                try {
                    Start-Process -FilePath $openCode -ArgumentList "`"$workDir`""
                    Write-Log "OpenCode lanzado OK"
                } catch {
                    Write-Log "ERROR al lanzar OpenCode: $_"
                }

                # Solo un lanzamiento por vuelta del loop
                break
            } else {
                # Registrar mtime para no releer si no cambia
                $state[$sessionId] = $mtime
            }
        }

        Save-State $state
    }
    catch {
        Write-Log "ERROR en loop: $_"
    }

    Start-Sleep -Seconds $IntervalSec
}
