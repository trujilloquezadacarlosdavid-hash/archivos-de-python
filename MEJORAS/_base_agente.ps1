# _base_agente.ps1
# Funciones compartidas para agentes de Carlos.
# Uso: en cada agente poner al inicio:
#   . "C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON\MEJORAS\_base_agente.ps1"
# Despues llamar: Assert-Hora, Assert-PCPrincipal, Write-AgenteLog, etc.

$script:PCPrincipal = "FRU-LOG-C1G4NN3"
$script:PCRespaldo  = "DORCAS-PC"  # ajustar al nombre real de la PC Dorcas cuando se confirme
$script:LogRoot     = "C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON\MEJORAS\LOGS_CENTRAL"

# ------------------------------------------------------------
# Assert-Hora
# Sale con exit 0 (sin ruido) si estamos FUERA de la ventana horaria.
# Default: 3AM-13 (1PM) para envios WA / rotulos.
# ------------------------------------------------------------
function Assert-Hora {
    param(
        [int]$Min = 3,
        [int]$Max = 13
    )
    $h = (Get-Date).Hour
    if ($h -lt $Min -or $h -ge $Max) {
        Write-AgenteLog -Agente "ASSERT_HORA" -Nivel "INFO" -Mensaje "Fuera de ventana $Min-$Max (hora actual $h). Saliendo."
        exit 0
    }
}

# ------------------------------------------------------------
# Assert-PCPrincipal
# Chequea failover: si estoy en Dorcas y la principal responde, salgo.
# ------------------------------------------------------------
function Assert-PCPrincipal {
    $esta = $env:COMPUTERNAME
    if ($esta -eq $script:PCPrincipal) { return }
    if ($esta -eq $script:PCRespaldo) {
        $ping = Test-Connection -ComputerName $script:PCPrincipal -Count 2 -Quiet -ErrorAction SilentlyContinue
        if ($ping) {
            Write-AgenteLog -Agente "FAILOVER" -Nivel "INFO" -Mensaje "PC principal viva, Dorcas se salta."
            exit 0
        }
        Write-AgenteLog -Agente "FAILOVER" -Nivel "WARN" -Mensaje "PC principal NO responde, Dorcas toma el relevo."
        return
    }
    # PC desconocida -> no corro
    Write-AgenteLog -Agente "FAILOVER" -Nivel "WARN" -Mensaje "PC $esta no autorizada. Saliendo."
    exit 0
}

# ------------------------------------------------------------
# Write-AgenteLog
# Escribe a log por-agente + log central unificado.
# ------------------------------------------------------------
function Write-AgenteLog {
    param(
        [Parameter(Mandatory=$true)][string]$Agente,
        [Parameter(Mandatory=$true)][string]$Mensaje,
        [ValidateSet('INFO','WARN','ERROR','OK')][string]$Nivel = 'INFO'
    )
    if (-not (Test-Path $script:LogRoot)) {
        New-Item -ItemType Directory -Path $script:LogRoot -Force | Out-Null
    }
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $linea = "$ts | $Nivel | $Agente | $Mensaje"

    # Log por agente
    $logPorAgente = Join-Path $script:LogRoot "$Agente.log"
    $linea | Out-File -Append -FilePath $logPorAgente -Encoding UTF8

    # Log central
    $logCentral = Join-Path $script:LogRoot "_central.log"
    $linea | Out-File -Append -FilePath $logCentral -Encoding UTF8

    # Tambien a consola
    Write-Host $linea
}

# ------------------------------------------------------------
# Get-OneDriveRoot
# Resuelve la ruta OneDrive segun el usuario que corre (trujillo.c o recepciones).
# ------------------------------------------------------------
function Get-OneDriveRoot {
    if ($env:OneDrive) { return $env:OneDrive }
    $up = [Environment]::GetFolderPath('UserProfile')
    return (Join-Path $up "OneDrive - Me Llega, S.A")
}

# ------------------------------------------------------------
# Release-Excel / Release-Outlook
# Libera correctamente los objetos COM para no dejar procesos zombie.
# ------------------------------------------------------------
function Release-Excel {
    param($Excel, $Workbook = $null, $Worksheet = $null)
    try {
        if ($Workbook) { $Workbook.Close($false); [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) }
        if ($Worksheet) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet) }
        if ($Excel) {
            $Excel.Quit()
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
        }
    } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

function Release-Outlook {
    param($Outlook, $Namespace = $null)
    try {
        if ($Namespace) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) }
        if ($Outlook) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) }
    } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# ------------------------------------------------------------
# Get-KorberConnection
# Devuelve una conexion abierta a Korber (IN_E) con los defaults estandar.
# ------------------------------------------------------------
function Get-KorberConnection {
    $cs = "Server=172.16.5.43;Database=IN_E;Integrated Security=True;TrustServerCertificate=True"
    $cn = New-Object System.Data.SqlClient.SqlConnection($cs)
    $cn.Open()
    return $cn
}

# ------------------------------------------------------------
# Invoke-KorberQuery
# Ejecuta una query y devuelve DataTable.  Maneja reintentos.
# ------------------------------------------------------------
function Invoke-KorberQuery {
    param(
        [Parameter(Mandatory=$true)][string]$Sql,
        [int]$Timeout = 120,
        [int]$Retries = 3
    )
    for ($i = 1; $i -le $Retries; $i++) {
        try {
            $cn = Get-KorberConnection
            $cmd = $cn.CreateCommand()
            $cmd.CommandText = $Sql
            $cmd.CommandTimeout = $Timeout
            $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
            $dt = New-Object System.Data.DataTable
            [void]$da.Fill($dt)
            $cn.Close()
            return $dt
        } catch {
            Write-AgenteLog -Agente "KORBER" -Nivel "WARN" -Mensaje "Intento $i/$Retries: $($_.Exception.Message)"
            if ($i -eq $Retries) {
                Write-AgenteLog -Agente "KORBER" -Nivel "ERROR" -Mensaje "SQL no responde despues de $Retries intentos."
                throw
            }
            Start-Sleep -Seconds 10
        }
    }
}

# ------------------------------------------------------------
# Marcar-Resultado
# Al final del agente, deja un marcador que la salud puede leer.
# ------------------------------------------------------------
function Marcar-Resultado {
    param(
        [Parameter(Mandatory=$true)][string]$Agente,
        [ValidateSet('OK','WARN','ERROR')][string]$Estado = 'OK',
        [string]$Detalle = ''
    )
    if (-not (Test-Path $script:LogRoot)) {
        New-Item -ItemType Directory -Path $script:LogRoot -Force | Out-Null
    }
    $marcador = @{
        agente   = $Agente
        estado   = $Estado
        detalle  = $Detalle
        ts       = Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'
        pc       = $env:COMPUTERNAME
        usuario  = $env:USERNAME
    } | ConvertTo-Json -Compress
    $marcadorFile = Join-Path $script:LogRoot "_$Agente.json"
    $marcador | Out-File -FilePath $marcadorFile -Encoding UTF8
    Write-AgenteLog -Agente $Agente -Nivel $Estado -Mensaje "Marcador: $Detalle"
}
