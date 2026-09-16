param([string]$ScriptPath, [string[]]$ScriptArgs)

$pcCarlos = "FRU-LOG-C1G4NN3"
$online = Test-Connection $pcCarlos -Count 1 -Quiet -TimeoutSeconds 3

if ($online) {
    exit 0
}

if ($ScriptPath) {
    $ext = [System.IO.Path]::GetExtension($ScriptPath).ToLower()
    switch ($ext) {
        '.ps1'  { powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File $ScriptPath @ScriptArgs }
        '.py'   { pythonw.exe $ScriptPath @ScriptArgs }
        '.vbs'  { wscript.exe $ScriptPath @ScriptArgs }
        '.bat'  { cmd.exe /c $ScriptPath @ScriptArgs }
        default { & $ScriptPath @ScriptArgs }
    }
}
