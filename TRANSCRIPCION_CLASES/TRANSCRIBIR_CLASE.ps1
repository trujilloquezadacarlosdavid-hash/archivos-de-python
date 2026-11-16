# TRANSCRIBIR_CLASE.ps1
# Herramienta para transcribir clases grabadas (Zoom, Teams, mp4, mp3, etc.)
# usando faster-whisper (modelo open source de OpenAI).
#
# USO:
#   1. Descarga el .mp4 de la clase desde Zoom a esta carpeta
#   2. Doble-click al .bat
#   3. Elegi el archivo a transcribir
#   4. Espera 10-20 min (depende duracion del video + CPU)
#   5. El .txt sale al lado del .mp4 con el mismo nombre

$ErrorActionPreference = "Stop"

$carpeta = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "============================================"
Write-Host " TRANSCRIPCION DE CLASES (faster-whisper)"
Write-Host "============================================"
Write-Host ""

# Chequear si faster-whisper esta instalado
$check = & py -m pip show faster-whisper 2>&1
if ($LASTEXITCODE -ne 0 -or $check -match "WARNING: Package\(s\) not found") {
    Write-Host "ERROR: faster-whisper NO esta instalado." -ForegroundColor Red
    Write-Host ""
    Write-Host "Para instalarlo (una sola vez), pedile a Claude:"
    Write-Host '  "instalame Whisper"'
    Write-Host ""
    Write-Host "Claude te confirma y corre la instalacion por vos."
    Write-Host ""
    pause
    exit 1
}

# Chequear ffmpeg (necesario para extraer audio de .mp4)
$ffmpeg = Get-Command ffmpeg -ErrorAction SilentlyContinue
if (-not $ffmpeg) {
    Write-Host "AVISO: ffmpeg no encontrado en PATH." -ForegroundColor Yellow
    Write-Host "faster-whisper lo necesita para .mp4. Si falla, pedile a Claude que lo instale."
}

# Listar archivos de audio/video en la carpeta
$archivos = Get-ChildItem -Path $carpeta -File | Where-Object { $_.Extension -match '\.(mp4|mp3|wav|m4a|webm|mkv|avi|flac|ogg)$' } | Sort-Object LastWriteTime -Descending

if ($archivos.Count -eq 0) {
    Write-Host "No hay archivos de audio/video en $carpeta" -ForegroundColor Yellow
    Write-Host "Copia aqui el .mp4 de la clase primero."
    pause
    exit 1
}

Write-Host "Archivos disponibles:"
for ($i = 0; $i -lt $archivos.Count; $i++) {
    $sizeMB = [math]::Round($archivos[$i].Length / 1MB, 1)
    Write-Host "  [$($i+1)] $($archivos[$i].Name) ($sizeMB MB)"
}
Write-Host ""

$sel = Read-Host "Elegi el numero del archivo a transcribir"
$idx = [int]$sel - 1
if ($idx -lt 0 -or $idx -ge $archivos.Count) {
    Write-Host "Seleccion invalida." -ForegroundColor Red
    pause
    exit 1
}

$input = $archivos[$idx].FullName
$output = [System.IO.Path]::ChangeExtension($input, ".txt")

Write-Host ""
Write-Host "Transcribiendo: $($archivos[$idx].Name)" -ForegroundColor Green
Write-Host "Esto puede tardar entre 10 y 30 minutos segun duracion y tu CPU."
Write-Host "(Si tu PC tiene GPU NVIDIA, se usa automaticamente y es mucho mas rapido.)"
Write-Host ""

# Script Python inline
$pyScript = @"
import sys
from faster_whisper import WhisperModel

# Usar modelo 'medium' (buena calidad, razonable en CPU) o 'small' si CPU muy lenta
# large-v3 es el mejor pero necesita GPU practicamente
modelo = 'small'

print(f'Cargando modelo whisper {modelo}...', flush=True)
model = WhisperModel(modelo, device='auto', compute_type='int8')

print(f'Transcribiendo audio...', flush=True)
segments, info = model.transcribe(r'$input', language='es', beam_size=5, vad_filter=True)

print(f'Idioma detectado: {info.language} (confianza {info.language_probability:.2f})', flush=True)
print(f'Duracion del audio: {info.duration:.0f} segundos', flush=True)
print(f'Escribiendo texto a: $output', flush=True)

with open(r'$output', 'w', encoding='utf-8') as f:
    f.write(f'# Transcripcion automatica (faster-whisper {modelo})\n')
    f.write(f'# Archivo: $($archivos[$idx].Name)\n')
    f.write(f'# Duracion: {int(info.duration // 60)} min {int(info.duration % 60)} seg\n')
    f.write('\n')
    for seg in segments:
        mins = int(seg.start // 60)
        secs = int(seg.start % 60)
        f.write(f'[{mins:02d}:{secs:02d}] {seg.text.strip()}\n')
        print(f'  [{mins:02d}:{secs:02d}] {seg.text.strip()[:80]}', flush=True)

print('LISTO.', flush=True)
"@

$tmpPy = "$env:TEMP\whisper_run_$(Get-Random).py"
$pyScript | Out-File -Encoding UTF8 -FilePath $tmpPy

try {
    & py $tmpPy
    if (Test-Path $output) {
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Green
        Write-Host " LISTO - Transcripcion guardada en:" -ForegroundColor Green
        Write-Host "   $output" -ForegroundColor Green
        Write-Host "============================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Tamano: $((Get-Item $output).Length) bytes"
        Write-Host ""
        Write-Host "Pasale este .txt a Claude y le decis que arme el resumen o las tareas."
    } else {
        Write-Host "ERROR: no se genero el archivo de salida." -ForegroundColor Red
    }
} finally {
    Remove-Item $tmpPy -ErrorAction SilentlyContinue
}

Write-Host ""
pause
