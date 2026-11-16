# Transcripción automática de clases grabadas

Herramienta para convertir la clase grabada (mp4/mp3) en texto (.txt)
y pasárselo a Claude, sin que tengas que escuchar vos la clase completa.

## Cuándo usarla

Cuando el Lic. del curso NO activó la transcripción nativa de Zoom y vos
querés que Claude trabaje con lo que realmente dijo el catedrático,
no con un resumen genérico del tema.

## Qué necesitás una sola vez (instalación)

**Pedile a Claude:** `instalame Whisper`

Claude confirma, instala `faster-whisper` y `ffmpeg` (~500 MB a 2 GB
según componentes), y queda listo para siempre.

## Flujo de cada clase (después de instalado)

1. Entrá a Zoom con el link de la grabación.
2. Clickeá **Descargar** (botón de descarga en la UI del video).
3. Guardá el `.mp4` en esta carpeta.
4. **Doble-click a `TRANSCRIBIR_CLASE.bat`**.
5. El script lista los archivos. Elegí el número del .mp4 nuevo.
6. Esperá 10-30 min (más largo si el video es de >1h o no tenés GPU).
7. Se genera un `.txt` con el mismo nombre al lado del `.mp4`.
8. Pasale ese `.txt` a Claude en el chat:
   *"Acá tenés la transcripción de la clase de X, armá los resúmenes individuales que pidió el licenciado"*.

## Qué contiene el .txt

```
# Transcripcion automatica (faster-whisper small)
# Archivo: clase_tema_2.mp4
# Duracion: 65 min 23 seg

[00:00] Buenos días, espero que estén bien. Hoy vamos a ver los 8 tipos de escucha...
[00:12] Antes de empezar quiero que tomen nota de la siguiente definición...
[00:25] ...
```

Timestamps + texto palabra por palabra.

## Alternativa: pedirle al profe que active transcripción

Zoom tiene transcripción nativa que se activa con 1 click del host. Si
el licenciado la activa, la grabación viene con panel "Audio Transcript"
y podés descargar un `.vtt` directo sin necesitar esta herramienta.

Pedíselo al profe: *"¿Puede activar la transcripción automática
cuando grabe las clases? Nos ayuda a estudiar."*

## Limitaciones

- No reemplaza escuchar la clase en vivo (perdés el tono, lo visual).
- Si varias personas hablan al mismo tiempo, mezcla voces.
- Términos técnicos muy específicos puede escribirlos mal (ej.
  nombres propios o siglas).
- Requiere el `.mp4` descargado — si Zoom no deja descargar, no sirve.
