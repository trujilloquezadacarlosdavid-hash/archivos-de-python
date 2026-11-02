---
name: outlook-com
description: Reglas y gotchas para trabajar con Outlook vía COM/MAPI desde PowerShell/Python en los agentes de Carlos (David, FSD Mondelez, FSD Casa, Carta Aceptación, Notas Envío, MTTO Montacargas, Faltantes). Usar SIEMPRE que haya que leer, buscar, archivar, responder o crear borradores en Outlook. Cubre Restrict/DASL, Items.Find trampas, iterar inboxes seguros, borrador vs envío automático, adjuntos inline (CID), thread/reply, carpetas por mes/día, y manejo de cuentas múltiples (trujillo.c vs recepciones@).
---

# Outlook COM - Patrones y gotchas

6 agentes de Carlos usan Outlook (David archiva Mondelez, FSD Mondelez/Casa, Carta Aceptación, Notas Envío, MTTO Montacargas, Faltantes). Todos cayeron en los mismos bugs. Esta skill los centraliza.

## Setup estándar

```powershell
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace('MAPI')
$inbox = $ns.GetDefaultFolder(6)  # 6 = Inbox
# Otros FolderType: 5=SentItems, 9=Calendar, 16=Drafts, 3=DeletedItems
```

En Python:
```python
import win32com.client as win32
ol = win32.Dispatch('Outlook.Application')
ns = ol.GetNamespace('MAPI')
inbox = ns.GetDefaultFolder(6)
```

## Cuentas múltiples — Carlos vs recepciones@

Carlos tiene 2 buzones:
- **trujillo.c@ariumcorp.com** (principal)
- **recepciones@ariumcorp.com** (buzón de bodega, arrancó 16-jul-2026)

Para acceder a `recepciones@`:
```powershell
$store = $ns.Folders | Where-Object { $_.Name -match 'recepciones' } | Select-Object -First 1
$inboxR = $store.Folders | Where-Object { $_.Name -eq 'Bandeja de entrada' -or $_.Name -eq 'Inbox' }
```

Nunca asumir que `GetDefaultFolder(6)` apunta a `recepciones@` — apunta al PRINCIPAL (trujillo.c).

## Trampa 1 — Restrict con DASL (la forma correcta de buscar)

❌ **NO usar `Items.Find`** — es frágil con caracteres especiales (`'`, `"`, acentos, `/`).

❌ **NO iterar todo el inbox con `foreach`** — en inboxes de 20k+ correos muere por timeout.

✅ **Usar `Restrict` con DASL**:

```powershell
# Por asunto que contenga texto
$filter = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%Reporte de descarga MONDELEZ%'"
$items = $inbox.Items.Restrict($filter)

# Por fecha de recibido >= hoy
$ayer = (Get-Date).AddDays(-1).ToString('MM/dd/yyyy HH:mm')
$filter = "[ReceivedTime] >= '$ayer'"
$items = $inbox.Items.Restrict($filter)

# Combinar: asunto + fecha
$filter = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%MONDELEZ%' AND ""urn:schemas:httpmail:datereceived"" >= '$ayer'"
```

Trampa dentro de la trampa: las comillas DASL son dobles DOBLES (`""`) adentro del string PowerShell.

## Trampa 2 — Sort DESCENDING antes de iterar

Para procesar los más recientes primero (y poder cortar temprano):
```powershell
$items.Sort('[ReceivedTime]', $true)  # true = descendente
foreach ($item in $items) {
    if ($item.ReceivedTime -lt $limite) { break }
    # procesar
}
```

Sin `Sort` el orden es impredecible → romper con `break` falla.

## Trampa 3 — Borrador vs envío automático

**Default correcto para Carlos: BORRADOR, NUNCA envío automático** (excepto agentes aprobados como David que archivan, no envían).

```powershell
$mail = $ol.CreateItem(0)  # 0 = MailItem
$mail.To = "destino@example.com"
$mail.Subject = "Asunto"
$mail.HTMLBody = $html
# ✅ Dejar en borrador para que Carlos revise y envíe
$mail.Save()
$mail.Display()  # opcional: abrir en pantalla
# ❌ NO llamar $mail.Send() salvo que Carlos lo autorice explícito
```

Casos donde SÍ se puede `.Send()` automático:
- Agente David archivando (no manda nada, solo mueve)
- Agente MTTO Montacargas reply al thread (ya aprobado, va al mismo thread)
- Agente FSD Mondelez (borrador que Carlos revisa primero — NO auto-send)

Si el usuario pide "manda este correo", preguntar: ¿borrador o enviar directo? Default siempre borrador.

## Trampa 4 — Adjuntos inline (CID) vs attachment

Para fotos/videos inline en el cuerpo del correo (ej. MTTO Montacargas):
```powershell
$attach = $mail.Attachments.Add($rutaFoto)
$attach.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
    "img001"
)
$mail.HTMLBody = $mail.HTMLBody + '<img src="cid:img001" />'
```

Sin SetProperty del CID, la imagen aparece como adjunto adjunto pero NO inline.

## Trampa 5 — Reply al thread vs correo nuevo

Para responder al mismo thread (que no se corte la conversación):
```powershell
$reply = $itemOriginal.ReplyAll()  # o .Reply() para solo al remitente
$reply.HTMLBody = $miTexto + "<br><br>" + $reply.HTMLBody  # mantener el histórico abajo
$reply.Display()
```

❌ NO crear `$ol.CreateItem(0)` y setear manualmente To/Subject con "RE: ..." — rompe el threading en el cliente del destinatario.

## Trampa 6 — Archivar a carpeta por mes/día

Patrón de agente David/Notas Envío:
```powershell
$asunto = $item.Subject
# Extraer fecha del asunto (ej "26P06" o "02112026")
if ($asunto -match '(\d{2})(\d{2})(\d{4})') {
    $dd = $matches[1]; $mm = $matches[2]; $yyyy = $matches[3]
    $mesStr = "{0}P{1:00}" -f $yyyy.Substring(2,2), [int]$mm  # 26P11
    $diaStr = "{0}{1}{2}" -f $dd, $mm, $yyyy                   # 02112026
    $destino = Join-Path $baseFolder "$mesStr\$diaStr"
    if (-not (Test-Path $destino)) { New-Item -ItemType Directory -Path $destino -Force | Out-Null }
    # ... guardar
}
```

## Trampa 7 — Guardar adjunto sin corromper nombre

```powershell
foreach ($att in $item.Attachments) {
    $nombreLimpio = $att.FileName -replace '[\\/:*?"<>|]', '_'
    $ruta = Join-Path $destino $nombreLimpio
    # Si existe, sumarle timestamp
    if (Test-Path $ruta) {
        $base = [IO.Path]::GetFileNameWithoutExtension($nombreLimpio)
        $ext  = [IO.Path]::GetExtension($nombreLimpio)
        $ruta = Join-Path $destino "${base}_$(Get-Date -Format 'HHmmss')$ext"
    }
    $att.SaveAsFile($ruta)
}
```

## Trampa 8 — Liberar COM al final

```powershell
# Al final del script
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
```

Sin esto queda `OUTLOOK.EXE` zombie en memoria.

## Checklist antes de dar por terminado el agente

- [ ] Usa `Restrict` DASL, NO `Items.Find`
- [ ] Sort DESC si itera por fecha
- [ ] Es BORRADOR (`.Save()` + `.Display()`), NO `.Send()` automático salvo excepción aprobada
- [ ] Si hay inline, usa SetProperty del CID
- [ ] Si es reply, usa `.ReplyAll()`/`.Reply()`, NO `CreateItem`
- [ ] Libera COM al final
- [ ] Chequeé si corresponde cuenta `trujillo.c` o `recepciones@`
- [ ] Si es tarea programada → cumple `/tarea-silenciosa` (hidden, pythonw/vbs)
