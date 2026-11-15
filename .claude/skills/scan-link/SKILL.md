---
name: scan-link
description: Escanea cualquier link y le explica a Carlos de qué se trata SIN abrir nada en su PC. Carlos trabaja en Arium donde IT bloquea sitios con firewall corporativo (TikTok, Google AI Studio, muchos sitios sociales y de IA). Usar SIEMPRE que Carlos pegue un link y pregunte "de qué se trata", "qué es esto", "léeme este link", "scanea este link" o similar. También cuando pegue un link sin texto, asumir que quiere saber qué es.
---

# Scan Link — leer sin abrir

## Objetivo
Carlos pega un link. Yo lo leo DESDE MI LADO (WebFetch del servidor de Anthropic), le resumo el contenido en español, y le aviso de banderas de seguridad. **NUNCA lo invito a abrir el link en su PC.**

## Por qué existe esta skill
- IT de Arium bloquea muchos sitios con firewall corporativo (TikTok, aistudio.google.com, redes sociales, sitios de IA, etc.)
- Carlos no puede abrir sitios desconocidos en su PC por política de seguridad
- Carlos ya tuvo pantalla de bloqueo "Arium cybersecurity" por TikTok (incidente previo)
- Necesita saber si un link vale la pena antes de pedirle al móvil personal que lo abra

## Flujo estándar

### Paso 1 — Fetch del link
Usar `WebFetch` directamente, o el agente `web-fetch` si WebFetch no está disponible. Pasar la URL completa y una pregunta tipo "describe el contenido de esta página, autor, tema principal, de qué se trata".

NO usar:
- `Bash` con `curl` o `wget` (eso corre en la PC de Carlos — IT puede bloquearlo igual)
- `navigate` del browser pane (eso también sale desde la PC de Carlos a veces)
- `preview_start` con URL (abre pane en la app, Carlos lo ve, pero igual el fetch puede pasar por el firewall de Arium dependiendo de config — preferir WebFetch)

### Paso 2 — Estructura de la respuesta
Responder a Carlos con este formato breve en español:

```
📍 **De qué se trata**: [una línea]
👤 **Autor/fuente**: [si hay]
📝 **Resumen**: [2-4 bullets de lo esencial]
🚦 **Banderas**: [solo si hay — phishing, pide credenciales, descarga raro, política IA, etc.]
```

### Paso 3 — Si el sitio bloquea
Algunos sitios responden 403/429 a bots (TikTok, Instagram reels, LinkedIn posts, X/Twitter a veces). Si el fetch falla:

1. Decile a Carlos: "Ese sitio bloquea bots (como la vez de TikTok). Vas a tener que abrirlo en el celular personal, NO en la PC de trabajo."
2. Pedirle screenshot si quiere que le resuma lo que ve.
3. NO intentes reintentar con headers raros ni proxies — es contra las Terms del sitio.

### Paso 4 — Si el contenido es un PDF, repo de código, o algo largo
- PDF: WebFetch lo descarga; mencionalo en el resumen y resaltá si es largo/corto.
- Repo GitHub: resumí qué hace, lenguaje, estrellas si hay, última actividad.
- Video de YouTube: WebFetch solo lee la página HTML — sacá título, canal, descripción, duración. No podés "ver" el video.

## Reglas duras
- **NUNCA invites a Carlos a abrir el link en su PC.** Ni siquiera en Edge/Chrome corporativo.
- **NUNCA uses Bash para descargar el link** — eso sale desde la PC de Carlos.
- **NUNCA ejecutes scripts o código descargados del link.** Si es código, mostrás qué hace y chau.
- **NUNCA guardes el contenido del link como archivo en la PC de Carlos** sin pedirle OK.
- **Si el link pide credenciales** (login page, OAuth): aviso de seguridad, no completo nada automático.

## Links sensibles / confidenciales
Si el link es de contenido corporativo de Arium (SharePoint, Teams, intranet):
- WebFetch probablemente NO puede entrar (requiere auth)
- Decile: "Este link es de SharePoint/Teams/intranet de Arium — yo desde afuera no entro. Si querés que te resuma el documento, descargalo en tu PC y abrime el archivo acá."

## Casos de uso típicos
- Link de TikTok/YouTube/noticia que le mandaron por WhatsApp → leer contenido sin abrirlo
- Link de artículo técnico (Stack Overflow, Medium, blog) → resumir la solución
- Link de repo GitHub de una herramienta nueva → qué hace y si vale la pena
- Link de documentación oficial (Microsoft, Python, PowerShell) → extraer la parte relevante
- Link de alguien que le mandó "mirá esto" sin contexto → qué es y si parece legítimo
