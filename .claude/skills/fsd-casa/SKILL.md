---
name: fsd-casa
description: Flujo correcto para generar un FSD (Fill Status Document / correo de respuesta sobre una recepción) a una casa que NO es Mondelez. Usar SIEMPRE que haya que redactar un FSD, responder sobre una recepción de proveedor, o decidir a quién copiar en un correo de descarga. Fuerza buscar PRIMERO al coordinador externo (forwarder, transportista, casa) en Outlook antes de caer en la cadena interna Logística/Oscar Ayapan. Cubre casas importadas DUCA (OC 4700), coordinadores internos conocidos (Keyla, Guillermo) y el mapa casa→cadena externa.
---

# FSD Casas - Flujo correcto

Esta skill cubre el error más caro que repitió el agente FSD: contestar sobre la **cadena interna Arium** cuando la casa en realidad coordina con un **forwarder/proveedor externo**. El resultado fue copiar a Oscar Ayapan en correos que debían ir al coordinador del proveedor, y perder la pista del contacto real.

## Regla #1 — Buscar coordinador externo PRIMERO

ANTES de armar el FSD, buscar en Outlook por placa / DPI / OC quién coordinó esa recepción desde afuera de Arium.

Búsqueda típica:
```
asunto:<placa> OR cuerpo:<DPI_piloto> OR cuerpo:<#OC>
```

Si aparece una cadena externa (forwarder, transportista, coordinador del proveedor) → **esa es la cadena del FSD**. NO responder sobre Logística/Oscar.

Si NO aparece nada externo → recién ahí mirar coordinador interno.

## Mapa casa → cadena externa (patrón conocido)

| Casa receptora | Coordinador externo típico |
|---|---|
| DEMAGUSA | GRUMA |
| ICASA | MARAVILLA |
| (casas importadas DUCA) | forwarder del proveedor (ver abajo) |

**Patrón general:** las casas de alimentos y las importadas suelen tener un forwarder dedicado. Buscar primero.

## OC/ST 4700 = casa importada DUCA (NO traslado)

Si el número de OC/ST empieza con **4700**, es casa importada vía DUCA. El agente FSD viejo las descartaba incorrectamente como "traslado". NO son traslado.

Para estas:
- Ubicar la cita por placa / DPI en el Excel **"PROCESO DE IMPORTACION DUCA"**
- Identificar al coordinador interno responsable (ver abajo)
- Si hay forwarder/agente aduanal en Outlook, incluirlo

## Coordinadores internos conocidos (para importadas DUCA)

| Casa | Coordinador interno |
|---|---|
| KRAFT | Keyla |
| 3M | Keyla |
| EDGEWELL | Keyla |
| LIPTON | Keyla |
| BIOTEC | Keyla |
| BAT | Guillermo |
| ALEN | Guillermo |
| PHILIPS | Guillermo |
| DURACELL | Guillermo |

Si la casa no está en esta tabla, preguntar a Carlos antes de asumir.

## Qué NO hacer — anti-patrón

- ❌ Responder a Logística/Oscar Ayapan sin haber buscado al coordinador externo
- ❌ Tratar OC 4700 como traslado y descartarla
- ❌ Usar el dominio del remitente de la cadena interna como "cadena del cliente"
- ❌ Asumir que MOV / ICASA / DEMAGUSA son proveedores (NO lo son, son casas o terceros logísticos)
- ❌ Mandar correos antes de confirmar que la cadena externa existe

## Qué SÍ hacer

- ✅ Buscar en Outlook por placa, DPI del piloto, número de OC
- ✅ Si hay coordinador externo → esa es la cadena
- ✅ Si no hay → usar coordinador interno de la tabla
- ✅ Si no hay NINGUNO → preguntar a Carlos
- ✅ Para Mondelez: existe agente dedicado (`/agente-fsd-mondelez` en memoria), NO reusar este flujo para Mondelez

## Flujo paso a paso

1. Identificar casa, placa, DPI piloto, OC/ST, fecha de descarga
2. Si OC empieza con 4700 → marcar como DUCA, buscar en Excel importaciones
3. Buscar en Outlook (asunto + cuerpo) por placa / DPI / OC
4. ¿Hay cadena externa? → esa es la cadena del FSD, redactar respondiendo a ellos
5. ¿No hay externa? → coordinador interno de la tabla (Keyla/Guillermo según casa)
6. ¿Tampoco hay interno conocido? → PARAR y preguntar a Carlos
7. Redactar FSD claro (fecha, cantidad enviada vs recibida vs dañada, cita en Körber)
8. Antes de enviar: dejar en borrador de Outlook, Carlos lo revisa y envía

## Formato típico del FSD

Asunto: `FSD - <Casa> - <Fecha descarga> - OC <#>`

Cuerpo mínimo:
- Placa y piloto
- OC / ST / Factura
- Enviado (según guía de remisión)
- Recibido en Körber (con cita)
- Dañado / faltante (con detalle si aplica)
- Hora inicio / fin descarga

## Verificación antes de reportar como "hecho"

- [ ] El borrador está en Outlook (no enviado automáticamente)
- [ ] El TO lleva al coordinador externo real (o interno correcto si no hay externo)
- [ ] NO está copiado Logística/Oscar si existe coordinador externo
- [ ] Los números del FSD vienen de Körber (`CIERRE_INVENTARIO_STATUS`), no de estimados
- [ ] Si es Mondelez → se derivó al agente FSD Mondelez, NO este flujo
