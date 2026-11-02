---
name: alcance
description: Entrega una TARJETA DE ALCANCE antes de calcular CUALQUIER análisis sobre datos (SQL Körber, Excel, SAP, reportes). Fuerza validar fuente, filtros, exclusiones y supuestos ANTES de ejecutar o reportar números. Usar siempre que el usuario pida un análisis, un conteo, un total, una comparación, un KPI, un cuadre o cualquier cifra derivada de datos. También antes de generar un reporte o dashboard.
---

# Tarjeta de Alcance

Esta skill fuerza el patrón que más molesta a Carlos cuando no se cumple: analizar sin validar la población.

## Regla

ANTES de ejecutar la consulta o calcular el número, entregá la Tarjeta de Alcance y **esperá OK explícito**. No calcules, no muestres totales, no armes el reporte hasta tener el visto bueno.

## Formato de la tarjeta

```
🎯 TARJETA DE ALCANCE
---------------------
Fuente         : <tabla/vista/archivo exacto>
Período        : <rango de fechas o "sin filtro temporal">
Filtros        : <WHERE / criterios aplicados>
Exclusiones    : <qué se deja fuera y POR QUÉ>
Supuestos      : <lo que estoy asumiendo y NO verifiqué>
---------------------
¿OK para calcular?
```

## Checklist de exclusiones típicas (siempre considerar)

Según el tipo de análisis, preguntate si corresponde EXCLUIR y declararlo:

- **Inventario Körber**: `status='H'` (hold/dañado), scrap, vencido, P20-P23 (bodega auxiliar), filas TOTALES/RESUMEN, SKUs sin costo (Q0), ubicaciones sin matrícula
- **Recepciones**: 4700 = casa importada (NO traslado), citas canceladas, OC/ST no finalizadas
- **Despachos**: `TranType 341` (NO 322), filtrar por `WhsCode` correcto (1110A001 vs 1110)
- **Casos Heinz**: EXCLUIR Philips/Listerine/Durex (son casas separadas)
- **Mondelez NC**: usar `#Factura MDLZ` como llave de cruce, NO el STO interno
- **ABC/DEF**: cortar al último producto real, excluir filas TOTALES/RESUMEN que contaminan sumas
- **Casos reporte diario**: fuente correcta = `CIERRE_INVENTARIO_STATUS`

## Supuestos que SIEMPRE marcar

- Mapeo UOM (CJ vs UND) si el reporte pide cajas
- Si "Heinz" en el contexto incluye o excluye las otras casas Kraft
- Si "hoy" es fecha del server o fecha local
- Si el status es el actual o el histórico
- Si F_Expira=1900 cuenta como vencido o como "sin fecha" (ojo centinela)

## Qué NO hacer

- ❌ Correr la query primero y entregar números con una "nota" al pie
- ❌ Asumir que el usuario ya conoce los filtros
- ❌ Omitir exclusiones por brevedad ("no te marco Hold porque es obvio")
- ❌ Decir "ya corrí la consulta, estos son los totales" sin OK previo

## Qué SÍ hacer

- ✅ Mostrar la tarjeta ANTES de la primera query productiva
- ✅ Preguntar sobre cada supuesto que no esté explícito
- ✅ Si el usuario pide "rápido, solo un número", igual pegar la tarjeta en 4 líneas y esperar OK
- ✅ Cuando el usuario diga "OK", recién ahí ejecutar

## Ejemplo real

Usuario: "cuánto de Heinz tenemos recibido este año"

Respuesta correcta:
```
🎯 TARJETA DE ALCANCE
Fuente     : HJ_TRANSACTIONS + HJ_MasterData (unidades convertidas a cajas vía UOM='CJ')
Período    : 2026-01-01 a hoy
Filtros    : TranType=341, WhsCode LIKE '1110%', Casa='HEINZ'
Exclusiones: Philips, Listerine, Durex (otras casas Kraft), status='H', filas RESUMEN
Supuestos  : "recibido" = finalizado en Körber (NO citas en descarga). Cajas = Qty/ConversionFactor.
¿OK para calcular?
```

Respuesta INCORRECTA: `"Son 19,383 cajas. Nota: excluí dañado."` (sin tarjeta, sin OK)
