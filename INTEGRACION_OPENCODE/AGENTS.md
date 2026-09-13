# Contexto de Carlos

- **Nombre:** Carlos David Trujillo Quezada (`trujillo.c@ariumcorp.com`)
- **Rol:** Supervisor de Recepción, Bodega 1110 CEDI, Distribuidora Me Llega / Arium
- **Jefes:** Oscar Ayapan, Mario Pérez, Andrés Castillo
- **KPIs:** Fill Rate, FSD, IRA, efectividad

## Reglas críticas
1. **Auditar antes de entregar.** No decir "listo" sin verificar el efecto real.
2. **Reusar lo que ya funciona.** Antes de reescribir integraciones (WhatsApp, Outlook, Körber), leer los scripts existentes.
3. **Nunca duplicar envíos WhatsApp.** Un solo disparo por acción, un solo sistema por grupo.
4. **Envío WA solo con `enviar_grupo_verificado.py`.** Verifica destino por DOM. Prohibido a ciegas.

## Entorno
- Windows 11, PowerShell primary. `py` disponible.
- SQL Körber: `172.16.5.43 / IN_E`, `Trusted_Connection=yes`.
- Outlook COM OK. Excel COM bloqueado por IT.
- Working dir: `C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON`.

## Proyectos activos
Dashboard Recepciones (SQL directo vía sp_DataRecepciones), Validador MDLZ, Cierre Turno, Medición Empleados (Q500), Monitor W2W, Cuadre SAP↔Körber, Agente Castro (WA descargas), Agente David (archivar PDFs MDLZ), FSD Mondelez.

## Comunicación
Terse, directo, español guatemalteco. Nunca inventar datos. Verificación > velocidad.
