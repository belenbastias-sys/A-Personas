---
name: nomina-contratos-obra
description: >
  Cruza el Reporte de Dotacion (BUK) con el Reporte de Firma de Documentos
  (BUK) de Orizon S.A. / Nutrisco para producir la Nomina de Gestion de
  Contratos: identifica el tipo real de cada contrato (Obra / Plazo fijo /
  Indefinido) comparando el campo del maestro contra el nombre del archivo
  firmado, detecta a quien le falta el contrato, marca discrepancias entre
  BUK y el documento firmado, y saca aparte cualquier Tipo de Contrato
  vacio o invalido como dato sucio a corregir. Usa esta skill SIEMPRE que
  Belen pida cruzar dotacion con firma de documentos, identificar contratos
  por obra/faena/extension de temporada, revisar anexos de renovacion de
  plazo fijo o PF, encontrar a quien le falta firmar su contrato, auditar
  el campo Tipo de Contrato de la dotacion, o arme una nomina de gestion de
  contratos para un corte de fecha especifico (por ejemplo "contratos que
  iniciaron despues del X"). Tambien activa si menciona "contratos de obra",
  "contrato faena", "extension temporada" + firma, o pide revisar el
  informe mensual de documentos para identificar contratos y anexos de un
  mes en particular.
---

# Nomina de Gestion de Contratos (Dotacion x Firma Documentos)

Esta skill automatiza el cruce que se arma manualmente cada vez que Belen
necesita saber, dentro de la dotacion activa, quien tiene contrato firmado,
de que tipo es realmente ese contrato, y donde el maestro BUK y el
documento firmado no coinciden. Antes se rehacia a mano cada corte de mes;
ahora es un script parametrizado por fecha y lugar de trabajo.

## Por que existe el cruce (no basta con mirar el maestro)

El campo "Tipo de Contrato" del Reporte de Dotacion (Obra / Plazo fijo /
Indefinido) es la fuente de verdad para agrupar a la gente — nunca lo
reemplaces por lo que diga el nombre del archivo. Pero ese campo puede
estar desactualizado si alguien renovo un plazo fijo o firmo un anexo de
extension de temporada y el maestro no se actualizo. El cruce con Firma
Documentos existe para **detectar esa brecha**, no para pisar el dato del
maestro. Por eso toda discrepancia queda marcada en amarillo para que
RRHH/Relaciones Laborales decida, nunca se corrige sola.

## Antes de correr el script

Sigue primero `reglas-datos-atencion-personas` si esta disponible — valida
duplicados, lista de exclusion y separaciones obligatorias antes de tocar
cualquier nomina de personas. Esta skill hereda esas reglas, no las
reemplaza.

Confirma con Belen (no asumas) si el pedido trae alguno de estos matices,
porque cambian el resultado:

1. **Fecha de corte.** Si pide "contratos que iniciaron despues de una
   fecha", usa `--fecha-corte`. Si pide la dotacion completa sin filtro de
   fecha, omite el parametro.
2. **Lugar de trabajo a excluir.** Por defecto se excluyen Coquimbo y
   Santiago (alcance estandar del area Bio Bio, igual que el resto de las
   skills de Atencion Personas). Si el pedido es para otra zona, ajusta
   `--excluir-lugar`.
3. **Que hacer con los hallazgos.** El script nunca corrige el maestro BUK
   ni asume clasificaciones ambiguas — deja "Revisar manual" cuando el
   nombre del archivo no calza con ningun patron conocido (por ejemplo
   "Fena" en vez de "Faena"). Pregunta a Belen antes de reinterpretar esos
   casos.

## Como correrlo

```bash
python3 skills/nomina-contratos-obra/scripts/generar_reporte.py \
  --dotacion "<ruta Reporte Dotacion Coordinador.xlsx>" \
  --firma "<ruta Reporte de Firma Documentos.xlsx>" \
  --output "<ruta de salida .xlsx>" \
  [--fecha-corte DD-MM-YYYY] \
  [--excluir-lugar Coquimbo Santiago]
```

El script imprime un resumen por consola (totales, discrepancias,
anomalias, mes con mas firmas) — usalo para redactar el mensaje al usuario
antes de abrir el Excel. Si algo no calza con las columnas esperadas del
maestro BUK, el script falla con un mensaje explicito en vez de adivinar
una columna parecida; revisa si BUK cambio el nombre de algun campo.

## Que trae el Excel de salida

1. **Resumen** — indicadores con formulas (`COUNTA`/`COUNTIF`, no
   hardcodeados) + Recomendacion de Gestion redactada segun los hallazgos
   de esa corrida.
2. **Nomina Gestion** — todo el universo filtrado, con RUT, nombre, codigo
   de ficha, filtro documental, observacion (fuero), gerencia, area/centro
   de costo, RUT y nombre del supervisor, codigo GeoVictoria, tipo de
   contrato del maestro vs. tipo detectado por archivo, y estado de firma.
3. **Contratos y Anexos - Mes Top** — los documentos de contrato/anexo
   firmados en el mes con mayor concentracion de firmas dentro del
   universo filtrado. Ese mes se calcula automaticamente en cada corrida
   (nunca queda fijo en "julio" ni en ningun otro mes).
4. **Alertas - Tipo Discrepante** — donde el maestro y el archivo firmado
   no coinciden. Resaltado completo en amarillo: es la hoja que Belen
   revisa primero.
5. **Anomalias Tipo Contrato** — activos cuyo campo "Tipo de Contrato" en
   BUK esta vacio o no es Obra/Plazo fijo/Indefinido. Son datos sucios que
   hay que corregir en BUK, no forman parte de la nomina de gestion.
6. **Excluidos por Lugar** — quienes cumplian el resto del filtro pero su
   Lugar de Trabajo esta en la lista de exclusion, para trazabilidad.

## Reglas de clasificacion por nombre de archivo

Ver el docstring de `scripts/generar_reporte.py` para el detalle exacto de
tokens (esta documentado ahi para no duplicarlo). En resumen:

- "extension" + "temporada" -> Obra (Extension Temporada)
- "faena" -> Obra (Faena)
- "obra" -> Obra
- "pf", o "plazo"+"fijo" -> Plazo Fijo
- "renovacion" + ("plazo" o "pf") -> Plazo Fijo (Renovacion)
- "fena" (sin "faena") -> Revisar manual, posible error de tipeo
- nada de lo anterior -> Sin clasificar - revisar

El documento vigente de cada persona es el de **fecha de firma del
trabajador mas reciente** entre los que califican como "documento de
contrato" (nombre contiene "contrato", o "anexo" + renovacion/plazo/pf;
se excluyen certificados y anexos de politica/conflicto de intereses).

## Despues de correr el script

1. Resume los hallazgos en el chat (totales de cada hoja) antes de mandar
   el archivo — no mandes el Excel en silencio.
2. Cierra siempre con Recomendacion de Gestion priorizando: primero las
   discrepancias de tipo (exposicion legal), despues las anomalias de dato
   en BUK, despues los "sin contrato".
3. Todo pendiente detectado (discrepancias, anomalias, sin contrato) va a
   la Bitacora de Compromisos en Notion via la skill `notion-bitacora` —
   nunca en un informe aparte. Pregunta antes de cargarlos si no te lo
   pidieron explicitamente.
