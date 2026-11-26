# Seguimiento de firma — Procedimientos de SST (Orizon S.A.)

Scripts para generar en cada corte (lunes, miércoles y viernes 08:00):

1. `Control_Avance_Firma_SST_Respaldo_Gestion_<AAAAMMDD>.xlsx`: control automatizado y respaldo de gestión.
2. `Correos_Seguimiento_Firma_SST_Supervisores_<AAAAMMDD>_<HHMM>.docx`: un correo por supervisor con pendientes.

Ningún archivo de datos (exports de Buk, Excel o Word generados) se guarda en este repositorio, porque contienen datos personales.

## Uso

```bash
# 1) Excel: export actual + export del corte anterior (para la columna Δ)
python3 build_control.py <reporte_firma_actual.xlsx> <reporte_firma_anterior.xlsx> Control_Avance_Firma_SST_Respaldo_Gestion_AAAAMMDD.xlsx

# 2) Recalcular fórmulas (LibreOffice)
python3 <ruta-skill-xlsx>/scripts/recalc.py Control_Avance_Firma_SST_Respaldo_Gestion_AAAAMMDD.xlsx 350

# 3) Word: usa los .pkl que deja el paso 1 junto al Excel
python3 build_correos.py Control_Avance_Firma_SST_Respaldo_Gestion_AAAAMMDD.xlsx Correos_Seguimiento_Firma_SST_Supervisores_AAAAMMDD_HHMM.docx "DD-MM-AAAA HH:MM"
```

Fuente esperada: Buk, *Reporte de Firma Documentos OK 2026*, con 5 filas de metadatos y el encabezado en la fila 6. La fecha y hora de generación se leen de la fila 2.

## Criterios fijos

- Campaña: el nombre de archivo contiene `SST-OR-MR` (procedimientos MR-005/006/007/008/010). El resto de los documentos del export queda fuera.
- Solo trabajadores Activos. Incluye Coquimbo y El Golf como bloques separados de "Coronel y plantas".
- Pendiente = `Estado Firma Trabajador(a) == "No firmado"`. La empresa no firma (No Requiere Firma).
- Plazo: 5 días hábiles desde la emisión (24-11-2026), vence 01-12-2026. Feriados 08-12 y 25-12. 🟢 ≤ plazo−2 · 🟡 ≤ plazo · 🔴 > plazo.
- Semáforo por supervisor según % firmado: ✅ completo · 🟢 ≥80% · 🟡 50–79% · 🔴 <50%.
- Códigos de `Filtro documental` obsoletos (ETI, A1–A4, BB, E): se dejan en blanco, sin borrar la fila. Fuero se separa del resto, y las fichas F se separan de T/E/D.
- Control de duplicados: RUT + Nombre de Archivo (se detiene si encuentra alguno).
- Paleta azul (#1F4E79 / #2E75B6 / #DDEBF7), sin naranja. Firma de correos: "Equipo SST Orizon S.A.".

## Histórico de cortes

`build_control.py` mantiene `historico_cortes_sst.csv` junto al Excel de salida (una fila por corte, se actualiza por fecha y hora de generación, sin duplicar). Ese CSV alimenta el Registro de cortes de la hoja Control. Guárdalo junto a los Excel: si se pierde, el registro vuelve a partir desde los tres cortes del 24-11.

La hoja Export_Buk se lee hasta max(30.000, filas del export + 1.000). Los exports de Buk cambian de rango de fechas entre un día y otro; el filtro de campaña deja solo los procedimientos SST.
