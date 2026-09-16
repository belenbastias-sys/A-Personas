#!/usr/bin/env python3
"""
Genera la Nomina de Gestion de Contratos (Orizon S.A. / Nutrisco) cruzando
el Reporte de Dotacion (BUK) con el Reporte de Firma de Documentos (BUK).

Uso:
    python generar_reporte.py \
        --dotacion "Reporte Dotacion Coordinador.xlsx" \
        --firma "Nuevo reporte de Firma Documentos.xlsx" \
        --output "Nomina_Gestion_Contratos.xlsx" \
        [--fecha-corte 04-07-2026] \
        [--excluir-lugar Coquimbo Santiago]

Reglas de negocio fijas (confirmadas con Belen Bastias, Atencion Personas):
  - Universo: dotacion con Estado = Activo.
  - Si se pasa --fecha-corte, se filtra Fecha Ingreso Compania > esa fecha.
    Sin el parametro, se reporta toda la dotacion activa.
  - Se excluyen por defecto los Lugar de Trabajo = Coquimbo y Santiago
    (alcance estandar del area Bio Bio). Ajustable con --excluir-lugar.
  - Cruce con Firma Documentos por RUT (Numero de Documento).
  - "Documento de contrato" = nombre de archivo contiene 'contrato', o
    contiene 'anexo' + ('renovacion'/'plazo'/'pf'). Se excluyen
    certificados y anexos de politica/conflicto de intereses.
  - El documento vigente de cada persona es el de firma mas reciente
    (Fecha Firma Trabajador(a)).
  - Clasificacion del documento por nombre de archivo (tokens, sin acentos):
      'pf' o 'plazo'+'fijo'            -> Plazo Fijo
      'extension'+'temporada'          -> Obra (Extension Temporada)
      'faena'                          -> Obra (Faena)
      'obra'                           -> Obra
      'renovacion'+('plazo' o 'pf')    -> Plazo Fijo (Renovacion)
      'fena' (sin 'faena')             -> Revisar manual (posible typo)
      nada de lo anterior              -> Sin clasificar - revisar
  - El Tipo de Contrato del maestro BUK manda para agrupar la nomina;
    el cruce con el archivo firmado es para verificar, no para redefinir.
    Toda discrepancia queda marcada, nunca corregida en silencio.
  - Cualquier registro cuyo campo 'Tipo de Contrato' este vacio o no sea
    Obra/Plazo fijo/Indefinido se reporta aparte como anomalia de dato.
  - La hoja de "Contratos y Anexos recientes" usa el mes con mas firmas
    de documentos de contrato dentro del universo filtrado (calculado
    automaticamente en cada corrida, no hardcodeado).

Este script no decide nada de negocio por si solo: si detecta datos que no
calzan con estas reglas (columna faltante, formato de fecha raro, etc.),
falla con un mensaje claro en vez de adivinar.
"""
import argparse
import re
import sys
import unicodedata
from collections import Counter
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ORANGE = "F4811F"
WHITE = "FFFFFF"
YELLOW = "FFF3B0"
GREY = "F2F2F2"
FONT_NAME = "Arial"

TIPOS_VALIDOS = {"Obra", "Plazo fijo", "Indefinido"}
DEFAULT_EXCLUIR_LUGAR = ["Coquimbo", "Santiago"]


# ---------------------------------------------------------------- utilidades

def norm(s):
    if s is None:
        return ""
    s = str(s).lower()
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))


def tokens(s):
    return re.split(r"[_\-\s\.]+", norm(s))


def parse_fecha_corte(s):
    if not s:
        return None
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise SystemExit(f"No pude interpretar --fecha-corte='{s}'. Usa DD-MM-YYYY.")


def load_libro(path, min_expected_cols):
    """Los libros BUK traen 4 filas de metadata + 1 fila vacia antes del
    encabezado real. Buscamos la primera fila con >= min_expected_cols
    valores no vacios en vez de asumir que siempre es la fila 6, por si
    BUK cambia el preambulo."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    header = None
    header_row_idx = None
    rows_iter = ws.iter_rows(values_only=True)
    all_rows = list(rows_iter)
    for i, r in enumerate(all_rows):
        non_empty = sum(1 for v in r if v not in (None, ""))
        if non_empty >= min_expected_cols:
            header = r
            header_row_idx = i
            break
    if header is None:
        raise SystemExit(f"No encontre una fila de encabezado valida en {path}")
    idx = {name: j for j, name in enumerate(header) if name}
    data_rows = all_rows[header_row_idx + 1:]
    return idx, data_rows


def require_cols(idx, cols, filelabel):
    faltantes = [c for c in cols if c not in idx]
    if faltantes:
        raise SystemExit(
            f"Faltan columnas esperadas en {filelabel}: {faltantes}. "
            "Revisa si BUK cambio el nombre de alguna columna antes de reintentar."
        )


# ------------------------------------------------------------- clasificacion

def clasificar_archivo(nombre):
    tk = set(tokens(nombre))
    if "pf" in tk:
        return "Plazo Fijo"
    if "plazo" in tk and "fijo" in tk:
        return "Plazo Fijo"
    if "extension" in tk and "temporada" in tk:
        return "Obra (Extension Temporada)"
    if "faena" in tk:
        return "Obra (Faena)"
    if "obra" in tk:
        return "Obra"
    if "renovacion" in tk and ("plazo" in tk or "pf" in tk):
        return "Plazo Fijo (Renovacion)"
    if "fena" in tk:
        return "Revisar manual (posible typo de Faena)"
    return None


def categoria(tipo_detectado):
    if tipo_detectado is None:
        return None
    if "Plazo Fijo" in tipo_detectado:
        return "Plazo fijo"
    if "Obra" in tipo_detectado:
        return "Obra"
    return None


def es_documento_contrato(nombre):
    tk = set(tokens(nombre))
    if "certificado" in tk:
        return False
    if "anexo" in tk and ("politica" in tk or "conflicto" in tk or "conflictos" in tk):
        return False
    if "contrato" in tk:
        return True
    if "anexo" in tk and ("renovacion" in tk or "plazo" in tk or "pf" in tk):
        return True
    return False


def es_anexo(nombre):
    return "anexo" in set(tokens(nombre))


# -------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dotacion", required=True, help="Reporte Dotacion Coordinador (xlsx)")
    ap.add_argument("--firma", required=True, help="Reporte de Firma Documentos (xlsx)")
    ap.add_argument("--output", required=True, help="Ruta del Excel de salida")
    ap.add_argument("--fecha-corte", default=None, help="DD-MM-YYYY. Filtra Fecha Ingreso Compania > esta fecha. Si se omite, se reporta toda la dotacion activa.")
    ap.add_argument("--excluir-lugar", nargs="*", default=DEFAULT_EXCLUIR_LUGAR, help="Lugares de Trabajo a excluir (default: Coquimbo Santiago)")
    args = ap.parse_args()

    fecha_corte = parse_fecha_corte(args.fecha_corte)

    # ---- Dotacion ----
    dot_idx, dot_rows = load_libro(args.dotacion, min_expected_cols=30)
    require_cols(dot_idx, [
        "Estado", "Tipo de Contrato", "Fecha Ingreso Compañía", "Lugar de Trabajo",
        "Número de Documento", "Nombre Completo", "Código de Ficha", "Filtro documental",
        "Observación (Fuero u otros)", "Nombre Gerencia", "Nombre Área", "Centro de Costo",
        "Número de Documento Supervisor", "Nombre Supervisor", "CODIGO GEOVICTORIA",
    ], args.dotacion)

    def g(row, col):
        j = dot_idx.get(col)
        return None if j is None else row[j]

    activos = [r for r in dot_rows if g(r, "Estado") == "Activo"]

    anomalias_tipo = [r for r in activos if g(r, "Tipo de Contrato") not in TIPOS_VALIDOS]
    validos = [r for r in activos if g(r, "Tipo de Contrato") in TIPOS_VALIDOS]

    filtrados = []
    excluidos_lugar = []
    for row in validos:
        fi = g(row, "Fecha Ingreso Compañía")
        if fecha_corte is not None:
            if not isinstance(fi, datetime) or fi <= fecha_corte:
                continue
        if g(row, "Lugar de Trabajo") in args.excluir_lugar:
            excluidos_lugar.append(row)
            continue
        filtrados.append(row)

    # ---- Firma Documentos ----
    firma_idx, firma_rows = load_libro(args.firma, min_expected_cols=10)
    require_cols(firma_idx, [
        "Número de Documento", "Nombre de Archivo", "Estado Firma Trabajador(a)",
        "Fecha Firma Trabajador(a)",
    ], args.firma)

    firma_by_rut = {}
    for r in firma_rows:
        rut = r[firma_idx["Número de Documento"]]
        if rut is None:
            continue
        firma_by_rut.setdefault(rut, []).append(r)

    def fidx(row, col):
        if row is None:
            return None
        j = firma_idx.get(col)
        return None if j is None else row[j]

    # ---- Cruce ----
    nomina = []
    todos_docs_contrato_fecha = []  # para detectar el mes dominante
    docs_por_persona = {}

    for row in filtrados:
        rut = g(row, "Número de Documento")
        docs = firma_by_rut.get(rut, [])
        docs_contrato = [d for d in docs if es_documento_contrato(fidx(d, "Nombre de Archivo"))]
        docs_por_persona[rut] = docs_contrato

        doc_elegido = None
        if docs_contrato:
            def fkey(d):
                f = fidx(d, "Fecha Firma Trabajador(a)")
                return f if isinstance(f, datetime) else datetime.min
            doc_elegido = sorted(docs_contrato, key=fkey, reverse=True)[0]

        for d in docs_contrato:
            ff = fidx(d, "Fecha Firma Trabajador(a)")
            if isinstance(ff, datetime):
                todos_docs_contrato_fecha.append((ff.year, ff.month))

        nombre_archivo = fidx(doc_elegido, "Nombre de Archivo") if doc_elegido else None
        tipo_detectado = clasificar_archivo(nombre_archivo) if doc_elegido else None
        tipo_maestro = g(row, "Tipo de Contrato")
        cat_detectada = categoria(tipo_detectado)
        discrepancia = cat_detectada is not None and cat_detectada != tipo_maestro

        nomina.append({
            "RUT": rut,
            "Nombre Completo": g(row, "Nombre Completo"),
            "Codigo Ficha": g(row, "Código de Ficha"),
            "Filtro documental": g(row, "Filtro documental"),
            "Observacion": g(row, "Observación (Fuero u otros)"),
            "Gerencia": g(row, "Nombre Gerencia"),
            "Area": g(row, "Nombre Área"),
            "Centro de Costo": g(row, "Centro de Costo"),
            "RUT Supervisor": g(row, "Número de Documento Supervisor"),
            "Nombre Supervisor": g(row, "Nombre Supervisor"),
            "Codigo GeoVictoria": g(row, "CODIGO GEOVICTORIA"),
            "Fecha Ingreso Compania": g(row, "Fecha Ingreso Compañía"),
            "Tipo Contrato Maestro": tipo_maestro,
            "Nombre Archivo Contrato": nombre_archivo,
            "Tipo Detectado por Archivo": tipo_detectado if tipo_detectado else ("Sin clasificar - revisar" if doc_elegido else "SIN CONTRATO"),
            "Estado Firma Trabajador": fidx(doc_elegido, "Estado Firma Trabajador(a)") if doc_elegido else None,
            "Fecha Firma Trabajador": fidx(doc_elegido, "Fecha Firma Trabajador(a)") if doc_elegido else None,
            "Discrepancia Maestro vs Archivo": "SI - revisar" if discrepancia else "No",
            "Contrato Faltante": "SI - FALTA CONTRATO" if doc_elegido is None else "No",
        })

    # ---- Mes dominante para hoja de "recientes" ----
    if todos_docs_contrato_fecha:
        (mes_top_year, mes_top_month), _n = Counter(todos_docs_contrato_fecha).most_common(1)[0]
    else:
        mes_top_year, mes_top_month = None, None

    docs_mes_top = []
    if mes_top_year is not None:
        for row in filtrados:
            rut = g(row, "Número de Documento")
            for d in docs_por_persona.get(rut, []):
                ff = fidx(d, "Fecha Firma Trabajador(a)")
                if isinstance(ff, datetime) and ff.year == mes_top_year and ff.month == mes_top_month:
                    nm = fidx(d, "Nombre de Archivo")
                    docs_mes_top.append({
                        "RUT": rut,
                        "Nombre Completo": g(row, "Nombre Completo"),
                        "Codigo Ficha": g(row, "Código de Ficha"),
                        "Tipo Documento": "Anexo de Contrato" if es_anexo(nm) else "Contrato",
                        "Clasificacion Detectada": clasificar_archivo(nm) or "Sin clasificar",
                        "Nombre Archivo": nm,
                        "Estado Firma Trabajador": fidx(d, "Estado Firma Trabajador(a)"),
                        "Fecha Firma Trabajador": ff,
                        "Gerencia": g(row, "Nombre Gerencia"),
                        "Nombre Supervisor": g(row, "Nombre Supervisor"),
                    })

    anomalias_out = [{
        "RUT": g(r, "Número de Documento"),
        "Nombre Completo": g(r, "Nombre Completo"),
        "Codigo Ficha": g(r, "Código de Ficha"),
        "Tipo de Contrato (valor en BUK)": g(r, "Tipo de Contrato") or "(vacio)",
        "Gerencia": g(r, "Nombre Gerencia"),
        "RUT Supervisor": g(r, "Número de Documento Supervisor"),
        "Nombre Supervisor": g(r, "Nombre Supervisor"),
    } for r in anomalias_tipo]

    excluidos_out = [{
        "RUT": g(r, "Número de Documento"),
        "Nombre Completo": g(r, "Nombre Completo"),
        "Lugar de Trabajo": g(r, "Lugar de Trabajo"),
        "Tipo de Contrato": g(r, "Tipo de Contrato"),
    } for r in excluidos_lugar]

    write_excel(args.output, nomina, docs_mes_top, anomalias_out, excluidos_out,
                mes_top_year, mes_top_month, fecha_corte, args.excluir_lugar)

    disc = sum(1 for o in nomina if o["Discrepancia Maestro vs Archivo"] != "No")
    falta = sum(1 for o in nomina if o["Contrato Faltante"] != "No")
    print(f"Nomina generada: {len(nomina)} personas")
    print(f"  Discrepancias Tipo Contrato Maestro vs Archivo: {disc}")
    print(f"  Sin contrato firmado: {falta}")
    print(f"  Anomalias de Tipo de Contrato (fuera de Obra/Plazo fijo/Indefinido): {len(anomalias_out)}")
    print(f"  Excluidos por lugar de trabajo {args.excluir_lugar}: {len(excluidos_out)}")
    if mes_top_year:
        print(f"  Mes con mas firmas de contrato/anexo: {mes_top_month:02d}-{mes_top_year} ({len(docs_mes_top)} documentos)")
    print(f"Guardado en: {args.output}")


# -------------------------------------------------------------------- excel

def write_excel(out_path, nomina, docs_mes_top, anomalias, excluidos, mes_top_year, mes_top_month, fecha_corte, excluir_lugar):
    wb = openpyxl.Workbook()

    header_font = Font(name=FONT_NAME, bold=True, color=WHITE, size=10)
    header_fill = PatternFill("solid", fgColor=ORANGE)
    cell_font = Font(name=FONT_NAME, size=10)
    title_font = Font(name=FONT_NAME, bold=True, size=13, color=ORANGE)
    sub_font = Font(name=FONT_NAME, italic=True, size=9, color="666666")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    yellow_fill = PatternFill("solid", fgColor=YELLOW)

    def write_table(ws, title, subtitle, headers, rows, col_widths, highlight_col=None, highlight_val=None, highlight_all=False):
        ws["A1"] = title
        ws["A1"].font = title_font
        ws["A2"] = subtitle
        ws["A2"].font = sub_font
        start_row = 4
        for j, h in enumerate(headers, start=1):
            c = ws.cell(row=start_row, column=j, value=h)
            c.font = header_font
            c.fill = header_fill
            c.border = border
            c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
        for i, row in enumerate(rows, start=start_row + 1):
            for j, h in enumerate(headers, start=1):
                val = row.get(h)
                c = ws.cell(row=i, column=j, value=val)
                if isinstance(val, datetime):
                    c.number_format = "DD-MM-YYYY"
                c.font = cell_font
                c.border = border
                if highlight_all:
                    c.fill = yellow_fill
                elif i % 2 == 0:
                    c.fill = PatternFill("solid", fgColor=GREY)
                if highlight_col and row.get(highlight_col) == highlight_val:
                    c.fill = yellow_fill
        for j, w in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w
        if rows:
            ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(headers))}{start_row + len(rows)}"

    ws1 = wb.active
    ws1.title = "Nomina Gestion"
    headers1 = ["RUT", "Nombre Completo", "Codigo Ficha", "Filtro documental", "Observacion",
                "Gerencia", "Area", "Centro de Costo", "RUT Supervisor", "Nombre Supervisor",
                "Codigo GeoVictoria", "Fecha Ingreso Compania", "Tipo Contrato Maestro",
                "Nombre Archivo Contrato", "Tipo Detectado por Archivo", "Estado Firma Trabajador",
                "Fecha Firma Trabajador", "Discrepancia Maestro vs Archivo", "Contrato Faltante"]
    widths1 = [14, 28, 10, 16, 16, 20, 26, 22, 14, 26, 14, 16, 16, 50, 26, 16, 16, 20, 18]
    corte_txt = f"ingreso > {fecha_corte.strftime('%d-%m-%Y')}" if fecha_corte else "sin filtro de fecha"
    write_table(ws1, "Nomina de Gestion - Contratos (Obra / Plazo fijo / Indefinido)",
                f"Activos, {corte_txt}, excluye {excluir_lugar}. Total: {len(nomina)} personas. "
                f"Generado {datetime.now().strftime('%d-%m-%Y %H:%M')}.",
                headers1, nomina, widths1,
                highlight_col="Discrepancia Maestro vs Archivo", highlight_val="SI - revisar")

    ws2 = wb.create_sheet("Contratos y Anexos - Mes Top")
    headers2 = ["RUT", "Nombre Completo", "Codigo Ficha", "Tipo Documento", "Clasificacion Detectada",
                "Nombre Archivo", "Estado Firma Trabajador", "Fecha Firma Trabajador", "Gerencia", "Nombre Supervisor"]
    widths2 = [14, 28, 10, 18, 24, 55, 18, 16, 22, 26]
    mes_label = f"{mes_top_month:02d}-{mes_top_year}" if mes_top_year else "N/A"
    write_table(ws2, f"Contratos y Anexos firmados en el mes con mayor concentracion ({mes_label})",
                f"Mes calculado automaticamente segun Fecha Firma Trabajador(a) dentro de la nomina filtrada. Total: {len(docs_mes_top)}.",
                headers2, docs_mes_top, widths2)

    ws3 = wb.create_sheet("Alertas - Tipo Discrepante")
    disc = [o for o in nomina if o["Discrepancia Maestro vs Archivo"] != "No"]
    headers3 = ["RUT", "Nombre Completo", "Codigo Ficha", "Gerencia", "RUT Supervisor", "Nombre Supervisor",
                "Tipo Contrato Maestro", "Nombre Archivo Contrato", "Tipo Detectado por Archivo", "Fecha Firma Trabajador"]
    widths3 = [14, 28, 10, 20, 14, 26, 16, 55, 20, 16]
    write_table(ws3, "ALERTA - El maestro BUK y el documento firmado no coinciden en tipo de contrato",
                f"No se corrigio el maestro. Requiere validacion de RRHH/Relaciones Laborales. Total: {len(disc)} casos.",
                headers3, disc, widths3, highlight_all=True)

    ws4 = wb.create_sheet("Anomalias Tipo Contrato")
    headers4 = ["RUT", "Nombre Completo", "Codigo Ficha", "Tipo de Contrato (valor en BUK)", "Gerencia", "RUT Supervisor", "Nombre Supervisor"]
    widths4 = [14, 28, 10, 26, 20, 14, 26]
    write_table(ws4, "Anomalias - Tipo de Contrato vacio o distinto de Obra/Plazo fijo/Indefinido",
                f"Dato sucio a corregir en BUK, no incluido en la nomina de gestion. Total: {len(anomalias)}.",
                headers4, anomalias, widths4, highlight_all=True)

    ws5 = wb.create_sheet("Excluidos por Lugar")
    headers5 = ["RUT", "Nombre Completo", "Lugar de Trabajo", "Tipo de Contrato"]
    widths5 = [14, 28, 18, 16]
    write_table(ws5, "Excluidos de la nomina - Fuera de alcance",
                f"Cumplian el resto del filtro pero su Lugar de Trabajo esta en {excluir_lugar}. Total: {len(excluidos)}.",
                headers5, excluidos, widths5)

    # ---- Resumen con formulas ----
    ws0 = wb.create_sheet("Resumen", 0)
    ws0["A1"] = "Resumen de Gestion - Nomina de Contratos"
    ws0["A1"].font = title_font
    ws0["A2"] = f"Fuente: Reporte Dotacion Coordinador + Reporte Firma Documentos. Corte generado {datetime.now().strftime('%d-%m-%Y')}."
    ws0["A2"].font = sub_font

    last1 = 4 + len(nomina)
    last2 = 4 + len(docs_mes_top)
    last4 = 4 + len(anomalias)
    last5 = 4 + len(excluidos)

    r = 4
    ws0.cell(row=r, column=1, value="Indicador").font = header_font
    ws0.cell(row=r, column=1).fill = header_fill
    ws0.cell(row=r, column=2, value="Valor").font = header_font
    ws0.cell(row=r, column=2).fill = header_fill
    r += 1
    formulas = [
        ("Total nomina de gestion", f"=COUNTA('Nomina Gestion'!A5:A{last1})" if nomina else 0),
        ("Con contrato firmado identificado", f"=COUNTIF('Nomina Gestion'!S5:S{last1},\"No\")" if nomina else 0),
        ("SIN CONTRATO (falta documento)", f"=COUNTIF('Nomina Gestion'!S5:S{last1},\"SI - FALTA CONTRATO\")" if nomina else 0),
        ("Discrepancia Tipo Maestro vs Archivo", f"=COUNTIF('Nomina Gestion'!R5:R{last1},\"SI - revisar\")" if nomina else 0),
        (f"Contratos/Anexos firmados en mes top ({mes_top_month:02d}-{mes_top_year})" if mes_top_year else "Contratos/Anexos firmados en mes top",
         f"=COUNTA('Contratos y Anexos - Mes Top'!A5:A{last2})" if docs_mes_top else 0),
        ("Anomalias de Tipo de Contrato (dato sucio BUK)", f"=COUNTA('Anomalias Tipo Contrato'!A5:A{last4})" if anomalias else 0),
        ("Excluidos por lugar de trabajo fuera de alcance", f"=COUNTA('Excluidos por Lugar'!A5:A{last5})" if excluidos else 0),
    ]
    for label, formula in formulas:
        ws0.cell(row=r, column=1, value=label).font = cell_font
        ws0.cell(row=r, column=1).border = border
        c = ws0.cell(row=r, column=2, value=formula)
        c.font = Font(name=FONT_NAME, size=10, bold=True)
        c.border = border
        r += 1

    ws0.column_dimensions["A"].width = 55
    ws0.column_dimensions["B"].width = 20

    r += 2
    ws0.cell(row=r, column=1, value="RECOMENDACION DE GESTION").font = Font(name=FONT_NAME, bold=True, size=11, color=ORANGE)
    r += 1
    recs = [
        "1. Revisar y confirmar con Relaciones Laborales/Jefatura cada caso de la hoja 'Alertas - Tipo Discrepante' antes del cierre de mes: el documento firmado no coincide con el tipo registrado en BUK.",
        "2. Corregir en BUK los registros de 'Anomalias Tipo Contrato' (Tipo de Contrato vacio o invalido) - son datos que no deberian existir en el maestro.",
        "3. Priorizar la cobranza de firma de los documentos en estado 'No firmado' dentro de la hoja de mes top.",
        "4. Cargar toda discrepancia y anomalia nueva en la Bitacora de Compromisos de Notion, con responsable y fecha de seguimiento, antes de cerrar el periodo.",
    ]
    for rec in recs:
        ws0.cell(row=r, column=1, value=rec).font = cell_font
        ws0.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
        ws0.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws0.row_dimensions[r].height = 30
        r += 1

    wb.save(out_path)


if __name__ == "__main__":
    main()
