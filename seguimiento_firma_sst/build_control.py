"""Respaldo de gestión + Control de Avance automatizado (pegar export Buk 'Reporte de Firma Documentos OK')."""
import sys, re, datetime as dt
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, Reference

SRC, PREV, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
EMISION = dt.date(2026, 11, 24)
OBSOLETOS = {"ETI", "A1", "A2", "A3", "A4", "BB", "E"}
CAMP = "SST-OR-MR"

raw = list(openpyxl.load_workbook(SRC, read_only=True).worksheets[0].iter_rows(values_only=True))
gen_txt = raw[1][0]
m = re.search(r"(\d{2})/(\d{2})/(\d{4}) a las (\d{2}):(\d{2})(AM|PM)", gen_txt)
hh = int(m[4]) % 12 + (12 if m[6] == "PM" else 0)
GEN = dt.datetime(int(m[3]), int(m[2]), int(m[1]), hh, int(m[5]))

x = pd.read_excel(SRC, header=5)
n_rows = len(x)
camp = x[x["Nombre de Archivo"].str.contains(CAMP, na=False)]
n_fuera = n_rows - len(camp)
n_inact = (camp.Estado != "Activo").sum()
d = camp[camp.Estado == "Activo"].copy()
assert d.duplicated(["Número de Documento", "Nombre de Archivo"]).sum() == 0
bl = lambda l: {"Coquimbo": "Coquimbo", "El Golf": "El Golf (Santiago)"}.get(l, "Coronel y plantas")
d["Bloque"] = d["Lugar de Trabajo"].map(bl)
d["Procedimiento SST"] = d["Nombre de Archivo"].str.extract(r"(SST-OR-MR-\d+)")[0]
d["Prefijo ficha"] = d["Código de Ficha"].astype(str).str[0]
fd = d["Filtro documental"].astype("string").str.strip()
d["Fuero"] = fd.str.upper().str.contains("FUERO").fillna(False).map({True: "Con fuero", False: "Sin fuero"})
d["Filtro Documental (vigente)"] = fd.where(~fd.isin(OBSOLETOS), "").fillna("")
d["Sindicato"] = d["Sindicato"].fillna("(Sin sindicato)")
d["Fecha Creación"] = pd.to_datetime(d["Fecha Creación"])
d["Pendiente"] = (d["Estado Firma Trabajador(a)"] == "No firmado").astype(int)
pend = d[d.Pendiente == 1].copy()

pv = pd.read_excel(PREV, header=5)
pv = pv[pv["Nombre de Archivo"].str.contains(CAMP, na=False) & (pv.Estado == "Activo")]
NORM = lambda t: " ".join(str(t).split()).casefold()
prev_sup = pv[pv["Estado Firma Trabajador(a)"] == "No firmado"].groupby(pv["Nombre Supervisor"].map(NORM)).size()

# ---------------- estilos
NAR, AZ = "2E75B6", "1F4E79"
fH = PatternFill("solid", fgColor=AZ); fN = PatternFill("solid", fgColor=NAR)
fIn = PatternFill("solid", fgColor="DDEBF7"); fY = PatternFill("solid", fgColor="FFF2CC")
fG = PatternFill("solid", fgColor="E2EFDA"); fA = PatternFill("solid", fgColor="FFEB9C"); fR = PatternFill("solid", fgColor="F8CBAD")
fT = PatternFill("solid", fgColor="BDD7EE")
thin = Side(style="thin", color="BFBFBF"); B = Border(left=thin, right=thin, top=thin, bottom=thin)
WH = Font(bold=True, color="FFFFFF"); BOLD = Font(bold=True); BLUE = Font(color="0000FF")
def title(ws, t, sub=None):
    ws["A1"] = t; ws["A1"].font = Font(bold=True, size=14, color=AZ)
    if sub: ws["A2"] = sub; ws["A2"].font = Font(italic=True, color="595959")
def header(ws, r, cols, c0=1, fill=fH):
    for i, c in enumerate(cols):
        z = ws.cell(r, c0 + i, c); z.font = WH; z.fill = fill; z.border = B
        z.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 32
def widths(ws, w):
    for i, v in enumerate(w, 1): ws.column_dimensions[L(i)].width = v
def box(ws, r, c1, c2, fill=None):
    for j in range(c1, c2 + 1):
        ws.cell(r, j).border = B
        if fill: ws.cell(r, j).fill = fill

wb = Workbook()
# ================= Export_Buk (pegado crudo) =================
ex = wb.active; ex.title = "Export_Buk"
for i, row in enumerate(raw, 1):
    for j, v in enumerate(row, 1):
        if v is not None: ex.cell(i, j, v)
ex.sheet_properties.tabColor = NAR
MAXR = 5000
def X(col):  # columna del export por nombre de encabezado (robusto a cambios de orden)
    return f'INDEX(Export_Buk!$A$7:$AZ${MAXR},0,MATCH("{col}",Export_Buk!$A$6:$AZ$6,0))'
BASE = f'{X("Estado")},"Activo",{X("Nombre de Archivo")},Control!$C$5'
FIRM = f'{X("Estado Firma Trabajador(a)")},"Firmado"'
PEND = f'{X("Estado Firma Trabajador(a)")},"No firmado"'
BLOQ = {"Coronel y plantas": f'{X("Lugar de Trabajo")},"<>Coquimbo",{X("Lugar de Trabajo")},"<>El Golf"',
        "Coquimbo": f'{X("Lugar de Trabajo")},"Coquimbo"', "El Golf (Santiago)": f'{X("Lugar de Trabajo")},"El Golf"'}

# ================= Control =================
c = wb.create_sheet("Control", 0)
title(c, "Control de Avance de Firma — Procedimientos de SST (automatizado)",
      "Para actualizar: en Buk exporta 'Reporte de Firma Documentos OK', copia TODO (desde A1) y pégalo en la hoja Export_Buk reemplazando lo anterior. Todo se recalcula.")
par = [("Parámetros", None),
       ("Filtro de campaña (Nombre de Archivo)", f"*{CAMP}*"),
       ("Fecha de emisión", EMISION),
       ("Plazo de firma (días hábiles)", 5),
       ("Feriados", dt.date(2026, 12, 8)),
       ("", dt.date(2026, 12, 25))]
c["B4"] = "PARÁMETROS"; c["B4"].font = Font(bold=True, color=AZ)
for i, (k, v) in enumerate(par[1:], 5):
    c.cell(i, 2, k).font = BOLD; z = c.cell(i, 3, v); z.fill = fY; z.border = B
    if isinstance(v, dt.date): z.number_format = "dd-mm-yyyy"
FER = "Control!$C$8:$C$9"
c["E4"] = "CORTE ACTUAL (desde el export pegado)"; c["E4"].font = Font(bold=True, color=AZ)
lab = ["Fecha/hora generación export", "Días hábiles desde emisión", "Fecha de vencimiento", "Estado del plazo",
       "Filas en export", "Documentos fuera de campaña (no SST)", "Docs campaña de Inactivos (excluidos)"]
fml = ['=DATE(MID(Export_Buk!A2,FIND("/",Export_Buk!A2)+4,4),MID(Export_Buk!A2,FIND("/",Export_Buk!A2)+1,2),MID(Export_Buk!A2,FIND("/",Export_Buk!A2)-2,2))'
       '+TIME(MOD(VALUE(MID(Export_Buk!A2,FIND("a las",Export_Buk!A2)+6,2)),12)+IF(MID(Export_Buk!A2,FIND("a las",Export_Buk!A2)+11,2)="PM",12,0),VALUE(MID(Export_Buk!A2,FIND("a las",Export_Buk!A2)+9,2)),0)',
       f"=NETWORKDAYS(C6,INT(F5),{FER})-1",
       f"=WORKDAY(C6,C7,{FER})",
       '=IF(F6>C7,"🔴 Vencido",IF(F6>C7-2,"🟡 Por vencer","🟢 En plazo"))',
       f'=COUNTA({X("Nombre de Archivo")})',
       f'=F9-COUNTIFS({X("Nombre de Archivo")},C5)',
       f'=COUNTIFS({X("Nombre de Archivo")},C5,{X("Estado")},"<>Activo")']
for i, (k, f) in enumerate(zip(lab, fml), 5):
    c.cell(i, 5, k).font = BOLD; z = c.cell(i, 6, f); z.border = B
c["F5"].number_format = "dd-mm-yyyy hh:mm"; c["F7"].number_format = "dd-mm-yyyy"
# KPI
r = 13
c.cell(r, 2, "AVANCE GENERAL (trabajadores Activos, documentos de la campaña)").font = Font(bold=True, color=AZ)
header(c, r + 1, ["Emitidos", "Firmados", "Pendientes", "% Avance", "Pend. con fuero", "Pend. ficha F", "Pend. ficha T/E/D"], c0=2)
k = r + 2
c.cell(k, 2, f"=COUNTIFS({BASE})"); c.cell(k, 3, f"=COUNTIFS({BASE},{FIRM})"); c.cell(k, 4, f"=COUNTIFS({BASE},{PEND})")
c.cell(k, 5, f"=IF(B{k}=0,0,C{k}/B{k})").number_format = "0.0%"
c.cell(k, 6, f'=IFERROR(COUNTIFS({BASE},{PEND},{X("Filtro documental")},"*fuero*"),"Falta columna")')
c.cell(k, 7, f'=IFERROR(COUNTIFS({BASE},{PEND},{X("Código de Ficha")},"F*"),"Falta columna")'); c.cell(k, 8, f'=IFERROR(D{k}-G{k},"Falta columna")')
box(c, k, 2, 8); KPI = k
for j in range(2, 9): c.cell(k, j).font = Font(bold=True, size=13)
# Cuadratura
c.cell(k + 2, 2, "Cuadratura con Resumen por Supervisor:").font = BOLD
c.cell(k + 2, 4, f"=IF(D{KPI}='Resumen por Supervisor'!H{{TOT}},\"✅ Cuadra\",\"⚠️ No cuadra: hay un supervisor nuevo o con nombre distinto en el export; revísalo en Resumen por Supervisor\")")
CUAD = (k + 2, 4)
# Bloque
r = k + 5
c.cell(r, 2, "POR BLOQUE").font = Font(bold=True, color=AZ)
header(c, r + 1, ["Bloque", "Emitidos", "Firmados", "Pendientes", "% Avance"], c0=2)
for i, (b, crit) in enumerate(BLOQ.items()):
    rr = r + 2 + i
    c.cell(rr, 2, b); c.cell(rr, 3, f"=COUNTIFS({BASE},{crit})"); c.cell(rr, 4, f"=COUNTIFS({BASE},{crit},{FIRM})")
    c.cell(rr, 5, f"=COUNTIFS({BASE},{crit},{PEND})"); c.cell(rr, 6, f"=IF(C{rr}=0,0,D{rr}/C{rr})").number_format = "0.0%"
    box(c, rr, 2, 6)
# Sindicato
r = r + 7
c.cell(r, 2, "POR SINDICATO").font = Font(bold=True, color=AZ)
header(c, r + 1, ["Sindicato", "Emitidos", "Firmados", "Pendientes", "% Avance"], c0=2)
sinds = d["Sindicato"].value_counts().index.tolist()
for i, s in enumerate(sinds):
    rr = r + 2 + i
    crit = f'{X("Sindicato")},"="' if s == "(Sin sindicato)" else f'{X("Sindicato")},"{s}"'
    c.cell(rr, 2, s); c.cell(rr, 3, f"=COUNTIFS({BASE},{crit})"); c.cell(rr, 4, f"=COUNTIFS({BASE},{crit},{FIRM})")
    c.cell(rr, 5, f"=COUNTIFS({BASE},{crit},{PEND})"); c.cell(rr, 6, f"=IF(C{rr}=0,0,D{rr}/C{rr})").number_format = "0.0%"
    box(c, rr, 2, 6)
rr += 1; c.cell(rr, 2, "TOTAL").font = BOLD
for j in (3, 4, 5): c.cell(rr, j, f"=SUM({L(j)}{r+2}:{L(j)}{rr-1})").font = BOLD
c.cell(rr, 6, f"=IF(C{rr}=0,0,D{rr}/C{rr})").number_format = "0.0%"; box(c, rr, 2, 6, fT)
# Procedimiento SST
r = rr + 3
c.cell(r, 2, "POR PROCEDIMIENTO SST").font = Font(bold=True, color=AZ)
header(c, r + 1, ["Procedimiento", "Emitidos", "Firmados", "Pendientes", "% Avance"], c0=2)
for i, doc in enumerate(sorted(d["Procedimiento SST"].unique())):
    rr = r + 2 + i; crit = f'{X("Nombre de Archivo")},"*{doc}*"'
    c.cell(rr, 2, doc); c.cell(rr, 3, f"=COUNTIFS({BASE},{crit})"); c.cell(rr, 4, f"=COUNTIFS({BASE},{crit},{FIRM})")
    c.cell(rr, 5, f"=COUNTIFS({BASE},{crit},{PEND})"); c.cell(rr, 6, f"=IF(C{rr}=0,0,D{rr}/C{rr})").number_format = "0.0%"
    box(c, rr, 2, 6)
# Área
r = rr + 3
c.cell(r, 2, "POR ÁREA").font = Font(bold=True, color=AZ)
header(c, r + 1, ["Área", "Emitidos", "Firmados", "Pendientes", "% Avance"], c0=2)
for i, a in enumerate(d["Nombre Área"].value_counts().index):
    rr = r + 2 + i; crit = f'{X("Nombre Área")},"{a}"'
    c.cell(rr, 2, a); c.cell(rr, 3, f"=COUNTIFS({BASE},{crit})"); c.cell(rr, 4, f"=COUNTIFS({BASE},{crit},{FIRM})")
    c.cell(rr, 5, f"=COUNTIFS({BASE},{crit},{PEND})"); c.cell(rr, 6, f"=IF(C{rr}=0,0,D{rr}/C{rr})").number_format = "0.0%"
    box(c, rr, 2, 6)
for col in ("F", "E"):
    pass
c.conditional_formatting.add("F8", FormulaRule(formula=['LEFT(F8,2)="🔴"'], fill=fR))
c.conditional_formatting.add("F8", FormulaRule(formula=['LEFT(F8,2)="🟡"'], fill=fA))
c.conditional_formatting.add("F8", FormulaRule(formula=['LEFT(F8,2)="🟢"'], fill=fG))
# Registro de cortes (a la derecha)
c["J13"] = "REGISTRO DE CORTES (lunes, miércoles y viernes 08:00)"; c["J13"].font = Font(bold=True, color=AZ)
header(c, 14, ["Corte", "Emitidos", "Firmados", "Pendientes", "% Avance", "Firmaron en el período", "Fuente"], c0=10)
cortes = [(dt.datetime(2026, 11, 24, 13, 18), 758, 98, 660, "Informe Mensual 13:18 (línea base)"),
          (dt.datetime(2026, 11, 24, 14, 15), 758, 146, 612, "Reporte Firma OK 14:15"),
          (GEN, len(d), int(len(d) - d.Pendiente.sum()), int(d.Pendiente.sum()), f"Reporte Firma OK {GEN:%H:%M}")]
for i, (f, e, fi, p, src) in enumerate(cortes):
    rr = 15 + i
    for j, v in enumerate([f, e, fi, p], 10): z = c.cell(rr, j, v); z.font = BLUE
    c.cell(rr, 10).number_format = "dd-mm-yyyy hh:mm"
    c.cell(rr, 14, f"=IF(K{rr}=0,0,L{rr}/K{rr})").number_format = "0.0%"
    c.cell(rr, 15, "" if i == 0 else f"=M{rr-1}-M{rr}"); c.cell(rr, 16, src); box(c, rr, 10, 16)
for i in range(3, 15):
    rr = 15 + i
    for j in range(10, 14): c.cell(rr, j).fill = fIn
    c.cell(rr, 10).number_format = "dd-mm-yyyy hh:mm"
    c.cell(rr, 14, f'=IF(K{rr}="","",IF(K{rr}=0,0,L{rr}/K{rr}))').number_format = "0.0%"
    c.cell(rr, 15, f'=IF(M{rr}="","",M{rr-1}-M{rr})'); box(c, rr, 10, 16)
c["J30"] = "Cómo registrar: tras pegar el export, copia la fila 'Corte actual' y pégala como VALORES en la primera fila celeste libre."
c["J31"] = "Corte actual →"; c["J31"].font = BOLD
for j, f in zip(range(10, 15), ["=F5", f"=B{KPI}", f"=C{KPI}", f"=D{KPI}", f"=E{KPI}"]):
    z = c.cell(32, j, f); z.border = B; z.fill = fT
c["J32"].number_format = "dd-mm-yyyy hh:mm"; c["N32"].number_format = "0.0%"
ch = LineChart(); ch.title = "Documentos pendientes por corte"; ch.height = 7; ch.width = 16
ch.add_data(Reference(c, min_col=13, min_row=14, max_row=28), titles_from_data=True)
ch.set_categories(Reference(c, min_col=10, min_row=15, max_row=28)); ch.legend = None
c.add_chart(ch, "J34")
widths(c, [2, 44, 22, 11, 44, 16, 13, 13, 3, 17, 10, 10, 11, 10, 13, 32])
c.sheet_view.showGridLines = False

# ================= Resumen por Supervisor (vivo) =================
rs = wb.create_sheet("Resumen por Supervisor", 1)
title(rs, "Avance de firma por Supervisor — se actualiza al pegar un nuevo export",
      "Emitidos/Firmados/Pendientes son fórmulas vivas sobre Export_Buk. Columnas celestes = registro de gestión.")
rs["B3"] = "Corte:"; rs["C3"] = "=Control!F5"; rs["C3"].number_format = "dd-mm-yyyy hh:mm"; rs["B3"].font = BOLD
rs["D3"] = "Estado del plazo:"; rs["E3"] = "=Control!F8"; rs["D3"].font = BOLD
rh = ["N°", "Supervisor", "Correo Supervisor", "Bloque(s)", "Área principal", "Emitidos", "Firmados", "Pendientes",
      "% Avance", "Semáforo avance", "Pend. corte 14:15", "Δ vs 14:15", "Pend. con fuero", "Personas pend. (corte " + f"{GEN:%H:%M})",
      "Fecha correo seguimiento", "Respuesta / compromiso", "Estado gestión"]
header(rs, 5, rh)
for j in (15, 16, 17): rs.cell(5, j).fill = fN
g = (d.groupby("Nombre Supervisor")
       .agg(correo=("Correo Supervisor", lambda s: s.mode().iat[0] if s.notna().any() else ""),
            bloques=("Bloque", lambda s: " / ".join(sorted(s.unique()))),
            area=("Nombre Área", lambda s: s.mode().iat[0]),
            pend=("Pendiente", "sum"), emit=("Pendiente", "size"))
       .sort_values(["pend", "emit"], ascending=False))
pers = pend.groupby("Nombre Supervisor")["Número de Documento"].nunique()
r0 = 6
for i, (sup, r) in enumerate(g.iterrows()):
    rr = r0 + i; S = f"$B{rr}"; SC = f'{X("Nombre Supervisor")},{S}'
    for j, v in enumerate([i + 1, sup, r.correo, r.bloques, r.area], 1): rs.cell(rr, j, v)
    rs.cell(rr, 6, f"=COUNTIFS({BASE},{SC})"); rs.cell(rr, 7, f"=COUNTIFS({BASE},{SC},{FIRM})")
    rs.cell(rr, 8, f"=COUNTIFS({BASE},{SC},{PEND})")
    rs.cell(rr, 9, f"=IF(F{rr}=0,0,G{rr}/F{rr})").number_format = "0.0%"
    rs.cell(rr, 10, f'=IF(F{rr}=0,"—",IF(H{rr}=0,"✅ Completo",IF(I{rr}>=0.8,"🟢 ≥80%",IF(I{rr}>=0.5,"🟡 50-79%","🔴 <50%"))))')
    z = rs.cell(rr, 11, int(prev_sup.get(NORM(sup), 0))); z.font = BLUE
    rs.cell(rr, 12, f"=H{rr}-K{rr}")
    rs.cell(rr, 13, f'=IFERROR(COUNTIFS({BASE},{SC},{PEND},{X("Filtro documental")},"*fuero*"),"s/d")')
    z = rs.cell(rr, 14, int(pers.get(sup, 0))); z.font = BLUE
    box(rs, rr, 1, 17)
    for j in (15, 16, 17): rs.cell(rr, j).fill = fIn
    rs.cell(rr, 15).number_format = "dd-mm-yyyy"
rs_last = r0 + len(g) - 1; TOT = rs_last + 1
rs.cell(TOT, 2, "TOTAL").font = BOLD
for j in (6, 7, 8, 11, 12, 13, 14): rs.cell(TOT, j, f"=SUM({L(j)}{r0}:{L(j)}{rs_last})").font = BOLD
rs.cell(TOT, 14, pend["Número de Documento"].nunique()).font = Font(bold=True, color="0000FF")
rs.cell(TOT, 9, f"=IF(F{TOT}=0,0,G{TOT}/F{TOT})").number_format = "0.0%"; box(rs, TOT, 1, 17, fT)
dv = DataValidation(type="list", formula1='"Correo enviado,Comprometido,Sin respuesta,Escalado a jefatura,Cerrado"', allow_blank=True)
rs.add_data_validation(dv); dv.add(f"Q{r0}:Q{rs_last}")
J = f"J{r0}:J{rs_last}"
for pre, fill in (("🔴", fR), ("🟡", fA), ("🟢", fG), ("✅", fG)):
    rs.conditional_formatting.add(J, FormulaRule(formula=[f'LEFT(J{r0},1)="{pre[0]}"'], fill=fill))
rs.conditional_formatting.add(f"L{r0}:L{rs_last}", FormulaRule(formula=[f"L{r0}<0"], fill=fG))
lg = TOT + 2
for i, t in enumerate(["¿QUÉ MIDE EL SEMÁFORO DE AVANCE? (% de documentos firmados del equipo del supervisor)",
                       "✅ Completo = sin pendientes · 🟢 ≥80% firmado · 🟡 50–79% firmado · 🔴 <50% firmado.",
                       "Estado del plazo (arriba) = días hábiles desde la emisión vs plazo: 🟢 en plazo · 🟡 vence en ≤2 días · 🔴 vencido.",
                       "Azul = valor fijo calculado en Python (corte anterior y personas únicas); no cambia al pegar un nuevo export.",
                       "Si Control muestra '⚠️ No cuadra', apareció un supervisor nuevo: agrega su nombre en una fila nueva y copia las fórmulas."]):
    rs.cell(lg + i, 2, t).font = BOLD if i == 0 else Font(color="404040")
rs.auto_filter.ref = f"A5:Q{rs_last}"; rs.freeze_panes = "C6"
widths(rs, [5, 34, 34, 22, 36, 9, 9, 10, 9, 13, 10, 9, 9, 12, 13, 34, 18])
rs.sheet_view.showGridLines = False
c.cell(*CUAD).value = c.cell(*CUAD).value.replace("{TOT}", str(TOT))

# ================= Detalle Pendientes (respaldo del corte) =================
det = wb.create_sheet("Detalle Pendientes", 2)
title(det, f"Respaldo: pendientes al corte {GEN:%d-%m-%Y %H:%M} (lista enviada a supervisores)",
      "Foto del corte, no cambia al pegar un nuevo export. Columnas celestes = seguimiento por corte (L-Mi-V 08:00).")
SEG = ["Mié 25-11", "Vie 27-11", "Lun 30-11", "Mié 02-12", "Vie 04-12"]
cols = ["Bloque", "Nombre Área", "Nombre Sub-área Asignada(o)", "Nombre Supervisor", "Correo Supervisor", "Nombre Completo",
        "Número de Documento", "Código de Ficha", "Prefijo ficha", "Cargo", "Lugar de Trabajo", "Sindicato", "Procedimiento SST",
        "Nombre de Archivo", "Fecha Creación", "Fuero", "Filtro Documental (vigente)", "Email Personal"]
calc = ["Fecha vencimiento", "Días hábiles al corte", "Semáforo plazo"]
hdr = cols + calc + [f"Seguimiento {s}" for s in SEG] + ["Observación"]
HR = 4; header(det, HR, hdr)
for j in range(len(cols) + len(calc) + 1, len(hdr) + 1): det.cell(HR, j).fill = fN
pend = pend.sort_values(["Bloque", "Nombre Supervisor", "Nombre Completo"])
cF, cD, cS = [L(len(cols) + i) for i in (1, 2, 3)]; cCre = L(cols.index("Fecha Creación") + 1)
for i, (_, r) in enumerate(pend.iterrows()):
    rr = HR + 1 + i
    for j, col in enumerate(cols, 1):
        v = r[col]; v = v.to_pydatetime() if isinstance(v, pd.Timestamp) else v
        det.cell(rr, j, None if pd.isna(v) else v).border = B
    det[f"{cCre}{rr}"].number_format = "dd-mm-yyyy hh:mm"
    det[f"{cF}{rr}"] = f"=WORKDAY(INT({cCre}{rr}),Control!$C$7,{FER})"; det[f"{cF}{rr}"].number_format = "dd-mm-yyyy"
    det[f"{cD}{rr}"] = f"=NETWORKDAYS(INT({cCre}{rr}),INT(Control!$F$5),{FER})-1"
    det[f"{cS}{rr}"] = f'=IF({cD}{rr}>Control!$C$7,"🔴 Vencido",IF({cD}{rr}>Control!$C$7-2,"🟡 Por vencer","🟢 En plazo"))'
    for j in range(len(cols) + 1, len(hdr) + 1): det.cell(rr, j).border = B
    for j in range(len(cols) + len(calc) + 1, len(hdr) + 1): det.cell(rr, j).fill = fIn
det_last = HR + len(pend)
dv2 = DataValidation(type="list", formula1='"Firmado,Avisado,Comprometido,Sin respuesta,Licencia/Ausente"', allow_blank=True)
det.add_data_validation(dv2); dv2.add(f"{L(len(cols)+len(calc)+1)}{HR+1}:{L(len(hdr)-1)}{det_last}")
rng = f"{cS}{HR+1}:{cS}{det_last}"
for pre, fill in (("🔴", fR), ("🟡", fA), ("🟢", fG)):
    det.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({cS}{HR+1},2)="{pre}"'], fill=fill))
det.auto_filter.ref = f"A{HR}:{L(len(hdr))}{det_last}"; det.freeze_panes = f"G{HR+1}"
widths(det, [16, 30, 30, 30, 30, 32, 13, 8, 7, 28, 13, 24, 15, 50, 16, 11, 14, 28, 12, 10, 13] + [12] * 5 + [30])
det.sheet_view.showGridLines = False

# ================= Notas =================
nt = wb.create_sheet("Notas y Supuestos", 0)
title(nt, "Notas y Supuestos — Control y respaldo de gestión, firma Procedimientos de SST (Orizon S.A.)", f"Corte: {GEN:%d-%m-%Y %H:%M}")
obs = d["Filtro documental"].astype("string").str.strip().isin(OBSOLETOS)
n_two = int((pend.groupby("Número de Documento").size() > 1).sum())
top = g.head(10)
notas = [
 ("1. Fuente y alcance", [
  f"Fuente: Buk, 'Reporte de Firma Documentos OK 2026', generado {GEN:%d-%m-%Y %H:%M} ({n_rows} documentos creados el 24-11-2026, 29 columnas).",
  f"Campaña = documentos cuyo Nombre de Archivo contiene '{CAMP}' (procedimientos de SST MR-005/006/007/008/010, confirmado por Belén 24-11-2026). Se excluyen {n_fuera} documentos del export que no son de la campaña (finiquitos, certificados, contrato, anexo IPC/colación); de ellos 5 están 'No firmado' y no se gestionan aquí.",
  f"Solo Activos: se excluyen {n_inact} procedimientos de SST de trabajadores Inactivos.",
  "Incluye Coquimbo y El Golf como bloques separados de 'Coronel y plantas' (instrucción de Belén 24-11-2026). Bloque según Lugar de Trabajo.",
  "Se agrupa por Nombre Área (Nombre Gerencia viene como 'ORIZON S.A' en Buk)."]),
 ("2. Pendiente y semáforos", [
  "Pendiente = Estado Firma Trabajador(a) 'No firmado'. Firmantes empresa: 'No Requiere Firma' en toda la campaña → todo depende del trabajador(a).",
  f"Resultado del corte: {len(pend)} pendientes / {pend['Número de Documento'].nunique()} personas / {int((g.pend>0).sum())} supervisores con pendientes, de {len(d)} emitidos ({1-len(pend)/len(d):.1%} firmado).",
  "Estado del plazo: días hábiles desde la emisión (no desde el ingreso). Plazo 5 días hábiles editable en Control!C7; vence martes 01-12-2026. 🟢 ≤ plazo−2 · 🟡 ≤ plazo · 🔴 > plazo.",
  "Semáforo de avance por supervisor: % firmado de su equipo (🟢 ≥80% · 🟡 50–79% · 🔴 <50% · ✅ completo)."]),
 ("3. Cómo funciona la automatización", [
  "Pegar el export completo de Buk (desde A1, con las 5 filas de encabezado) en Export_Buk. Control y Resumen por Supervisor se recalculan solos; las columnas se buscan por nombre de encabezado, así que un cambio de orden en Buk no rompe las fórmulas.",
  "Detalle Pendientes es la foto del corte enviado a los supervisores (respaldo de gestión); no cambia al pegar.",
  "Registro de cortes: pegar como valores la fila 'Corte actual' en cada seguimiento lunes, miércoles y viernes 08:00.",
  "Cuadratura automática (Control): si un export trae un supervisor que no está en la lista, aparece '⚠️ No cuadra'.",
  "Límite: el export se lee hasta la fila 5.000."]),
 ("4. Calidad de datos", [
  "Duplicados (RUT + Nombre de Archivo): 0.",
  f"{n_two} personas con 2 documentos pendientes: son procedimientos de SST distintos, no duplicados.",
  f"Filtro documental con código obsoleto (ETI/A1–A4/BB/E): {int(obs.sum())} documentos ({int((obs & (d.Pendiente==1)).sum())} pendientes) — en blanco en 'Filtro Documental (vigente)'.",
  f"Pendientes con fuero: {int((pend.Fuero=='Con fuero').sum())}. Fichas F vs T/E/D separadas (Control).",
  "Este export no trae Visibilidad ni Email corporativo; el Informe Mensual de las 13:18 mostraba 0 pendientes ocultos.",
  "Corte 13:18 viene del Informe Mensual (mismos 769 documentos, calce 769/769); cortes 14:15 y siguientes del Reporte de Firma OK."]),
]
rr = 4
for t, items in notas:
    nt.cell(rr, 1, t).font = Font(bold=True, color=AZ, size=12); rr += 1
    for it in items:
        z = nt.cell(rr, 1, "• " + it); z.alignment = Alignment(wrap_text=True, vertical="top"); rr += 1
    rr += 1
nt.cell(rr, 1, "RECOMENDACIÓN DE GESTIÓN").font = Font(bold=True, color="FFFFFF", size=12); nt.cell(rr, 1).fill = fN; rr += 1
reco = (f"En las primeras 2,5 horas firmaron 58 documentos (660 → {len(pend)}), sin gestión activa: el avance orgánico ya se está frenando "
        f"(48 en la 1ª hora, 10 en los 20 minutos siguientes). Enviar hoy el correo de seguimiento a los {int((g.pend>0).sum())} supervisores; "
        f"los 10 primeros concentran {int(top.pend.sum())} pendientes ({top.pend.sum()/len(pend):.0%}) y son la prioridad de llamada del lunes 30-11. "
        "Validar fuero/licencias antes de escalar. Al cierre (04-12), los no firmados van a la Bitácora de Compromisos en Notion, previa validación de jefatura.")
z = nt.cell(rr, 1, reco); z.alignment = Alignment(wrap_text=True, vertical="top"); nt.row_dimensions[rr].height = 80
nt.column_dimensions["A"].width = 150; nt.sheet_view.showGridLines = False

wb._sheets = [wb[n] for n in ["Notas y Supuestos", "Control", "Resumen por Supervisor", "Detalle Pendientes", "Export_Buk"]]
wb.active = 1
wb.save(OUT)
pend.to_pickle(OUT + ".pend.pkl"); g.to_pickle(OUT + ".sup.pkl"); d.to_pickle(OUT + ".all.pkl")
print("ok", GEN, len(d), len(pend), pend["Número de Documento"].nunique(), int((g.pend > 0).sum()), int(top.pend.sum()), "TOT", TOT)
