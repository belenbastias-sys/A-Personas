import sys, pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
F=sys.argv[1]; OUT=sys.argv[2]; CORTE=sys.argv[3]
p=pd.read_pickle(F+'.pend.pkl'); g=pd.read_pickle(F+'.sup.pkl'); g=g[g.pend>0]
NAR=RGBColor(0x2E,0x75,0xB6); AZ=RGBColor(0x1F,0x4E,0x79); W=RGBColor(255,255,255)
doc=Document()
for s in doc.sections: s.left_margin=s.right_margin=Cm(2); s.top_margin=s.bottom_margin=Cm(1.8)
doc.styles['Normal'].font.name='Calibri'; doc.styles['Normal'].font.size=Pt(10.5)
def shade(cell,h):
    pr=cell._tc.get_or_add_tcPr(); sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:fill'),h); pr.append(sh)
def hrow(t,labels,color):
    for c,x in zip(t.rows[0].cells,labels):
        c.text=x; r=c.paragraphs[0].runs[0]; r.font.bold=True; r.font.color.rgb=W; r.font.size=Pt(9.5); shade(c,color)
h=doc.add_heading('Correos de seguimiento — Firma de procedimientos de SST',0)
for r in h.runs: r.font.color.rgb=AZ
doc.add_paragraph(f'Corte: {CORTE} · Emisión: 24-11-2026 · Plazo de firma: martes 01-12-2026 (5 días hábiles). '
  f'{len(p)} procedimientos pendientes de firma de {p["Número de Documento"].nunique()} personas, {len(g)} supervisores. '
  'Un correo por supervisor, ordenados de mayor a menor pendiente. Respaldo: '+__import__('os').path.basename(F)+' (hoja Detalle Pendientes).')
t=doc.add_table(rows=1,cols=5); t.style='Table Grid'; hrow(t,['N°','Supervisor','Bloque(s)','Firmados / Emitidos','Pendientes'],'1F4E79')
for i,(s,r) in enumerate(g.iterrows(),1):
    cs=t.add_row().cells
    for c,v in zip(cs,[i,s,r.bloques,f'{int(r.emit-r.pend)} / {int(r.emit)}',int(r.pend)]): c.text=str(v); c.paragraphs[0].runs[0].font.size=Pt(9)
for i,(s,r) in enumerate(g.iterrows(),1):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    hh=doc.add_heading(f'{i}. {s}',level=1)
    for x in hh.runs: x.font.color.rgb=NAR
    q=p[p['Nombre Supervisor']==s].sort_values(['Fuero','Nombre Completo'])
    n=len(q); npp=q['Número de Documento'].nunique(); firm=int(r.emit-r.pend)
    for k,v in [('Para:',f'{s} <{r.correo}>'),('Asunto:',f'Seguimiento firma procedimientos de SST — {npp} trabajador(es) pendiente(s) — plazo 01-12-2026')]:
        pp=doc.add_paragraph(); a=pp.add_run(k+' '); a.bold=True; pp.add_run(v)
    doc.add_paragraph(f'Hola {s.split()[0]}:')
    pp=doc.add_paragraph('El 24-11-2026 se publicaron en Buk procedimientos de Seguridad y Salud en el Trabajo (SST) que tu equipo debe firmar. ')
    pp.add_run(f'Al corte de hoy ({CORTE}) van {firm} de {int(r.emit)} procedimientos firmados ({firm/r.emit:.0%}), ')
    b=pp.add_run(f'quedan {n} procedimiento(s) pendiente(s) de firma de {npp} persona(s):'); b.bold=True
    tb=doc.add_table(rows=1,cols=5); tb.style='Table Grid'; hrow(tb,['Nombre','RUT','Cargo','Procedimiento SST','Lugar'],'2E75B6')
    for _,x in q.iterrows():
        cs=tb.add_row().cells
        nm=x['Nombre Completo']+(' (fuero)' if x['Fuero']=='Con fuero' else '')
        for c,v in zip(cs,[nm,x['Número de Documento'],x['Cargo'],x['Procedimiento SST'],x['Lugar de Trabajo']]):
            c.text=str(v); c.paragraphs[0].runs[0].font.size=Pt(9)
    doc.add_paragraph()
    pp=doc.add_paragraph('Te pedimos apoyo para que firmen desde la app o portal de Buk a más tardar el ')
    pp.add_run('martes 01-12-2026').bold=True
    pp.add_run('. Si alguien está con licencia, vacaciones o no puede ingresar a Buk, respóndenos este correo para gestionarlo. '
               'Haremos seguimiento los lunes, miércoles y viernes a las 08:00 y te enviaremos la lista actualizada.')
    doc.add_paragraph('Saludos,\nEquipo SST Orizon S.A.')
doc.save(OUT); print('ok',len(g))
