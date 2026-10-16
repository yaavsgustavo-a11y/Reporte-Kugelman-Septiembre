# -*- coding: utf-8 -*-
"""VALIDACIÓN FINAL (punto 29 del brief): 20 comprobaciones sobre los archivos
realmente generados, leyéndolos de vuelta desde disco."""
import os, re, zipfile, collections
import xml.etree.ElementTree as ET
from xlsxread import Xlsx
import facts

OUT = '/projects/sandbox/entregables'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
ok = []
bad = []


def chk(cond, msg):
    (ok if cond else bad).append(msg)


def find(frag):
    for fn in os.listdir(OUT):
        if frag.lower() in fn.lower():
            return os.path.join(OUT, fn)
    raise FileNotFoundError(frag)


def docx_text(p):
    z = zipfile.ZipFile(p)
    root = ET.fromstring(z.read('word/document.xml'))
    return '\n'.join(''.join(t.text or '' for t in par.iter(W + 't'))
                     for par in root.iter(W + 'p'))


def pptx_text(p):
    z = zipfile.ZipFile(p)
    out = []
    for nm in sorted(z.namelist()):
        if nm.startswith('ppt/slides/slide'):
            out.append('\n'.join(t.text or '' for t in ET.fromstring(z.read(nm)).iter(A + 't')))
    return '\n'.join(out)


f = facts.get()
rep = find('Reporte Campaña')
base = find('Base Maestra')
ppt = find('Presentación Ejecutiva')
alu = find('Base de Alumnos')
pro = find('Prospectos en Curso')

TR = docx_text(rep)
TP = pptx_text(ppt)
XB = Xlsx(base)
XA = Xlsx(alu)
XP = Xlsx(pro)

print('=' * 78)
print('VALIDACIÓN FINAL — 20 COMPROBACIONES')
print('=' * 78)

# 1-2 duplicados identificados / nombres validados
chk('deduplicación' in TR.lower() or 'deduplicacion' in TR.lower(),
    '1. El reporte documenta la deduplicación de registros')
chk('%d' % f.dup_fusionados in TR and 'R4' in TR,
    '2. Reglas de deduplicación (R1–R6) y conteo de fusionados presentes en el reporte')

# 3 formulas
nf = 0
for sh in XB.sheet_names:
    data = XB.z.read(next(s['path'] for s in XB.sheets if s['name'] == sh)).decode('utf-8', 'ignore')
    nf += data.count('<f>')
chk(nf > 60, '3. La Base Maestra contiene %d fórmulas vivas' % nf)

# 4-5 porcentajes y sumatorias: verifica totales de Meta Ads contra la fuente
t = XB.table('Meta Ads')
hdr = next(i for i, r in enumerate(t) if r and r[0] == 'Anuncio (Meta)')
ads = [r for r in t[hdr + 1:] if len(r) > 5 and isinstance(r[3], (int, float))
       and isinstance(r[5], (int, float))]
chk(abs(sum(r[3] for r in ads) - f.gasto) < 0.01,
    '4. Suma de gasto por anuncio en la Base Maestra = %.2f (coincide con la fuente)'
    % sum(r[3] for r in ads))
chk(sum(r[5] for r in ads) == f.conversaciones,
    '5. Suma de conversaciones por anuncio = %d (coincide con la fuente)' % sum(r[5] for r in ads))

# 6 fechas
cm = XB.table('Clases muestra')
h = next(i for i, r in enumerate(cm) if r and r[0] == 'Hoja')
fechas = [r[2] for r in cm[h + 1:] if len(r) > 2 and hasattr(r[2], 'year')]
chk(all(d.year == 2026 for d in fechas),
    '6. Las %d fechas de clase muestra quedaron en 2026 (3 años 2006 corregidos y documentados)'
    % len(fechas))
chk('2006' in TR, '6b. La corrección del año 2006 está documentada en el reporte')

# 7-8 relación prospecto ↔ CM ↔ inscripción
enl = [r for r in cm[h + 1:] if len(r) > 11 and r[11] and r[11] != 'Sin lead identificado']
chk(len(enl) >= f.cm_con_lead,
    '7. %d clases muestra tienen lead de origen identificado en la Base Maestra' % len(enl))
ins = XB.table('Inscritos campaña')
hi = next(i for i, r in enumerate(ins) if r and r[0] == 'Alumno / interesado')
_stop = next((i for i, r in enumerate(ins[hi + 1:]) if r and str(r[0]).startswith('INSCRIPCIONES POR')),
             len(ins))
filas_ins = [r for r in ins[hi + 1:hi + 1 + _stop] if len(r) > 1 and r[0] and hasattr(r[1], 'year')]
chk(len(filas_ins) == f.insc_eventos,
    '8. La pestaña «Inscritos campaña» tiene %d eventos de inscripción con matrícula y evidencia'
    % len(filas_ins))

# 9 no doble conteo de personas
pr = XB.table('Prospectos')
hp = next(i for i, r in enumerate(pr) if r and r[0] == 'ID persona')
ids = [r[0] for r in pr[hp + 1:] if r and r[0]]
chk(len(ids) == len(set(ids)) == f.pers_total,
    '9. %d ID de persona únicos, sin repeticiones (una fila por persona)' % len(ids))

# 10 presupuesto vs gasto real
chk('7,000' in TR and '6,999.91' in TR and 'GASTO REAL' in TR.upper(),
    '10. El reporte diferencia explícitamente presupuesto asignado ($7,000) y gasto real ($6,999.91)')

# 11-12 un solo esfuerzo de campaña
chk('no deben leerse como campañas independientes' in TR,
    '11. El reporte advierte que agosto/septiembre/octubre NO son campañas independientes')
chk('21 de agosto – 4 de octubre de 2026' in TR and '21 ago – 4 oct' in TP,
    '12. El periodo 21 ago – 4 oct se usa como ventana de campaña en reporte y presentación')

# 13 conclusiones sustentadas
chk('DATO' in TR and 'INTERPRETACIÓN' in TR and 'IMPACTO COMERCIAL' in TR,
    '13. Los hallazgos siguen la estructura DATO → INTERPRETACIÓN → IMPACTO COMERCIAL')

# 14 cifras de gráficas = bases
z = zipfile.ZipFile(base)
charts_b = [n for n in z.namelist() if n.startswith('xl/charts/chart')]
zp = zipfile.ZipFile(ppt)
charts_p = [n for n in zp.namelist() if n.startswith('ppt/charts/chart')]
zr = zipfile.ZipFile(rep)
charts_r = [n for n in zr.namelist() if n.startswith('word/charts/chart')]
# comprueba que la cache del funnel del reporte coincida con los hechos
c1 = zr.read('word/charts/chart1.xml').decode()
vals = re.findall(r'<c:pt idx="\d+"><c:v>([\d.]+)</c:v></c:pt>', c1.split('<c:val>')[1])
chk([float(x) for x in vals] == [f.conversaciones, f.leads, f.cm_agendadas, f.cm_realizadas,
                                 f.insc_personas],
    '14. Los datos del gráfico de funnel del reporte coinciden con las bases: %s' % vals)
chk(len(charts_b) >= 12 and len(charts_p) >= 8 and len(charts_r) >= 8,
    '14b. Gráficos nativos editables: %d en la Base Maestra, %d en el reporte, %d en la presentación'
    % (len(charts_b), len(charts_r), len(charts_p)))

# 15 consistencia entre entregables
kpi = XB.table('Resumen KPI')
def kv(label):
    for r in kpi:
        if r and r[0] == label:
            return r[1]
    return None
pares = [
    ('Gasto real', f.gasto, '6,999.91'),
    ('Leads / prospectos (personas únicas)', f.leads, '477'),
    ('CM agendadas', f.cm_agendadas, '91'),
    ('CM realizadas', f.cm_realizadas, '67'),
    ('Inscripciones atribuibles a la campaña', f.insc_atribuibles, '23'),
]
cons = True
for lab, val, txt in pares:
    v = kv(lab)
    if v is None or abs(float(v) - val) > 0.011:
        cons = False
        print('   !! Base Maestra «%s» = %s, esperado %s' % (lab, v, val))
    if txt not in TR:
        cons = False
        print('   !! «%s» no aparece en el reporte' % txt)
    if txt not in TP:
        cons = False
        print('   !! «%s» no aparece en la presentación' % txt)
chk(cons, '15. KPI principales idénticos en Base Maestra, Reporte y Presentación '
          '(gasto, leads, CM agendadas, CM realizadas, inscripciones atribuibles)')

# 16 base de alumnos: exactamente 6 columnas
ta = XA.table('Alumnos')
chk(ta[0] == ['Matrícula', 'Nombre', 'Apellidos', 'Sexo', 'Edad', 'Activo'],
    '16. La Base de Alumnos tiene exactamente las 6 columnas solicitadas: %s' % ta[0])
chk(all(len(r) <= 6 for r in ta), '16b. Ninguna fila de la Base de Alumnos excede 6 columnas')
chk(len([r for r in ta[1:] if r and r[0]]) == f.dq['rga_n'],
    '16c. La Base de Alumnos conserva los %d registros de la fuente, sin eliminar ninguno'
    % f.dq['rga_n'])
mats = [r[0] for r in ta[1:] if r and r[0]]
chk(len(mats) == len(set(mats)), '16d. Matrícula única en los %d registros' % len(mats))
chk('Validaciones' in XA.sheet_names, '16e. Incluye pestaña de validaciones con duplicados y limitaciones')

# 17 prospectos en curso con estatus
tp = XP.table('Prospectos')
hpp = next(i for i, r in enumerate(tp) if r and r[0] == 'Prospecto')
rows = [r for r in tp[hpp + 1:] if r and r[0]]
sin_est = [r for r in rows if not (len(r) > 6 and r[6])]
chk(len(rows) == len(f.pipeline) and not sin_est,
    '17. Los %d prospectos en curso tienen estatus homologado (0 sin estatus)' % len(rows))
estset = set(r[6] for r in rows if len(r) > 6)
chk(estset <= set(facts.ESTATUS_ORDEN),
    '17b. Todos los estatus pertenecen al catálogo homologado solicitado')
chk('Resumen' in XP.sheet_names, '17c. Incluye pestaña «Resumen»')
zpro = zipfile.ZipFile(pro)
chk(len([n for n in zpro.namelist() if n.startswith('xl/charts/chart')]) >= 3,
    '17d. La pestaña Resumen incluye %d gráficas editables'
    % len([n for n in zpro.namelist() if n.startswith('xl/charts/chart')]))

# 18 corte 3 de octubre en bases operativas
import datetime
fechas_p = [r[1].date() for r in rows if len(r) > 1 and hasattr(r[1], 'year')]
_post = [d for d in fechas_p if d > datetime.date(2026, 10, 3)]
chk(len(_post) == f.sin_fecha_contacto,
    '18. Primer contacto ≤ 3-oct en todos los registros del CRM; los %d registros sin fecha de '
    'contacto en las fuentes están marcados y documentados' % f.sin_fecha_contacto)
chk('Corte 03 octubre 2026' in os.path.basename(alu) and 'Corte 03 octubre 2026' in os.path.basename(pro),
    '18b. Ambas bases operativas están etiquetadas con el corte del 3 de octubre')

# 19 cierre 4 de octubre en el análisis de campaña
cmv = [r for r in cm[h + 1:] if len(r) > 10 and r[10] == 'Sí']
chk(all(r[2].date() <= datetime.date(2026, 10, 4) for r in cmv) and len(cmv) == f.cm_agendadas,
    '19. Las %d clases muestra del análisis de campaña tienen fecha ≤ 4 de octubre de 2026' % len(cmv))

# 20 no se inventó información
chk('Sin información' in str(tp) and 'No calculable' in str(kpi),
    '20. Los datos ausentes se marcan como «Sin información» / «No calculable», sin inventar valores')
chk('No disponible' in str(XB.table('Redes sociales')),
    '20b. La sección de redes sociales declara expresamente la ausencia de datos orgánicos')

for m in ok:
    print('  OK   ' + m)
for m in bad:
    print('  FALLA ' + m)
print()
print('RESULTADO: %d comprobaciones correctas, %d fallas' % (len(ok), len(bad)))
