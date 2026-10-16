"""Graficos nativos con libro de datos embebido, para usar en DOCX y PPTX.

Genera: la parte chartN.xml (DrawingML Chart), su .rels y el libro
embeddings/DatosGraficoN.xlsx con las categorias y series, de modo que el
grafico queda totalmente editable ("Editar datos" abre el libro).
"""
import io, zipfile
import xlsxw as X
from xlsxw import Sheet, C
from ooxml import esc, escA, rels, R_OFFICE


class GChart:
    def __init__(self, kind, title, cats, series, stacked=False, show_val=True,
                 colors=None, numfmt='General', cat_title='', val_title='', legend=None):
        """series = [(nombre, [valores])]"""
        self.kind, self.title, self.cats, self.series = kind, title, cats, series
        self.stacked, self.show_val = stacked, show_val
        self.colors = colors or X.SERIES
        self.numfmt = numfmt
        self.cat_title, self.val_title = cat_title, val_title
        self.legend = legend if legend is not None else (len(series) > 1 or kind in ('pie', 'doughnut'))


def _wb_bytes(ch):
    """Libro de datos embebido del grafico."""
    sh = Sheet('Hoja1', cols=[28] + [14] * len(ch.series))
    sh.row(C('', 'base'), *[C(nm, 'bold') for nm, _ in ch.series])
    for i, cat in enumerate(ch.cats):
        sh.row(C(cat, 'txt'), *[C(vals[i] if i < len(vals) else None, 'num2') for _, vals in ch.series])
    buf = io.BytesIO()
    import tempfile, os
    fd, tmp = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    X.write(tmp, [sh], ch.title or 'Datos')
    data = open(tmp, 'rb').read()
    os.unlink(tmp)
    return data


def chart_part_xml(ch, with_external=True):
    n = len(ch.cats)
    cat_ref = "Hoja1!$A$2:$A$%d" % (n + 1)
    sers = []
    for i, (nm, vals) in enumerate(ch.series):
        col = X.colname(i + 1)
        sers.append((nm, "Hoja1!$%s$2:$%s$%d" % (col, col, n + 1), vals,
                     "Hoja1!$%s$1" % col))
    tmp = X.Chart(ch.kind, ch.title, 'Hoja1', cat_ref, ch.cats,
                  [(nm, ref, vals) for nm, ref, vals, _ in sers], 'A1',
                  stacked=ch.stacked, show_val=ch.show_val, colors=ch.colors,
                  cat_title=ch.cat_title, val_title=ch.val_title, numfmt=ch.numfmt)
    xml = X.chart_xml(tmp)
    if not ch.legend:
        import re
        xml = re.sub(r'<c:legend>.*?</c:legend>', '', xml, flags=re.S)
    if with_external:
        xml = xml.replace('</c:chart>', '</c:chart><c:externalData r:id="rIdData">'
                                        '<c:autoUpdate val="0"/></c:externalData>')
        # externalData debe ir despues de spPr/txPr: se reubica al final de chartSpace
        xml = xml.replace('</c:chart><c:externalData r:id="rIdData">'
                          '<c:autoUpdate val="0"/></c:externalData>', '</c:chart>')
        xml = xml.replace('</c:chartSpace>', '<c:externalData r:id="rIdData">'
                                             '<c:autoUpdate val="0"/></c:externalData></c:chartSpace>')
    return xml


def add_charts(pkg, folder, charts, start=1):
    """Agrega las partes de los graficos al paquete.
    folder = 'word' o 'ppt'.  Devuelve lista de rutas de chart parts (relativas al folder)."""
    paths = []
    for i, ch in enumerate(charts, start=start):
        cpath = '%s/charts/chart%d.xml' % (folder, i)
        epath = '%s/embeddings/DatosGrafico%d.xlsx' % (folder, i)
        pkg.add(cpath, chart_part_xml(ch),
                'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
        pkg.add(epath, _wb_bytes(ch),
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        pkg.add('%s/charts/_rels/chart%d.xml.rels' % (folder, i), rels([
            ('rIdData', R_OFFICE + '/package', '../embeddings/DatosGrafico%d.xlsx' % i),
        ]))
        paths.append('charts/chart%d.xml' % i)
    return paths
