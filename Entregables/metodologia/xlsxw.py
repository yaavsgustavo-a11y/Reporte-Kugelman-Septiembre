"""Escritor XLSX minimo pero completo (stdlib):
hojas, estilos, formatos numericos, autofiltro, paneles fijos, anchos de columna,
formulas, tablas nativas y graficos nativos vinculados a los rangos de la hoja.
"""
import datetime
from ooxml import Pkg, esc, escA, rels, core_props, R_OFFICE, R_PKG

# ------------------------------------------------------------------ paleta
AZUL = '1F3864'     # titulo / encabezado
AZUL2 = '2E5C9A'
AZUL3 = 'D9E2F3'
GRIS = 'F2F2F2'
VERDE = '2E7D5B'
AMBAR = 'C98A1B'
ROJO = 'B23A33'
BLANCO = 'FFFFFF'
SERIES = ['1F3864', '2E8BC0', 'E0A030', '2E7D5B', 'B23A33', '7A5EA8', '6B7280', '46A3B0']


def colname(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


# formatos numericos personalizados (ids >= 164)
NUMFMTS = [
    (164, '&quot;$&quot;#,##0.00'),
    (165, '&quot;$&quot;#,##0'),
    (166, '#,##0'),
    (167, '0.0%'),
    (168, 'dd/mm/yyyy'),
    (169, '0.00'),
    (170, '0.0'),
    (171, '0%'),
    (172, '#,##0.00'),
]

# (nombre, bold, size, color, fill, numfmt, align, wrap, border, italic)
STYLES = [
    ('base',        0, 10, '000000', None,  0,   'left',   0, 0, 0),
    ('titulo',      1, 16, AZUL,     None,  0,   'left',   0, 0, 0),
    ('subtitulo',   1, 12, AZUL2,    None,  0,   'left',   0, 0, 0),
    ('hdr',         1, 10, BLANCO,   AZUL,  0,   'center', 1, 1, 0),
    ('hdr2',        1, 10, AZUL,     AZUL3, 0,   'center', 1, 1, 0),
    ('txt',         0, 10, '000000', None,  0,   'left',   0, 1, 0),
    ('txtw',        0, 10, '000000', None,  0,   'left',   1, 1, 0),
    ('num',         0, 10, '000000', None,  166, 'right',  0, 1, 0),
    ('money',       0, 10, '000000', None,  164, 'right',  0, 1, 0),
    ('money0',      0, 10, '000000', None,  165, 'right',  0, 1, 0),
    ('pct',         0, 10, '000000', None,  167, 'right',  0, 1, 0),
    ('pct0',        0, 10, '000000', None,  171, 'right',  0, 1, 0),
    ('dec1',        0, 10, '000000', None,  170, 'right',  0, 1, 0),
    ('dec2',        0, 10, '000000', None,  169, 'right',  0, 1, 0),
    ('num2',        0, 10, '000000', None,  172, 'right',  0, 1, 0),
    ('fecha',       0, 10, '000000', None,  168, 'center', 0, 1, 0),
    ('fecha_warn',  1, 10, AMBAR,    'FFF4DF', 168, 'center', 0, 1, 0),
    ('kpi',         1, 12, AZUL,     GRIS,  0,   'left',   0, 1, 0),
    ('kpinum',      1, 12, AZUL,     GRIS,  166, 'right',  0, 1, 0),
    ('kpimoney',    1, 12, AZUL,     GRIS,  164, 'right',  0, 1, 0),
    ('kpipct',      1, 12, AZUL,     GRIS,  167, 'right',  0, 1, 0),
    ('bold',        1, 10, '000000', None,  0,   'left',   0, 1, 0),
    ('boldnum',     1, 10, '000000', None,  166, 'right',  0, 1, 0),
    ('boldmoney',   1, 10, '000000', None,  164, 'right',  0, 1, 0),
    ('boldpct',     1, 10, '000000', None,  167, 'right',  0, 1, 0),
    ('nota',        0,  9, '595959', None,  0,   'left',   1, 0, 1),
    ('ok',          1, 10, VERDE,    None,  0,   'left',   0, 1, 0),
    ('warn',        1, 10, AMBAR,    None,  0,   'left',   0, 1, 0),
    ('bad',         1, 10, ROJO,     None,  0,   'left',   0, 1, 0),
    ('seccion',     1, 11, BLANCO,   AZUL2, 0,   'left',   0, 0, 0),
]
SMAP = {s[0]: i for i, s in enumerate(STYLES)}


def styles_xml():
    fonts, fills, borders, xfs = [], [], [], []
    # fills obligatorios
    fills.append('<fill><patternFill patternType="none"/></fill>')
    fills.append('<fill><patternFill patternType="gray125"/></fill>')
    fillmap = {None: 0}
    fontmap = {}
    for name, b, sz, col, fill, nf, al, wrap, bd, it in STYLES:
        fk = (b, sz, col, it)
        if fk not in fontmap:
            fontmap[fk] = len(fonts)
            fonts.append('<font>%s%s<sz val="%d"/><color rgb="FF%s"/><name val="Calibri"/><family val="2"/></font>'
                         % ('<b/>' if b else '', '<i/>' if it else '', sz, col))
        if fill not in fillmap:
            fillmap[fill] = len(fills)
            fills.append('<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/>'
                         '<bgColor indexed="64"/></patternFill></fill>' % fill)
    borders.append('<border><left/><right/><top/><bottom/><diagonal/></border>')
    borders.append('<border><left style="thin"><color rgb="FFBFBFBF"/></left>'
                   '<right style="thin"><color rgb="FFBFBFBF"/></right>'
                   '<top style="thin"><color rgb="FFBFBFBF"/></top>'
                   '<bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>')
    for name, b, sz, col, fill, nf, al, wrap, bd, it in STYLES:
        xfs.append('<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"'
                   ' applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1">'
                   '<alignment horizontal="%s" vertical="center"%s/></xf>'
                   % (nf, fontmap[(b, sz, col, it)], fillmap[fill], 1 if bd else 0, al,
                      ' wrapText="1"' if wrap else ''))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<numFmts count="%d">%s</numFmts>'
            '<fonts count="%d">%s</fonts><fills count="%d">%s</fills>'
            '<borders count="%d">%s</borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="%d">%s</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '<dxfs count="0"/>'
            '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
            '</styleSheet>'
            % (len(NUMFMTS), ''.join('<numFmt numFmtId="%d" formatCode="%s"/>' % n for n in NUMFMTS),
               len(fonts), ''.join(fonts), len(fills), ''.join(fills),
               len(borders), ''.join(borders), len(xfs), ''.join(xfs)))


EPOCH = datetime.date(1899, 12, 30)


def to_serial(d):
    if isinstance(d, datetime.datetime):
        d = d.date()
    return (d - EPOCH).days


class Cell:
    __slots__ = ('v', 'st', 'f')

    def __init__(self, v, st='base', f=None):
        self.v, self.st, self.f = v, st, f


def C(v, st='base'):
    return Cell(v, st)


def F(formula, st='base'):
    return Cell(None, st, formula)


class Chart:
    """Grafico nativo vinculado a rangos de la hoja (totalmente editable).

    cats    = lista de etiquetas de categoria (para la cache de visualizacion)
    series  = [(nombre, ref_valores, [valores])]
    """

    def __init__(self, kind, title, sheet, cat_ref, cats, series, anchor,
                 width=12, height=9, stacked=False, show_val=True, colors=None,
                 cat_title='', val_title='', numfmt='General'):
        self.kind, self.title, self.sheet = kind, title, sheet
        self.cat_ref, self.cats, self.series, self.anchor = cat_ref, cats, series, anchor
        self.w, self.h = width, height
        self.stacked, self.show_val = stacked, show_val
        self.colors = colors or SERIES
        self.cat_title, self.val_title = cat_title, val_title
        self.numfmt = numfmt


class Sheet:
    def __init__(self, name, cols=None, freeze=None, autofilter=None, tab_color=None):
        self.name = name
        self.rows = []            # lista de listas de Cell/valores
        self.cols = cols or []    # anchos
        self.freeze = freeze      # 'A3'
        self.autofilter = autofilter  # 'A5:M200'
        self.merges = []
        self.charts = []
        self.tables = []          # (nombre, ref, [encabezados])
        self.tab_color = tab_color

    def row(self, *cells):
        self.rows.append(list(cells))
        return len(self.rows)

    def blank(self, n=1):
        for _ in range(n):
            self.rows.append([])

    def merge(self, ref):
        self.merges.append(ref)

    def addrows(self, rows):
        for r in rows:
            self.rows.append(list(r))


def _cell_xml(ci, ri, c):
    ref = '%s%d' % (colname(ci), ri)
    if isinstance(c, Cell):
        v, st, f = c.v, c.st, c.f
    else:
        v, st, f = c, 'base', None
    si = SMAP.get(st, 0)
    if f is not None:
        return '<c r="%s" s="%d"><f>%s</f></c>' % (ref, si, esc(f))
    if v is None or v == '':
        return '<c r="%s" s="%d"/>' % (ref, si)
    if isinstance(v, bool):
        return '<c r="%s" s="%d" t="b"><v>%d</v></c>' % (ref, si, 1 if v else 0)
    if isinstance(v, (datetime.date, datetime.datetime)):
        return '<c r="%s" s="%d"><v>%d</v></c>' % (ref, si, to_serial(v))
    if isinstance(v, (int, float)):
        return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, si, repr(v) if isinstance(v, float) else v)
    return ('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, si, esc(v)))


def sheet_xml(sh, has_drawing, table_ids):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
           ' xmlns:r="%s">' % R_OFFICE]
    if sh.tab_color:
        out.append('<sheetPr><tabColor rgb="FF%s"/></sheetPr>' % sh.tab_color)
    out.append('<sheetViews><sheetView workbookViewId="0" showGridLines="0">')
    if sh.freeze:
        col = ''.join(ch for ch in sh.freeze if ch.isalpha())
        rw = ''.join(ch for ch in sh.freeze if ch.isdigit())
        xs = 0
        for ch in col:
            xs = xs * 26 + (ord(ch) - 64)
        out.append('<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
                   % (xs, int(rw) - 1, sh.freeze))
    out.append('</sheetView></sheetViews>')
    out.append('<sheetFormatPr defaultRowHeight="15"/>')
    if sh.cols:
        out.append('<cols>')
        for i, w in enumerate(sh.cols):
            out.append('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, w))
        out.append('</cols>')
    out.append('<sheetData>')
    for ri, row in enumerate(sh.rows, start=1):
        if not row:
            continue
        cells = ''.join(_cell_xml(ci, ri, c) for ci, c in enumerate(row) if c is not None)
        if cells:
            out.append('<row r="%d">%s</row>' % (ri, cells))
    out.append('</sheetData>')
    if sh.autofilter:
        out.append('<autoFilter ref="%s"/>' % sh.autofilter)
    if sh.merges:
        out.append('<mergeCells count="%d">%s</mergeCells>'
                   % (len(sh.merges), ''.join('<mergeCell ref="%s"/>' % m for m in sh.merges)))
    out.append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
    if has_drawing:
        out.append('<drawing r:id="rIdDraw"/>')
    if table_ids:
        out.append('<tableParts count="%d">%s</tableParts>'
                   % (len(table_ids), ''.join('<tablePart r:id="%s"/>' % t for t in table_ids)))
    out.append('</worksheet>')
    return ''.join(out)


# ------------------------------------------------------------------ graficos
def _strcache(vals):
    pts = ''.join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, esc(v)) for i, v in enumerate(vals))
    return '<c:strCache><c:ptCount val="%d"/>%s</c:strCache>' % (len(vals), pts)


def _numcache(vals, fmt='General'):
    pts = ''.join('' if v is None else '<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, v)
                  for i, v in enumerate(vals))
    return ('<c:numCache><c:formatCode>%s</c:formatCode><c:ptCount val="%d"/>%s</c:numCache>'
            % (esc(fmt), len(vals), pts))


def chart_xml(ch):
    kind = ch.kind
    plot = []
    if kind in ('bar', 'col'):
        plot.append('<c:barChart><c:barDir val="%s"/><c:grouping val="%s"/><c:varyColors val="0"/>'
                    % ('bar' if kind == 'bar' else 'col', 'stacked' if ch.stacked else 'clustered'))
    elif kind == 'line':
        plot.append('<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>')
    else:
        plot.append('<c:%sChart><c:varyColors val="1"/>' % ('doughnut' if kind == 'doughnut' else 'pie'))

    for i, (nm, ref, vals) in enumerate(ch.series):
        col = ch.colors[i % len(ch.colors)]
        sp = ('<c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
              '<a:ln><a:noFill/></a:ln></c:spPr>' % col)
        extra = ''
        if kind == 'line':
            sp = ('<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
                  '<a:round/></a:ln></c:spPr>' % col)
            extra = '<c:marker><c:symbol val="circle"/><c:size val="5"/>' \
                    '<c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill></c:spPr></c:marker>' % col
        dpts = ''
        if kind in ('pie', 'doughnut'):
            sp = ''
            dpts = ''.join('<c:dPt><c:idx val="%d"/><c:bubble3D val="0"/><c:spPr>'
                           '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
                           '<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>'
                           '</c:spPr></c:dPt>' % (j, ch.colors[j % len(ch.colors)])
                           for j in range(len(ch.cats)))
        dlbl = ''
        if ch.show_val:
            pos = ''
            if kind in ('bar', 'col'):
                pos = '<c:dLblPos val="%s"/>' % ('ctr' if ch.stacked else 'outEnd')
            elif kind == 'line':
                pos = '<c:dLblPos val="t"/>'
            dlbl = ('<c:dLbls><c:numFmt formatCode="%s" sourceLinked="0"/>'
                    '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
                    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1"/></a:pPr>'
                    '<a:endParaRPr lang="es-MX"/></a:p></c:txPr>%s'
                    '<c:showLegendKey val="0"/><c:showVal val="1"/>'
                    '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
                    '<c:showBubbleSize val="0"/></c:dLbls>' % (esc(ch.numfmt), pos))
        plot.append('<c:ser><c:idx val="%d"/><c:order val="%d"/>'
                    '<c:tx><c:v>%s</c:v></c:tx>%s%s%s%s'
                    '<c:cat><c:strRef><c:f>%s</c:f>%s</c:strRef></c:cat>'
                    '<c:val><c:numRef><c:f>%s</c:f>%s</c:numRef></c:val>%s</c:ser>'
                    % (i, i, esc(nm), sp, extra, dpts, dlbl,
                       escA(ch.cat_ref), _strcache(ch.cats),
                       escA(ref), _numcache(vals, ch.numfmt),
                       '<c:smooth val="0"/>' if kind == 'line' else ''))

    if kind in ('bar', 'col'):
        plot.append('<c:gapWidth val="60"/>')
        if ch.stacked:
            plot.append('<c:overlap val="100"/>')
        else:
            plot.append('<c:overlap val="-20"/>' if len(ch.series) > 1 else '<c:overlap val="0"/>')
        plot.append('<c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>')
    elif kind == 'line':
        plot.append('<c:marker val="1"/><c:axId val="111111111"/><c:axId val="222222222"/></c:lineChart>')
    else:
        if kind == 'doughnut':
            plot.append('<c:holeSize val="52"/>')
        else:
            plot.append('<c:firstSliceAng val="0"/>')
        plot.append('</c:%sChart>' % ('doughnut' if kind == 'doughnut' else 'pie'))

    axes = ''
    if kind in ('bar', 'col', 'line'):
        axes = (
            '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="%s"/>%s'
            '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>'
            '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
            '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr>'
            '<a:endParaRPr lang="es-MX"/></a:p></c:txPr>'
            '<c:crossAx val="222222222"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
            '<c:noMultiLvlLbl val="0"/></c:catAx>'
            '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="%s"/>'
            '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E7E7E7"/>'
            '</a:solidFill></a:ln></c:spPr></c:majorGridlines>%s'
            '<c:numFmt formatCode="%s" sourceLinked="0"/><c:majorTickMark val="none"/>'
            '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
            '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr>'
            '<a:endParaRPr lang="es-MX"/></a:p></c:txPr>'
            '<c:crossAx val="111111111"/></c:valAx>'
            % ('l' if kind == 'bar' else 'b', _axtitle(ch.cat_title),
               'b' if kind == 'bar' else 'l', _axtitle(ch.val_title), esc(ch.numfmt)))

    legend = ''
    if len(ch.series) > 1 or kind in ('pie', 'doughnut'):
        legend = ('<c:legend><c:legendPos val="b"/><c:overlay val="0"/>'
                  '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr>'
                  '<a:endParaRPr lang="es-MX"/></a:p></c:txPr></c:legend>')
    title = ('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
             '<a:defRPr sz="1200" b="1"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:defRPr></a:pPr>'
             '<a:r><a:rPr lang="es-MX" sz="1200" b="1"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
             '</a:rPr><a:t>%s</a:t></a:r></a:p></c:rich></c:tx>'
             '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
             % (AZUL, AZUL, esc(ch.title)))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:r="%s"><c:roundedCorners val="0"/><c:chart>%s'
            '<c:plotArea><c:layout/>%s%s<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
            '</c:plotArea>%s<c:plotVisOnly val="1"/>'
            '<c:dispBlanksAs val="gap"/></c:chart>'
            '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            '<a:ln><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr>'
            '<a:endParaRPr lang="es-MX"/></a:p></c:txPr>'
            '</c:chartSpace>' % (R_OFFICE, title, ''.join(plot), axes, legend))


def _axtitle(t):
    if not t:
        return ''
    return ('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
            '<a:defRPr sz="900" b="0"/></a:pPr><a:r><a:rPr lang="es-MX" sz="900"/>'
            '<a:t>%s</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' % esc(t))


EMU_COL = 640080     # ancho aproximado de columna por defecto
EMU_ROW = 190500


def drawing_xml(charts):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
           ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
           ' xmlns:r="%s">' % R_OFFICE]
    for i, ch in enumerate(charts):
        col = 0
        for c in ch.anchor:
            if c.isalpha():
                col = col * 26 + (ord(c.upper()) - 64)
        row = int(''.join(c for c in ch.anchor if c.isdigit()))
        out.append(
            '<xdr:twoCellAnchor editAs="oneCell">'
            '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
            '<xdr:cNvPr id="%d" name="Gráfico %d"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
            ' xmlns:r="%s" r:id="rId%d"/></a:graphicData></a:graphic></xdr:graphicFrame>'
            '<xdr:clientData/></xdr:twoCellAnchor>'
            % (col - 1, row - 1, col - 1 + ch.w, row - 1 + ch.h, i + 2, i + 1, R_OFFICE, i + 1))
    out.append('</xdr:wsDr>')
    return ''.join(out)


def table_xml(tid, name, ref, headers):
    cols = ''.join('<tableColumn id="%d" name="%s"/>' % (i + 1, escA(h)) for i, h in enumerate(headers))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' id="%d" name="%s" displayName="%s" ref="%s" totalsRowShown="0">'
            '<autoFilter ref="%s"/><tableColumns count="%d">%s</tableColumns>'
            '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0"'
            ' showRowStripes="1" showColumnStripes="0"/></table>'
            % (tid, escA(name), escA(name), ref, ref, len(headers), cols))


def write(path, sheets, title='Reporte'):
    pkg = Pkg()
    pkg.default('rels', 'application/vnd.openxmlformats-package.relationships+xml')
    pkg.default('xml', 'application/xml')
    pkg.add('_rels/.rels', rels([
        ('rId1', R_OFFICE + '/officeDocument', 'xl/workbook.xml'),
        ('rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
         'docProps/core.xml'),
        ('rId3', R_OFFICE + '/extended-properties', 'docProps/app.xml'),
    ]))
    pkg.add('docProps/core.xml', core_props(title),
            'application/vnd.openxmlformats-package.core-properties+xml')
    pkg.add('docProps/app.xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
            ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity>'
            '<ScaleCrop>false</ScaleCrop><Company>Kugelman Academy</Company>'
            '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
            '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>',
            'application/vnd.openxmlformats-officedocument.extended-properties+xml')

    wb_rels = []
    sheets_xml_tags = []
    tcount = 0
    for i, sh in enumerate(sheets, start=1):
        rid = 'rId%d' % i
        wb_rels.append((rid, R_OFFICE + '/worksheet', 'worksheets/sheet%d.xml' % i))
        sheets_xml_tags.append('<sheet name="%s" sheetId="%d" r:id="%s"/>' % (escA(sh.name), i, rid))
        srels = []
        table_ids = []
        for j, (tname, tref, theaders) in enumerate(sh.tables, start=1):
            tcount += 1
            trid = 'rIdTbl%d' % j
            pkg.add('xl/tables/table%d.xml' % tcount, table_xml(tcount, tname, tref, theaders),
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml')
            srels.append((trid, R_OFFICE + '/table', '../tables/table%d.xml' % tcount))
            table_ids.append(trid)
        if sh.charts:
            pkg.add('xl/drawings/drawing%d.xml' % i, drawing_xml(sh.charts),
                    'application/vnd.openxmlformats-officedocument.drawing+xml')
            drels = []
            for k, ch in enumerate(sh.charts, start=1):
                cidx = '%d_%d' % (i, k)
                pkg.add('xl/charts/chart%s.xml' % cidx, chart_xml(ch),
                        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
                drels.append(('rId%d' % k, R_OFFICE + '/chart', '../charts/chart%s.xml' % cidx))
            pkg.add('xl/drawings/_rels/drawing%d.xml.rels' % i, rels(drels))
            srels.append(('rIdDraw', R_OFFICE + '/drawing', '../drawings/drawing%d.xml' % i))
        if srels:
            pkg.add('xl/worksheets/_rels/sheet%d.xml.rels' % i, rels(srels))
        pkg.add('xl/worksheets/sheet%d.xml' % i, sheet_xml(sh, bool(sh.charts), table_ids),
                'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')

    n = len(sheets)
    wb_rels.append(('rId%d' % (n + 1), R_OFFICE + '/styles', 'styles.xml'))
    wb_rels.append(('rId%d' % (n + 2), R_OFFICE + '/theme', 'theme/theme1.xml'))
    pkg.add('xl/_rels/workbook.xml.rels', rels(wb_rels))
    pkg.add('xl/styles.xml', styles_xml(),
            'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml')
    pkg.add('xl/theme/theme1.xml', THEME, 'application/vnd.openxmlformats-officedocument.theme+xml')
    pkg.add('xl/workbook.xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="%s"><fileVersion appName="xl" lastEdited="7" lowestEdited="7"/>'
            '<workbookPr defaultThemeVersion="166925"/>'
            '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15000"/></bookViews>'
            '<sheets>%s</sheets><calcPr calcId="181029" fullCalcOnLoad="1"/></workbook>'
            % (R_OFFICE, ''.join(sheets_xml_tags)),
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
    pkg.save(path)
    return path


# tema minimo valido (requerido por Excel para resolver colores de tema)
def _theme():
    sc = ('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
          '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
          '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
          '<a:accent1><a:srgbClr val="1F3864"/></a:accent1><a:accent2><a:srgbClr val="2E8BC0"/></a:accent2>'
          '<a:accent3><a:srgbClr val="E0A030"/></a:accent3><a:accent4><a:srgbClr val="2E7D5B"/></a:accent4>'
          '<a:accent5><a:srgbClr val="B23A33"/></a:accent5><a:accent6><a:srgbClr val="7A5EA8"/></a:accent6>'
          '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>')
    fs = ('<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
          '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
          '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
          '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
          '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
          '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
          '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>'
          '<a:effectStyle><a:effectLst/></a:effectStyle>'
          '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
          '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
          '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
          '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>')
    font = ('<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
            '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Kugelman">'
            '<a:themeElements><a:clrScheme name="Kugelman">%s</a:clrScheme>'
            '<a:fontScheme name="Kugelman">%s</a:fontScheme>'
            '<a:fmtScheme name="Kugelman">%s</a:fmtScheme></a:themeElements>'
            '<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>' % (sc, font, fs))


THEME = _theme()
