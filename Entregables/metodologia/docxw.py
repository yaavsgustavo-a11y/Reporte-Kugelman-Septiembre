"""Escritor DOCX (stdlib): titulos, parrafos, vinetas, tablas con formato,
graficos nativos editables (con libro de datos embebido), saltos de pagina,
encabezado/pie y numeracion de paginas.
"""
from ooxml import Pkg, esc, escA, rels, core_props, R_OFFICE
import chartsx
from chartsx import GChart

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
AZUL = '1F3864'
AZUL2 = '2E5C9A'
AZUL3 = 'DCE4F2'
GRIS = 'F4F4F4'
GRIS2 = 'EDEDED'
VERDE = '2E7D5B'
AMBAR = 'B07A12'
ROJO = 'B23A33'

EMU_PER_CM = 360000


class Doc:
    def __init__(self, title, subject=''):
        self.title, self.subject = title, subject
        self.body = []
        self.charts = []

    # ----------------------------------------------------------- texto
    def h1(self, t):
        self.body.append('<w:p><w:pPr><w:pStyle w:val="Titulo1"/></w:pPr>%s</w:p>' % _runs(t, b=1))

    def h2(self, t):
        self.body.append('<w:p><w:pPr><w:pStyle w:val="Titulo2"/></w:pPr>%s</w:p>' % _runs(t, b=1))

    def h3(self, t):
        self.body.append('<w:p><w:pPr><w:pStyle w:val="Titulo3"/></w:pPr>%s</w:p>' % _runs(t, b=1))

    def p(self, t='', size=20, color='262626', b=0, i=0, space_after=120, align='both'):
        self.body.append('<w:p><w:pPr><w:jc w:val="%s"/><w:spacing w:after="%d" w:line="264" '
                         'w:lineRule="auto"/></w:pPr>%s</w:p>'
                         % (align, space_after, _runs(t, size=size, color=color, b=b, i=i)))

    def small(self, t, color='595959'):
        self.p(t, size=16, color=color, space_after=100)

    def bullet(self, t, level=0):
        self.body.append('<w:p><w:pPr><w:pStyle w:val="Vineta"/><w:numPr>'
                         '<w:ilvl w:val="%d"/><w:numId w:val="1"/></w:numPr>'
                         '<w:spacing w:after="60" w:line="264" w:lineRule="auto"/></w:pPr>%s</w:p>'
                         % (level, _runs(t, size=20)))

    def numbered(self, t):
        self.body.append('<w:p><w:pPr><w:pStyle w:val="Vineta"/><w:numPr>'
                         '<w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'
                         '<w:spacing w:after="60"/></w:pPr>%s</w:p>' % _runs(t, size=20))

    def kv(self, label, value):
        self.body.append('<w:p><w:pPr><w:spacing w:after="40"/></w:pPr>%s%s</w:p>'
                         % (_runs(label + ': ', b=1, size=20), _runs(value, size=20)))

    def callout(self, titulo, texto, color=AZUL3, barra=AZUL):
        """Caja destacada de una celda."""
        cells = [[('%s' % titulo, {'b': 1, 'color': AZUL, 'size': 20})],
                 [(texto, {'size': 19, 'color': '262626'})]]
        rows = []
        for r in cells:
            rows.append('<w:tr><w:trPr/><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>'
                        '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>'
                        '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>'
                        '<w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar>'
                        '<w:tcBorders><w:left w:val="single" w:sz="18" w:color="%s"/></w:tcBorders>'
                        '</w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>%s</w:p></w:tc></w:tr>'
                        % (color, barra, ''.join(_runs(t, **o) for t, o in r)))
        self.body.append('<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>'
                         '<w:tblLayout w:type="fixed"/></w:tblPr>'
                         '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>%s</w:tbl>' % ''.join(rows))
        self.p('', space_after=0)

    def pagebreak(self):
        self.body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    def spacer(self, n=1):
        for _ in range(n):
            self.body.append('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>')

    # ----------------------------------------------------------- tablas
    def table(self, headers, rows, widths=None, aligns=None, header_fill=AZUL,
              zebra=True, font=18, total_row=False, caption=None, no_header=False):
        n = len(headers)
        widths = widths or [int(9360 / n)] * n
        total = sum(widths)
        if total != 9360:
            widths = [int(w * 9360 / total) for w in widths]
        aligns = aligns or ['left'] * n
        if caption:
            self.body.append('<w:p><w:pPr><w:spacing w:before="80" w:after="60"/></w:pPr>%s</w:p>'
                             % _runs(caption, b=1, size=18, color=AZUL2))
        grid = ''.join('<w:gridCol w:w="%d"/>' % w for w in widths)
        trs = []
        hc = []
        for j, h in enumerate(headers):
            hc.append(_tc(widths[j], _runs(h, b=1, size=font, color='FFFFFF'),
                          fill=header_fill, align='center'))
        if not no_header:
            trs.append('<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>%s</w:tr>' % ''.join(hc))
        for ri, row in enumerate(rows):
            fill = (GRIS if (zebra and ri % 2 == 1) else None)
            bold = 0
            if total_row and ri == len(rows) - 1:
                fill = AZUL3
                bold = 1
            tcs = []
            for j in range(n):
                v = row[j] if j < len(row) else ''
                style = {}
                if isinstance(v, tuple):
                    v, style = v
                tcs.append(_tc(widths[j], _runs(v, size=font, b=style.get('b', bold),
                                                color=style.get('color', '262626')),
                               fill=style.get('fill', fill), align=style.get('align', aligns[j])))
            trs.append('<w:tr>%s</w:tr>' % ''.join(tcs))
        self.body.append('<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>'
                         '<w:tblBorders>'
                         '<w:top w:val="single" w:sz="4" w:color="BFBFBF"/>'
                         '<w:left w:val="none" w:sz="0" w:color="auto"/>'
                         '<w:bottom w:val="single" w:sz="4" w:color="BFBFBF"/>'
                         '<w:right w:val="none" w:sz="0" w:color="auto"/>'
                         '<w:insideH w:val="single" w:sz="2" w:color="D9D9D9"/>'
                         '<w:insideV w:val="none" w:sz="0" w:color="auto"/></w:tblBorders>'
                         '<w:tblLayout w:type="fixed"/>'
                         '<w:tblCellMar><w:top w:w="50" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/>'
                         '<w:left w:w="90" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar>'
                         '</w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>' % (grid, ''.join(trs)))
        self.p('', space_after=60)

    def kpi_grid(self, items, cols=4):
        """Tarjetas de KPI: items = [(etiqueta, valor, nota)]"""
        w = int(9360 / cols)
        trs = []
        for i in range(0, len(items), cols):
            chunk = items[i:i + cols]
            tcs = []
            for lab, val, nota in chunk:
                inner = (_runs(lab.upper(), size=14, color='5A5A5A', b=1)
                         + '</w:p><w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="left"/></w:pPr>'
                         + _runs(val, size=30, color=AZUL, b=1))
                if nota:
                    inner += ('</w:p><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                              + _runs(nota, size=13, color='7A7A7A'))
                tcs.append(_tc(w, inner, fill=GRIS2, align='left', pad=120))
            for _ in range(cols - len(chunk)):
                tcs.append(_tc(w, '', fill=None))
            trs.append('<w:tr>%s</w:tr>' % ''.join(tcs))
        self.body.append('<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>'
                         '<w:tblBorders><w:insideV w:val="single" w:sz="12" w:color="FFFFFF"/>'
                         '<w:insideH w:val="single" w:sz="12" w:color="FFFFFF"/>'
                         '<w:top w:val="single" w:sz="12" w:color="FFFFFF"/>'
                         '<w:bottom w:val="single" w:sz="12" w:color="FFFFFF"/>'
                         '<w:left w:val="single" w:sz="12" w:color="FFFFFF"/>'
                         '<w:right w:val="single" w:sz="12" w:color="FFFFFF"/></w:tblBorders>'
                         '<w:tblLayout w:type="fixed"/></w:tblPr>'
                         '<w:tblGrid>%s</w:tblGrid>%s</w:tbl>'
                         % (''.join('<w:gridCol w:w="%d"/>' % w for _ in range(cols)), ''.join(trs)))
        self.p('', space_after=120)

    def funnel(self, etapas):
        """Embudo visual con tabla: etapas = [(nombre, valor_txt, ancho_rel 0-1, nota)]"""
        trs = []
        for nm, val, rel, nota in etapas:
            bar = max(1, int(6200 * rel))
            resto = 6200 - bar
            tcs = [_tc(2000, _runs(nm, b=1, size=17, color=AZUL), fill=None, align='left'),
                   _tc(bar, _runs(val, b=1, size=18, color='FFFFFF'), fill=AZUL2, align='center')]
            if resto > 40:
                tcs.append(_tc(resto, '', fill='F0F2F6'))
            tcs.append(_tc(1160, _runs(nota, size=15, color='595959'), fill=None, align='left'))
            trs.append('<w:tr><w:trPr><w:trHeight w:val="300"/></w:trPr>%s</w:tr>' % ''.join(tcs))
        self.body.append('<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>'
                         '<w:tblBorders><w:insideH w:val="single" w:sz="12" w:color="FFFFFF"/>'
                         '<w:insideV w:val="single" w:sz="8" w:color="FFFFFF"/></w:tblBorders>'
                         '<w:tblLayout w:type="fixed"/></w:tblPr>'
                         '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="6200"/>'
                         '<w:gridCol w:w="1160"/></w:tblGrid>%s</w:tbl>' % ''.join(trs))
        self.p('', space_after=120)

    # ----------------------------------------------------------- graficos
    def chart(self, gchart, width_cm=16.0, height_cm=8.4):
        self.charts.append(gchart)
        idx = len(self.charts)
        cx, cy = int(width_cm * EMU_PER_CM), int(height_cm * EMU_PER_CM)
        self.body.append(
            '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="60" w:after="120"/></w:pPr>'
            '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" '
            'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
            '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:docPr id="%d" name="Gráfico %d"/><wp:cNvGraphicFramePr/>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
            'xmlns:r="%s" r:id="rIdChart%d"/></a:graphicData></a:graphic>'
            '</wp:inline></w:drawing></w:r></w:p>'
            % (cx, cy, 100 + idx, idx, R_OFFICE, idx))

    # ----------------------------------------------------------- guardar
    def save(self, path):
        pkg = Pkg()
        pkg.default('rels', 'application/vnd.openxmlformats-package.relationships+xml')
        pkg.default('xml', 'application/xml')
        pkg.default('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        pkg.add('_rels/.rels', rels([
            ('rId1', R_OFFICE + '/officeDocument', 'word/document.xml'),
            ('rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
             'docProps/core.xml'),
            ('rId3', R_OFFICE + '/extended-properties', 'docProps/app.xml'),
        ]))
        pkg.add('docProps/core.xml', core_props(self.title, subject=self.subject),
                'application/vnd.openxmlformats-package.core-properties+xml')
        pkg.add('docProps/app.xml',
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
                ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                '<Application>Microsoft Office Word</Application><Company>Kugelman Academy</Company>'
                '<AppVersion>16.0000</AppVersion></Properties>',
                'application/vnd.openxmlformats-officedocument.extended-properties+xml')

        drels = [('rIdStyles', R_OFFICE + '/styles', 'styles.xml'),
                 ('rIdNum', R_OFFICE + '/numbering', 'numbering.xml'),
                 ('rIdSettings', R_OFFICE + '/settings', 'settings.xml'),
                 ('rIdTheme', R_OFFICE + '/theme', 'theme/theme1.xml'),
                 ('rIdFooter', R_OFFICE + '/footer', 'footer1.xml'),
                 ('rIdHeader', R_OFFICE + '/header', 'header1.xml')]
        if self.charts:
            chartsx.add_charts(pkg, 'word', self.charts)
            for i in range(1, len(self.charts) + 1):
                drels.append(('rIdChart%d' % i, R_OFFICE + '/chart', 'charts/chart%d.xml' % i))
        pkg.add('word/_rels/document.xml.rels', rels(drels))
        pkg.add('word/styles.xml', STYLES,
                'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml')
        pkg.add('word/numbering.xml', NUMBERING,
                'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml')
        pkg.add('word/settings.xml', SETTINGS,
                'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml')
        pkg.add('word/theme/theme1.xml', _THEME,
                'application/vnd.openxmlformats-officedocument.theme+xml')
        pkg.add('word/footer1.xml', FOOTER,
                'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
        pkg.add('word/header1.xml', HEADER % esc(self.title),
                'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
        sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>'
                '<w:footerReference w:type="default" r:id="rIdFooter"/>'
                '<w:pgSz w:w="12240" w:h="15840"/>'
                '<w:pgMar w:top="1300" w:right="1440" w:bottom="1200" w:left="1440"'
                ' w:header="708" w:footer="708" w:gutter="0"/>'
                '<w:titlePg/><w:docGrid w:linePitch="360"/></w:sectPr>')
        pkg.add('word/document.xml',
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="%s" xmlns:r="%s" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
                '<w:body>%s%s</w:body></w:document>'
                % (W, R_OFFICE, ''.join(self.body), sect),
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
        pkg.save(path)
        return path


def _runs(text, size=20, color='262626', b=0, i=0, font=None):
    if text is None:
        text = ''
    out = []
    for k, part in enumerate(str(text).split('\n')):
        if k:
            out.append('<w:r><w:br/></w:r>')
        rpr = '<w:rPr>%s%s<w:color w:val="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/>%s</w:rPr>' % (
            '<w:b/>' if b else '', '<w:i/>' if i else '', color, size, size,
            '<w:rFonts w:ascii="%s" w:hAnsi="%s"/>' % (font, font) if font else '')
        out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, esc(part)))
    return ''.join(out)


def _tc(w, inner_runs, fill=None, align='left', pad=90, valign='center'):
    shd = '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % fill if fill else ''
    return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s'
            '<w:tcMar><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/>'
            '<w:left w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/></w:tcMar>'
            '<w:vAlign w:val="%s"/></w:tcPr>'
            '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="%s"/></w:pPr>%s</w:p></w:tc>'
            % (w, shd, pad, pad, valign, align, inner_runs))


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="%s">'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="es-MX"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/>'
          '</w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:qFormat/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Titulo1"><w:name w:val="heading 1"/>'
          '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
          '<w:pPr><w:keepNext/><w:outlineLvl w:val="0"/>'
          '<w:pBdr><w:bottom w:val="single" w:sz="10" w:color="%s"/></w:pBdr>'
          '<w:spacing w:before="320" w:after="160"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/>'
          '<w:color w:val="%s"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Titulo2"><w:name w:val="heading 2"/>'
          '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
          '<w:pPr><w:keepNext/><w:outlineLvl w:val="1"/><w:spacing w:before="240" w:after="100"/></w:pPr>'
          '<w:rPr><w:b/><w:color w:val="%s"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Titulo3"><w:name w:val="heading 3"/>'
          '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
          '<w:pPr><w:keepNext/><w:outlineLvl w:val="2"/><w:spacing w:before="180" w:after="80"/></w:pPr>'
          '<w:rPr><w:b/><w:color w:val="404040"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Vineta"><w:name w:val="Lista con viñetas"/>'
          '<w:basedOn w:val="Normal"/><w:qFormat/>'
          '<w:pPr><w:ind w:left="360" w:hanging="220"/></w:pPr></w:style>'
          '</w:styles>' % (W, AZUL, AZUL, AZUL2))

NUMBERING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<w:numbering xmlns:w="%s">'
             '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
             '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="▪"/>'
             '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="220"/></w:pPr>'
             '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:hint="default"/>'
             '<w:color w:val="%s"/></w:rPr></w:lvl>'
             '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="–"/>'
             '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="220"/></w:pPr>'
             '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:hint="default"/></w:rPr></w:lvl>'
             '</w:abstractNum>'
             '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>'
             '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%%1."/>'
             '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="260"/></w:pPr>'
             '<w:rPr><w:b/><w:color w:val="%s"/></w:rPr></w:lvl></w:abstractNum>'
             '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
             '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
             '</w:numbering>' % (W, AZUL2, AZUL2))

SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="%s"><w:zoom w:percent="110"/>'
            '<w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
            '<w:themeFontLang w:val="es-MX"/></w:settings>' % W)

FOOTER = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:ftr xmlns:w="%s"><w:p><w:pPr><w:jc w:val="center"/>'
          '<w:pBdr><w:top w:val="single" w:sz="4" w:color="D9D9D9"/></w:pBdr>'
          '<w:spacing w:before="60" w:after="0"/></w:pPr>'
          '<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="15"/></w:rPr>'
          '<w:t xml:space="preserve">Kugelman Academy · Campaña Septiembre 2026 (21 ago – 4 oct) · Página </w:t></w:r>'
          '<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="15"/></w:rPr>'
          '<w:fldChar w:fldCharType="begin"/></w:r>'
          '<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="15"/></w:rPr>'
          '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
          '<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="15"/></w:rPr>'
          '<w:fldChar w:fldCharType="separate"/></w:r>'
          '<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="15"/></w:rPr><w:t>1</w:t></w:r>'
          '<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="15"/></w:rPr>'
          '<w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>' % W)

HEADER = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:hdr xmlns:w="' + W + '"><w:p><w:pPr><w:jc w:val="right"/>'
          '<w:spacing w:after="0"/></w:pPr>'
          '<w:r><w:rPr><w:color w:val="9A9A9A"/><w:sz w:val="15"/><w:caps/></w:rPr>'
          '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:hdr>')

_THEME = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Kugelman">'
          '<a:themeElements><a:clrScheme name="Kugelman">'
          '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
          '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
          '<a:dk2><a:srgbClr val="1F3864"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
          '<a:accent1><a:srgbClr val="1F3864"/></a:accent1><a:accent2><a:srgbClr val="2E8BC0"/></a:accent2>'
          '<a:accent3><a:srgbClr val="E0A030"/></a:accent3><a:accent4><a:srgbClr val="2E7D5B"/></a:accent4>'
          '<a:accent5><a:srgbClr val="B23A33"/></a:accent5><a:accent6><a:srgbClr val="7A5EA8"/></a:accent6>'
          '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
          '</a:clrScheme><a:fontScheme name="Kugelman">'
          '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
          '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
          '</a:fontScheme><a:fmtScheme name="Kugelman">'
          '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
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
          '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
          '</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')
