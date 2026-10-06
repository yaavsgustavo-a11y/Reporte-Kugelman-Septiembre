"""Escritor PPTX (stdlib): diapositivas 16:9 con formas, cuadros de texto,
tablas y graficos nativos editables (libro de datos embebido).
"""
from ooxml import Pkg, esc, escA, rels, core_props, R_OFFICE
import chartsx

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

EMU_CM = 360000
W_SLIDE, H_SLIDE = 12192000, 6858000      # 33.87 x 19.05 cm (16:9)

AZUL = '1F3864'
AZUL2 = '2E5C9A'
AZUL3 = 'DCE4F2'
AZULCL = 'EDF1F8'
GRIS = 'F2F2F2'
GRIS_T = '595959'
BLANCO = 'FFFFFF'
VERDE = '2E7D5B'
AMBAR = 'C98A1B'
ROJO = 'B23A33'
SERIES = ['1F3864', '2E8BC0', 'E0A030', '2E7D5B', 'B23A33', '7A5EA8', '6B7280', '46A3B0']


def _prst(g):
    if g == 'roundRect':
        return ('<a:prstGeom prst="roundRect"><a:avLst>'
                '<a:gd name="adj" fmla="val 6000"/></a:avLst></a:prstGeom>')
    return '<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>' % g


def cm(v):
    return int(v * EMU_CM)


class Slide:
    def __init__(self):
        self.shapes = []
        self.charts = []        # (GChart, (x,y,w,h))
        self._id = 1

    def nid(self):
        self._id += 1
        return self._id

    # ------------------------------------------------------------- formas
    def rect(self, x, y, w, h, fill=None, line=None, radius=False, lw=12700,
             shadow=False, geom=None):
        prst = _prst(geom or ('roundRect' if radius else 'rect'))
        f = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill) if fill else '<a:noFill/>'
        ln = ('<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>' % (lw, line)) \
            if line else '<a:ln><a:noFill/></a:ln>'
        self.shapes.append(
            '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="Forma %d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>%s%s%s</p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
            % (self.nid(), self._id, x, y, w, h, prst, f, ln))

    def text(self, x, y, w, h, runs, align='l', anchor='t', fill=None, line=None,
             radius=False, wrap=True, spc=0, geom=None, rot=None):
        """runs = [(texto, {size, b, i, color, spc})] o string"""
        if isinstance(runs, str):
            runs = [(runs, {})]
        paras = []
        cur = []
        for t, o in runs:
            if t == '\n':
                paras.append(cur)
                cur = []
            else:
                cur.append((t, o))
        paras.append(cur)
        body = []
        for pr in paras:
            if not pr:
                body.append('<a:p><a:pPr algn="%s"/><a:endParaRPr lang="es-MX"/></a:p>' % align)
                continue
            rs = []
            o0 = pr[0][1]
            for t, o in pr:
                bullet = o.get('bullet')
                rs.append('<a:r><a:rPr lang="es-MX" sz="%d" b="%d" i="%d" dirty="0">'
                          '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
                          '<a:latin typeface="%s"/></a:rPr><a:t>%s</a:t></a:r>'
                          % (int(o.get('size', 14) * 100), 1 if o.get('b') else 0,
                             1 if o.get('i') else 0, o.get('color', '262626'),
                             o.get('font', 'Calibri'), esc(t)))
            ppr = '<a:pPr algn="%s" marL="%d" indent="%d"%s>%s</a:pPr>' % (
                o0.get('align', align), o0.get('marL', 0), o0.get('indent', 0),
                ' lvl="%d"' % o0['lvl'] if o0.get('lvl') else '',
                '<a:lnSpc><a:spcPct val="%d"/></a:lnSpc><a:spcBef><a:spcPts val="%d"/></a:spcBef>'
                '%s' % (int(o0.get('line', 1.0) * 100000), int(o0.get('before', 0) * 100),
                        '<a:buChar char="▪"/>' if o0.get('bullet') else '<a:buNone/>'))
            body.append('<a:p>%s%s</a:p>' % (ppr, ''.join(rs)))
        f = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill) if fill else '<a:noFill/>'
        ln = ('<a:ln w="12700"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>' % line) \
            if line else '<a:ln><a:noFill/></a:ln>'
        prst = _prst(geom or ('roundRect' if radius else 'rect'))
        xfrm = '<a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>' % (
            ' rot="%d"' % rot if rot else '', x, y, w, h)
        self.shapes.append(
            '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="Texto %d"/>'
            '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr>%s%s%s%s</p:spPr>'
            '<p:txBody><a:bodyPr wrap="%s" lIns="%d" tIns="%d" rIns="%d" bIns="%d" anchor="%s">'
            '<a:normAutofit/></a:bodyPr><a:lstStyle/>%s</p:txBody></p:sp>'
            % (self.nid(), self._id, xfrm, prst, f, ln,
               'square' if wrap else 'none', cm(0.15), cm(0.08), cm(0.15), cm(0.08),
               anchor, ''.join(body)))

    def line(self, x, y, w, h, color=AZUL, lw=25400):
        self.shapes.append(
            '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="%d" name="Linea %d"/><p:cNvCxnSpPr/><p:nvPr/>'
            '</p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
            '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln></p:spPr>'
            '</p:cxnSp>' % (self.nid(), self._id, x, y, w, h, lw, color))

    # ------------------------------------------------------------- tabla
    def table(self, x, y, w, col_w, headers, rows, row_h=cm(0.78), hdr_h=None,
              font=10, hdr_font=10, hdr_fill=AZUL, aligns=None, zebra=True):
        n = len(headers)
        tot = sum(col_w)
        col_w = [int(c * w / tot) for c in col_w]
        aligns = aligns or ['l'] * n
        hdr_h = hdr_h or row_h
        grid = ''.join('<a:gridCol w="%d"/>' % c for c in col_w)
        trs = ['<a:tr h="%d">%s</a:tr>' % (hdr_h, ''.join(
            _tc(h, hdr_fill, BLANCO, hdr_font, 1, 'ctr') for h in headers))]
        for ri, row in enumerate(rows):
            fill = GRIS if (zebra and ri % 2 == 1) else BLANCO
            cells = []
            for j in range(n):
                v = row[j] if j < len(row) else ''
                o = {}
                if isinstance(v, tuple):
                    v, o = v
                cells.append(_tc(v, o.get('fill', fill), o.get('color', '262626'),
                                 o.get('size', font), o.get('b', 0), o.get('align', aligns[j])))
            trs.append('<a:tr h="%d">%s</a:tr>' % (row_h, ''.join(cells)))
        h = hdr_h + row_h * len(rows)
        self.shapes.append(
            '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="Tabla %d"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/>'
            '</p:nvGraphicFramePr><p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            '<a:tbl><a:tblPr firstRow="1" bandRow="0"/><a:tblGrid>%s</a:tblGrid>%s</a:tbl>'
            '</a:graphicData></a:graphic></p:graphicFrame>'
            % (self.nid(), self._id, x, y, w, h, grid, ''.join(trs)))
        return h

    # ------------------------------------------------------------- grafico
    def chart(self, gchart, x, y, w, h):
        self.charts.append((gchart, (x, y, w, h)))


def _tc(txt, fill, color, size, b, align):
    return ('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="%s"/>'
            '<a:r><a:rPr lang="es-MX" sz="%d" b="%d"><a:solidFill><a:srgbClr val="%s"/>'
            '</a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>%s</a:t></a:r></a:p></a:txBody>'
            '<a:tcPr marL="%d" marR="%d" marT="%d" marB="%d" anchor="ctr">'
            '<a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR>'
            '<a:lnT w="3175"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnT>'
            '<a:lnB w="3175"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnB>'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:tcPr></a:tc>'
            % (align, int(size * 100), 1 if b else 0, color, esc(txt),
               cm(0.12), cm(0.12), cm(0.04), cm(0.04), fill))


class Pres:
    def __init__(self, title, subject=''):
        self.title, self.subject = title, subject
        self.slides = []

    def slide(self):
        s = Slide()
        self.slides.append(s)
        return s

    def save(self, path):
        pkg = Pkg()
        pkg.default('rels', 'application/vnd.openxmlformats-package.relationships+xml')
        pkg.default('xml', 'application/xml')
        pkg.default('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        pkg.add('_rels/.rels', rels([
            ('rId1', R_OFFICE + '/officeDocument', 'ppt/presentation.xml'),
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
                '<Application>Microsoft Office PowerPoint</Application>'
                '<Slides>%d</Slides><Company>Kugelman Academy</Company><AppVersion>16.0000</AppVersion>'
                '</Properties>' % len(self.slides),
                'application/vnd.openxmlformats-officedocument.extended-properties+xml')

        pkg.add('ppt/theme/theme1.xml', THEME,
                'application/vnd.openxmlformats-officedocument.theme+xml')
        pkg.add('ppt/slideMasters/slideMaster1.xml', MASTER,
                'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml')
        pkg.add('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([
            ('rId1', R_OFFICE + '/slideLayout', '../slideLayouts/slideLayout1.xml'),
            ('rId2', R_OFFICE + '/theme', '../theme/theme1.xml'),
        ]))
        pkg.add('ppt/slideLayouts/slideLayout1.xml', LAYOUT,
                'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml')
        pkg.add('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels([
            ('rId1', R_OFFICE + '/slideMaster', '../slideMasters/slideMaster1.xml'),
        ]))

        nchart = 0
        sids = []
        for i, s in enumerate(self.slides, start=1):
            srels = [('rId1', R_OFFICE + '/slideLayout', '../slideLayouts/slideLayout1.xml')]
            frames = []
            for k, (gc, (x, y, w, h)) in enumerate(s.charts, start=1):
                nchart += 1
                chartsx.add_charts(pkg, 'ppt', [gc], start=nchart)
                rid = 'rIdC%d' % k
                srels.append((rid, R_OFFICE + '/chart', '../charts/chart%d.xml' % nchart))
                frames.append(
                    '<p:graphicFrame><p:nvGraphicFramePr>'
                    '<p:cNvPr id="%d" name="Gráfico %d"/>'
                    '<p:cNvGraphicFramePr><a:graphicFrameLocks/></p:cNvGraphicFramePr><p:nvPr/>'
                    '</p:nvGraphicFramePr><p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>'
                    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
                    'xmlns:r="%s" r:id="%s"/></a:graphicData></a:graphic></p:graphicFrame>'
                    % (900 + k, k, x, y, w, h, R_OFFICE, rid))
            pkg.add('ppt/slides/_rels/slide%d.xml.rels' % i, rels(srels))
            pkg.add('ppt/slides/slide%d.xml' % i,
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld><p:spTree>'
                    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
                    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
                    '%s%s</p:spTree></p:cSld>'
                    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
                    '</p:sld>' % (A, R_OFFICE, P, ''.join(s.shapes), ''.join(frames)),
                    'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
            sids.append(i)

        prels = [('rIdM', R_OFFICE + '/slideMaster', 'slideMasters/slideMaster1.xml')]
        for i in sids:
            prels.append(('rIdS%d' % i, R_OFFICE + '/slide', 'slides/slide%d.xml' % i))
        prels.append(('rIdT', R_OFFICE + '/theme', 'theme/theme1.xml'))
        pkg.add('ppt/_rels/presentation.xml.rels', rels(prels))
        pkg.add('ppt/presentation.xml',
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" saveSubsetFonts="1">'
                '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst>'
                '<p:sldIdLst>%s</p:sldIdLst>'
                '<p:sldSz cx="%d" cy="%d"/><p:notesSz cx="%d" cy="%d"/>'
                '<p:defaultTextStyle><a:defPPr><a:defRPr lang="es-MX"/></a:defPPr></p:defaultTextStyle>'
                '</p:presentation>'
                % (A, R_OFFICE, P,
                   ''.join('<p:sldId id="%d" r:id="rIdS%d"/>' % (255 + i, i) for i in sids),
                   W_SLIDE, H_SLIDE, H_SLIDE, W_SLIDE),
                'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml')
        pkg.save(path)
        return path


# ------------------------------------------------------------------ partes fijas
def _txstyles():
    lv = []
    for i in range(1, 10):
        lv.append('<a:lvl%dpPr marL="%d" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" '
                  'latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200">'
                  '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
                  '<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl%dpPr>'
                  % (i, (i - 1) * 342900, i))
    return ''.join(lv)


MASTER = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<p:sldMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld>'
          '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
          '<a:effectLst/></p:bgPr></p:bg>'
          '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
          '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
          '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
          '</p:spTree></p:cSld>'
          '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"'
          ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"'
          ' hlink="hlink" folHlink="folHlink"/>'
          '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
          '<p:txStyles><p:titleStyle>%s</p:titleStyle><p:bodyStyle>%s</p:bodyStyle>'
          '<p:otherStyle>%s</p:otherStyle></p:txStyles></p:sldMaster>'
          % (A, R_OFFICE, P, _txstyles(), _txstyles(), _txstyles()))

LAYOUT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<p:sldLayout xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" type="blank" preserve="1">'
          '<p:cSld name="En blanco"><p:spTree>'
          '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
          '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
          '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
          '</p:spTree></p:cSld>'
          '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>' % (A, R_OFFICE, P))

THEME = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<a:theme xmlns:a="%s" name="Kugelman"><a:themeElements>'
         '<a:clrScheme name="Kugelman">'
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
         '</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>' % A)
