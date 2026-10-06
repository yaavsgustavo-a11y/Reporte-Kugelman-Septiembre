"""Minimal XLSX writer (OOXML) using only the Python standard library.

Supports: multiple sheets, inline strings, numbers, dates, percentages,
named cell styles, column widths, frozen panes, merged cells and real
Excel tables (structured tables with autofilter).
"""
import datetime
import zipfile

EPOCH = datetime.datetime(1899, 12, 30)

# ---------------------------------------------------------------- styles ----
FONTS = [
    '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',          # 0 default
    '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',        # 1 bold
    '<font><b/><sz val="16"/><color rgb="FF1F3864"/><name val="Calibri"/><family val="2"/></font>',   # 2 title
    '<font><b/><sz val="12"/><color rgb="FF1F3864"/><name val="Calibri"/><family val="2"/></font>',   # 3 subtitle
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>',   # 4 header
    '<font><b/><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/><family val="2"/></font>',   # 5 green
    '<font><b/><sz val="11"/><color rgb="FF9C5700"/><name val="Calibri"/><family val="2"/></font>',   # 6 amber
    '<font><b/><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/><family val="2"/></font>',   # 7 red
    '<font><i/><sz val="10"/><color rgb="FF7F7F7F"/><name val="Calibri"/><family val="2"/></font>',   # 8 note
    '<font><b/><sz val="14"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',        # 9 big
]

FILLS = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>',  # 2 header
    '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>',  # 3 green
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/><bgColor indexed="64"/></patternFill></fill>',  # 4 amber
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>',  # 5 red
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2F3"/><bgColor indexed="64"/></patternFill></fill>',  # 6 soft blue
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>',  # 7 light gray
]

_THIN = ('<border><left style="thin"><color rgb="FFD9D9D9"/></left>'
         '<right style="thin"><color rgb="FFD9D9D9"/></right>'
         '<top style="thin"><color rgb="FFD9D9D9"/></top>'
         '<bottom style="thin"><color rgb="FFD9D9D9"/></bottom></border>')
BORDERS = [
    '<border><left/><right/><top/><bottom/><diagonal/></border>',  # 0 none
    _THIN,                                                         # 1 thin
]

NUMFMTS = {164: 'dd/mm/yyyy', 165: '0.0%', 166: '#,##0'}

# name -> (font, fill, border, numFmtId, horiz, wrap, vert)
STYLES = [
    ('default',   0, 0, 0, 0,   None,     False, None),
    ('title',     2, 0, 0, 0,   'left',   False, 'center'),
    ('subtitle',  3, 0, 0, 0,   'left',   False, 'center'),
    ('note',      8, 0, 0, 0,   'left',   True,  'top'),
    ('header',    4, 2, 1, 0,   'center', True,  'center'),
    ('text',      0, 0, 1, 0,   'left',   True,  'top'),
    ('center',    0, 0, 1, 0,   'center', False, 'top'),
    ('num',       0, 0, 1, 166, 'center', False, 'top'),
    ('date',      0, 0, 1, 164, 'center', False, 'top'),
    ('pct',       0, 0, 1, 165, 'center', False, 'top'),
    ('green',     5, 3, 1, 0,   'left',   True,  'center'),
    ('amber',     6, 4, 1, 0,   'left',   True,  'center'),
    ('red',       7, 5, 1, 0,   'left',   True,  'center'),
    ('label',     1, 7, 1, 0,   'left',   True,  'center'),
    ('metric',    9, 0, 1, 166, 'center', False, 'center'),
    ('metricpct', 9, 0, 1, 165, 'center', False, 'center'),
    ('section',   4, 2, 1, 0,   'left',   False, 'center'),
    ('bold',      1, 0, 1, 0,   'left',   True,  'top'),
    ('boldblue',  3, 6, 1, 0,   'left',   True,  'center'),
]
SIDX = {s[0]: i for i, s in enumerate(STYLES)}


def _styles_xml():
    nf = ''.join('<numFmt numFmtId="%d" formatCode="%s"/>' % (k, v) for k, v in sorted(NUMFMTS.items()))
    xfs = []
    for _name, f, fl, b, nfid, hor, wrap, vert in STYLES:
        al = ''
        parts = []
        if hor:
            parts.append('horizontal="%s"' % hor)
        if vert:
            parts.append('vertical="%s"' % vert)
        if wrap:
            parts.append('wrapText="1"')
        if parts:
            al = '<alignment %s/>' % ' '.join(parts)
        xfs.append('<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"'
                   ' applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"'
                   ' applyAlignment="1">%s</xf>' % (nfid, f, fl, b, al))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<numFmts count="%d">%s</numFmts>'
        '<fonts count="%d">%s</fonts>'
        '<fills count="%d">%s</fills>'
        '<borders count="%d">%s</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="%d">%s</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '<dxfs count="0"/>'
        '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
        '</styleSheet>' % (len(NUMFMTS), nf, len(FONTS), ''.join(FONTS), len(FILLS), ''.join(FILLS),
                           len(BORDERS), ''.join(BORDERS), len(STYLES), ''.join(xfs))
    )


# ---------------------------------------------------------------- helpers ---
def col_letter(idx):
    s = ''
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def esc(t):
    t = str(t)
    out = []
    for ch in t:
        o = ord(ch)
        if ch == '&':
            out.append('&amp;')
        elif ch == '<':
            out.append('&lt;')
        elif ch == '>':
            out.append('&gt;')
        elif ch == '"':
            out.append('&quot;')
        elif o < 0x20 and ch not in '\t\n\r':
            continue  # control chars illegal in XML 1.0
        else:
            out.append(ch)
    return ''.join(out)


class Pct(float):
    """Marker type: write as number with percent format (0.25 -> 25.0%)."""


class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}        # (row, col) -> (value, style_name)
        self.widths = []
        self.freeze = None     # (row, col) zero-based top-left of scrollable area
        self.merges = []
        self.table = None      # (name, first_row, first_col, last_row, last_col, [colnames])
        self.row_heights = {}

    def write(self, row, col, value, style='default'):
        self.cells[(row, col)] = (value, style)

    def write_row(self, row, col, values, style='default'):
        for i, v in enumerate(values):
            st = style[i] if isinstance(style, (list, tuple)) else style
            self.write(row, col + i, v, st)

    def set_widths(self, widths):
        self.widths = widths

    def set_freeze(self, row, col=0):
        self.freeze = (row, col)

    def merge(self, r1, c1, r2, c2):
        self.merges.append('%s%d:%s%d' % (col_letter(c1), r1 + 1, col_letter(c2), r2 + 1))

    def add_table(self, name, first_row, first_col, last_row, last_col, colnames):
        self.table = (name, first_row, first_col, last_row, last_col, colnames)

    def _xml(self, rel_table_id=None):
        rows = {}
        for (r, c), v in self.cells.items():
            rows.setdefault(r, {})[c] = v
        maxr = max(rows) if rows else 0
        maxc = max((max(cs) for cs in rows.values()), default=0)
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
               ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
        out.append('<dimension ref="A1:%s%d"/>' % (col_letter(maxc), maxr + 1))
        pane = ''
        if self.freeze:
            fr, fc = self.freeze
            ref = '%s%d' % (col_letter(fc), fr + 1)
            if fr and fc:
                split = ' xSplit="%d" ySplit="%d" activePane="bottomRight"' % (fc, fr)
                act = 'bottomRight'
            elif fr:
                split = ' ySplit="%d" activePane="bottomLeft"' % fr
                act = 'bottomLeft'
            else:
                split = ' xSplit="%d" activePane="topRight"' % fc
                act = 'topRight'
            pane = ('<pane%s topLeftCell="%s" state="frozen"/>'
                    '<selection pane="%s" activeCell="%s" sqref="%s"/>' % (split, ref, act, ref, ref))
        out.append('<sheetViews><sheetView workbookViewId="0" showGridLines="0">%s</sheetView></sheetViews>' % pane)
        out.append('<sheetFormatPr defaultRowHeight="15"/>')
        if self.widths:
            cols = ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, w)
                           for i, w in enumerate(self.widths))
            out.append('<cols>%s</cols>' % cols)
        out.append('<sheetData>')
        for r in sorted(rows):
            h = ' ht="%s" customHeight="1"' % self.row_heights[r] if r in self.row_heights else ''
            out.append('<row r="%d"%s>' % (r + 1, h))
            for c in sorted(rows[r]):
                val, style = rows[r][c]
                ref = '%s%d' % (col_letter(c), r + 1)
                s = SIDX[style]
                if val is None or val == '':
                    out.append('<c r="%s" s="%d"/>' % (ref, s))
                elif isinstance(val, bool):
                    out.append('<c r="%s" s="%d" t="b"><v>%d</v></c>' % (ref, s, int(val)))
                elif isinstance(val, datetime.datetime):
                    out.append('<c r="%s" s="%d"><v>%s</v></c>' % (ref, s, (val - EPOCH).days))
                elif isinstance(val, datetime.date):
                    d = datetime.datetime(val.year, val.month, val.day)
                    out.append('<c r="%s" s="%d"><v>%s</v></c>' % (ref, s, (d - EPOCH).days))
                elif isinstance(val, (int, float)):
                    out.append('<c r="%s" s="%d"><v>%s</v></c>' % (ref, s, repr(float(val)) if isinstance(val, float) else val))
                else:
                    out.append('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                               % (ref, s, esc(val)))
            out.append('</row>')
        out.append('</sheetData>')
        if self.merges:
            out.append('<mergeCells count="%d">%s</mergeCells>'
                       % (len(self.merges), ''.join('<mergeCell ref="%s"/>' % m for m in self.merges)))
        if self.table and rel_table_id:
            out.append('<tableParts count="1"><tablePart r:id="%s"/></tableParts>' % rel_table_id)
        out.append('</worksheet>')
        return ''.join(out)


class Workbook:
    def __init__(self):
        self.sheets = []

    def add_sheet(self, name):
        sh = Sheet(name)
        self.sheets.append(sh)
        return sh

    def save(self, path):
        z = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)
        tables = []  # (table_file_index, sheet_index, table def)
        for si, sh in enumerate(self.sheets):
            if sh.table:
                tables.append((len(tables) + 1, si, sh.table))

        ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
              '<Default Extension="xml" ContentType="application/xml"/>'
              '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.'
              'spreadsheetml.sheet.main+xml"/>'
              '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.'
              'spreadsheetml.styles+xml"/>'
              '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.'
              'core-properties+xml"/>'
              '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.'
              'extended-properties+xml"/>']
        for i in range(len(self.sheets)):
            ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.'
                      'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1))
        for ti, _si, _t in tables:
            ct.append('<Override PartName="/xl/tables/table%d.xml" ContentType="application/vnd.'
                      'openxmlformats-officedocument.spreadsheetml.table+xml"/>' % ti)
        ct.append('</Types>')
        z.writestr('[Content_Types].xml', ''.join(ct))

        z.writestr('_rels/.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                   'relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/'
                   'relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                   '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                   'relationships/extended-properties" Target="docProps/app.xml"/>'
                   '</Relationships>')

        now = datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
        z.writestr('docProps/core.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/'
                   'core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/'
                   'dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                   '<dc:title>Kugelman Academy - Corte Comercial</dc:title>'
                   '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
                   '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
                   '</cp:coreProperties>' % (now, now))
        z.writestr('docProps/app.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
                   ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                   '<Application>Microsoft Excel</Application></Properties>')

        sheets_xml = ''.join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (esc(s.name), i + 1, i + 1)
                             for i, s in enumerate(self.sheets))
        z.writestr('xl/workbook.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                   ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                   '<sheets>%s</sheets></workbook>' % sheets_xml)

        rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for i in range(len(self.sheets)):
            rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                        'relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1))
        rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                    'relationships/styles" Target="styles.xml"/>' % (len(self.sheets) + 1))
        rels.append('</Relationships>')
        z.writestr('xl/_rels/workbook.xml.rels', ''.join(rels))
        z.writestr('xl/styles.xml', _styles_xml())

        tbl_by_sheet = {si: ti for ti, si, _t in tables}
        for i, sh in enumerate(self.sheets):
            rid = 'rId1' if i in tbl_by_sheet else None
            z.writestr('xl/worksheets/sheet%d.xml' % (i + 1), sh._xml(rid))
            if i in tbl_by_sheet:
                ti = tbl_by_sheet[i]
                z.writestr('xl/worksheets/_rels/sheet%d.xml.rels' % (i + 1),
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                           'relationships/table" Target="../tables/table%d.xml"/></Relationships>' % ti)

        for ti, _si, (tname, r1, c1, r2, c2, colnames) in tables:
            ref = '%s%d:%s%d' % (col_letter(c1), r1 + 1, col_letter(c2), r2 + 1)
            cols = ''.join('<tableColumn id="%d" name="%s"/>' % (j + 1, esc(n)) for j, n in enumerate(colnames))
            z.writestr('xl/tables/table%d.xml' % ti,
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="%d"'
                       ' name="%s" displayName="%s" ref="%s" totalsRowShown="0">'
                       '<autoFilter ref="%s"/><tableColumns count="%d">%s</tableColumns>'
                       '<tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0"'
                       ' showRowStripes="1" showColumnStripes="0"/></table>'
                       % (ti, esc(tname), esc(tname), ref, ref, len(colnames), cols))
        z.close()
