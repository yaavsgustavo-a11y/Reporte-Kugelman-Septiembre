"""Minimal XLSX reader using only the Python standard library.

Reads sheet names, cell values (inline/shared strings, numbers, booleans),
and converts Excel serial dates to date/datetime when the cell style says so.
"""
import zipfile, re, datetime
import xml.etree.ElementTree as ET

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKGREL = 'http://schemas.openxmlformats.org/package/2006/relationships'


def _q(tag, ns=NS_MAIN):
    return '{%s}%s' % (ns, tag)


# Number format ids that are built-in date/time formats
BUILTIN_DATE_FMTS = set(list(range(14, 23)) + list(range(45, 48)) + [27, 30, 36, 50, 57])


def col_to_idx(ref):
    """'BC12' -> 54 (0-based column index)"""
    m = re.match(r'([A-Z]+)', ref)
    letters = m.group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def idx_to_col(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def serial_to_datetime(v):
    """Excel 1900 date system."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return None
    if v <= 0:
        return None
    # Excel leap-year bug: serial 60 == 1900-02-29 (nonexistent)
    days = int(v)
    frac = v - days
    if days >= 61:
        days -= 1
    base = datetime.datetime(1899, 12, 31)
    try:
        dt = base + datetime.timedelta(days=days - 1, seconds=round(frac * 86400))
    except OverflowError:
        return None
    return dt


class Xlsx:
    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path)
        self.shared = self._read_shared()
        self.date_styles = self._read_styles()
        self.sheets = self._read_sheet_index()

    # ---------- package plumbing ----------
    def _read_shared(self):
        out = []
        try:
            data = self.z.read('xl/sharedStrings.xml')
        except KeyError:
            return out
        root = ET.fromstring(data)
        for si in root.findall(_q('si')):
            parts = [t.text or '' for t in si.iter(_q('t'))]
            out.append(''.join(parts))
        return out

    def _read_styles(self):
        """Return set of cellXf indexes whose numFmt is a date format."""
        date_xf = set()
        try:
            data = self.z.read('xl/styles.xml')
        except KeyError:
            return date_xf
        root = ET.fromstring(data)
        custom = {}
        numfmts = root.find(_q('numFmts'))
        if numfmts is not None:
            for nf in numfmts.findall(_q('numFmt')):
                fid = int(nf.get('numFmtId'))
                code = nf.get('formatCode') or ''
                custom[fid] = code
        cellxfs = root.find(_q('cellXfs'))
        if cellxfs is None:
            return date_xf
        for i, xf in enumerate(cellxfs.findall(_q('xf'))):
            fid = int(xf.get('numFmtId') or 0)
            if fid in BUILTIN_DATE_FMTS:
                date_xf.add(i)
            elif fid in custom:
                code = custom[fid]
                code_nq = re.sub(r'\[[^\]]*\]', '', re.sub(r'"[^"]*"', '', code))
                if re.search(r'[dmyhs]', code_nq, re.I) and not re.search(r'(General|0\.00)', code_nq):
                    date_xf.add(i)
        return date_xf

    def _read_sheet_index(self):
        wb = ET.fromstring(self.z.read('xl/workbook.xml'))
        rels = ET.fromstring(self.z.read('xl/_rels/workbook.xml.rels'))
        rmap = {r.get('Id'): r.get('Target') for r in rels.findall(_q('Relationship', NS_PKGREL))}
        out = []
        for sh in wb.find(_q('sheets')).findall(_q('sheet')):
            rid = sh.get(_q('id', NS_REL))
            target = rmap.get(rid, '')
            if target.startswith('/'):
                target = target[1:]
            elif not target.startswith('xl/'):
                target = 'xl/' + target
            out.append({'name': sh.get('name'), 'sheetId': sh.get('sheetId'),
                        'state': sh.get('state') or 'visible', 'path': target})
        return out

    @property
    def sheet_names(self):
        return [s['name'] for s in self.sheets]

    # ---------- data ----------
    def rows(self, sheet_name, keep_dates=True):
        """Yield list-of-values per row (ragged -> padded to max col seen in row)."""
        info = next(s for s in self.sheets if s['name'] == sheet_name)
        data = self.z.read(info['path'])
        root = ET.fromstring(data)
        sd = root.find(_q('sheetData'))
        if sd is None:
            return
        for row in sd.findall(_q('row')):
            cells = {}
            for c in row.findall(_q('c')):
                ref = c.get('r') or ''
                ci = col_to_idx(ref) if ref else len(cells)
                t = c.get('t')
                s_attr = c.get('s')
                v_el = c.find(_q('v'))
                is_el = c.find(_q('is'))
                val = None
                if t == 'inlineStr' and is_el is not None:
                    val = ''.join(x.text or '' for x in is_el.iter(_q('t')))
                elif t == 's' and v_el is not None:
                    try:
                        val = self.shared[int(v_el.text)]
                    except (ValueError, IndexError):
                        val = v_el.text
                elif t == 'b' and v_el is not None:
                    val = (v_el.text == '1')
                elif t == 'e':
                    val = '#ERR:' + (v_el.text if v_el is not None else '')
                elif t == 'str' and v_el is not None:
                    val = v_el.text
                elif v_el is not None:
                    txt = v_el.text
                    try:
                        num = float(txt)
                        if keep_dates and s_attr is not None and int(s_attr) in self.date_styles:
                            dt = serial_to_datetime(num)
                            val = dt if dt else num
                        else:
                            val = int(num) if num == int(num) and abs(num) < 1e15 else num
                    except (TypeError, ValueError):
                        val = txt
                cells[ci] = val
            if not cells:
                yield []
                continue
            maxc = max(cells)
            yield [cells.get(i) for i in range(maxc + 1)]

    def table(self, sheet_name):
        return list(self.rows(sheet_name))
