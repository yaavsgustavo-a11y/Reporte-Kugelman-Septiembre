"""Minimal read-only XLSX reader using only the Python standard library."""
import zipfile
import re
import datetime
from xml.etree import ElementTree as ET

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

EPOCH = datetime.datetime(1899, 12, 30)


def col_to_idx(ref):
    m = re.match(r'([A-Z]+)', ref)
    s = m.group(1)
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def shared_strings(z):
    try:
        data = z.read('xl/sharedStrings.xml')
    except KeyError:
        return []
    root = ET.fromstring(data)
    out = []
    for si in root.findall('m:si', NS):
        out.append(''.join(t.text or '' for t in si.iter('{%s}t' % NS['m'])))
    return out


def number_formats(z):
    """Return map styleIdx -> numFmt code string."""
    try:
        root = ET.fromstring(z.read('xl/styles.xml'))
    except KeyError:
        return {}
    builtin = {14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
               18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
               22: 'm/d/yy h:mm', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0'}
    custom = {}
    nfs = root.find('m:numFmts', NS)
    if nfs is not None:
        for nf in nfs.findall('m:numFmt', NS):
            custom[int(nf.get('numFmtId'))] = nf.get('formatCode')
    out = {}
    cxfs = root.find('m:cellXfs', NS)
    if cxfs is not None:
        for i, xf in enumerate(cxfs.findall('m:xf', NS)):
            fid = int(xf.get('numFmtId', '0'))
            out[i] = custom.get(fid, builtin.get(fid, ''))
    return out


def sheet_names(z):
    root = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    relmap = {rel.get('Id'): rel.get('Target') for rel in rels}
    out = []
    for sh in root.find('m:sheets', NS).findall('m:sheet', NS):
        rid = sh.get('{%s}id' % NS['r'])
        target = relmap.get(rid, '')
        if not target.startswith('/'):
            target = 'xl/' + target.lstrip('/')
        else:
            target = target.lstrip('/')
        out.append((sh.get('name'), target, sh.get('state', 'visible')))
    return out


def is_date_fmt(code):
    if not code:
        return False
    c = re.sub(r'\[[^\]]*\]', '', code)
    c = re.sub(r'"[^"]*"', '', c)
    return bool(re.search(r'[ymdhs]', c, re.I)) and not re.search(r'^[#0,.%\s]+$', c)


def read_sheet(path, sheet=None):
    z = zipfile.ZipFile(path)
    ss = shared_strings(z)
    fmts = number_formats(z)
    sheets = sheet_names(z)
    if sheet is None:
        name, target, _ = sheets[0]
    else:
        match = [s for s in sheets if s[0] == sheet]
        if not match:
            raise KeyError('%s not in %s' % (sheet, [s[0] for s in sheets]))
        name, target, _ = match[0]
    root = ET.fromstring(z.read(target))
    rows = []
    sd = root.find('m:sheetData', NS)
    for row in sd.findall('m:row', NS):
        cells = {}
        for c in row.findall('m:c', NS):
            ref = c.get('r')
            idx = col_to_idx(ref)
            t = c.get('t', 'n')
            s = c.get('s')
            v = c.find('m:v', NS)
            if t == 'inlineStr':
                isn = c.find('m:is', NS)
                val = ''.join(x.text or '' for x in isn.iter('{%s}t' % NS['m'])) if isn is not None else ''
            elif t == 's':
                val = ss[int(v.text)] if v is not None and v.text else ''
            elif t == 'str':
                val = v.text if v is not None else ''
            elif t == 'b':
                val = (v is not None and v.text == '1')
            elif v is None or v.text is None or v.text == '':
                val = None
            else:
                try:
                    num = float(v.text)
                except ValueError:
                    val = v.text
                else:
                    code = fmts.get(int(s)) if s is not None else ''
                    if is_date_fmt(code):
                        val = EPOCH + datetime.timedelta(days=num)
                    else:
                        val = int(num) if num == int(num) else num
            cells[idx] = val
        if cells:
            width = max(cells) + 1
            rows.append([cells.get(i) for i in range(width)])
        else:
            rows.append([])
    z.close()
    return name, rows


def all_sheets(path):
    z = zipfile.ZipFile(path)
    names = sheet_names(z)
    z.close()
    return names
