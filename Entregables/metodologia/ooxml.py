"""Utilidades comunes para escribir paquetes OOXML (xlsx / docx / pptx)
usando unicamente la biblioteca estandar."""
import zipfile, datetime, re

NOW = '2026-10-05T00:00:00Z'


def esc(s):
    if s is None:
        return ''
    s = str(s)
    s = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    s = s.replace('"', '&quot;')
    # elimina caracteres de control no validos en XML 1.0
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', s)
    return s


def escA(s):
    return esc(s).replace("'", '&apos;')


class Pkg:
    """Acumula partes y escribe el zip final."""

    def __init__(self):
        self.parts = {}            # ruta -> bytes/str
        self.overrides = []        # (PartName, ContentType)
        self.defaults = {}         # ext -> ContentType

    def add(self, path, data, content_type=None):
        if isinstance(data, str):
            data = data.encode('utf-8')
        self.parts[path] = data
        if content_type:
            self.overrides.append(('/' + path, content_type))

    def default(self, ext, ct):
        self.defaults[ext] = ct

    def content_types(self):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">']
        for ext, ct in sorted(self.defaults.items()):
            out.append('<Default Extension="%s" ContentType="%s"/>' % (ext, ct))
        for pn, ct in self.overrides:
            out.append('<Override PartName="%s" ContentType="%s"/>' % (pn, ct))
        out.append('</Types>')
        return ''.join(out)

    def save(self, path):
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', self.content_types())
            for p in sorted(self.parts):
                z.writestr(p, self.parts[p])


def rels(items):
    """items = [(id, type, target[, mode])]"""
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for it in items:
        rid, typ, tgt = it[0], it[1], it[2]
        mode = it[3] if len(it) > 3 else None
        out.append('<Relationship Id="%s" Type="%s" Target="%s"%s/>' %
                   (rid, typ, escA(tgt), ' TargetMode="%s"' % mode if mode else ''))
    out.append('</Relationships>')
    return ''.join(out)


def core_props(title, author='Kugelman Academy', subject=''):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
            ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
            ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title><dc:subject>%s</dc:subject><dc:creator>%s</dc:creator>'
            '<cp:lastModifiedBy>%s</cp:lastModifiedBy>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
            '</cp:coreProperties>' % (esc(title), esc(subject), esc(author), esc(author), NOW, NOW))


R_OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
R_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
