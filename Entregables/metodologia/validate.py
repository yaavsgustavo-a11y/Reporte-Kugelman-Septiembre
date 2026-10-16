"""Validador de paquetes OOXML generados: XML bien formado, content-types y
relaciones completas (todo r:id referenciado debe existir en el .rels del part)."""
import zipfile, re, sys, os
import xml.etree.ElementTree as ET

NS_CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
NS_R = '{http://schemas.openxmlformats.org/package/2006/relationships}'
RID_ATTR = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
RID_EMBED = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'


def check(path, verbose=True):
    errs, warns = [], []
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    if '[Content_Types].xml' not in names:
        errs.append('falta [Content_Types].xml')
        return errs, warns
    ct = ET.fromstring(z.read('[Content_Types].xml'))
    defaults = {d.get('Extension').lower(): d.get('ContentType') for d in ct.findall(NS_CT + 'Default')}
    overrides = {o.get('PartName').lstrip('/'): o.get('ContentType') for o in ct.findall(NS_CT + 'Override')}

    # 1) todo part xml bien formado
    for n in sorted(names):
        if n.endswith('.xml') or n.endswith('.rels'):
            try:
                ET.fromstring(z.read(n))
            except Exception as e:
                errs.append('XML mal formado %s: %s' % (n, e))

    # 2) cada part debe tener content type
    for n in sorted(names):
        if n == '[Content_Types].xml':
            continue
        ext = n.rsplit('.', 1)[-1].lower()
        if n not in overrides and ext not in defaults:
            errs.append('part sin content-type: %s' % n)

    # 3) relaciones: targets existen y r:id referenciados estan declarados
    for n in sorted(names):
        if not n.endswith('.rels'):
            continue
        base = os.path.dirname(os.path.dirname(n))
        try:
            rl = ET.fromstring(z.read(n))
        except Exception:
            continue
        for r in rl.findall(NS_R + 'Relationship'):
            if r.get('TargetMode') == 'External':
                continue
            tgt = r.get('Target')
            p = os.path.normpath(os.path.join(base, tgt)).replace('\\', '/').lstrip('/')
            if p not in names:
                errs.append('%s -> target inexistente %s (%s)' % (n, tgt, p))

    for n in sorted(names):
        if not (n.endswith('.xml')) or n == '[Content_Types].xml':
            continue
        d, f = os.path.dirname(n), os.path.basename(n)
        relp = (d + '/_rels/' + f + '.rels') if d else '_rels/' + f + '.rels'
        declared = set()
        if relp in names:
            rl = ET.fromstring(z.read(relp))
            declared = set(r.get('Id') for r in rl.findall(NS_R + 'Relationship'))
        try:
            root = ET.fromstring(z.read(n))
        except Exception:
            continue
        used = set()
        for el in root.iter():
            for a in (RID_ATTR, RID_EMBED, 'r:id'):
                v = el.get(a)
                if v:
                    used.add(v)
        miss = used - declared
        if miss:
            errs.append('%s usa r:id no declarados: %s' % (n, sorted(miss)))

    if verbose:
        tag = 'OK ' if not errs else 'ERR'
        print('%s %-62s parts=%-3d  errores=%d' % (tag, os.path.basename(path), len(names), len(errs)))
        for e in errs:
            print('    ! ' + e)
        for w in warns:
            print('    ~ ' + w)
    return errs, warns


if __name__ == '__main__':
    bad = 0
    for p in sys.argv[1:]:
        e, w = check(p)
        bad += len(e)
    sys.exit(1 if bad else 0)
