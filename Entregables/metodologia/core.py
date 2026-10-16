"""Core data loading, normalization and record-linkage for the
Kugelman Academy September 2026 campaign analysis.

Sources (all inside the repository):
  1. CAMPAÑAS (2).xlsx            -> hojas JULIO / AGOSTOSEPTIEMBRE / Octubre / OCTUBRE 1  (CRM de leads)
  2. CLASES MUESTRA 2026__ (1).xlsx -> hojas ENERO..OCTUBRE (bitacora de clases muestra)
  3. RGA_1791241975.xlsx          -> General de alumnos (padron, export 05/Oct/2026)
  4. Kugelman-Academy-Anuncios-21-ago-2026---4-oct-2026.csv  (Meta Ads, nivel anuncio)
"""
import csv, re, datetime, unicodedata, difflib, collections
import paths
from xlsxread import Xlsx

WIN_INI = datetime.date(2026, 8, 21)
WIN_FIN = datetime.date(2026, 10, 4)
CORTE_OPER = datetime.date(2026, 10, 3)
PRESUPUESTO = 7000.0

# ---------------------------------------------------------------- normalizers

def nstr(s):
    if s is None:
        return ''
    return re.sub(r'\s+', ' ', str(s)).strip()


def norm_name(s):
    """lowercase, sin acentos, sin signos, espacios colapsados"""
    s = nstr(s).lower()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def norm_phone(s):
    """ultimos 10 digitos; '' si no hay 10 digitos utilizables"""
    if s is None:
        return ''
    d = re.sub(r'\D', '', str(s))
    if len(d) > 10:
        d = d[-10:]
    return d if len(d) == 10 else ''


def name_sim(a, b):
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def tokens(n):
    return set(t for t in n.split() if len(t) > 2)


def name_match_level(a, b):
    """Nivel de coincidencia entre dos nombres normalizados.

    'exacta'  -> cadena identica
    'fuerte'  -> coincide el NOMBRE DE PILA y al menos un apellido
                 (o la similitud global es >= 0.85)
    'debil'   -> coincide solo el nombre de pila, o solo apellidos
    None      -> sin coincidencia
    Regla clave: exigir coincidencia de nombre de pila evita fusionar hermanos
    que comparten apellidos (p.ej. 'Mateo Garcia Vazquez' vs 'Heliana Garcia Vazquez').
    """
    if not a or not b:
        return None
    if a == b:
        return 'exacta'
    ta, tb = a.split(), b.split()
    if not ta or not tb:
        return None
    fa, fb = ta[0], tb[0]
    given_ok = (
        fa == fb
        or name_sim(fa, fb) >= 0.80
        or (len(fa) >= 4 and len(fb) >= 4 and (fa.startswith(fb) or fb.startswith(fa)))
        or fa in tb or fb in ta          # el nombre de pila aparece en el otro registro
    )
    sa, sb = set(ta[1:]), set(tb[1:])
    shared_sur = len(sa & sb)
    if given_ok:
        if shared_sur >= 1:
            return 'fuerte'
        if len(ta) == 1 or len(tb) == 1:
            return 'debil'            # solo nombre de pila disponible
        if name_sim(a, b) >= 0.85:
            return 'fuerte'
        return 'debil'
    if shared_sur >= 2:
        return 'debil'                # hermanos / mismo apellido, persona distinta
    return None


def same_person_names(a, b):
    """Coincidencia aceptable sin apoyo de telefono: exige nivel exacta/fuerte."""
    return name_match_level(a, b) in ('exacta', 'fuerte')


# ---------------------------------------------------------------- fix de fechas
YEAR_TYPOS = {}  # (hoja, fila) -> nota


def fix_year(dt, sheet, rownum):
    """Las hojas de clases muestra traen 3 fechas con anio 2006 (error de captura).
    Se corrige a 2026 y se documenta."""
    if isinstance(dt, datetime.datetime):
        d = dt.date()
    elif isinstance(dt, datetime.date):
        d = dt
    else:
        return None, False
    if d.year == 2006:
        YEAR_TYPOS[(sheet, rownum)] = d.isoformat()
        return d.replace(year=2026), True
    return d, False


# ---------------------------------------------------------------- 1. LEADS CRM
LEAD_COLS = ['fecha', 'nombre', 'tel', 'idioma', 'medio', 'anuncio', 'canal',
             'segmento', 'etapa', 'registrado_cm', 'asistio_cm', 'resultado', 'comentario']


def load_leads():
    x = Xlsx(paths.CAMPANAS)
    out = []
    # --- hoja principal de la campana
    t = x.table('AGOSTOSEPTIEMBRE')
    for i, r in enumerate(t):
        if i < 3 or not any(c not in (None, '') for c in r):
            continue
        rec = {k: (r[j] if j < len(r) else None) for j, k in enumerate(LEAD_COLS)}
        rec['_hoja'] = 'AGOSTOSEPTIEMBRE'
        rec['_fila'] = i + 1
        out.append(rec)
    # --- hoja "Octubre": copia de 5 registros del 30-sep (se marca, no se suma)
    t2 = x.table('Octubre')
    oct_rows = []
    for i, r in enumerate(t2):
        if i < 1 or not any(c not in (None, '') for c in r):
            continue
        rec = {k: (r[j] if j < len(r) else None) for j, k in enumerate(LEAD_COLS)}
        rec['_hoja'] = 'Octubre'
        rec['_fila'] = i + 1
        oct_rows.append(rec)
    # normaliza
    for rec in out + oct_rows:
        f = rec['fecha']
        rec['fecha'] = f.date() if isinstance(f, datetime.datetime) else f
        rec['nombre'] = nstr(rec['nombre'])
        rec['nn'] = norm_name(rec['nombre'])
        rec['tel10'] = norm_phone(rec['tel'])
        for k in ('idioma', 'medio', 'anuncio', 'canal', 'segmento', 'etapa',
                  'registrado_cm', 'asistio_cm', 'resultado'):
            rec[k] = nstr(rec[k])
    return out, oct_rows


def load_leads_julio():
    x = Xlsx(paths.CAMPANAS)
    t = x.table('JULIO')
    out = []
    for i, r in enumerate(t):
        if i < 2 or not any(c not in (None, '') for c in r):
            continue
        rec = {k: (r[j] if j < len(r) else None) for j, k in enumerate(LEAD_COLS)}
        f = rec['fecha']
        rec['fecha'] = f.date() if isinstance(f, datetime.datetime) else f
        rec['nombre'] = nstr(rec['nombre'])
        rec['nn'] = norm_name(rec['nombre'])
        rec['tel10'] = norm_phone(rec['tel'])
        rec['_hoja'] = 'JULIO'
        rec['_fila'] = i + 1
        out.append(rec)
    return out


# ------------------------------------------------------- 2. CLASES MUESTRA
CM_COLS = ['num', 'fecha', 'nombre', 'tel', 'categoria', 'edad', 'idioma',
           'resultado', 'medio_agenda', 'hora', 'nota1', 'nota2']


def load_cm(sheets=('AGOSTO', 'SEPTIEMBRE', 'OCTUBRE')):
    x = Xlsx(paths.CLASES)
    out = []
    for sh in sheets:
        t = x.table(sh)
        for i, r in enumerate(t):
            if i < 2 or not any(c not in (None, '') for c in r):
                continue
            rec = {k: (r[j] if j < len(r) else None) for j, k in enumerate(CM_COLS)}
            d, fixed = fix_year(rec['fecha'], sh, i + 1)
            rec['fecha'] = d
            rec['_anio_corregido'] = fixed
            rec['_hoja'] = sh
            rec['_fila'] = i + 1
            rec['nombre'] = nstr(rec['nombre'])
            rec['nn'] = norm_name(rec['nombre'])
            rec['tel10'] = norm_phone(rec['tel'])
            for k in ('categoria', 'idioma', 'resultado', 'medio_agenda', 'hora'):
                rec[k] = nstr(rec[k])
            rec['nota'] = ' / '.join([nstr(rec['nota1']), nstr(rec['nota2'])]).strip(' /')
            try:
                rec['edad'] = int(rec['edad']) if rec['edad'] not in (None, '') else None
            except (TypeError, ValueError):
                rec['edad'] = None
            out.append(rec)
    return out


# Homologacion del campo RESULTADO de la bitacora de clases muestra
CM_EST = {
    'inscrito': 'Realizada - Inscrito',
    'show': 'Realizada - Sin cierre',
    'no asistio': 'No show',
    'cita': 'Agendada pendiente',
    'cita pospuesta': 'Reprogramada',
}


def cm_estado(res):
    return CM_EST.get(norm_name(res), 'Estatus por validar')


# ------------------------------------------------------------- 3. ALUMNOS RGA

def load_alumnos():
    x = Xlsx(paths.RGA)
    t = x.table('General de alumnos')
    out = []
    for i, r in enumerate(t):
        if i < 5 or not any(c not in (None, '') for c in r):
            continue
        mat, nom, ap, am, sexo, edad, activo = (list(r) + [None] * 7)[:7]
        if mat in (None, ''):
            continue
        rec = {
            'matricula': mat, 'nombre': nstr(nom), 'ap': nstr(ap), 'am': nstr(am),
            'sexo': nstr(sexo), 'edad': edad, 'activo': nstr(activo), '_fila': i + 1,
        }
        rec['apellidos'] = nstr((rec['ap'] + ' ' + rec['am']).strip())
        rec['completo'] = nstr(rec['nombre'] + ' ' + rec['apellidos'])
        rec['nn'] = norm_name(rec['completo'])
        out.append(rec)
    meta = {'encabezado': nstr(t[0][0]), 'fecha_export': nstr(t[0][6]),
            'grupos': nstr(t[1][0])}
    return out, meta


# ------------------------------------------------------------------ 4. META ADS

def load_ads():
    rows = list(csv.DictReader(open(paths.ADS, encoding='utf-8')))
    out = []
    for r in rows:
        out.append({
            'anuncio': r['Nombre del anuncio'],
            'conjunto': r['Nombre del conjunto de anuncios'],
            'entrega': r['Entrega del anuncio'],
            'gasto': float(r['Importe gastado (MXN)']),
            'resultados': int(r['Resultados']),
            'impresiones': int(r['Impresiones']),
            'alcance': int(r['Alcance']),
            'costo_res': float(r['Costo por resultados']),
            'calidad': r['Clasificación de calidad'],
            'interaccion': r['Clasificación del porcentaje de interacción'],
            'conversion': r['Clasificación del porcentaje de conversiones'],
            'ini': r['Inicio del informe'], 'fin': r['Fin del informe'],
            'indicador': r['Indicador de resultado'],
        })
    return out


# Mapeo nombre de anuncio en Meta  <->  etiqueta usada en el CRM
AD_MAP = {
    'video villas': 'Reel Villas de Pachuca',
    'video sin libros': 'Reel sin libros',
    'post grupos': 'Post Grupos Abiertos',
    'video testimoniales': 'Reel Testimoniales',
    'video clases': 'Reel informativo',
    'carrete': 'Post Carrete',
    'post clase': 'Post Clase de Inglés Gratis',
    'video switch': 'Reel switch (sin registros en CRM)',
}


def in_window(d):
    return d is not None and WIN_INI <= d <= WIN_FIN


def name_match_strict(a, b):
    """Criterio estricto para vincular con el padron de alumnos (RGA).

    Exige nombre + al menos un apellido en ambos registros y, ademas,
    que un conjunto de tokens contenga al otro (p.ej. 'Mara Sanchez' dentro de
    'Mara Montserrat Sanchez Lopez') o una similitud textual >= 0.85.
    Evita falsos positivos del tipo 'Maria Luisa Hernandez' vs
    'Maria Luisa Rojas Bravo' (solo comparten nombre de pila).
    """
    if not a or not b:
        return None
    if a == b:
        return 'exacta'
    ta, tb = a.split(), b.split()
    if len(ta) < 2 or len(tb) < 2:
        return None
    sa, sb = set(ta), set(tb)
    fa, fb = ta[0], tb[0]
    given_ok = (fa == fb or name_sim(fa, fb) >= 0.85 or fa in tb or fb in ta)
    if given_ok and (sa <= sb or sb <= sa):
        return 'fuerte'
    if name_sim(a, b) >= 0.85:
        return 'fuerte'
    return None
