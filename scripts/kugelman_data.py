# -*- coding: utf-8 -*-
"""Carga y normalizacion de las fuentes de Kugelman Academy (solo lectura)."""
import datetime
import glob
import re
import unicodedata

import xlsx_read as xr

CORTE = datetime.date(2026, 10, 3)

SRC_DIR = '../Reporte-Kugelman-Septiembre'
F_ALUMNOS = SRC_DIR + '/RGA_1791241975.xlsx'
F_CM = SRC_DIR + '/CLASES MUESTRA 2026__ (1).xlsx'
F_CAMP = [p for p in glob.glob(SRC_DIR + '/*.xlsx') if 'CAMPA' in p][0]

PARTICULAS = {'de', 'del', 'la', 'las', 'los', 'y', 'da', 'di'}
ABREV = {
    'hdz': 'hernandez', 'hdez': 'hernandez', 'hzz': 'hernandez',
    'gzz': 'gonzalez', 'mtz': 'martinez', 'rdz': 'rodriguez',
}
# Variantes ortograficas observadas en las fuentes (mismo apellido / nombre)
EQUIV = {
    'pedrosa': 'pedroza', 'cordova': 'cordoba', 'olguin': 'olguin',
    'gonzales': 'gonzalez', 'gonzalez': 'gonzalez', 'sanches': 'sanchez',
    'vazques': 'vazquez', 'nanacy': 'nancy', 'nikol': 'nicol',
    'dego': 'diego', 'alinne': 'aline', 'ayleen': 'aylen', 'ayelen': 'aylen',
    'garca': 'garcia', 'guzman': 'guzman', 'jacqueline': 'jacqueline',
    'jaqueline': 'jacqueline',
}


def strip_acc(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s)
                   if unicodedata.category(c) != 'Mn')


def norm(s):
    """Texto normalizado: sin acentos, minusculas, sin signos, espacios simples."""
    if s is None:
        return ''
    s = strip_acc(str(s)).lower()
    s = re.sub(r'[^a-z0-9 ]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def tokens(name):
    """Tokens significativos de un nombre (sin particulas, con abreviaturas resueltas)."""
    out = []
    for t in norm(name).split():
        t = ABREV.get(t, t)
        t = EQUIV.get(t, t)
        if t in PARTICULAS or len(t) < 2:
            continue
        out.append(t)
    return out


def norm_phone(v):
    """Ultimos 10 digitos del telefono; '' si no hay digitos utiles."""
    if v is None:
        return ''
    d = re.sub(r'\D', '', str(v))
    if len(d) < 10:
        return ''
    return d[-10:]


def clean_text(v):
    if v is None:
        return ''
    if isinstance(v, datetime.datetime):
        return v.date().isoformat()
    return re.sub(r'\s+', ' ', str(v)).strip()


def as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None


# ------------------------------------------------------------------ FUENTE 1
def load_alumnos():
    _, rows = xr.read_sheet(F_ALUMNOS)
    hdr = None
    for i, r in enumerate(rows):
        if r and clean_text(r[0]) == 'Matricula' or (r and clean_text(r[0]) == u'Matrícula'):
            hdr = i
            break
    out = []
    for r in rows[hdr + 1:]:
        r = list(r) + [None] * (7 - len(r))
        if r[0] in (None, ''):
            continue
        ap = clean_text(r[2])
        am = clean_text(r[3])
        nom = clean_text(r[1])
        apellidos = (ap + ' ' + am).strip()
        out.append({
            'matricula': r[0],
            'nombre': nom,
            'ap_pat': ap,
            'ap_mat': am,
            'apellidos': apellidos,
            'completo': (nom + ' ' + apellidos).strip(),
            'sexo': clean_text(r[4]),
            'edad': r[5] if isinstance(r[5], (int, float)) else None,
            'activo': clean_text(r[6]),
            'tok': set(tokens(nom + ' ' + apellidos)),
            'tok_nombre': tokens(nom),
            'tok_ap': tokens(apellidos),
        })
    return out


# ------------------------------------------------------------------ FUENTE 2
MESES = {'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
         'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10}


def load_clases_muestra():
    out = []
    for sname, _, _ in xr.all_sheets(F_CM):
        _, rows = xr.read_sheet(F_CM, sname)
        # localizar fila de encabezado
        hi = None
        for i, r in enumerate(rows[:6]):
            if r and any(clean_text(c).upper() == 'NOMBRE' for c in r if c):
                hi = i
                break
        if hi is None:
            continue
        hdr = [clean_text(c).upper() for c in rows[hi]]

        def ix(*names):
            for n in names:
                for j, h in enumerate(hdr):
                    if h == n or h.startswith(n):
                        return j
            return None

        i_f, i_n = ix('FECHA'), ix('NOMBRE')
        i_t = ix('NO. CONTACTO')
        i_c, i_e = ix('CATEGORIA', u'CATEGORÍA'), ix('EDAD')
        i_i, i_r = ix('IDIOMA'), ix('RESULTADO')
        i_m = ix('MEDIO CONTACTO', 'MEDIO')
        i_h = ix('HORA CLASE MUESTRA', 'HORARIO')
        extra_ix = [j for j, h in enumerate(hdr)
                    if h.startswith('COLUMNA') or h.startswith('COLUMN ')]

        for ri, r in enumerate(rows[hi + 1:], start=hi + 2):
            r = list(r) + [None] * (max(len(hdr), 25) - len(r))
            nombre = clean_text(r[i_n]) if i_n is not None else ''
            if not nombre:
                continue
            fecha = as_date(r[i_f]) if i_f is not None else None
            # Correccion de anios tecleados mal (2006 / 2025) conservando el mes de la hoja
            mes_hoja = MESES.get(sname.upper())
            fecha_orig = fecha
            corr = ''
            if fecha and mes_hoja and fecha.year != 2026:
                try:
                    fecha = datetime.date(2026, fecha.month, fecha.day)
                    corr = u'Año capturado como %d; corregido a 2026' % fecha_orig.year
                except ValueError:
                    pass
            if fecha and mes_hoja and fecha.month != mes_hoja:
                corr = (corr + '; ' if corr else '') + \
                    u'Mes capturado (%02d) no coincide con la hoja %s' % (fecha.month, sname)
            notas = []
            for j in extra_ix:
                v = clean_text(r[j]) if j < len(r) else ''
                if v:
                    notas.append(v)
            out.append({
                'mes': sname,
                'fila_origen': ri,
                'consec': clean_text(r[0]),
                'fecha': fecha,
                'fecha_corr': corr,
                'nombre': nombre,
                'telefono': clean_text(r[i_t]) if i_t is not None else '',
                'tel_n': norm_phone(r[i_t]) if i_t is not None else '',
                'categoria': clean_text(r[i_c]) if i_c is not None else '',
                'edad': r[i_e] if (i_e is not None and isinstance(r[i_e], (int, float))) else None,
                'edad_txt': clean_text(r[i_e]) if i_e is not None else '',
                'idioma': clean_text(r[i_i]) if i_i is not None else '',
                'resultado': clean_text(r[i_r]) if i_r is not None else '',
                'medio': clean_text(r[i_m]) if i_m is not None else '',
                'horario': clean_text(r[i_h]) if i_h is not None else '',
                'notas': ' | '.join(notas),
                'tok': tokens(nombre),
            })
    return out


# ----------------------------------------------- FUENTE 3 (apoyo): CAMPANAS
def load_campanas():
    out = []
    for sname, _, _ in xr.all_sheets(F_CAMP):
        _, rows = xr.read_sheet(F_CAMP, sname)
        hi = None
        for i, r in enumerate(rows[:6]):
            if r and any(clean_text(c).startswith('Nombre del interesado') for c in r if c):
                hi = i
                break
        if hi is None:
            continue
        hdr = [clean_text(c) for c in rows[hi]]

        def ix(pref):
            for j, h in enumerate(hdr):
                if h.startswith(pref):
                    return j
            return None

        i_f = ix('Fecha')
        i_n = ix('Nombre del interesado')
        i_t = ix(u'Teléfono')
        i_med = ix('Medio de contacto')
        i_anu = ix('Anuncio')
        i_seg = ix('Segmento')
        i_et = ix('Etapa del proceso')
        i_reg = ix('Registrado a clase muestra')
        i_asi = ix(u'Asistió a clase muestra')
        i_res = ix('Resultado del seguimiento')
        i_com = ix('Comentario')
        for r in rows[hi + 1:]:
            r = list(r) + [None] * (len(hdr) + 2 - len(r))
            nombre = clean_text(r[i_n]) if i_n is not None else ''
            tel = clean_text(r[i_t]) if i_t is not None else ''
            if not nombre and not tel:
                continue
            out.append({
                'campana': sname,
                'fecha': as_date(r[i_f]) if i_f is not None else None,
                'nombre': nombre,
                'telefono': tel,
                'tel_n': norm_phone(r[i_t]) if i_t is not None else '',
                'medio': clean_text(r[i_med]) if i_med is not None else '',
                'anuncio': clean_text(r[i_anu]) if i_anu is not None else '',
                'segmento': clean_text(r[i_seg]) if i_seg is not None else '',
                'etapa': clean_text(r[i_et]) if i_et is not None else '',
                'registrado': clean_text(r[i_reg]) if i_reg is not None else '',
                'asistio': clean_text(r[i_asi]) if i_asi is not None else '',
                'resultado': clean_text(r[i_res]) if i_res is not None else '',
                'comentario': clean_text(r[i_com]) if i_com is not None else '',
                'tok': tokens(nombre),
            })
    return out


# ------------------------------------------------------------------ matching
def match_score(cm_tok, al):
    """Puntua la coincidencia de un nombre de clase muestra contra un alumno.

    Devuelve (score, motivo). score >= 90 => coincidencia confiable.
    """
    if not cm_tok:
        return 0, ''
    ct, at = set(cm_tok), al['tok']
    inter = ct & at
    nom_set, ap_set = set(al['tok_nombre']), set(al['tok_ap'])
    hit_nom = ct & nom_set
    hit_ap = ct & ap_set
    if ct == at:
        return 100, 'Nombre completo identico (normalizado)'
    if ct <= at and len(hit_nom) >= 1 and len(hit_ap) >= 2:
        return 98, u'Nombre + 2 apellidos contenidos en el registro del alumno'
    if ct <= at and len(hit_nom) >= 1 and len(hit_ap) >= 1:
        return 95, u'Nombre + apellido contenidos en el registro del alumno'
    if at <= ct and len(hit_nom) >= 1 and len(hit_ap) >= 1:
        return 93, u'Registro de clase muestra mas extenso; contiene nombre y apellidos del alumno'
    if len(hit_nom) >= 1 and len(hit_ap) >= 2:
        return 92, u'Nombre de pila y dos apellidos coinciden'
    if len(hit_nom) >= 2 and len(hit_ap) >= 1:
        return 91, u'Dos nombres de pila y un apellido coinciden'
    if len(hit_nom) >= 1 and len(hit_ap) >= 1:
        return 70, u'Solo un nombre de pila y un apellido coinciden'
    if len(inter) >= 2:
        return 55, u'Dos tokens coinciden sin estructura clara'
    if len(inter) == 1 and len(ct) == 1:
        return 45, u'Unico token disponible coincide'
    return 0, ''
