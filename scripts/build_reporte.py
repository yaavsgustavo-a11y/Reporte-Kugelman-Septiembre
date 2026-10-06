# -*- coding: utf-8 -*-
"""Construye el archivo ejecutivo de Kugelman Academy (corte sabado 03/10/2026).

Solo lectura sobre las fuentes originales; genera un .xlsx nuevo.
"""
import collections
import datetime

import kugelman_data as k
import xlsx_write as w

CORTE = datetime.date(2026, 10, 3)
SIN_INFO = u'Sin información'
SALIDA = u'../Reporte-Kugelman-Septiembre/Kugelman_Academy_Corte_Comercial_2026-10-03.xlsx'

# ---------------------------------------------------------------------------
# Overrides auditados manualmente (revision caso por caso)
# ---------------------------------------------------------------------------
CONFIRMA = {
    u'mairy alexandra': (1067, u'Ambos nombres de pila coinciden y son únicos en la base; categoría Kids y edad consistentes'),
    u'nanacy de la o': (1082, u'"Nanacy" = Nancy (error de captura). Teléfono 7712033713 compartido con María José Vargas De la O (misma familia). Edad 45 = 45'),
    u'victor manuel de la o': (1083, u'Nombres de pila completos + apellido materno "De la O"; misma familia y misma fecha que María José Vargas De la O'),
    u'luisa fernanda': (1089, u'Ambos nombres de pila coinciden y son únicos en la base; edad 9 = 9; categoría Kids'),
    u'laura lizeth pedrosa': (1130, u'"Pedrosa" = Pedroza (error de captura). El registro trae las edades 11, 14 y 36 (familia Márquez Pedroza)'),
    u'lianne': (1149, u'Nombre de pila único en la base. Teléfono 5517795341 compartido con Yara Guzmán Flores (misma familia)'),
}
RECHAZA = {
    u'daniela valeria hernandez pineda': u'No es Daniela Hernández Hernández: apellidos Hernández Pineda vs Hernández Hernández, Kids 11 años vs Certifications 26 años, teléfonos distintos',
    u'dulce abigahil hernandez lopez': u'No es Dulce Eli Ortega Hernández: apellidos Hernández López vs Ortega Hernández, 17 vs 12 años',
    u'ambar montserrat garcia sanchez': u'No es Mara Montserrat Sánchez López: Kids 9 años vs College 36 años. Son familiares (teléfono 5587932898 compartido)',
    u'zoe hernandez garcia': u'No es Mairy Alexandra García Hernández: nombre de pila distinto, 15 vs 13 años, teléfonos distintos',
    u'yatsiri alexandra sanchez hernandez': u'No es Mairy Alexandra García Hernández: apellidos Sánchez Hernández vs García Hernández',
    u'maria fernanda hernandez pineda': u'No es Fernanda Montserrat Hernández Vite: apellidos Hernández Pineda vs Hernández Vite, 7 vs 9 años',
    u'maria luisa hernandez': u'No es María Luisa Rojas Bravo: teléfono 5543564382 (compartido con Roxana Hernández Torres) vs 7711437295; 48 vs 47 años. Sólo coincide el nombre de pila',
    u'abigail sanchez hernandez': u'No es Jessica Sánchez Hernández: es su hermana (teléfono 7712444410 compartido), 23 vs 17 años, categoría College vs Junior',
    u'daniela': u'Registro con un solo nombre de pila; teléfono 7721029628 no coincide con ninguna alumna. No es posible confirmar identidad',
}
VALIDAR_EXTRA = {
    u'denisse e ortega': u'El registro dice "Inscrito" y comparte el apellido Ortega con Dulce Eli Ortega Hernández (matrículas 1014/1015), pero el nombre de pila y la edad (10 vs 12) no coinciden',
    u'alejandra mendez rebeca castelan': u'Una sola fila contiene dos personas. "Rebeca Castelán" podría ser Cecilia Rebeka Castelán Hernández (matrícula 1120); la ortografía Rebeca/Rebeka y la edad (14 vs 13) no permiten confirmarlo',
    u'cunada zuleyka': u'El campo nombre no contiene una identidad real ("Cuñada Zuleyka"): se requiere el nombre de la persona para darle seguimiento',
}
# Revision sugerida de captura en el Reporte General de Alumnos
REVISION_RGA = [
    (1087, u'Sexo registrado como "F" para el nombre "Santiago"'),
    (1115, u'Sexo registrado como "F" para el nombre "Roel Tomas"'),
    (1122, u'Sexo registrado como "F" para el nombre "Jose Ricardo"'),
    (1128, u'Sexo registrado como "F" para el nombre "Diego"'),
    (1140, u'Sexo registrado como "F"; la matrícula duplicada 1153 registra "M" para la misma persona'),
    (1092, u'Nombre capturado como "SSaid" (posible duplicación de la letra inicial)'),
    (1038, u'Apellido capturado como "Esptia" (posible "Espitia")'),
    (1110, u'Apellido materno vacío'),
    (1063, u'Apellido materno "Prueba": el registro parece ser una prueba del sistema, no un alumno real'),
]

ACTIVOS_CIERRE = {u'Interesado / por cerrar', u'Seguimiento post clase muestra', u'Pendiente de respuesta'}
ACTIVOS_CM = {u'Clase muestra por realizar', u'Reprogramación pendiente', u'Lista de espera',
              u'Clase muestra realizada'}
DESCARTADOS = {u'No interesado', u'Sin respuesta'}
ORDEN_ESTATUS = [u'Interesado / por cerrar', u'Seguimiento post clase muestra', u'Pendiente de respuesta',
                 u'Clase muestra por realizar', u'Reprogramación pendiente', u'Clase muestra realizada',
                 u'Lista de espera', u'No asistió', u'Información insuficiente', u'Sin respuesta',
                 u'No interesado']

ACCIONES = {
    u'Clase muestra por realizar': u'Confirmar asistencia 24 h antes (recordatorio por WhatsApp)',
    u'Clase muestra realizada': u'Registrar el resultado de la clase y agendar llamada de cierre',
    u'Seguimiento post clase muestra': u'Llamada de cierre: presentar propuesta de inscripción con fecha límite',
    u'Pendiente de respuesta': u'Segundo intento de contacto; si no responde en 7 días, reclasificar',
    u'Interesado / por cerrar': u'Cerrar inscripción: agendar firma y primer pago',
    u'Reprogramación pendiente': u'Reagendar clase muestra con fecha y hora confirmadas',
    u'No asistió': u'Reagendar clase muestra; si no responde en 7 días, marcar sin respuesta',
    u'Lista de espera': u'Avisar en cuanto abra grupo y horario compatible',
    u'No interesado': u'No contactar en el ciclo actual; conservar para campañas futuras',
    u'Sin respuesta': u'Último intento de contacto y cierre del registro',
    u'Información insuficiente': u'Validar y completar el registro (nombre completo, teléfono y resultado)',
}
SEMAFORO = {
    u'Interesado / por cerrar': 'green',
    u'Seguimiento post clase muestra': 'amber',
    u'Pendiente de respuesta': 'amber',
    u'Reprogramación pendiente': 'amber',
    u'Clase muestra por realizar': 'amber',
    u'Clase muestra realizada': 'amber',
    u'Lista de espera': 'amber',
    u'Información insuficiente': 'amber',
    u'No asistió': 'red',
    u'No interesado': 'red',
    u'Sin respuesta': 'red',
}
RES_COMERCIAL = {
    u'Inscrito': u'Se inscribió',
    u'Interesado / por cerrar': u'Prospecto activo',
    u'Clase muestra por realizar': u'Prospecto activo',
    u'Reprogramación pendiente': u'Prospecto activo',
    u'Lista de espera': u'Prospecto activo',
    u'Clase muestra realizada': u'Seguimiento',
    u'Seguimiento post clase muestra': u'Seguimiento',
    u'Pendiente de respuesta': u'Seguimiento',
    u'No asistió': u'No asistió',
    u'No interesado': u'No interesado',
    u'Sin respuesta': u'Sin respuesta',
    u'Información insuficiente': u'Sin información suficiente',
}


def nkey(name):
    return k.norm(name)


def fdate(d):
    return d.strftime('%d/%m/%Y') if d else ''


def tokset(lst):
    return set(lst)


# ===========================================================================
# 1. CARGA DE FUENTES
# ===========================================================================
alumnos = k.load_alumnos()
cms = k.load_clases_muestra()
camps = k.load_campanas()
al_by_mat = {a['matricula']: a for a in alumnos}

# ===========================================================================
# 2. CRUCE CLASES MUESTRA  <->  ALUMNOS
# ===========================================================================
for r in cms:
    key = nkey(r['nombre'])
    r['al'] = None
    r['match_why'] = u'Sin coincidencia en la base de alumnos'
    r['match_score'] = 0
    if key in CONFIRMA:
        mat, why = CONFIRMA[key]
        r['al'] = al_by_mat[mat]
        r['match_why'] = u'Validación manual – ' + why
        r['match_score'] = 100
        continue
    if key in RECHAZA:
        r['match_why'] = u'Coincidencia descartada tras revisión manual – ' + RECHAZA[key]
        continue
    best = (0, '', None)
    for a in alumnos:
        s, why = k.match_score(r['tok'], a)
        if s > best[0]:
            best = (s, why, a)
    s, why, a = best
    r['match_score'] = s
    if s >= 90:
        r['al'] = a
        r['match_why'] = why
    elif a and s >= 45:
        r['match_why'] = u'Coincidencia débil descartada (%s)' % a['completo']

# ===========================================================================
# 3. ENLACE CON CAMPANAS
#    - por nombre  -> sirve para inferir estatus
#    - por telefono -> solo contexto / fecha de primer contacto (familias
#      comparten numero, no se heredan resultados)
# ===========================================================================
cm_tels = set(r['tel_n'] for r in cms if r['tel_n'])
camp_by_tel = {}
for c in camps:
    c['link_nombre'] = False
    if c['tel_n']:
        camp_by_tel.setdefault(c['tel_n'], []).append(c)

for r in cms:
    por_nombre, por_tel = [], []
    rt = tokset(r['tok'])
    for c in camps:
        if not c['tok']:
            continue
        ct = tokset(c['tok'])
        if ct == rt or (len(ct) >= 2 and ct <= rt) or (len(rt) >= 2 and rt <= ct):
            por_nombre.append(c)
            c['link_nombre'] = True
    if r['tel_n']:
        for c in camp_by_tel.get(r['tel_n'], []):
            if c not in por_nombre:
                por_tel.append(c)
    r['camp_n'] = por_nombre
    r['camp_t'] = por_tel

# ===========================================================================
# 4. AGRUPACION POR PERSONA
# ===========================================================================
# matriculas duplicadas = misma persona capturada dos veces
sig_map = {}
for a in alumnos:
    sig_map.setdefault((nkey(a['completo']), a['edad']), []).append(a)
dup_pairs = [v for v in sig_map.values() if len(v) > 1]
mat_canon = {}
for grp in dup_pairs:
    canon = min(x['matricula'] for x in grp)
    for x in grp:
        mat_canon[x['matricula']] = canon

personas = {}
for r in cms:
    if r['al']:
        key = ('AL', mat_canon.get(r['al']['matricula'], r['al']['matricula']))
    else:
        key = ('NM', ' '.join(sorted(tokset(r['tok']))) or nkey(r['nombre']))
    p = personas.setdefault(key, {'regs': [], 'al': None, 'camp_n': [], 'camp_t': []})
    p['regs'].append(r)
    if r['al'] and not p['al']:
        p['al'] = al_by_mat[mat_canon.get(r['al']['matricula'], r['al']['matricula'])]
    for c in r['camp_n']:
        if c not in p['camp_n']:
            p['camp_n'].append(c)
    for c in r['camp_t']:
        if c not in p['camp_t']:
            p['camp_t'].append(c)

for p in personas.values():
    p['regs'].sort(key=lambda x: (x['fecha'] or datetime.date(1900, 1, 1)))
    p['solo_campana'] = False

# ===========================================================================
# 5. CLASIFICACION DE ESTATUS
# ===========================================================================
def camp_ultimo(lst, campo):
    vals = [(c['fecha'] or datetime.date(1900, 1, 1), c) for c in lst if c[campo]]
    vals.sort(key=lambda t: t[0])
    return vals[-1][1] if vals else None


def clasifica(p):
    regs = p['regs']
    res_list = [r['resultado'] for r in regs]
    notas = ' | '.join(r['notas'] for r in regs if r['notas'])
    notas_u = k.strip_acc(notas).upper()
    cr = camp_ultimo(p['camp_n'], 'resultado')
    cres = cr['resultado'] if cr else ''
    asistio = p['asistio']

    # (a0) clase muestra posterior al corte marcada como ya realizada
    post = [r for r in regs if r['fecha'] and r['fecha'] > CORTE and r['resultado'] in ('Show', 'Inscrito')]
    if post and not asistio:
        return (u'Clase muestra por realizar',
                u'Clase muestra agendada para el %s (posterior al corte del 03/10/2026). La fuente ya la '
                u'marca como "%s", dato que no puede darse por realizado a la fecha de corte'
                % (fdate(post[-1]['fecha']), post[-1]['resultado']))
    # (a) la fuente declara inscripcion pero no hay respaldo en la base de alumnos
    if any(x == 'Inscrito' for x in res_list) or cres == 'Inscrito':
        return (u'Información insuficiente',
                u'La fuente marca "Inscrito" pero la persona NO aparece en el Reporte General de Alumnos: '
                u'posible baja, inscripción no concretada o matrícula no localizada')
    # (b) evidencia explicita de perdida
    if u'NO LES GUSTO EL METODO' in notas_u:
        return u'No interesado', u'Nota en CLASES MUESTRA: "NO LES GUSTÓ EL MÉTODO"'
    if cres == u'No interesado':
        extra = u' – "%s"' % cr['comentario'] if cr['comentario'] else ''
        return u'No interesado', u'CAMPAÑAS, Resultado del seguimiento: "No interesado"' + extra
    if cres == u'Canceló proceso':
        return u'No interesado', u'CAMPAÑAS, Resultado del seguimiento: "Canceló proceso"'
    if cres == u'Sin respuesta':
        return u'Sin respuesta', u'CAMPAÑAS, Resultado del seguimiento: "Sin respuesta"'
    # (c) interes explicito de cierre
    if u'POSIBLE INSCRIPCION' in notas_u:
        return u'Interesado / por cerrar', u'Nota en CLASES MUESTRA: "%s"' % notas
    if cres == u'Inscripción pospuesta':
        return u'Interesado / por cerrar', u'CAMPAÑAS, Resultado del seguimiento: "Inscripción pospuesta"'
    # (d) clase muestra agendada despues del corte
    fut = [r for r in regs if r['fecha'] and r['fecha'] > CORTE and r['resultado'] in ('Cita', 'Cita pospuesta')]
    if fut:
        return (u'Clase muestra por realizar',
                u'Cita agendada para el %s (posterior al corte del 03/10/2026)' % fdate(fut[-1]['fecha']))
    if any(x == 'Cita pospuesta' for x in res_list):
        return u'Reprogramación pendiente', u'RESULTADO en CLASES MUESTRA: "Cita pospuesta"'
    # (e) asistio a clase muestra
    if asistio:
        if u'ESPERANDO RESPUESTA' in notas_u:
            return u'Pendiente de respuesta', u'Nota en CLASES MUESTRA: "esperando respuesta"'
        return (u'Seguimiento post clase muestra',
                u'Asistió a clase muestra (RESULTADO "Show") y no aparece en la base de alumnos')
    # (f) no asistio
    if any(x == u'NO ASISTIÓ' for x in res_list):
        if 'EN ESPERA' in notas_u or any(c['resultado'] == 'En espera' or c['asistio'] == 'En espera'
                                         for c in p['camp_n']):
            return u'Reprogramación pendiente', u'No asistió y el seguimiento sigue marcado "En espera"'
        return u'No asistió', u'RESULTADO en CLASES MUESTRA: "NO ASISTIÓ", sin evidencia posterior'
    # (g) cita vencida sin resultado
    vieja = [r for r in regs if r['resultado'] == 'Cita']
    if vieja:
        return (u'Información insuficiente',
                u'Cita agendada el %s sin resultado registrado en la fuente' % fdate(vieja[-1]['fecha']))
    # (h) lead
    if any(x == 'Lead' for x in res_list):
        return u'Pendiente de respuesta', u'RESULTADO en CLASES MUESTRA: "Lead" (sin clase muestra realizada)'
    return u'Información insuficiente', u'RESULTADO no registrado en la fuente'


def set_clase(p):
    e = p['estatus']
    if e == u'Inscrito':
        p['clase'] = u'Convertido'
    elif p.get('inscrito_sin_confirmar'):
        p['clase'] = u'Por validar'
    elif e in ACTIVOS_CIERRE:
        p['clase'] = u'Activo – en cierre'
    elif e in ACTIVOS_CM:
        p['clase'] = u'Activo – clase muestra pendiente'
    elif e == u'No asistió':
        p['clase'] = u'Recuperable – no asistió'
    elif e in DESCARTADOS:
        p['clase'] = u'Descartado'
    else:
        p['clase'] = u'Por validar'


for p in personas.values():
    p['inscrito'] = p['al'] is not None
    # "realizo clase muestra" se evalua a la fecha de corte: una clase agendada
    # despues del 03/10/2026 no puede contarse como realizada
    p['asistio'] = any(r['resultado'] in ('Show', 'Inscrito') and
                       (r['fecha'] is None or r['fecha'] <= CORTE) for r in p['regs']) or \
        any(c['asistio'] == 'Si' for c in p['camp_n'])
    p['inscrito_sin_confirmar'] = (not p['inscrito']) and (
        any(r['resultado'] == 'Inscrito' for r in p['regs']) or
        any(c['resultado'] == 'Inscrito' for c in p['camp_n']))
    if p['inscrito']:
        p['estatus'] = u'Inscrito'
        p['motivo'] = u'Confirmado en el Reporte General de Alumnos (matrícula %s)' % p['al']['matricula']
    else:
        p['estatus'], p['motivo'] = clasifica(p)
    set_clase(p)

# ===========================================================================
# 6. CAMPOS DERIVADOS
# ===========================================================================
PART = {'de', 'del', 'la', 'las', 'los', 'y'}


def split_nombre(p):
    if p['al']:
        return p['al']['nombre'], p['al']['apellidos'], p['al']['completo']
    raw = p['regs'][-1]['nombre'].strip().rstrip(',') if p['regs'] else p['completo']
    parts = [t for t in raw.split() if t]
    if len(parts) <= 1:
        return raw, '', raw
    idx = 1 if len(parts) == 2 else len(parts) - 2
    while idx > 0 and parts[idx - 1].lower() in PART:
        idx -= 1
    return ' '.join(parts[:idx]), ' '.join(parts[idx:]), raw


for p in personas.values():
    p['nombre'], p['apellidos'], p['completo'] = split_nombre(p)
    tels = [r['telefono'] for r in p['regs'] if r['telefono']] + \
           [c['telefono'] for c in p['camp_n'] if c['telefono']]
    p['telefono'] = tels[0] if tels else ''
    edades = [r['edad'] for r in p['regs'] if r['edad']]
    p['edad_cm'] = edades[-1] if edades else None
    p['edad'] = p['al']['edad'] if (p['al'] and p['al']['edad']) else p['edad_cm']
    p['sexo'] = p['al']['sexo'] if p['al'] else ''
    fc = [c['fecha'] for c in (p['camp_n'] + p['camp_t']) if c['fecha']]
    p['primer_contacto'] = min(fc) if fc else None
    cm_f = [r['fecha'] for r in p['regs'] if r['fecha']]
    p['cm_fecha'] = max(cm_f) if cm_f else None
    p['cm_fecha_min'] = min(cm_f) if cm_f else None
    p['horario'] = next((r['horario'] for r in reversed(p['regs']) if r['horario']), '')
    p['resultado_cm'] = (p['regs'][-1]['resultado'] if p['regs'] else '') or SIN_INFO
    p['categoria'] = next((r['categoria'] for r in reversed(p['regs']) if r['categoria']), '')
    p['idioma'] = next((r['idioma'] for r in reversed(p['regs']) if r['idioma']), '')
    p['medio'] = next((r['medio'] for r in reversed(p['regs']) if r['medio']), '')
    p['notas'] = ' | '.join(sorted(set(r['notas'] for r in p['regs'] if r['notas'])))
    p['n_cm'] = len(p['regs'])
    mov = []
    if p['regs']:
        r = p['regs'][-1]
        mov.append(u'Clase muestra %s – resultado "%s" (CLASES MUESTRA, hoja %s)'
                   % (fdate(r['fecha']) or SIN_INFO, r['resultado'] or SIN_INFO, r['mes']))
    ce = camp_ultimo(p['camp_n'], 'etapa')
    if ce:
        t = u'CAMPAÑAS %s – etapa "%s"' % (fdate(ce['fecha']), ce['etapa'])
        cr = camp_ultimo(p['camp_n'], 'resultado')
        if cr:
            t += u', resultado "%s"' % cr['resultado']
        mov.append(t)
    p['ultimo_mov'] = ' ; '.join(mov) or SIN_INFO

# ===========================================================================
# 7. LEADS DE CAMPANAS SIN REGISTRO DE CLASE MUESTRA
# ===========================================================================
camp_extra = {}
camp_desc = {'sin_nombre': 0, 'sin_avance': 0, 'ya_alumno': 0, 'mismo_telefono': 0}
for c in camps:
    if c['link_nombre']:
        continue
    if not c['nombre']:
        camp_desc['sin_nombre'] += 1
        continue
    avance = bool(c['registrado'] or c['asistio'] or c['resultado']) or \
        c['etapa'] in (u'Pendiente de decisión', u'Asistencia a Clase muestra')
    if not avance:
        camp_desc['sin_avance'] += 1
        continue
    if any(k.match_score(c['tok'], a)[0] >= 93 for a in alumnos):
        camp_desc['ya_alumno'] += 1
        continue
    if c['tel_n'] and c['tel_n'] in cm_tels:
        camp_desc['mismo_telefono'] += 1
        continue
    camp_extra.setdefault(' '.join(sorted(tokset(c['tok']))), []).append(c)

for key, lst in camp_extra.items():
    lst.sort(key=lambda c: c['fecha'] or datetime.date(1900, 1, 1))
    c = lst[-1]
    res = c['resultado']
    pocos_tokens = len(tokset(c['tok'])) < 2
    if res == u'Inscrito':
        est = u'Información insuficiente'
        mot = (u'CAMPAÑAS marca "Inscrito" pero la persona NO aparece en el Reporte General de Alumnos '
               u'y no tiene registro en CLASES MUESTRA')
    elif res in (u'No interesado', u'Canceló proceso'):
        est, mot = u'No interesado', u'CAMPAÑAS, Resultado del seguimiento: "%s"' % res
    elif res == u'Sin respuesta':
        est, mot = u'Sin respuesta', u'CAMPAÑAS, Resultado del seguimiento: "Sin respuesta"'
    elif res == u'Inscripción pospuesta':
        est, mot = u'Interesado / por cerrar', u'CAMPAÑAS, Resultado del seguimiento: "Inscripción pospuesta"'
    elif pocos_tokens:
        est = u'Información insuficiente'
        mot = (u'Lead de campaña identificado sólo con apodo o nombre parcial ("%s"): no es posible '
               u'cruzarlo contra alumnos ni contra clases muestra' % c['nombre'])
    elif c['asistio'] == 'Si':
        est, mot = (u'Seguimiento post clase muestra',
                    u'CAMPAÑAS indica asistencia a clase muestra sin registro en la base de CLASES MUESTRA')
    elif c['registrado'] == 'Si':
        est, mot = (u'Clase muestra por realizar',
                    u'CAMPAÑAS indica registro a clase muestra sin resultado capturado')
    elif c['etapa'] == u'Pendiente de decisión':
        est, mot = u'Pendiente de respuesta', u'CAMPAÑAS, Etapa del proceso: "Pendiente de decisión"'
    elif 'En espera' in (c['registrado'], c['asistio'], c['resultado']):
        est, mot = u'Pendiente de respuesta', u'CAMPAÑAS, seguimiento marcado "En espera"'
    else:
        est, mot = (u'Información insuficiente',
                    u'Lead de campaña sin resultado de seguimiento registrado')
    nombre = c['nombre'].strip()
    parts = nombre.split()
    p = {
        'regs': [], 'camp_n': lst, 'camp_t': [], 'al': None, 'inscrito': False,
        'estatus': est, 'motivo': mot, 'solo_campana': True,
        'nombre': parts[0] if parts else nombre,
        'apellidos': ' '.join(parts[1:]) if len(parts) > 1 else '',
        'completo': nombre,
        'telefono': next((x['telefono'] for x in lst if x['telefono']), ''),
        'edad': None, 'edad_cm': None, 'sexo': '',
        'primer_contacto': next((x['fecha'] for x in lst if x['fecha']), None),
        'cm_fecha': None, 'cm_fecha_min': None, 'horario': '',
        'resultado_cm': u'Sin registro en CLASES MUESTRA',
        'categoria': next((x['segmento'] for x in lst if x['segmento']), ''),
        'idioma': next((x['nombre'] and '' for x in lst[:1]), ''),
        'medio': next((x['medio'] for x in lst if x['medio']), ''),
        'notas': ' | '.join(sorted(set(x['comentario'] for x in lst if x['comentario']))),
        'n_cm': 0, 'asistio': any(x['asistio'] == 'Si' for x in lst),
        'inscrito_sin_confirmar': any(x['resultado'] == 'Inscrito' for x in lst),
        'ultimo_mov': u'CAMPAÑAS %s (hoja %s) – etapa "%s"; registrado a clase muestra "%s"; '
                      u'asistió "%s"; resultado "%s"' % (
                          fdate(c['fecha']) or SIN_INFO, c['campana'], c['etapa'] or SIN_INFO,
                          c['registrado'] or SIN_INFO, c['asistio'] or SIN_INFO, c['resultado'] or SIN_INFO),
    }
    set_clase(p)
    personas[('CP', key)] = p

prospectos = [p for p in personas.values() if not p['inscrito']]
convertidos = [p for p in personas.values() if p['inscrito']]

# ===========================================================================
# 8. ANTIGUEDAD DEL ULTIMO MOVIMIENTO
# ===========================================================================
for p in personas.values():
    fechas = [f for f in [p['cm_fecha'], p['primer_contacto']] if f]
    fechas += [c['fecha'] for c in p['camp_n'] if c['fecha']]
    p['fecha_mov'] = max(fechas) if fechas else None
    if p['fecha_mov'] and p['fecha_mov'] <= CORTE:
        p['dias'] = (CORTE - p['fecha_mov']).days
    elif p['fecha_mov']:
        p['dias'] = 0          # movimiento agendado a futuro
    else:
        p['dias'] = None

# ===========================================================================
# 9. REGISTROS POR VALIDAR
# ===========================================================================
val = []        # (tipo, persona, referencia, fuente, info, que_validar, prioridad)
GENERICO_MAX_TOK = 2


def add_val(tipo, persona, ref, fuente, info, que, prio):
    val.append([tipo, persona, ref, fuente, info, que, prio])


# (1) matriculas duplicadas
for grp in dup_pairs:
    mats = ', '.join(str(x['matricula']) for x in grp)
    sexos = set(x['sexo'] for x in grp)
    extra = u' Además el sexo difiere entre matrículas (%s).' % ' / '.join(sorted(sexos)) if len(sexos) > 1 else ''
    add_val(u'Matrícula duplicada', grp[0]['completo'], mats, u'FUENTE 1 – Reporte General de Alumnos',
            u'%d matrículas con nombre, edad y sexo idénticos (%s años).' % (len(grp), grp[0]['edad'] or 0) + extra,
            u'Confirmar si son dos personas distintas o un alta duplicada; cancelar la matrícula sobrante. '
            u'Mientras no se resuelva, el conteo de alumnos varía entre %d y %d.' % (len(alumnos) - len(dup_pairs), len(alumnos)),
            u'Alta')

# (2) revision de captura en RGA
for mat, det in REVISION_RGA:
    a = al_by_mat.get(mat)
    if not a:
        continue
    add_val(u'Posible error de captura (revisión sugerida)', a['completo'], str(mat),
            u'FUENTE 1 – Reporte General de Alumnos', det,
            u'Verificar el dato contra el expediente del alumno y corregir en el sistema.',
            u'Alta' if mat == 1063 else u'Media')

# (3) inscripcion declarada sin respaldo en la base de alumnos
for p in sorted(prospectos, key=lambda x: x['completo'].lower()):
    if not p['inscrito_sin_confirmar']:
        continue
    fuente = u'FUENTE 2 – Clases Muestra' if p['regs'] else u'FUENTE 3 – Campañas'
    ref = (u'%s, fila %s' % (p['regs'][-1]['mes'], p['regs'][-1]['fila_origen'])) if p['regs'] \
        else (p['camp_n'][-1]['campana'] if p['camp_n'] else '')
    add_val(u'Inscripción declarada sin confirmar', p['completo'], ref, fuente,
            u'Resultado registrado "Inscrito" el %s. Teléfono: %s. No aparece en el Reporte General de Alumnos.'
            % (fdate(p['cm_fecha']) or SIN_INFO, p['telefono'] or SIN_INFO),
            u'Confirmar si la persona se inscribió y causó baja, si la inscripción no se concretó, o si '
            u'la matrícula existe con otro nombre. Determina si es alumno, prospecto o registro a depurar.',
            u'Alta')

# (4) coincidencias confirmadas y descartadas manualmente
for key, (mat, why) in sorted(CONFIRMA.items()):
    reg = next((r for r in cms if nkey(r['nombre']) == key), None)
    add_val(u'Cruce confirmado por validación manual', reg['nombre'] if reg else key,
            u'Matrícula %d' % mat, u'FUENTE 1 vs FUENTE 2',
            u'Se asoció con %s. Criterio: %s' % (al_by_mat[mat]['completo'], why),
            u'Ratificar el cruce con el expediente. Si no es la misma persona, el alumno pasa a '
            u'"sin procedencia identificada" y el registro vuelve a prospecto.', u'Media')
for key, why in sorted(RECHAZA.items()):
    reg = next((r for r in cms if nkey(r['nombre']) == key), None)
    add_val(u'Coincidencia de nombre descartada', reg['nombre'] if reg else key, u'—',
            u'FUENTE 1 vs FUENTE 2', why,
            u'Ratificar que NO se trata de la misma persona. Si lo fuera, debe moverse a '
            u'"Alumnos inscritos" y se reduce el total de prospectos.', u'Media')
for key, why in sorted(VALIDAR_EXTRA.items()):
    reg = next((r for r in cms if nkey(r['nombre']) == key), None)
    add_val(u'Información contradictoria', reg['nombre'] if reg else key,
            (u'%s, fila %s' % (reg['mes'], reg['fila_origen'])) if reg else u'—',
            u'FUENTE 2 – Clases Muestra', why,
            u'Identificar a la persona con nombre y apellidos completos y definir su estatus real.', u'Alta')

# (5) nombres incompletos / genericos
vistos = set()
for p in sorted(prospectos, key=lambda x: x['completo'].lower()):
    if p['inscrito_sin_confirmar'] or nkey(p['completo']) in VALIDAR_EXTRA:
        continue
    ntok = len(tokset(p['regs'][-1]['tok'] if p['regs'] else k.tokens(p['completo'])))
    if ntok >= GENERICO_MAX_TOK:
        continue
    if p['completo'] in vistos:
        continue
    vistos.add(p['completo'])
    fuente = u'FUENTE 2 – Clases Muestra' if p['regs'] else u'FUENTE 3 – Campañas'
    ref = (u'%s, fila %s' % (p['regs'][-1]['mes'], p['regs'][-1]['fila_origen'])) if p['regs'] else \
        (p['camp_n'][-1]['campana'] if p['camp_n'] else '')
    add_val(u'Nombre incompleto o no identificable', p['completo'], ref, fuente,
            u'Registro con un solo nombre o apodo. Estatus asignado: %s. Teléfono: %s. Fecha: %s.'
            % (p['estatus'], p['telefono'] or SIN_INFO, fdate(p['cm_fecha'] or p['primer_contacto']) or SIN_INFO),
            u'Completar nombre y apellidos para poder cruzarlo contra alumnos y detectar duplicados.',
            u'Media')
for nombre in [u'2 niños']:
    reg = next((r for r in cms if r['nombre'] == nombre), None)
    if reg:
        add_val(u'Nombre incompleto o no identificable', nombre,
                u'%s, fila %s' % (reg['mes'], reg['fila_origen']), u'FUENTE 2 – Clases Muestra',
                u'El campo nombre dice "2 niños": el registro representa a dos personas sin identificar. '
                u'Teléfono %s, clase muestra del %s con resultado "%s".'
                % (reg['telefono'] or SIN_INFO, fdate(reg['fecha']), reg['resultado']),
                u'Separar en dos registros con nombre y apellidos y determinar el estatus de cada uno. '
                u'El conteo de clases muestra está subestimado en 1.', u'Alta')

# (6) personas con mas de una clase muestra
for p in sorted(personas.values(), key=lambda x: x['completo'].lower()):
    if p['n_cm'] < 2:
        continue
    det = ' ; '.join(u'%s %s (%s)' % (r['mes'], fdate(r['fecha']), r['resultado'] or SIN_INFO) for r in p['regs'])
    add_val(u'Más de una clase muestra para la misma persona', p['completo'],
            (u'Matrícula %s' % p['al']['matricula']) if p['al'] else u'—',
            u'FUENTE 2 – Clases Muestra',
            u'%d registros de clase muestra: %s' % (p['n_cm'], det),
            u'Confirmar que se trata de la misma persona (no un duplicado de captura). Los registros se '
            u'contabilizaron como %d clases muestra y 1 sola persona.' % p['n_cm'], u'Baja')

# (7) fechas con error de captura
for r in cms:
    if not r['fecha_corr']:
        continue
    add_val(u'Fecha con error de captura', r['nombre'], u'%s, fila %s' % (r['mes'], r['fila_origen']),
            u'FUENTE 2 – Clases Muestra', r['fecha_corr'] + u'. Fecha usada en el análisis: %s' % fdate(r['fecha']),
            u'Corregir la fecha en la fuente. No cambia el estatus comercial de la persona.', u'Baja')

# (8) discrepancias de edad entre fuentes
for p in sorted(convertidos, key=lambda x: x['completo'].lower()):
    if not (p['al'] and p['al']['edad'] and p['edad_cm']):
        continue
    dif = abs(p['al']['edad'] - p['edad_cm'])
    if dif < 2:
        continue
    add_val(u'Edad inconsistente entre fuentes', p['completo'], u'Matrícula %s' % p['al']['matricula'],
            u'FUENTE 1 vs FUENTE 2',
            u'Edad en alumnos: %s años. Edad en clases muestra: %s años. Diferencia de %d años.'
            % (p['al']['edad'], p['edad_cm'], dif),
            u'Verificar la fecha de nacimiento en el expediente. Si la diferencia es grande, confirmar '
            u'que ambos registros son de la misma persona.', u'Media' if dif >= 5 else u'Baja')

# (9) citas vencidas sin resultado
for p in sorted(prospectos, key=lambda x: x['completo'].lower()):
    if p['estatus'] != u'Información insuficiente' or not p['regs']:
        continue
    if not any(r['resultado'] == 'Cita' for r in p['regs']):
        continue
    add_val(u'Cita sin resultado registrado', p['completo'],
            u'%s, fila %s' % (p['regs'][-1]['mes'], p['regs'][-1]['fila_origen']),
            u'FUENTE 2 – Clases Muestra',
            u'Cita agendada el %s (hace %s días respecto al corte) y el campo RESULTADO sigue en "Cita".'
            % (fdate(p['cm_fecha']), p['dias'] if p['dias'] is not None else '?'),
            u'Registrar si la persona asistió, no asistió o canceló. Sin este dato no puede contarse '
            u'ni como clase muestra realizada ni como no asistencia.', u'Alta')

# (9b) resultado posterior al corte marcado como realizado
for r in cms:
    if not (r['fecha'] and r['fecha'] > CORTE and r['resultado'] in ('Show', 'Inscrito')):
        continue
    add_val(u'Resultado anticipado a la fecha de corte', r['nombre'],
            u'%s, fila %s' % (r['mes'], r['fila_origen']), u'FUENTE 2 – Clases Muestra',
            u'La clase muestra está agendada para el %s (posterior al corte del 03/10/2026) y el campo '
            u'RESULTADO ya dice "%s".' % (fdate(r['fecha']), r['resultado']),
            u'Confirmar la fecha real de la clase. En este corte la persona se contabilizó como '
            u'"Clase muestra por realizar" y NO como clase muestra realizada.', u'Media')

# (10) telefonos invalidos
for r in cms:
    tel = r['telefono']
    if not tel:
        continue
    digs = ''.join(ch for ch in tel if ch.isdigit())
    problema = ''
    if not digs:
        problema = u'El campo teléfono no contiene dígitos (valor capturado: "%s").' % tel
    elif len(digs) > 12:
        problema = u'El teléfono tiene %d dígitos (valor capturado: "%s").' % (len(digs), tel)
    if not problema:
        continue
    add_val(u'Teléfono inválido', r['nombre'], u'%s, fila %s' % (r['mes'], r['fila_origen']),
            u'FUENTE 2 – Clases Muestra', problema,
            u'Corregir el número de contacto; sin él no es posible dar seguimiento ni detectar '
            u'registros de la misma familia.', u'Media')

# (11) leads de campana sin nombre
if camp_desc['sin_nombre']:
    add_val(u'Registros sin identificación', u'(varios)', u'—', u'FUENTE 3 – Campañas',
            u'%d filas de seguimiento de leads no tienen nombre del interesado (sólo teléfono o medio '
            u'de contacto).' % camp_desc['sin_nombre'],
            u'Completar el nombre para poder cruzarlos. Estos registros quedaron fuera del tablero '
            u'de prospectos por falta de identificación.', u'Media')
if camp_desc['sin_avance']:
    add_val(u'Leads sin resultado de seguimiento', u'(varios)', u'—', u'FUENTE 3 – Campañas',
            u'%d filas quedan en etapa inicial ("Información enviada", "Nuevo lead" o "Contactado") sin '
            u'ningún campo de resultado, registro o asistencia capturado.' % camp_desc['sin_avance'],
            u'Definir si siguen en proceso o deben cerrarse. No se incluyeron en "Prospectos en curso" '
            u'porque la fuente no registra avance comercial alguno.', u'Media')

for i, row in enumerate(val, 1):
    row.insert(0, i)

# ===========================================================================
# 10. METRICAS
# ===========================================================================
n_reg_alumnos = len(alumnos)
n_dup = sum(len(g) - 1 for g in dup_pairs)
n_alumnos_unicos = n_reg_alumnos - n_dup
n_prueba = sum(1 for a in alumnos if nkey(a['ap_mat']) == 'prueba')
n_alumnos_dep = n_alumnos_unicos - n_prueba

cm_reg = len(cms)


def _post(r):
    return bool(r['fecha']) and r['fecha'] > CORTE


cm_realizadas = sum(1 for r in cms if r['resultado'] in ('Show', 'Inscrito') and not _post(r))
cm_futuras = sum(1 for r in cms if _post(r) and r['resultado'] in ('Cita', 'Cita pospuesta', 'Show', 'Inscrito'))
cm_vencidas = sum(1 for r in cms if not _post(r) and r['resultado'] in ('Cita', 'Cita pospuesta'))
cm_agendadas = cm_futuras + cm_vencidas
cm_noasistio = sum(1 for r in cms if r['resultado'] == u'NO ASISTIÓ')
cm_lead = sum(1 for r in cms if r['resultado'] == 'Lead')
cm_otros = cm_reg - cm_realizadas - cm_agendadas - cm_noasistio - cm_lead

personas_cm = [p for p in personas.values() if p['n_cm'] > 0]
cm_asistieron = [p for p in personas_cm if p['asistio']]
cm_convertidas = [p for p in cm_asistieron if p['inscrito']]
cm_declarada = [p for p in personas_cm if p['inscrito_sin_confirmar']]
conv = (float(len(cm_convertidas)) / len(cm_asistieron)) if cm_asistieron else 0.0
conv_max = (float(len(cm_convertidas) + len(cm_declarada)) / len(cm_asistieron)) if cm_asistieron else 0.0

act_cierre = [p for p in prospectos if p['clase'] == u'Activo – en cierre']
act_cm = [p for p in prospectos if p['clase'] == u'Activo – clase muestra pendiente']
recuperables = [p for p in prospectos if p['clase'] == u'Recuperable – no asistió']
descartados = [p for p in prospectos if p['clase'] == u'Descartado']
por_validar = [p for p in prospectos if p['clase'] == u'Por validar']
activos = act_cierre + act_cm

n_seguimiento = sum(1 for p in prospectos if p['estatus'] == u'Seguimiento post clase muestra')
cnt_est = collections.Counter(p['estatus'] for p in prospectos)


def recencia(lst, lo, hi):
    out = 0
    for p in lst:
        d = p['dias']
        if d is None:
            continue
        if lo <= d <= hi:
            out += 1
    return out


sin_fecha = sum(1 for p in activos if p['dias'] is None)
camp_total = len(camps)
camp_nombre_link = sum(1 for c in camps if c['link_nombre'])
camp_extra_filas = sum(len(v) for v in camp_extra.values())
# registros de alumno (no personas) con procedencia identificada
mats_con_cm = set()
for p in convertidos:
    for a in alumnos:
        if mat_canon.get(a['matricula'], a['matricula']) == p['al']['matricula']:
            mats_con_cm.add(a['matricula'])
n_reg_con_cm = len(mats_con_cm)

# ===========================================================================
# 11. GENERACION DEL EXCEL
# ===========================================================================
wb = w.Workbook()
HOY = datetime.date.today()

# ------------------------------------------------------------- 1. RESUMEN --
sh = wb.add_sheet(u'RESUMEN')
sh.set_widths([3, 52, 14, 14, 62])
sh.set_freeze(3)
sh.row_heights[0] = 26
sh.write(0, 1, u'KUGELMAN ACADEMY · Corte comercial al sábado 3 de octubre de 2026', 'title')
sh.merge(0, 1, 0, 4)
sh.write(1, 1, u'Alumnos inscritos y prospectos en curso · Documento de trabajo para Dirección', 'subtitle')
sh.merge(1, 1, 1, 4)
sh.write(2, 1, u'Generado el %s a partir de: RGA_1791241975.xlsx (alumnos), CLASES MUESTRA 2026__ (1).xlsx '
               u'y CAMPAÑAS (2).xlsx. Las fuentes originales no fueron modificadas.' % HOY.strftime('%d/%m/%Y'),
         'note')
sh.merge(2, 1, 2, 4)

row = 4


def seccion(titulo):
    global row
    sh.write(row, 1, titulo, 'section')
    sh.write(row, 2, '', 'section')
    sh.write(row, 3, '', 'section')
    sh.write(row, 4, '', 'section')
    sh.merge(row, 1, row, 4)
    row += 1


def kpi(label, valor, nota='', estilo='metric'):
    global row
    sh.write(row, 1, label, 'label')
    sh.write(row, 2, valor, estilo)
    sh.write(row, 3, '', 'text')
    sh.merge(row, 2, row, 3)
    sh.write(row, 4, nota, 'text')
    row += 1


seccion(u'1 · ALUMNOS INSCRITOS')
kpi(u'Total de alumnos inscritos actuales (registros en la base)', n_reg_alumnos,
    u'Todos con estatus "Activo = Sí" en el Reporte General de Alumnos.')
kpi(u'Personas únicas tras depurar matrículas duplicadas', n_alumnos_unicos,
    u'Se detectaron %d matrículas duplicadas (ver pestaña REGISTROS POR VALIDAR).' % n_dup)
kpi(u'Alumnos reales (sin el registro de prueba del sistema)', n_alumnos_dep,
    u'Se excluye la matrícula 1063 "Gustavo Martinez Prueba".')
kpi(u'Alumnos con procedencia de clase muestra identificada', len(cm_convertidas),
    u'Personas únicas. En la pestaña ALUMNOS INSCRITOS son %d registros marcados "Sí" (la diferencia '
    u'proviene de matrículas duplicadas). El resto se inscribió antes de que iniciara el registro de '
    u'clases muestra o no fue posible identificar su origen.' % n_reg_con_cm)

row += 1
seccion(u'2 · CLASES MUESTRA (enero – 8 de octubre de 2026)')
kpi(u'Total de clases muestra registradas', cm_reg,
    u'Filas con nombre en la base CLASES MUESTRA 2026 (10 hojas mensuales).')
kpi(u'Clases muestra realizadas', cm_realizadas,
    u'Resultado "Show" o "Inscrito" con fecha igual o anterior al 3 de octubre de 2026.')
kpi(u'Clases muestra pendientes / agendadas', cm_agendadas, u'Suma de las dos líneas siguientes.')
kpi(u'   · agendadas para después del corte', cm_futuras,
    u'Citas del 5, 6 y 8 de octubre: cuentan como prospecto activo.')
kpi(u'   · citas vencidas sin resultado registrado', cm_vencidas,
    u'Citas anteriores al corte en las que nunca se capturó si la persona asistió.')
kpi(u'No asistencias', cm_noasistio, u'Resultado "NO ASISTIÓ".')
kpi(u'Registros marcados sólo como "Lead"', cm_lead, u'Captados pero sin clase muestra realizada.')
kpi(u'Registros sin resultado capturado', cm_otros,
    u'Verificación: %d + %d + %d + %d + %d = %d registros.'
    % (cm_realizadas, cm_agendadas, cm_noasistio, cm_lead, cm_otros, cm_reg))
kpi(u'Personas distintas en la base de clases muestra', len(personas_cm),
    u'%d registros corresponden a personas con más de una clase muestra.'
    % (cm_reg - len(personas_cm)))

row += 1
seccion(u'3 · EMBUDO DE PROSPECTOS (personas que NO aparecen como alumnos)')
kpi(u'Total de prospectos activos', len(activos),
    u'Suma de las dos líneas siguientes. No incluye descartados ni registros por validar.')
kpi(u'   · Activos en cierre (ya tuvieron contacto o clase muestra)', len(act_cierre),
    u'Seguimiento post clase muestra, pendientes de respuesta e interesados por cerrar.')
kpi(u'   · Activos con clase muestra pendiente', len(act_cm),
    u'Citas agendadas a futuro y reprogramaciones pendientes.')
kpi(u'Total de prospectos en seguimiento post clase muestra', n_seguimiento,
    u'Asistieron a la clase muestra y todavía no se inscriben: es la bolsa de cierre más inmediata.')
kpi(u'Recuperables – no asistieron', len(recuperables),
    u'Sin evidencia de desinterés: deben reagendarse antes de descartarlos.')
kpi(u'Descartados con evidencia explícita', len(descartados),
    u'La fuente registra "No interesado", "Canceló proceso" o "Sin respuesta".')
kpi(u'Registros por validar antes de contarlos como prospectos', len(por_validar),
    u'Información contradictoria o insuficiente. Ver pestaña REGISTROS POR VALIDAR.')

row += 1
seccion(u'4 · CONVERSIÓN CLASE MUESTRA → INSCRIPCIÓN')
kpi(u'Personas que realizaron clase muestra', len(cm_asistieron), u'Denominador de la conversión.')
kpi(u'Personas que realizaron clase muestra y hoy son alumnos', len(cm_convertidas),
    u'Numerador: confirmadas en el Reporte General de Alumnos.')
kpi(u'Conversión confirmada clase muestra → inscripción', w.Pct(conv),
    u'%d de %d personas.' % (len(cm_convertidas), len(cm_asistieron)), 'metricpct')
kpi(u'Conversión máxima posible (techo)', w.Pct(conv_max),
    u'Incluye %d personas que la fuente marca "Inscrito" pero que no aparecen en la base de alumnos. '
    u'La conversión real está entre ambos valores.' % len(cm_declarada), 'metricpct')

row += 1
seccion(u'5 · PROSPECTOS POR ESTATUS')
sh.write(row, 1, u'Estatus', 'header')
sh.write(row, 2, u'Personas', 'header')
sh.write(row, 3, u'% del total', 'header')
sh.write(row, 4, u'Clasificación del embudo', 'header')
row += 1
CLASE_DE_ESTATUS = {}
for p in prospectos:
    CLASE_DE_ESTATUS.setdefault(p['estatus'], set()).add(p['clase'])
tot_p = len(prospectos)
for est in ORDEN_ESTATUS:
    n = cnt_est.get(est, 0)
    if not n:
        continue
    sh.write(row, 1, est, SEMAFORO.get(est, 'text'))
    sh.write(row, 2, n, 'num')
    sh.write(row, 3, w.Pct(float(n) / tot_p), 'pct')
    sh.write(row, 4, ' / '.join(sorted(CLASE_DE_ESTATUS.get(est, []))), 'text')
    row += 1
sh.write(row, 1, u'TOTAL DE PROSPECTOS (no inscritos)', 'label')
sh.write(row, 2, tot_p, 'metric')
sh.write(row, 3, w.Pct(1.0), 'pct')
sh.write(row, 4, u'Incluye activos, recuperables, descartados y registros por validar.', 'text')
row += 2

seccion(u'6 · ANTIGÜEDAD DEL ÚLTIMO MOVIMIENTO (prospectos activos)')
for etq, lo, hi in [(u'Movimiento en los últimos 30 días', 0, 30),
                    (u'Entre 31 y 60 días sin movimiento', 31, 60),
                    (u'Entre 61 y 120 días sin movimiento', 61, 120),
                    (u'Más de 120 días sin movimiento', 121, 10 ** 6)]:
    kpi(etq, recencia(activos, lo, hi), '')
kpi(u'Sin fecha disponible', sin_fecha, u'No es posible medir la antigüedad del contacto.')

row += 1
seccion(u'7 · CONCILIACIÓN DE REGISTROS')
for lab, v, nota in [
    (u'FUENTE 1 · Registros en el Reporte General de Alumnos', n_reg_alumnos,
     u'= %d personas únicas + %d matrículas duplicadas.' % (n_alumnos_unicos, n_dup)),
    (u'FUENTE 2 · Registros de clase muestra', cm_reg,
     u'%d registros corresponden a personas que tomaron más de una clase muestra.'
     % (cm_reg - len(personas_cm))),
    (u'   · personas distintas', len(personas_cm),
     u'Las dos líneas siguientes suman este total.'),
    (u'   · de ellas, confirmadas como alumnos inscritos', len(cm_convertidas),
     u'Pasan a la pestaña ALUMNOS INSCRITOS.'),
    (u'   · de ellas, prospectos o registros por validar', len(personas_cm) - len(cm_convertidas),
     u'Pasan a la pestaña PROSPECTOS EN CURSO.'),
    (u'FUENTE 3 · Registros de seguimiento de leads (CAMPAÑAS)', camp_total,
     u'Fuente de apoyo: aporta teléfono, fecha de primer contacto y resultado de seguimiento.'),
    (u'   · vinculados por nombre a un registro de clase muestra', camp_nombre_link, u''),
    (u'   · sin nombre del interesado', camp_desc['sin_nombre'], u'No cruzables.'),
    (u'   · en etapa inicial, sin avance comercial registrado', camp_desc['sin_avance'],
     u'Tope del embudo: no se listan individualmente.'),
    (u'   · ya identificados como alumnos inscritos', camp_desc['ya_alumno'], u''),
    (u'   · mismo teléfono que un registro de clase muestra', camp_desc['mismo_telefono'],
     u'Mismo hogar: no se duplican como prospecto nuevo.'),
    (u'   · incorporados como prospectos adicionales', camp_extra_filas,
     u'%d registros que corresponden a %d personas: leads con avance comercial y sin registro en '
     u'CLASES MUESTRA.' % (camp_extra_filas, len(camp_extra))),
]:
    kpi(lab, v, nota)
row += 1
kpi(u'PERSONAS EN EL TABLERO', len(personas),
    u'= %d convertidas (alumnos con origen en clase muestra) + %d prospectos.'
    % (len(convertidos), len(prospectos)))
kpi(u'Filas en ALUMNOS INSCRITOS', n_reg_alumnos, u'Se conservan todos los registros de la FUENTE 1.')
kpi(u'Filas en PROSPECTOS EN CURSO', tot_p, u'Ninguna persona aparece en las dos pestañas.')
kpi(u'Filas en HISTÓRICO CLASES MUESTRA', cm_reg, u'Trazabilidad completa de la FUENTE 2.')
kpi(u'Filas en REGISTROS POR VALIDAR', len(val), u'Hallazgos que requieren revisión manual.')

row += 1
seccion(u'8 · NOTAS METODOLÓGICAS')
NOTAS = [
    u'Corte de la información: sábado 3 de octubre de 2026. Las clases muestra agendadas para el 5, 6 y 8 de '
    u'octubre se conservan porque, según el criterio comercial, una cita futura es un prospecto activo.',
    u'Una persona se considera INSCRITA sólo si aparece en el Reporte General de Alumnos. Si la base de clases '
    u'muestra dice "Inscrito" pero no hay respaldo en alumnos, el registro NO se cuenta como alumno ni como '
    u'prospecto activo: pasa a REGISTROS POR VALIDAR.',
    u'Haber asistido a una clase muestra sin inscribirse NO se interpreta como prospecto perdido. Sólo se '
    u'descarta a quien tiene evidencia explícita en la fuente ("No interesado", "Canceló proceso", '
    u'"Sin respuesta" o una nota equivalente).',
    u'El cruce entre fuentes se hizo por teléfono y por nombre normalizado (sin acentos, sin mayúsculas, sin '
    u'espacios dobles, resolviendo abreviaturas como "Hdz" = Hernández). Las coincidencias dudosas se revisaron '
    u'una por una y quedaron documentadas en REGISTROS POR VALIDAR.',
    u'Los resultados de seguimiento de CAMPAÑAS se heredan únicamente cuando el nombre coincide. El teléfono '
    u'sólo se usa para la fecha de primer contacto, porque las familias comparten el mismo número.',
    u'La columna "Procedencia de clase muestra" usa "No identificado" en lugar de "No": la base de clases '
    u'muestra empieza en enero de 2026 y no permite demostrar que un alumno anterior no tuvo clase muestra.',
    u'Ningún dato fue inventado. Lo que no existe en las fuentes aparece vacío o como "Sin información".',
]
for n in NOTAS:
    sh.write(row, 1, u'•', 'center')
    sh.write(row, 2, n, 'text')
    sh.merge(row, 2, row, 4)
    sh.row_heights[row] = 30
    row += 1


def hoja(nombre, titulo, subtitulo, cols, widths, tabla):
    """Crea una hoja con titulo, encabezado y tabla estructurada."""
    s = wb.add_sheet(nombre)
    s.set_widths(widths)
    s.row_heights[0] = 22
    s.write(0, 0, titulo, 'title')
    s.merge(0, 0, 0, len(cols) - 1)
    s.write(1, 0, subtitulo, 'note')
    s.merge(1, 0, 1, len(cols) - 1)
    s.row_heights[2] = 32
    for j, c in enumerate(cols):
        s.write(2, j, c, 'header')
    s.set_freeze(3)
    s._tabla = tabla
    return s


# ------------------------------------------------- 2. ALUMNOS INSCRITOS ----
COLS_AL = [u'ID (matrícula)', u'Nombre', u'Apellidos', u'Nombre completo', u'Edad', u'Sexo',
           u'Estatus actual', u'Procedencia de clase muestra', u'Fecha de clase muestra',
           u'Categoría de la clase muestra', u'Medio de contacto original', u'Observaciones']
sh2 = hoja(u'ALUMNOS INSCRITOS',
           u'ALUMNOS INSCRITOS · Corte al 3 de octubre de 2026',
           u'Fuente: RGA_1791241975.xlsx (Reporte General de Alumnos). Se conservan todos los registros, '
           u'con o sin antecedente de clase muestra. La procedencia se obtuvo cruzando contra '
           u'CLASES MUESTRA 2026.',
           COLS_AL, [14, 22, 24, 34, 7, 7, 16, 16, 15, 15, 20, 70], 'tblAlumnos')

per_por_mat = {}
for p in convertidos:
    per_por_mat[p['al']['matricula']] = p
rev_map = dict(REVISION_RGA)

r0 = 3
for i, a in enumerate(sorted(alumnos, key=lambda x: x['matricula'])):
    canon = mat_canon.get(a['matricula'], a['matricula'])
    p = per_por_mat.get(canon)
    obs = []
    if a['matricula'] in mat_canon:
        otras = [str(x['matricula']) for g in dup_pairs for x in g
                 if mat_canon.get(x['matricula']) == canon and x['matricula'] != a['matricula']]
        obs.append(u'Posible matrícula duplicada con %s (ver REGISTROS POR VALIDAR).' % ', '.join(otras))
    if a['matricula'] in rev_map:
        obs.append(u'Revisión de captura: %s.' % rev_map[a['matricula']])
    if p:
        obs.append(u'Cruce con clase muestra: %s.' % p['regs'][-1]['match_why'])
        if p['n_cm'] > 1:
            obs.append(u'Registra %d clases muestra (%s).'
                       % (p['n_cm'], ', '.join(fdate(r['fecha']) for r in p['regs'])))
        if p['edad_cm'] and a['edad'] and abs(p['edad_cm'] - a['edad']) >= 2:
            obs.append(u'Edad distinta entre fuentes: %s años en alumnos vs %s en clase muestra.'
                       % (a['edad'], p['edad_cm']))
        if any(r['resultado'] == 'Show' for r in p['regs']) and \
                not any(r['resultado'] == 'Inscrito' for r in p['regs']):
            obs.append(u'La clase muestra quedó registrada como "Show": la inscripción se concretó después '
                       u'y no se actualizó el resultado en la fuente.')
    else:
        obs.append(u'Sin registro de clase muestra localizado en la base 2026.')
    vals = [
        a['matricula'], a['nombre'], a['apellidos'], a['completo'],
        a['edad'], a['sexo'],
        u'Inscrito (activo)' if a['activo'] == u'Sí' else (a['activo'] or SIN_INFO),
        u'Sí' if p else u'No identificado',
        p['cm_fecha_min'] if p else '',
        (p['categoria'] or SIN_INFO) if p else '',
        (p['medio'] or SIN_INFO) if p else '',
        ' '.join(obs),
    ]
    estilos = ['center', 'text', 'text', 'text', 'center', 'center', 'center',
               'center', 'date', 'center', 'center', 'text']
    for j, v in enumerate(vals):
        sh2.write(r0 + i, j, v, estilos[j])
sh2.add_table('tblAlumnos', 2, 0, r0 + len(alumnos) - 1, len(COLS_AL) - 1, COLS_AL)

# ------------------------------------------------ 3. PROSPECTOS EN CURSO ---
COLS_PR = [u'Nombre', u'Apellidos', u'Nombre completo', u'Edad', u'Sexo', u'Teléfono',
           u'Fecha de primer contacto', u'Fecha de clase muestra', u'Horario de clase muestra',
           u'Resultado de clase muestra', u'Estatus actual', u'Clasificación del embudo',
           u'Días sin movimiento', u'Categoría', u'Medio de contacto',
           u'Último movimiento disponible', u'Próxima acción sugerida', u'Observaciones']
sh3 = hoja(u'PROSPECTOS EN CURSO',
           u'PROSPECTOS EN CURSO · Personas que NO aparecen como alumnos inscritos · Corte al 3 de octubre de 2026',
           u'Semáforo en "Estatus actual": verde = interesado / por cerrar · amarillo = seguimiento, '
           u'pendiente o reprogramación · rojo = no asistió, no interesado o sin respuesta. '
           u'Ninguna persona de esta pestaña aparece en ALUMNOS INSCRITOS.',
           COLS_PR, [20, 24, 32, 7, 7, 16, 14, 14, 13, 17, 28, 26, 11, 13, 15, 62, 46, 70],
           'tblProspectos')

ORD_CLASE = {u'Activo – en cierre': 0, u'Activo – clase muestra pendiente': 1,
             u'Recuperable – no asistió': 2, u'Por validar': 3, u'Descartado': 4}
ORD_EST = {e: i for i, e in enumerate(ORDEN_ESTATUS)}
prospectos_ord = sorted(prospectos, key=lambda p: (ORD_CLASE.get(p['clase'], 9),
                                                   ORD_EST.get(p['estatus'], 9),
                                                   p['dias'] if p['dias'] is not None else 10 ** 6,
                                                   p['completo'].lower()))
for i, p in enumerate(prospectos_ord):
    obs = [p['motivo'] + '.']
    if p['notas']:
        obs.append(u'Notas de la fuente: "%s".' % p['notas'])
    if p['n_cm'] > 1:
        obs.append(u'%d clases muestra registradas: %s.'
                   % (p['n_cm'], ', '.join(u'%s (%s)' % (fdate(r['fecha']), r['resultado'] or SIN_INFO)
                                           for r in p['regs'])))
    if p['solo_campana']:
        obs.append(u'Procede de la base de CAMPAÑAS; no tiene registro en CLASES MUESTRA.')
    if p['camp_t']:
        obs.append(u'El teléfono coincide con %d registro(s) de CAMPAÑAS de otra persona del mismo hogar; '
                   u'no se heredó su resultado.' % len(p['camp_t']))
    if p['regs'] and p['regs'][-1]['match_score'] >= 45 and not p['inscrito']:
        obs.append(p['regs'][-1]['match_why'] + '.')
    if p['inscrito_sin_confirmar']:
        obs.append(u'REQUIERE VALIDACIÓN: la fuente lo marca como inscrito.')
    vals = [
        p['nombre'], p['apellidos'], p['completo'],
        p['edad'] if p['edad'] else '',
        p['sexo'] or SIN_INFO,
        p['telefono'] or SIN_INFO,
        p['primer_contacto'] or '',
        p['cm_fecha'] or '',
        p['horario'] or SIN_INFO,
        p['resultado_cm'],
        p['estatus'],
        p['clase'],
        p['dias'] if p['dias'] is not None else '',
        p['categoria'] or SIN_INFO,
        p['medio'] or SIN_INFO,
        p['ultimo_mov'],
        ACCIONES.get(p['estatus'], ''),
        ' '.join(obs),
    ]
    estilos = ['text', 'text', 'text', 'center', 'center', 'center', 'date', 'date', 'center',
               'center', SEMAFORO.get(p['estatus'], 'text'), 'text', 'num', 'center', 'center',
               'text', 'text', 'text']
    for j, v in enumerate(vals):
        sh3.write(3 + i, j, v, estilos[j])
sh3.add_table('tblProspectos', 2, 0, 3 + len(prospectos_ord) - 1, len(COLS_PR) - 1, COLS_PR)

# -------------------------------------- 4. HISTORICO CLASES MUESTRA -------
COLS_H = [u'#', u'Hoja de origen', u'Consecutivo en la fuente', u'Fila en la fuente',
          u'Fecha de clase muestra', u'Nombre registrado', u'Teléfono', u'Categoría', u'Edad',
          u'Idioma', u'Medio de contacto', u'Horario', u'Resultado original',
          u'Notas de la fuente', u'Matrícula del alumno', u'Alumno confirmado',
          u'Base del cruce', u'Estatus de la persona', u'Resultado comercial',
          u'Observaciones de calidad del dato']
sh4 = hoja(u'HISTÓRICO CLASES MUESTRA',
           u'HISTÓRICO DE CLASES MUESTRA 2026 · Información original depurada para auditoría',
           u'Un renglón por cada registro de CLASES MUESTRA 2026__ (1).xlsx (10 hojas mensuales). '
           u'Se conservan los valores originales y se agregan el cruce contra alumnos y el resultado comercial.',
           COLS_H, [5, 14, 11, 10, 15, 34, 16, 13, 7, 9, 15, 12, 15, 30, 13, 32, 46, 28, 20, 44],
           'tblHistorico')

per_de_reg = {}
for p in personas.values():
    for r in p['regs']:
        per_de_reg[id(r)] = p

cms_ord = sorted(cms, key=lambda r: (r['fecha'] or datetime.date(1900, 1, 1), r['nombre'].lower()))
for i, r in enumerate(cms_ord):
    p = per_de_reg[id(r)]
    if p['inscrito']:
        rc = u'Se inscribió'
    elif p['inscrito_sin_confirmar']:
        rc = u'Por validar (posible inscripción)'
    else:
        rc = RES_COMERCIAL.get(p['estatus'], u'Sin información suficiente')
    qual = []
    if r['fecha_corr']:
        qual.append(r['fecha_corr'] + '.')
    if not r['telefono']:
        qual.append(u'Sin teléfono de contacto.')
    if not r['resultado']:
        qual.append(u'Campo RESULTADO vacío.')
    if len(tokset(r['tok'])) < 2:
        qual.append(u'Nombre incompleto o apodo.')
    if p['n_cm'] > 1:
        qual.append(u'La persona tiene %d clases muestra registradas.' % p['n_cm'])
    vals = [
        i + 1, r['mes'], r['consec'], r['fila_origen'], r['fecha'] or '', r['nombre'],
        r['telefono'] or SIN_INFO, r['categoria'] or SIN_INFO,
        r['edad'] if r['edad'] else (r['edad_txt'] or ''),
        r['idioma'] or SIN_INFO, r['medio'] or SIN_INFO, r['horario'] or SIN_INFO,
        r['resultado'] or SIN_INFO, r['notas'] or '',
        r['al']['matricula'] if r['al'] else '',
        p['al']['completo'] if p['inscrito'] else '',
        r['match_why'], p['estatus'], rc, ' '.join(qual),
    ]
    estilos = ['num', 'center', 'center', 'center', 'date', 'text', 'center', 'center', 'center',
               'center', 'center', 'center', 'center', 'text', 'center', 'text', 'text',
               'center', 'center', 'text']
    for j, v in enumerate(vals):
        sh4.write(3 + i, j, v, estilos[j])
sh4.add_table('tblHistorico', 2, 0, 3 + len(cms_ord) - 1, len(COLS_H) - 1, COLS_H)

# ------------------------------------------ 5. REGISTROS POR VALIDAR ------
COLS_V = [u'#', u'Tipo de hallazgo', u'Registro o persona', u'Referencia en la fuente', u'Fuente',
          u'Información disponible', u'Qué debe validarse', u'Prioridad']
sh5 = hoja(u'REGISTROS POR VALIDAR',
           u'REGISTROS POR VALIDAR · Casos que requieren revisión manual antes de darlos por definitivos',
           u'Ningún registro dudoso fue eliminado. Cada renglón indica exactamente qué debe confirmarse y '
           u'qué efecto tendría en los totales del RESUMEN.',
           COLS_V, [5, 38, 34, 22, 30, 78, 78, 11], 'tblValidar')
PRIO = {u'Alta': 'red', u'Media': 'amber', u'Baja': 'text'}
val_ord = sorted(val, key=lambda v: ({u'Alta': 0, u'Media': 1, u'Baja': 2}[v[7]], v[1], str(v[2])))
for i, rowv in enumerate(val_ord):
    estilos = ['num', 'text', 'text', 'center', 'center', 'text', 'text', PRIO[rowv[7]]]
    for j, v in enumerate(rowv):
        sh5.write(3 + i, j, (i + 1) if j == 0 else v, estilos[j])
sh5.add_table('tblValidar', 2, 0, 3 + len(val_ord) - 1, len(COLS_V) - 1, COLS_V)

wb.save(SALIDA)
print(u'Archivo generado: %s' % SALIDA)
print()
print(u'== RESUMEN DE CONTROL ==')
print(u'Alumnos: %d registros / %d personas unicas / %d sin registro de prueba' % (n_reg_alumnos, n_alumnos_unicos, n_alumnos_dep))
print(u'Clases muestra: %d registros = realizadas %d + agendadas %d (futuras %d + vencidas %d) + no asistio %d + lead %d + otros %d'
      % (cm_reg, cm_realizadas, cm_agendadas, cm_futuras, cm_vencidas, cm_noasistio, cm_lead, cm_otros))
print(u'Personas distintas en clases muestra: %d' % len(personas_cm))
print(u'Prospectos: %d (activos %d = cierre %d + CM pendiente %d, recuperables %d, descartados %d, por validar %d)'
      % (tot_p, len(activos), len(act_cierre), len(act_cm), len(recuperables), len(descartados), len(por_validar)))
print(u'Conversion confirmada: %d/%d = %.1f%%  (techo %.1f%%)' % (len(cm_convertidas), len(cm_asistieron), conv * 100, conv_max * 100))
print(u'Registros por validar: %d' % len(val))
print(u'Conciliacion prospectos: %d = %d + %d + %d + %d' % (tot_p, len(activos), len(recuperables), len(descartados), len(por_validar)))
print()
print(u'Estatus:')
for e in ORDEN_ESTATUS:
    if cnt_est.get(e):
        print(u'  %-34s %3d' % (e, cnt_est[e]))
print()
print(u'Hallazgos por tipo:')
for t, n in collections.Counter(v[1] for v in val).most_common():
    print(u'  %-52s %3d' % (t, n))
