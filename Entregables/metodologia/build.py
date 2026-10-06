"""Construye el modelo de datos consolidado de la Campana Septiembre 2026.
Salida: diccionario M con todos los hechos y KPIs, usado por los generadores
de entregables.  No contiene ningun dato inventado: todo proviene de las
4 fuentes del repositorio.
"""
import collections, datetime, json
import core, link
from core import (WIN_INI, WIN_FIN, CORTE_OPER, PRESUPUESTO, in_window,
                  norm_name, norm_phone, nstr, cm_estado, AD_MAP)


def build():
    M = {}

    # ===================================================== FUENTES
    leads_raw, leads_oct_copia = core.load_leads()
    leads_julio = core.load_leads_julio()
    cms_all = core.load_cm()
    alumnos, rga_meta = core.load_alumnos()
    ads = core.load_ads()

    M['rga_meta'] = rga_meta
    M['n_leads_hoja'] = len(leads_raw)
    M['n_leads_oct_copia'] = len(leads_oct_copia)
    M['n_leads_julio'] = len(leads_julio)

    # --- duplicidad entre hoja "Octubre" y hoja principal
    key_main = set((r['fecha'], r['nn'], r['tel10']) for r in leads_raw)
    dup_oct = [r for r in leads_oct_copia if (r['fecha'], r['nn'], r['tel10']) in key_main]
    M['leads_oct_duplicados'] = len(dup_oct)
    M['leads_oct_detalle'] = [(r['fecha'].isoformat(), r['nombre'] or '(sin nombre)', r['tel10'])
                              for r in leads_oct_copia]

    # ===================================================== 1. LEADS
    for r in leads_raw:
        r['en_ventana'] = in_window(r['fecha'])
    M['leads_fuera_ventana'] = [r for r in leads_raw if not r['en_ventana']]

    grupos, notas, hogar = link.dedupe_leads(leads_raw)

    personas = []
    for g in grupos:
        rs = [leads_raw[i] for i in g]
        rs_sorted = sorted(rs, key=lambda r: r['fecha'])
        nombres = [r['nombre'] for r in rs if r['nombre']]
        tels = sorted(set(r['tel10'] for r in rs if r['tel10']))
        anuncios = [r['anuncio'] for r in rs if r['anuncio']]
        medios = [r['medio'] for r in rs if r['medio']]
        canales = [r['canal'] for r in rs if r['canal']]
        idiomas = [r['idioma'] for r in rs if r['idioma']]
        segs = [r['segmento'] for r in rs if r['segmento']]
        etapas = [r['etapa'] for r in rs if r['etapa']]
        p = {
            'idx': g,
            'n_registros': len(g),
            'nombre': nombres[0] if nombres else '',
            'nombres': sorted(set(nombres)),
            'nombres_norm': sorted(set(norm_name(x) for x in nombres if x)),
            'tels': tels,
            'tel': tels[0] if tels else '',
            'f_primer': rs_sorted[0]['fecha'],
            'f_ultimo': rs_sorted[-1]['fecha'],
            'anuncio': collections.Counter(anuncios).most_common(1)[0][0] if anuncios else '',
            'anuncios': sorted(set(anuncios)),
            'medio': collections.Counter(medios).most_common(1)[0][0] if medios else '',
            'canal': collections.Counter(canales).most_common(1)[0][0] if canales else '',
            'idioma': '; '.join(sorted(set(idiomas))),
            'segmento': '; '.join(sorted(set(segs))),
            'etapas': etapas,
            'etapa': etapas[-1] if etapas else '',
            'registrado_cm': 'Si' if any(norm_name(r['registrado_cm']) == 'si' for r in rs) else
                             ('En espera' if any(norm_name(r['registrado_cm']) == 'en espera' for r in rs) else
                              ('No' if any(norm_name(r['registrado_cm']) == 'no' for r in rs) else '')),
            'asistio_cm': '; '.join(sorted(set(r['asistio_cm'] for r in rs if r['asistio_cm']))),
            'resultado_crm': '; '.join(sorted(set(r['resultado'] for r in rs if r['resultado']))),
            'notas_dedup': sorted(set(sum((notas.get(i, []) for i in g), []))),
            'hogar': sorted(set(hogar[i] for i in g if i in hogar)),
            'filas': [leads_raw[i]['_fila'] for i in g],
        }
        p['en_ventana'] = in_window(p['f_primer'])
        p['sin_identificador'] = (not p['nombre']) and (not p['tels'])
        personas.append(p)
    personas.sort(key=lambda p: (p['f_primer'], p['nombre']))

    M['personas'] = personas
    M['n_personas'] = len(personas)
    M['n_personas_ventana'] = sum(1 for p in personas if p['en_ventana'])
    M['n_reg_ventana'] = sum(1 for r in leads_raw if r['en_ventana'])
    M['dup_fusionados'] = M['n_leads_hoja'] - M['n_personas']

    # ===================================================== 2. CLASES MUESTRA
    for c in cms_all:
        c['estado'] = cm_estado(c['resultado'])
        c['en_ventana'] = in_window(c['fecha'])
    cms = [c for c in cms_all if c['en_ventana']]
    M['cms_all'] = cms_all
    M['cms'] = cms
    M['cms_fuera'] = [c for c in cms_all if not c['en_ventana']]

    # dedup de personas dentro de la bitacora de CM (misma persona, 2 clases)
    cm_person = {}
    cm_groups = collections.defaultdict(list)
    for ci, c in enumerate(cms_all):
        key = None
        for k, idxs in list(cm_groups.items()):
            for j in idxs:
                o = cms_all[j]
                same_tel = c['tel10'] and o['tel10'] and c['tel10'] == o['tel10']
                if core.same_person_names(c['nn'], o['nn']) and (same_tel or not (c['tel10'] and o['tel10'])):
                    key = k
                    break
            if key:
                break
        if key is None:
            key = ci
        cm_groups[key].append(ci)
        cm_person[ci] = key
    M['cm_person'] = cm_person
    M['cm_groups'] = dict(cm_groups)

    # enlace CM -> persona lead
    cm2p = link.link_cm(personas, cms_all)
    M['cm2p'] = cm2p

    # enlace CM -> alumno padron
    cm2a = link.link_alumnos(cms_all, alumnos)
    M['cm2a'] = cm2a
    M['alumnos'] = alumnos

    # ===================================================== 3. ADS
    M['ads'] = ads
    M['gasto'] = round(sum(a['gasto'] for a in ads), 2)
    M['impresiones'] = sum(a['impresiones'] for a in ads)
    M['alcance_suma'] = sum(a['alcance'] for a in ads)
    M['alcance_max_anuncio'] = max(a['alcance'] for a in ads)
    M['conversaciones'] = sum(a['resultados'] for a in ads)
    M['presupuesto'] = PRESUPUESTO

    # ===================================================== 4. FUNNEL
    cmv = cms
    est = collections.Counter(c['estado'] for c in cmv)
    M['cm_estados'] = est
    M['cm_agendadas'] = len(cmv)
    M['cm_realizadas'] = est['Realizada - Inscrito'] + est['Realizada - Sin cierre']
    M['cm_noshow'] = est['No show']
    M['cm_pendientes'] = est['Agendada pendiente']
    M['cm_reprogramadas'] = est['Reprogramada']
    M['cm_inscritos'] = est['Realizada - Inscrito']

    # personas unicas en CM dentro de ventana
    pers_cm = set(cm_person[ci] for ci, c in enumerate(cms_all) if c['en_ventana'])
    M['cm_personas_unicas'] = len(pers_cm)
    pers_insc = set(cm_person[ci] for ci, c in enumerate(cms_all)
                    if c['en_ventana'] and c['estado'] == 'Realizada - Inscrito')
    M['insc_personas_unicas'] = len(pers_insc)
    pers_real = set(cm_person[ci] for ci, c in enumerate(cms_all)
                    if c['en_ventana'] and c['estado'].startswith('Realizada'))
    M['cm_realizadas_personas'] = len(pers_real)

    # atribucion: CM con lead de campana identificado
    def atribuido(ci):
        lk = cm2p.get(ci)
        return lk is not None and lk[2] in ('Alta', 'Media', 'Hogar')

    M['atribuido'] = atribuido
    M['cm_con_lead'] = sum(1 for ci, c in enumerate(cms_all) if c['en_ventana'] and atribuido(ci))
    M['cm_insc_con_lead'] = sum(1 for ci, c in enumerate(cms_all)
                                if c['en_ventana'] and c['estado'] == 'Realizada - Inscrito' and atribuido(ci))
    M['insc_pers_con_lead'] = len(set(cm_person[ci] for ci, c in enumerate(cms_all)
                                      if c['en_ventana'] and c['estado'] == 'Realizada - Inscrito' and atribuido(ci)))
    M['cm_pers_con_lead'] = len(set(cm_person[ci] for ci, c in enumerate(cms_all)
                                    if c['en_ventana'] and atribuido(ci)))
    M['cm_conf'] = collections.Counter((cm2p[ci][2] if ci in cm2p else 'Sin enlace')
                                       for ci, c in enumerate(cms_all) if c['en_ventana'])

    # inscritos validados contra padron RGA
    M['insc_en_padron'] = sum(1 for ci, c in enumerate(cms_all)
                              if c['en_ventana'] and c['estado'] == 'Realizada - Inscrito' and ci in cm2a)
    M['insc_pers_padron'] = len(set(cm2a[ci][0] for ci, c in enumerate(cms_all)
                                    if c['en_ventana'] and c['estado'] == 'Realizada - Inscrito' and ci in cm2a))

    # ===================================================== 5. TEMPORAL
    def etapa_temporal(d):
        if d is None:
            return 'Sin fecha'
        if d < datetime.date(2026, 9, 1):
            return '21-31 ago (arranque)'
        if d <= datetime.date(2026, 9, 30):
            return 'Septiembre (principal)'
        return '1-4 oct (cierre)'

    M['etapa_temporal'] = etapa_temporal
    tl = collections.OrderedDict()
    for k in ('21-31 ago (arranque)', 'Septiembre (principal)', '1-4 oct (cierre)'):
        tl[k] = {'leads_reg': 0, 'personas': 0, 'cm_agend': 0, 'cm_real': 0, 'insc': 0}
    for r in leads_raw:
        if r['en_ventana']:
            tl[etapa_temporal(r['fecha'])]['leads_reg'] += 1
    for p in personas:
        if p['en_ventana']:
            tl[etapa_temporal(p['f_primer'])]['personas'] += 1
    for c in cmv:
        k = etapa_temporal(c['fecha'])
        tl[k]['cm_agend'] += 1
        if c['estado'].startswith('Realizada'):
            tl[k]['cm_real'] += 1
        if c['estado'] == 'Realizada - Inscrito':
            tl[k]['insc'] += 1
    M['timeline'] = tl

    # ===================================================== 6. ANUNCIOS x CRM
    crm_por_anuncio = collections.Counter()
    crm_pers_por_anuncio = collections.Counter()
    for r in leads_raw:
        if r['en_ventana']:
            crm_por_anuncio[r['anuncio'] or '(sin anuncio registrado)'] += 1
    for p in personas:
        if p['en_ventana']:
            crm_pers_por_anuncio[p['anuncio'] or '(sin anuncio registrado)'] += 1
    M['crm_por_anuncio'] = crm_por_anuncio
    M['crm_pers_por_anuncio'] = crm_pers_por_anuncio

    # CM e inscritos por anuncio de origen (via lead enlazado)
    cm_por_anuncio = collections.Counter()
    insc_por_anuncio = collections.Counter()
    for ci, c in enumerate(cms_all):
        if not c['en_ventana'] or not atribuido(ci):
            continue
        p = personas[cm2p[ci][0]]
        a = p['anuncio'] or '(sin anuncio registrado)'
        cm_por_anuncio[a] += 1
        if c['estado'] == 'Realizada - Inscrito':
            insc_por_anuncio[a] += 1
    M['cm_por_anuncio'] = cm_por_anuncio
    M['insc_por_anuncio'] = insc_por_anuncio

    # ===================================================== 7. SEGMENTOS
    M['cm_por_categoria'] = collections.Counter(c['categoria'] or 'Sin información' for c in cmv)
    M['cm_por_idioma'] = collections.Counter(c['idioma'] or 'Sin información' for c in cmv)
    M['insc_por_categoria'] = collections.Counter(c['categoria'] or 'Sin información'
                                                  for c in cmv if c['estado'] == 'Realizada - Inscrito')
    M['insc_por_idioma'] = collections.Counter(c['idioma'] or 'Sin información'
                                               for c in cmv if c['estado'] == 'Realizada - Inscrito')
    M['cm_por_hora'] = collections.Counter(nstr(c['hora']).lower().replace(' ', '') or 'Sin información' for c in cmv)
    M['cm_por_medio_agenda'] = collections.Counter(c['medio_agenda'] or 'Sin información' for c in cmv)
    M['insc_por_medio_agenda'] = collections.Counter(c['medio_agenda'] or 'Sin información'
                                                     for c in cmv if c['estado'] == 'Realizada - Inscrito')
    # edades
    eda = [c['edad'] for c in cmv if c['edad'] is not None]
    M['edad_cm'] = {'n': len(eda), 'min': min(eda) if eda else None, 'max': max(eda) if eda else None,
                    'prom': round(sum(eda) / len(eda), 1) if eda else None}

    # ===================================================== 8. CANALES CRM
    M['leads_por_medio'] = collections.Counter((r['medio'] or 'Sin información')
                                               for r in leads_raw if r['en_ventana'])
    M['leads_por_canal'] = collections.Counter((r['canal'] or 'Sin información')
                                               for r in leads_raw if r['en_ventana'])
    M['leads_por_etapa'] = collections.Counter((r['etapa'] or 'Sin información')
                                               for r in leads_raw if r['en_ventana'])
    M['pers_por_etapa'] = collections.Counter((p['etapa'] or 'Sin información')
                                              for p in personas if p['en_ventana'])

    # serie diaria
    serie = collections.Counter()
    for r in leads_raw:
        if r['en_ventana']:
            serie[r['fecha']] += 1
    M['serie_leads'] = collections.OrderedDict(sorted(serie.items()))
    serie_cm = collections.Counter()
    for c in cmv:
        serie_cm[c['fecha']] += 1
    M['serie_cm'] = collections.OrderedDict(sorted(serie_cm.items()))

    # ===================================================== 9. CALIDAD DE DATOS
    M['dq'] = {
        'leads_sin_nombre': sum(1 for r in leads_raw if r['en_ventana'] and not r['nombre']),
        'leads_sin_tel': sum(1 for r in leads_raw if r['en_ventana'] and not r['tel10']),
        'leads_sin_nombre_ni_tel': sum(1 for r in leads_raw if r['en_ventana'] and not r['nombre'] and not r['tel10']),
        'leads_sin_idioma': sum(1 for r in leads_raw if r['en_ventana'] and not r['idioma']),
        'leads_sin_segmento': sum(1 for r in leads_raw if r['en_ventana'] and not r['segmento']),
        'leads_sin_anuncio': sum(1 for r in leads_raw if r['en_ventana'] and not r['anuncio']),
        'leads_sin_etapa': sum(1 for r in leads_raw if r['en_ventana'] and not r['etapa']),
        'leads_sin_resultado': sum(1 for r in leads_raw if r['en_ventana'] and not r['resultado']),
        'leads_comentario_vacio': sum(1 for r in leads_raw if not nstr(r['comentario'])),
        'cm_sin_tel': sum(1 for c in cmv if not c['tel10']),
        'cm_sin_categoria': sum(1 for c in cmv if not c['categoria']),
        'cm_sin_edad': sum(1 for c in cmv if c['edad'] is None),
        'cm_anio_erroneo': len(core.YEAR_TYPOS),
        'cm_anio_detalle': dict(core.YEAR_TYPOS),
        'cm_sin_lead': sum(1 for ci, c in enumerate(cms_all) if c['en_ventana'] and not atribuido(ci)),
        'leads_revisar': sum(1 for p in personas if any(n.startswith('R5') for n in p['notas_dedup'])),
        'personas_sin_id': sum(1 for p in personas if p['sin_identificador']),
    }

    # duplicados en padron de alumnos
    byname = collections.defaultdict(list)
    for a in alumnos:
        byname[a['nn']].append(a)
    M['rga_dup'] = {k: v for k, v in byname.items() if len(v) > 1}
    M['rga_activos'] = collections.Counter(a['activo'] for a in alumnos)
    M['rga_sexo'] = collections.Counter(a['sexo'] for a in alumnos)
    M['rga_n'] = len(alumnos)

    return M


if __name__ == '__main__':
    M = build()
    g = M['gasto']
    print('==== FUENTES ====')
    print('leads hoja AGOSTOSEPTIEMBRE   :', M['n_leads_hoja'], 'registros')
    print('  dentro de ventana           :', M['n_reg_ventana'])
    print('  fuera de ventana            :', [(r['fecha'].isoformat(), r['nombre']) for r in M['leads_fuera_ventana']])
    print('hoja Octubre (copia)          :', M['n_leads_oct_copia'], 'de los cuales duplicados exactos:', M['leads_oct_duplicados'])
    print('leads JULIO (otra campana)    :', M['n_leads_julio'])
    print('personas unicas (dedup)       :', M['n_personas'], ' en ventana:', M['n_personas_ventana'])
    print('registros fusionados          :', M['dup_fusionados'])
    print()
    print('==== PAUTA ====')
    print('gasto real %.2f / presupuesto %.0f' % (g, M['presupuesto']))
    print('impresiones', M['impresiones'], 'suma alcance', M['alcance_suma'], 'max alcance 1 anuncio', M['alcance_max_anuncio'])
    print('conversaciones Meta', M['conversaciones'])
    print()
    print('==== CLASES MUESTRA (en ventana) ====')
    for k, v in M['cm_estados'].most_common():
        print('  %-24s %d' % (k, v))
    print('  agendadas %d | realizadas %d | no show %d | pendientes %d | reprogramadas %d | inscritos %d'
          % (M['cm_agendadas'], M['cm_realizadas'], M['cm_noshow'], M['cm_pendientes'], M['cm_reprogramadas'], M['cm_inscritos']))
    print('  personas unicas en CM %d | personas realizadas %d | personas inscritas %d'
          % (M['cm_personas_unicas'], M['cm_realizadas_personas'], M['insc_personas_unicas']))
    print('  CM con lead de campana identificado: %d de %d' % (M['cm_con_lead'], M['cm_agendadas']))
    print('  inscritos con lead identificado (registros): %d | personas: %d' % (M['cm_insc_con_lead'], M['insc_pers_con_lead']))
    print('  inscritos validados en padron RGA: %d registros / %d alumnos' % (M['insc_en_padron'], M['insc_pers_padron']))
    print('  confianza de enlace CM->lead:', dict(M['cm_conf']))
    print()
    print('==== TEMPORAL ====')
    for k, v in M['timeline'].items():
        print('  %-24s %s' % (k, v))
    print()
    print('==== CRM x ANUNCIO ====')
    for k, v in M['crm_por_anuncio'].most_common():
        print('  %-32s reg=%-4d pers=%-4d cm=%-4d insc=%d' % (k, v, M['crm_pers_por_anuncio'][k], M['cm_por_anuncio'][k], M['insc_por_anuncio'][k]))
    print()
    print('==== CALIDAD ====')
    for k, v in M['dq'].items():
        if k != 'cm_anio_detalle':
            print('  %-28s %s' % (k, v))
    print('  anio erroneo:', M['dq']['cm_anio_detalle'])
    print('  RGA n=%d activos=%s sexo=%s' % (M['rga_n'], dict(M['rga_activos']), dict(M['rga_sexo'])))
    print('  RGA duplicados por nombre:', [(k, [a['matricula'] for a in v]) for k, v in M['rga_dup'].items()])
