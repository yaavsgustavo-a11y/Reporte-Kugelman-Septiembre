# -*- coding: utf-8 -*-
"""ENTREGABLE 3 — Presentación Ejecutiva — Campaña Septiembre 2026 (.pptx)"""
import os
import facts
import pptxw as PX
from pptxw import Pres, cm, W_SLIDE, H_SLIDE, AZUL, AZUL2, AZUL3, AZULCL, GRIS, VERDE, AMBAR, ROJO, BLANCO
from chartsx import GChart

OUT = '/projects/sandbox/entregables'
M = 1.5          # margen cm
CW = 33.87 - 2 * M
SER = ['1F3864', '2E8BC0', 'E0A030', '2E7D5B', 'B23A33', '7A5EA8', '6B7280', '46A3B0']


def mx(v, dec=2):
    return '$' + ('{:,.%df}' % dec).format(v)


def n(v):
    return '{:,}'.format(v)


def pc(v, dec=1):
    return ('{:.%df}%%' % dec).format(v * 100)


def header(s, num, titulo, sub=None):
    s.rect(0, 0, W_SLIDE, cm(0.22), fill=AZUL)
    s.text(cm(M), cm(0.75), cm(CW - 3), cm(1.25),
           [(titulo, {'size': 26, 'b': 1, 'color': AZUL})], anchor='b')
    if sub:
        s.text(cm(M), cm(2.05), cm(CW - 3), cm(0.8), [(sub, {'size': 12.5, 'color': '636363'})])
    s.text(cm(33.87 - M - 2.2), cm(0.8), cm(2.2), cm(1.0),
           [(str(num).zfill(2), {'size': 26, 'b': 1, 'color': 'D2D8E4', 'align': 'r'})], align='r')
    s.line(cm(M), cm(2.82), cm(CW), 0, color='D9D9D9', lw=12700)


def footer(s, texto):
    s.line(cm(M), cm(17.85), cm(CW), 0, color='E3E3E3', lw=9525)
    s.text(cm(M), cm(17.95), cm(CW), cm(0.7),
           [(texto, {'size': 9.5, 'color': '8C8C8C'})])


def card(s, x, y, w, h, label, valor, nota='', fill=AZULCL, color=AZUL, vsize=27):
    s.rect(x, y, w, h, fill=fill, radius=True)
    s.rect(x, y, cm(0.13), h, fill=color)
    s.text(x + cm(0.42), y + cm(0.3), w - cm(0.6), cm(0.6),
           [(label.upper(), {'size': 9.5, 'b': 1, 'color': '5E5E5E'})])
    s.text(x + cm(0.42), y + cm(0.82), w - cm(0.6), cm(1.3),
           [(valor, {'size': vsize, 'b': 1, 'color': color})])
    if nota:
        s.text(x + cm(0.42), y + h - cm(0.95), w - cm(0.6), cm(0.8),
               [(nota, {'size': 9.5, 'color': '6E6E6E'})])


def bullets(s, x, y, w, items, size=12.5, gap=1.05, color='333333'):
    for i, it in enumerate(items):
        yy = y + cm(gap * i)
        s.rect(x, yy + cm(0.16), cm(0.16), cm(0.16), fill=AZUL2)
        if isinstance(it, tuple):
            t, b = it
            s.text(x + cm(0.42), yy, w - cm(0.42), cm(gap),
                   [(t, {'size': size, 'b': 1, 'color': AZUL}), (' ' + b, {'size': size, 'color': color})])
        else:
            s.text(x + cm(0.42), yy, w - cm(0.42), cm(gap), [(it, {'size': size, 'color': color})])


# =============================================================== DIAPOSITIVAS
def d1_portada(p, f):
    s = p.slide()
    s.rect(0, 0, W_SLIDE, H_SLIDE, fill=AZUL)
    s.rect(0, cm(14.2), W_SLIDE, cm(4.85), fill='172C4D')
    s.rect(cm(M), cm(3.6), cm(2.6), cm(0.17), fill='E0A030')
    s.text(cm(M), cm(4.2), cm(24), cm(1.1), [('KUGELMAN ACADEMY', {'size': 17, 'b': 1, 'color': 'A9BEDF'})])
    s.text(cm(M), cm(5.5), cm(28), cm(2.3),
           [('CAMPAÑA SEPTIEMBRE 2026', {'size': 46, 'b': 1, 'color': BLANCO})])
    s.text(cm(M), cm(8.1), cm(28), cm(1.2),
           [('21 de agosto – 4 de octubre de 2026', {'size': 20, 'color': 'D7E0EE'})])
    s.text(cm(M), cm(9.5), cm(28), cm(1.0),
           [('Marketing, captación y resultados comerciales', {'size': 14, 'color': '9FB4D6'})])
    xs = cm(M)
    wc = cm(7.3)
    for lab, val in [('Gasto real', mx(f.gasto)), ('Leads', n(f.leads)),
                     ('Clases muestra', n(f.cm_agendadas)), ('Inscripciones', n(f.insc_personas))]:
        s.text(xs, cm(15.0), wc, cm(0.7), [(lab.upper(), {'size': 10, 'b': 1, 'color': '8FA6CC'})])
        s.text(xs, cm(15.6), wc, cm(1.5), [(val, {'size': 28, 'b': 1, 'color': BLANCO})])
        xs += wc + cm(0.6)
    s.text(cm(M), cm(17.6), cm(30), cm(0.8),
           [('Presentación ejecutiva para Dirección  ·  Todos los elementos son editables', 
             {'size': 10, 'color': '7E94B8'})])


def d2_numeros(p, f):
    s = p.slide()
    header(s, 2, 'La campaña en números', 'Periodo completo 21 ago – 4 oct 2026 · gasto real, no presupuesto')
    w = cm((CW - 3 * 0.55) / 4)
    h = cm(3.35)
    y1, y2 = cm(3.35), cm(7.1)
    datos = [
        ('Gasto real', mx(f.gasto), '%s de $7,000 asignados' % pc(f.uso_presupuesto)),
        ('Impresiones', n(f.impresiones), 'CPM %s' % mx(f.cpm)),
        ('Conversaciones', n(f.conversaciones), '%s cada una' % mx(f.costo_conversacion)),
        ('Leads (personas)', n(f.leads), 'CPL %s' % mx(f.cpl_persona)),
        ('CM agendadas', n(f.cm_agendadas), '%s de los leads' % pc(f.t_lead_cm)),
        ('CM realizadas', n(f.cm_realizadas), '%s de asistencia' % pc(f.t_cm_real)),
        ('Inscripciones', n(f.insc_personas), '%d atribuibles a campaña' % f.insc_atribuibles),
        ('CAC', mx(f.cac), 'sobre inscripciones atribuibles'),
    ]
    for i, (a, b, c) in enumerate(datos):
        x = cm(M + (i % 4) * ((CW - 3 * 0.55) / 4 + 0.55))
        y = y1 if i < 4 else y2
        col = AZUL if i < 4 else (VERDE if i >= 6 else AZUL2)
        card(s, x, y, w, h, a, b, c, fill=AZULCL if i < 4 else 'F1F5EC' if i >= 6 else AZULCL, color=col)
    s.rect(cm(M), cm(11.1), cm(CW), cm(5.9), fill='FAFAFA', radius=True)
    s.rect(cm(M), cm(11.1), cm(0.13), cm(5.9), fill=AMBAR)
    s.text(cm(M + 0.5), cm(11.45), cm(CW - 1), cm(0.8),
           [('Lectura ejecutiva', {'size': 14, 'b': 1, 'color': AZUL})])
    bullets(s, cm(M + 0.5), cm(12.3), cm(CW - 1.2), [
        ('Demanda barata.', 'Cada persona interesada costó %s. Generar interés no fue el problema.'
         % mx(f.cpl_persona)),
        ('Conversión alta al final del embudo.', '%s de asistencia a clase muestra y %s de cierre.'
         % (pc(f.t_cm_real), pc(f.t_real_insc))),
        ('Pérdida concentrada en un solo paso.', 'Solo %s de los leads llegó a clase muestra: '
         '%d personas se quedaron en el camino.' % (pc(f.t_lead_cm), f.leads - f.cm_personas)),
        ('Presupuesto ejecutado al 100%.', 'Gasto real %s de %s asignados.'
         % (mx(f.gasto), mx(f.presupuesto, 0))),
    ], size=12)
    footer(s, 'Fuente: Meta Ads (8 anuncios) · CRM de campaña deduplicado · bitácora de clases muestra · padrón de alumnos')


def d3_alcance(p, f):
    s = p.slide()
    header(s, 3, 'Alcance y generación de prospectos',
           'Qué compró la inversión: entrega publicitaria y demanda identificable')
    w = cm((CW - 3 * 0.55) / 4)
    for i, (a, b, c, col) in enumerate([
        ('Inversión', mx(f.gasto), '%s del presupuesto' % pc(f.uso_presupuesto), AZUL),
        ('Impresiones', n(f.impresiones), 'Frecuencia %.2f – %.2f' % (f.frec_sobre_suma, f.frec_sobre_min), AZUL),
        ('Alcance', '≥ ' + n(f.alcance_min), 'rango %s – %s' % (n(f.alcance_min), n(f.alcance_suma)), AZUL2),
        ('Conversaciones', n(f.conversaciones), '%s por conversación' % mx(f.costo_conversacion), AZUL),
    ]):
        card(s, cm(M + i * ((CW - 3 * 0.55) / 4 + 0.55)), cm(3.35), w, cm(3.2), a, b, c, color=col)
    s.chart(GChart('col', 'Captación por semana (leads-persona)',
                   [lab for lab, l, ca, cr, ins in f.semanas],
                   [('Leads', [l for lab, l, ca, cr, ins in f.semanas])], colors=['2E8BC0']),
            cm(M), cm(7.1), cm(CW / 2 - 0.4), cm(9.1))
    s.chart(GChart('doughnut', 'Medio de contacto de los leads',
                   [k for k, v in f.leads_medio.most_common()],
                   [('Registros', [v for k, v in f.leads_medio.most_common()])], colors=SER),
            cm(M + CW / 2 + 0.4), cm(7.1), cm(CW / 2 - 0.4), cm(9.1))
    s.text(cm(M), cm(16.5), cm(CW), cm(1.2),
           [('%d conversaciones en Meta frente a %d registros capturados en el CRM: %d conversaciones no '
             'dejaron rastro (cobertura de registro %s). El alcance único de campaña no fue exportado por '
             'Meta, por eso se reporta como rango.'
             % (f.conversaciones, f.reg_ventana, f.gap_registro, pc(f.cobertura_crm)),
             {'size': 11, 'color': '4D4D4D'})])
    footer(s, 'El alcance por anuncio no es sumable sin duplicación · CTR y CPC no son calculables: la exportación no incluye clics')


def d4_funnel(p, f):
    s = p.slide()
    header(s, 4, 'Funnel comercial', 'Leads → CM agendadas → CM realizadas → Inscritos')
    etapas = [
        ('Leads (personas)', f.leads, 1.0, '—', AZUL),
        ('CM agendadas', f.cm_agendadas, 0.60, pc(f.t_lead_cm), AZUL2),
        ('CM realizadas', f.cm_realizadas, 0.44, pc(f.t_cm_real), '46A3B0'),
        ('Inscripciones', f.insc_personas, 0.26, pc(f.t_real_insc), VERDE),
    ]
    y = cm(3.5)
    maxw = cm(17.0)
    for nm, val, rel, conv, col in etapas:
        bw = int(maxw * rel)
        s.rect(cm(M), y, bw, cm(2.3), fill=col, geom='rect')
        s.text(cm(M + 0.5), y, bw - cm(1), cm(2.3),
               [(nm, {'size': 13, 'b': 1, 'color': BLANCO})], anchor='ctr')
        s.text(cm(M) + bw - cm(4.2), y, cm(3.7), cm(2.3),
               [(n(val), {'size': 22, 'b': 1, 'color': BLANCO, 'align': 'r'})], anchor='ctr', align='r')
        s.text(cm(M) + maxw + cm(0.6), y, cm(4.2), cm(2.3),
               [(conv, {'size': 17, 'b': 1, 'color': col})], anchor='ctr')
        y += cm(2.9)
    s.text(cm(M) + maxw + cm(0.6), cm(3.0), cm(4.2), cm(0.6),
           [('CONVERSIÓN', {'size': 9.5, 'b': 1, 'color': '7A7A7A'})])
    x2 = cm(M + 23.2)
    s.rect(x2, cm(3.5), cm(CW - 23.2 + M - M), cm(13.2), fill='FAFAFA', radius=True)
    s.rect(x2, cm(3.5), cm(0.13), cm(13.2), fill=ROJO)
    s.text(x2 + cm(0.5), cm(3.85), cm(8.3), cm(0.8),
           [('Dónde se pierde', {'size': 14, 'b': 1, 'color': AZUL})])
    p1 = f.leads - f.cm_personas
    p2 = f.cm_noshow
    p3 = f.cm_pers_realizadas - f.insc_personas
    yy = cm(4.85)
    for lab, val, pct, col in [
        ('Nunca agenda clase muestra', p1, pc(p1 / f.leads), ROJO),
        ('Agenda y no asiste', p2, pc(p2 / f.leads), AMBAR),
        ('Asiste y no se inscribe', p3, pc(p3 / f.leads), '8C8C8C'),
    ]:
        s.text(x2 + cm(0.5), yy, cm(8.3), cm(0.7), [(lab, {'size': 11, 'color': '4D4D4D'})])
        s.text(x2 + cm(0.5), yy + cm(0.62), cm(5), cm(1.1),
               [(n(val), {'size': 23, 'b': 1, 'color': col}), ('  ' + pct, {'size': 12, 'color': '7A7A7A'})])
        yy += cm(2.15)
    s.text(x2 + cm(0.5), cm(11.6), cm(8.3), cm(4.7),
           [('El %s de toda la pérdida del embudo ocurre ANTES de que exista una clase muestra agendada.'
             % pc(p1 / (p1 + p2 + p3)), {'size': 12.5, 'b': 1, 'color': ROJO}), ('\n', {}),
            ('No es un problema de pauta ni de cierre: es agendamiento y seguimiento.',
             {'size': 11.5, 'color': '4D4D4D', 'before': 8})])
    footer(s, 'Lead → CM %s · CM → realizada %s · CM realizada → inscripción %s · Lead → inscripción %s'
           % (pc(f.t_lead_cm), pc(f.t_cm_real), pc(f.t_real_insc), pc(f.t_lead_insc)))


def d5_clases(p, f):
    s = p.slide()
    header(s, 5, 'Clases muestra', 'El activo comercial más rentable de la academia')
    w = cm((CW - 4 * 0.5) / 5)
    for i, (a, b, c, col) in enumerate([
        ('Agendadas', n(f.cm_agendadas), '%d personas' % f.cm_personas, AZUL),
        ('Realizadas', n(f.cm_realizadas), pc(f.t_cm_real) + ' asistencia', VERDE),
        ('No show', n(f.cm_noshow), pc(f.tasa_noshow), ROJO),
        ('Reprogramadas', n(f.cm_reprogramadas), pc(f.tasa_reprog), AMBAR),
        ('Pendientes', n(f.cm_pendientes), 'al 4 de octubre', '6B7280'),
    ]):
        card(s, cm(M + i * ((CW - 4 * 0.5) / 5 + 0.5)), cm(3.35), w, cm(3.1), a, b, c,
             fill='FAFAFA', color=col, vsize=24)
    s.chart(GChart('col', 'Agendadas → Realizadas → Inscripciones',
                   ['Agendadas', 'Realizadas', 'Con inscripción'],
                   [('Clases muestra', [f.cm_agendadas, f.cm_realizadas, f.cm_inscritos_reg])],
                   colors=['1F3864']),
            cm(M), cm(6.9), cm(14.6), cm(9.3))
    rows = [[k, n(v[0]), n(v[1]), n(v[2]), pc(v[1] / v[0]), pc(v[2] / v[1]) if v[1] else '—']
            for k, v in sorted(f.conv_cat.items(), key=lambda kv: -kv[1][0]) if k != 'Sin información']
    s.text(cm(M + 15.4), cm(6.9), cm(CW - 15.4), cm(0.8),
           [('Desempeño por programa', {'size': 13, 'b': 1, 'color': AZUL})])
    s.table(cm(M + 15.4), cm(7.7), cm(CW - 15.4), [2.3, 1.3, 1.3, 1.3, 1.3, 1.4],
            ['Programa', 'Agend.', 'Realiz.', 'Insc.', 'Asist.', 'Cierre'], rows,
            row_h=cm(0.82), font=10.5, aligns=['l', 'r', 'r', 'r', 'r', 'r'])
    s.text(cm(M + 15.4), cm(13.3), cm(CW - 15.4), cm(3.0),
           [('Diagnóstico: ', {'size': 11.5, 'b': 1, 'color': AZUL}),
            ('la asistencia (%s) y el cierre (%s) están sanos. La reprogramación es irrelevante '
             '(%d caso). El problema no está aquí.' % (pc(f.t_cm_real), pc(f.t_real_insc),
                                                       f.cm_reprogramadas),
             {'size': 11.5, 'color': '4D4D4D'})])
    footer(s, 'Horario más demandado: 16:00 h (%d clases) · %s de las clases se agendaron por WhatsApp'
           % (f.cm_hora.get('4:00pm', 0),
              pc(f.cm_medio_ag.get('Whatsaap', 0) / sum(f.cm_medio_ag.values()))))


def d6_inscripciones(p, f):
    s = p.slide()
    header(s, 6, 'Inscripciones y conversión', 'Resultado comercial del periodo y su atribución')
    w = cm((CW - 3 * 0.55) / 4)
    for i, (a, b, c, col) in enumerate([
        ('Inscripciones confirmadas', n(f.insc_personas), '%d eventos de inscripción' % f.insc_eventos, VERDE),
        ('Atribuibles a campaña', n(f.insc_atribuibles),
         pc(f.insc_atribuibles / f.insc_personas) + ' del total', AZUL),
        ('Cierre tras clase muestra', pc(f.t_real_insc), 'punto más fuerte del embudo', VERDE),
        ('CAC', mx(f.cac), '%s si se cuentan todas' % mx(f.cac_total), AZUL),
    ]):
        card(s, cm(M + i * ((CW - 3 * 0.55) / 4 + 0.55)), cm(3.35), w, cm(3.2), a, b, c,
             fill='F1F5EC' if col == VERDE else AZULCL, color=col, vsize=25)
    s.chart(GChart('doughnut', 'Inscripciones por programa',
                   [k for k, v in f.insc_cat.most_common() if k != 'Sin información'],
                   [('Inscripciones', [v for k, v in f.insc_cat.most_common() if k != 'Sin información'])],
                   colors=SER), cm(M), cm(7.0), cm(10.6), cm(9.2))
    s.chart(GChart('bar', 'Conversión general del embudo (%)',
                   ['Lead → CM', 'CM → realizada', 'CM realizada → inscripción', 'Lead → inscripción'],
                   [('Conversión', [round(f.t_lead_cm * 100, 1), round(f.t_cm_real * 100, 1),
                                    round(f.t_real_insc * 100, 1), round(f.t_lead_insc * 100, 1)])],
                   numfmt='0.0"%"', colors=['2E8BC0']), cm(M + 11.2), cm(7.0), cm(10.6), cm(9.2))
    x3 = cm(M + 22.4)
    s.rect(x3, cm(7.0), cm(CW - 22.4), cm(9.2), fill='FAFAFA', radius=True)
    s.rect(x3, cm(7.0), cm(0.13), cm(9.2), fill=AMBAR)
    s.text(x3 + cm(0.45), cm(7.3), cm(8.1), cm(0.8),
           [('De dónde vinieron', {'size': 13, 'b': 1, 'color': AZUL})])
    bullets(s, x3 + cm(0.45), cm(8.25), cm(8.3), [
        ('%d de %d' % (f.insc_atribuibles, f.insc_personas), 'con lead identificado en el CRM de campaña.'),
        ('%d' % (f.insc_personas - f.insc_atribuibles), 'sin origen de pauta: 1 recomendación explícita '
                                                        'y 6 sin registro en el CRM.'),
        ('%d' % f.insc_idioma.get('Francés', 0), 'inscripciones de francés, todas de alumnos que ya '
                                                 'habían tomado clase de inglés: venta cruzada.'),
        ('3 contactos', 'inscribieron a dos integrantes de la familia cada uno.'),
    ], size=11, gap=1.9)
    footer(s, 'Evidencia cruzada: bitácora de clases muestra + CRM + padrón de alumnos (export 05-oct-2026)')


def d7_pauta(p, f):
    s = p.slide()
    header(s, 7, 'Resultados de pauta', 'Volumen frente a eficiencia: el reparto del gasto fue el gran costo oculto')
    s.chart(GChart('col', 'Conversaciones por anuncio',
                   [t['anuncio'] for t in sorted(f.tab_ads, key=lambda r: -r['conv'])],
                   [('Conversaciones', [t['conv'] for t in sorted(f.tab_ads, key=lambda r: -r['conv'])])],
                   colors=['1F3864']), cm(M), cm(3.3), cm(15.2), cm(6.5))
    s.chart(GChart('col', 'Costo por conversación (MXN) — menor es mejor',
                   [t['anuncio'] for t in sorted(f.tab_ads, key=lambda r: r['costo_conv'])],
                   [('Costo por conversación',
                     [round(t['costo_conv'], 2) for t in sorted(f.tab_ads, key=lambda r: r['costo_conv'])])],
                   numfmt='"$"#,##0', colors=['E0A030']), cm(M), cm(10.1), cm(15.2), cm(6.5))
    x = cm(M + 15.9)
    wq = cm(CW - 15.9)
    s.text(x, cm(3.3), wq, cm(0.8), [('Los 3 anuncios que importan', {'size': 13, 'b': 1, 'color': AZUL})])
    rows = [[t['anuncio'], mx(t['gasto'], 0), pc(t['pct_gasto'], 0), n(t['conv']), mx(t['costo_conv'], 0)]
            for t in sorted(f.tab_ads, key=lambda r: -r['gasto'])]
    s.table(x, cm(4.1), wq, [2.6, 1.5, 1.0, 1.1, 1.4],
            ['Anuncio', 'Gasto', '%', 'Conv.', '$/conv'], rows, row_h=cm(0.78), font=10,
            aligns=['l', 'r', 'r', 'r', 'r'])
    s.rect(x, cm(12.4), wq, cm(4.2), fill='FFF6E5', radius=True)
    s.rect(x, cm(12.4), cm(0.13), cm(4.2), fill=AMBAR)
    s.text(x + cm(0.45), cm(12.7), wq - cm(0.8), cm(3.7),
           [('La oportunidad más fácil', {'size': 12.5, 'b': 1, 'color': AZUL}), ('\n', {}),
            ('%s (%s del gasto) se repartió entre 6 anuncios y produjo 35 conversaciones a %s cada una. '
             'Al costo del anuncio ganador (%s) ese mismo dinero habría generado ~136 conversaciones.'
             % (mx(1353.61, 0), pc(0.193, 0), mx(38.67, 0), mx(9.98)),
             {'size': 11, 'color': '4D4D4D', 'before': 6})])
    footer(s, 'Un solo conjunto de anuncios con presupuesto a nivel campaña: el reparto lo decidió el algoritmo, no una asignación manual')


def d8_creativos(p, f):
    s = p.slide()
    header(s, 8, 'Creativos y contenidos ganadores', 'Qué mensaje funcionó y por qué')
    v = f.ad_top_vol
    s.rect(cm(M), cm(3.3), cm(CW / 2 - 0.4), cm(6.4), fill='F1F5EC', radius=True)
    s.rect(cm(M), cm(3.3), cm(0.13), cm(6.4), fill=VERDE)
    s.text(cm(M + 0.5), cm(3.6), cm(CW / 2 - 1.2), cm(0.7),
           [('CREATIVO GANADOR', {'size': 10, 'b': 1, 'color': VERDE})])
    s.text(cm(M + 0.5), cm(4.3), cm(CW / 2 - 1.2), cm(1.1),
           [(v['anuncio'], {'size': 24, 'b': 1, 'color': AZUL})])
    s.text(cm(M + 0.5), cm(5.5), cm(CW / 2 - 1.2), cm(4.0),
           [('Mensaje de proximidad geográfica en formato video.', {'size': 12, 'color': '4D4D4D'}),
            ('\n', {}),
            ('%s del gasto · %s de las conversaciones · %s por conversación · único anuncio «por encima '
             'del promedio» en interacción Y conversiones a la vez.'
             % (pc(v['pct_gasto'], 0), pc(v['pct_conv'], 0), mx(v['costo_conv'])),
             {'size': 12, 'color': '4D4D4D', 'before': 6}), ('\n', {}),
            ('%d de las %d inscripciones atribuibles salieron de aquí.'
             % (f.insc_ad.get('Reel Villas de Pachuca', 0), f.insc_atribuibles),
             {'size': 12, 'b': 1, 'color': AZUL, 'before': 6})])
    s.rect(cm(M + CW / 2 + 0.4), cm(3.3), cm(CW / 2 - 0.4), cm(6.4), fill='FAFAFA', radius=True)
    s.rect(cm(M + CW / 2 + 0.4), cm(3.3), cm(0.13), cm(6.4), fill=AZUL2)
    s.text(cm(M + CW / 2 + 0.9), cm(3.6), cm(CW / 2 - 1.2), cm(0.7),
           [('SEGUNDO MEJOR', {'size': 10, 'b': 1, 'color': AZUL2})])
    s.text(cm(M + CW / 2 + 0.9), cm(4.3), cm(CW / 2 - 1.2), cm(1.1),
           [('video sin libros', {'size': 24, 'b': 1, 'color': AZUL})])
    s.text(cm(M + CW / 2 + 0.9), cm(5.5), cm(CW / 2 - 1.2), cm(4.0),
           [('Mensaje de diferenciación metodológica.', {'size': 12, 'color': '4D4D4D'}), ('\n', {}),
            ('%s por conversación, el segundo mejor de la campaña, con solo %s del gasto. '
             'Clasificación de interacción por encima del promedio.'
             % (mx(19.09), pc(0.106, 0)), {'size': 12, 'color': '4D4D4D', 'before': 6})])
    s.chart(GChart('bar', 'Eficacia del creativo: conversaciones por cada 1,000 impresiones',
                   [t['anuncio'] for t in sorted(f.tab_ads,
                                                 key=lambda r: r['conv'] / r['impresiones'])],
                   [('Conv. / 1,000 impresiones',
                     [round(t['conv'] / t['impresiones'] * 1000, 2)
                      for t in sorted(f.tab_ads, key=lambda r: r['conv'] / r['impresiones'])])],
                   numfmt='0.00', colors=['2E8BC0']), cm(M), cm(10.1), cm(16.0), cm(6.6))
    x = cm(M + 16.7)
    s.text(x, cm(10.1), cm(CW - 16.7), cm(0.8),
           [('Lo que comparten los ganadores', {'size': 13, 'b': 1, 'color': AZUL})])
    bullets(s, x, cm(11.0), cm(CW - 16.7), [
        ('Formato video/reel: el %s de las conversaciones vino de video.'
         % pc(sum(t['conv'] for t in f.tab_ads if t['formato'] == 'Video / Reel') / f.conversaciones)),
        'Mensaje ESPECÍFICO —un lugar o un método— no una oferta genérica.',
        'Clasificación de interacción por encima del promedio en Meta.',
        'Los creativos de oferta genérica con volumen comparable costaron $36.61 («post grupos») '
        'y $44.99 («carrete») por conversación.',
    ], size=11, gap=1.45)
    footer(s, 'Las piezas creativas no están en el repositorio: esta lectura se basa en el nombre del anuncio y en sus métricas de desempeño')


def d9_hallazgos(p, f):
    s = p.slide()
    header(s, 9, 'Hallazgos y áreas de oportunidad', 'Solo los puntos con mayor impacto comercial')
    s.text(cm(M), cm(3.25), cm(CW / 2 - 0.4), cm(0.8),
           [('HALLAZGOS', {'size': 12, 'b': 1, 'color': AZUL})])
    y = cm(4.1)
    for i, (t, d) in enumerate([
        ('El cuello de botella es el agendamiento',
         '%d de %d leads (%s) nunca llegaron a clase muestra. Asistencia y cierre están sanos.'
         % (f.leads - f.cm_personas, f.leads, pc((f.leads - f.cm_personas) / f.leads))),
        ('Un creativo sostuvo la campaña',
         '«video villas»: %s del gasto, %s de las conversaciones, %d de las %d inscripciones atribuibles.'
         % (pc(0.700, 0), pc(0.869, 0), f.insc_ad.get('Reel Villas de Pachuca', 0), f.insc_atribuibles)),
        ('La clase muestra convierte muy bien',
         '%s de asistencia y %s de cierre. Cada clase realizada costó %s de pauta.'
         % (pc(f.t_cm_real), pc(f.t_real_insc), mx(f.costo_cm_realizada))),
        ('Hay inventario pagado sin trabajar',
         '%d oportunidades abiertas; %d en descubrimiento con 15+ días sin movimiento.'
         % (f.activos, f.dormidos)),
    ], start=1):
        s.rect(cm(M), y, cm(CW / 2 - 0.4), cm(3.0), fill='FAFAFA', radius=True)
        s.rect(cm(M), y, cm(0.13), cm(3.0), fill=AZUL)
        s.text(cm(M + 0.45), y + cm(0.28), cm(0.9), cm(0.8),
               [(str(i), {'size': 17, 'b': 1, 'color': 'C3CBDC'})])
        s.text(cm(M + 1.35), y + cm(0.3), cm(CW / 2 - 2.1), cm(0.9),
               [(t, {'size': 12.5, 'b': 1, 'color': AZUL})])
        s.text(cm(M + 1.35), y + cm(1.2), cm(CW / 2 - 2.1), cm(1.7),
               [(d, {'size': 11, 'color': '4D4D4D'})])
        y += cm(3.2)
    x = cm(M + CW / 2 + 0.4)
    s.text(x, cm(3.25), cm(CW / 2 - 0.4), cm(0.8),
           [('ÁREAS DE OPORTUNIDAD', {'size': 12, 'b': 1, 'color': ROJO})])
    y = cm(4.1)
    for t, d, col in [
        ('Agendamiento bajo', 'Lead → CM %s. Es la pérdida más grande y la de mayor retorno si se corrige.'
         % pc(f.t_lead_cm), ROJO),
        ('Seguimiento sin registro', '%s de los registros del CRM sin resultado de seguimiento; '
                                     'columna de comentarios vacía en el 100%% de la hoja.'
         % pc(f.dq['leads_sin_resultado'] / f.reg_ventana), ROJO),
        ('Gasto en anuncios ineficientes', '%s del gasto en 6 anuncios 3.9 veces más caros que el ganador.'
         % pc(0.193, 0), AMBAR),
        ('No show del %s' % pc(f.tasa_noshow), '%d clases perdidas; no hay recordatorios ni '
                                               'confirmaciones registradas.' % f.cm_noshow, AMBAR),
        ('Trazabilidad incompleta', '%d conversaciones de Meta sin registro en el CRM (cobertura %s).'
         % (f.gap_registro, pc(f.cobertura_crm)), AMBAR),
    ]:
        s.rect(x, y, cm(CW / 2 - 0.4), cm(2.4), fill='FDF7F6', radius=True)
        s.rect(x, y, cm(0.13), cm(2.4), fill=col)
        s.text(x + cm(0.45), y + cm(0.25), cm(CW / 2 - 1.2), cm(0.8),
               [(t, {'size': 12.5, 'b': 1, 'color': AZUL})])
        s.text(x + cm(0.45), y + cm(1.0), cm(CW / 2 - 1.2), cm(1.3),
               [(d, {'size': 11, 'color': '4D4D4D'})])
        y += cm(2.6)
    footer(s, 'Cada hallazgo y cada área de oportunidad está sustentada en los datos de las cuatro fuentes · detalle en el Reporte de Campaña')


def d10_acciones(p, f):
    s = p.slide()
    header(s, 10, 'Aprendizajes y siguientes acciones',
           'Qué repetir, qué modificar y en qué orden')
    s.rect(cm(M), cm(3.3), cm(CW / 2 - 0.4), cm(5.2), fill='F1F5EC', radius=True)
    s.rect(cm(M), cm(3.3), cm(0.13), cm(5.2), fill=VERDE)
    s.text(cm(M + 0.5), cm(3.6), cm(CW / 2 - 1.2), cm(0.8),
           [('QUÉ REPETIR', {'size': 12, 'b': 1, 'color': VERDE})])
    bullets(s, cm(M + 0.5), cm(4.5), cm(CW / 2 - 1.2), [
        'El creativo de proximidad geográfica en formato video.',
        'El objetivo de mensajes con WhatsApp como canal operativo.',
        'El formato de la clase muestra tal como está hoy.',
        'La ejecución total del presupuesto asignado.',
    ], size=11.5, gap=0.95)
    s.rect(cm(M + CW / 2 + 0.4), cm(3.3), cm(CW / 2 - 0.4), cm(5.2), fill='FDF7F6', radius=True)
    s.rect(cm(M + CW / 2 + 0.4), cm(3.3), cm(0.13), cm(5.2), fill=ROJO)
    s.text(cm(M + CW / 2 + 0.9), cm(3.6), cm(CW / 2 - 1.2), cm(0.8),
           [('QUÉ MODIFICAR', {'size': 12, 'b': 1, 'color': ROJO})])
    bullets(s, cm(M + CW / 2 + 0.9), cm(4.5), cm(CW / 2 - 1.2), [
        'Proponer horarios de clase muestra en el primer mensaje, no enviar información.',
        'Cadencia de seguimiento con registro obligatorio en el CRM.',
        'Dejar de repartir presupuesto entre anuncios ya probados como ineficientes.',
        'Confirmar asistencia 24 h y 2 h antes de cada clase.',
    ], size=11.5, gap=0.95)
    s.text(cm(M), cm(8.9), cm(CW), cm(0.8),
           [('PRIORIDADES PARA LA SIGUIENTE CAMPAÑA', {'size': 12, 'b': 1, 'color': AZUL})])
    rows = [
        [('Alta', {'b': 1, 'color': ROJO}), 'Responder y agendar en menos de 30 minutos, con responsable por turno',
         'Lead → CM de %s' % pc(f.t_lead_cm), 'Lead → CM agendada'],
        [('Alta', {'b': 1, 'color': ROJO}), 'Primer mensaje con dos horarios concretos en lugar de información',
         '%d registros detenidos en «información enviada» o «contactado»'
         % (f.M['leads_por_etapa'].get('Información enviada', 0)
            + f.M['leads_por_etapa'].get('Contactado', 0)), 'Lead → CM agendada'],
        [('Alta', {'b': 1, 'color': ROJO}), 'Recuperar los %d prospectos sin movimiento antes de abrir presupuesto nuevo'
         % f.dormidos, '%d oportunidades abiertas' % f.activos, 'Inscripciones sin costo incremental'],
        [('Alta', {'b': 1, 'color': ROJO}), 'Reasignar el gasto de los 6 anuncios ineficientes al creativo ganador',
         '%s del gasto produjo %s de las conversaciones' % (pc(0.193, 0), pc(0.062, 0)), 'CPL'],
        [('Media', {'b': 1, 'color': AMBAR}), 'Confirmación de asistencia 24 h y 2 h antes de la clase',
         'No show %s' % pc(f.tasa_noshow), 'Tasa de asistencia'],
        [('Media', {'b': 1, 'color': AMBAR}), 'Dos creativos nuevos de mensaje familiar (hermanos, madre e hijo)',
         'Kids cierra %s frente a %s de College'
         % (pc(f.conv_cat['Kids'][2] / f.conv_cat['Kids'][1]),
            pc(f.conv_cat['College'][2] / f.conv_cat['College'][1])), 'Inscripciones por CM'],
    ]
    s.table(cm(M), cm(9.7), cm(CW), [1.3, 8.0, 6.5, 3.6],
            ['Prioridad', 'Acción', 'Problema que resuelve (dato)', 'KPI afectado'], rows,
            row_h=cm(0.88), font=10.5, aligns=['c', 'l', 'l', 'l'])
    s.rect(cm(M), cm(16.0), cm(CW), cm(1.55), fill=AZUL, radius=True)
    s.text(cm(M + 0.5), cm(16.0), cm(CW - 1), cm(1.55),
           [('Si solo se puede hacer una cosa: atacar el agendamiento. ',
             {'size': 13, 'b': 1, 'color': BLANCO}),
            ('Con el mismo presupuesto de %s, llevar Lead → CM de %s a 25%% produciría ~10 inscripciones '
             'más (+%s), porque asistencia y cierre ya funcionan.'
             % (mx(f.presupuesto, 0), pc(f.t_lead_cm), pc(10.5 / f.insc_personas)),
             {'size': 12.5, 'color': 'CFDAEC'})], anchor='ctr')
    footer(s, 'Kugelman Academy · Campaña Septiembre 2026 · 21 ago – 4 oct · Detalle completo en el Reporte de Campaña y la Base Maestra')


def build():
    f = facts.get()
    p = Pres('Presentación Ejecutiva — Campaña Septiembre 2026',
             subject='Kugelman Academy · 21 ago – 4 oct 2026')
    d1_portada(p, f)
    d2_numeros(p, f)
    d3_alcance(p, f)
    d4_funnel(p, f)
    d5_clases(p, f)
    d6_inscripciones(p, f)
    d7_pauta(p, f)
    d8_creativos(p, f)
    d9_hallazgos(p, f)
    d10_acciones(p, f)
    path = os.path.join(OUT, 'Presentación Ejecutiva — Campaña Septiembre 2026.pptx')
    p.save(path)
    return path


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    print(build())
