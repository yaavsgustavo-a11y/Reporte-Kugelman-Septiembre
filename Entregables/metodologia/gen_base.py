# -*- coding: utf-8 -*-
"""ENTREGABLE 2 — Base Maestra Campaña Septiembre 2026 (.xlsx)"""
import collections, datetime, os
import facts, core, xlsxw as X
from xlsxw import Sheet, Chart, C, F

OUT = '/projects/sandbox/entregables'
PER = '21 de agosto – 4 de octubre de 2026'


def q(sheet, ref):
    return "'%s'!%s" % (sheet, ref)


def hdr(sh, cols, widths):
    sh.cols = widths
    sh.row(*[C(c, 'hdr') for c in cols])


def titulo(sh, t, sub=None, ancho=8):
    sh.row(C(t, 'titulo'))
    if sub:
        sh.row(C(sub, 'nota'))
    sh.blank()


def build():
    f = facts.get()
    sheets = []

    # ============================================================ 1. PORTADA
    s = Sheet('Portada', cols=[34, 62], tab_color=X.AZUL)
    s.row(C('BASE MAESTRA — CAMPAÑA SEPTIEMBRE 2026', 'titulo'))
    s.row(C('Kugelman Academy  ·  Marketing, captación y resultados comerciales', 'subtitulo'))
    s.blank()
    for k, v in [
        ('Periodo de campaña', PER),
        ('Presupuesto asignado', '$7,000.00 MXN'),
        ('Gasto real (Meta Ads)', '$%s MXN (%.1f%% del presupuesto)' % ('{:,.2f}'.format(f.gasto), f.uso_presupuesto * 100)),
        ('Corte de bases operativas', 'sábado 3 de octubre de 2026 (archivos 4 y 5, independientes)'),
        ('Fuentes utilizadas', '4 archivos del repositorio (ver pestaña «Fuentes»)'),
        ('Regla de atribución', 'Publicidad → Lead → Seguimiento → Clase muestra → Inscripción'),
    ]:
        s.row(C(k, 'bold'), C(v, 'txt'))
    s.blank()
    s.row(C('CÓMO ESTÁ ORGANIZADA ESTA BASE', 'seccion'), C('', 'seccion'))
    for k, v in [
        ('Resumen KPI', 'Tablero de indicadores de la campaña con fórmulas vivas a las demás pestañas.'),
        ('Prospectos', 'Una fila por PERSONA (deduplicado). Es la base de conteo de leads.'),
        ('Prospectos (registros)', 'Los 504 registros originales del CRM con su ID de persona, para auditoría.'),
        ('Clases muestra', 'Los 107 registros de la bitácora (ago/sep/oct) con estado homologado y enlace al lead.'),
        ('Inscritos campaña', 'Una fila por evento de inscripción con su nivel de evidencia y matrícula.'),
        ('Meta Ads', 'Exportación de anuncios con métricas de eficiencia calculadas por fórmula.'),
        ('Redes sociales', 'Declaración de disponibilidad: el repositorio NO contiene métricas orgánicas.'),
        ('Funnel comercial', 'Embudo completo con conversiones, pérdida absoluta y pérdida porcentual.'),
        ('Análisis temporal', 'Evolución semanal y por fase (arranque / septiembre / cierre).'),
        ('Segmentos', 'Clases muestra e inscripciones por programa, idioma, horario y canal de agenda.'),
        ('Gráficas', 'Tablas de datos + 10 gráficos nativos editables ligados a esas tablas.'),
        ('Fuentes', 'Trazabilidad: qué archivo alimenta cada grupo de indicadores.'),
        ('Calidad de datos', 'Faltantes, duplicados, inconsistencias y conflictos entre fuentes.'),
    ]:
        s.row(C(k, 'bold'), C(v, 'txtw'))
    s.blank()
    s.row(C('REGLA DE ORO DE ESTE ARCHIVO', 'bold'))
    s.row(C('Ninguna celda contiene datos estimados o inventados. Cuando un dato no existe en las fuentes se '
            'escribe "Sin información" o "No calculable con los datos disponibles". Las inferencias están '
            'marcadas como tales.', 'txtw'))
    s.merge('B%d:F%d' % (len(s.rows), len(s.rows)))
    sheets.append(s)

    # ========================================================= 2. RESUMEN KPI
    s = Sheet('Resumen KPI', cols=[44, 18, 16, 54], freeze='A5', tab_color=X.AZUL2)
    s.row(C('RESUMEN KPI — CAMPAÑA SEPTIEMBRE 2026', 'titulo'))
    s.row(C('Periodo: %s  ·  Todas las cifras provienen de las 4 fuentes del repositorio' % PER, 'nota'))
    s.blank()
    s.row(C('Indicador', 'hdr'), C('Campaña Septiembre', 'hdr'), C('Unidad', 'hdr'), C('Fuente / nota', 'hdr'))
    kpi_rows = [
        ('INVERSIÓN', None, None, None),
        ('Presupuesto asignado', f.presupuesto, 'MXN', 'Dato proporcionado por Dirección', 'money'),
        ('Gasto real', f.gasto, 'MXN', 'Meta Ads — suma de "Importe gastado" de los 8 anuncios', 'money'),
        ('% de presupuesto utilizado', f.uso_presupuesto, '%', '=Gasto real / Presupuesto asignado', 'pct'),
        ('ALCANCE Y ENTREGA', None, None, None),
        ('Impresiones', f.impresiones, 'impresiones', 'Meta Ads — suma por anuncio', 'num'),
        ('Alcance (cota mínima)', f.alcance_min, 'personas',
         'Meta no reporta alcance único de campaña; es el alcance del anuncio mayor', 'num'),
        ('Alcance (suma por anuncio, con duplicación)', f.alcance_suma, 'personas',
         'Suma de alcance de los 8 anuncios: sobreestima por traslape entre anuncios', 'num'),
        ('Frecuencia', f.frec_sobre_suma, 'veces', 'Impresiones / suma de alcance (cota inferior)', 'dec2'),
        ('CPM', f.cpm, 'MXN', 'Gasto real / impresiones × 1,000', 'money'),
        ('CTR / CPC', 'No calculable con los datos disponibles', '',
         'La exportación de Meta no incluye clics ni enlaces', 'txt'),
        ('GENERACIÓN DE DEMANDA', None, None, None),
        ('Conversaciones iniciadas (Meta)', f.conversaciones, 'conversaciones',
         'Meta Ads — resultado optimizado del objetivo (mensajes)', 'num'),
        ('Registros capturados en el CRM', f.reg_ventana, 'registros',
         'Hoja AGOSTOSEPTIEMBRE, fecha dentro de la ventana', 'num'),
        ('Leads / prospectos (personas únicas)', f.leads, 'personas',
         'CRM deduplicado — CIFRA OFICIAL DE LEADS', 'num'),
        ('Registros duplicados fusionados', f.dup_fusionados, 'registros',
         '%.1f%% de los 504 registros de la hoja' % (f.tasa_duplicidad * 100), 'num'),
        ('Cobertura de registro (CRM / Meta)', f.cobertura_crm, '%',
         '%d conversaciones de Meta sin registro equivalente en el CRM' % f.gap_registro, 'pct'),
        ('CLASES MUESTRA', None, None, None),
        ('CM solicitadas', 'No disponible de forma confiable', '',
         'No existe campo de solicitud; el CRM marca %d personas como registradas a CM, '
         'muy por debajo de las %d de la bitácora' % (f.cm_solicitadas_crm, f.cm_agendadas), 'txt'),
        ('CM agendadas', f.cm_agendadas, 'clases', 'Bitácora de clases muestra, fecha en ventana', 'num'),
        ('CM realizadas', f.cm_realizadas, 'clases', 'Resultado "Inscrito" o "Show"', 'num'),
        ('CM no show', f.cm_noshow, 'clases', 'Resultado "NO ASISTIÓ"', 'num'),
        ('CM reprogramadas', f.cm_reprogramadas, 'clases', 'Resultado "Cita pospuesta"', 'num'),
        ('CM pendientes al 4-oct', f.cm_pendientes, 'clases', 'Resultado "Cita" con fecha ≤ 4-oct', 'num'),
        ('Personas distintas con CM', f.cm_personas, 'personas', 'Bitácora deduplicada', 'num'),
        ('RESULTADO COMERCIAL', None, None, None),
        ('Inscripciones confirmadas en el periodo', f.insc_personas, 'personas',
         '%d eventos de inscripción (2 personas se inscribieron en 2 idiomas)' % f.insc_eventos, 'num'),
        ('Inscripciones atribuibles a la campaña', f.insc_atribuibles, 'personas',
         'Con lead identificado en el CRM de campaña — CIFRA OFICIAL DE ATRIBUCIÓN', 'num'),
        ('Inscripciones sin origen de campaña', f.insc_personas - f.insc_atribuibles, 'personas',
         'Recomendación, sitio o sin registro en el CRM', 'num'),
        ('CONVERSIÓN', None, None, None),
        ('Lead → CM agendada', f.t_lead_cm, '%', 'CM agendadas / leads', 'pct'),
        ('CM agendada → CM realizada (asistencia)', f.t_cm_real, '%', 'CM realizadas / CM agendadas', 'pct'),
        ('CM realizada → Inscripción (cierre)', f.t_real_insc, '%',
         '%d personas inscritas / %d personas con CM realizada' % (f.insc_personas, f.cm_pers_realizadas), 'pct'),
        ('Lead → Inscripción (total periodo)', f.t_lead_insc, '%', 'Inscripciones / leads', 'pct'),
        ('Lead → Inscripción (solo atribuibles)', f.t_lead_insc_atrib, '%',
         'Inscripciones atribuibles / leads', 'pct'),
        ('No show rate', f.tasa_noshow, '%', 'No show / CM agendadas', 'pct'),
        ('Tasa de reprogramación', f.tasa_reprog, '%', 'Reprogramadas / CM agendadas', 'pct'),
        ('COSTOS', None, None, None),
        ('Costo por conversación (Meta)', f.costo_conversacion, 'MXN', 'Gasto real / conversaciones', 'money'),
        ('Costo por lead (persona única)', f.cpl_persona, 'MXN', 'Gasto real / leads', 'money'),
        ('Costo por CM agendada', f.costo_cm_agendada, 'MXN', 'Gasto real / CM agendadas', 'money'),
        ('Costo por CM realizada', f.costo_cm_realizada, 'MXN', 'Gasto real / CM realizadas', 'money'),
        ('CAC (inscripciones atribuibles)', f.cac, 'MXN', 'Gasto real / %d inscripciones atribuibles' % f.insc_atribuibles, 'money'),
        ('CAC (todas las inscripciones del periodo)', f.cac_total, 'MXN',
         'Gasto real / %d inscripciones' % f.insc_personas, 'money'),
    ]
    for r in kpi_rows:
        if r[1] is None:
            s.row(C(r[0], 'seccion'), C('', 'seccion'), C('', 'seccion'), C('', 'seccion'))
        else:
            lab, val, uni, nota, st = r
            s.row(C(lab, 'txt'), C(val, st if st != 'txt' else 'txtw'), C(uni, 'txt'), C(nota, 'txtw'))
    s.blank()
    s.row(C('Nota de método: "leads" se cuenta a nivel de PERSONA (CRM deduplicado). Las conversaciones de '
            'Meta (%d) son eventos de mensaje y no personas únicas; por eso ambas cifras no coinciden.'
            % f.conversaciones, 'nota'))
    sheets.append(s)

    # ======================================================== 3. PROSPECTOS
    s = Sheet('Prospectos', cols=[9, 30, 14, 13, 13, 13, 24, 14, 20, 26, 11, 11, 26, 38],
              freeze='A5', tab_color=X.AZUL2)
    s.row(C('PROSPECTOS DE LA CAMPAÑA — UNA FILA POR PERSONA (CRM deduplicado)', 'titulo'))
    s.row(C('Base de conteo de leads. %d personas totales, %d con primer contacto dentro de la ventana %s'
            % (f.pers_total, f.leads, PER), 'nota'))
    s.blank()
    cols = ['ID persona', 'Nombre del interesado', 'Teléfono', 'Primer contacto', 'Último registro',
            'Medio de contacto', 'Anuncio de contacto', 'Canal seguimiento', 'Idioma de interés',
            'Segmento / programa', 'En ventana', '# registros', 'Etapa CRM (último)', 'Notas de deduplicación']
    s.row(*[C(c, 'hdr') for c in cols])
    for i, p in enumerate(f.P):
        s.row(C('P%03d' % (i + 1), 'txt'), C(p['nombre'] or '(sin nombre registrado)', 'txt'),
              C(p['tel'] or 'Sin información', 'txt'), C(p['f_primer'], 'fecha'), C(p['f_ultimo'], 'fecha'),
              C(p['medio'] or 'Sin información', 'txt'), C(p['anuncio'] or 'Sin anuncio registrado', 'txt'),
              C(p['canal'] or 'Sin información', 'txt'), C(p['idioma'] or 'Sin información', 'txt'),
              C(p['segmento'] or 'Sin información', 'txt'),
              C('Sí' if p['en_ventana'] else 'No (19–20 ago)', 'txt'), C(p['n_registros'], 'num'),
              C(p['etapa'] or 'Sin información', 'txt'),
              C('; '.join(p['notas_dedup']) if p['notas_dedup'] else '', 'txtw'))
    s.autofilter = 'A4:N%d' % len(s.rows)
    sheets.append(s)

    # ============================================= 4. PROSPECTOS (REGISTROS)
    s = Sheet('Prospectos (registros)', cols=[8, 9, 12, 28, 14, 12, 12, 22, 13, 18, 24, 13, 12, 16, 10],
              freeze='A5', tab_color=X.AZUL3)
    s.row(C('REGISTROS ORIGINALES DEL CRM DE CAMPAÑA — PARA AUDITORÍA', 'titulo'))
    s.row(C('Los %d registros de la hoja «AGOSTOSEPTIEMBRE» del archivo CAMPAÑAS (2).xlsx, '
            'con el ID de persona asignado por la deduplicación.' % f.reg_hoja, 'nota'))
    s.blank()
    cols = ['Fila orig.', 'ID persona', 'Fecha', 'Nombre del interesado', 'Teléfono', 'Idioma',
            'Medio', 'Anuncio de contacto', 'Canal', 'Segmento', 'Etapa del proceso',
            'Registrado a CM', 'Asistió a CM', 'Resultado seguimiento', 'En ventana']
    s.row(*[C(c, 'hdr') for c in cols])
    leads_raw, _ = core.load_leads()
    row2p = {}
    for i, p in enumerate(f.P):
        for fl in p['filas']:
            row2p[fl] = 'P%03d' % (i + 1)
    for r in leads_raw:
        s.row(C(r['_fila'], 'num'), C(row2p.get(r['_fila'], ''), 'txt'), C(r['fecha'], 'fecha'),
              C(r['nombre'] or '(sin nombre)', 'txt'), C(r['tel10'] or '', 'txt'), C(r['idioma'], 'txt'),
              C(r['medio'], 'txt'), C(r['anuncio'] or 'Sin anuncio registrado', 'txt'), C(r['canal'], 'txt'),
              C(r['segmento'], 'txt'), C(r['etapa'] or 'Sin información', 'txt'),
              C(r['registrado_cm'], 'txt'), C(r['asistio_cm'], 'txt'), C(r['resultado'], 'txt'),
              C('Sí' if core.in_window(r['fecha']) else 'No', 'txt'))
    s.autofilter = 'A4:O%d' % len(s.rows)
    sheets.append(s)

    # ===================================================== 5. CLASES MUESTRA
    s = Sheet('Clases muestra', cols=[8, 12, 12, 34, 13, 13, 7, 10, 18, 24, 11, 28, 12, 10, 26],
              freeze='A5', tab_color=X.AZUL2)
    s.row(C('BITÁCORA DE CLASES MUESTRA — ESTADO HOMOLOGADO Y ENLACE AL LEAD', 'titulo'))
    s.row(C('Hojas AGOSTO, SEPTIEMBRE y OCTUBRE del archivo «CLASES MUESTRA 2026__ (1).xlsx». '
            '%d registros totales, %d con fecha dentro de la ventana de campaña.'
            % (len(f.cms), f.cm_agendadas), 'nota'))
    s.blank()
    cols = ['Hoja', 'Fila orig.', 'Fecha CM', 'Nombre', 'Teléfono', 'Categoría', 'Edad', 'Idioma',
            'Resultado original', 'Estado homologado', 'En ventana', 'Lead de origen (CRM)',
            'Confianza enlace', 'Hora', 'Medio para agendar / nota']
    s.row(*[C(c, 'hdr') for c in cols])
    for ci, c in enumerate(f.cms):
        lk = f.cm2p.get(ci)
        s.row(C(c['_hoja'], 'txt'), C(c['_fila'], 'num'), C(c['fecha'], 'fecha'), C(c['nombre'], 'txt'),
              C(c['tel10'] or 'Sin información', 'txt'), C(c['categoria'] or 'Sin información', 'txt'),
              C(c['edad'], 'num'), C(c['idioma'] or 'Sin información', 'txt'),
              C(c['resultado'] or 'Sin información', 'txt'), C(c['estado'], 'txt'),
              C('Sí' if c['en_ventana'] else 'No', 'txt'),
              C(f.P[lk[0]]['nombre'] if lk else 'Sin lead identificado', 'txt'),
              C(lk[2] if lk else 'Sin enlace', 'txt'), C(c['hora'] or '', 'txt'),
              C(' / '.join(x for x in [c['medio_agenda'], c['nota']] if x) or '', 'txtw'))
    s.autofilter = 'A4:O%d' % len(s.rows)
    s.blank()
    s.row(C('Homologación aplicada: Inscrito → «Realizada - Inscrito»; Show → «Realizada - Sin cierre»; '
            'NO ASISTIÓ → «No show»; Cita → «Agendada pendiente»; Cita pospuesta → «Reprogramada».', 'nota'))
    s.row(C('Confianza del enlace: Alta = teléfono + nombre · Media = nombre completo · '
            'Hogar = mismo teléfono con nombre distinto (el lead es el contacto familiar) · '
            'Baja = solo un nombre de pila (NO se usa para atribución) · Sin enlace = no localizado en el CRM.', 'nota'))
    sheets.append(s)

    # ================================================= 6. INSCRITOS CAMPAÑA
    s = Sheet('Inscritos campaña', cols=[30, 12, 13, 9, 7, 10, 18, 12, 32, 7, 26, 11, 11, 46],
              freeze='A6', tab_color=X.VERDE)
    s.row(C('INSCRIPCIONES DEL PERIODO Y SU ATRIBUCIÓN A LA CAMPAÑA', 'titulo'))
    s.row(C('%d eventos de inscripción = %d personas distintas. Atribuibles a la campaña: %d personas.'
            % (f.insc_eventos, f.insc_personas, f.insc_atribuibles), 'nota'))
    s.row(C('Una misma persona puede aparecer 2 veces si se inscribió en 2 idiomas (inglés y francés) '
            'y tiene 2 matrículas en el padrón.', 'nota'))
    s.blank()
    cols = ['Alumno / interesado', 'Fecha CM', 'Programa', 'Idioma', 'Edad', 'Hora CM',
            'Medio para agendar', 'Matrícula', 'Nombre en padrón RGA', 'Edad padrón',
            'Lead de origen (CRM)', 'Confianza', 'Atribuible', 'Evidencia utilizada']
    s.row(*[C(c, 'hdr') for c in cols])
    for fi in f.insc_filas:
        s.row(C(fi['nombre'], 'txt'), C(fi['fecha_cm'], 'fecha'), C(fi['categoria'] or 'Sin información', 'txt'),
              C(fi['idioma'], 'txt'), C(fi['edad'], 'num'), C(fi['hora'] or '', 'txt'),
              C(fi['medio_agenda'] or 'Sin información', 'txt'),
              C(fi['matricula'] if fi['matricula'] else 'No localizado', 'txt'),
              C(fi['nombre_padron'] or 'No localizado', 'txt'), C(fi['edad_padron'], 'num'),
              C(fi['lead'] or 'Sin lead identificado', 'txt'), C(fi['confianza'], 'txt'),
              C('Sí' if fi['atribuible'] else 'No', 'ok' if fi['atribuible'] else 'warn'),
              C(fi['evidencia'], 'txtw'))
    s.autofilter = 'A5:N%d' % len(s.rows)
    s.blank()
    s.row(C('INSCRIPCIONES POR VALIDAR (NO contabilizadas)', 'seccion'))
    s.row(*[C(c, 'hdr2') for c in ['Nombre', 'Fecha CM', 'Estado en bitácora', 'Matrícula padrón',
                                   'Lead', 'Observación']])
    for r in f.insc_por_validar:
        s.row(C(r['nombre'], 'txt'), C(r['fecha_cm'], 'fecha'), C(r['estado_cm'], 'txt'),
              C(r['matricula'], 'txt'), C(r['lead'] or 'Sin lead', 'txt'), C(r['evidencia'], 'txtw'))
    s.blank()
    s.row(C('SEGMENTACIÓN DE LAS INSCRIPCIONES', 'seccion'))
    s.row(C('Por programa', 'hdr2'), C('Inscripciones', 'hdr2'), C('', 'hdr2'),
          C('Por idioma', 'hdr2'), C('Inscripciones', 'hdr2'), C('', 'hdr2'),
          C('Por canal de agenda', 'hdr2'), C('Inscripciones', 'hdr2'))
    cat = f.insc_cat.most_common()
    idi = f.insc_idioma.most_common()
    med = f.insc_medio_ag.most_common()
    for i in range(max(len(cat), len(idi), len(med))):
        row = []
        row += [C(cat[i][0], 'txt'), C(cat[i][1], 'num')] if i < len(cat) else [C(''), C('')]
        row += [C('')]
        row += [C(idi[i][0], 'txt'), C(idi[i][1], 'num')] if i < len(idi) else [C(''), C('')]
        row += [C('')]
        row += [C(med[i][0], 'txt'), C(med[i][1], 'num')] if i < len(med) else [C(''), C('')]
        s.row(*row)
    sheets.append(s)

    # ============================================================ 7. META ADS
    s = Sheet('Meta Ads', cols=[20, 26, 13, 12, 10, 13, 12, 12, 11, 10, 9, 10, 9, 8, 16, 18, 18, 17],
              freeze='A6', tab_color=X.AMBAR)
    s.row(C('RESULTADOS DE PAUTA — META ADS (NIVEL ANUNCIO)', 'titulo'))
    s.row(C('Archivo: Kugelman-Academy-Anuncios-21-ago-2026---4-oct-2026.csv  ·  '
            'Conjunto de anuncios: «campaña agosto A»  ·  Ventana del informe: 2026-08-21 a 2026-10-04', 'nota'))
    s.row(C('Indicador de resultado reportado por Meta: conversaciones de mensaje iniciadas '
            '(atribución 7 días tras clic / 1 día tras visualización).', 'nota'))
    s.blank()
    cols = ['Anuncio (Meta)', 'Etiqueta en el CRM', 'Formato', 'Gasto (MXN)', '% gasto',
            'Conversaciones', '% conv.', 'Costo/conv.', 'Impresiones', 'Alcance', 'Frec.', 'CPM',
            'Leads CRM', 'CM', 'Inscripciones atrib.', 'Clasif. calidad',
            'Clasif. interacción', 'Clasif. conversión']
    s.row(*[C(c, 'hdr') for c in cols])
    r0 = len(s.rows) + 1
    for t in f.tab_ads:
        r = len(s.rows) + 1
        s.row(C(t['anuncio'], 'txt'), C(t['etiqueta'] or 'Sin registros en CRM', 'txt'),
              C(t['formato'], 'txt'), C(t['gasto'], 'money'),
              F('=D%d/$D$%d' % (r, r0 + len(f.tab_ads)), 'pct'),
              C(t['conv'], 'num'), F('=F%d/$F$%d' % (r, r0 + len(f.tab_ads)), 'pct'),
              F('=IF(F%d=0,"",D%d/F%d)' % (r, r, r), 'money'), C(t['impresiones'], 'num'),
              C(t['alcance'], 'num'), F('=I%d/J%d' % (r, r), 'dec2'),
              F('=D%d/I%d*1000' % (r, r), 'money'),
              C(t['crm_pers'], 'num'), C(t['cm'], 'num'), C(t['insc'], 'num'),
              C(t['calidad'] if t['calidad'] != '-' else 'Sin dato (volumen insuficiente)', 'txt'),
              C(t['interaccion'] if t['interaccion'] != '-' else 'Sin dato', 'txt'),
              C(t['conversion'] if t['conversion'] != '-' else 'Sin dato', 'txt'))
    rt = len(s.rows) + 1
    s.row(C('TOTAL CAMPAÑA', 'bold'), C('', 'bold'), C('', 'bold'),
          F('=SUM(D%d:D%d)' % (r0, rt - 1), 'boldmoney'), F('=D%d/D%d' % (rt, rt), 'boldpct'),
          F('=SUM(F%d:F%d)' % (r0, rt - 1), 'boldnum'), F('=F%d/F%d' % (rt, rt), 'boldpct'),
          F('=D%d/F%d' % (rt, rt), 'boldmoney'), F('=SUM(I%d:I%d)' % (r0, rt - 1), 'boldnum'),
          F('=SUM(J%d:J%d)' % (r0, rt - 1), 'boldnum'), F('=I%d/J%d' % (rt, rt), 'dec2'),
          F('=D%d/I%d*1000' % (rt, rt), 'boldmoney'),
          F('=SUM(M%d:M%d)' % (r0, rt - 1), 'boldnum'), F('=SUM(N%d:N%d)' % (r0, rt - 1), 'boldnum'),
          F('=SUM(O%d:O%d)' % (r0, rt - 1), 'boldnum'))
    s.blank()
    s.row(C('ADVERTENCIAS METODOLÓGICAS', 'seccion'))
    for t in [
        'El ALCANCE no es aditivo: la suma de los 8 anuncios (%s) contiene duplicación entre anuncios. '
        'El alcance único real de la campaña está entre %s (anuncio mayor) y %s. '
        'Meta no exportó un alcance único a nivel campaña.'
        % ('{:,}'.format(f.alcance_suma), '{:,}'.format(f.alcance_min), '{:,}'.format(f.alcance_suma)),
        'La columna "Leads CRM" cuenta personas únicas cuyo campo «Anuncio de contacto» coincide con la '
        'etiqueta del anuncio. No todos los registros del CRM traen anuncio (%d sin anuncio).'
        % f.dq['leads_sin_anuncio'],
        'Mapeo inferido: los anuncios «video clases» y «video switch» de Meta suman 8 conversaciones, '
        'exactamente las 8 registradas en el CRM como «Reel informativo». La correspondencia 1 a 1 no '
        'puede confirmarse con los datos disponibles y está marcada como inferida.',
        'CTR, CPC y clics no están en la exportación: no son calculables.',
        'Los 8 anuncios pertenecen a un único conjunto («campaña agosto A») con presupuesto a nivel '
        'campaña (CBO), por lo que el reparto de gasto fue decidido por el algoritmo, no manualmente.',
    ]:
        s.row(C('•  ' + t, 'txtw'))
        s.merge('A%d:R%d' % (len(s.rows), len(s.rows)))
    sheets.append(s)

    # ====================================================== 8. REDES SOCIALES
    s = Sheet('Redes sociales', cols=[38, 22, 62], tab_color=X.ROJO)
    s.row(C('REDES SOCIALES — DECLARACIÓN DE DISPONIBILIDAD DE DATOS', 'titulo'))
    s.row(C('El repositorio NO contiene ninguna exportación de métricas orgánicas. '
            'Esta pestaña documenta exactamente qué falta para poder analizarlas.', 'nota'))
    s.blank()
    s.row(C('Métrica solicitada', 'hdr'), C('Disponibilidad', 'hdr'), C('Observación', 'hdr'))
    for m in ['Alcance orgánico (Facebook)', 'Alcance orgánico (Instagram)', 'Alcance orgánico (TikTok)',
              'Impresiones orgánicas', 'Visualizaciones / reproducciones de video',
              'Interacciones (reacciones, comentarios, compartidos)', 'Guardados',
              'Visitas al perfil', 'Mensajes recibidos (orgánicos)', 'Crecimiento de seguidores',
              'Desempeño por publicación / reel orgánico']:
        s.row(C(m, 'txt'), C('No disponible', 'bad'),
              C('No existe archivo fuente en el repositorio (se requiere export de Meta Business Suite '
                'o TikTok Analytics del periodo).', 'txtw'))
    s.blank()
    s.row(C('LO QUE SÍ PUEDE AFIRMARSE CON LOS DATOS EXISTENTES', 'seccion'))
    tot = sum(f.leads_medio.values())
    s.row(C('Medio de contacto declarado en el CRM', 'hdr2'), C('Registros', 'hdr2'), C('% del total', 'hdr2'))
    rr = len(s.rows) + 1
    for k, v in f.leads_medio.most_common():
        s.row(C(k, 'txt'), C(v, 'num'), C(v / tot, 'pct'))
    s.row(C('Total', 'bold'), F('=SUM(B%d:B%d)' % (rr, len(s.rows)), 'boldnum'), C(1.0, 'boldpct'))
    s.blank()
    s.row(C('Estos %d registros son el medio por el que llegó la conversación (Facebook, WhatsApp o '
            'Instagram), no una medición de desempeño orgánico. El 100%% de los registros con anuncio '
            'identificado proviene de los 8 anuncios pagados; no hay forma de separar tráfico orgánico '
            'de pagado con las fuentes actuales.' % tot, 'txtw'))
    s.merge('A%d:C%d' % (len(s.rows), len(s.rows)))
    sheets.append(s)

    # ===================================================== 9. FUNNEL COMERCIAL
    s = Sheet('Funnel comercial', cols=[38, 16, 16, 16, 16, 48], freeze='A5', tab_color=X.AZUL2)
    s.row(C('FUNNEL COMERCIAL COMPLETO', 'titulo'))
    s.row(C('Alcance → Leads → CM agendadas → CM realizadas → Inscripciones  ·  %s' % PER, 'nota'))
    s.blank()
    s.row(C('Etapa', 'hdr'), C('Cantidad', 'hdr'), C('Conv. vs etapa previa', 'hdr'),
          C('Pérdida absoluta', 'hdr'), C('Pérdida %', 'hdr'), C('Fuente', 'hdr'))
    fr0 = len(s.rows) + 1
    for i, (nm, val, fu) in enumerate(f.etapas):
        r = len(s.rows) + 1
        if i == 0:
            s.row(C(nm, 'txt'), C(val, 'num'), C('—', 'txt'), C('—', 'txt'), C('—', 'txt'), C(fu, 'txtw'))
        else:
            s.row(C(nm, 'txt'), C(val, 'num'), F('=B%d/B%d' % (r, r - 1), 'pct'),
                  F('=B%d-B%d' % (r - 1, r), 'num'), F('=1-B%d/B%d' % (r, r - 1), 'pct'), C(fu, 'txtw'))
    s.blank()
    s.row(C('TASAS CLAVE DEL EMBUDO COMERCIAL', 'seccion'))
    s.row(C('Indicador', 'hdr2'), C('Fórmula', 'hdr2'), C('Resultado', 'hdr2'), C('', 'hdr2'),
          C('', 'hdr2'), C('Lectura', 'hdr2'))
    for lab, form, val, lect in [
        ('Lead → CM agendada', 'CM agendadas / Leads', f.t_lead_cm,
         'De cada 100 prospectos, %d llegaron a tener clase muestra agendada.' % round(f.t_lead_cm * 100)),
        ('CM agendada → realizada', 'CM realizadas / CM agendadas', f.t_cm_real,
         'Asistencia saludable: %d de cada 100 clases agendadas se realizaron.' % round(f.t_cm_real * 100)),
        ('CM realizada → Inscripción', 'Inscripciones / personas con CM realizada', f.t_real_insc,
         'Casi 1 de cada 2 personas que tomó la clase se inscribió.'),
        ('Lead → Inscripción (periodo)', 'Inscripciones / Leads', f.t_lead_insc,
         'Conversión global del periodo.'),
        ('Lead → Inscripción (atribuible)', 'Inscripciones atribuibles / Leads', f.t_lead_insc_atrib,
         'Conversión imputable exclusivamente a la pauta.'),
        ('No show rate', 'No show / CM agendadas', f.tasa_noshow,
         '%d clases agendadas se perdieron por inasistencia.' % f.cm_noshow),
        ('Tasa de reprogramación', 'Reprogramadas / CM agendadas', f.tasa_reprog,
         'Solo %d reprogramación registrada: no es un problema.' % f.cm_reprogramadas),
    ]:
        s.row(C(lab, 'txt'), C(form, 'txt'), C(val, 'pct'), C(''), C(''), C(lect, 'txtw'))
    s.blank()
    s.row(C('DÓNDE SE PIERDE EL EMBUDO', 'seccion'))
    s.row(C('Punto de pérdida', 'hdr2'), C('Prospectos perdidos', 'hdr2'), C('% del total de leads', 'hdr2'),
          C('', 'hdr2'), C('', 'hdr2'), C('Diagnóstico sustentado en datos', 'hdr2'))
    perd = [
        ('Lead que nunca agenda clase muestra', f.leads - f.cm_personas,
         (f.leads - f.cm_personas) / f.leads,
         'CUELLO DE BOTELLA PRINCIPAL. %d de %d personas (%.1f%%) no llegaron a clase muestra. '
         'El %d%% de los registros del CRM se quedó en etapa «Información enviada» o «Contactado».'
         % (f.leads - f.cm_personas, f.leads, (f.leads - f.cm_personas) / f.leads * 100,
            round((f.M['leads_por_etapa'].get('Información enviada', 0)
                   + f.M['leads_por_etapa'].get('Contactado', 0)) / f.reg_ventana * 100))),
        ('Clase agendada a la que no asiste', f.cm_noshow, f.cm_noshow / f.leads,
         'Pérdida secundaria: %.1f%% de no show sobre clases agendadas.' % (f.tasa_noshow * 100)),
        ('Clase realizada sin inscripción', f.cm_pers_realizadas - f.insc_personas,
         (f.cm_pers_realizadas - f.insc_personas) / f.leads,
         '%d personas tomaron la clase y no cerraron: la tasa de cierre (%.1f%%) es el punto más '
         'sano del embudo.' % (f.cm_pers_realizadas - f.insc_personas, f.t_real_insc * 100)),
    ]
    for a, b, c, d in perd:
        s.row(C(a, 'txt'), C(b, 'num'), C(c, 'pct'), C(''), C(''), C(d, 'txtw'))
    sheets.append(s)

    # ================================================== 10. ANÁLISIS TEMPORAL
    s = Sheet('Análisis temporal', cols=[26, 14, 14, 14, 14, 14, 14, 40], freeze='A5', tab_color=X.AZUL3)
    s.row(C('EVOLUCIÓN DE RESULTADOS DURANTE LA CAMPAÑA', 'titulo'))
    s.row(C('Un solo esfuerzo de campaña: 21 ago → septiembre → 4 oct. Las fases NO son campañas '
            'independientes; se muestran solo para entender el comportamiento temporal.', 'nota'))
    s.blank()
    s.row(C('Fase', 'hdr'), C('Días', 'hdr'), C('Leads', 'hdr'), C('Leads/día', 'hdr'),
          C('CM agendadas', 'hdr'), C('CM realizadas', 'hdr'), C('Inscripciones', 'hdr'), C('Lectura', 'hdr'))
    lect = {
        '21–31 ago (arranque)': 'Arranque: la pauta se calienta y se concentra la captación, pero la '
                                'operación de clases muestra aún no responde.',
        'Septiembre (principal)': 'Periodo principal: concentra el %d%% de los leads y el %d%% de las '
                                  'inscripciones.',
        '1–4 oct (cierre)': 'Cierre: 4 días con poca pauta efectiva; quedan 3 clases agendadas que '
                            'se resolvieron después del 4 de octubre.',
    }
    fr = len(s.rows) + 1
    for k, v in f.fases.items():
        r = len(s.rows) + 1
        t = lect[k]
        if '%d%%' in t:
            t = t % (round(v['leads'] / f.leads * 100), round(v['insc'] / f.insc_eventos * 100))
        s.row(C(k, 'txt'), C(v['dias'], 'num'), C(v['leads'], 'num'), F('=C%d/B%d' % (r, r), 'dec1'),
              C(v['cm_ag'], 'num'), C(v['cm_re'], 'num'), C(v['insc'], 'num'), C(t, 'txtw'))
    rt = len(s.rows) + 1
    s.row(C('TOTAL CAMPAÑA', 'bold'), F('=SUM(B%d:B%d)' % (fr, rt - 1), 'boldnum'),
          F('=SUM(C%d:C%d)' % (fr, rt - 1), 'boldnum'), F('=C%d/B%d' % (rt, rt), 'dec1'),
          F('=SUM(E%d:E%d)' % (fr, rt - 1), 'boldnum'), F('=SUM(F%d:F%d)' % (fr, rt - 1), 'boldnum'),
          F('=SUM(G%d:G%d)' % (fr, rt - 1), 'boldnum'),
          C('Métrica principal: el acumulado 21 ago – 4 oct.', 'txtw'))
    s.blank()
    s.row(C('DETALLE SEMANAL', 'seccion'))
    s.row(C('Semana', 'hdr'), C('Leads', 'hdr'), C('CM agendadas', 'hdr'), C('CM realizadas', 'hdr'),
          C('Inscripciones', 'hdr'), C('Lead→CM', 'hdr'), C('', 'hdr'), C('', 'hdr'))
    sr = len(s.rows) + 1
    for lab, l, ca, cr, ins in f.semanas:
        r = len(s.rows) + 1
        s.row(C(lab, 'txt'), C(l, 'num'), C(ca, 'num'), C(cr, 'num'), C(ins, 'num'),
              F('=IF(B%d=0,"",C%d/B%d)' % (r, r, r), 'pct'))
    s.blank()
    s.row(C('CAPTACIÓN DIARIA (registros del CRM)', 'seccion'))
    s.row(C('Fecha', 'hdr'), C('Registros', 'hdr'), C('', 'hdr'), C('', 'hdr'), C('', 'hdr'),
          C('', 'hdr'), C('', 'hdr'), C('', 'hdr'))
    for d, v in f.serie_leads.items():
        s.row(C(d, 'fecha'), C(v, 'num'))
    s.blank()
    s.row(C('Pico de captación: %s con %d registros en un día.'
            % (f.pico_leads[0].strftime('%d/%m/%Y'), f.pico_leads[1]), 'nota'))
    sheets.append(s)

    # ======================================================== 11. SEGMENTOS
    s = Sheet('Segmentos', cols=[26, 14, 14, 14, 14, 14, 44], freeze='A5', tab_color=X.AZUL3)
    s.row(C('SEGMENTOS: PROGRAMA, IDIOMA, HORARIO Y CANAL DE AGENDA', 'titulo'))
    s.row(C('Base: los %d registros de clase muestra con fecha en la ventana de campaña.' % f.cm_agendadas, 'nota'))
    s.blank()
    s.row(C('Programa', 'hdr'), C('CM agendadas', 'hdr'), C('CM realizadas', 'hdr'),
          C('Inscripciones', 'hdr'), C('Asistencia', 'hdr'), C('Cierre post-CM', 'hdr'), C('Lectura', 'hdr'))
    pr = len(s.rows) + 1
    for k, (ag, re_, ins) in sorted(f.conv_cat.items(), key=lambda kv: -kv[1][0]):
        r = len(s.rows) + 1
        s.row(C(k, 'txt'), C(ag, 'num'), C(re_, 'num'), C(ins, 'num'),
              F('=C%d/B%d' % (r, r), 'pct'), F('=IF(C%d=0,"",D%d/C%d)' % (r, r, r), 'pct'), C('', 'txtw'))
    rt = len(s.rows) + 1
    s.row(C('TOTAL', 'bold'), F('=SUM(B%d:B%d)' % (pr, rt - 1), 'boldnum'),
          F('=SUM(C%d:C%d)' % (pr, rt - 1), 'boldnum'), F('=SUM(D%d:D%d)' % (pr, rt - 1), 'boldnum'),
          F('=C%d/B%d' % (rt, rt), 'boldpct'), F('=D%d/C%d' % (rt, rt), 'boldpct'))
    s.blank()
    s.row(C('Idioma', 'hdr'), C('CM agendadas', 'hdr'), C('Inscripciones', 'hdr'), C('', 'hdr'),
          C('', 'hdr'), C('', 'hdr'), C('Lectura', 'hdr'))
    for k, v in f.cm_idioma.most_common():
        s.row(C(k, 'txt'), C(v, 'num'), C(f.insc_idioma.get(k, 0), 'num'))
    s.blank()
    s.row(C('Horario de la clase muestra', 'hdr'), C('CM agendadas', 'hdr'), C('', 'hdr'), C('', 'hdr'),
          C('', 'hdr'), C('', 'hdr'), C('', 'hdr'))
    for k, v in f.cm_hora.most_common():
        s.row(C(k, 'txt'), C(v, 'num'))
    s.blank()
    s.row(C('Canal usado para agendar la clase', 'hdr'), C('CM agendadas', 'hdr'), C('Inscripciones', 'hdr'),
          C('Cierre', 'hdr'), C('', 'hdr'), C('', 'hdr'), C('', 'hdr'))
    cr = len(s.rows) + 1
    for k, v in f.cm_medio_ag.most_common():
        r = len(s.rows) + 1
        s.row(C(k, 'txt'), C(v, 'num'), C(f.insc_medio_ag.get(k, 0), 'num'),
              F('=IF(B%d=0,"",C%d/B%d)' % (r, r, r), 'pct'))
    s.blank()
    s.row(C('Edades', 'hdr'), C('n', 'hdr'), C('Mín', 'hdr'), C('Máx', 'hdr'), C('Promedio', 'hdr'),
          C('', 'hdr'), C('', 'hdr'))
    s.row(C('Clases muestra', 'txt'), C(f.edad_cm['n'], 'num'), C(f.edad_cm['min'], 'num'),
          C(f.edad_cm['max'], 'num'), C(f.edad_cm['prom'], 'dec1'))
    s.row(C('Inscripciones', 'txt'), C(f.edad_insc['n'], 'num'), C(f.edad_insc['min'], 'num'),
          C(f.edad_insc['max'], 'num'), C(f.edad_insc['prom'], 'dec1'))
    s.blank()
    s.row(C('Medio de contacto del lead (CRM)', 'hdr'), C('Registros', 'hdr'), C('', 'hdr'), C('', 'hdr'),
          C('', 'hdr'), C('', 'hdr'), C('', 'hdr'))
    for k, v in f.leads_medio.most_common():
        s.row(C(k, 'txt'), C(v, 'num'))
    sheets.append(s)

    # ========================================================= 12. GRÁFICAS
    s = _graficas(f)
    sheets.append(s)

    # ========================================================== 13. FUENTES
    s = Sheet('Fuentes', cols=[40, 46, 26, 60], freeze='A5', tab_color=X.GRIS)
    s.row(C('FUENTES UTILIZADAS Y TRAZABILIDAD', 'titulo'))
    s.row(C('Cada cifra del análisis puede rastrearse hasta uno de estos cuatro archivos.', 'nota'))
    s.blank()
    s.row(C('Fuente', 'hdr'), C('Información utilizada', 'hdr'), C('Periodo', 'hdr'), C('Observaciones', 'hdr'))
    for a, b, c, d in [
        ('Kugelman-Academy-Anuncios-21-ago-2026---4-oct-2026.csv',
         'Gasto real, impresiones, alcance, conversaciones iniciadas, frecuencia, CPM, costo por '
         'conversación y clasificaciones de calidad por anuncio.',
         '21 ago – 4 oct 2026 (fijo en el archivo)',
         'Nivel anuncio (8 filas), un solo conjunto «campaña agosto A». No incluye clics, CTR, CPC '
         'ni desglose diario. El alcance por anuncio no es sumable sin duplicación.'),
        ('CAMPAÑAS (2).xlsx — hoja «AGOSTOSEPTIEMBRE»',
         'Leads/prospectos de la campaña: fecha, nombre, teléfono, idioma, medio, anuncio de contacto, '
         'canal, segmento, etapa del proceso, registro y asistencia a CM, resultado.',
         '19 ago – 4 oct 2026 (504 registros)',
         'Base del conteo de leads. 5 registros del 19–20 ago quedan fuera de la ventana y se marcan. '
         'Campos de resultado muy incompletos (%d de %d registros sin resultado).'
         % (f.dq['leads_sin_resultado'], f.reg_ventana)),
        ('CAMPAÑAS (2).xlsx — hoja «Octubre»',
         'Verificación de duplicidad: 5 registros del 30 de septiembre.',
         '30 sep 2026',
         'Los %d registros son copia exacta de filas ya presentes en AGOSTOSEPTIEMBRE. '
         'NO se suman, para no duplicar leads.' % f.oct_dup),
        ('CAMPAÑAS (2).xlsx — hoja «OCTUBRE 1»',
         'Ninguna: solo contiene encabezados.',
         '—', 'Hoja vacía (formato preparado para la campaña de octubre).'),
        ('CAMPAÑAS (2).xlsx — hoja «JULIO»',
         'Referencia comparativa y revisión de prospectos abiertos de la campaña anterior.',
         'Julio 2026 (%d registros)' % f.reg_julio,
         'EXCLUIDA de todos los KPI de la Campaña de Septiembre. Se reporta como referencia en el '
         'archivo «Prospectos en Curso».'),
        ('CLASES MUESTRA 2026__ (1).xlsx — hojas AGOSTO / SEPTIEMBRE / OCTUBRE',
         'Clases muestra: fecha, nombre, teléfono, categoría, edad, idioma, resultado, medio para '
         'agendar y hora. Es la fuente de CM agendadas, realizadas, no show e inscripciones.',
         '12 ago – 7 oct 2026 (107 registros; %d en ventana)' % f.cm_agendadas,
         'Fuente única de la inscripción. 3 registros traen el año 2006 por error de captura '
         '(corregidos a 2026 y documentados). Las hojas ENERO–JULIO corresponden a meses previos '
         'y no se usan.'),
        ('RGA_1791241975.xlsx — «General de alumnos»',
         'Padrón de alumnos: matrícula, nombre, apellidos, sexo, edad, activo. Se usa para validar '
         'las inscripciones y para construir el Entregable 4.',
         'Exportado el 05/Oct/2026 (%d alumnos)' % f.dq['rga_n'],
         'NO contiene fecha de inscripción, programa, idioma ni grupo por alumno, por lo que no es '
         'posible filtrar el padrón al corte del 3 de octubre ni segmentar inscripciones desde aquí.'),
    ]:
        s.row(C(a, 'txt'), C(b, 'txtw'), C(c, 'txt'), C(d, 'txtw'))
    s.blank()
    s.row(C('INFORMACIÓN SOLICITADA QUE NO EXISTE EN EL REPOSITORIO', 'seccion'))
    for t in ['Métricas orgánicas de Facebook, Instagram y TikTok (alcance, interacciones, guardados, '
              'visitas al perfil, seguidores, desempeño por publicación).',
              'Clics, CTR y CPC de los anuncios.',
              'Desglose diario o semanal del gasto publicitario.',
              'Hora/fecha del primer contacto y del primer seguimiento (no se puede medir tiempo de respuesta).',
              'Motivo de pérdida de los prospectos que no cerraron.',
              'Fecha de inscripción, programa, grupo, modalidad y mensualidad por alumno.',
              'Campo de "clase muestra solicitada" distinto de "agendada".',
              'Identificador de campaña/conjunto por lead (solo se registra el nombre del anuncio).']:
        s.row(C('•  ' + t, 'txtw'))
        s.merge('A%d:D%d' % (len(s.rows), len(s.rows)))
    sheets.append(s)

    # ================================================= 14. CALIDAD DE DATOS
    s = _calidad(f)
    sheets.append(s)

    path = os.path.join(OUT, 'Base Maestra Campaña Septiembre 2026.xlsx')
    X.write(path, sheets, 'Base Maestra Campaña Septiembre 2026 — Kugelman Academy')
    return path


# ------------------------------------------------------------------ GRÁFICAS
def _graficas(f):
    s = Sheet('Gráficas', cols=[30, 16, 16, 16, 16, 3] + [11] * 14, tab_color=X.AMBAR)
    s.row(C('TABLAS DE DATOS Y GRÁFICOS EDITABLES', 'titulo'))
    s.row(C('Cada gráfico está ligado a su tabla: al cambiar un número de la tabla, el gráfico se '
            'actualiza. Los gráficos son objetos nativos de Excel, totalmente editables.', 'nota'))
    s.blank()
    SH = 'Gráficas'

    def blk(titulo, cols, rows, chart_kind, chart_title, anchor, w=9, h=17, stacked=False,
            numfmt='General', series_cols=None, cat_title='', val_title=''):
        s.row(C(titulo, 'seccion'), *[C('', 'seccion') for _ in cols[:-1]])
        s.row(*[C(c, 'hdr') for c in cols])
        r0 = len(s.rows) + 1
        for row in rows:
            cells = [C(row[0], 'txt')]
            for v in row[1:]:
                st = 'money' if numfmt.startswith('"$"') else ('pct' if numfmt.endswith('%') else 'num')
                cells.append(C(v, st))
            s.row(*cells)
        r1 = len(s.rows)
        cats = [row[0] for row in rows]
        sc = series_cols or [(cols[i], X.colname(i)) for i in range(1, len(cols))]
        series = []
        for ci, (nm, col) in enumerate(sc, start=1):
            vals = [row[ci] for row in rows]
            series.append((nm, "'%s'!$%s$%d:$%s$%d" % (SH, col, r0, col, r1), vals))
        s.charts.append(Chart(chart_kind, chart_title, SH,
                              "'%s'!$A$%d:$A$%d" % (SH, r0, r1), cats, series, anchor,
                              width=w, height=h, stacked=stacked, numfmt=numfmt,
                              cat_title=cat_title, val_title=val_title))
        s.blank()
        return r0, r1

    # G1 funnel
    blk('GRÁFICA 1 · Funnel comercial completo',
        ['Etapa', 'Personas / eventos'],
        [('Conversaciones Meta', f.conversaciones), ('Leads (personas)', f.leads),
         ('CM agendadas', f.cm_agendadas), ('CM realizadas', f.cm_realizadas),
         ('Inscripciones', f.insc_personas)],
        'bar', 'Funnel comercial: Conversaciones → Leads → CM → Inscripciones', 'G3', w=9, h=16)

    # G2 conversión por etapa
    blk('GRÁFICA 2 · Conversión por etapa del embudo',
        ['Paso', 'Conversión'],
        [('Lead → CM agendada', round(f.t_lead_cm, 4)),
         ('CM agendada → realizada', round(f.t_cm_real, 4)),
         ('CM realizada → inscripción', round(f.t_real_insc, 4)),
         ('Lead → inscripción', round(f.t_lead_insc, 4))],
        'col', 'Conversión por etapa (%)', 'Q3', w=9, h=16, numfmt='0.0%')

    # G3 CM agendadas vs realizadas vs inscritos
    blk('GRÁFICA 3 · Clases muestra: agendadas vs realizadas vs inscripciones',
        ['Resultado de la clase muestra', 'Clases'],
        [('Realizadas con inscripción', f.cm_inscritos_reg),
         ('Realizadas sin cierre', f.cm_realizadas - f.cm_inscritos_reg),
         ('No show', f.cm_noshow), ('Pendientes al 4-oct', f.cm_pendientes),
         ('Reprogramadas', f.cm_reprogramadas)],
        'col', 'Resultado de las %d clases muestra agendadas' % f.cm_agendadas, 'G22', w=9, h=16)

    # G4 resultados por anuncio
    rows = [(t['anuncio'], t['conv'], round(t['costo_conv'], 2) if t['costo_conv'] else 0)
            for t in sorted(f.tab_ads, key=lambda r: -r['conv'])]
    blk('GRÁFICA 4 · Resultados por anuncio: volumen vs eficiencia',
        ['Anuncio', 'Conversaciones', 'Costo por conversación (MXN)'], rows,
        'col', 'Conversaciones por anuncio', 'Q22', w=9, h=16,
        series_cols=[('Conversaciones', 'B')])
    # grafico adicional de eficiencia sobre la misma tabla
    r1 = len(s.rows) - 1
    r0 = r1 - len(rows) + 1
    s.charts.append(Chart('col', 'Costo por conversación por anuncio (MXN) — menor es mejor', SH,
                          "'%s'!$A$%d:$A$%d" % (SH, r0, r1), [r[0] for r in rows],
                          [('Costo por conversación', "'%s'!$C$%d:$C$%d" % (SH, r0, r1),
                            [r[2] for r in rows])], 'G41', width=9, height=16,
                          numfmt='"$"#,##0.00', colors=[X.AMBAR]))

    # G5 distribución de inversión
    blk('GRÁFICA 5 · Distribución de la inversión por anuncio',
        ['Anuncio', 'Gasto (MXN)'],
        [(t['anuncio'], round(t['gasto'], 2)) for t in sorted(f.tab_ads, key=lambda r: -r['gasto'])],
        'doughnut', 'Reparto del gasto real ($%s)' % '{:,.2f}'.format(f.gasto), 'Q41',
        w=9, h=16, numfmt='"$"#,##0')

    # G6 leads vs inscritos por anuncio
    etiquetas = sorted(set(list(f.crm_pers_ad.keys())), key=lambda k: -f.crm_pers_ad[k])
    rows = [(k, f.crm_pers_ad[k], f.cm_ad.get(k, 0), f.insc_ad.get(k, 0)) for k in etiquetas]
    blk('GRÁFICA 6 · Leads, clases muestra e inscripciones por anuncio de contacto (CRM)',
        ['Anuncio de contacto (CRM)', 'Leads (personas)', 'Clases muestra', 'Inscripciones atribuibles'],
        rows, 'col', 'Del lead a la inscripción, por anuncio de contacto', 'G60', w=9, h=16)

    # G7 distribución de inscritos por programa
    blk('GRÁFICA 7 · Distribución de las inscripciones por programa',
        ['Programa', 'Inscripciones'], [(k, v) for k, v in f.insc_cat.most_common()],
        'doughnut', 'Inscripciones por programa (%d eventos)' % f.insc_eventos, 'Q60', w=9, h=16)

    # G8 mejores contenidos (formato / clasificación)
    rows = []
    for t in sorted(f.tab_ads, key=lambda r: -(r['conv'] / r['impresiones'] * 1000 if r['impresiones'] else 0)):
        rows.append((t['anuncio'], round(t['conv'] / t['impresiones'] * 1000, 3) if t['impresiones'] else 0))
    blk('GRÁFICA 8 · Mejores contenidos: conversaciones por cada 1,000 impresiones',
        ['Anuncio', 'Conversaciones / 1,000 impresiones'], rows,
        'bar', 'Eficacia del creativo: conversaciones por mil impresiones', 'G79', w=9, h=16,
        numfmt='0.00')

    # G9 evolución semanal
    blk('GRÁFICA 9 · Evolución semanal de resultados',
        ['Semana', 'Leads', 'CM agendadas', 'CM realizadas', 'Inscripciones'],
        [(lab, l, ca, cr, ins) for lab, l, ca, cr, ins in f.semanas],
        'line', 'Evolución semanal: leads, clases muestra e inscripciones', 'Q79', w=9, h=16)

    # G10 prospectos por estatus
    blk('GRÁFICA 10 · Prospectos por estatus homologado (pipeline al 3-oct)',
        ['Estatus', 'Prospectos'],
        [(k, f.estatus_count.get(k, 0)) for k in facts.ESTATUS_ORDEN],
        'bar', 'Pipeline comercial por estatus al corte del 3 de octubre', 'G98', w=9, h=18)

    # G11 fases
    blk('GRÁFICA 11 · Resultados por fase de la campaña',
        ['Fase', 'Leads', 'CM agendadas', 'Inscripciones'],
        [(k, v['leads'], v['cm_ag'], v['insc']) for k, v in f.fases.items()],
        'col', 'Un solo esfuerzo, tres fases: arranque, septiembre y cierre', 'Q98', w=9, h=18)
    return s


# ------------------------------------------------------------ CALIDAD DE DATOS
def _calidad(f):
    s = Sheet('Calidad de datos', cols=[48, 14, 14, 62], freeze='A5', tab_color=X.ROJO)
    s.row(C('CALIDAD Y LIMITACIONES DE LOS DATOS', 'titulo'))
    s.row(C('No se corrigió ninguna inconsistencia en silencio: todo lo detectado queda documentado aquí.', 'nota'))
    s.blank()
    s.row(C('Hallazgo de calidad', 'hdr'), C('Cantidad', 'hdr'), C('% de la base', 'hdr'),
          C('Implicación para el análisis', 'hdr'))
    N = f.reg_ventana
    filas = [
        ('CAMPOS INCOMPLETOS EN EL CRM DE CAMPAÑA (base: %d registros en ventana)' % N, None, None, None),
        ('Registros sin nombre del interesado', f.dq['leads_sin_nombre'], f.dq['leads_sin_nombre'] / N,
         'No se pueden cruzar con clases muestra ni con el padrón; limitan la atribución.'),
        ('Registros sin teléfono', f.dq['leads_sin_tel'], f.dq['leads_sin_tel'] / N,
         'El teléfono es el identificador más fiable: sin él, el enlace depende solo del nombre.'),
        ('Registros sin nombre NI teléfono', f.dq['leads_sin_nombre_ni_tel'],
         f.dq['leads_sin_nombre_ni_tel'] / N, 'Imposibles de rastrear en el embudo.'),
        ('Registros sin idioma de interés', f.dq['leads_sin_idioma'], f.dq['leads_sin_idioma'] / N,
         'Impide segmentar la demanda por idioma desde el CRM.'),
        ('Registros sin segmento / programa', f.dq['leads_sin_segmento'], f.dq['leads_sin_segmento'] / N,
         'El perfil del prospecto solo se conoce cuando llega a clase muestra.'),
        ('Registros sin anuncio de contacto', f.dq['leads_sin_anuncio'], f.dq['leads_sin_anuncio'] / N,
         'Esos leads no pueden imputarse a un creativo específico.'),
        ('Registros sin etapa del proceso', f.dq['leads_sin_etapa'], f.dq['leads_sin_etapa'] / N,
         'Quedan como «Estatus por validar» en el pipeline.'),
        ('Registros sin resultado del seguimiento', f.dq['leads_sin_resultado'],
         f.dq['leads_sin_resultado'] / N,
         'PROBLEMA MÁS GRAVE: el CRM no cierra el ciclo. El resultado comercial hubo que '
         'reconstruirlo desde la bitácora de clases muestra y el padrón.'),
        ('Columna «Comentario adicional» vacía en toda la hoja', f.dq['leads_comentario_vacio'], 1.0,
         'No hay contexto cualitativo: no se puede analizar motivo de pérdida ni objeciones.'),
        ('DUPLICIDAD', None, None, None),
        ('Registros duplicados detectados y fusionados', f.dup_fusionados, f.tasa_duplicidad,
         'Sin deduplicar, los leads se sobreestimarían en %.1f%%.' % (f.tasa_duplicidad * 100),),
        ('Registros marcados para revisión (no fusionados)', len(f.revisar), len(f.revisar) / N,
         'Nombres de una sola palabra repetidos con teléfono distinto o ausente: no hay evidencia '
         'suficiente para fusionarlos.'),
        ('Personas sin ningún identificador', len(f.sin_id), len(f.sin_id) / N,
         'Sin nombre ni teléfono: se conservan en la base pero no son rastreables.'),
        ('Registros duplicados entre hojas (hoja «Octubre»)', f.oct_dup, None,
         'Los 5 registros del 30-sep de la hoja «Octubre» ya estaban en AGOSTOSEPTIEMBRE. '
         'Se excluyeron del conteo.'),
        ('INCONSISTENCIAS Y CONFLICTOS ENTRE FUENTES', None, None, None),
        ('Fechas con año erróneo en la bitácora de CM', f.dq['cm_anio_erroneo'], None,
         'Filas %s de la hoja SEPTIEMBRE traen año 2006. Se interpretaron como 2026 (inferencia '
         'documentada); el resto de la hoja es consistente con septiembre de 2026.'
         % ', '.join(str(k[1]) for k in f.anio_typos)),
        ('Inscripción registrada en el CRM pero NO en la bitácora de CM', len(f.insc_conflicto), None,
         'Caso detectado: %s. El CRM la marca «Inscrito» y aparece en el padrón, pero la bitácora '
         'dice «Show». Se contabilizó como inscripción por tener doble evidencia.'
         % ', '.join(r['nombre'] for r in f.insc_conflicto)),
        ('Alumnos del padrón con clase muestra sin cierre marcado', len(f.insc_por_validar), None,
         'Casos: %s. Aparecen en el padrón con matrícula reciente pero la bitácora no los marca '
         'inscritos y el CRM no lo confirma. NO se contabilizaron; requieren validación.'
         % ', '.join('%s (mat. %s)' % (r['nombre'], r['matricula']) for r in f.insc_por_validar)),
        ('Inscripciones de la bitácora sin correspondencia en el padrón', len(f.insc_sin_padron), None,
         'Casos: %s. Marcadas «Inscrito» en la bitácora pero no localizadas en el padrón del 05-oct. '
         'Puede ser alta no procesada, baja posterior o error de captura.'
         % ', '.join(r['nombre'] for r in f.insc_sin_padron)),
        ('Diferencias de edad entre bitácora y padrón (≥3 años)', len(f.edad_dif), None,
         'Casos: %s. No invalidan el enlace (nombre completo coincide) pero indican captura poco '
         'cuidada de la edad.'
         % ('; '.join('%s: %s vs %s' % (a, b, d) for a, b, c, d, e in f.edad_dif) or 'ninguno')),
        ('Clases muestra sin lead identificado en el CRM', f.cm_sin_lead,
         f.cm_sin_lead / f.cm_agendadas,
         'El %.0f%% de las clases muestra no puede ligarse a un lead de la pauta. Incluye '
         'recomendaciones y visitas a sitio, pero también fallas de registro.'
         % (f.cm_sin_lead / f.cm_agendadas * 100)),
        ('Brecha de registro entre Meta y el CRM', f.gap_registro, 1 - f.cobertura_crm,
         'Meta reporta %d conversaciones y el CRM capturó %d registros: %d conversaciones '
         'no dejaron rastro en el CRM.' % (f.conversaciones, f.reg_ventana, f.gap_registro)),
        ('Subregistro del campo «Registrado a clase muestra»', f.cm_agendadas - f.cm_solicitadas_crm, None,
         'El CRM marca %d personas registradas a CM frente a %d clases de la bitácora: el campo '
         'no es utilizable como métrica.' % (f.cm_solicitadas_crm, f.cm_agendadas)),
        ('Valor no válido en «Resultado del seguimiento»', 1, None,
         'Un registro tiene el valor «1» en lugar de un estatus.'),
        ('PADRÓN DE ALUMNOS (RGA)', None, None, None),
        ('Alumnos en el padrón', f.dq['rga_n'], None,
         'Exportado el 05/Oct/2026. Todos con Activo = «Sí».'),
        ('Pares de matrículas con nombre, sexo y edad idénticos', len(f.dq['rga_dup']), None,
         'Casos: %s. Dos de ellos (García Vazquez) se explican por una segunda inscripción en '
         'francés; los otros dos requieren validación administrativa. No se eliminó ningún registro.'
         % '; '.join('%s → matrículas %s' % (v[0]['completo'], '/'.join(str(a['matricula']) for a in v))
                     for v in f.dq['rga_dup'].values())),
        ('Registro con apellido «Prueba»', 1, None,
         'Matrícula 1063 «Gustavo Martinez Prueba»: parece un registro de prueba del sistema. '
         'Se conserva porque está en la fuente; se marca para validación.'),
        ('Campos ausentes en el padrón', None, None,
         'No hay fecha de inscripción, programa, idioma, grupo, modalidad ni estatus de pago. '
         'Por eso el padrón NO puede filtrarse al corte del 3 de octubre ni segmentarse.'),
    ]
    for a, b, c, d in filas:
        if b is None and c is None and d is None:
            s.row(C(a, 'seccion'), C('', 'seccion'), C('', 'seccion'), C('', 'seccion'))
        else:
            s.row(C(a, 'txt'), C(b if b is not None else 'n/a', 'num' if isinstance(b, int) else 'txt'),
                  C(c if c is not None else '', 'pct'), C(d, 'txtw'))
    s.blank()
    s.row(C('QUÉ DEBERÍA EMPEZAR A REGISTRARSE DE FORMA SISTEMÁTICA', 'seccion'),
          C('', 'seccion'), C('', 'seccion'), C('', 'seccion'))
    s.row(C('Campo', 'hdr'), C('', 'hdr'), C('', 'hdr'), C('Para qué sirve', 'hdr'))
    for a, b in [
        ('Fecha y hora exactas del lead', 'Medir tiempo de respuesta, hoy no calculable.'),
        ('Fecha y hora del primer contacto de vuelta', 'Medir velocidad de atención y su efecto en el agendamiento.'),
        ('Anuncio, conjunto y campaña (ID, no nombre libre)', 'Atribuir sin ambigüedad y evitar mapeos inferidos.'),
        ('Teléfono normalizado obligatorio (10 dígitos)', 'Es el único identificador fiable para cruzar bases.'),
        ('Idioma, programa y edad desde el primer contacto', 'Segmentar la demanda sin esperar a la clase muestra.'),
        ('Fecha de solicitud de CM, distinta de la fecha de agenda', 'Separar «solicitada» de «agendada» en el embudo.'),
        ('Asistencia a CM como campo obligatorio', 'Medir no show de forma directa.'),
        ('Bitácora de seguimientos con fecha por interacción', 'Saber cuántos toques recibió cada prospecto.'),
        ('Estatus homologado con lista cerrada de valores', 'Evitar los %d valores libres y campos vacíos actuales.'
         % len(set(f.M['leads_por_etapa']))),
        ('Fecha de inscripción y motivo de pérdida', 'Cerrar el ciclo y calcular CAC real por cohorte.'),
        ('Export mensual de métricas orgánicas de FB/IG/TikTok', 'Hoy no existe: la sección orgánica no es analizable.'),
        ('Export de anuncios con clics y desglose diario', 'Permitiría calcular CTR, CPC y curvas de entrega.'),
    ]:
        s.row(C(a, 'txt'), C(''), C(''), C(b, 'txtw'))
    return s


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    print(build())
