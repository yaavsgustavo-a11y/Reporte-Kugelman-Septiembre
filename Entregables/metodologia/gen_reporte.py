# -*- coding: utf-8 -*-
"""ENTREGABLE 1 — Reporte Campaña Septiembre 2026 — Kugelman Academy (.docx)"""
import os, collections
import facts
from docxw import Doc, AZUL, AZUL2, AZUL3, VERDE, AMBAR, ROJO, GRIS
from chartsx import GChart

OUT = '/projects/sandbox/entregables'
PER = '21 de agosto – 4 de octubre de 2026'


def mx(v, dec=2):
    return '$' + ('{:,.%df}' % dec).format(v)


def n(v):
    return '{:,}'.format(v)


def pc(v, dec=1):
    return ('{:.%df}%%' % dec).format(v * 100)


def portada(d, f):
    d.spacer(3)
    d.p('KUGELMAN ACADEMY', size=18, color=AZUL2, b=1, space_after=40, align='left')
    d.body.append('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="18" w:color="%s"/></w:pBdr>'
                  '<w:spacing w:after="240"/></w:pPr></w:p>' % AZUL)
    d.p('REPORTE DE CAMPAÑA', size=52, color=AZUL, b=1, space_after=0, align='left')
    d.p('SEPTIEMBRE 2026', size=52, color=AZUL, b=1, space_after=160, align='left')
    d.p('Marketing, captación y resultados comerciales', size=26, color='404040', space_after=320,
        align='left')
    d.kv('Periodo de evaluación', PER)
    d.kv('Presupuesto asignado', mx(f.presupuesto, 0) + ' MXN')
    d.kv('Gasto real', mx(f.gasto) + ' MXN  (' + pc(f.uso_presupuesto) + ' del presupuesto)')
    d.kv('Inscripciones atribuibles a la campaña', '%d' % f.insc_atribuibles)
    d.kv('Costo de adquisición (CAC)', mx(f.cac))
    d.spacer(1)
    d.callout('Alcance de este documento',
              'Analiza el esfuerzo completo de la campaña del 21 de agosto al 4 de octubre de 2026 como '
              'un solo periodo, cruzando las cuatro fuentes disponibles en el repositorio: la exportación '
              'de anuncios de Meta, el CRM de leads, la bitácora de clases muestra y el padrón de alumnos. '
              'Las bases operativas al corte del 3 de octubre (alumnos inscritos y prospectos en curso) '
              'son entregables independientes y no modifican el periodo de campaña.')
    d.spacer(1)
    d.small('Documento totalmente editable. Todas las tablas son editables y cada gráfica está ligada a '
            'su tabla de datos (clic derecho › Editar datos). Las tablas fuente completas están en el '
            'archivo «Base Maestra Campaña Septiembre 2026.xlsx».')
    d.pagebreak()


def resumen(d, f):
    d.h1('1. Resumen ejecutivo')
    d.p('Periodo %s. La campaña consumió prácticamente todo el presupuesto asignado y generó un volumen '
        'de demanda alto y barato. El resultado comercial, en cambio, quedó limitado por un solo eslabón: '
        'el paso de la conversación inicial a la clase muestra agendada.' % PER)
    d.kpi_grid([
        ('Gasto real', mx(f.gasto), '%s del presupuesto' % pc(f.uso_presupuesto)),
        ('Impresiones', n(f.impresiones), 'CPM %s' % mx(f.cpm)),
        ('Conversaciones', n(f.conversaciones), '%s por conversación' % mx(f.costo_conversacion)),
        ('Leads (personas)', n(f.leads), 'CPL %s' % mx(f.cpl_persona)),
    ])
    d.kpi_grid([
        ('CM realizadas', n(f.cm_realizadas), 'de %d agendadas' % f.cm_agendadas),
        ('Inscripciones', n(f.insc_personas), '%d atribuibles a campaña' % f.insc_atribuibles),
        ('Cierre tras CM', pc(f.t_real_insc), 'punto más fuerte'),
        ('CAC', mx(f.cac), 'sobre %d atribuibles' % f.insc_atribuibles),
    ])

    d.h2('1.1 Dashboard ejecutivo')
    rows = [
        ['Presupuesto asignado', mx(f.presupuesto, 0), 'Dato de Dirección'],
        ['Gasto real', mx(f.gasto), 'Meta Ads — 8 anuncios'],
        ['% de presupuesto utilizado', pc(f.uso_presupuesto), 'Gasto / presupuesto'],
        ['Alcance (personas)', '%s – %s' % (n(f.alcance_min), n(f.alcance_suma)),
         'Meta no exportó alcance único de campaña'],
        ['Impresiones', n(f.impresiones), 'Meta Ads'],
        ['Frecuencia', '%.2f – %.2f' % (f.frec_sobre_suma, f.frec_sobre_min), 'Impresiones / alcance'],
        ['Conversaciones iniciadas', n(f.conversaciones), 'Resultado optimizado de Meta'],
        ['Leads / prospectos (personas únicas)', n(f.leads), 'CRM deduplicado'],
        ['Registros capturados en el CRM', n(f.reg_ventana), '%d duplicados fusionados' % f.dup_fusionados],
        ['CM solicitadas', 'No disponible', 'No existe campo de solicitud en las fuentes'],
        ['CM agendadas', n(f.cm_agendadas), 'Bitácora de clases muestra'],
        ['CM realizadas', n(f.cm_realizadas), 'Resultado Inscrito o Show'],
        ['Inscripciones atribuibles', n(f.insc_atribuibles), 'Con lead identificado en el CRM'],
        ['Lead → CM agendada', pc(f.t_lead_cm), 'CM agendadas / leads'],
        ['CM realizada → Inscripción', pc(f.t_real_insc), 'Inscripciones / personas con CM realizada'],
        ['Lead → Inscripción', pc(f.t_lead_insc), 'Inscripciones / leads'],
        ['Costo por lead', mx(f.cpl_persona), 'Gasto / leads'],
        ['Costo por CM agendada', mx(f.costo_cm_agendada), 'Gasto / CM agendadas'],
        ['Costo por CM realizada', mx(f.costo_cm_realizada), 'Gasto / CM realizadas'],
        ['CAC', mx(f.cac), 'Gasto / %d inscripciones atribuibles' % f.insc_atribuibles],
    ]
    d.table(['Indicador', 'Campaña Septiembre 2026', 'Cómo se calcula / fuente'], rows,
            widths=[3500, 2100, 3760], aligns=['left', 'right', 'left'],
            caption='Tabla 1. Dashboard ejecutivo de la campaña (21 ago – 4 oct 2026)')
    d.small('Los dos indicadores marcados como rango o no disponibles se explican en la sección 11: '
            'Meta no exportó un alcance único a nivel campaña y las fuentes no distinguen una clase '
            'muestra «solicitada» de una «agendada».')

    d.h2('1.2 Conclusiones ejecutivas')
    for i, (t, txt) in enumerate([
        ('La campaña compró demanda de forma eficiente',
         'Con %s de inversión se generaron %d conversaciones y %d prospectos identificables, a %s por '
         'persona. Para una academia de idiomas con ticket recurrente, ese costo de entrada es bajo: el '
         'problema de esta campaña no fue conseguir interesados.'
         % (mx(f.gasto), f.conversaciones, f.leads, mx(f.cpl_persona))),
        ('El cuello de botella está entre el lead y la clase muestra, no en la pauta ni en el cierre',
         'De %d personas interesadas solo %d llegaron a tener clase muestra (%s). En cambio, %s de las '
         'clases agendadas se realizaron y %s de quienes asistieron se inscribieron. El embudo se rompe '
         'en un solo punto, y es un punto de operación comercial, no de marketing.'
         % (f.leads, f.cm_personas, pc(f.t_lead_cm), pc(f.t_cm_real), pc(f.t_real_insc))),
        ('La clase muestra es el activo comercial más rentable de la academia',
         'Cada clase muestra realizada costó %s de pauta y produjo 0.45 inscripciones. Es la palanca con '
         'mayor retorno: cualquier acción que incremente clases muestra se traduce casi directamente en '
         'inscripciones, sin necesidad de más presupuesto.' % mx(f.costo_cm_realizada)),
        ('Un solo creativo sostuvo la campaña completa',
         'El reel «video villas» absorbió %s del gasto y produjo %s de las conversaciones a %s cada una, '
         'frente a %s del promedio de los otros seis anuncios. Es una fortaleza que hay que explotar y, '
         'al mismo tiempo, una dependencia que hay que diversificar.'
         % (pc(0.700), pc(0.869), mx(9.98), mx(38.67))),
        ('Hay un inventario de prospectos ya pagados sin trabajar',
         'Al corte del 3 de octubre, %d oportunidades siguen abiertas y %d de ellas están en etapa de '
         'descubrimiento con 15 días o más sin ningún registro de movimiento. Recuperarlas no requiere '
         'inversión publicitaria nueva.' % (f.activos, f.dormidos)),
    ], start=1):
        d.h3('%d. %s' % (i, t))
        d.p(txt)
    d.pagebreak()


def contexto(d, f):
    d.h1('2. Periodo y contexto de la campaña')
    rows = [['Nombre', 'Campaña Septiembre 2026'],
            ['Inicio', '21 de agosto de 2026'],
            ['Fin', '4 de octubre de 2026'],
            ['Duración', '45 días'],
            ['Presupuesto asignado', mx(f.presupuesto, 0) + ' MXN'],
            ['Conjunto de anuncios', '«campaña agosto A» (un solo conjunto, presupuesto a nivel campaña)'],
            ['Anuncios activos', '8 creativos'],
            ['Objetivo de la pauta', 'Conversaciones de mensaje (Messenger / WhatsApp / Instagram Direct)']]
    d.table(['Concepto', 'Definición'], rows, widths=[2600, 6760],
            caption='Tabla 2. Ficha de la campaña')

    d.h2('2.1 Línea de tiempo')
    d.funnel([
        ('21 – 31 AGO', 'Arranque · %d leads · %d CM · %d inscripciones'
         % (f.fases['21–31 ago (arranque)']['leads'], f.fases['21–31 ago (arranque)']['cm_ag'],
            f.fases['21–31 ago (arranque)']['insc']), 0.26, '11 días'),
        ('SEPTIEMBRE', 'Periodo principal · %d leads · %d CM · %d inscripciones'
         % (f.fases['Septiembre (principal)']['leads'], f.fases['Septiembre (principal)']['cm_ag'],
            f.fases['Septiembre (principal)']['insc']), 1.0, '30 días'),
        ('1 – 4 OCT', 'Cierre · %d leads · %d CM · %d inscripciones'
         % (f.fases['1–4 oct (cierre)']['leads'], f.fases['1–4 oct (cierre)']['cm_ag'],
            f.fases['1–4 oct (cierre)']['insc']), 0.16, '4 días'),
    ])
    d.p('Los últimos once días de agosto funcionaron como arranque de la pauta, septiembre fue el periodo '
        'principal de captación y operación, y los primeros cuatro días de octubre fueron el cierre. '
        'Las tres fases pertenecen al mismo esfuerzo y no deben leerse como campañas independientes: '
        'la métrica principal de este reporte es siempre el acumulado %s.' % PER)
    d.callout('Criterio de atribución temporal',
              'Un prospecto captado el 28 de agosto que tomó su clase muestra el 3 de septiembre y se '
              'inscribió después forma parte de esta campaña. Igualmente, un prospecto de septiembre que '
              'concretó su clase o su inscripción entre el 1 y el 4 de octubre se contabiliza aquí. '
              'Las clases muestra se cuentan por la fecha de la clase y los leads por la fecha de primer '
              'contacto; %d registros del CRM (19 y 20 de agosto) quedan dos días antes del inicio de la '
              'pauta y están identificados por separado en la Base Maestra.' % f.reg_fuera)
    d.pagebreak()


def metodologia(d, f):
    d.h1('3. Metodología, atribución y homologación de registros')
    d.p('El análisis no trata cada base por separado: reconstruye el recorrido completo '
        'Publicidad → Lead → Seguimiento → Clase muestra → Inscripción, relacionando a la misma persona '
        'a través de las cuatro fuentes.')

    d.h2('3.1 Identificación de personas y eliminación de duplicados')
    d.p('Se normalizaron nombres (minúsculas, sin acentos ni signos) y teléfonos (últimos 10 dígitos) y '
        'se aplicaron reglas deterministas, documentadas una a una para que cualquier fusión sea auditable:')
    d.table(['Regla', 'Criterio', 'Decisión'], [
        ['R1', 'Mismo teléfono y mismo nombre normalizado', 'Misma persona — se fusiona'],
        ['R2', 'Mismo teléfono y un nombre contenido en el otro o muy similar',
         'Misma persona — se fusiona'],
        ['R3', 'Mismo nombre completo (2+ palabras) y ningún teléfono que lo contradiga',
         'Misma persona — se fusiona'],
        ['R4', 'Mismo teléfono con nombres distintos',
         'Personas distintas — se vinculan como mismo hogar/contacto'],
        ['R5', 'Mismo nombre de una sola palabra con teléfono distinto o ausente',
         'No se fusiona — se marca «Revisar duplicado»'],
        ['R6', 'Registro sin nombre y sin teléfono', 'No fusionable — se marca «No identificado»'],
    ], widths=[700, 4900, 3760], caption='Tabla 3. Reglas de deduplicación aplicadas')
    d.p('La regla R4 es clave en una academia: varios padres registraron a dos o tres hijos desde el mismo '
        'teléfono. Fusionarlos habría subestimado los prospectos; tratarlos como independientes sin '
        'vincularlos habría roto la atribución. Resultado: de %d registros del CRM se identificaron '
        '%d personas distintas (%d registros fusionados, %s de duplicidad) y %d registros quedaron '
        'marcados para revisión en lugar de fusionarse arbitrariamente.'
        % (f.reg_hoja, f.pers_total, f.dup_fusionados, pc(f.tasa_duplicidad), len(f.revisar)))

    d.h2('3.2 Enlace entre lead, clase muestra e inscripción')
    d.p('Cada clase muestra se intentó ligar a un lead del CRM. El nivel de confianza del enlace se '
        'registra explícitamente y solo los tres primeros niveles se usan para atribuir resultados a la pauta:')
    conf = f.M['cm_conf']
    d.table(['Confianza', 'Criterio', 'Clases muestra', 'Se usa para atribuir'], [
        ['Alta', 'Teléfono idéntico y nombre compatible', n(conf.get('Alta', 0)), 'Sí'],
        ['Media', 'Nombre completo compatible (nombre de pila + apellido)', n(conf.get('Media', 0)), 'Sí'],
        ['Hogar', 'Mismo teléfono con nombre distinto: el lead es el contacto familiar',
         n(conf.get('Hogar', 0)), 'Sí'],
        ['Baja', 'Única coincidencia por un nombre de pila', n(conf.get('Baja', 0)), 'No — se marca'],
        ['Sin enlace', 'No localizado en el CRM de campaña', n(conf.get('Sin enlace', 0)), 'No'],
    ], widths=[1200, 4800, 1600, 1760], aligns=['left', 'left', 'right', 'center'],
        caption='Tabla 4. Niveles de confianza del enlace clase muestra → lead (%d clases en ventana)'
                % f.cm_agendadas)
    d.p('Para validar las inscripciones se exigió además coincidencia de nombre de pila y al menos un '
        'apellido contra el padrón de alumnos, descartando coincidencias que solo compartían nombre de '
        'pila. Ese filtro evitó falsos positivos reales detectados en los datos, como «María Luisa '
        'Hernández» frente a «María Luisa Rojas Bravo».')

    d.h2('3.3 Homologación de estatus')
    d.p('Las fuentes usan vocabularios distintos. Se homologaron así:')
    d.table(['Valor en la fuente', 'Fuente', 'Estatus homologado'], [
        ['Nuevo lead', 'CRM', 'Nuevo'],
        ['Contactado · Información enviada', 'CRM', 'Descubrimiento'],
        ['Esperando respuesta', 'CRM', 'Seguimiento'],
        ['Pendiente de decisión', 'CRM', 'Por cerrar'],
        ['Cerrado + «Canceló proceso»', 'CRM', 'No interesado'],
        ['Resultado «Sin respuesta»', 'CRM', 'Sin respuesta'],
        ['Resultado «Inscrito»', 'CRM', 'Inscrito / Cerrado'],
        ['Cita', 'Bitácora CM', 'Clase muestra agendada'],
        ['Cita pospuesta', 'Bitácora CM', 'Reprogramación'],
        ['Show', 'Bitácora CM', 'Clase muestra realizada'],
        ['NO ASISTIÓ', 'Bitácora CM', 'Seguimiento (no asistió — se anota en observaciones)'],
        ['Inscrito', 'Bitácora CM', 'Inscrito / Cerrado'],
        ['Sin valor o valor no válido', 'Cualquiera', 'Estatus por validar'],
    ], widths=[3000, 1600, 4760], caption='Tabla 5. Homologación de estatus')
    d.callout('Lista de espera: no existe en las fuentes',
              'Ninguna de las cuatro fuentes contiene un campo o un valor que registre lista de espera. '
              'Los valores «En espera» del CRM significan que el seguimiento está pendiente de respuesta, '
              'no que haya un cupo reservado. Por eso el estatus «Lista de espera» aparece con 0 '
              'registros: no es un hallazgo, es una ausencia de dato.')
    d.pagebreak()


def funnel(d, f):
    d.h1('4. Funnel comercial')
    d.p('El embudo completo, de la impresión publicitaria a la inscripción. Las cantidades de las dos '
        'primeras etapas provienen de Meta Ads y las tres últimas de las bases internas; por eso la '
        'conversión entre impresiones y conversaciones debe leerse como indicador de entrega, no de venta.')
    d.funnel([
        ('Impresiones', n(f.impresiones), 1.0, '—'),
        ('Alcance', '≥ ' + n(f.alcance_min), 0.42, 'cota mín.'),
        ('Conversaciones', n(f.conversaciones), 0.17, pc(f.conversaciones / f.alcance_min)),
        ('Leads (personas)', n(f.leads), 0.145, pc(f.leads / f.conversaciones)),
        ('CM agendadas', n(f.cm_agendadas), 0.055, pc(f.t_lead_cm)),
        ('CM realizadas', n(f.cm_realizadas), 0.042, pc(f.t_cm_real)),
        ('Inscripciones', n(f.insc_personas), 0.022, pc(f.t_real_insc)),
    ])
    rows = []
    prev = None
    for nm, val, fu in f.etapas:
        if prev is None:
            rows.append([nm, n(val), '—', '—', '—', fu])
        else:
            rows.append([nm, n(val), pc(val / prev), n(prev - val), pc(1 - val / prev), fu])
        prev = val
    d.table(['Etapa', 'Cantidad', 'Conv. vs etapa previa', 'Pérdida absoluta', 'Pérdida %', 'Fuente'],
            rows, widths=[2100, 1150, 1250, 1150, 900, 2810],
            aligns=['left', 'right', 'right', 'right', 'right', 'left'],
            caption='Tabla 6. Funnel completo con pérdida por etapa')
    d.chart(GChart('bar', 'Funnel comercial de la campaña',
                   ['Conversaciones Meta', 'Leads (personas)', 'CM agendadas', 'CM realizadas',
                    'Inscripciones'],
                   [('Personas / eventos', [f.conversaciones, f.leads, f.cm_agendadas,
                                            f.cm_realizadas, f.insc_personas])],
                   colors=['1F3864']), height_cm=7.6)

    d.h2('4.1 Tasas clave')
    d.table(['Indicador', 'Fórmula', 'Resultado', 'Lectura comercial'], [
        ['Lead → CM agendada', 'CM agendadas / leads', pc(f.t_lead_cm),
         'De cada 100 interesados, solo %d llegaron a una clase muestra.' % round(f.t_lead_cm * 100)],
        ['CM agendada → realizada', 'CM realizadas / CM agendadas', pc(f.t_cm_real),
         'Buena asistencia: 3 de cada 4 clases agendadas se dieron.'],
        ['CM realizada → inscripción', 'Inscripciones / personas con CM realizada', pc(f.t_real_insc),
         'Casi una de cada dos personas que asistió se inscribió.'],
        ['Lead → inscripción', 'Inscripciones / leads', pc(f.t_lead_insc),
         'Conversión global del periodo.'],
        ['Lead → inscripción (atribuible)', 'Inscripciones atribuibles / leads', pc(f.t_lead_insc_atrib),
         'Conversión imputable exclusivamente a la pauta.'],
        ['No show', 'No show / CM agendadas', pc(f.tasa_noshow),
         '%d clases perdidas por inasistencia.' % f.cm_noshow],
        ['Reprogramación', 'Reprogramadas / CM agendadas', pc(f.tasa_reprog),
         'Solo %d caso: la reprogramación no es un problema.' % f.cm_reprogramadas],
    ], widths=[1900, 2100, 900, 4460], aligns=['left', 'left', 'right', 'left'],
        caption='Tabla 7. Tasas de conversión del embudo')
    d.chart(GChart('col', 'Conversión por etapa del embudo',
                   ['Lead → CM\nagendada', 'CM agendada →\nrealizada', 'CM realizada →\ninscripción',
                    'Lead →\ninscripción'],
                   [('Conversión', [round(f.t_lead_cm * 100, 1), round(f.t_cm_real * 100, 1),
                                    round(f.t_real_insc * 100, 1), round(f.t_lead_insc * 100, 1)])],
                   numfmt='0.0"%"', colors=['2E8BC0']), height_cm=7.2)

    d.h2('4.2 Dónde se pierde el embudo')
    p1 = f.leads - f.cm_personas
    p2 = f.cm_noshow
    p3 = f.cm_pers_realizadas - f.insc_personas
    d.table(['Punto de pérdida', 'Personas', '% de leads', 'Diagnóstico sustentado en datos'], [
        [('Lead que nunca agenda clase muestra', {'b': 1}), (n(p1), {'b': 1}),
         (pc(p1 / f.leads), {'b': 1}),
         ('CUELLO DE BOTELLA PRINCIPAL. %d de %d registros del CRM (%s) se quedaron en «Información '
          'enviada» o «Contactado», sin avanzar a ninguna etapa posterior.'
          % (f.M['leads_por_etapa'].get('Información enviada', 0)
             + f.M['leads_por_etapa'].get('Contactado', 0), f.reg_ventana,
             pc((f.M['leads_por_etapa'].get('Información enviada', 0)
                 + f.M['leads_por_etapa'].get('Contactado', 0)) / f.reg_ventana)), {'b': 1})],
        ['Clase agendada a la que no asiste', n(p2), pc(p2 / f.leads),
         'Pérdida secundaria y manejable: %s de no show, sin recordatorios ni confirmación '
         'registrados en las fuentes.' % pc(f.tasa_noshow)],
        ['Clase realizada sin inscripción', n(p3), pc(p3 / f.leads),
         'La tasa de cierre (%s) es el eslabón más sano del embudo; no es donde hay que intervenir primero.'
         % pc(f.t_real_insc)],
    ], widths=[2300, 900, 900, 5260], aligns=['left', 'right', 'right', 'left'],
        caption='Tabla 8. Pérdida del embudo por punto de fuga')
    d.callout('El problema no está en marketing',
              'Los datos descartan generación (%d conversaciones a %s cada una), asistencia (%s) y cierre '
              '(%s) como el problema principal. El %s de la pérdida total ocurre antes de que exista una '
              'clase muestra agendada, es decir, en la conversión de la conversación en cita. '
              'Es un problema de operación comercial y de cadencia de seguimiento.'
              % (f.conversaciones, mx(f.costo_conversacion), pc(f.t_cm_real), pc(f.t_real_insc),
                 pc(p1 / (p1 + p2 + p3))))
    d.pagebreak()


def clases(d, f):
    d.h1('5. Clases muestra')
    d.p('Se analizaron las %d clases muestra con fecha dentro del periodo, correspondientes a %d personas '
        'distintas. La bitácora es la única fuente que registra el cierre, por lo que es también la fuente '
        'de la inscripción.' % (f.cm_agendadas, f.cm_personas))
    d.table(['Estado de la clase muestra', 'Clases', '% de agendadas', 'Qué significa'], [
        ['Realizadas con inscripción', n(f.cm_inscritos_reg), pc(f.cm_inscritos_reg / f.cm_agendadas),
         'La clase terminó en alta'],
        ['Realizadas sin cierre', n(f.cm_realizadas - f.cm_inscritos_reg),
         pc((f.cm_realizadas - f.cm_inscritos_reg) / f.cm_agendadas),
         'Asistió pero aún no decide: oportunidad abierta'],
        ['No show', n(f.cm_noshow), pc(f.tasa_noshow), 'Agendada y no asistió'],
        ['Pendientes al 4 de octubre', n(f.cm_pendientes), pc(f.cm_pendientes / f.cm_agendadas),
         'Cita vigente al cierre del periodo'],
        ['Reprogramadas', n(f.cm_reprogramadas), pc(f.tasa_reprog), 'Cita pospuesta a nueva fecha'],
        [('Total agendadas', {'b': 1}), (n(f.cm_agendadas), {'b': 1}), ('100.0%', {'b': 1}), ''],
    ], widths=[2500, 900, 1100, 4860], aligns=['left', 'right', 'right', 'left'],
        caption='Tabla 9. Resultado de las clases muestra del periodo', total_row=True)
    d.chart(GChart('col', 'Clases muestra: agendadas → realizadas → inscripciones',
                   ['Agendadas', 'Realizadas', 'Con inscripción'],
                   [('Clases muestra', [f.cm_agendadas, f.cm_realizadas, f.cm_inscritos_reg])],
                   colors=['1F3864']), height_cm=7.2)
    d.p('Canceladas: las fuentes no registran un estado de cancelación; los casos que no se concretaron '
        'aparecen como «NO ASISTIÓ» o «Cita pospuesta». No es posible separar cancelación de no show '
        'con los datos disponibles.')

    d.h2('5.1 ¿Agendamiento, asistencia, reprogramación o cierre?')
    d.table(['Posible problema', 'Dato observado', 'Veredicto'], [
        ['Agendamiento', 'Lead → CM agendada = %s (%d de %d personas)'
         % (pc(f.t_lead_cm), f.cm_personas, f.leads), ('SÍ — es el problema principal', {'b': 1, 'color': ROJO})],
        ['Asistencia', 'CM realizadas / agendadas = %s' % pc(f.t_cm_real),
         ('No — está en nivel saludable', {'color': VERDE})],
        ['Reprogramación', '%d caso de %d clases (%s)' % (f.cm_reprogramadas, f.cm_agendadas, pc(f.tasa_reprog)),
         ('No — irrelevante', {'color': VERDE})],
        ['No show', '%d de %d clases (%s)' % (f.cm_noshow, f.cm_agendadas, pc(f.tasa_noshow)),
         ('Parcial — mejorable con confirmación', {'color': AMBAR})],
        ['Cierre', 'Inscripciones / personas con CM realizada = %s' % pc(f.t_real_insc),
         ('No — es la fortaleza del proceso', {'color': VERDE})],
    ], widths=[1700, 4200, 3460], caption='Tabla 10. Diagnóstico del proceso de clase muestra')

    d.h2('5.2 Horarios y canal de agenda')
    tot_h = sum(f.cm_hora.values())
    d.table(['Horario de la clase', 'Clases muestra', '% del total'],
            [[k, n(v), pc(v / tot_h)] for k, v in f.cm_hora.most_common(8)],
            widths=[2600, 1600, 1600], aligns=['left', 'right', 'right'],
            caption='Tabla 11. Horarios más demandados para clase muestra')
    d.p('Las 16:00 h concentran %s de las clases muestra y el bloque de 16:00 a 19:00 h agrupa %s del '
        'total. Es información directamente útil para dimensionar la disponibilidad de docentes.'
        % (pc(f.cm_hora.get('4:00pm', 0) / tot_h),
           pc(sum(v for k, v in f.cm_hora.items() if k in ('4:00pm', '5:00pm', '6:00pm', '7:00pm')) / tot_h)))
    tot_m = sum(f.cm_medio_ag.values())
    d.table(['Canal usado para agendar', 'Clases muestra', 'Inscripciones', 'Cierre'],
            [[k, n(v), n(f.insc_medio_ag.get(k, 0)),
              pc(f.insc_medio_ag.get(k, 0) / v) if v else '—'] for k, v in f.cm_medio_ag.most_common()],
            widths=[2600, 1400, 1400, 1400], aligns=['left', 'right', 'right', 'right'],
            caption='Tabla 12. Canal de agenda y su tasa de cierre')
    d.p('WhatsApp es el canal operativo real: %s de las clases muestra se agendaron por ahí. '
        'Facebook aparece con una tasa de cierre superior (%s frente a %s de WhatsApp), aunque sobre una '
        'base pequeña (%d clases), por lo que debe leerse como indicio y no como conclusión.'
        % (pc(f.cm_medio_ag.get('Whatsaap', 0) / tot_m),
           pc(f.insc_medio_ag.get('Facebook', 0) / f.cm_medio_ag.get('Facebook', 1)),
           pc(f.insc_medio_ag.get('Whatsaap', 0) / f.cm_medio_ag.get('Whatsaap', 1)),
           f.cm_medio_ag.get('Facebook', 0)))
    d.pagebreak()


def pauta(d, f):
    d.h1('6. Resultados de pauta')
    d.callout('Presupuesto asignado frente a gasto real',
              'Presupuesto asignado: %s MXN. Gasto real reportado por Meta: %s MXN (%s). '
              'Todos los indicadores de eficiencia de este reporte se calculan sobre el GASTO REAL, '
              'no sobre el presupuesto. El remanente fue de %s.'
              % (mx(f.presupuesto, 0), mx(f.gasto), pc(f.uso_presupuesto),
                 mx(f.presupuesto - f.gasto)))
    d.table(['Indicador de pauta', 'Valor', 'Observación'], [
        ['Gasto real', mx(f.gasto), 'Suma de los 8 anuncios'],
        ['Impresiones', n(f.impresiones), 'Suma por anuncio'],
        ['Alcance', '%s – %s' % (n(f.alcance_min), n(f.alcance_suma)),
         'Meta no exportó alcance único de campaña'],
        ['Frecuencia', '%.2f – %.2f' % (f.frec_sobre_suma, f.frec_sobre_min),
         'Impresiones / alcance, según la cota usada'],
        ['Conversaciones iniciadas', n(f.conversaciones), 'Resultado optimizado del objetivo'],
        ['CPM', mx(f.cpm), 'Gasto / impresiones × 1,000'],
        ['Costo por conversación', mx(f.costo_conversacion), 'Gasto / conversaciones'],
        ['Costo por lead (persona única)', mx(f.cpl_persona), 'Gasto / %d leads' % f.leads],
        ['Costo por CM agendada', mx(f.costo_cm_agendada), 'Gasto / %d clases' % f.cm_agendadas],
        ['Costo por CM realizada', mx(f.costo_cm_realizada), 'Gasto / %d clases' % f.cm_realizadas],
        ['CAC (atribuibles)', mx(f.cac), 'Gasto / %d inscripciones atribuibles' % f.insc_atribuibles],
        ['CAC (todas las inscripciones)', mx(f.cac_total), 'Gasto / %d inscripciones' % f.insc_personas],
        ['CTR', 'No calculable', 'La exportación no incluye clics'],
        ['CPC', 'No calculable', 'La exportación no incluye clics'],
    ], widths=[2800, 1700, 4860], aligns=['left', 'right', 'left'],
        caption='Tabla 13. Indicadores de pauta y eficiencia')
    d.small('Nota sobre el alcance: sumar el alcance de los 8 anuncios da %s personas, pero esa cifra '
            'contiene duplicación porque una misma persona pudo ver varios anuncios. El alcance único '
            'real de la campaña está entre %s (el alcance del anuncio de mayor entrega) y %s. '
            'No se reporta un número único porque la fuente no lo permite.'
            % (n(f.alcance_suma), n(f.alcance_min), n(f.alcance_suma)))
    d.pagebreak()

    d.h1('7. Anuncios y creativos: volumen frente a eficiencia')
    d.p('Los ocho anuncios compartieron un solo conjunto con presupuesto a nivel campaña, de modo que el '
        'reparto del gasto lo decidió el algoritmo de entrega y no una asignación manual. Eso explica la '
        'enorme concentración observada.')
    rows = []
    for t in sorted(f.tab_ads, key=lambda r: -r['gasto']):
        rows.append([t['anuncio'], t['formato'], mx(t['gasto']), pc(t['pct_gasto'], 1),
                     n(t['conv']), mx(t['costo_conv']) if t['costo_conv'] else '—',
                     n(t['impresiones']), n(t['alcance']), mx(t['cpm'])])
    rows.append([('TOTAL', {'b': 1}), '', (mx(f.gasto), {'b': 1}), ('100.0%', {'b': 1}),
                 (n(f.conversaciones), {'b': 1}), (mx(f.costo_conversacion), {'b': 1}),
                 (n(f.impresiones), {'b': 1}), (n(f.alcance_suma), {'b': 1}), (mx(f.cpm), {'b': 1})])
    d.table(['Anuncio', 'Formato', 'Gasto', '% gasto', 'Conv.', 'Costo/conv.', 'Impresiones',
             'Alcance', 'CPM'], rows,
            widths=[1500, 1000, 1050, 750, 650, 1000, 1200, 1050, 1160], font=16,
            aligns=['left', 'left', 'right', 'right', 'right', 'right', 'right', 'right', 'right'],
            caption='Tabla 14. Desempeño por anuncio (ordenado por gasto)', total_row=True)
    d.chart(GChart('col', 'Volumen frente a eficiencia por anuncio',
                   [t['anuncio'] for t in sorted(f.tab_ads, key=lambda r: -r['conv'])],
                   [('Conversaciones', [t['conv'] for t in sorted(f.tab_ads, key=lambda r: -r['conv'])])],
                   colors=['1F3864']), height_cm=7.0)
    d.chart(GChart('col', 'Costo por conversación (MXN) — menor es mejor',
                   [t['anuncio'] for t in sorted(f.tab_ads, key=lambda r: r['costo_conv'] or 0)],
                   [('Costo por conversación',
                     [round(t['costo_conv'], 2) for t in sorted(f.tab_ads, key=lambda r: r['costo_conv'] or 0)])],
                   numfmt='"$"#,##0.00', colors=['E0A030']), height_cm=7.0)

    d.h2('7.1 Ganadores y perdedores')
    v = f.ad_top_vol
    d.table(['Categoría', 'Anuncio', 'Evidencia'], [
        ['Mayor volumen de resultados', v['anuncio'],
         '%s conversaciones (%s del total) con %s de las impresiones'
         % (n(v['conv']), pc(v['pct_conv']), pc(v['impresiones'] / f.impresiones))],
        ['Más eficiente', f.ad_top_efic['anuncio'],
         '%s por conversación, el más bajo entre los anuncios con volumen significativo'
         % mx(f.ad_top_efic['costo_conv'])],
        ['Mayor alcance', max(f.tab_ads, key=lambda r: r['alcance'])['anuncio'],
         '%s personas alcanzadas' % n(max(f.tab_ads, key=lambda r: r['alcance'])['alcance'])],
        ['Mejor CPM', min(f.tab_ads, key=lambda r: r['cpm'])['anuncio'],
         '%s por mil impresiones' % mx(min(f.tab_ads, key=lambda r: r['cpm'])['cpm'])],
        ['Menos eficiente con gasto relevante', f.ad_peor_abs['anuncio'],
         '%s de gasto para %s conversaciones = %s cada una'
         % (mx(f.ad_peor_abs['gasto']), n(f.ad_peor_abs['conv']), mx(f.ad_peor_abs['costo_conv']))],
        ['Mejor formato', 'Video / Reel',
         'Los 5 videos concentran %s del gasto y %s de las conversaciones'
         % (pc(sum(t['gasto'] for t in f.tab_ads if t['formato'] == 'Video / Reel') / f.gasto),
            pc(sum(t['conv'] for t in f.tab_ads if t['formato'] == 'Video / Reel') / f.conversaciones))],
        ['Creativo ganador', v['anuncio'],
         'Único anuncio con clasificación «Por encima del promedio» en interacción y en conversiones '
         'de forma simultánea'],
    ], widths=[2300, 1700, 5360], caption='Tabla 15. Lectura de ganadores y perdedores')
    d.callout('Volumen no es eficiencia, pero esta vez coincidieron',
              '«%s» es a la vez el de mayor volumen y el más eficiente: %s por conversación frente a %s '
              'promedio de los otros seis anuncios de bajo gasto. Los %s (%s del gasto) repartidos entre '
              'esos seis anuncios produjeron solo %d conversaciones; al costo del anuncio ganador ese '
              'mismo dinero habría generado del orden de %d conversaciones. Es la oportunidad de '
              'eficiencia más grande y más fácil de capturar.'
              % (v['anuncio'], mx(9.98), mx(38.67), mx(1353.61), pc(0.193), 35, 136))
    d.p('Mensaje que mejor funcionó: el creativo ganador es un reel de geolocalización («Villas de '
        'Pachuca»), es decir, un mensaje de proximidad física. El segundo mejor («video sin libros», '
        '%s por conversación) es un mensaje de propuesta metodológica diferencial. Los creativos de '
        'oferta genérica con volumen comparable fueron los más caros: «post grupos» %s y «carrete» %s. '
        'Los datos no incluyen las piezas creativas, por lo que esta lectura se basa en el nombre del '
        'anuncio y en sus métricas, no en el contenido visual.' % (mx(19.09), mx(36.61), mx(44.99)))
    d.pagebreak()


def redes(d, f):
    d.h1('8. Redes sociales')
    d.callout('No hay datos orgánicos en el repositorio',
              'El repositorio no contiene ninguna exportación de métricas orgánicas de Facebook, '
              'Instagram o TikTok. No es posible analizar alcance orgánico, impresiones orgánicas, '
              'visualizaciones, interacciones, comentarios, compartidos, guardados, visitas al perfil, '
              'mensajes orgánicos ni crecimiento de seguidores, ni separar resultados orgánicos de '
              'pagados. Esta sección se declara NO CALCULABLE con los datos disponibles.',
              color='FBE9E7', barra=ROJO)
    d.p('Lo único que las fuentes permiten afirmar sobre canales es por dónde llegó cada conversación, '
        'tal como lo capturó el CRM:')
    tot = sum(f.leads_medio.values())
    d.table(['Medio de contacto declarado', 'Registros', '% del total', 'Interpretación'],
            [[k, n(v), pc(v / tot),
              {'Facebook': 'Conversación iniciada desde Facebook (Messenger o comentarios del anuncio)',
               'WhatsApp': 'Conversación entrante directa a WhatsApp, incluido el clic a WhatsApp del anuncio',
               'Instagram': 'Conversación iniciada desde Instagram Direct',
               'Sin información': 'Medio no capturado'}.get(k, '')]
             for k, v in f.leads_medio.most_common()],
            widths=[2200, 1100, 1100, 4960], aligns=['left', 'right', 'right', 'left'],
            caption='Tabla 16. Medio de contacto de los leads (no es medición orgánica)')
    d.p('El %s de los registros del CRM tiene identificado como origen uno de los ocho anuncios '
        'pagados, por lo que el volumen observado es esencialmente de pauta. No hay ningún registro '
        'cuyo origen esté marcado como publicación orgánica. Para la próxima campaña se requiere un '
        'export mensual de Meta Business Suite y de TikTok Analytics si se quiere evaluar el aporte '
        'orgánico.' % pc(1 - f.dq['leads_sin_anuncio'] / f.reg_ventana))
    d.pagebreak()


def inscripciones(d, f):
    d.h1('9. Inscripciones atribuibles al esfuerzo de campaña')
    d.p('Esta sección analiza únicamente las inscripciones del periodo y su origen. No debe confundirse '
        'con la base administrativa general de alumnos, que es el Entregable 4 y tiene corte al 3 de '
        'octubre.')
    d.kpi_grid([
        ('Inscripciones confirmadas', n(f.insc_personas), 'personas, %d eventos' % f.insc_eventos),
        ('Atribuibles a campaña', n(f.insc_atribuibles), pc(f.insc_atribuibles / f.insc_personas) + ' del total'),
        ('Sin origen de campaña', n(f.insc_personas - f.insc_atribuibles), 'recomendación, sitio o sin registro'),
        ('CAC atribuible', mx(f.cac), 'gasto / atribuibles'),
    ], cols=4)
    d.h2('9.1 Evidencia utilizada para cada inscripción')
    d.table(['Nivel de evidencia', 'Personas', 'Detalle'], [
        ['Bitácora de clases muestra marcada «Inscrito» y localizada en el padrón',
         n(f.insc_personas - len(f.insc_sin_padron) - len(f.insc_conflicto)), 'Doble evidencia'],
        ['Bitácora marcada «Inscrito» pero no localizada en el padrón del 05-oct',
         n(len(f.insc_sin_padron)),
         ', '.join(r['nombre'] for r in f.insc_sin_padron) + ' — requieren validación administrativa'],
        ['CRM marcado «Inscrito» y presente en el padrón, aunque la bitácora dice «Show»',
         n(len(f.insc_conflicto)),
         ', '.join(r['nombre'] for r in f.insc_conflicto) + ' — se contabiliza por doble evidencia'],
        [('Total inscripciones confirmadas', {'b': 1}), (n(f.insc_personas), {'b': 1}), ''],
    ], widths=[4300, 900, 4160], aligns=['left', 'right', 'left'],
        caption='Tabla 17. Trazabilidad de las inscripciones del periodo', total_row=True)
    if f.insc_por_validar:
        d.p('Además se detectaron %d personas que aparecen en el padrón con matrícula reciente y que '
            'tomaron clase muestra en el periodo, pero cuya bitácora no las marca inscritas y cuyo CRM '
            'no lo confirma: %s. NO se contabilizaron como inscripciones; se reportan para validación '
            'administrativa. Si se confirmaran, las inscripciones del periodo subirían a %d.'
            % (len(f.insc_por_validar),
               '; '.join('%s (matrícula %s)' % (r['nombre'], r['matricula']) for r in f.insc_por_validar),
               f.insc_personas + len(f.insc_por_validar)))

    d.h2('9.2 Segmentación de las inscripciones')
    d.table(['Programa', 'CM agendadas', 'CM realizadas', 'Inscripciones', 'Asistencia', 'Cierre tras CM'],
            [[k, n(v[0]), n(v[1]), n(v[2]), pc(v[1] / v[0]), pc(v[2] / v[1]) if v[1] else '—']
             for k, v in sorted(f.conv_cat.items(), key=lambda kv: -kv[1][0])],
            widths=[2000, 1500, 1500, 1400, 1400, 1560],
            aligns=['left', 'right', 'right', 'right', 'right', 'right'],
            caption='Tabla 18. Clases muestra e inscripciones por programa')
    d.chart(GChart('col', 'Clases muestra e inscripciones por programa',
                   [k for k, v in sorted(f.conv_cat.items(), key=lambda kv: -kv[1][0])],
                   [('CM agendadas', [v[0] for k, v in sorted(f.conv_cat.items(), key=lambda kv: -kv[1][0])]),
                    ('CM realizadas', [v[1] for k, v in sorted(f.conv_cat.items(), key=lambda kv: -kv[1][0])]),
                    ('Inscripciones', [v[2] for k, v in sorted(f.conv_cat.items(), key=lambda kv: -kv[1][0])])]),
            height_cm=7.4)
    d.p('College genera el mayor volumen de clases muestra (%d) pero tiene la tasa de cierre más baja '
        '(%s). Kids, Junior y Parenthood cierran entre %s y %s. Es un dato operativo relevante: la '
        'demanda adulta es más fácil de atraer y más difícil de cerrar, mientras que la demanda familiar '
        'es más escasa pero más rentable por clase muestra.'
        % (f.conv_cat['College'][0], pc(f.conv_cat['College'][2] / f.conv_cat['College'][1]),
           pc(f.conv_cat['Junior'][2] / f.conv_cat['Junior'][1]),
           pc(f.conv_cat['Kids'][2] / f.conv_cat['Kids'][1])))
    d.table(['Idioma', 'CM agendadas', 'Inscripciones'],
            [[k, n(v), n(f.insc_idioma.get(k, 0))] for k, v in f.cm_idioma.most_common()],
            widths=[2600, 1600, 1600], aligns=['left', 'right', 'right'],
            caption='Tabla 19. Clases muestra e inscripciones por idioma')
    d.p('El francés aparece con %d clases muestra y %d inscripciones, todas concentradas en el 21 de '
        'septiembre y todas de alumnos que ya habían tomado clase muestra de inglés. Es demanda cruzada '
        'de familias existentes, no captación nueva: una vía de ingreso adicional que hoy no se promueve '
        'en la pauta.' % (f.cm_idioma.get('Francés', 0), f.insc_idioma.get('Francés', 0)))
    d.p('Edad de los inscritos: de %d a %d años, promedio %.1f. El rango confirma que la academia atiende '
        'tanto a niños como a adultos y que la pauta trajo ambos perfiles.'
        % (f.edad_insc['min'], f.edad_insc['max'], f.edad_insc['prom']))

    d.h2('9.3 Las inscripciones no atribuibles a la pauta')
    d.table(['Persona', 'Canal de agenda', 'Por qué no es atribuible'],
            [[r['nombre'], r['medio_agenda'] or 'Sin información',
              'Registrada explícitamente como recomendación'
              if r['medio_agenda'] == 'Recomendación' else
              ('Agendó por Facebook pero no existe registro suyo en el CRM de campaña'
               if r['medio_agenda'] == 'Facebook' else
               'Agendó por WhatsApp pero no existe registro suyo en el CRM de campaña')]
             for r in sorted(f.insc_no_atribuibles, key=lambda r: r['nombre'])],
            widths=[2800, 1800, 4760], caption='Tabla 20. Inscripciones del periodo sin origen de campaña')
    d.p('Seis de estas siete inscripciones llegaron por canales que la academia no está midiendo. '
        'Solo una está explícitamente marcada como recomendación. Esto tiene dos implicaciones: el CAC '
        'real de la pauta es %s y no %s, y existe un canal de referidos que produce altas sin ningún '
        'seguimiento formal.' % (mx(f.cac), mx(f.cac_total)))
    d.pagebreak()


def prospectos(d, f):
    d.h1('10. Prospectos, estatus y eficiencia comercial')
    d.p('Pipeline consolidado al corte del 3 de octubre de 2026: %d registros de prospecto, resultado de '
        'unir el CRM de campaña deduplicado con las personas que solo aparecen en la bitácora de clases '
        'muestra. El detalle individual está en el Entregable 5.' % len(f.pipeline))
    tot = len(f.pipeline)
    d.table(['Estatus homologado', 'Prospectos', '% del total', 'Oportunidad abierta'],
            [[k, n(f.estatus_count.get(k, 0)), pc(f.estatus_count.get(k, 0) / tot),
              'Sí' if k in facts.ACTIVOS else 'No'] for k in facts.ESTATUS_ORDEN],
            widths=[2800, 1300, 1300, 1800], aligns=['left', 'right', 'right', 'center'],
            caption='Tabla 21. Pipeline comercial por estatus al 3 de octubre de 2026')
    d.chart(GChart('bar', 'Prospectos por estatus al corte del 3 de octubre',
                   [k for k in facts.ESTATUS_ORDEN if f.estatus_count.get(k, 0) > 0],
                   [('Prospectos', [f.estatus_count.get(k, 0) for k in facts.ESTATUS_ORDEN
                                    if f.estatus_count.get(k, 0) > 0])],
                   colors=['2E5C9A']), height_cm=7.8)
    d.callout('El tamaño del pipeline esconde su calidad',
              'De los %d registros en estados abiertos, %d están en «Descubrimiento» con 15 días o más '
              'sin ningún movimiento registrado. Son prospectos ya pagados que no están siendo '
              'trabajados. A la conversión observada de la propia campaña (%s de lead a inscripción), '
              'ese inventario representa del orden de %d inscripciones potenciales. Es una proyección '
              'con la tasa histórica de la campaña, no un dato confirmado.'
              % (f.activos, f.dormidos, pc(f.t_lead_insc), round(f.dormidos * f.t_lead_insc)))

    d.h2('10.1 Eficiencia comercial: dónde estuvo el reto')
    d.table(['Posible reto', 'Dato que lo mide', 'Veredicto'], [
        ['Generación', '%d conversaciones y %d leads-persona con %s de inversión'
         % (f.conversaciones, f.leads, mx(f.gasto)), ('Descartado', {'color': VERDE, 'b': 1})],
        ['Calidad del lead', '%s de las personas que llegaron a clase muestra asistieron y %s de ellas '
                             'se inscribieron' % (pc(f.t_cm_real), pc(f.t_real_insc)),
         ('Descartado', {'color': VERDE, 'b': 1})],
        ['Contacto', '%d de %d registros sin teléfono; %d sin nombre ni teléfono'
         % (f.dq['leads_sin_tel'], f.reg_ventana, f.dq['leads_sin_nombre_ni_tel']),
         ('Contribuye', {'color': AMBAR, 'b': 1})],
        ['Agendamiento', 'Solo %s de los leads llegó a clase muestra: %d personas perdidas'
         % (pc(f.t_lead_cm), f.leads - f.cm_personas),
         ('RETO PRINCIPAL', {'color': ROJO, 'b': 1})],
        ['Asistencia', '%s de no show (%d clases)' % (pc(f.tasa_noshow), f.cm_noshow),
         ('Reto secundario', {'color': AMBAR, 'b': 1})],
        ['Seguimiento', '%d registros sin resultado de seguimiento (%s) y %d prospectos activos con 15+ '
                        'días sin movimiento' % (f.dq['leads_sin_resultado'],
                                                 pc(f.dq['leads_sin_resultado'] / f.reg_ventana), f.dormidos),
         ('RETO PRINCIPAL', {'color': ROJO, 'b': 1})],
        ['Cierre', '%s de las personas con clase muestra realizada se inscribieron' % pc(f.t_real_insc),
         ('Descartado — es la fortaleza', {'color': VERDE, 'b': 1})],
    ], widths=[1700, 4800, 2860], caption='Tabla 22. Diagnóstico de eficiencia comercial')
    d.p('La conclusión no es interpretativa: es aritmética. El %s de toda la pérdida del embudo ocurre '
        'antes de que exista una clase muestra agendada, y el CRM no registra resultado de seguimiento '
        'en el %s de sus registros. Agendamiento y seguimiento son el mismo problema visto desde dos '
        'ángulos.'
        % (pc((f.leads - f.cm_personas) / (f.leads - f.insc_personas)),
           pc(f.dq['leads_sin_resultado'] / f.reg_ventana)))
    d.pagebreak()


def temporal(d, f):
    d.h1('11. Análisis temporal')
    d.p('Todo pertenece a la misma campaña. Esta sección es secundaria y sirve para entender el '
        'comportamiento en el tiempo, no para comparar tres campañas.')
    d.table(['Fase', 'Días', 'Leads', 'Leads/día', 'CM agendadas', 'CM realizadas', 'Inscripciones'],
            [[k, n(v['dias']), n(v['leads']), '%.1f' % (v['leads'] / v['dias']), n(v['cm_ag']),
              n(v['cm_re']), n(v['insc'])] for k, v in f.fases.items()]
            + [[('TOTAL CAMPAÑA', {'b': 1}), ('45', {'b': 1}), (n(f.leads), {'b': 1}),
                ('%.1f' % (f.leads / 45), {'b': 1}), (n(f.cm_agendadas), {'b': 1}),
                (n(f.cm_realizadas), {'b': 1}), (n(f.insc_eventos), {'b': 1})]],
            widths=[2400, 800, 1000, 1100, 1400, 1400, 1260],
            aligns=['left', 'right', 'right', 'right', 'right', 'right', 'right'],
            caption='Tabla 23. Resultados por fase de la campaña', total_row=True)
    d.table(['Semana', 'Leads', 'CM agendadas', 'CM realizadas', 'Inscripciones', 'Lead → CM'],
            [[lab, n(l), n(ca), n(cr), n(ins), pc(ca / l) if l else '—']
             for lab, l, ca, cr, ins in f.semanas],
            widths=[1800, 1400, 1700, 1700, 1500, 1260],
            aligns=['left', 'right', 'right', 'right', 'right', 'right'],
            caption='Tabla 24. Evolución semanal de la campaña')
    d.chart(GChart('line', 'Evolución semanal: leads, clases muestra e inscripciones',
                   [lab for lab, l, ca, cr, ins in f.semanas],
                   [('Leads', [l for lab, l, ca, cr, ins in f.semanas]),
                    ('CM agendadas', [ca for lab, l, ca, cr, ins in f.semanas]),
                    ('Inscripciones', [ins for lab, l, ca, cr, ins in f.semanas])]), height_cm=7.4)
    d.callout('Patrón relevante: a mayor volumen semanal de leads, menor tasa de agendamiento',
              'La tasa Lead → CM pasa de %s en la semana de mayor volumen (%d leads) a %s en la de menor '
              'volumen (%d leads). La relación es consistente a lo largo de las siete semanas. '
              'Se trata de tasas por semana calendario, no por cohorte de lead, por lo que el patrón es '
              'indicativo y no una medición causal; aun así apunta a que la capacidad de atención, y no '
              'la demanda, fue el factor limitante.'
              % (pc(f.semanas[1][2] / f.semanas[1][1]), f.semanas[1][1],
                 pc(f.semanas[5][2] / f.semanas[5][1]), f.semanas[5][1]))
    d.p('El pico de captación fue el %s con %d registros en un solo día, equivalente al %s de todo el '
        'periodo. Ese día no tiene un incremento proporcional de clases muestra en los días siguientes.'
        % (f.pico_leads[0].strftime('%d de septiembre de 2026'), f.pico_leads[1],
           pc(f.pico_leads[1] / f.reg_ventana)))
    d.pagebreak()


def hallazgos(d, f):
    d.h1('12. Principales hallazgos')
    d.p('Estructura: DATO → INTERPRETACIÓN → IMPACTO COMERCIAL.')
    H = [
        ('Un solo creativo sostuvo toda la campaña',
         'El reel «video villas» consumió %s del gasto (%s), generó %s de las impresiones (%s) y %s de '
         'las conversaciones (%d de %d), a %s por conversación. Es el único anuncio con clasificación '
         '«Por encima del promedio» simultáneamente en interacción y en conversiones.'
         % (pc(0.700), mx(4901.80), pc(0.830), n(200179), pc(0.869), 491, f.conversaciones, mx(9.98)),
         'Con presupuesto a nivel campaña, el algoritmo identificó al ganador y concentró la entrega. '
         'La campaña no funcionó porque hubiera ocho buenos anuncios: funcionó porque uno fue muy bueno.',
         '%d de las %d inscripciones atribuibles (%s) provienen de ese creativo. Es la mayor fortaleza '
         'de la campaña y, a la vez, su mayor riesgo: si ese reel se satura, no hay segundo creativo '
         'probado que lo sustituya.'
         % (f.insc_ad.get('Reel Villas de Pachuca', 0), f.insc_atribuibles,
            pc(f.insc_ad.get('Reel Villas de Pachuca', 0) / f.insc_atribuibles))),

        ('El cuello de botella está entre el lead y la clase muestra, no en marketing',
         'De %d leads-persona, solo %d llegaron a tener clase muestra (%s). Se perdieron %d personas '
         '(%s) antes de agendar. En contraste, la asistencia fue de %s y el cierre tras la clase de %s.'
         % (f.leads, f.cm_personas, pc(f.t_lead_cm), f.leads - f.cm_personas,
            pc((f.leads - f.cm_personas) / f.leads), pc(f.t_cm_real), pc(f.t_real_insc)),
         'La generación fue abundante y barata, y el proceso de clase muestra funciona muy bien. '
         'El eslabón roto es la conversión de la conversación en cita agendada.',
         'Con la misma pauta y el mismo gasto, subir Lead → CM de %s a 25%% habría producido unas 28 '
         'clases muestra adicionales y, a las tasas de asistencia y cierre observadas, del orden de '
         '10 inscripciones más. Sería un incremento del %s sobre el resultado obtenido, sin un peso '
         'adicional de inversión.'
         % (pc(f.t_lead_cm), pc(10.5 / f.insc_personas))),

        ('%d oportunidades abiertas y %d de ellas sin movimiento desde hace 15 días o más'
         % (f.activos, f.dormidos),
         'Al corte del 3 de octubre hay %d registros en estados abiertos. De ellos, %d están en '
         '«Descubrimiento» con 15 o más días sin ningún registro nuevo. Además, el CRM no tiene '
         'resultado de seguimiento en %d de %d registros (%s) y la columna de comentarios está vacía en '
         'el 100%% de la hoja.'
         % (f.activos, f.dormidos, f.dq['leads_sin_resultado'], f.reg_ventana,
            pc(f.dq['leads_sin_resultado'] / f.reg_ventana)),
         'No es un problema de volumen ni de interés del mercado: es un problema de cadencia y de '
         'documentación del seguimiento. Un prospecto sin registro de seguimiento es, en la práctica, '
         'un prospecto no trabajado.',
         'Ese inventario ya está pagado. Trabajarlo no requiere inversión publicitaria nueva. A la '
         'conversión observada de la propia campaña representa del orden de %d inscripciones '
         'potenciales (proyección con tasa propia, no dato confirmado).'
         % round(f.dormidos * f.t_lead_insc)),

        ('La clase muestra es el activo comercial más rentable de la academia',
         '%d de %d clases agendadas se realizaron (%s) y %d de las %d personas que asistieron se '
         'inscribieron (%s). El no show fue de %s y solo hubo %d reprogramación.'
         % (f.cm_realizadas, f.cm_agendadas, pc(f.t_cm_real), f.insc_personas, f.cm_pers_realizadas,
            pc(f.t_real_insc), pc(f.tasa_noshow), f.cm_reprogramadas),
         'Cuando el prospecto entra al salón, cierra casi la mitad de las veces. El proceso pedagógico '
         'y comercial de la clase muestra está resuelto.',
         'Cada clase muestra realizada costó %s de pauta y produjo 0.45 inscripciones. Toda acción que '
         'incremente el número de clases muestra tiene retorno casi directo, lo que convierte al '
         'agendamiento en la prioridad número uno de inversión de esfuerzo.'
         % mx(f.costo_cm_realizada)),

        ('%s del gasto se fue en seis anuncios que aportaron %s de los resultados'
         % (pc(0.193), pc(0.062)),
         '%s (%s del gasto) se repartieron entre los seis anuncios distintos de «video villas» y '
         '«video sin libros». Produjeron %d conversaciones (%s del total) a %s cada una, frente a %s '
         'del anuncio ganador: 3.9 veces más caro.'
         % (mx(1353.61), pc(0.193), 35, pc(0.062), mx(38.67), mx(9.98)),
         'El presupuesto a nivel campaña reparte, pero no elimina: los anuncios de bajo rendimiento '
         'siguieron consumiendo entrega durante todo el periodo sin aportar resultados.',
         'Reasignar ese %s al creativo eficiente habría generado del orden de %d conversaciones en '
         'lugar de %d, es decir unas 100 conversaciones adicionales. Al CPL observado equivale a unas '
         '85 personas más en el embudo.' % (pc(0.193), 136, 35)),

        ('%d de las %d inscripciones del periodo no provienen de la pauta y seis no se están midiendo'
         % (f.insc_personas - f.insc_atribuibles, f.insc_personas),
         '%d de %d inscripciones (%s) son atribuibles a la campaña. De las %d restantes, una está '
         'marcada explícitamente como «Recomendación», tres corresponden a una familia que agendó por '
         'Facebook sin registro en el CRM y tres llegaron por WhatsApp sin registro.'
         % (f.insc_atribuibles, f.insc_personas, pc(f.insc_atribuibles / f.insc_personas),
            f.insc_personas - f.insc_atribuibles),
         'Existen canales productivos —recomendación y contacto directo— que no tienen ningún registro '
         'formal. La atribución subestima lo que la pauta genera y, al mismo tiempo, oculta un canal '
         'orgánico que funciona.',
         'El CAC atribuible es %s frente a %s si se consideran todas las inscripciones. Formalizar el '
         'registro de referidos permitiría medir y escalar un canal que hoy produce altas sin costo '
         'publicitario.' % (mx(f.cac), mx(f.cac_total))),

        ('La demanda adulta atrae más y cierra menos; la demanda familiar es escasa y cierra mejor',
         'College generó %d clases muestra (%s del total) con un cierre de %s. Kids cerró %s, '
         'Parenthood %s y Junior %s.'
         % (f.conv_cat['College'][0], pc(f.conv_cat['College'][0] / f.cm_agendadas),
            pc(f.conv_cat['College'][2] / f.conv_cat['College'][1]),
            pc(f.conv_cat['Kids'][2] / f.conv_cat['Kids'][1]),
            pc(f.conv_cat['Parenthood'][2] / f.conv_cat['Parenthood'][1]),
            pc(f.conv_cat['Junior'][2] / f.conv_cat['Junior'][1])),
         'El mensaje de la campaña atrajo predominantemente adultos que deciden solos y tardan más en '
         'comprometerse. Las familias llegan menos pero deciden en bloque: varios casos inscribieron a '
         'dos o tres integrantes.',
         'Tres contactos del pipeline concretaron dos inscripciones cada uno. Orientar parte del '
         'presupuesto a mensajes familiares eleva el valor por clase muestra sin aumentar el CPL.'),

        ('Hay demanda cruzada de un segundo idioma que no se está promoviendo',
         'Las %d clases muestra de francés y sus %d inscripciones se concentran en un solo día '
         '(21 de septiembre) y todas corresponden a personas que ya habían tomado clase muestra de '
         'inglés. Dos alumnos tienen dos matrículas en el padrón, una por idioma.'
         % (f.cm_idioma.get('Francés', 0), f.insc_idioma.get('Francés', 0)),
         'El francés no captó demanda nueva: capitalizó demanda existente. Es venta cruzada, no '
         'adquisición.',
         '%s de las inscripciones del periodo vinieron de venta cruzada sin inversión publicitaria '
         'dedicada. Es el camino más económico para subir ingreso por alumno.'
         % pc(f.insc_idioma.get('Francés', 0) / f.insc_eventos)),
    ]
    for i, (t, dato, inter, imp) in enumerate(H, start=1):
        d.h2('Hallazgo %d. %s' % (i, t))
        d.table(['', ''], [
            [('DATO', {'b': 1, 'color': AZUL}), dato],
            [('INTERPRETACIÓN', {'b': 1, 'color': AZUL}), inter],
            [('IMPACTO COMERCIAL', {'b': 1, 'color': AZUL}), imp],
        ], widths=[1800, 7560], zebra=False, no_header=True)
    d.pagebreak()


def funciono(d, f):
    d.h1('13. Qué funcionó')
    S = [
        ('Pauta',
         ['Ejecución presupuestal prácticamente perfecta: %s de %s (%s), sin subejecución.'
          % (mx(f.gasto), mx(f.presupuesto, 0), pc(f.uso_presupuesto)),
          'Costo por conversación de %s y CPM de %s: costos de entrada bajos para el sector.'
          % (mx(f.costo_conversacion), mx(f.cpm)),
          'El objetivo de mensajes fue el correcto: %s de las clases muestra se agendaron por WhatsApp, '
          'el mismo canal al que llevaba la pauta.'
          % pc(f.cm_medio_ag.get('Whatsaap', 0) / sum(f.cm_medio_ag.values()))]),
        ('Creativos',
         ['El formato video/reel concentró %s de las conversaciones.'
          % pc(sum(t['conv'] for t in f.tab_ads if t['formato'] == 'Video / Reel') / f.conversaciones),
          'El mensaje de proximidad geográfica («Villas de Pachuca») fue el más eficiente de la campaña: '
          '%s por conversación.' % mx(9.98),
          'El segundo mejor mensaje fue de diferenciación metodológica («sin libros»), %s por '
          'conversación: la propuesta de valor vende mejor que la oferta genérica.' % mx(19.09)]),
        ('Contenido',
         ['Los dos creativos ganadores comparten tres características verificables en los datos: '
          'formato video, mensaje específico (lugar o método) y clasificación de interacción por encima '
          'del promedio.',
          'Los creativos de oferta genérica fueron los más caros por conversación: «post grupos» %s '
          'con 11 conversaciones, «carrete» %s con 4 y «video clases» %s con 7. «post clase» y '
          '«video switch» registran costos menores pero con una sola conversación cada uno, por lo '
          'que no son comparables.' % (mx(36.61), mx(44.99), mx(59.65))]),
        ('Generación de prospectos',
         ['%d conversaciones y %d personas identificables en 45 días.' % (f.conversaciones, f.leads),
          'Costo por lead de %s a nivel persona única.' % mx(f.cpl_persona),
          'Captación sostenida: las siete semanas del periodo tuvieron leads, con pico de %d en un solo día.'
          % f.pico_leads[1]]),
        ('Seguimiento',
         ['El equipo registró el anuncio de origen en %s de los registros, lo que permitió '
          'reconstruir la atribución creativo por creativo.'
          % pc(1 - f.dq['leads_sin_anuncio'] / f.reg_ventana),
          'Cuando el seguimiento se concretó en una cita, funcionó: %s de asistencia.' % pc(f.t_cm_real),
          'Nota: el seguimiento es también un área de oportunidad grave (sección 14); lo que funcionó '
          'fue la ejecución de la cita, no la cadencia previa.']),
        ('Clases muestra',
         ['%s de asistencia sobre clases agendadas.' % pc(f.t_cm_real),
          'Solo %d reprogramación en todo el periodo (%s).' % (f.cm_reprogramadas, pc(f.tasa_reprog)),
          'Programación concentrada en los horarios de mayor demanda: %s de las clases entre 16:00 y 19:00 h.'
          % pc(sum(v for k, v in f.cm_hora.items()
                   if k in ('4:00pm', '5:00pm', '6:00pm', '7:00pm')) / sum(f.cm_hora.values()))]),
        ('Conversión',
         ['Cierre tras clase muestra de %s: casi una de cada dos personas que asistió se inscribió.'
          % pc(f.t_real_insc),
          'Tres contactos cerraron dos inscripciones cada uno (hermanos o madre e hijo).',
          '%d inscripciones en el periodo con %s de inversión.' % (f.insc_personas, mx(f.gasto))]),
    ]
    for t, items in S:
        d.h2(t)
        for it in items:
            d.bullet(it)
    d.pagebreak()


def oportunidades(d, f):
    d.h1('14. Áreas de oportunidad')
    d.p('Cada problema está vinculado al dato que lo evidencia. No se incluye ninguna observación '
        'genérica sin soporte.')
    rows = [
        ['Agendamiento bajo', 'Lead → CM = %s. %d de %d personas nunca llegaron a clase muestra.'
         % (pc(f.t_lead_cm), f.leads - f.cm_personas, f.leads),
         'Es la pérdida más grande del embudo y la de mayor retorno si se corrige.'],
        ['Seguimiento sin cadencia ni registro',
         '%d de %d registros (%s) sin resultado de seguimiento; columna de comentarios vacía en el 100%% '
         'de la hoja; %d prospectos activos con 15+ días sin movimiento.'
         % (f.dq['leads_sin_resultado'], f.reg_ventana,
            pc(f.dq['leads_sin_resultado'] / f.reg_ventana), f.dormidos),
         'Imposible saber cuántos toques recibió cada prospecto ni por qué se perdió.'],
        ['Gasto concentrado en anuncios ineficientes',
         '%s (%s del gasto) en seis anuncios que aportaron %d conversaciones a %s cada una.'
         % (mx(1353.61), pc(0.193), 35, mx(38.67)),
         'Dinero gastado a 3.9 veces el costo del anuncio ganador.'],
        ['No show del %s' % pc(f.tasa_noshow),
         '%d de %d clases agendadas no se realizaron por inasistencia. Las fuentes no registran '
         'recordatorios ni confirmaciones.' % (f.cm_noshow, f.cm_agendadas),
         'Cada no show consume un espacio de agenda y %s de pauta.' % mx(f.costo_cm_agendada)],
        ['Pérdida de trazabilidad entre Meta y el CRM',
         'Meta reporta %d conversaciones; el CRM capturó %d registros. %d conversaciones no dejaron rastro.'
         % (f.conversaciones, f.reg_ventana, f.gap_registro),
         'Hasta el %s de la demanda pagada puede estar perdiéndose antes de entrar al proceso.'
         % pc(1 - f.cobertura_crm)],
        ['Datos de perfil capturados demasiado tarde',
         '%s de los registros sin idioma de interés y %s sin segmento. El perfil solo se conoce cuando '
         'el prospecto llega a clase muestra.'
         % (pc(f.dq['leads_sin_idioma'] / f.reg_ventana), pc(f.dq['leads_sin_segmento'] / f.reg_ventana)),
         'No se puede priorizar ni segmentar el seguimiento por valor potencial.'],
        ['Baja conversión del segmento College',
         'Cierre tras clase muestra de %s frente a %s de Kids, sobre el segmento de mayor volumen '
         '(%d clases).' % (pc(f.conv_cat['College'][2] / f.conv_cat['College'][1]),
                           pc(f.conv_cat['Kids'][2] / f.conv_cat['Kids'][1]), f.conv_cat['College'][0]),
         'El segmento que más cuesta atraer es el que menos cierra.'],
        ['Asistencia desigual por segmento',
         'College %s de asistencia y Kids %s, frente a %s de Junior y %s de Parenthood.'
         % (pc(f.conv_cat['College'][1] / f.conv_cat['College'][0]),
            pc(f.conv_cat['Kids'][1] / f.conv_cat['Kids'][0]),
            pc(f.conv_cat['Junior'][1] / f.conv_cat['Junior'][0]),
            pc(f.conv_cat['Parenthood'][1] / f.conv_cat['Parenthood'][0])),
         'El no show se concentra en los dos segmentos de mayor volumen.'],
        ['Padrón de alumnos sin campos de gestión',
         'El padrón no tiene fecha de inscripción, programa, idioma ni grupo. %d pares de matrículas '
         'con nombre, sexo y edad idénticos.' % len(f.dq['rga_dup']),
         'No permite cortes por fecha, cohortes ni conciliación con la campaña.'],
        ['Sin medición orgánica',
         'No existe ninguna exportación de métricas de Facebook, Instagram o TikTok en el repositorio.',
         'Se desconoce por completo el aporte del contenido no pagado.'],
    ]
    d.table(['Área de oportunidad', 'Dato que lo evidencia', 'Por qué importa'], rows,
            widths=[2100, 4300, 2960], caption='Tabla 25. Áreas de oportunidad con su evidencia')
    d.pagebreak()


def recomendaciones(d, f):
    d.h1('15. Recomendaciones para la siguiente campaña')
    d.p('Acciones derivadas del comportamiento observado en esta campaña, priorizadas por impacto '
        'esperado sobre el embudo.')
    R = [
        ('Alta', 'Responder y agendar en menos de 30 minutos durante las horas de mayor entrada, con '
                 'un responsable asignado por turno.',
         'Lead → CM de %s: %d personas se perdieron antes de agendar' % (pc(f.t_lead_cm), f.leads - f.cm_personas),
         'Lead → CM agendada'),
        ('Alta', 'Cambiar el guion del primer mensaje: proponer dos horarios concretos de clase muestra '
                 'en lugar de enviar información. Los horarios de mayor demanda observados son 16:00, '
                 '18:00 y 09:00 h.',
         '%d de %d registros se quedaron en «Información enviada» o «Contactado»'
         % (f.M['leads_por_etapa'].get('Información enviada', 0)
            + f.M['leads_por_etapa'].get('Contactado', 0), f.reg_ventana),
         'Lead → CM agendada'),
        ('Alta', 'Campaña de recuperación sobre los %d prospectos en «Descubrimiento» sin movimiento, '
                 'antes de abrir presupuesto nuevo.' % f.dormidos,
         '%d oportunidades abiertas de las cuales %d llevan 15+ días sin registro'
         % (f.activos, f.dormidos), 'Inscripciones sin costo incremental'),
        ('Alta', 'Reasignar el presupuesto de los seis anuncios de bajo rendimiento al creativo ganador '
                 'y a dos variantes nuevas del mismo concepto (proximidad geográfica).',
         '%s del gasto produjo solo %s de las conversaciones a %s cada una'
         % (pc(0.193), pc(0.062), mx(38.67)), 'CPL y costo por conversación'),
        ('Alta', 'Hacer obligatorios cuatro campos en el CRM al crear el lead: teléfono a 10 dígitos, '
                 'idioma, programa/edad y anuncio de origen.',
         '%s de registros sin teléfono, %s sin idioma, %s sin segmento'
         % (pc(f.dq['leads_sin_tel'] / f.reg_ventana), pc(f.dq['leads_sin_idioma'] / f.reg_ventana),
            pc(f.dq['leads_sin_segmento'] / f.reg_ventana)), 'Trazabilidad y atribución'),
        ('Media', 'Confirmación de asistencia 24 h y 2 h antes de cada clase muestra, con reagenda '
                  'inmediata ofrecida en el mismo mensaje.',
         'No show de %s (%d clases) y ninguna confirmación registrada en las fuentes'
         % (pc(f.tasa_noshow), f.cm_noshow), 'Tasa de asistencia'),
        ('Media', 'Producir dos creativos nuevos de mensaje familiar (hermanos, madre e hijo) para '
                  'atacar el segmento que mejor cierra.',
         'Kids %s y Parenthood %s de cierre frente a %s de College; tres contactos inscribieron a dos '
         'integrantes cada uno'
         % (pc(f.conv_cat['Kids'][2] / f.conv_cat['Kids'][1]),
            pc(f.conv_cat['Parenthood'][2] / f.conv_cat['Parenthood'][1]),
            pc(f.conv_cat['College'][2] / f.conv_cat['College'][1])), 'Inscripciones por clase muestra'),
        ('Media', 'Guion específico de cierre para College: la objeción del adulto no es la misma que la '
                  'de un padre de familia.',
         'College es el segmento de mayor volumen (%d clases) y el de menor cierre (%s)'
         % (f.conv_cat['College'][0], pc(f.conv_cat['College'][2] / f.conv_cat['College'][1])),
         'CM realizada → inscripción'),
        ('Media', 'Registrar formalmente el canal de recomendación con un campo propio y un incentivo '
                  'de referidos.',
         '%d de %d inscripciones del periodo llegaron sin origen de pauta, una marcada explícitamente '
         'como recomendación' % (f.insc_personas - f.insc_atribuibles, f.insc_personas), 'CAC'),
        ('Media', 'Remarketing a quienes iniciaron conversación y no agendaron, usando el creativo '
                  'ganador con mensaje de clase muestra.',
         '%d personas con conversación y sin clase muestra' % (f.leads - f.cm_personas),
         'Lead → CM agendada'),
        ('Media', 'Promover el segundo idioma a la base de alumnos actuales como venta cruzada.',
         'Las %d inscripciones de francés vinieron de alumnos de inglés existentes, sin pauta dedicada'
         % f.insc_idioma.get('Francés', 0), 'Ingreso por alumno'),
        ('Baja', 'Separar el conjunto de anuncios por segmento (familias / adultos) para controlar el '
                 'reparto del gasto en lugar de dejarlo al algoritmo.',
         'Un solo conjunto con presupuesto a nivel campaña concentró %s del gasto en un anuncio'
         % pc(0.700), 'Control de la inversión'),
        ('Baja', 'Exportar cada mes las métricas orgánicas de Facebook, Instagram y TikTok.',
         'No existe ninguna fuente orgánica en el repositorio', 'Medición del aporte orgánico'),
        ('Baja', 'Agregar fecha de inscripción, programa y grupo al padrón de alumnos.',
         'El padrón no permite cortes por fecha ni análisis de cohortes', 'Conciliación comercial'),
        ('Baja', 'Resolver los %d pares de matrículas duplicadas y el registro de prueba del padrón.'
         % len(f.dq['rga_dup']),
         '%d pares con nombre, sexo y edad idénticos; matrícula 1063 con apellido «Prueba»'
         % len(f.dq['rga_dup']), 'Calidad del padrón'),
    ]
    d.table(['Prioridad', 'Acción', 'Problema que resuelve (dato)', 'KPI afectado'],
            [[(p, {'b': 1, 'color': ROJO if p == 'Alta' else (AMBAR if p == 'Media' else '595959')}),
              a, pr, k] for p, a, pr, k in R],
            widths=[900, 3300, 3400, 1760], font=17,
            caption='Tabla 26. Recomendaciones priorizadas')
    d.callout('Si solo se puede hacer una cosa',
              'Atacar el agendamiento. Es el punto donde se pierde el %s de todo el embudo y el único '
              'donde una mejora se convierte casi linealmente en inscripciones, porque las etapas '
              'posteriores ya funcionan: %s de asistencia y %s de cierre. No requiere más presupuesto '
              'de pauta.'
              % (pc((f.leads - f.cm_personas) / (f.leads - f.insc_personas)),
                 pc(f.t_cm_real), pc(f.t_real_insc)))
    d.pagebreak()


def calidad(d, f):
    d.h1('16. Calidad y limitaciones de los datos')
    d.p('Ninguna inconsistencia se corrigió en silencio. Todo lo detectado se documenta aquí y en la '
        'pestaña «Calidad de datos» de la Base Maestra.')
    d.h2('16.1 Campos incompletos')
    N = f.reg_ventana
    d.table(['Campo', 'Registros sin dato', '% de la base', 'Consecuencia'], [
        ['Resultado del seguimiento', n(f.dq['leads_sin_resultado']),
         pc(f.dq['leads_sin_resultado'] / N),
         'El CRM no cierra el ciclo: el resultado comercial se reconstruyó desde la bitácora y el padrón'],
        ['Comentario adicional', n(f.dq['leads_comentario_vacio']), '100.0%',
         'Sin contexto cualitativo ni motivo de pérdida'],
        ['Segmento / programa', n(f.dq['leads_sin_segmento']), pc(f.dq['leads_sin_segmento'] / N),
         'No se puede segmentar el seguimiento por valor'],
        ['Idioma de interés', n(f.dq['leads_sin_idioma']), pc(f.dq['leads_sin_idioma'] / N),
         'No se puede dimensionar la demanda por idioma'],
        ['Teléfono', n(f.dq['leads_sin_tel']), pc(f.dq['leads_sin_tel'] / N),
         'El enlace entre bases depende solo del nombre'],
        ['Anuncio de contacto', n(f.dq['leads_sin_anuncio']), pc(f.dq['leads_sin_anuncio'] / N),
         'Esos leads no se pueden imputar a un creativo'],
        ['Nombre del interesado', n(f.dq['leads_sin_nombre']), pc(f.dq['leads_sin_nombre'] / N),
         'No rastreable en el embudo'],
        ['Etapa del proceso', n(f.dq['leads_sin_etapa']), pc(f.dq['leads_sin_etapa'] / N),
         'Quedan como «Estatus por validar»'],
    ], widths=[2100, 1300, 1100, 4860], aligns=['left', 'right', 'right', 'left'],
        caption='Tabla 27. Campos incompletos en el CRM de campaña (base: %d registros)' % N)

    d.h2('16.2 Duplicados')
    d.bullet('%d registros duplicados detectados y fusionados sobre %d (%s de duplicidad). Sin '
             'deduplicar, los leads se sobreestimarían en esa proporción.'
             % (f.dup_fusionados, f.reg_hoja, pc(f.tasa_duplicidad)))
    d.bullet('%d personas quedaron marcadas para revisión en lugar de fusionarse: nombres de una sola '
             'palabra repetidos con teléfono distinto o ausente.' % len(f.revisar))
    d.bullet('%d personas sin ningún identificador (ni nombre ni teléfono): se conservan pero no son '
             'rastreables.' % len(f.sin_id))
    d.bullet('Los %d registros del 30 de septiembre de la hoja «Octubre» son copia exacta de filas ya '
             'presentes en «AGOSTOSEPTIEMBRE». Se excluyeron del conteo para no duplicar leads.' % f.oct_dup)
    d.bullet('El padrón de alumnos tiene %d pares de matrículas con nombre, sexo y edad idénticos: %s. '
             'Dos de ellos se explican por una segunda inscripción en francés; los otros dos requieren '
             'validación. No se eliminó ningún registro.'
             % (len(f.dq['rga_dup']),
                '; '.join('%s (%s)' % (v[0]['completo'], '/'.join(str(a['matricula']) for a in v))
                          for v in f.dq['rga_dup'].values())))

    d.h2('16.3 Inconsistencias y conflictos entre fuentes')
    d.table(['Inconsistencia', 'Detalle', 'Tratamiento aplicado'], [
        ['Año erróneo en la bitácora de clases muestra',
         '%d registros de la hoja SEPTIEMBRE traen año 2006 (filas %s)'
         % (f.dq['cm_anio_erroneo'], ', '.join(str(k[1]) for k in f.anio_typos)),
         'Interpretados como 2026. Es una INFERENCIA: el resto de la hoja y el orden de los registros '
         'son consistentes con septiembre de 2026'],
        ['Inscripción en el CRM no reflejada en la bitácora',
         '%s: el CRM la marca «Inscrito» y aparece en el padrón, la bitácora dice «Show»'
         % ', '.join(r['nombre'] for r in f.insc_conflicto),
         'Contabilizada como inscripción por tener doble evidencia (CRM y padrón)'],
        ['Alumnos del padrón sin cierre marcado',
         '%s: matrícula reciente y clase muestra en el periodo, pero la bitácora no los marca inscritos'
         % '; '.join('%s (mat. %s)' % (r['nombre'], r['matricula']) for r in f.insc_por_validar),
         'NO contabilizados. Reportados para validación administrativa'],
        ['Inscripciones de la bitácora sin correspondencia en el padrón',
         '%s' % ', '.join(r['nombre'] for r in f.insc_sin_padron),
         'Se conservan como inscripciones (la bitácora es la fuente del cierre) y se señala la '
         'discrepancia: puede ser alta no procesada, baja posterior o error de captura'],
        ['Diferencias de edad entre bitácora y padrón',
         '%d casos con 3 o más años de diferencia: %s'
         % (len(f.edad_dif), '; '.join('%s (%s vs %s)' % (a, b, d) for a, b, c, d, e in f.edad_dif)),
         'No invalidan el enlace (nombre completo coincide); se documenta la captura imprecisa'],
        ['Valor no válido en el CRM', 'Un registro tiene «1» en «Resultado del seguimiento»',
         'Tratado como sin información'],
        ['Subregistro del campo «Registrado a clase muestra»',
         'El CRM marca %d personas registradas a clase muestra frente a %d clases en la bitácora'
         % (f.cm_solicitadas_crm, f.cm_agendadas),
         'El campo no se usa como métrica; la bitácora es la fuente única de clases muestra'],
        ['Clases muestra sin lead identificado',
         '%d de %d clases (%s) no se pueden ligar a un lead del CRM'
         % (f.cm_sin_lead, f.cm_agendadas, pc(f.cm_sin_lead / f.cm_agendadas)),
         'Incluye recomendaciones y visitas a sitio, pero también fallas de registro. No se atribuyen '
         'a la pauta'],
        ['Brecha entre Meta y el CRM',
         '%d conversaciones en Meta frente a %d registros en el CRM' % (f.conversaciones, f.reg_ventana),
         'Se reporta como indicador de cobertura de registro (%s), no como pérdida confirmada'
         % pc(f.cobertura_crm)],
        ['Registro de prueba en el padrón', 'Matrícula 1063 «Gustavo Martinez Prueba»',
         'Se conserva porque existe en la fuente; se marca para validación'],
    ], widths=[2200, 3800, 3360], font=17,
        caption='Tabla 28. Inconsistencias detectadas y tratamiento documentado')

    d.h2('16.4 Información que no existe y que impide calcular indicadores')
    for t in ['Métricas orgánicas de Facebook, Instagram y TikTok: la sección de redes sociales '
              'orgánicas no es calculable.',
              'Clics de los anuncios: CTR y CPC no son calculables.',
              'Desglose diario del gasto: no se puede cruzar inversión con captación día por día.',
              'Alcance único a nivel campaña: solo se puede reportar un rango.',
              'Hora del lead y hora del primer contacto de vuelta: el tiempo de respuesta no es medible.',
              'Campo de clase muestra «solicitada» distinto de «agendada»: esa etapa del embudo no es '
              'reportable.',
              'Motivo de pérdida: no se puede analizar por qué no cerraron los prospectos.',
              'Fecha de inscripción, programa, grupo y modalidad en el padrón: impide filtrar el padrón '
              'al 3 de octubre y segmentar inscripciones desde esa fuente.']:
        d.bullet(t)

    d.h2('16.5 Qué debería empezar a registrarse de forma sistemática')
    d.table(['Campo a registrar', 'Para qué sirve'], [
        ['Fecha y hora exactas del lead', 'Medir tiempo de respuesta'],
        ['Fecha y hora del primer contacto de vuelta', 'Correlacionar velocidad con agendamiento'],
        ['ID de campaña, conjunto y anuncio (no texto libre)', 'Atribuir sin mapeos inferidos'],
        ['Teléfono normalizado obligatorio', 'Único identificador fiable entre bases'],
        ['Idioma, programa y edad desde el primer contacto', 'Segmentar el seguimiento por valor'],
        ['Fecha de solicitud de clase muestra, separada de la fecha de agenda',
         'Completar la etapa «solicitadas» del embudo'],
        ['Asistencia a clase muestra como campo obligatorio', 'Medir no show directamente'],
        ['Bitácora de seguimientos con fecha por interacción', 'Saber cuántos toques recibió cada prospecto'],
        ['Estatus con lista cerrada de valores', 'Evitar campos vacíos y valores no válidos'],
        ['Fecha de inscripción y motivo de pérdida', 'Cerrar el ciclo y calcular CAC por cohorte'],
        ['Export mensual de métricas orgánicas', 'Hacer analizable el aporte del contenido'],
        ['Export de anuncios con clics y desglose diario', 'Calcular CTR, CPC y curvas de entrega'],
    ], widths=[4000, 5360], caption='Tabla 29. Registro mínimo recomendado para la próxima campaña')
    d.pagebreak()


def conclusion(d, f):
    d.h1('17. Conclusión ejecutiva')
    d.h3('¿Cómo funcionó la Campaña de Septiembre?')
    d.p('Funcionó bien como motor de demanda y de forma incompleta como motor de ventas. Con %s de '
        'inversión ejecutada al %s generó %d conversaciones, %d prospectos identificables, %d clases '
        'muestra y %d inscripciones confirmadas, de las cuales %d son atribuibles a la pauta.'
        % (mx(f.gasto), pc(f.uso_presupuesto), f.conversaciones, f.leads, f.cm_agendadas,
           f.insc_personas, f.insc_atribuibles))
    d.h3('¿Qué obtuvimos con la inversión realizada?')
    d.p('Un costo de %s por persona interesada, %s por clase muestra realizada y un CAC de %s. '
        'Para una academia con ingreso recurrente mensual, ese CAC se recupera en el corto plazo; el '
        'problema no es el precio de la adquisición sino el volumen de adquisiciones que se dejó sobre '
        'la mesa.' % (mx(f.cpl_persona), mx(f.costo_cm_realizada), mx(f.cac)))
    d.h3('¿Cuál fue la principal fortaleza?')
    d.p('La clase muestra. %s de asistencia y %s de cierre significan que el proceso presencial está '
        'resuelto. A eso se suma un creativo ganador claramente identificado, con %s por conversación.'
        % (pc(f.t_cm_real), pc(f.t_real_insc), mx(9.98)))
    d.h3('¿Cuál fue el principal cuello de botella?')
    d.p('El agendamiento y su gemelo, el seguimiento. %d de %d personas (%s) nunca llegaron a una clase '
        'muestra, y el CRM no registra resultado de seguimiento en el %s de sus registros. Ahí se '
        'concentra el %s de toda la pérdida del embudo.'
        % (f.leads - f.cm_personas, f.leads, pc((f.leads - f.cm_personas) / f.leads),
           pc(f.dq['leads_sin_resultado'] / f.reg_ventana),
           pc((f.leads - f.cm_personas) / (f.leads - f.insc_personas))))
    d.h3('¿Qué deberíamos repetir?')
    d.p('El creativo de proximidad geográfica en formato video, el objetivo de mensajes con WhatsApp como '
        'canal operativo, la ejecución total del presupuesto y el formato de la clase muestra tal como '
        'está diseñado hoy.')
    d.h3('¿Qué deberíamos cambiar?')
    d.p('Tres cosas concretas: proponer horarios de clase muestra en el primer mensaje en lugar de enviar '
        'información; establecer una cadencia de seguimiento con registro obligatorio; y dejar de '
        'repartir presupuesto entre anuncios que ya demostraron ser 3.9 veces más caros.')
    d.h3('¿Qué oportunidades quedan abiertas?')
    d.p('%d oportunidades comerciales activas al 3 de octubre, de las cuales %d están en descubrimiento '
        'sin movimiento reciente; %d clases muestra pendientes al cierre del periodo; venta cruzada de '
        'un segundo idioma a la base actual; y un canal de recomendación que produce inscripciones sin '
        'estar medido.' % (f.activos, f.dormidos, f.cm_pendientes))
    d.h3('¿Qué implicaciones tiene esto para la siguiente campaña?')
    d.p('Con el mismo presupuesto de %s y sin mejorar un solo indicador de pauta, llevar el agendamiento '
        'de %s a 25%% produciría del orden de 10 inscripciones adicionales, un %s más que el resultado '
        'de esta campaña. La prioridad de la siguiente campaña no es comprar más demanda: es convertir '
        'la que ya se está comprando.'
        % (mx(f.presupuesto, 0), pc(f.t_lead_cm), pc(10.5 / f.insc_personas)))
    d.pagebreak()


def anexo(d, f):
    d.h1('Anexo. Tablas de datos de las gráficas')
    d.p('Cada gráfica de este reporte está ligada a su propia tabla de datos editable (clic derecho › '
        'Editar datos). Estas tablas reproducen esos datos para permitir reconstruir cualquier gráfica, '
        'y están además en la pestaña «Gráficas» de la Base Maestra.')
    d.table(['Funnel comercial', 'Valor'],
            [[nm, n(v)] for nm, v, fu in f.etapas], widths=[5000, 4360],
            aligns=['left', 'right'], caption='A1. Funnel completo')
    d.table(['Conversión por etapa', 'Valor'],
            [['Lead → CM agendada', pc(f.t_lead_cm)], ['CM agendada → realizada', pc(f.t_cm_real)],
             ['CM realizada → inscripción', pc(f.t_real_insc)], ['Lead → inscripción', pc(f.t_lead_insc)]],
            widths=[5000, 4360], aligns=['left', 'right'], caption='A2. Conversión por etapa')
    d.table(['Resultado de clase muestra', 'Clases'],
            [['Realizadas con inscripción', n(f.cm_inscritos_reg)],
             ['Realizadas sin cierre', n(f.cm_realizadas - f.cm_inscritos_reg)],
             ['No show', n(f.cm_noshow)], ['Pendientes al 4-oct', n(f.cm_pendientes)],
             ['Reprogramadas', n(f.cm_reprogramadas)]],
            widths=[5000, 4360], aligns=['left', 'right'], caption='A3. Clases muestra por resultado')
    d.table(['Anuncio', 'Gasto', 'Conversaciones', 'Costo/conv.', 'Conv. por mil impresiones'],
            [[t['anuncio'], mx(t['gasto']), n(t['conv']), mx(t['costo_conv']),
              '%.2f' % (t['conv'] / t['impresiones'] * 1000)]
             for t in sorted(f.tab_ads, key=lambda r: -r['conv'])],
            widths=[2200, 1700, 1800, 1800, 1860],
            aligns=['left', 'right', 'right', 'right', 'right'],
            caption='A4, A5 y A8. Resultados, inversión y eficacia por anuncio')
    d.table(['Anuncio de contacto (CRM)', 'Leads', 'Clases muestra', 'Inscripciones atribuibles'],
            [[k, n(f.crm_pers_ad[k]), n(f.cm_ad.get(k, 0)), n(f.insc_ad.get(k, 0))]
             for k in sorted(f.crm_pers_ad, key=lambda k: -f.crm_pers_ad[k])],
            widths=[3600, 1800, 1900, 2060], aligns=['left', 'right', 'right', 'right'],
            caption='A6. Del lead a la inscripción por anuncio de contacto')
    d.table(['Programa', 'Inscripciones'],
            [[k, n(v)] for k, v in f.insc_cat.most_common()], widths=[5000, 4360],
            aligns=['left', 'right'], caption='A7. Distribución de inscripciones por programa')
    d.table(['Semana', 'Leads', 'CM agendadas', 'CM realizadas', 'Inscripciones'],
            [[lab, n(l), n(ca), n(cr), n(ins)] for lab, l, ca, cr, ins in f.semanas],
            widths=[2000, 1800, 1900, 1900, 1760],
            aligns=['left', 'right', 'right', 'right', 'right'],
            caption='A9. Evolución semanal')
    d.table(['Estatus homologado', 'Prospectos'],
            [[k, n(f.estatus_count.get(k, 0))] for k in facts.ESTATUS_ORDEN],
            widths=[5000, 4360], aligns=['left', 'right'],
            caption='A10. Prospectos por estatus al 3 de octubre')
    d.spacer(1)
    d.small('Fin del reporte. Entregables complementarios: Base Maestra Campaña Septiembre 2026.xlsx · '
            'Presentación Ejecutiva — Campaña Septiembre 2026.pptx · Base de Alumnos Inscritos — Corte 03 '
            'octubre 2026.xlsx · Prospectos en Curso — Corte 03 octubre 2026.xlsx')


def build():
    f = facts.get()
    d = Doc('Reporte Campaña Septiembre 2026 — Kugelman Academy',
            subject='Marketing, captación y resultados comerciales · 21 ago – 4 oct 2026')
    portada(d, f)
    resumen(d, f)
    contexto(d, f)
    metodologia(d, f)
    funnel(d, f)
    clases(d, f)
    pauta(d, f)
    redes(d, f)
    inscripciones(d, f)
    prospectos(d, f)
    temporal(d, f)
    hallazgos(d, f)
    funciono(d, f)
    oportunidades(d, f)
    recomendaciones(d, f)
    calidad(d, f)
    conclusion(d, f)
    anexo(d, f)
    path = os.path.join(OUT, 'Reporte Campaña Septiembre 2026 — Kugelman Academy.docx')
    d.save(path)
    return path


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    print(build())
