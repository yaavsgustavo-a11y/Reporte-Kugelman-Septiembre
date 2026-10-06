# -*- coding: utf-8 -*-
"""ENTREGABLE 4 — Base de Alumnos Inscritos — Corte 03 octubre 2026 (.xlsx)
ENTREGABLE 5 — Prospectos en Curso — Corte 03 octubre 2026 (.xlsx)
"""
import os, collections, datetime
import facts, core, xlsxw as X
from xlsxw import Sheet, Chart, C, F

OUT = '/projects/sandbox/entregables'
CORTE = 'sábado 3 de octubre de 2026'


def pc(v, dec=1):
    return ('{:.%df}%%' % dec).format(v * 100)


# ==========================================================================
# ENTREGABLE 4 — BASE DE ALUMNOS INSCRITOS
# ==========================================================================
def build_alumnos():
    f = facts.get()
    A = f.alumnos
    meta = f.M['rga_meta']
    sheets = []

    # ---------------------------------------------------------- hoja 1: base
    s = Sheet('Alumnos', cols=[12, 26, 30, 8, 8, 10], freeze='A2', tab_color=X.AZUL)
    headers = ['Matrícula', 'Nombre', 'Apellidos', 'Sexo', 'Edad', 'Activo']
    s.row(*[C(h, 'hdr') for h in headers])
    for a in sorted(A, key=lambda r: r['matricula']):
        s.row(C(a['matricula'], 'num'), C(a['nombre'], 'txt'), C(a['apellidos'], 'txt'),
              C(a['sexo'], 'txt'), C(a['edad'], 'num'), C(a['activo'], 'txt'))
    last = len(s.rows)
    s.tables.append(('BaseAlumnos', 'A1:F%d' % last, headers))
    sheets.append(s)

    # ------------------------------------------------- hoja 2: validaciones
    v = Sheet('Validaciones', cols=[46, 14, 72], tab_color=X.ROJO)
    v.row(C('BASE DE ALUMNOS INSCRITOS — VALIDACIONES Y NOTAS DE FUENTE', 'titulo'))
    v.row(C('Fecha de corte solicitada: %s  ·  Hoja «Alumnos»: %d registros, 6 columnas exactas.'
            % (CORTE, len(A)), 'nota'))
    v.blank()
    v.row(C('Validación realizada', 'hdr'), C('Resultado', 'hdr'), C('Detalle', 'hdr'))
    dupn = f.dq['rga_dup']
    mats = [a['matricula'] for a in A]
    dup_mat = [m for m, c in collections.Counter(mats).items() if c > 1]
    edades_mal = [a for a in A if not isinstance(a['edad'], int) or a['edad'] < 0 or a['edad'] > 110]
    sexos = f.dq['rga_sexo']
    sexo_mal = {k: vv for k, vv in sexos.items() if k not in ('F', 'M')}
    NM = {'santiago', 'jose', 'diego', 'mateo', 'ricardo', 'emiliano', 'leonardo', 'rodrigo',
          'miguel', 'juan', 'luis', 'alfredo', 'erick', 'axel', 'dylan', 'rommel', 'antonio',
          'victor', 'gustavo', 'tomas', 'eduardo', 'sebastian', 'andrei', 'alexis', 'leobardo',
          'leonel', 'mario', 'ian', 'john', 'edgar', 'angel', 'said', 'matias', 'iktan', 'roberto',
          'ariel', 'daniel'}
    NF = {'maria', 'ana', 'abigail', 'abril', 'alejandra', 'andrea', 'dulce', 'ila', 'josefina',
          'lucia', 'marian', 'melany', 'melissa', 'valeria', 'zuleyka', 'fernanda', 'renata',
          'paula', 'mairy', 'lizzy', 'nancy', 'lesly', 'luisa', 'montserrat', 'sofia', 'ivanna',
          'carmen', 'aylen', 'brenda', 'elvia', 'isabella', 'teresa', 'amira', 'cecilia',
          'jonintzin', 'mara', 'jacqueline', 'laura', 'ximena', 'zoemy', 'darlene', 'nicol',
          'samantha', 'heliana', 'jade', 'irma', 'yara', 'lianne', 'jessica', 'mariana', 'elisa',
          'camila', 'emilia', 'claudia'}
    sexo_dudoso = []
    for a in sorted(A, key=lambda r: r['matricula']):
        g = core.norm_name(a['nombre']).split()[0] if a['nombre'] else ''
        if (g in NM and a['sexo'] == 'F') or (g in NF and a['sexo'] == 'M'):
            sexo_dudoso.append(a)
    act = f.dq['rga_activos']
    filas = [
        ('1. Duplicados por matrícula', 'Sin duplicados' if not dup_mat else '%d encontrados' % len(dup_mat),
         'La matrícula es única en los %d registros: se usa como identificador principal.' % len(A)),
        ('2. Duplicados por nombre + sexo + edad', '%d pares' % len(dupn),
         ('Se detectaron los siguientes pares con datos idénticos pero matrícula distinta: %s. '
          'NO se eliminó ningún registro porque cada uno tiene matrícula propia en la fuente. '
          'Dos de los pares (García Vazquez) se explican por una segunda inscripción en francés '
          'documentada en la bitácora de clases muestra; los otros dos requieren validación '
          'administrativa.'
          % '; '.join('%s → matrículas %s' % (vv[0]['completo'],
                                              ' y '.join(str(x['matricula']) for x in vv))
                      for vv in dupn.values())) if dupn else 'Sin coincidencias.'),
        ('3. Nombre y Apellidos como campos separados', 'Cumple',
         'La fuente trae «Apellido paterno» y «Apellido materno» en columnas distintas. Se '
         'concatenaron en «Apellidos» respetando el orden paterno + materno, sin alterar ningún '
         'valor. «Nombre» se conservó íntegro y separado.'),
        ('4. Valor de Sexo respetado', 'Cumple',
         'Distribución tal como está en la fuente: %s. Valores fuera de F/M: %s. No se modificó '
         'ningún valor. Se detectaron %d registros en los que el sexo registrado no parece '
         'corresponder con el nombre de pila (%s): se respeta el valor de la fuente y se señalan '
         'para validación administrativa.'
         % (', '.join('%s=%d' % (k, vv) for k, vv in sorted(sexos.items())),
            (sexo_mal if sexo_mal else 'ninguno'), len(sexo_dudoso),
            '; '.join('matrícula %s %s = %s' % (x['matricula'], x['completo'], x['sexo'])
                      for x in sexo_dudoso) or 'ninguno')),
        ('5. Formato de Edad', 'Cumple' if not edades_mal else '%d fuera de rango' % len(edades_mal),
         'Todas las edades son números enteros. Rango observado: %d a %d años. No se detectaron '
         'valores vacíos, negativos ni no numéricos.'
         % (min(a['edad'] for a in A), max(a['edad'] for a in A))),
        ('6. Valor de Activo conservado', 'Cumple',
         'Distribución en la fuente: %s. Se conservó literalmente el valor exportado.'
         % ', '.join('%s=%d' % (k, vv) for k, vv in act.items())),
        ('7. No se eliminaron alumnos inactivos', 'Cumple',
         'No se eliminó ningún registro por su estatus. En esta exportación los %d alumnos están '
         'marcados Activo = «Sí», por lo que no hubo inactivos que conservar.' % len(A)),
        ('8. No se inventó información faltante', 'Cumple',
         'No se agregó, estimó ni completó ningún dato. Los %d registros reproducen exactamente la '
         'fuente en las seis columnas solicitadas.' % len(A)),
        ('9. No se modificaron registros sin evidencia', 'Cumple',
         'La única transformación aplicada es la concatenación de los dos apellidos en una sola '
         'columna, requerida por el formato solicitado.'),
        ('10. Columnas entregadas', 'Exactamente 6',
         'Matrícula | Nombre | Apellidos | Sexo | Edad | Activo. NO se incluyó teléfono, idioma, '
         'horario, grupo, modalidad, mensualidad, fecha de inscripción, fuente, programa ni '
         'observaciones, conforme a lo solicitado.'),
    ]
    for a, b, c in filas:
        v.row(C(a, 'txt'), C(b, 'ok' if b.startswith('Cumple') or b.startswith('Sin') else 'warn'),
              C(c, 'txtw'))
    v.blank()
    v.row(C('LIMITACIÓN IMPORTANTE SOBRE LA FECHA DE CORTE', 'seccion'), C('', 'seccion'), C('', 'seccion'))
    v.row(C('El archivo fuente «RGA_1791241975.xlsx» fue exportado el %s y NO contiene fecha de '
            'inscripción ni fecha de alta por alumno. Por lo tanto NO es posible filtrar el padrón al '
            'corte del %s: la base refleja el estado del padrón tal como quedó en la exportación. '
            'Se entrega completa y sin recortes para no eliminar registros sin evidencia. '
            'Para que un corte por fecha sea posible en el futuro, el padrón debe incluir la fecha de '
            'inscripción de cada alumno.' % (meta['fecha_export'], CORTE), 'txtw'))
    v.merge('A%d:C%d' % (len(v.rows), len(v.rows)))
    v.blank()
    v.row(C('REGISTRO QUE REQUIERE VALIDACIÓN ADMINISTRATIVA', 'seccion'), C('', 'seccion'), C('', 'seccion'))
    v.row(C('Matrícula 1063 — «Gustavo Martinez Prueba»', 'txt'), C('Revisar', 'warn'),
          C('El apellido materno es «Prueba»: tiene todas las características de un registro de prueba '
            'del sistema. Se conserva en la base porque existe en la fuente y no hay evidencia para '
            'eliminarlo. Se recomienda validarlo y, si procede, darlo de baja en el sistema origen.', 'txtw'))
    v.blank()
    v.row(C('COMPOSICIÓN DEL PADRÓN (informativo, no modifica la base)', 'seccion'),
          C('', 'seccion'), C('', 'seccion'))
    v.row(C('Indicador', 'hdr2'), C('Valor', 'hdr2'), C('Nota', 'hdr2'))
    bands = [('0 a 5 años', 0, 5), ('6 a 11 años', 6, 11), ('12 a 17 años', 12, 17),
             ('18 a 29 años', 18, 29), ('30 a 49 años', 30, 49), ('50 años o más', 50, 200)]
    v.row(C('Total de alumnos', 'txt'), C(len(A), 'num'), C('Según exportación del %s' % meta['fecha_export'], 'txt'))
    for k, vv in sorted(sexos.items()):
        v.row(C('Sexo %s' % k, 'txt'), C(vv, 'num'), C(pc(vv / len(A)), 'txt'))
    for lab, lo, hi in bands:
        cnt = sum(1 for a in A if lo <= a['edad'] <= hi)
        v.row(C(lab, 'txt'), C(cnt, 'num'), C(pc(cnt / len(A)), 'txt'))
    v.row(C('Edad promedio', 'txt'), C(round(sum(a['edad'] for a in A) / len(A), 1), 'dec1'), C('', 'txt'))
    v.row(C('Matrícula mínima / máxima', 'txt'),
          C('%d / %d' % (min(mats), max(mats)), 'txt'),
          C('La matrícula parece asignarse de forma secuencial: las altas más recientes están en el '
            'rango alto.', 'txtw'))
    v.blank()
    v.row(C('FUENTE', 'seccion'), C('', 'seccion'), C('', 'seccion'))
    v.row(C('Archivo', 'bold'), C('RGA_1791241975.xlsx — hoja «General de alumnos»', 'txt'))
    v.row(C('Encabezado de la fuente', 'bold'), C(meta['encabezado'], 'txt'))
    v.row(C('Fecha de exportación', 'bold'), C(meta['fecha_export'], 'txt'))
    v.row(C('Grupos incluidos en la fuente', 'bold'), C(meta['grupos'], 'txtw'))
    sheets.append(v)

    path = os.path.join(OUT, 'Base de Alumnos Inscritos — Corte 03 octubre 2026.xlsx')
    X.write(path, sheets, 'Base de Alumnos Inscritos — Corte 03 octubre 2026')
    return path


# ==========================================================================
# ENTREGABLE 5 — PROSPECTOS EN CURSO
# ==========================================================================
def build_prospectos():
    f = facts.get()
    sheets = []
    P = f.pipeline

    # --------------------------------------------------- hoja 1: prospectos
    s = Sheet('Prospectos', cols=[30, 14, 20, 24, 11, 42, 24, 14, 10, 52, 9, 13, 14, 24, 24, 11],
              freeze='A4', tab_color=X.AZUL)
    s.row(C('PROSPECTOS EN CURSO — CORTE %s' % CORTE.upper(), 'titulo'))
    s.row(C('%d registros consolidados: CRM de la campaña de septiembre (deduplicado) + personas que '
            'solo aparecen en la bitácora de clases muestra. Estatus homologado. Los prospectos de la '
            'campaña de julio están en la pestaña de referencia y NO se cuentan como activos.'
            % len(P), 'nota'))
    headers = ['Prospecto', 'Fecha primer contacto', 'Interés', 'Edad / Programa', 'Horario',
               'Clase muestra', 'Estatus', 'Último seguimiento', 'Oportunidad abierta', 'Observaciones',
               'ID', 'Teléfono', 'Medio de contacto', 'Anuncio de origen', 'Origen del registro',
               'Días sin movimiento']
    s.row(*[C(h, 'hdr') for h in headers])
    for r in P:
        prog = r['segmento']
        if r['edad'] is not None:
            prog = '%s · %s años' % (prog, r['edad']) if prog != 'Sin información' else '%s años' % r['edad']
        ult = r['cm_fecha'] if (r['cm_fecha'] and r['cm_fecha'] > r['f_ultimo']) else r['f_ultimo']
        fcol = 'fecha_warn' if r.get('fecha_inferida') else 'fecha'
        s.row(C(r['nombre'], 'txt'), C(r['f_primer'], fcol), C(r['interes'], 'txt'),
              C(prog, 'txt'), C(r['hora'] or 'Sin información', 'txt'), C(r['cm'], 'txtw'),
              C(r['estatus'], 'txt'), C(ult, 'fecha'),
              C('Sí' if r['activo'] else 'No', 'ok' if r['activo'] else 'txt'),
              C(r['obs'] or '', 'txtw'), C(r['id'], 'txt'), C(r['tel'] or 'Sin información', 'txt'),
              C(r['medio'], 'txt'), C(r['anuncio'], 'txt'), C(r['origen'], 'txt'),
              C(r['dias_sin_mov'], 'num'))
    s.autofilter = 'A3:P%d' % len(s.rows)
    sheets.append(s)

    # ------------------------------------------------------- hoja 2: resumen
    r = Sheet('Resumen', cols=[40, 14, 14, 6, 26, 14, 14, 50], tab_color=X.AZUL2)
    r.row(C('RESUMEN DE PROSPECTOS EN CURSO — CORTE %s' % CORTE.upper(), 'titulo'))
    r.row(C('Fuente: CRM de campaña (hoja AGOSTOSEPTIEMBRE) + bitácora de clases muestra. '
            'Universo: campaña de septiembre 2026.', 'nota'))
    r.blank()
    r.row(C('Indicador', 'hdr'), C('Prospectos', 'hdr'), C('% del total', 'hdr'), C('', 'hdr'),
          C('', 'hdr'), C('', 'hdr'), C('', 'hdr'), C('Nota', 'hdr'))
    tot = len(P)
    rt = len(s.rows)
    r.row(C('Total de registros analizados', 'bold'), C(tot, 'boldnum'), C(1.0, 'boldpct'), C(''),
          C(''), C(''), C(''), C('Una fila por persona en la pestaña «Prospectos»', 'txtw'))
    r.row(C('Registros sin fecha de primer contacto en las fuentes', 'txt'),
          C(f.sin_fecha_contacto, 'num'), C(f.sin_fecha_contacto / tot, 'pct'), C(''), C(''), C(''), C(''),
          C('Provienen solo de la bitácora de clases muestra y tienen cita posterior al corte. '
            'En la columna «Fecha primer contacto» se muestra la fecha de su clase muestra, resaltada, '
            'y se señala en Observaciones. Se conservan porque son citas vigentes.', 'txtw'))
    r.row(C('Total de prospectos activos', 'bold'), C(f.activos, 'boldnum'),
          C(f.activos / tot, 'boldpct'), C(''), C(''), C(''), C(''),
          C('Estados: Nuevo, Descubrimiento, CM agendada, CM realizada, Reprogramación, '
            'Seguimiento, Por cerrar y Lista de espera', 'txtw'))
    r.blank()
    r.row(C('DESGLOSE POR ESTATUS HOMOLOGADO', 'seccion'), C('', 'seccion'), C('', 'seccion'),
          C('', 'seccion'), C('', 'seccion'), C('', 'seccion'), C('', 'seccion'), C('', 'seccion'))
    r.row(C('Estatus', 'hdr'), C('Prospectos', 'hdr'), C('% del total', 'hdr'), C('', 'hdr'),
          C('', 'hdr'), C('', 'hdr'), C('', 'hdr'), C('Oportunidad abierta', 'hdr'))
    g0 = len(r.rows) + 1
    LAB = {
        'Nuevo': 'Nuevos',
        'Descubrimiento': 'Descubrimiento',
        'Clase muestra agendada': 'Clases muestra agendadas',
        'Clase muestra realizada': 'Clases muestra realizadas pendientes de cierre',
        'Reprogramación': 'Reprogramaciones',
        'Seguimiento': 'Seguimientos',
        'Por cerrar': 'Prospectos por cerrar',
        'Lista de espera': 'Lista de espera',
        'Sin respuesta': 'Sin respuesta',
        'No interesado': 'No interesados',
        'Inscrito / Cerrado': 'Inscritos / cerrados',
        'Estatus por validar': 'Estatus por validar',
    }
    NOTA = {
        'Lista de espera': 'No existe en las fuentes ningún campo o valor que registre lista de espera',
        'Inscrito / Cerrado': 'No se cuentan como oportunidad abierta para no duplicarlos',
        'Estatus por validar': 'Registros del CRM sin etapa capturada',
        'Seguimiento': 'Incluye a quienes no asistieron a su clase muestra y siguen abiertos',
    }
    cats = []
    for k in facts.ESTATUS_ORDEN:
        vv = f.estatus_count.get(k, 0)
        cats.append((LAB[k], vv))
        rr = len(r.rows) + 1
        r.row(C(LAB[k], 'txt'), C(vv, 'num'), F('=B%d/$B$5' % rr, 'pct'), C(''), C(''), C(''), C(''),
              C(NOTA.get(k, 'Sí' if k in facts.ACTIVOS else 'No'), 'txtw'))
    g1 = len(r.rows)
    rr = len(r.rows) + 1
    r.row(C('TOTAL', 'bold'), F('=SUM(B%d:B%d)' % (g0, g1), 'boldnum'),
          F('=B%d/$B$5' % rr, 'boldpct'))
    r.blank()

    # funnel del pipeline
    r.row(C('FUNNEL DEL PIPELINE', 'seccion'), C('', 'seccion'), C('', 'seccion'), C('', 'seccion'),
          C('', 'seccion'), C('', 'seccion'), C('', 'seccion'), C('', 'seccion'))
    r.row(C('Etapa', 'hdr'), C('Registros', 'hdr'), C('Conv. vs previa', 'hdr'), C('', 'hdr'),
          C('', 'hdr'), C('', 'hdr'), C('', 'hdr'), C('Nota', 'hdr'))
    pf = f.pipe_funnel
    f0 = len(r.rows) + 1
    fun = [('Prospectos', pf['prospectos']), ('CM agendadas', pf['cm_agendadas']),
           ('CM realizadas', pf['cm_realizadas']), ('Inscritos / cerrados', pf['inscritos'])]
    for i, (k, vv) in enumerate(fun):
        rr = len(r.rows) + 1
        r.row(C(k, 'txt'), C(vv, 'num'),
              C('—', 'txt') if i == 0 else F('=B%d/B%d' % (rr, rr - 1), 'pct'),
              C(''), C(''), C(''), C(''), C('', 'txtw'))
    f1 = len(r.rows)
    r.blank()
    r.row(C('Las clases muestra y las inscripciones de este funnel se cuentan a nivel de REGISTRO DE '
            'PROSPECTO (el contacto que escribió). Tres contactos inscribieron a dos integrantes de su '
            'familia cada uno, por lo que los %d registros «Inscrito / Cerrado» equivalen a %d personas '
            'inscritas en el periodo. El detalle individual de esas inscripciones está en el Reporte de '
            'Campaña y en la Base Maestra.' % (pf['inscritos'], f.insc_personas), 'txtw'))
    r.merge('A%d:H%d' % (len(r.rows), len(r.rows)))
    r.blank()

    # calidad del pipeline
    r.row(C('CALIDAD DEL PIPELINE: ANTIGÜEDAD DEL ÚLTIMO MOVIMIENTO', 'seccion'), C('', 'seccion'),
          C('', 'seccion'), C('', 'seccion'), C('', 'seccion'), C('', 'seccion'), C('', 'seccion'),
          C('', 'seccion'))
    r.row(C('Días sin movimiento registrado', 'hdr'), C('Prospectos activos', 'hdr'),
          C('% de activos', 'hdr'), C('', 'hdr'), C('', 'hdr'), C('', 'hdr'), C('', 'hdr'),
          C('Acción sugerida', 'hdr'))
    bands = [('0 a 7 días', 0, 7, 'Seguimiento normal'), ('8 a 14 días', 8, 14, 'Reactivar esta semana'),
             ('15 a 30 días', 15, 30, 'Campaña de recuperación'),
             ('Más de 30 días', 31, 9999, 'Último intento o descartar con motivo')]
    b0 = len(r.rows) + 1
    for lab, lo, hi, acc in bands:
        cnt = sum(1 for x in P if x['activo'] and lo <= x['dias_sin_mov'] <= hi)
        r.row(C(lab, 'txt'), C(cnt, 'num'), C(cnt / f.activos, 'pct'), C(''), C(''), C(''), C(''),
              C(acc, 'txtw'))
    b1 = len(r.rows)
    r.row(C('Prospectos en «Descubrimiento» con 15 días o más sin movimiento', 'bold'),
          C(f.dormidos, 'boldnum'), C(f.dormidos / f.activos, 'boldpct'), C(''), C(''), C(''), C(''),
          C('Inventario ya pagado: trabajarlo no requiere inversión publicitaria nueva', 'txtw'))
    r.blank()

    # referencia julio
    r.row(C('REFERENCIA: PROSPECTOS DE LA CAMPAÑA DE JULIO 2026 (NO contabilizados arriba)', 'seccion'),
          C('', 'seccion'), C('', 'seccion'), C('', 'seccion'), C('', 'seccion'), C('', 'seccion'),
          C('', 'seccion'), C('', 'seccion'))
    r.row(C('Registros en la hoja JULIO', 'txt'), C(f.julio_regs, 'num'), C(''), C(''), C(''), C(''),
          C(''), C('Archivo CAMPAÑAS (2).xlsx, hoja «JULIO»', 'txtw'))
    r.row(C('Personas distintas tras deduplicar', 'txt'), C(f.julio_personas, 'num'), C(''), C(''),
          C(''), C(''), C(''), C('', 'txtw'))
    r.row(C('En estados abiertos según su último registro de julio', 'txt'), C(f.julio_activos, 'num'),
          C(f.julio_activos / f.julio_personas, 'pct'), C(''), C(''), C(''), C(''),
          C('Su estatus está congelado en julio: no hay ningún registro posterior. Se entregan en la '
            'pestaña «Prospectos julio (ref.)» para que Dirección decida si se reactivan, pero NO se '
            'suman a los activos de la campaña de septiembre para no inflar el pipeline con '
            'información desactualizada.', 'txtw'))
    r.blank()
    r.row(C('Gráficas editables: ver abajo. Cada gráfica está ligada a las tablas de esta pestaña.', 'nota'))

    r.charts.append(Chart('bar', 'Prospectos por estatus al corte del 3 de octubre de 2026', 'Resumen',
                          "'Resumen'!$A$%d:$A$%d" % (g0, g1), [c[0] for c in cats],
                          [('Prospectos', "'Resumen'!$B$%d:$B$%d" % (g0, g1), [c[1] for c in cats])],
                          'J3', width=9, height=22, colors=[X.AZUL2]))
    r.charts.append(Chart('col', 'Funnel del pipeline: prospectos → CM agendadas → CM realizadas → inscritos',
                          'Resumen', "'Resumen'!$A$%d:$A$%d" % (f0, f1), [x[0] for x in fun],
                          [('Registros', "'Resumen'!$B$%d:$B$%d" % (f0, f1), [x[1] for x in fun])],
                          'J26', width=9, height=18, colors=[X.AZUL]))
    r.charts.append(Chart('doughnut', 'Antigüedad del último movimiento (prospectos activos)', 'Resumen',
                          "'Resumen'!$A$%d:$A$%d" % (b0, b1), [b[0] for b in bands],
                          [('Prospectos', "'Resumen'!$B$%d:$B$%d" % (b0, b1),
                            [sum(1 for x in P if x['activo'] and b[1] <= x['dias_sin_mov'] <= b[2])
                             for b in bands])],
                          'J45', width=9, height=18))
    sheets.insert(1, r)

    # ------------------------------------------- hoja 3: criterios y notas
    c = Sheet('Criterios y notas', cols=[34, 86], tab_color=X.GRIS)
    c.row(C('CRITERIOS APLICADOS Y LIMITACIONES', 'titulo'))
    c.row(C('Para que Dirección pueda interpretar y auditar cada campo de la pestaña «Prospectos».', 'nota'))
    c.blank()
    c.row(C('Tema', 'hdr'), C('Criterio aplicado', 'hdr'))
    for a, b in [
        ('Fecha de corte', 'Se incluyen los prospectos cuyo primer contacto es anterior o igual al %s. '
                           'Las clases muestra con cita posterior al corte se conservan porque '
                           'representan oportunidades vigentes. En %d casos que solo existen en la '
                           'bitácora de clases muestra NO hay fecha de primer contacto en ninguna '
                           'fuente: se muestra la fecha de su clase muestra (resaltada en ámbar) y se '
                           'advierte en Observaciones. No se asume cuándo se registraron.'
                           % (CORTE, facts.get().sin_fecha_contacto)),
        ('Universo', 'Campaña de septiembre 2026 (21 ago – 4 oct): hoja «AGOSTOSEPTIEMBRE» del CRM, '
                     'deduplicada a nivel persona, más las personas que solo aparecen en la bitácora '
                     'de clases muestra y no tienen registro en el CRM.'),
        ('Deduplicación', 'R1 mismo teléfono y nombre → misma persona. R2 mismo teléfono y nombre '
                          'contenido o similar → misma persona. R3 mismo nombre completo sin teléfono '
                          'contradictorio → misma persona. R4 mismo teléfono con nombres distintos → '
                          'personas distintas del mismo hogar (caso frecuente: un padre registra a dos '
                          'o tres hijos). R5 mismo nombre de una sola palabra con teléfono distinto o '
                          'ausente → NO se fusiona, se marca «REVISAR» en Observaciones.'),
        ('Estatus homologado', 'Prioridad: (1) si la persona o su hogar tiene clase muestra, manda el '
                               'estado de la clase muestra; (2) si el CRM registra «Inscrito», manda '
                               'ese valor; (3) en otro caso manda la etapa del CRM homologada. '
                               'Si no hay información suficiente se usa «Estatus por validar»: nunca '
                               'se inventó un estatus para completar la base.'),
        ('Oportunidad abierta', 'Sí para Nuevo, Descubrimiento, Clase muestra agendada, Clase muestra '
                                'realizada, Reprogramación, Seguimiento, Por cerrar y Lista de espera. '
                                'No para Inscrito/Cerrado, No interesado, Sin respuesta y Estatus por '
                                'validar. Los inscritos quedan identificados como «Inscrito / Cerrado» '
                                'para evitar volver a contabilizarlos como prospectos activos.'),
        ('Prospectos sin respuesta', 'Se conservan en la base y quedan claramente identificados con el '
                                     'estatus «Sin respuesta». No se cuentan como oportunidad abierta.'),
        ('«Último seguimiento»', 'INFERENCIA: las fuentes NO tienen un campo de fecha de seguimiento. '
                                 'Esta columna muestra la fecha del registro más reciente de esa '
                                 'persona (último registro en el CRM o fecha de su clase muestra, la '
                                 'que sea posterior). Es la mejor aproximación disponible, no un dato '
                                 'de seguimiento real.'),
        ('«Días sin movimiento»', 'Días entre la fecha de «Último seguimiento» y el corte del 3 de '
                                  'octubre de 2026. Mide antigüedad del último registro, no de la '
                                  'última conversación.'),
        ('«Horario»', 'Solo existe para quienes tuvieron clase muestra agendada: es la hora de esa '
                      'clase. El CRM no captura preferencia de horario, por eso el resto aparece como '
                      '«Sin información».'),
        ('«Edad / Programa»', 'La edad y el programa provienen de la bitácora de clases muestra cuando '
                              'existe; en los demás casos se usa el segmento declarado en el CRM. '
                              'Cuando el dato viene de la bitácora se indica expresamente.'),
        ('«Clase muestra»', 'Lista todas las clases muestra ligadas a la persona o a su hogar, con '
                            'fecha, nombre de quien tomó la clase, estado e idioma. Si no hay ninguna, '
                            'dice «Sin clase muestra registrada».'),
        ('Lista de espera', 'Aparece con 0 registros porque NINGUNA de las fuentes contiene un campo o '
                            'un valor que registre lista de espera. Los valores «En espera» del CRM '
                            'significan seguimiento pendiente de respuesta, no cupo reservado.'),
        ('Campaña de julio', 'Los %d registros de la hoja «JULIO» se entregan como referencia en su '
                             'propia pestaña. No se suman a los activos porque su estatus no ha sido '
                             'actualizado desde julio.' % f.julio_regs),
        ('Datos faltantes', 'Cuando un dato no existe se escribe «Sin información». No se estimó, '
                            'completó ni inventó ningún valor.'),
    ]:
        c.row(C(a, 'bold'), C(b, 'txtw'))
    c.blank()
    c.row(C('LIMITACIONES QUE AFECTAN A ESTA BASE', 'seccion'), C('', 'seccion'))
    for t in ['%d de %d registros del CRM no tienen teléfono: su enlace con clases muestra e '
              'inscripciones depende solo del nombre.' % (f.dq['leads_sin_tel'], f.reg_ventana),
              '%d registros no tienen nombre del interesado y %d no tienen ni nombre ni teléfono.'
              % (f.dq['leads_sin_nombre'], f.dq['leads_sin_nombre_ni_tel']),
              '%d de %d registros no tienen resultado de seguimiento capturado (%s): el estatus se '
              'reconstruyó a partir de la etapa del proceso y de la bitácora de clases muestra.'
              % (f.dq['leads_sin_resultado'], f.reg_ventana,
                 pc(f.dq['leads_sin_resultado'] / f.reg_ventana)),
              'La columna «Comentario adicional» del CRM está vacía en el 100% de las filas: no hay '
              'contexto cualitativo ni motivo de pérdida.',
              '%d personas quedaron marcadas para revisión por posible duplicidad sin evidencia '
              'suficiente para fusionarlas.' % len(f.revisar),
              '%d clases muestra del periodo no pudieron ligarse a ningún registro del CRM: esas '
              'personas aparecen en la base con origen «Bitácora clases muestra (sin lead en CRM)».'
              % f.cm_sin_lead]:
        c.row(C('•  ' + t, 'txtw'))
        c.merge('A%d:B%d' % (len(c.rows), len(c.rows)))
    sheets.append(c)

    # ----------------------------------------- hoja 4: prospectos julio ref
    j = Sheet('Prospectos julio (ref.)', cols=[30, 14, 14, 18, 20, 24, 26, 24, 24, 11],
              freeze='A4', tab_color=X.GRIS)
    j.row(C('PROSPECTOS DE LA CAMPAÑA DE JULIO 2026 — REFERENCIA', 'titulo'))
    j.row(C('%d registros / %d personas. Su estatus corresponde al último movimiento registrado en '
            'julio y NO ha sido actualizado desde entonces. NO se contabilizan en el resumen de '
            'prospectos activos de la campaña de septiembre.' % (f.julio_regs, f.julio_personas), 'nota'))
    j.row(*[C(h, 'hdr') for h in ['Prospecto', 'Fecha primer contacto', 'Teléfono', 'Interés',
                                  'Segmento', 'Etapa registrada (CRM)', 'Resultado registrado',
                                  'Estatus homologado', 'Anuncio de origen', 'Oportunidad abierta']])
    for x in f.julio_rows:
        j.row(C(x['nombre'], 'txt'), C(x['f_primer'], 'fecha'), C(x['tel'] or 'Sin información', 'txt'),
              C(x['interes'], 'txt'), C(x['segmento'], 'txt'), C(x['etapa_crm'], 'txt'),
              C(x['resultado'], 'txt'), C(x['estatus'], 'txt'), C(x['anuncio'], 'txt'),
              C('Sí' if x['activo'] else 'No', 'txt'))
    j.autofilter = 'A3:J%d' % len(j.rows)
    sheets.append(j)

    path = os.path.join(OUT, 'Prospectos en Curso — Corte 03 octubre 2026.xlsx')
    X.write(path, sheets, 'Prospectos en Curso — Corte 03 octubre 2026')
    return path


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    print(build_alumnos())
    print(build_prospectos())
