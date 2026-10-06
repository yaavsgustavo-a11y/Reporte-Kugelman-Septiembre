"""FUENTE UNICA DE VERDAD de la Campana Septiembre 2026.

Todos los entregables (Reporte, Base Maestra, Presentacion, bases operativas)
leen de aqui, lo que garantiza consistencia de cifras entre archivos (req. 32).
Ninguna cifra esta escrita a mano: todas se calculan de las 4 fuentes del repo.
"""
import collections, datetime
import core, link, build
from core import nstr, norm_name

WIN_INI = datetime.date(2026, 8, 21)
WIN_FIN = datetime.date(2026, 10, 4)
CORTE = datetime.date(2026, 10, 3)

ETIQ_ADS = {
    'video villas': 'Reel Villas de Pachuca',
    'video sin libros': 'Reel sin libros',
    'post grupos': 'Post Grupos Abiertos',
    'video testimoniales': 'Reel Testimoniales',
    'carrete': 'Post Carrete',
    'post clase': 'Post Clase de Inglés Gratis',
    'video clases': 'Reel informativo (inferido)',
    'video switch': 'Reel informativo (inferido)',
}
FORMATO_ADS = {
    'video villas': 'Video / Reel', 'video sin libros': 'Video / Reel',
    'video testimoniales': 'Video / Reel', 'video clases': 'Video / Reel',
    'video switch': 'Video / Reel', 'post grupos': 'Imagen / Post',
    'post clase': 'Imagen / Post', 'carrete': 'Carrusel',
}

ESTATUS_ORDEN = [
    'Nuevo', 'Descubrimiento', 'Clase muestra agendada', 'Clase muestra realizada',
    'Reprogramación', 'Seguimiento', 'Por cerrar', 'Lista de espera',
    'Sin respuesta', 'No interesado', 'Inscrito / Cerrado', 'Estatus por validar',
]
ACTIVOS = {'Nuevo', 'Descubrimiento', 'Clase muestra agendada', 'Clase muestra realizada',
           'Reprogramación', 'Seguimiento', 'Por cerrar', 'Lista de espera'}


def semana(d):
    k = (d - WIN_INI).days // 7
    ini = WIN_INI + datetime.timedelta(days=7 * k)
    fin = min(ini + datetime.timedelta(days=6), WIN_FIN)
    return k, '%d/%02d–%d/%02d' % (ini.day, ini.month, fin.day, fin.month)


def etapa_fase(d):
    if d is None:
        return 'Sin fecha'
    if d < datetime.date(2026, 9, 1):
        return '21–31 ago (arranque)'
    if d <= datetime.date(2026, 9, 30):
        return 'Septiembre (principal)'
    return '1–4 oct (cierre)'


class Facts:
    def __init__(self):
        M = build.build()
        self.M = M
        self.P = M['personas']
        self.cms = M['cms_all']
        self.cmv = M['cms']
        self.alumnos = M['alumnos']
        self.cm2p = M['cm2p']
        self.cm2a = M['cm2a']
        self.cm_person = M['cm_person']
        self.atribuido = M['atribuido']
        self._pauta()
        self._crm()
        self._clases()
        self._inscripciones()
        self._funnel()
        self._anuncios()
        self._temporal()
        self._segmentos()
        self._estatus()
        self._calidad()

    # ------------------------------------------------------------- PAUTA
    def _pauta(self):
        M = self.M
        self.presupuesto = M['presupuesto']
        self.gasto = M['gasto']
        self.uso_presupuesto = self.gasto / self.presupuesto
        self.impresiones = M['impresiones']
        self.alcance_suma = M['alcance_suma']
        self.alcance_min = M['alcance_max_anuncio']
        self.conversaciones = M['conversaciones']
        self.cpm = self.gasto / self.impresiones * 1000
        self.costo_conversacion = self.gasto / self.conversaciones
        self.frec_sobre_suma = self.impresiones / self.alcance_suma
        self.frec_sobre_min = self.impresiones / self.alcance_min
        self.ads = sorted(M['ads'], key=lambda a: -a['gasto'])

    # ------------------------------------------------------------- CRM
    def _crm(self):
        M = self.M
        self.reg_hoja = M['n_leads_hoja']
        self.reg_ventana = M['n_reg_ventana']
        self.reg_fuera = len(M['leads_fuera_ventana'])
        self.pers_total = M['n_personas']
        self.leads = M['n_personas_ventana']          # <-- LEADS oficiales (personas unicas)
        self.dup_fusionados = M['dup_fusionados']
        self.tasa_duplicidad = self.dup_fusionados / self.reg_hoja
        self.cobertura_crm = self.reg_ventana / self.conversaciones
        self.gap_registro = self.conversaciones - self.reg_ventana
        self.reg_julio = M['n_leads_julio']
        self.oct_copia = M['n_leads_oct_copia']
        self.oct_dup = M['leads_oct_duplicados']

    # ------------------------------------------------------- CLASES MUESTRA
    def _clases(self):
        M = self.M
        self.cm_agendadas = M['cm_agendadas']
        self.cm_realizadas = M['cm_realizadas']
        self.cm_noshow = M['cm_noshow']
        self.cm_pendientes = M['cm_pendientes']
        self.cm_reprogramadas = M['cm_reprogramadas']
        self.cm_inscritos_reg = M['cm_inscritos']
        self.cm_personas = M['cm_personas_unicas']
        self.cm_pers_realizadas = M['cm_realizadas_personas']
        self.cm_con_lead = M['cm_con_lead']
        self.cm_sin_lead = self.cm_agendadas - self.cm_con_lead
        self.tasa_asistencia = self.cm_realizadas / self.cm_agendadas
        self.tasa_noshow = self.cm_noshow / self.cm_agendadas
        self.tasa_reprog = self.cm_reprogramadas / self.cm_agendadas
        # CM "solicitadas" segun CRM (campo Registrado a clase muestra)
        self.cm_solicitadas_crm = sum(1 for p in self.P if p['en_ventana'] and p['registrado_cm'] == 'Si')
        self.cm_etapa_crm = sum(1 for p in self.P if p['en_ventana']
                                and 'Asistencia a Clase muestra' in p['etapas'])

    # ------------------------------------------------------- INSCRIPCIONES
    def _inscripciones(self):
        """Universo de inscripciones con nivel de evidencia explicito."""
        filas = []
        vistos = {}
        for ci, c in enumerate(self.cms):
            if not c['en_ventana'] or c['estado'] != 'Realizada - Inscrito':
                continue
            pk = self.cm_person[ci]
            a = self.cm2a.get(ci)
            lk = self.cm2p.get(ci)
            filas.append({
                'pk': pk, 'nombre': c['nombre'], 'fecha_cm': c['fecha'], 'categoria': c['categoria'],
                'idioma': c['idioma'], 'edad': c['edad'], 'hora': c['hora'],
                'medio_agenda': c['medio_agenda'],
                'matricula': self.alumnos[a[0]]['matricula'] if a else None,
                'nombre_padron': self.alumnos[a[0]]['completo'] if a else '',
                'edad_padron': self.alumnos[a[0]]['edad'] if a else None,
                'lead': self.P[lk[0]]['nombre'] if lk else '',
                'anuncio': (self.P[lk[0]]['anuncio'] or 'Sin anuncio registrado') if lk else '',
                'confianza': lk[2] if lk else 'Sin enlace',
                'atribuible': bool(lk and lk[2] in ('Alta', 'Media', 'Hogar')),
                'evidencia': 'Bitácora CM = Inscrito' + (' + padrón RGA' if a else ''),
                'fuente': 'Bitácora clases muestra',
            })
        # --- inscripcion confirmada por CRM + padron que la bitacora no refleja
        self.insc_conflicto = []
        self.insc_por_validar = []
        for ci, c in enumerate(self.cms):
            if not c['en_ventana'] or c['estado'] == 'Realizada - Inscrito':
                continue
            a = self.cm2a.get(ci)
            if not a:
                continue
            lk = self.cm2p.get(ci)
            crm_insc = bool(lk and 'Inscrito' in (self.P[lk[0]]['resultado_crm'] or ''))
            if self.cm_person[ci] in [f['pk'] for f in filas]:
                continue
            reg = {
                'pk': self.cm_person[ci], 'nombre': c['nombre'], 'fecha_cm': c['fecha'],
                'categoria': c['categoria'], 'idioma': c['idioma'], 'edad': c['edad'],
                'hora': c['hora'], 'medio_agenda': c['medio_agenda'],
                'matricula': self.alumnos[a[0]]['matricula'],
                'nombre_padron': self.alumnos[a[0]]['completo'],
                'edad_padron': self.alumnos[a[0]]['edad'],
                'lead': self.P[lk[0]]['nombre'] if lk else '',
                'anuncio': (self.P[lk[0]]['anuncio'] or 'Sin anuncio registrado') if lk else '',
                'confianza': lk[2] if lk else 'Sin enlace',
                'atribuible': bool(lk and lk[2] in ('Alta', 'Media', 'Hogar')),
                'estado_cm': c['estado'],
            }
            if crm_insc:
                reg['evidencia'] = 'CRM = Inscrito + padrón RGA (bitácora CM dice "%s")' % c['estado']
                reg['fuente'] = 'CRM + padrón'
                filas.append(reg)
                self.insc_conflicto.append(reg)
            else:
                reg['evidencia'] = ('Aparece en padrón RGA (matrícula %s) pero bitácora CM dice "%s" '
                                    'y el CRM no registra inscripción' % (reg['matricula'], c['estado']))
                self.insc_por_validar.append(reg)

        self.insc_filas = sorted(filas, key=lambda f: (f['fecha_cm'], f['nombre']))
        pks = {}
        for f in self.insc_filas:
            pks.setdefault(f['pk'], []).append(f)
        self.insc_personas = len(pks)
        self.insc_eventos = len(self.insc_filas)
        self.insc_atribuibles = len([k for k, v in pks.items() if any(x['atribuible'] for x in v)])
        self.insc_eventos_atrib = len([f for f in self.insc_filas if f['atribuible']])
        self.insc_sin_padron = [f for f in self.insc_filas if not f['matricula']]
        self.insc_no_atribuibles = [v[0] for k, v in pks.items() if not any(x['atribuible'] for x in v)]
        self.cac = self.gasto / self.insc_atribuibles
        self.cac_total = self.gasto / self.insc_personas

    # ------------------------------------------------------------- FUNNEL
    def _funnel(self):
        self.t_lead_cm = self.cm_agendadas / self.leads
        self.t_cm_real = self.cm_realizadas / self.cm_agendadas
        self.t_real_insc = self.insc_personas / self.cm_pers_realizadas
        self.t_lead_insc = self.insc_personas / self.leads
        self.t_lead_insc_atrib = self.insc_atribuibles / self.leads
        self.cpl_persona = self.gasto / self.leads
        self.costo_cm_agendada = self.gasto / self.cm_agendadas
        self.costo_cm_realizada = self.gasto / self.cm_realizadas
        self.etapas = [
            ('Impresiones', self.impresiones, 'Meta Ads'),
            ('Alcance (personas)', self.alcance_min, 'Meta Ads – cota mínima'),
            ('Conversaciones iniciadas', self.conversaciones, 'Meta Ads'),
            ('Leads / prospectos (personas)', self.leads, 'CRM campaña, deduplicado'),
            ('CM agendadas', self.cm_agendadas, 'Bitácora clases muestra'),
            ('CM realizadas', self.cm_realizadas, 'Bitácora clases muestra'),
            ('Inscripciones', self.insc_personas, 'Bitácora + CRM + padrón'),
        ]

    # ----------------------------------------------------------- ANUNCIOS
    def _anuncios(self):
        self.crm_reg_ad = self.M['crm_por_anuncio']
        self.crm_pers_ad = self.M['crm_pers_por_anuncio']
        cm_ad = collections.Counter()
        insc_ad = collections.Counter()
        for ci, c in enumerate(self.cms):
            if not c['en_ventana'] or not self.atribuido(ci):
                continue
            p = self.P[self.cm2p[ci][0]]
            cm_ad[p['anuncio'] or 'Sin anuncio registrado'] += 1
        for f in self.insc_filas:
            if f['atribuible']:
                insc_ad[f['anuncio'] or 'Sin anuncio registrado'] += 1
        self.cm_ad = cm_ad
        self.insc_ad = insc_ad
        # tabla consolidada por anuncio de Meta
        tab = []
        for a in self.ads:
            etiq = ETIQ_ADS.get(a['anuncio'], '')
            base = etiq.replace(' (inferido)', '')
            tab.append({
                'anuncio': a['anuncio'], 'etiqueta': etiq, 'formato': FORMATO_ADS.get(a['anuncio'], ''),
                'gasto': a['gasto'], 'conv': a['resultados'],
                'costo_conv': a['gasto'] / a['resultados'] if a['resultados'] else None,
                'impresiones': a['impresiones'], 'alcance': a['alcance'],
                'cpm': a['gasto'] / a['impresiones'] * 1000,
                'frec': a['impresiones'] / a['alcance'],
                'pct_gasto': a['gasto'] / self.gasto,
                'pct_conv': a['resultados'] / self.conversaciones,
                'calidad': a['calidad'], 'interaccion': a['interaccion'], 'conversion': a['conversion'],
                'entrega': a['entrega'],
                'crm_reg': self.crm_reg_ad.get(base, 0), 'crm_pers': self.crm_pers_ad.get(base, 0),
                'cm': cm_ad.get(base, 0), 'insc': insc_ad.get(base, 0),
            })
        self.tab_ads = tab
        self.ad_top_vol = max(tab, key=lambda r: r['conv'])
        elig = [r for r in tab if r['conv'] >= 10]
        self.ad_top_efic = min(elig, key=lambda r: r['costo_conv'])
        self.ad_peor_efic = max(elig, key=lambda r: r['costo_conv'])
        self.ad_peor_abs = max([r for r in tab if r['gasto'] > 100], key=lambda r: r['costo_conv'])

    # ----------------------------------------------------------- TEMPORAL
    def _temporal(self):
        fases = collections.OrderedDict()
        for k in ('21–31 ago (arranque)', 'Septiembre (principal)', '1–4 oct (cierre)'):
            fases[k] = dict(leads=0, cm_ag=0, cm_re=0, insc=0, dias=0)
        fases['21–31 ago (arranque)']['dias'] = 11
        fases['Septiembre (principal)']['dias'] = 30
        fases['1–4 oct (cierre)']['dias'] = 4
        for p in self.P:
            if p['en_ventana']:
                fases[etapa_fase(p['f_primer'])]['leads'] += 1
        for c in self.cmv:
            f = fases[etapa_fase(c['fecha'])]
            f['cm_ag'] += 1
            if c['estado'].startswith('Realizada'):
                f['cm_re'] += 1
        for fi in self.insc_filas:
            fases[etapa_fase(fi['fecha_cm'])]['insc'] += 1
        self.fases = fases
        # serie semanal
        L = collections.Counter(); CA = collections.Counter()
        CR = collections.Counter(); IN = collections.Counter(); lab = {}
        for p in self.P:
            if p['en_ventana']:
                k, l = semana(p['f_primer']); L[k] += 1; lab[k] = l
        for c in self.cmv:
            k, l = semana(c['fecha']); lab[k] = l; CA[k] += 1
            if c['estado'].startswith('Realizada'):
                CR[k] += 1
        for fi in self.insc_filas:
            k, l = semana(fi['fecha_cm']); lab[k] = l; IN[k] += 1
        self.semanas = [(lab[k], L[k], CA[k], CR[k], IN[k]) for k in sorted(lab)]
        self.serie_leads = self.M['serie_leads']
        self.pico_leads = max(self.serie_leads.items(), key=lambda kv: kv[1])

    # ---------------------------------------------------------- SEGMENTOS
    def _segmentos(self):
        self.cm_cat = self.M['cm_por_categoria']
        self.cm_idioma = self.M['cm_por_idioma']
        self.insc_cat = collections.Counter(f['categoria'] or 'Sin información' for f in self.insc_filas)
        self.insc_idioma = collections.Counter(f['idioma'] or 'Sin información' for f in self.insc_filas)
        self.cm_medio_ag = self.M['cm_por_medio_agenda']
        self.insc_medio_ag = collections.Counter(f['medio_agenda'] or 'Sin información'
                                                 for f in self.insc_filas)
        self.cm_hora = self.M['cm_por_hora']
        self.edad_cm = self.M['edad_cm']
        eda = [f['edad'] for f in self.insc_filas if f['edad'] is not None]
        self.edad_insc = {'n': len(eda), 'min': min(eda), 'max': max(eda),
                          'prom': round(sum(eda) / len(eda), 1)}
        self.leads_medio = self.M['leads_por_medio']
        self.leads_canal = self.M['leads_por_canal']
        # conversion CM->inscripcion por categoria
        self.conv_cat = {}
        for k, v in self.cm_cat.items():
            real = sum(1 for c in self.cmv if (c['categoria'] or 'Sin información') == k
                       and c['estado'].startswith('Realizada'))
            self.conv_cat[k] = (v, real, self.insc_cat.get(k, 0))

    # ------------------------------------------------------------ ESTATUS
    def _estatus(self):
        """Homologa estatus por persona y construye el pipeline al corte 03-oct."""
        ETAPA = {
            'nuevo lead': 'Nuevo',
            'contactado': 'Descubrimiento',
            'informacion enviada': 'Descubrimiento',
            'esperando respuesta': 'Seguimiento',
            'pendiente de decision': 'Por cerrar',
            'asistencia a clase muestra': 'Clase muestra agendada',
            'cerrado': 'No interesado',
        }
        CM_EST = {
            'Realizada - Inscrito': 'Inscrito / Cerrado',
            'Realizada - Sin cierre': 'Clase muestra realizada',
            'No show': 'Seguimiento',
            'Agendada pendiente': 'Clase muestra agendada',
            'Reprogramada': 'Reprogramación',
        }
        # CM por persona-lead (incluye hogar) y CM sin lead
        cm_by_person = collections.defaultdict(list)
        cm_huerfanas = []
        for ci, c in enumerate(self.cms):
            lk = self.cm2p.get(ci)
            if lk and lk[2] in ('Alta', 'Media', 'Hogar'):
                cm_by_person[lk[0]].append(ci)
            else:
                cm_huerfanas.append(ci)
        insc_pk = set(f['pk'] for f in self.insc_filas)
        insc_nombres = set(norm_name(f['nombre']) for f in self.insc_filas)

        regs = []
        for pi, p in enumerate(self.P):
            if p['f_primer'] > CORTE:
                continue
            cis = sorted(cm_by_person.get(pi, []), key=lambda ci: self.cms[ci]['fecha'])
            cm_txt, cm_estado, cm_fecha, cm_det = '', '', None, []
            for ci in cis:
                c = self.cms[ci]
                cm_det.append('%s %s (%s%s)' % (c['fecha'].strftime('%d/%m'), c['nombre'],
                                                c['estado'], ', ' + c['idioma'] if c['idioma'] else ''))
            if cis:
                # prioridad: inscrito > agendada pendiente > reprogramada > realizada > no show
                prio = {'Realizada - Inscrito': 0, 'Agendada pendiente': 1, 'Reprogramada': 2,
                        'Realizada - Sin cierre': 3, 'No show': 4}
                best = sorted(cis, key=lambda ci: (prio.get(self.cms[ci]['estado'], 9),
                                                   -self.cms[ci]['fecha'].toordinal()))[0]
                cm_estado = self.cms[best]['estado']
                cm_fecha = self.cms[best]['fecha']
                cm_txt = '; '.join(cm_det)
            est = None
            if any(self.cms[ci]['estado'] == 'Realizada - Inscrito' for ci in cis):
                est = 'Inscrito / Cerrado'
            elif cm_estado:
                est = CM_EST.get(cm_estado, 'Estatus por validar')
            # inscripcion confirmada por CRM aunque la bitacora diga otra cosa
            if 'Inscrito' in (p['resultado_crm'] or ''):
                est = 'Inscrito / Cerrado'
            if est is None:
                e = norm_name(p['etapa'])
                est = ETAPA.get(e)
                if est == 'No interesado' and 'Inscrito' in (p['resultado_crm'] or ''):
                    est = 'Inscrito / Cerrado'
                if est is None:
                    est = 'Estatus por validar'
            if norm_name(p['resultado_crm']) == 'sin respuesta':
                est = 'Sin respuesta'
            if norm_name(p['resultado_crm']) == 'cancelo proceso':
                est = 'No interesado'
            obs = []
            if p['n_registros'] > 1:
                obs.append('%d registros fusionados (filas %s)'
                           % (p['n_registros'], ','.join(str(f) for f in p['filas'])))
            for n in p['notas_dedup']:
                if n.startswith('R4'):
                    obs.append(n[3:])
                if n.startswith('R5'):
                    obs.append('REVISAR: ' + n[3:])
            if cm_estado == 'No show':
                obs.append('No asistió a la clase muestra programada')
            if p['resultado_crm']:
                obs.append('CRM resultado: ' + p['resultado_crm'])
            regs.append({
                'id': 'P%03d' % (pi + 1), 'nombre': p['nombre'] or '(sin nombre registrado)',
                'tel': p['tel'], 'f_primer': p['f_primer'], 'f_ultimo': p['f_ultimo'],
                'medio': p['medio'] or 'Sin información', 'anuncio': p['anuncio'] or 'Sin anuncio registrado',
                'canal': p['canal'] or 'Sin información',
                'interes': p['idioma'] or 'Sin información',
                'segmento': p['segmento'] or 'Sin información',
                'cm': cm_txt or 'Sin clase muestra registrada',
                'cm_fecha': cm_fecha, 'cm_estado': cm_estado or 'Sin información',
                'estatus': est, 'etapa_crm': p['etapa'] or 'Sin información',
                'activo': est in ACTIVOS, 'obs': ' | '.join(obs),
                'origen': 'CRM campaña (hoja AGOSTOSEPTIEMBRE)',
                'en_ventana': p['en_ventana'],
                'edad': None, 'hora': '', 'fecha_inferida': False,
            })
        # completa edad / horario / programa desde la bitacora de CM
        idx = {r['id']: r for r in regs}
        for pi, p in enumerate(self.P):
            rid = 'P%03d' % (pi + 1)
            if rid not in idx:
                continue
            cis = cm_by_person.get(pi, [])
            if cis:
                c = self.cms[sorted(cis, key=lambda ci: -self.cms[ci]['fecha'].toordinal())[0]]
                idx[rid]['edad'] = c['edad']
                idx[rid]['hora'] = c['hora']
                if idx[rid]['segmento'] == 'Sin información' and c['categoria']:
                    idx[rid]['segmento'] = c['categoria'] + ' (de bitácora CM)'
                if idx[rid]['interes'] == 'Sin información' and c['idioma']:
                    idx[rid]['interes'] = c['idioma'] + ' (de bitácora CM)'

        # --- personas que solo existen en la bitacora de clases muestra
        huerf = collections.defaultdict(list)
        for ci in cm_huerfanas:
            huerf[self.cm_person[ci]].append(ci)
        for pk, cis in huerf.items():
            cis = sorted(cis, key=lambda ci: self.cms[ci]['fecha'])
            c = self.cms[cis[-1]]
            if c['fecha'] is None:
                continue
            prio = {'Realizada - Inscrito': 0, 'Agendada pendiente': 1, 'Reprogramada': 2,
                    'Realizada - Sin cierre': 3, 'No show': 4}
            best = sorted(cis, key=lambda ci: (prio.get(self.cms[ci]['estado'], 9),
                                               -self.cms[ci]['fecha'].toordinal()))[0]
            cb = self.cms[best]
            if cb['fecha'] < datetime.date(2026, 8, 21):
                continue
            est = CM_EST.get(cb['estado'], 'Estatus por validar')
            a = self.cm2a.get(best)
            obs = ['Sin registro en el CRM de campaña: no es atribuible a la pauta']
            post_corte = cb['fecha'] > CORTE
            if post_corte:
                obs.append('Fecha de primer contacto NO disponible en las fuentes: se muestra la fecha '
                           'de la clase muestra agendada, posterior al corte del 3-oct')
            if a:
                obs.append('Localizado en padrón RGA (matrícula %s)' % self.alumnos[a[0]]['matricula'])
            if cb['medio_agenda']:
                obs.append('Agendó por: ' + cb['medio_agenda'])
            regs.append({
                'id': 'CM%03d' % (pk + 1), 'nombre': cb['nombre'], 'tel': cb['tel10'],
                'f_primer': cb['fecha'], 'f_ultimo': self.cms[cis[-1]]['fecha'],
                'medio': cb['medio_agenda'] or 'Sin información',
                'anuncio': 'Sin anuncio registrado', 'canal': cb['medio_agenda'] or 'Sin información',
                'interes': cb['idioma'] or 'Sin información',
                'segmento': cb['categoria'] or 'Sin información',
                'cm': '; '.join('%s %s (%s)' % (self.cms[ci]['fecha'].strftime('%d/%m'),
                                                self.cms[ci]['nombre'], self.cms[ci]['estado'])
                                for ci in cis),
                'cm_fecha': cb['fecha'], 'cm_estado': cb['estado'], 'estatus': est,
                'etapa_crm': 'Sin información', 'activo': est in ACTIVOS,
                'obs': ' | '.join(obs), 'origen': 'Bitácora clases muestra (sin lead en CRM)',
                'en_ventana': True, 'edad': cb['edad'], 'hora': cb['hora'],
                'fecha_inferida': post_corte,
            })
        # dias sin movimiento al corte y marca de inactividad operativa
        for r in regs:
            r['dias_sin_mov'] = (CORTE - max(r['f_ultimo'], r['cm_fecha'] or r['f_ultimo'])).days
            r['dormido'] = (r['estatus'] == 'Descubrimiento' and r['dias_sin_mov'] >= 15)
        regs.sort(key=lambda r: (r['f_primer'], r['nombre']))
        self.pipeline = regs
        self.dormidos = sum(1 for r in regs if r['dormido'])
        self.sin_fecha_contacto = sum(1 for r in regs if r.get('fecha_inferida'))
        self.estatus_count = collections.Counter(r['estatus'] for r in regs)
        self.activos = sum(1 for r in regs if r['activo'])
        self.pipe_funnel = {
            'prospectos': len(regs),
            'cm_agendadas': sum(1 for r in regs if r['cm_estado'] != 'Sin información'),
            'cm_realizadas': sum(1 for r in regs if r['cm_estado'].startswith('Realizada')),
            'inscritos': sum(1 for r in regs if r['estatus'] == 'Inscrito / Cerrado'),
        }
        self.pipe_insc_regs = self.pipe_funnel['inscritos']
        # base julio (referencia)
        jl = core.load_leads_julio()
        gj, nj, hj = link.dedupe_leads(jl)
        self.julio_regs = len(jl)
        self.julio_personas = len(gj)
        jest = collections.Counter()
        self.julio_rows = []
        for g in gj:
            rs = [jl[i] for i in g]
            rs.sort(key=lambda r: r['fecha'])
            etapas = [r['etapa'] for r in rs if nstr(r['etapa'])]
            res = '; '.join(sorted(set(nstr(r['resultado']) for r in rs if nstr(r['resultado']))))
            e = norm_name(etapas[-1]) if etapas else ''
            est = ETAPA.get(e, 'Estatus por validar')
            if 'Inscrito' in res:
                est = 'Inscrito / Cerrado'
            elif norm_name(res) == 'sin respuesta':
                est = 'Sin respuesta'
            elif 'Canceló' in res or 'cancelo' in norm_name(res):
                est = 'No interesado'
            jest[est] += 1
            self.julio_rows.append({
                'nombre': nstr(rs[0]['nombre']) or '(sin nombre registrado)',
                'tel': rs[0]['tel10'], 'f_primer': rs[0]['fecha'],
                'medio': nstr(rs[0]['medio']) or 'Sin información',
                'anuncio': nstr(rs[0]['anuncio']) or 'Sin anuncio registrado',
                'interes': nstr(rs[0]['idioma']) or 'Sin información',
                'segmento': nstr(rs[0]['segmento']) or 'Sin información',
                'etapa_crm': etapas[-1] if etapas else 'Sin información',
                'resultado': res or 'Sin información', 'estatus': est,
                'activo': est in ACTIVOS,
            })
        self.julio_estatus = jest
        self.julio_activos = sum(1 for r in self.julio_rows if r['activo'])
        self.julio_rows.sort(key=lambda r: (r['f_primer'], r['nombre']))

    # ------------------------------------------------------------ CALIDAD
    def _calidad(self):
        dq = dict(self.M['dq'])
        dq['rga_n'] = self.M['rga_n']
        dq['rga_dup'] = self.M['rga_dup']
        dq['rga_activos'] = self.M['rga_activos']
        dq['rga_sexo'] = self.M['rga_sexo']
        self.dq = dq
        # inconsistencias de edad bitacora vs padron
        self.edad_dif = []
        for f in self.insc_filas:
            if f['matricula'] and f['edad'] is not None and isinstance(f['edad_padron'], int):
                d = abs(f['edad'] - f['edad_padron'])
                if d >= 3:
                    self.edad_dif.append((f['nombre'], f['edad'], f['matricula'], f['edad_padron'], d))
        # registros marcados para revision
        self.revisar = [p for p in self.P if any(n.startswith('R5') for n in p['notas_dedup'])]
        self.sin_id = [p for p in self.P if p['sin_identificador']]
        self.anio_typos = self.M['dq']['cm_anio_detalle']


_F = None


def get():
    global _F
    if _F is None:
        _F = Facts()
    return _F


if __name__ == '__main__':
    f = get()
    print('PAUTA    presupuesto %.0f gasto %.2f (%.1f%%) impr %d alcance>=%d conv %d CPM %.2f'
          % (f.presupuesto, f.gasto, f.uso_presupuesto * 100, f.impresiones, f.alcance_min,
             f.conversaciones, f.cpm))
    print('CRM      reg hoja %d / ventana %d / personas %d (dup fusionados %d = %.1f%%) cobertura %.1f%%'
          % (f.reg_hoja, f.reg_ventana, f.leads, f.dup_fusionados, f.tasa_duplicidad * 100,
             f.cobertura_crm * 100))
    print('CM       agendadas %d realizadas %d noshow %d (%.1f%%) reprog %d pend %d personas %d'
          % (f.cm_agendadas, f.cm_realizadas, f.cm_noshow, f.tasa_noshow * 100,
             f.cm_reprogramadas, f.cm_pendientes, f.cm_personas))
    print('INSCR    eventos %d personas %d atribuibles %d | sin padron %d | por validar %d | conflicto %d'
          % (f.insc_eventos, f.insc_personas, f.insc_atribuibles, len(f.insc_sin_padron),
             len(f.insc_por_validar), len(f.insc_conflicto)))
    print('FUNNEL   lead->CM %.1f%% CM->real %.1f%% real->insc %.1f%% lead->insc %.1f%% (atrib %.1f%%)'
          % (f.t_lead_cm * 100, f.t_cm_real * 100, f.t_real_insc * 100, f.t_lead_insc * 100,
             f.t_lead_insc_atrib * 100))
    print('COSTOS   CPL %.2f  C/CMag %.2f  C/CMreal %.2f  CAC %.2f (total %.2f)'
          % (f.cpl_persona, f.costo_cm_agendada, f.costo_cm_realizada, f.cac, f.cac_total))
    print('PIPELINE registros %d activos %d' % (len(f.pipeline), f.activos))
    for k in ESTATUS_ORDEN:
        print('   %-26s %d' % (k, f.estatus_count.get(k, 0)))
    print('   funnel pipeline:', f.pipe_funnel)
    print('JULIO    reg %d personas %d activos %d' % (f.julio_regs, f.julio_personas, f.julio_activos))
    print('SEMANAS', f.semanas)
    print('FASES')
    for k, v in f.fases.items():
        print('   %-24s %s' % (k, v))
