"""Deduplicacion de leads y enlace Lead -> Clase muestra -> Inscripcion.

CRITERIOS DOCUMENTADOS (ver pestana 'Calidad de datos'):
  R1  Mismo telefono (10 digitos) + mismo nombre normalizado           -> MISMA persona (fusion)
  R2  Mismo telefono + un nombre contiene al otro / alta similitud     -> MISMA persona (fusion)
  R3  Mismo nombre completo (>=2 palabras) y ambos SIN telefono        -> MISMA persona (fusion)
  R4  Mismo telefono + nombres distintos                              -> PERSONAS DISTINTAS
                                                                          (se vinculan como mismo hogar/contacto)
  R5  Mismo nombre de una sola palabra, telefono distinto o ausente    -> NO se fusiona, se marca 'Revisar duplicado'
  R6  Registro sin nombre y sin telefono                              -> no fusionable, se marca 'No identificado'
"""
import collections, datetime
import core
from core import norm_name, norm_phone, name_sim, same_person_names, tokens


class Union:
    def __init__(self, n):
        self.p = list(range(n))

    def find(self, a):
        while self.p[a] != a:
            self.p[a] = self.p[self.p[a]]
            a = self.p[a]
        return a

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.p[max(ra, rb)] = min(ra, rb)


def dedupe_leads(leads):
    """Devuelve (grupos, notas) donde grupos = lista de listas de indices."""
    n = len(leads)
    uf = Union(n)
    notas = collections.defaultdict(list)

    by_phone = collections.defaultdict(list)
    by_name = collections.defaultdict(list)
    for i, r in enumerate(leads):
        if r['tel10']:
            by_phone[r['tel10']].append(i)
        if r['nn']:
            by_name[r['nn']].append(i)

    # --- R1 / R2 / R4 sobre telefono compartido
    hogar = {}
    for ph, ix in by_phone.items():
        if len(ix) > 1:
            for i in ix:
                hogar[i] = ph
        for a in range(len(ix)):
            for b in range(a + 1, len(ix)):
                i, j = ix[a], ix[b]
                ni, nj = leads[i]['nn'], leads[j]['nn']
                if not ni or not nj:
                    uf.union(i, j)
                    notas[i].append('R2 tel igual, un registro sin nombre')
                    continue
                if ni == nj:
                    uf.union(i, j)
                    notas[i].append('R1 tel+nombre identicos')
                elif tokens(ni) <= tokens(nj) or tokens(nj) <= tokens(ni) or name_sim(ni, nj) >= 0.90:
                    uf.union(i, j)
                    notas[i].append('R2 tel igual, nombre contenido/similar')
                else:
                    notas[i].append('R4 tel compartido con %s (mismo hogar, persona distinta)' % leads[j]['nombre'])
                    notas[j].append('R4 tel compartido con %s (mismo hogar, persona distinta)' % leads[i]['nombre'])

    # --- R3 nombre completo identico sin telefono
    for nm, ix in by_name.items():
        if len(ix) < 2:
            continue
        multi = len(nm.split()) >= 2
        for a in range(len(ix)):
            for b in range(a + 1, len(ix)):
                i, j = ix[a], ix[b]
                if uf.find(i) == uf.find(j):
                    continue
                pi, pj = leads[i]['tel10'], leads[j]['tel10']
                if pi and pj and pi != pj:
                    notas[i].append('R5 mismo nombre, telefono distinto -> NO fusionado')
                    notas[j].append('R5 mismo nombre, telefono distinto -> NO fusionado')
                    continue
                if multi:
                    uf.union(i, j)
                    notas[i].append('R3 nombre completo identico sin telefono contradictorio')
                else:
                    notas[i].append('R5 nombre de una sola palabra repetido -> revisar')
                    notas[j].append('R5 nombre de una sola palabra repetido -> revisar')

    grupos = collections.defaultdict(list)
    for i in range(n):
        grupos[uf.find(i)].append(i)
    return [sorted(v) for v in grupos.values()], notas, hogar


# ----------------------------------------------------- enlace persona <-> CM

def link_cm(personas, cms):
    """Para cada clase muestra intenta ubicar a la persona-lead de origen.

    Devuelve dict idx_cm -> (idx_persona, criterio, confianza)
      confianza 'Alta'  : telefono identico + nombre compatible
      confianza 'Media' : nombre completo compatible (>=2 tokens) sin telefono contradictorio
      confianza 'Hogar' : telefono identico, nombre distinto -> mismo contacto/hogar
                          (el lead es el padre/madre o el contacto que escribio)
      confianza 'Baja'  : unica coincidencia por nombre de una sola palabra -> se marca, no se usa
                          para atribucion
    """
    out = {}
    tel_idx = collections.defaultdict(list)
    for pi, p in enumerate(personas):
        for t in p['tels']:
            tel_idx[t].append(pi)

    for ci, c in enumerate(cms):
        # ---- 1) telefono identico
        if c['tel10'] and c['tel10'] in tel_idx:
            cands = tel_idx[c['tel10']]
            fuerte = [pi for pi in cands
                      if any(core.name_match_level(c['nn'], nn) in ('exacta', 'fuerte')
                             for nn in personas[pi]['nombres_norm'])]
            if fuerte:
                out[ci] = (fuerte[0], 'Teléfono + nombre', 'Alta')
            else:
                debil = [pi for pi in cands
                         if any(core.name_match_level(c['nn'], nn) == 'debil'
                                for nn in personas[pi]['nombres_norm'])]
                if debil:
                    out[ci] = (debil[0], 'Teléfono + nombre parcial', 'Alta')
                else:
                    out[ci] = (cands[0], 'Teléfono del contacto (mismo hogar)', 'Hogar')
            continue
        # ---- 2) solo nombre, sin telefono contradictorio
        if not c['nn']:
            continue
        fuertes, debiles = [], []
        for pi, p in enumerate(personas):
            if c['tel10'] and p['tels'] and c['tel10'] not in p['tels']:
                continue  # telefonos conocidos y distintos -> no es la misma persona
            for nn in p['nombres_norm']:
                lvl = core.name_match_level(c['nn'], nn)
                if lvl in ('exacta', 'fuerte'):
                    fuertes.append((name_sim(c['nn'], nn), pi))
                    break
                if lvl == 'debil':
                    debiles.append((name_sim(c['nn'], nn), pi))
                    break
        if fuertes:
            fuertes.sort(reverse=True)
            out[ci] = (fuertes[0][1], 'Nombre completo', 'Media')
        elif len(debiles) == 1:
            out[ci] = (debiles[0][1], 'Nombre parcial (1 token)', 'Baja')
    return out


# ------------------------------------------------- enlace CM/persona <-> alumno

def link_alumnos(cms, alumnos):
    """Empareja un registro de clase muestra con el padron de alumnos (RGA).
    Criterio estricto: coincidencia exacta o fuerte (>=2 tokens).  Se descartan
    coincidencias por un solo nombre de pila para evitar falsos positivos.
    Devuelve idx_cm -> (idx_alumno, nivel)"""
    out = {}
    for ci, c in enumerate(cms):
        if not c['nn']:
            continue
        cand = []
        for ai, a in enumerate(alumnos):
            lvl = core.name_match_strict(c['nn'], a['nn'])
            if lvl in ('exacta', 'fuerte'):
                dif = None
                if c['edad'] is not None and isinstance(a['edad'], int):
                    dif = abs(c['edad'] - a['edad'])
                    if dif > 10 and lvl != 'exacta':
                        continue  # diferencia de edad incompatible -> no se asume misma persona
                cand.append((0 if lvl == 'exacta' else 1, -name_sim(c['nn'], a['nn']), ai, lvl, dif))
        if cand:
            cand.sort()
            out[ci] = (cand[0][2], cand[0][3], cand[0][4])
    return out


# --------------------------------------------------- homologacion de estatus
# Etapa del proceso (CRM)  ->  estatus homologado solicitado
ETAPA_MAP = {
    'nuevo lead': 'Nuevo',
    'contactado': 'Descubrimiento',
    'informacion enviada': 'Descubrimiento',
    'esperando respuesta': 'Seguimiento',
    'pendiente de decision': 'Por cerrar',
    'asistencia a clase muestra': 'Clase muestra realizada',
    'cerrado': 'Inscrito / Cerrado',
}

CM_TO_ESTATUS = {
    'Realizada - Inscrito': 'Inscrito / Cerrado',
    'Realizada - Sin cierre': 'Clase muestra realizada',
    'No show': 'Seguimiento',
    'Agendada pendiente': 'Clase muestra agendada',
    'Reprogramada': 'Reprogramación',
}
