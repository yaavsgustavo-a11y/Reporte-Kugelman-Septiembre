import glob, os
REPO = '/projects/sandbox/Reporte-Kugelman-Septiembre'
OUT = '/projects/sandbox/entregables'

def find(frag):
    """Resolve a filename by substring, tolerant of unicode normalization."""
    frag = frag.lower()
    for f in os.listdir(REPO):
        if frag in f.lower():
            return os.path.join(REPO, f)
    raise FileNotFoundError(frag)

CAMPANAS = find('(2).xlsx')
CLASES = find('clases muestra')
RGA = find('rga_')
ADS = find('anuncios-21-ago')
