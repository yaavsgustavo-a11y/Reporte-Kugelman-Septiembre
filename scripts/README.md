# Scripts del corte comercial

Reproducen el archivo `Kugelman_Academy_Corte_Comercial_2026-10-03.xlsx` a partir de las
tres fuentes del repositorio. **No modifican las fuentes originales.**

| Archivo | Función |
|---|---|
| `xlsx_read.py` | Lector de `.xlsx` (sólo biblioteca estándar de Python) |
| `xlsx_write.py` | Generador de `.xlsx` con estilos, filtros y tablas estructuradas |
| `kugelman_data.py` | Carga y normalización de las tres fuentes + algoritmo de cruce de nombres |
| `build_reporte.py` | Reglas de clasificación, métricas y armado de las cinco pestañas |

## Uso

```bash
cd scripts
python3 build_reporte.py
```

Requiere Python 3.6+ sin dependencias externas. El archivo se escribe en la raíz del repositorio.

## Decisiones que están codificadas en el script

- **Fecha de corte:** `CORTE = 2026-10-03` en `build_reporte.py`. Cambiarla recalcula todo el tablero.
- **Cruces validados a mano:** diccionarios `CONFIRMA` (coincidencias aceptadas), `RECHAZA`
  (coincidencias descartadas) y `VALIDAR_EXTRA` (casos irresolubles) al inicio de `build_reporte.py`.
  Cada entrada incluye el motivo y aparece en la pestaña *REGISTROS POR VALIDAR*.
- **Normalización de nombres:** `kugelman_data.py` quita acentos, mayúsculas y espacios dobles,
  ignora partículas (`de`, `la`, `del`…) y resuelve abreviaturas (`Hdz` → Hernández) y variantes
  de captura (`Pedrosa` → Pedroza, `Cordova` → Córdoba).
- **Herencia de resultados:** los resultados de `CAMPAÑAS` sólo se heredan cuando coincide el
  nombre. El teléfono se usa únicamente para la fecha de primer contacto, porque las familias
  comparten número.
