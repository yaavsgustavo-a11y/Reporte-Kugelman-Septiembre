# Campaña Septiembre 2026 — Kugelman Academy

**Periodo de campaña:** 21 de agosto – 4 de octubre de 2026 · **Presupuesto asignado:** $7,000 MXN
**Corte de bases operativas:** sábado 3 de octubre de 2026

Todos los archivos son **editables** (OOXML nativo): tablas editables y gráficas nativas ligadas a su
tabla de datos (clic derecho › *Editar datos*). No se entrega ningún PDF, imagen ni tabla renderizada
en sustitución de un archivo editable.

---

## Entregables

| # | Archivo | Formato | Contenido |
|---|---|---|---|
| 1 | `Reporte Campaña Septiembre 2026 — Kugelman Academy.docx` | Word | 17 secciones + anexo · 60 tablas editables · 8 gráficas nativas. Dashboard, funnel, pauta, creativos, inscripciones, hallazgos, oportunidades, recomendaciones priorizadas, calidad de datos y conclusión ejecutiva. |
| 2 | `Base Maestra Campaña Septiembre 2026.xlsx` | Excel | 14 pestañas · 107 fórmulas vivas · 12 gráficas nativas. Conserva los registros individuales: 482 prospectos-persona, los 504 registros originales del CRM, 107 clases muestra y 32 eventos de inscripción con su evidencia. |
| 3 | `Presentación Ejecutiva — Campaña Septiembre 2026.pptx` | PowerPoint | 10 diapositivas 16:9 para Dirección · 8 gráficas nativas con libro de datos embebido · 3 tablas editables. |
| 4 | `Base de Alumnos Inscritos — Corte 03 octubre 2026.xlsx` | Excel | 109 alumnos con **exactamente** las 6 columnas solicitadas (Matrícula, Nombre, Apellidos, Sexo, Edad, Activo), en formato de tabla con filtros, más una pestaña de validaciones. |
| 5 | `Prospectos en Curso — Corte 03 octubre 2026.xlsx` | Excel | 505 prospectos con estatus homologado · pestaña Resumen con distribución por estatus, funnel y 3 gráficas editables · criterios documentados · base de julio como referencia. |

---

## Cifras principales (consistentes en los archivos 1, 2 y 3)

| Indicador | Valor |
|---|---:|
| Presupuesto asignado | $7,000.00 |
| Gasto real (Meta Ads) | $6,999.91 (100.0%) |
| Impresiones | 241,265 |
| Alcance | 75,110 – 97,823 (Meta no exportó alcance único) |
| Conversaciones iniciadas | 565 |
| Leads / prospectos (personas únicas) | 477 |
| Clases muestra agendadas | 91 |
| Clases muestra realizadas | 67 |
| Inscripciones confirmadas en el periodo | 30 |
| Inscripciones atribuibles a la campaña | 23 |
| Lead → CM agendada | 19.1% |
| CM agendada → realizada | 73.6% |
| CM realizada → inscripción | 46.2% |
| Costo por lead | $14.67 |
| Costo por CM realizada | $104.48 |
| CAC (atribuibles) | $304.34 |

**Conclusión central:** el cuello de botella no está en marketing ni en el cierre, sino en el
**agendamiento**. 390 de 477 personas interesadas (81.8%) nunca llegaron a una clase muestra,
mientras que la asistencia (73.6%) y el cierre tras la clase (46.2%) están en niveles sanos.

---

## Fuentes utilizadas (las 4 del repositorio)

| Fuente | Alimenta |
|---|---|
| `Kugelman-Academy-Anuncios-21-ago-2026---4-oct-2026.csv` | Gasto real, impresiones, alcance, conversaciones, CPM, costo por conversación y clasificaciones por anuncio. |
| `CAMPAÑAS (2).xlsx` → hoja `AGOSTOSEPTIEMBRE` | Leads de la campaña (504 registros). Las hojas `Octubre` (5 registros duplicados), `OCTUBRE 1` (vacía) y `JULIO` (otra campaña) se excluyen de los KPI y se documentan. |
| `CLASES MUESTRA 2026__ (1).xlsx` → hojas `AGOSTO`, `SEPTIEMBRE`, `OCTUBRE` | Clases muestra, asistencia e inscripciones. |
| `RGA_1791241975.xlsx` → `General de alumnos` | Padrón de 109 alumnos; valida inscripciones y alimenta el entregable 4. |

### Información que NO existe en el repositorio (declarada, no estimada)

- Métricas orgánicas de Facebook, Instagram y TikTok → **la sección de redes sociales orgánicas no es calculable**.
- Clics de anuncios → **CTR y CPC no son calculables**.
- Desglose diario del gasto · alcance único de campaña (solo rango).
- Hora del lead y del primer contacto de vuelta → **el tiempo de respuesta no es medible**.
- Campo de clase muestra «solicitada» distinto de «agendada» · motivo de pérdida.
- Fecha de inscripción en el padrón → **el padrón no puede filtrarse al corte del 3 de octubre**.

---

## Trazabilidad

La carpeta `metodologia/` contiene el código que produce los cinco archivos a partir de las cuatro
fuentes. Reproduce **todas** las cifras: ninguna está escrita a mano. `facts.py` es la fuente única
de verdad que alimenta los cinco entregables, lo que garantiza que los indicadores sean idénticos
entre el Reporte, la Base Maestra y la Presentación.

```
python3 gen_reporte.py && python3 gen_base.py && python3 gen_ppt.py && python3 gen_bases.py
python3 final_check.py      # 31 comprobaciones de la validación final (punto 29 del brief)
python3 validate.py *.docx *.pptx *.xlsx   # integridad OOXML
```

Resultado de la validación final: **31 comprobaciones correctas, 0 fallas**.
