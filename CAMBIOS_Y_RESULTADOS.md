# Cambios y Resultados — Dashboard Etapa Productiva Virtual (SENA Centro ASTIN)

> Documento de trazabilidad de las modificaciones realizadas sobre el proyecto
> `Dashboard-de-Etapa-Productiva-Virtual---Centro-ASTIN`.
> Corte de datos: **25/11/2026** · Fuente primaria: **Excel de SharePoint/OneDrive**.

---

## 1. Resumen ejecutivo

El sistema genera semanalmente un dashboard HTML (lista de chequeo) para el
seguimiento de los **1.040 aprendices virtuales** (35 fichas) en Etapa Productiva
del Centro ASTIN. Desde la versión inicial (commit `fc5c594`) se incorporaron:

| # | Cambio | Motivo |
|---|--------|--------|
| 1 | Fechas de vencimiento reglamentarias en el semáforo | Acuerdo 9 de 2024 y Circular 120 |
| 2 | Vista de instructores de Etapa Productiva | Seguimiento a la carga de instructores |
| 3 | Clasificación de Fin de Etapa Lectiva por año | Atención inmediata a casos críticos |
| 4 | Filtro multi-año | Selección de 2 o más años a la vez |
| 5 | Seguimientos acreditados solo con formato GFPI-F-023 V6 | Regla de negocio: formato firmado por las 3 partes |
| 6 | Lectura por contenido de archivos ("GFPI-F-023 V6") | Los archivos tienen nombres variados |
| 7 | Evaluación de diligenciamiento del GFPI | Validar formato completo vs. incompleto |
| 8 | Excel de SharePoint/OneDrive como fuente primaria | Archivo oficial de la coordinación |

---

## 2. Detalle de los cambios

### 2.1 Fechas de vencimiento en el semáforo

Nuevas columnas consideradas del CSV/Excel:

| Fecha | Cálculo si falta en la base | Norma |
|-------|-----------------------------|-------|
| Venc. Selección Alternativa EP | Fin Lectiva **+ 6 meses** | Acuerdo 9 de 2024 |
| VENCIMIENTO (EP) | Terminación Ficha **+ 7 meses** | Circular 120 |

Matriz de estado resultante (7 reglas en cascada):

- 🔴 **ALERTA**: EP vencida sin cierre · sin alternativa con plazo de selección
  vencido · sin ARL · alternativa+ARL sin documentos.
- 🟡 **EN PROCESO**: sin alternativa con plazo de selección vigente · alternativa
  registrada con seguimientos pendientes.
- 🟢 **AL DÍA**: alternativa + ARL + seguimientos 1·2·3 realizados.

Preaviso de **30 días** para vencimientos próximos (⚠️).

### 2.2 Vista de instructores EP

Nueva pestaña del dashboard que agrega los 6 roles del CSV (líder, seguimiento
1·2·3, técnico, consultorio) por instructor, con fichas asignadas, aprendices a
cargo, su estado (semáforo) y barra de avance % al día. Filtros por nombre/ficha
y por rol.

### 2.3 Fin de Etapa Lectiva por año

Clasificación según el año de `FECHA FIN LECTIVA` frente al año del corte:

| Clasificación | Criterio | Visual |
|---------------|----------|--------|
| Atención inmediata | Año anterior al corte | 🔴 |
| Año en curso | Año del corte | 🟡 |
| A futuro | Años posteriores | 🟢 |

KPI propio + columna en la tabla + botón rápido "🔴 Atención" + filtro multi-año
(casillas de verificación para seleccionar varios años a la vez).

### 2.4 Seguimientos: formato GFPI-F-023 V6

**Regla de negocio**: un seguimiento (1, 2 o 3) solo se considera *realizado*
cuando existe el formato **GFPI-F-023 V6** (Planeación, Seguimiento y Evaluación
de Etapa Productiva) diligenciado y firmado por las tres partes (aprendiz,
instructor de seguimiento y ente co-formador). La asignación de instructor/fecha
en la base de datos **ya no** acredita el seguimiento.

Implementación:

1. **Detección por nombre**: `S1_`, `S2_`, `S3_`, `GFPI-F-023`, `F-023`, `F023`.
2. **Detección por contenido**: se lee el texto interno de **todos** los archivos
   .pdf/.docx/.doc (PyMuPDF y python-docx) buscando `GFPI-F-023 V6`. Resultados
   cacheados por mtime → corridas semanales rápidas.
3. **Asociación al aprendiz**: por cédula/ficha en el nombre, o por los
   identificadores leídos dentro del documento (cédula y "No. Grupo").
4. **Evaluación de diligenciamiento** contra el formato oficial:
   - **Críticos** (bloquean): nombre, cédula, ficha, alternativa, fechas inicio/fin
     de EP, instructor de seguimiento, ente co-formador.
   - **No críticos** (observaciones): correos del aprendiz, fecha Sofia Plus,
     fecha ARL, póliza ARL, jefe inmediato.
   - **Firma digital**: se valida con PyMuPDF (`sigflags`); si no existe, queda
     observación de revisión manual.

Estados resultantes de la columna Seguimiento:

| Estado | Visual | Significado |
|--------|--------|-------------|
| Realizado | ✅ | GFPI-F-023 V6 completo |
| Incompleto | ⚠️ | GFPI existe pero faltan campos críticos (tooltip con detalle) |
| Asignado | ⚠️ | Instructor/fecha en base de datos, sin formato |
| Pendiente | ❌ | Sin formato y sin asignación |

### 2.5 Excel de SharePoint como fuente primaria

El script lee `BASE DE DATOS ETAPA PRODUCTIVA.xlsx` (copia local sincronizada del
sitio SharePoint CentroAstin/coordinacion). Manejo defensivo:

- Si el Excel está abierto/bloqueado → copia temporal y lectura de la copia.
- Si no existe → respaldo automático con el CSV.
- La fuente usada queda registrada en `outputs/resumen_ultima_generacion.json`.

Nota: el Excel (corte 25/11/2026) trae fechas en formato datetime
(`2025-08-23 00:00:00`) y no trae las columnas de vencimiento; el script las
parsea y calcula automáticamente.

---

## 3. Resultados obtenidos (corte 25/11/2026)

### 3.1 Fuente y volumen

| Métrica | Valor |
|---------|-------|
| Fuente | Excel (BASE DE DATOS ETAPA PRODUCTIVA.xlsx) |
| Filas totales | 3.539 |
| Aprendices VIRTUALES | **1.040** (35 fichas) |
| Archivos OneDrive revisados | 4.462 (.pdf/.docx/.doc) |
| GFPI-F-023 V6 detectados por contenido | **0** |

### 3.2 Semáforo

| Estado | Aprendices | % |
|--------|-----------:|---:|
| 🟢 AL DÍA | 0 | 0.0 |
| 🟡 EN PROCESO | 814 | 78.3 |
| 🔴 ALERTA | 226 | 21.7 |

El 100% de expedientes pendientes es fiel a los datos: la columna ARL está vacía
en toda la base (0/3.539) y aún no hay formatos GFPI-F-023 V6 en OneDrive.

### 3.3 Seguimientos (con la nueva regla GFPI-F-023 V6)

| Seguimiento | ✅ Realizado | ⚠️ Incompleto | ⚠️ Asignado (BD) | ❌ Pendiente |
|------------|------------:|--------------:|-----------------:|------------:|
| 1.º | 0 | 0 | 271 | 769 |
| 2.º | 0 | 0 | 271 | 769 |
| 3.º | 0 | 0 | 271 | 769 |

Antes del cambio, la columna se basaba en instructor/fecha del CSV (26.1% marcado
como "realizado"). Con la regla GFPI-F-023 V6 el resultado es 0% realizado, que
refleja la ausencia real de los formatos en OneDrive.

### 3.4 Vencimientos

| Indicador | Aprendices |
|-----------|-----------:|
| Venc. selección alternativa vencido | 144 |
| Venc. selección próximo a vencer (≤30 días) | 24 |
| VENCIMIENTO EP vencido | 26 |
| Fin Lectiva año anterior (atención inmediata) | 26 (los 26 en ALERTA) |
| Fin Lectiva año en curso | 245 |
| Fin Lectiva a futuro | 769 |

### 3.5 Instructores EP

| Indicador | Valor |
|-----------|-------|
| Instructores con asignación | 11 |
| Asignaciones totales (6 roles) | 1.853 |
| Carga promedio | 120,3 aprendices/instructor |
| Instructores con ≥1 aprendiz en ALERTA | 10 |

### 3.6 Validación con el documento de ejemplo (`S1_Junio.pdf`)

El PDF de ejemplo entregado fue evaluado por la nueva lógica:

- ✅ Detectado como **GFPI-F-023 V6** por contenido (el nombre no contiene cédula).
- ✅ **Completo**: todos los campos críticos diligenciados.
- ✅ **Firmado digitalmente** (detectado por PyMuPDF).
- Identificadores extraídos: cédula `1012916660`, ficha `2834818`.
- Observaciones no críticas: falta correo institucional del aprendiz, fecha de
  registro Sofia Plus y fecha de afiliación ARL.

---

## 4. Rendimiento del escaneo de contenido

| Corrida | Tiempo |
|---------|-------:|
| Primera (lee el texto de 4.462 archivos) | ~36 min |
| Semanal con caché (mtime) | **~0.2 s** |

La caché de contenido se guarda en `data/gfpi_contenido_cache.json` (excluida del
repositorio). Para forzar relectura completa: `--rescan`.

---

## 5. Historial de versiones

| Commit | Descripción |
|--------|-------------|
| `fc5c594` | Dashboard base: semáforo, KPIs, filtros, escaneo OneDrive, vista instructores |
| *(siguiente)* | Vencimientos reglamentarios, GFPI-F-023 V6 por contenido, Excel SharePoint, multi-año, atención inmediata |

*Generado automáticamente como parte del ciclo de actualización semanal del proyecto.*
