# Proyecto Lista de Chequeo Etapa Productiva — SENA Centro ASTIN

Sistema de automatización en Python para el seguimiento semanal de los aprendices de
**Etapa Productiva en modalidad VIRTUAL** de la Coordinación de Etapa Productiva.

El script procesa la base de datos `BASE DE DATOS ETAPA PRODUCTIVA.csv`, escanea la
carpeta sincronizada de **OneDrive** en busca de los entregables digitales de cada
aprendiz y genera un **Dashboard HTML interactivo** (Lista de Chequeo) con semáforo
de cumplimiento.

---

## Estructura del proyecto

```
Proyecto Lista de Chequeo Etapa Productiva/
├── scripts/
│   └── 07_generar_dashboard_etapa_productiva.py   # Script principal (ejecutable e idempotente)
├── data/
│   └── onedrive_index.json                        # Caché del escaneo de OneDrive (auto-generada)
├── outputs/
│   └── resumen_ultima_generacion.json             # Trazabilidad del corte semanal (auto-generado)
├── dashboard_etapa_productiva_virtual.html        # SALIDA: Dashboard autónomo
├── requirements.txt
└── README.md
```

## Instalación

```powershell
python -m pip install pandas
```

## Ejecución

```powershell
cd "C:\Users\capatino\Documents\Proyecto Lista de Chequeo Etapa Productiva"
python scripts/07_generar_dashboard_etapa_productiva.py
```

Luego abra `dashboard_etapa_productiva_virtual.html` en cualquier navegador.

El dashboard tiene dos pestañas:
1. **👥 Aprendices** — Lista de chequeo por aprendiz con semáforo de cumplimiento.
2. **🧑🏫 Instructores EP** — Vista de seguimiento a los instructores de Etapa Productiva
   (líder, seguimiento 1·2·3, técnico y consultorio) con su carga asignada y el estado
   de los aprendices a su cargo.

### Comandos de prueba

```powershell
# Validar la lectura de aprendices virtuales (no genera HTML)
python scripts/07_generar_dashboard_etapa_productiva.py --solo-validar

# Corrida completa (usa caché de OneDrive si existe)
python scripts/07_generar_dashboard_etapa_productiva.py

# Forzar reescaneo completo de OneDrive
python scripts/07_generar_dashboard_etapa_productiva.py --rescan
```

## Reglas de negocio implementadas

| # | Regla | Implementación |
|---|-------|----------------|
| 1 | Exclusividad VIRTUAL | Filtro estricto `MODALIDAD == 'VIRTUAL'` (ignora presenciales y vacías) |
| 2 | Identificación del aprendiz | Clave única `(FICHA, NUMERO_DOCUMENTO)`; duplicados omitidos |
| 3 | Verificación en OneDrive | Escaneo recursivo de `.pdf/.docx/.doc` cuyo nombre o ruta contenga FICHA o CÉDULA |
| 4 | Matriz de estado (semáforo) | Ver tabla siguiente |

### Semáforo de cumplimiento (con fechas de vencimiento)

Fechas reglamentarias consideradas:

| Fecha | Fuente / Cálculo | Norma |
|-------|------------------|-------|
| Venc. Selección Alternativa EP | Columna del CSV; si falta: **Fin Lectiva + 6 meses** | Acuerdo 9 de 2024 (Reglamento del Aprendiz) |
| VENCIMIENTO (EP) | Columna del CSV; si falta: **Terminación Ficha + 7 meses** | Circular 120 — Reglamento del Aprendiz |

| Estado | Color | Condición |
|--------|-------|-----------|
| AL DÍA | 🟢 `#39A900` | Alternativa + ARL + seguimientos 1, 2 y 3 realizados (GFPI-F-023 V6) |
| EN PROCESO | 🟡 Amarillo | Alternativa registrada con seguimientos pendientes, **o** sin alternativa pero con plazo de selección aún vigente |
| ALERTA | 🔴 Rojo | Sin alternativa con plazo de selección vencido, sin ARL, sin documentos, o EP vencida sin cierre |

### Fuente de datos

El script usa como **fuente primaria** el Excel oficial de SharePoint
(`BASE DE DATOS ETAPA PRODUCTIVA.xlsx`, carpeta "coordinacion" del sitio CentroAstin),
leído desde su copia local sincronizada por OneDrive. Si el Excel no existe o está
bloqueado, usa la copia temporal del archivo; el CSV queda como respaldo automático.

## Seguimientos — formato GFPI-F-023 V6

Un seguimiento (1, 2 o 3) se considera **realizado** únicamente cuando existe para el
aprendiz el formato **GFPI-F-023 V6** "Planeación, Seguimiento y Evaluación de Etapa
Productiva" diligenciado y firmado por las tres partes (aprendiz, instructor de
seguimiento y ente co-formador).

- **Detección por contenido**: el script lee el texto interno de TODOS los archivos
  .pdf/.docx/.doc (los nombres varían: S1_Junio.pdf, escaneos, etc.) y busca el texto
  `GFPI-F-023 V6`. Los resultados se cachean por mtime (corridas semanales rápidas).
- **Detección por nombre**: `S1_`, `S2_`, `S3_`, `GFPI-F-023`, `F-023`, `F023` +
  número de seguimiento en el nombre o en el contenido (Momentos 2/3/4 → Seg 1/2/3).
- **Asociación al aprendiz**: por cédula/ficha en el nombre, o por los identificadores
  leídos DENTRO del PDF (cédula y "No. Grupo").
- **Validación de diligenciamiento**: el formato se evalúa contra el ejemplo oficial:
  campos críticos (nombre, cédula, ficha, alternativa, fechas de EP, instructor,
  ente co-formador) deben estar diligenciados; campos no críticos (correos, fecha
  Sofia Plus, ARL, póliza, jefe inmediato) se reportan como observaciones; la firma
  digital del PDF se valida, y si no existe se deja observación para revisión manual.

Estados de la columna Seguimiento:

| Estado | Visual | Significado |
|--------|--------|-------------|
| Realizado | ✅ | GFPI-F-023 V6 completo (campos críticos diligenciados) |
| Incompleto | ⚠️ | GFPI existe pero faltan campos críticos (ver tooltip) |
| Asignado | ⚠️ | Instructor/fecha en CSV, sin formato GFPI |
| Pendiente | ❌ | Sin formato GFPI y sin asignación |

La asignación de instructor/fecha en el CSV **ya no** acredita el seguimiento.

Preaviso: las fechas a **30 días o menos** de vencer se marcan como ⚠️ "Próximo a vencer".

### Fin de Etapa Lectiva — atención inmediata

La fecha `FECHA FIN LECTIVA` se clasifica según su **año** frente al corte:

| Clasificación | Criterio | Visual |
|---------------|----------|--------|
| Atención inmediata | Año de fin lectiva **anterior** al año del corte | 🔴 |
| Año en curso | Fin lectiva en el año del corte | 🟡 |
| A futuro | Fin lectiva en años posteriores | 🟢 |

El dashboard incluye KPI propio, columna "Fin Lectiva", **filtro multi-año** (casillas de verificación para seleccionar varios años a la vez) y botón rápido **🔴 Atención**.

Fuentes combinadas: columnas del CSV **+** archivos encontrados en OneDrive.

### Categorías de entregables detectadas en OneDrive

- **🚀 Inicio**: registro de alternativa, carta de aceptación, afiliación ARL.
- **📒 Bitácoras y seguimientos**: bitácoras 1–12, formatos de seguimiento 1, 2 y 3.
- **🏁 Cierre y certificación**: acta de cierre lectiva/productiva, evaluación RAP, paz y salvo.

## Configuración

Las rutas se ajustan en la sección `CONFIG` al inicio del script:

- `csv_path`: ruta del CSV de Etapa Productiva.
- `onedrive_root`: raíz del escaneo de entregables.
- `extensiones`: formatos permitidos (`.pdf`, `.docx`, `.doc`).
- `excluir_dirnames`: carpetas a omitir del escaneo.

## Notas técnicas

- Codificación: el CSV se lee con `encoding='utf-8-sig'` (preserva tildes/ñ); si falla,
  se reintenta con `latin-1`.
- Manejo defensivo de nulos: todos los valores `NaN`/vacíos se normalizan a `''`.
- Idempotencia: cada ejecución regenera el HTML completo desde cero (corte semanal).
- El dashboard es **autónomo**: un único archivo HTML con datos embebidos en JSON,
  sin dependencias externas ni conexión a internet.
