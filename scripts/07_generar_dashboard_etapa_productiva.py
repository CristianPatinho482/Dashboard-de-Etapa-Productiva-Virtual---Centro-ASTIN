#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 07_generar_dashboard_etapa_productiva.py
 SENA - Centro ASTIN | Coordinación de Etapa Productiva
================================================================================
 Genera el Dashboard HTML (Lista de Chequeo) para el seguimiento semanal de los
 aprendices de MODALIDAD = 'VIRTUAL' en Etapa Productiva.

 FUENTES DE DATOS:
   1) BASE DE DATOS ETAPA PRODUCTIVA.csv  (corte semanal)
   2) OneDrive sincronizado localmente    (entregables digitales .pdf/.docx/.doc)

 SALIDA:
   dashboard_etapa_productiva_virtual.html  (archivo único y autónomo)

 USO:
   python scripts/07_generar_dashboard_etapa_productiva.py                 # corrida completa
   python scripts/07_generar_dashboard_etapa_productiva.py --solo-validar  # valida lectura CSV
   python scripts/07_generar_dashboard_etapa_productiva.py --rescan        # reescanea OneDrive desde cero
   python scripts/07_generar_dashboard_etapa_productiva.py --no-scan       # no escanea OneDrive

 El script es IDEMPOTENTE: cada ejecución regenera el HTML completo desde cero
 reflejando el corte semanal vigente.
================================================================================
"""

import argparse
import json
import re
import sys
import time
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd

# Salida de consola en UTF-8 explícito (evita caracteres corruptos en Windows)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ----------------------------------------------------------------------------
# 1. CONFIGURACIÓN CENTRAL (ajustable sin tocar la lógica)
# ----------------------------------------------------------------------------
CONFIG = {
    "csv_path": Path(
        r"C:\Users\capatino\OneDrive - Servicio Nacional de Aprendizaje\BACKUP DISCO"
        r"\Mis Bases de Datos\AREA LABORAL\SENA\Contrato 2026 - SENA"
        r"\BASE DE DATOS ETAPA PRODUCTIVA.csv"
    ),
    # Raíz del escaneo de entregables en OneDrive (se recorre recursivamente)
    "onedrive_root": Path(r"C:\Users\capatino\OneDrive - Servicio Nacional de Aprendizaje"),
    "extensiones": {".pdf", ".docx", ".doc"},
    "cache_index": None,  # se resuelve en main() -> data/onedrive_index.json
    "output_html": None,  # se resuelve en main() -> dashboard_etapa_productiva_virtual.html
    "resumen_dir": None,  # se resuelve en main() -> outputs/
    "excluir_dirnames": {
        ".git", "__pycache__", "tmp_chrome_profile", "tmp_validation",
        "Shortcuts", "Aplicaciones",
    },
}

# Paleta institucional SENA
PALETA = {"verde": "#39A900", "navy": "#00324D", "fondo": "#F8F9FA"}

# ----------------------------------------------------------------------------
# 2. UTILIDADES DE TEXTO Y DEFENSA DE NULOS
# ----------------------------------------------------------------------------

def normalizar(s):
    """Normaliza texto para comparación: minúsculas, sin tildes, sin espacios extremos."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.replace("_", " ")  # guiones bajos como separadores (FECHA_TERMINACION_FICHA)
    s = re.sub(r"\s+", " ", s)  # colapsa espacios dobles
    return s


def solo_digitos(s):
    """Extrae únicamente los dígitos de una cadena (para números de documento)."""
    return re.sub(r"\D", "", "" if s is None else str(s))


def parsear_fecha(s):
    """Convierte texto de fecha (DD/MM/AAAA u otros) a datetime.date. None si es inválida."""
    s = "" if s is None else str(s).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def sumar_meses(d, meses):
    """Suma meses a una fecha respetando fin de mes (regla del calendario civil)."""
    import calendar
    total = d.month - 1 + meses
    y, m = d.year + total // 12, total % 12 + 1
    dia = min(d.day, calendar.monthrange(y, m)[1])
    return d.replace(year=y, month=m, day=dia)


def celda(df_col, valor):
    """Devuelve texto seguro de una celda pandas: convierte NaN/None a ''."""
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return str(valor).strip()


def texto(row, col_name):
    """Acceso defensivo a columna: si no existe o es nula devuelve ''."""
    if col_name not in row.index:
        return ""
    return celda(None, row.get(col_name, ""))


# ----------------------------------------------------------------------------
# 3. LECTURA Y FILTRADO DEL CSV
# ----------------------------------------------------------------------------

def mapear_columnas(df):
    """Crea diccionario nombre_normalizado -> nombre_real de columna."""
    return {normalizar(c): c for c in df.columns}


def leer_base_datos(csv_path):
    """Lee el CSV con codificación UTF-8-sig y separador ; de forma defensiva."""
    if not csv_path.exists():
        sys.exit(f"[ERROR] No se encontró el archivo CSV:\n  {csv_path}")
    try:
        df = pd.read_csv(csv_path, encoding="utf-8-sig", sep=";", dtype=str)
    except UnicodeDecodeError:
        df = pd.read_csv(csv_path, encoding="latin-1", sep=";", dtype=str)
    df = df.fillna("")
    return df


def filtrar_virtuales(df):
    """Regla de negocio 1: exclusividad VIRTUAL (ignora presenciales y vacías)."""
    col_mod = "MODALIDAD"
    col_mod_real = None
    for c in df.columns:
        if normalizar(c) == "modalidad":
            col_mod_real = c
            break
    if col_mod_real is None:
        sys.exit("[ERROR] El CSV no tiene la columna 'MODALIDAD'.")
    mask = df[col_mod_real].map(lambda v: normalizar(v) == "virtual")
    return df[mask].copy()


# ----------------------------------------------------------------------------
# 4. ESCANEO DE ONEDRIVE (entregables digitales)
# ----------------------------------------------------------------------------

RE_SEG_NUM = [
    # (número, regex) para seguimientos 1, 2 y 3 en nombres de archivo
    (1, re.compile(r"(?:1er|1ro|1°|primero?)\s*seguimiento|seguimiento\s*(?:n[o°]?\.?\s*)?1\b", re.I)),
    (2, re.compile(r"(?:2do|2°|segundo)\s*seguimiento|seguimiento\s*(?:n[o°]?\.?\s*)?2\b", re.I)),
    (3, re.compile(r"(?:3er|3ro|3°|tercero?)\s*seguimiento|seguimiento\s*(?:n[o°]?\.?\s*)?3\b", re.I)),
]
RE_BITACORA_NUM = re.compile(r"bit[aá]cora\s*(?:n[o°]?\.?\s*)?(\d{1,2})", re.I)

CATEGORIAS = {
    # categoría -> regex de palabras clave (sin tildes, en minúscula)
    "inicio": re.compile(
        r"\b(alternativa|aceptaci[oó]n|arl|afiliaci[oó]n|registro\s*de\s*alternativa)\b"
    ),
    "seguimiento": re.compile(r"\b(seguimientos?|bit[aá]coras?)\b"),
    "cierre": re.compile(
        r"\b(cierre|rap\b|paz\s*y\s*salvo|pazysalvo|certificado|certificaci[oó]n|"
        r"evaluaci[oó]n\s*rap)\b"
    ),
}


def clasificar_archivo(nombre):
    """
    Clasifica un archivo según las 3 categorías de entregables y detecta
    número de seguimiento / bitácora. Devuelve dict con flags.
    """
    n = normalizar(nombre)
    r = {"inicio": False, "seguimiento": False, "cierre": False,
         "seg_nums": set(), "bitacora_nums": set(), "arl": False}
    if CATEGORIAS["inicio"].search(n):
        r["inicio"] = True
        if re.search(r"\barl\b|\bafiliaci[oó]n\b", n):
            r["arl"] = True
    if CATEGORIAS["cierre"].search(n):
        r["cierre"] = True
    if CATEGORIAS["seguimiento"].search(n):
        r["seguimiento"] = True
    for num, rx in RE_SEG_NUM:
        if rx.search(n):
            r["seg_nums"].add(num)
    for m in RE_BITACORA_NUM.finditer(n):
        b = int(m.group(1))
        if 1 <= b <= 12:
            r["bitacora_nums"].add(b)
    if "bitacora" in n and not r["bitacora_nums"]:
        r["bitacora_nums"].add(0)  # bitácora sin número identificable
    return r


def escanear_onedrive(root, extensiones, excluir_dirnames, cache_file, forzar=False):
    """
    Recorre recursivamente OneDrive en busca de .pdf/.docx/.doc.
    Usa caché JSON (ruta -> nombre) para que las corridas semanales sean rápidas.
    Devuelve lista de rutas (Path) absolutas.
    """
    if cache_file and cache_file.exists() and not forzar:
        try:
            data = json.loads(cache_file.read_text(encoding="utf-8"))
            files = [Path(p) for p in data.get("files", []) if Path(p).exists()]
            if files:
                print(f"  [OneDrive] Caché cargada: {len(files)} archivos indexados "
                      f"(use --rescan para reescaneo completo).")
                return files
        except (json.JSONDecodeError, OSError):
            pass

    print("  [OneDrive] Escaneando entregables (.pdf/.docx/.doc)... esto puede tomar unos minutos.")
    t0 = time.time()
    files = []
    if not root.exists():
        print(f"  [WARN] La raíz de OneDrive no existe: {root}")
        return files
    for p in root.rglob("*"):
        try:
            if not p.is_file():
                continue
            if p.suffix.lower() not in extensiones:
                continue
            if any(part in excluir_dirnames for part in p.parts):
                continue
            files.append(p)
        except OSError:
            continue
    files.sort(key=lambda x: str(x).lower())
    if cache_file:
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        cache_file.write_text(
            json.dumps({"files": [str(f) for f in files]}, ensure_ascii=False),
            encoding="utf-8",
        )
    print(f"  [OneDrive] Escaneo completo: {len(files)} archivos en {time.time() - t0:.1f}s.")
    return files


def construir_indice(files):
    """
    Índice: token numérico (6+ dígitos) -> set de archivos.
    Cubre cédulas con puntos (1.001.244.069 -> 1001244069) y fichas.
    """
    idx = defaultdict(set)
    for f in files:
        nombre = f.name
        ruta_padre = f.parent.name
        texto = f"{nombre} {ruta_padre}"
        digitos = re.sub(r"\D", "", texto)
        for token in set(re.findall(r"\d{6,}", digitos)):
            idx[token].add(f)
    return idx


def buscar_archivos_aprendiz(idx, files, ficha, documento):
    """Regla de negocio 3: archivos cuyo nombre/ruta contenga FICHA o DOCUMENTO."""
    fichas_d = solo_digitos(ficha)
    docs_d = solo_digitos(documento)
    encontrados = set()
    if len(fichas_d) >= 6:
        encontrados |= idx.get(fichas_d, set())
    if len(docs_d) >= 6:
        encontrados |= idx.get(docs_d, set())
    # Fallback para identificadores cortos o con formatos raros
    cortos = []
    if len(fichas_d) < 6 and fichas_d:
        cortos.append(fichas_d)
    if len(docs_d) < 6 and docs_d:
        cortos.append(docs_d)
    if cortos:
        for f in files:
            n = normalizar(f.name)
            if any(c in n for c in cortos):
                encontrados.add(f)
    return sorted(encontrados, key=lambda x: str(x).lower())


# ----------------------------------------------------------------------------
# 5. MATRIZ DE ESTADO (SEMÁFORO DE CUMPLIMIENTO) — con fechas de vencimiento
# ----------------------------------------------------------------------------
# Fechas reglamentarias (Acuerdo 9 de 2024 / Circular 120 Reglamento del Aprendiz):
#   VENC_SEL_ALT  : fecha límite para SELECCIONAR la alternativa de EP
#                   = FECHA FIN LECTIVA + 6 meses (si el CSV no la trae, se calcula).
#   VENCIMIENTO   : vencimiento de la etapa productiva
#                   = FECHA_TERMINACION_FICHA + 7 meses (si el CSV no lo trae, se calcula).
#
# Semáforo:
#   AL DÍA      : alternativa + ARL + seguimientos 1,2,3 cargados y sin vencimientos críticos.
#   EN PROCESO  : alternativa registrada (o plazo de selección aún vigente) pero con
#                 seguimientos/bitácoras pendientes.
#   ALERTA      : sin alternativa habiendo vencido el plazo de selección, sin ARL,
#                 sin documentos, o etapa productiva vencida sin cierre.
# ----------------------------------------------------------------------------

DIAS_PREAVISO = 30  # ⚠️ "Próximo a vencer" si faltan 30 días o menos


def calcular_estado(alternativa_ok, arl_ok, seg_ok, tiene_docs, tiene_seg_csv,
                    venc_sel_alt, vencimiento, cierre_ok, hoy):
    """
    venc_sel_alt / vencimiento: datetime.date o None.
    """
    vencida_sel = venc_sel_alt is not None and hoy > venc_sel_alt
    vencida_ep = vencimiento is not None and hoy > vencimiento

    # 1) Etapa productiva vencida y sin cierre/certificación → ALERTA crítico
    if vencida_ep and not cierre_ok:
        return "ALERTA"
    # 2) Sin alternativa y ya venció el plazo de selección → ALERTA
    if not alternativa_ok and vencida_sel:
        return "ALERTA"
    # 3) Sin alternativa pero el plazo de selección sigue vigente → EN PROCESO
    if not alternativa_ok and venc_sel_alt is not None and not vencida_sel:
        return "EN PROCESO"
    # 4) Sin ARL → ALERTA
    if not arl_ok:
        return "ALERTA"
    # 5) Alternativa + ARL pero sin documentos ni seguimientos en ninguna fuente → ALERTA
    if not tiene_docs and not tiene_seg_csv:
        return "ALERTA"
    # 6) Alternativa + ARL + seguimientos 1,2,3 completos → AL DÍA
    if seg_ok[1] and seg_ok[2] and seg_ok[3]:
        return "AL DÍA"
    # 7) Alternativa + ARL con seguimientos pendientes → EN PROCESO
    return "EN PROCESO"


def clasificar_vencimiento(fecha, hoy):
    """
    Estado de un vencimiento respecto a hoy:
      'vencido'  si la fecha ya pasó,
      'proximo'  si faltan <= DIAS_PREAVISO días,
      'vigente'  en otro caso,
      'sin_info' si no hay fecha.
    """
    if fecha is None:
        return "sin_info"
    dias = (fecha - hoy).days
    if dias < 0:
        return "vencido"
    if dias <= DIAS_PREAVISO:
        return "proximo"
    return "vigente"


def clasificar_fin_lectiva(fecha, hoy):
    """
    Clasifica la fecha de FINALIZACIÓN DE ETAPA LECTIVA según su año, para
    detectar situaciones que requieren atención inmediata:
      'atencion' : año ANTERIOR al año del corte → sus plazos de EP ya corren/vencieron.
      'vigente'  : año del corte (en curso).
      'futuro'   : años posteriores.
      'sin_info' : sin fecha en el CSV.
    """
    if fecha is None:
        return "sin_info"
    if fecha.year < hoy.year:
        return "atencion"
    if fecha.year > hoy.year:
        return "futuro"
    return "vigente"


# ----------------------------------------------------------------------------
# 6. CONSTRUCCIÓN DE EXPEDIENTES POR APRENDIZ
# ----------------------------------------------------------------------------

def construir_expedientes(df, files, idx):
    """Construye la lista de expedientes (clave única FICHA + NUMERO_DOCUMENTO)."""
    cmap = mapear_columnas(df)

    def col(nombre):
        """Busca columna por nombre normalizado, tolerando tildes/espacios."""
        for k, v in cmap.items():
            if normalizar(nombre) == k:
                return v
        return None

    c_ficha = col("ficha")
    c_doc = col("numero_documento")
    c_tipo = col("tipo_documento")
    c_nom = col("nombre")
    c_ape1 = col("primer_apellido")
    c_ape2 = col("segundo_apellido")
    c_prog = col("programa")
    c_mod = col("modalidad")
    c_alt = col("alternativa registrada")
    c_emp = col("empresa")
    c_arl = col("fecha de afiliacion arl (sena)")
    c_i1 = col("instructor de seguimiento 1")
    c_f1 = col("fecha de 1er seguimiento")
    c_i2 = col("instructor de seguimiento 2")
    c_f2 = col("fecha de 2do seguimiento")
    c_i3 = col("instructor de seguimiento 3")
    c_f3 = col("fecha de 3er seguimiento")
    c_cie = col("acta de cierre de etapa productiva")
    c_rap = col("rap productiva evaluado por")
    c_cert = col("certificado")
    # Fechas reglamentarias nuevas
    c_fin_lectiva = col("fecha fin lectiva")
    c_term_ficha = col("fecha terminacion ficha")
    c_venc_sel = col("vencimiento seleccion alternativa ep")
    c_vencimiento = col("vencimiento")

    hoy = datetime.now().date()

    expedientes, vistos, duplicados = [], set(), 0
    for _, row in df.iterrows():
        ficha = texto(row, c_ficha) if c_ficha else ""
        doc = texto(row, c_doc) if c_doc else ""
        clave = (ficha, doc)
        if clave in vistos:
            duplicados += 1
            continue
        vistos.add(clave)

        # --- Archivos del aprendiz en OneDrive ----------------------------
        archivos = buscar_archivos_aprendiz(idx, files, ficha, doc)
        clasificados = [(f, clasificar_archivo(f.name)) for f in archivos]

        docs_inicio = [f for f, c in clasificados if c["inicio"]]
        docs_alt = [f for f, c in clasificados if c["inicio"] and not c["arl"]]
        docs_arl = [f for f, c in clasificados if c["arl"]]
        docs_bit = [f for f, c in clasificados if c["bitacora_nums"]]
        docs_seg = {}
        for n in (1, 2, 3):
            docs_seg[n] = [f for f, c in clasificados if n in c["seg_nums"]]
        docs_cierre = [f for f, c in clasificados if c["cierre"]]

        # --- Fuentes CSV ---------------------------------------------------
        alt_csv = texto(row, c_alt) if c_alt else ""
        arl_csv = texto(row, c_arl) if c_arl else ""
        seg_csv = {}
        seg_csv[1] = bool(texto(row, c_i1) or texto(row, c_f1)) if (c_i1 or c_f1) else False
        seg_csv[2] = bool(texto(row, c_i2) or texto(row, c_f2)) if (c_i2 or c_f2) else False
        seg_csv[3] = bool(texto(row, c_i3) or texto(row, c_f3)) if (c_i3 or c_f3) else False
        cierre_csv = bool(texto(row, c_cie) or texto(row, c_rap) or texto(row, c_cert)) \
            if (c_cie or c_rap or c_cert) else False

        # --- Fechas reglamentarias (vencimientos) ---------------------------
        # 0) Fecha de finalización de la Etapa Lectiva (para atención inmediata)
        fin_lectiva = parsear_fecha(texto(row, c_fin_lectiva) if c_fin_lectiva else "")
        estado_fin_lectiva = clasificar_fin_lectiva(fin_lectiva, hoy)
        # 1) Vencimiento selección alternativa EP:
        #    usa el valor del CSV; si está vacío, calcula Fin Lectiva + 6 meses.
        venc_sel_alt = parsear_fecha(texto(row, c_venc_sel) if c_venc_sel else "")
        if venc_sel_alt is None and c_fin_lectiva:
            fin_lectiva = parsear_fecha(texto(row, c_fin_lectiva))
            if fin_lectiva:
                venc_sel_alt = sumar_meses(fin_lectiva, 6)
        # 2) VENCIMIENTO de la etapa productiva (Circular 120):
        #    usa el valor del CSV; si está vacío, calcula Terminación Ficha + 7 meses.
        vencimiento = parsear_fecha(texto(row, c_vencimiento) if c_vencimiento else "")
        if vencimiento is None and c_term_ficha:
            term_ficha = parsear_fecha(texto(row, c_term_ficha))
            if term_ficha:
                vencimiento = sumar_meses(term_ficha, 7)

        estado_venc_sel = clasificar_vencimiento(venc_sel_alt, hoy)
        estado_venc_ep = clasificar_vencimiento(vencimiento, hoy)

        # --- Consolidación (CSV + OneDrive combinados) ---------------------
        alternativa_ok = bool(alt_csv) or bool(docs_alt)
        arl_ok = bool(arl_csv) or bool(docs_arl)
        seg_ok = {n: bool(seg_csv[n] or docs_seg[n]) for n in (1, 2, 3)}
        tiene_docs = bool(archivos)
        tiene_seg_csv = any(seg_csv.values())
        cierre_ok = cierre_csv or bool(docs_cierre)

        estado = calcular_estado(
            alternativa_ok, arl_ok, seg_ok, tiene_docs, tiene_seg_csv,
            venc_sel_alt, vencimiento, cierre_ok, hoy,
        )

        def _ruta_uri(f):
            return str(f)

        expedientes.append({
            "ficha": ficha,
            "documento": doc,
            "tipo_doc": texto(row, c_tipo) if c_tipo else "",
            "nombre": " ".join(
                p for p in [
                    texto(row, c_nom) if c_nom else "",
                    texto(row, c_ape1) if c_ape1 else "",
                    texto(row, c_ape2) if c_ape2 else "",
                ] if p
            ).strip(),
            "programa": texto(row, c_prog) if c_prog else "",
            "modalidad": texto(row, c_mod) if c_mod else "VIRTUAL",
            "alternativa": alt_csv,
            "alternativa_ok": alternativa_ok,
            "empresa": texto(row, c_emp) if c_emp else "",
            "arl": arl_csv,
            "arl_ok": arl_ok,
            "seg": {n: {"ok": seg_ok[n], "csv": seg_csv[n]} for n in (1, 2, 3)},
            "bitacoras_ok": bool(docs_bit),
            "bitacoras_n": len(docs_bit),
            "cierre_ok": cierre_ok,
            "venc_sel_alt": venc_sel_alt.isoformat() if venc_sel_alt else "",
            "venc_sel_alt_estado": estado_venc_sel,
            "vencimiento": vencimiento.isoformat() if vencimiento else "",
            "vencimiento_estado": estado_venc_ep,
            "fin_lectiva": fin_lectiva.isoformat() if fin_lectiva else "",
            "fin_lectiva_estado": estado_fin_lectiva,
            "estado": estado,
            "archivos": {
                "inicio": [_ruta_uri(f) for f in docs_inicio],
                "arl": [_ruta_uri(f) for f in docs_arl],
                "seg1": [_ruta_uri(f) for f in docs_seg[1]],
                "seg2": [_ruta_uri(f) for f in docs_seg[2]],
                "seg3": [_ruta_uri(f) for f in docs_seg[3]],
                "bitacoras": [_ruta_uri(f) for f in docs_bit],
                "cierre": [_ruta_uri(f) for f in docs_cierre],
            },
            "n_archivos": len(archivos),
        })

    expedientes.sort(key=lambda e: (e["ficha"], e["nombre"].lower()))
    return expedientes, duplicados


# ----------------------------------------------------------------------------
# 6b. VISTA DE INSTRUCTORES DE ETAPA PRODUCTIVA
# ----------------------------------------------------------------------------

ROLES_INSTRUCTOR = [
    ("instructor lider de la ficha", "LÍDER"),
    ("instructor de seguimiento 1", "SEGUIMIENTO 1"),
    ("instructor de seguimiento 2", "SEGUIMIENTO 2"),
    ("instructor de seguimiento 3", "SEGUIMIENTO 3"),
    ("instructor tecnico (proyecto productivo)", "TÉCNICO"),
    ("instructor de consultorio de proyecto (proyecto productivo)", "CONSULTORIO"),
]


def construir_instructores(df_virtuales, expedientes):
    """
    Agrega por instructor todas sus asignaciones en los 6 roles de EP y el
    estado (semáforo) de los aprendices a su cargo.
    """
    cmap = mapear_columnas(df_virtuales)

    def col(nombre):
        for k, v in cmap.items():
            if normalizar(nombre) == k:
                return v
        return None

    cols = [(col(n), rol) for n, rol in ROLES_INSTRUCTOR]
    estado_por_clave = {(e["ficha"], e["documento"]): e["estado"] for e in expedientes}
    nombre_por_clave = {(e["ficha"], e["documento"]): e["nombre"] for e in expedientes}
    c_ficha = col("ficha")
    c_doc = col("numero documento")

    instructores = {}
    for _, row in df_virtuales.iterrows():
        ficha = texto(row, c_ficha) if c_ficha else ""
        doc = texto(row, c_doc) if c_doc else ""
        clave = (ficha, doc)
        for c, rol in cols:
            if not c:
                continue
            nombre = texto(row, c)
            if not nombre:
                continue
            key = normalizar(nombre)
            if key not in instructores:
                instructores[key] = {
                    "nombre": nombre, "roles": {}, "fichas": set(), "aprendices": {},
                }
            inst = instructores[key]
            inst["roles"][rol] = inst["roles"].get(rol, 0) + 1
            inst["fichas"].add(ficha)
            ap = inst["aprendices"].setdefault(clave, {
                "nombre": nombre_por_clave.get(clave, ""),
                "ficha": ficha,
                "documento": doc,
                "roles": set(),
                "estado": estado_por_clave.get(clave, "ALERTA"),
            })
            ap["roles"].add(rol)

    orden_estado = {"ALERTA": 0, "EN PROCESO": 1, "AL DÍA": 2}
    lista = []
    for inst in instructores.values():
        aprendices = []
        for ap in inst["aprendices"].values():
            aprendices.append({**ap, "roles": sorted(ap["roles"])})
        aprendices.sort(key=lambda a: (orden_estado.get(a["estado"], 3), a["nombre"].lower()))
        estados = {"AL DÍA": 0, "EN PROCESO": 0, "ALERTA": 0}
        for a in aprendices:
            estados[a["estado"]] = estados.get(a["estado"], 0) + 1
        lista.append({
            "nombre": inst["nombre"],
            "roles": inst["roles"],
            "fichas": sorted(inst["fichas"]),
            "n_fichas": len(inst["fichas"]),
            "n_aprendices": len(aprendices),
            "aprendices": aprendices,
            "estados": estados,
        })
    lista.sort(key=lambda i: (-i["n_aprendices"], i["nombre"].lower()))
    return lista


def calcular_kpis_instructores(instructores):
    total = len(instructores)
    asignaciones = sum(sum(i["roles"].values()) for i in instructores)
    n_aprendices_unicos = sum(i["n_aprendices"] for i in instructores)
    promedio = round(n_aprendices_unicos / total, 1) if total else 0.0
    con_alerta = sum(1 for i in instructores if i["estados"].get("ALERTA", 0) > 0)
    return {
        "total": total,
        "asignaciones": asignaciones,
        "promedio": promedio,
        "con_alerta": con_alerta,
    }


# ----------------------------------------------------------------------------
# 7. KPIs
# ----------------------------------------------------------------------------

def calcular_kpis(expedientes):
    total = len(expedientes) or 1
    n_alt = sum(1 for e in expedientes if e["alternativa_ok"])
    n_seg = {n: sum(1 for e in expedientes if e["seg"][n]["ok"]) for n in (1, 2, 3)}
    n_estado = defaultdict(int)
    for e in expedientes:
        n_estado[e["estado"]] += 1
    n_al_dia = n_estado.get("AL DÍA", 0)
    n_pend = total - n_al_dia

    def pct(x):
        return round(100.0 * x / total, 1)

    n_venc_sel = sum(1 for e in expedientes if e["venc_sel_alt_estado"] == "vencido")
    n_prox_sel = sum(1 for e in expedientes if e["venc_sel_alt_estado"] == "proximo")
    n_venc_ep = sum(1 for e in expedientes if e["vencimiento_estado"] == "vencido")
    n_prox_ep = sum(1 for e in expedientes if e["vencimiento_estado"] == "proximo")
    n_fin_aten = sum(1 for e in expedientes if e["fin_lectiva_estado"] == "atencion")
    n_fin_aten_alerta = sum(
        1 for e in expedientes
        if e["fin_lectiva_estado"] == "atencion" and e["estado"] == "ALERTA"
    )
    n_fin_vigente = sum(1 for e in expedientes if e["fin_lectiva_estado"] == "vigente")
    n_fin_futuro = sum(1 for e in expedientes if e["fin_lectiva_estado"] == "futuro")

    return {
        "total": len(expedientes),
        "total_fichas": len({e["ficha"] for e in expedientes}),
        "pct_alternativa": pct(n_alt),
        "pct_seg1": pct(n_seg[1]),
        "pct_seg2": pct(n_seg[2]),
        "pct_seg3": pct(n_seg[3]),
        "n_al_dia": n_al_dia,
        "n_en_proceso": n_estado.get("EN PROCESO", 0),
        "n_alerta": n_estado.get("ALERTA", 0),
        "pct_completos": pct(n_al_dia),
        "pct_pendientes": pct(n_pend),
        "n_venc_sel": n_venc_sel,
        "n_prox_sel": n_prox_sel,
        "n_venc_ep": n_venc_ep,
        "n_prox_ep": n_prox_ep,
        "n_fin_aten": n_fin_aten,
        "n_fin_aten_alerta": n_fin_aten_alerta,
        "n_fin_vigente": n_fin_vigente,
        "n_fin_futuro": n_fin_futuro,
    }


# ----------------------------------------------------------------------------
# 8. GENERACIÓN DEL DASHBOARD HTML (autónomo y estático)
# ----------------------------------------------------------------------------

PLANTILLA_HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Etapa Productiva Virtual — SENA Centro ASTIN</title>
<style>
  :root{
    --verde:#39A900; --navy:#00324D; --fondo:#F8F9FA; --blanco:#FFFFFF;
    --rojo:#D92525; --amarillo:#E8A800; --borde:#DEE2E6; --texto:#212529; --gris:#6C757D;
  }
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:var(--fondo);color:var(--texto);}
  header{background:var(--navy);color:#fff;padding:18px 28px;display:flex;flex-wrap:wrap;gap:12px;
         align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(0,50,77,.25);}
  header .logo{display:flex;align-items:center;gap:12px;}
  header .logo .badge{background:var(--verde);color:#fff;font-weight:800;font-size:13px;
         padding:4px 10px;border-radius:20px;letter-spacing:.5px;}
  header h1{font-size:20px;font-weight:700;}
  header .sub{font-size:12.5px;opacity:.85;margin-top:2px;}
  header .corte{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);
         padding:8px 14px;border-radius:10px;font-size:12.5px;text-align:right;line-height:1.5;}
  main{padding:20px 28px 40px;max-width:1500px;margin:0 auto;}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px;margin-bottom:18px;}
  .kpi{background:#fff;border-radius:12px;padding:16px 18px;box-shadow:0 1px 4px rgba(0,0,0,.08);
       border-left:6px solid var(--verde);}
  .kpi.navy{border-left-color:var(--navy);} .kpi.amber{border-left-color:var(--amarillo);}
  .kpi.red{border-left-color:var(--rojo);}
  .kpi .titulo{font-size:12px;font-weight:600;color:var(--gris);text-transform:uppercase;letter-spacing:.4px;}
  .kpi .valor{font-size:30px;font-weight:800;color:var(--navy);margin:4px 0 2px;}
  .kpi .mini{font-size:12.5px;color:var(--gris);display:flex;flex-wrap:wrap;gap:8px;margin-top:4px;}
  .kpi .mini b{color:var(--navy);}
  .fecha-cell{font-size:11.5px;line-height:1.6;}
  .fecha-cell .d{font-weight:700;color:var(--navy);}
  .tabs{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;}
  .tab-btn{border:2px solid var(--navy);background:#fff;color:var(--navy);font-weight:700;
          font-size:13.5px;padding:8px 20px;border-radius:22px;cursor:pointer;transition:all .15s;}
  .tab-btn:hover{filter:brightness(.96);}
  .tab-btn.activo{background:var(--navy);color:#fff;}
  .tab-btn .cnt{opacity:.8;font-weight:600;margin-left:4px;}
  .oculto{display:none !important;}
  .multi-dd{position:relative;}
  .dd-btn{font-size:13.5px;padding:8px 10px;border:1px solid var(--borde);border-radius:8px;
          background:#fff;color:var(--texto);min-width:175px;max-width:230px;cursor:pointer;
          text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
  .dd-btn::after{content:"▾";float:right;color:var(--gris);}
  .dd-panel{position:absolute;top:calc(100% + 4px);left:0;z-index:50;background:#fff;
          border:1px solid var(--borde);border-radius:8px;box-shadow:0 4px 14px rgba(0,0,0,.16);
          padding:8px;min-width:210px;max-height:270px;overflow:auto;}
  .dd-panel .dd-item{display:flex;flex-direction:row;align-items:center;gap:8px;font-size:13px;font-weight:400;
          color:var(--texto);padding:5px 7px;border-radius:6px;cursor:pointer;}
  .dd-panel .dd-item input[type="checkbox"]{order:-1;margin:0;width:15px;height:15px;cursor:pointer;}
  .dd-panel .dd-item:hover{background:#F1F7EC;}
  .dd-actions{display:flex;gap:6px;padding:2px 4px 8px;border-bottom:1px solid var(--borde);
          margin-bottom:6px;}
  .dd-actions button{font-size:11.5px;font-weight:600;border:1px solid var(--borde);background:#fff;
          border-radius:6px;padding:4px 9px;cursor:pointer;color:var(--navy);}
  .dd-actions button:hover{background:var(--fondo);}
  .filtros{background:#fff;border-radius:12px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,.08);
           display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:16px;}
  .filtros label{font-size:12px;font-weight:600;color:var(--gris);display:flex;flex-direction:column;gap:4px;}
  .filtros select,.filtros input{font-size:13.5px;padding:8px 10px;border:1px solid var(--borde);
           border-radius:8px;background:#fff;color:var(--texto);min-width:170px;}
  .filtros input#buscar{flex:1;min-width:240px;}
  .botones{display:flex;flex-wrap:wrap;gap:8px;margin-left:auto;}
  .btn-estado{border:none;border-radius:20px;padding:8px 14px;font-size:12.5px;font-weight:700;cursor:pointer;
           background:#E9ECEF;color:var(--navy);transition:all .15s;}
  .btn-estado:hover{filter:brightness(.95);}
  .btn-estado.activo-todos{background:var(--navy);color:#fff;}
  .btn-estado.activo-aldia{background:var(--verde);color:#fff;}
  .btn-estado.activo-proceso{background:var(--amarillo);color:#fff;}
  .btn-estado.activo-alerta{background:var(--rojo);color:#fff;}
  .tabla-wrap{background:#fff;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,.08);overflow:auto;}
  table{width:100%;border-collapse:collapse;font-size:13px;min-width:1450px;}
  thead th{background:var(--navy);color:#fff;padding:11px 10px;text-align:left;font-size:12px;
           position:sticky;top:0;z-index:2;white-space:nowrap;}
  tbody td{padding:9px 10px;border-bottom:1px solid var(--borde);vertical-align:top;}
  tbody tr:hover{background:#F1F7EC;}
  .ficha{font-weight:700;color:var(--navy);white-space:nowrap;}
  .aprendiz b{color:var(--navy);}
  .aprendiz .doc{color:var(--gris);font-size:11.5px;}
  .programa{font-size:12px;color:#343A40;}
  .chip{display:inline-block;font-size:11px;font-weight:700;padding:2px 8px;border-radius:10px;white-space:nowrap;}
  .ok{background:#E3F5D8;color:#1E7A00;}
  .no{background:#FDE4E4;color:var(--rojo);}
  .parcial{background:#FFF3CC;color:#8F6A00;}
  .seg-cell{display:inline-flex;align-items:center;gap:3px;margin:1px 2px;font-size:11.5px;font-weight:700;
         padding:2px 7px;border-radius:12px;}
  .seg-ok{background:#E3F5D8;color:#1E7A00;} .seg-no{background:#FDE4E4;color:var(--rojo);}
  .estado{display:inline-block;font-weight:800;font-size:12px;padding:5px 12px;border-radius:16px;
         color:#fff;white-space:nowrap;}
  .est-AL{background:var(--verde);} .est-EN{background:var(--amarillo);} .est-ALERTA{background:var(--rojo);}
  .soportes{font-size:11.5px;color:var(--gris);line-height:1.7;}
  .soportes a{color:var(--navy);text-decoration:none;border-bottom:1px dotted var(--navy);}
  .soportes a:hover{color:var(--verde);}
  details summary{cursor:pointer;font-weight:600;color:var(--navy);font-size:11.5px;}
  details ul{list-style:none;margin:4px 0 0 10px;max-width:360px;}
  details li{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:360px;}
  .contador{font-size:12.5px;color:var(--gris);margin:10px 2px 0;}
  .leyenda{display:flex;flex-wrap:wrap;gap:16px;font-size:12px;color:var(--gris);margin-top:14px;}
  .leyenda span{display:inline-flex;align-items:center;gap:5px;}
  footer{background:var(--navy);color:rgba(255,255,255,.85);font-size:12px;text-align:center;
         padding:12px;margin-top:26px;}
  .vaciar{color:var(--gris);font-style:italic;padding:18px;text-align:center;}
  @media (max-width:720px){
    header{padding:14px;} main{padding:12px 10px 30px;}
    .filtros{flex-direction:column;align-items:stretch;} .botones{margin-left:0;}
  }
</style>
</head>
<body>
<header>
  <div class="logo">
    <span class="badge">SENA · ASTIN</span>
    <div>
      <h1>Dashboard Etapa Productiva — Aprendices Virtuales</h1>
      <div class="sub">Coordinación de Etapa Productiva · Lista de Chequeo de Expedientes</div>
    </div>
  </div>
  <div class="corte">Corte semanal<br><b id="fecha-corte"></b></div>
</header>
<main>
  <section class="tabs">
    <button class="tab-btn activo" id="tab-aprendices" onclick="mostrarTab('aprendices')">👥 Aprendices <span class="cnt" id="cnt-aprendices"></span></button>
    <button class="tab-btn" id="tab-instructores" onclick="mostrarTab('instructores')">🧑‍🏫 Instructores EP <span class="cnt" id="cnt-instructores"></span></button>
  </section>

  <section id="vista-aprendices">
  <section class="kpis" id="kpis"></section>

  <section class="filtros">
    <label>Ficha
      <select id="sel-ficha"><option value="">Todas las fichas</option></select>
    </label>
    <label>Programa
      <select id="sel-programa"><option value="">Todos los programas</option></select>
    </label>
    <label>Año Fin Lectiva (multi)
      <div class="multi-dd" id="dd-anio">
        <button type="button" class="dd-btn" id="btn-anio">Todos los años</button>
        <div class="dd-panel oculto" id="panel-anio"></div>
      </div>
    </label>
    <label style="flex:1;">Buscar (cédula, nombre o empresa)
      <input type="search" id="buscar" placeholder="Escriba para filtrar al instante...">
    </label>
    <div class="botones" id="botones-estado">
      <button class="btn-estado activo-todos" data-estado="">Todos</button>
      <button class="btn-estado" data-estado="AL DÍA">Al Día</button>
      <button class="btn-estado" data-estado="EN PROCESO">En Proceso</button>
      <button class="btn-estado" data-estado="ALERTA">Alerta</button>
      <button class="btn-estado" data-aten="1" style="background:#FDE4E4;color:var(--rojo);">🔴 Atención</button>
    </div>
  </section>

  <div class="tabla-wrap">
    <table id="tabla">
      <thead>
        <tr>
          <th>Ficha</th>
          <th>Aprendiz</th>
          <th>Programa</th>
          <th>Alternativa</th>
          <th>Empresa</th>
          <th>ARL</th>
          <th>Seguimientos (1·2·3)</th>
          <th>Fin Lectiva</th>
          <th>Venc. Selección Alt.</th>
          <th>Vencimiento EP</th>
          <th>Soportes OneDrive</th>
          <th>Estado Final</th>
        </tr>
      </thead>
      <tbody id="tbody"></tbody>
    </table>
  </div>

  <div class="contador" id="contador"></div>
  <div class="leyenda">
    <span>✅ Registrado / Entregado</span>
    <span>⚠️ Parcial / Próximo a vencer</span>
    <span>❌ Pendiente / Vencido</span>
    <span>·</span>
    <span>🟢 AL DÍA: alternativa + ARL + seguimientos 1·2·3</span>
    <span>🟡 EN PROCESO: alternativa registrada o plazo de selección vigente, seguimientos pendientes</span>
    <span>🔴 ALERTA: sin alternativa con plazo vencido, sin ARL, sin documentos o EP vencida sin cierre</span>
    <span>·</span>
    <span>🗓️ Venc. Selección = Fin Lectiva + 6 meses · Vencimiento EP = Terminación Ficha + 7 meses</span>
    <span>·</span>
    <span>🎓 Fin Lectiva: 🔴 año anterior al corte (atención inmediata) · 🟡 año en curso · 🟢 años futuros</span>
  </div>
  </section>

  <section id="vista-instructores" class="oculto">
  <section class="kpis" id="kpis-instr"></section>

  <section class="filtros">
    <label style="flex:1;">Buscar instructor (nombre o ficha)
      <input type="search" id="buscar-instr" placeholder="Escriba para filtrar al instante...">
    </label>
    <label>Rol
      <select id="sel-rol">
        <option value="">Todos los roles</option>
        <option value="LÍDER">Líder de Ficha</option>
        <option value="SEGUIMIENTO 1">Seguimiento 1</option>
        <option value="SEGUIMIENTO 2">Seguimiento 2</option>
        <option value="SEGUIMIENTO 3">Seguimiento 3</option>
        <option value="TÉCNICO">Técnico (Proyecto Productivo)</option>
        <option value="CONSULTORIO">Consultorio de Proyecto</option>
      </select>
    </label>
  </section>

  <div class="tabla-wrap">
    <table id="tabla-instr" style="min-width:1100px;">
      <thead>
        <tr>
          <th>Instructor</th>
          <th>Roles</th>
          <th>Fichas</th>
          <th>Aprendices</th>
          <th>🟢 Al Día</th>
          <th>🟡 En Proceso</th>
          <th>🔴 Alerta</th>
          <th>Carga</th>
        </tr>
      </thead>
      <tbody id="tbody-instr"></tbody>
    </table>
  </div>

  <div class="contador" id="contador-instr"></div>
  </section>
</main>
<footer>Generado por scripts/07_generar_dashboard_etapa_productiva.py · SENA Centro ASTIN · Coordinación de Etapa Productiva</footer>

<script>
const DATA = __DATA_JSON__;
const INSTR = __INSTR_JSON__;
const META = __META_JSON__;

const ICON_OK = "✅", ICON_PAR = "⚠️", ICON_NO = "❌";
const estadoActual = { ficha:"", programa:"", anios:[], buscar:"", estado:"", atencion:false };
let ANIOS_DISPONIBLES = [];

function esc(s){ return String(s ?? "").replace(/[&<>"']/g, c => (
  {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }

function fileUri(p){
  const s = String(p).replace(/\\/g, "/");
  const partes = s.split("/").map((seg, i) => i === 0 ? seg : encodeURIComponent(seg));
  return "file:///" + partes.join("/");
}

function renderKPIs(){
  const k = META;
  document.getElementById("kpis").innerHTML = `
    <div class="kpi navy"><div class="titulo">Total Aprendices Virtuales</div>
      <div class="valor">${k.total}</div><div class="mini"><b>${k.total_fichas}</b> fichas activas</div></div>
    <div class="kpi"><div class="titulo">Alternativa Registrada</div>
      <div class="valor">${k.pct_alternativa}%</div><div class="mini">de los aprendices virtuales</div></div>
    <div class="kpi amber"><div class="titulo">Seguimientos Asignados</div>
      <div class="valor">${k.pct_seg1}% · ${k.pct_seg2}% · ${k.pct_seg3}%</div>
      <div class="mini"><b>1.º</b> ${k.pct_seg1}% &nbsp;·&nbsp; <b>2.º</b> ${k.pct_seg2}% &nbsp;·&nbsp; <b>3.º</b> ${k.pct_seg3}%</div></div>
    <div class="kpi ${k.pct_completos >= 60 ? '' : 'red'}"><div class="titulo">Expedientes</div>
      <div class="valor">${k.pct_completos}%</div>
      <div class="mini"><b>Completos:</b> ${k.n_al_dia} (${k.pct_completos}%) &nbsp;·&nbsp;
      <b>Pendientes:</b> ${k.n_en_proceso + k.n_alerta} (${k.pct_pendientes}%)</div></div>
    <div class="kpi red"><div class="titulo">Vencimientos</div>
      <div class="valor">${k.n_venc_sel + k.n_venc_ep}</div>
      <div class="mini"><b>Sel. Alt. vencida:</b> ${k.n_venc_sel} (próx: ${k.n_prox_sel}) &nbsp;·&nbsp;
      <b>EP vencida:</b> ${k.n_venc_ep} (próx: ${k.n_prox_ep})</div></div>
    <div class="kpi red"><div class="titulo">Fin Lectiva · Atención Inmediata</div>
      <div class="valor">${k.n_fin_aten}</div>
      <div class="mini"><b>${k.n_fin_aten_alerta}</b> en ALERTA &nbsp;·&nbsp; año en curso: ${k.n_fin_vigente} &nbsp;·&nbsp; a futuro: ${k.n_fin_futuro}</div></div>`;
}

function poblarSelects(){
  const fichas = [...new Set(DATA.map(d => d.ficha))].sort();
  const programas = [...new Set(DATA.map(d => d.programa).filter(Boolean))].sort();
  const anios = [...new Set(DATA.map(d => d.fin_lectiva ? d.fin_lectiva.slice(0,4) : "").filter(Boolean))].sort();
  const sf = document.getElementById("sel-ficha");
  const sp = document.getElementById("sel-programa");
  fichas.forEach(f => { const o = document.createElement("option"); o.value = f; o.textContent = f; sf.appendChild(o); });
  programas.forEach(p => { const o = document.createElement("option"); o.value = p; o.textContent = p; sp.appendChild(o); });
  // Multi-select de años de Fin Lectiva (checkboxes)
  const panel = document.getElementById("panel-anio");
  panel.innerHTML = `
    <div class="dd-actions">
      <button type="button" id="btn-anio-todos">Todos</button>
      <button type="button" id="btn-anio-ninguno">Ninguno</button>
    </div>` + anios.map(a => `
    <label class="dd-item"><input type="checkbox" class="chk-anio" value="${a}" checked> ${a}</label>`).join("");
  ANIOS_DISPONIBLES = anios;
  estadoActual.anios = [...anios]; // inicialmente: todos los años seleccionados
}

function actualizarBotonAnio(){
  const btn = document.getElementById("btn-anio");
  if (estadoActual.anios.length === 0) btn.textContent = "Ningún año";
  else if (estadoActual.anios.length === ANIOS_DISPONIBLES.length) btn.textContent = "Todos los años";
  else btn.textContent = estadoActual.anios.slice().sort().join(", ") + " (" + estadoActual.anios.length + ")";
}

function filtrar(){
  const q = estadoActual.buscar.trim().toLowerCase();
  return DATA.filter(d => {
    if (estadoActual.ficha && d.ficha !== estadoActual.ficha) return false;
    if (estadoActual.programa && d.programa !== estadoActual.programa) return false;
    if (estadoActual.anios.length && (!d.fin_lectiva || !estadoActual.anios.includes(d.fin_lectiva.slice(0,4)))) return false;
    if (estadoActual.estado && d.estado !== estadoActual.estado) return false;
    if (estadoActual.atencion && d.fin_lectiva_estado !== "atencion") return false;
    if (q){
      const blob = (d.documento + " " + d.nombre + " " + d.empresa + " " + d.ficha).toLowerCase();
      if (!blob.includes(q)) return false;
    }
    return true;
  });
}

function fmtFecha(iso){
  if (!iso) return "—";
  const p = iso.split("-");
  return p[2] + "/" + p[1] + "/" + p[0];
}

function celdaFecha(iso, estadoFecha){
  if (!iso) return `<span class="chip no">${ICON_NO} Sin dato</span>`;
  const m = { vencido: ["no", ICON_NO], proximo: ["parcial", ICON_PAR], vigente: ["ok", ICON_OK] };
  const [cls, ico] = m[estadoFecha] || ["ok", ICON_OK];
  const etiqueta = estadoFecha === "vencido" ? "VENCIDO" :
                   estadoFecha === "proximo" ? "Próximo a vencer" : "Dentro de plazo";
  return `<div class="fecha-cell"><span class="d">${fmtFecha(iso)}</span><br>
    <span class="chip ${cls}">${ico} ${etiqueta}</span></div>`;
}

function celdaFinLectiva(d){
  if (!d.fin_lectiva) return `<span class="chip no">${ICON_NO} Sin dato</span>`;
  const m = { atencion: ["no", "🔴"], vigente: ["parcial", "🟡"], futuro: ["ok", "🟢"] };
  const [cls, ico] = m[d.fin_lectiva_estado] || ["ok", "🟢"];
  const etiqueta = d.fin_lectiva_estado === "atencion" ? "Atención inmediata" :
                   d.fin_lectiva_estado === "vigente" ? "Año en curso" : "A futuro";
  return `<div class="fecha-cell"><span class="d">${fmtFecha(d.fin_lectiva)}</span><br>
    <span class="chip ${cls}">${ico} ${etiqueta}</span></div>`;
}

function celdaAlternativa(d){
  if (!d.alternativa_ok) return `<span class="chip no">${ICON_NO} Sin registrar</span>`;
  return `<span class="chip ok">${ICON_OK} ${esc(d.alternativa || "Registrada")}</span>`;
}

function celdaARL(d){
  if (!d.arl_ok) return `<span class="chip no">${ICON_NO} Pendiente</span>`;
  return `<span class="chip ok">${ICON_OK} Afiliado${d.arl ? " · " + esc(d.arl) : ""}</span>`;
}

function celdaSeguimientos(d){
  const labels = ["1.º", "2.º", "3.º"];
  return [1,2,3].map(n => {
    const s = d.seg[n];
    const cls = s.ok ? "seg-ok" : "seg-no";
    const ico = s.ok ? ICON_OK : ICON_NO;
    return `<span class="seg-cell ${cls}" title="Seguimiento ${n}">${labels[n-1]} ${ico}</span>`;
  }).join("");
}

function listaArchivos(lista, etiqueta, icono){
  if (!lista || !lista.length) return `<span class="chip no">${ICON_NO} ${etiqueta}</span>`;
  const max = 8;
  const visibles = lista.slice(0, max);
  const extra = lista.length - max;
  const items = visibles.map(p => `<li><a href="${fileUri(p)}" target="_blank" rel="noopener"
    title="${esc(p)}">${esc(p.split(/[\\/]/).pop())}</a></li>`).join("");
  return `<details><summary>${icono} ${etiqueta} (${lista.length})</summary>
    <ul>${items}${extra > 0 ? `<li>+ ${extra} más…</li>` : ""}</ul></details>`;
}

function celdaSoportes(d){
  const a = d.archivos;
  const partes = [];
  if (a.inicio.length) partes.push(listaArchivos(a.inicio, "Inicio", "🚀"));
  if (a.arl.length) partes.push(listaArchivos(a.arl, "ARL", "🛡️"));
  if (a.bitacoras.length) partes.push(listaArchivos(a.bitacoras, `Bitácoras (${a.bitacoras.length})`, "📒"));
  if (a.seg1.length) partes.push(listaArchivos(a.seg1, "Seguimiento 1", "1️⃣"));
  if (a.seg2.length) partes.push(listaArchivos(a.seg2, "Seguimiento 2", "2️⃣"));
  if (a.seg3.length) partes.push(listaArchivos(a.seg3, "Seguimiento 3", "3️⃣"));
  if (a.cierre.length) partes.push(listaArchivos(a.cierre, "Cierre", "🏁"));
  if (!partes.length) return `<span class="chip no">${ICON_NO} Sin soportes</span>`;
  return `<div class="soportes">${partes.join("")}</div>`;
}

function renderTabla(){
  const filas = filtrar();
  const tbody = document.getElementById("tbody");
  if (!filas.length){
    tbody.innerHTML = `<tr><td colspan="12" class="vaciar">No hay aprendices que coincidan con los filtros.</td></tr>`;
  } else {
    tbody.innerHTML = filas.map(d => `
      <tr>
        <td class="ficha">${esc(d.ficha)}</td>
        <td class="aprendiz"><b>${esc(d.nombre || "—")}</b><br>
            <span class="doc">${esc(d.tipo_doc)} ${esc(d.documento)}</span></td>
        <td class="programa">${esc(d.programa)}</td>
        <td>${celdaAlternativa(d)}</td>
        <td>${d.empresa ? esc(d.empresa) : `<span class="chip no">${ICON_NO} Sin empresa</span>`}</td>
        <td>${celdaARL(d)}</td>
        <td>${celdaSeguimientos(d)}</td>
        <td>${celdaFinLectiva(d)}</td>
        <td>${celdaFecha(d.venc_sel_alt, d.venc_sel_alt_estado)}</td>
        <td>${celdaFecha(d.vencimiento, d.vencimiento_estado)}</td>
        <td>${celdaSoportes(d)}</td>
        <td><span class="estado est-${d.estado === "AL DÍA" ? "AL" : d.estado === "EN PROCESO" ? "EN" : "ALERTA"}">${esc(d.estado)}</span></td>
      </tr>`).join("");
  }
  document.getElementById("contador").textContent =
    `Mostrando ${filas.length} de ${DATA.length} aprendices virtuales.`;
}

function wireEvents(){
  document.getElementById("sel-ficha").addEventListener("change", e => {
    estadoActual.ficha = e.target.value; renderTabla(); });
  document.getElementById("sel-programa").addEventListener("change", e => {
    estadoActual.programa = e.target.value; renderTabla(); });
  document.getElementById("buscar").addEventListener("input", e => {
    estadoActual.buscar = e.target.value; renderTabla(); });
  // Multi-select de años de Fin Lectiva
  document.getElementById("btn-anio").addEventListener("click", e => {
    e.stopPropagation();
    document.getElementById("panel-anio").classList.toggle("oculto");
  });
  document.addEventListener("click", e => {
    const dd = document.getElementById("dd-anio");
    if (!dd.contains(e.target)) document.getElementById("panel-anio").classList.add("oculto");
  });
  document.getElementById("panel-anio").addEventListener("change", e => {
    if (e.target.classList.contains("chk-anio")){
      estadoActual.anios = [...document.querySelectorAll(".chk-anio:checked")].map(c => c.value);
      actualizarBotonAnio();
      renderTabla();
    }
  });
  document.getElementById("btn-anio-todos").addEventListener("click", () => {
    document.querySelectorAll(".chk-anio").forEach(c => c.checked = true);
    estadoActual.anios = [...ANIOS_DISPONIBLES];
    actualizarBotonAnio();
    renderTabla();
  });
  document.getElementById("btn-anio-ninguno").addEventListener("click", () => {
    document.querySelectorAll(".chk-anio").forEach(c => c.checked = false);
    estadoActual.anios = [];
    actualizarBotonAnio();
    renderTabla();
  });
  document.getElementById("botones-estado").addEventListener("click", e => {
    const btn = e.target.closest(".btn-estado"); if (!btn) return;
    if (btn.dataset.aten){
      estadoActual.atencion = !estadoActual.atencion;
      btn.style.background = estadoActual.atencion ? "var(--rojo)" : "#FDE4E4";
      btn.style.color = estadoActual.atencion ? "#fff" : "var(--rojo)";
      renderTabla();
      return;
    }
    estadoActual.estado = btn.dataset.estado;
    document.querySelectorAll(".btn-estado").forEach(b => b.className = "btn-estado");
    btn.classList.add("activo-" + (btn.dataset.estado === "" ? "todos" :
      btn.dataset.estado === "AL DÍA" ? "aldia" : btn.dataset.estado === "EN PROCESO" ? "proceso" : "alerta"));
    renderTabla();
  });
}

// ============================ VISTA INSTRUCTORES EP ============================
const estadoInstr = { buscar: "", rol: "" };

function mostrarTab(tab){
  const esAprendices = tab === "aprendices";
  document.getElementById("vista-aprendices").classList.toggle("oculto", !esAprendices);
  document.getElementById("vista-instructores").classList.toggle("oculto", esAprendices);
  document.getElementById("tab-aprendices").classList.toggle("activo", esAprendices);
  document.getElementById("tab-instructores").classList.toggle("activo", !esAprendices);
}

function renderKPIsInstr(){
  const k = META.instr || {};
  document.getElementById("kpis-instr").innerHTML = `
    <div class="kpi navy"><div class="titulo">Instructores EP</div>
      <div class="valor">${k.total}</div><div class="mini">con asignación de aprendices virtuales</div></div>
    <div class="kpi"><div class="titulo">Asignaciones Totales</div>
      <div class="valor">${k.asignaciones}</div><div class="mini">roles asignados (seg. 1·2·3, líder, técnico, consultorio)</div></div>
    <div class="kpi amber"><div class="titulo">Carga Promedio</div>
      <div class="valor">${k.promedio}</div><div class="mini">aprendices por instructor</div></div>
    <div class="kpi red"><div class="titulo">Instructores con Alertas</div>
      <div class="valor">${k.con_alerta}</div><div class="mini">con al menos 1 aprendiz en ALERTA</div></div>`;
}

function filtrarInstr(){
  const q = estadoInstr.buscar.trim().toLowerCase();
  return INSTR.filter(i => {
    if (estadoInstr.rol && !(i.roles[estadoInstr.rol])) return false;
    if (q){
      const blob = (i.nombre + " " + i.fichas.join(" ")).toLowerCase();
      if (!blob.includes(q)) return false;
    }
    return true;
  });
}

function chipRol(rol, n){
  const mapa = {"LÍDER":"📋","SEGUIMIENTO 1":"1️⃣","SEGUIMIENTO 2":"2️⃣",
                "SEGUIMIENTO 3":"3️⃣","TÉCNICO":"🔧","CONSULTORIO":"🗂️"};
  return `<span class="chip" style="background:#E9ECEF;color:var(--navy);margin:1px 2px;"
    title="${esc(rol)}">${mapa[rol] || ""} ${esc(rol)} ×${n}</span>`;
}

function renderTablaInstr(){
  const filas = filtrarInstr();
  const tbody = document.getElementById("tbody-instr");
  if (!filas.length){
    tbody.innerHTML = `<tr><td colspan="8" class="vaciar">No hay instructores que coincidan con los filtros.</td></tr>`;
  } else {
    tbody.innerHTML = filas.map(i => {
      const e = i.estados;
      const total = i.n_aprendices;
      const alDiaPct = total ? Math.round(100 * e["AL DÍA"] / total) : 0;
      const roles = Object.keys(i.roles).map(r => chipRol(r, i.roles[r])).join("");
      const aprendicesDetalle = i.aprendices.map(a => `
        <li>${esc(a.nombre)} · ficha ${esc(a.ficha)} · ${esc(a.documento)}
          <span class="chip ${a.estado === "AL DÍA" ? "ok" : a.estado === "EN PROCESO" ? "parcial" : "no"}">${esc(a.estado)}</span>
          <span class="chip" style="background:#E9ECEF;color:var(--navy);">${a.roles.map(esc).join(", ")}</span></li>`).join("");
      const barraColor = alDiaPct >= 60 ? "var(--verde)" : alDiaPct >= 30 ? "var(--amarillo)" : "var(--rojo)";
      return `
      <tr>
        <td class="aprendiz"><b>${esc(i.nombre)}</b><br><span class="doc">${i.n_fichas} ficha(s)</span></td>
        <td>${roles}</td>
        <td><details><summary>${i.n_fichas}</summary><ul>${i.fichas.map(f => `<li>${esc(f)}</li>`).join("")}</ul></details></td>
        <td><b>${total}</b> aprendices</td>
        <td><span class="chip ok">${e["AL DÍA"]}</span></td>
        <td><span class="chip parcial">${e["EN PROCESO"]}</span></td>
        <td><span class="chip no">${e["ALERTA"]}</span></td>
        <td>
          <div style="display:flex;align-items:center;gap:6px;">
            <div style="flex:1;min-width:90px;height:9px;background:#E9ECEF;border-radius:5px;overflow:hidden;">
              <div style="width:${alDiaPct}%;height:100%;background:${barraColor};"></div>
            </div>
            <span style="font-size:11.5px;font-weight:700;color:var(--navy);">${alDiaPct}%</span>
          </div>
          <div style="margin-top:4px;"><details><summary style="font-size:11px;">Ver aprendices</summary>
            <ul>${aprendicesDetalle}</ul></details></div>
        </td>
      </tr>`;
    }).join("");
  }
  document.getElementById("contador-instr").textContent =
    `Mostrando ${filas.length} de ${INSTR.length} instructores de Etapa Productiva.`;
}

function wireEventsInstr(){
  document.getElementById("buscar-instr").addEventListener("input", e => {
    estadoInstr.buscar = e.target.value; renderTablaInstr(); });
  document.getElementById("sel-rol").addEventListener("change", e => {
    estadoInstr.rol = e.target.value; renderTablaInstr(); });
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("fecha-corte").textContent = META.fecha_corte + " · " + META.generado;
  document.getElementById("cnt-aprendices").textContent = "(" + DATA.length + ")";
  document.getElementById("cnt-instructores").textContent = "(" + (META.instr ? META.instr.total : 0) + ")";
  renderKPIs();
  poblarSelects();
  renderTabla();
  wireEvents();
  renderKPIsInstr();
  renderTablaInstr();
  wireEventsInstr();
});
</script>
</body>
</html>
"""


def generar_html(expedientes, kpis, instructores, kpis_instr, fecha_corte):
    """Inyecta los datos como JSON embebido (archivo único y autónomo)."""
    data_json = json.dumps(expedientes, ensure_ascii=False)
    # Evita romper la etiqueta <script> ante cualquier "</" dentro de los datos
    data_json = data_json.replace("</", "<\\/")
    instr_json = json.dumps(instructores, ensure_ascii=False).replace("</", "<\\/")
    meta = {
        "fecha_corte": fecha_corte,
        "generado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        **kpis,
        "instr": kpis_instr,
    }
    meta_json = json.dumps(meta, ensure_ascii=False).replace("</", "<\\/")
    html = (
        PLANTILLA_HTML
        .replace("__DATA_JSON__", data_json)
        .replace("__INSTR_JSON__", instr_json)
        .replace("__META_JSON__", meta_json)
    )
    return html


# ----------------------------------------------------------------------------
# 9. MAIN (idempotente)
# ----------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Genera el Dashboard HTML de Etapa Productiva Virtual.")
    parser.add_argument("--solo-validar", action="store_true",
                        help="Solo lee el CSV y reporta cuántos aprendices VIRTUALES hay. No genera HTML.")
    parser.add_argument("--rescan", action="store_true",
                        help="Fuerza un reescaneo completo de OneDrive (ignora la caché).")
    parser.add_argument("--no-scan", action="store_true",
                        help="No escanea OneDrive (usa solo datos del CSV).")
    args = parser.parse_args()

    raiz = Path(__file__).resolve().parent.parent
    CONFIG["cache_index"] = raiz / "data" / "onedrive_index.json"
    CONFIG["output_html"] = raiz / "dashboard_etapa_productiva_virtual.html"
    CONFIG["resumen_dir"] = raiz / "outputs"

    print("=" * 78)
    print(" SENA CENTRO ASTIN · Generador del Dashboard de Etapa Productiva Virtual")
    print("=" * 78)

    # --- Lectura del CSV -----------------------------------------------------
    print("[1/4] Leyendo base de datos...")
    df = leer_base_datos(CONFIG["csv_path"])
    print(f"       Filas totales en CSV: {len(df)}")
    virtuales = filtrar_virtuales(df)
    print(f"       Aprendices MODALIDAD='VIRTUAL': {len(virtuales)}")
    if args.solo_validar:
        cmap = mapear_columnas(virtuales)
        n_fichas = virtuales.iloc[:, 0].nunique() if len(virtuales) else 0
        print(f"       Fichas virtuales únicas: {n_fichas}")
        print(f"       Columnas detectadas ({len(virtuales.columns)}):")
        for c in virtuales.columns:
            print(f"         - {c}")
        print("\n[OK] Validación exitosa: la lectura de aprendices virtuales funciona.")
        return

    # --- Escaneo de OneDrive --------------------------------------------------
    print("[2/4] Escaneando OneDrive...")
    if args.no_scan:
        files, idx = [], defaultdict(set)
        print("       Escaneo omitido (--no-scan).")
    else:
        files = escanear_onedrive(
            CONFIG["onedrive_root"], CONFIG["extensiones"],
            CONFIG["excluir_dirnames"], CONFIG["cache_index"], forzar=args.rescan,
        )
        idx = construir_indice(files)

    # --- Expedientes y KPIs ----------------------------------------------------
    print("[3/4] Construyendo expedientes (clave FICHA + DOCUMENTO)...")
    expedientes, duplicados = construir_expedientes(virtuales, files, idx)
    kpis = calcular_kpis(expedientes)
    print(f"       Expedientes: {len(expedientes)} | Duplicados omitidos: {duplicados}")
    print(f"       AL DÍA: {kpis['n_al_dia']} | EN PROCESO: {kpis['n_en_proceso']} | "
          f"ALERTA: {kpis['n_alerta']}")
    instructores = construir_instructores(virtuales, expedientes)
    kpis_instr = calcular_kpis_instructores(instructores)
    print(f"       Instructores EP: {kpis_instr['total']} | Asignaciones: {kpis_instr['asignaciones']}")

    # --- HTML ----------------------------------------------------------------
    print("[4/4] Generando dashboard HTML...")
    fecha_corte = datetime.now().strftime("%d/%m/%Y")
    html = generar_html(expedientes, kpis, instructores, kpis_instr, fecha_corte)
    CONFIG["output_html"].write_text(html, encoding="utf-8")
    print(f"       Dashboard generado: {CONFIG['output_html']}")
    print(f"       Tamaño: {CONFIG['output_html'].stat().st_size / 1024:.1f} KB")

    # --- Resumen JSON (trazabilidad del corte semanal) -------------------------
    CONFIG["resumen_dir"].mkdir(parents=True, exist_ok=True)
    resumen = {
        "generado": datetime.now().isoformat(timespec="seconds"),
        "csv_filas": len(df),
        "virtuales": len(virtuales),
        "expedientes": len(expedientes),
        "duplicados_omitidos": duplicados,
        "archivos_onedrive": len(files),
        "kpis": kpis,
        "kpis_instructores": kpis_instr,
    }
    resumen_path = CONFIG["resumen_dir"] / "resumen_ultima_generacion.json"
    resumen_path.write_text(json.dumps(resumen, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"       Resumen JSON: {resumen_path}")

    print("=" * 78)
    print(" [OK] Proceso completado. Abra el dashboard en su navegador.")
    print("=" * 78)


if __name__ == "__main__":
    main()
