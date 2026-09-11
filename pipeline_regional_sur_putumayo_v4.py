"""
Pipeline de indicadores operativos — Reporte Regional Sur y Putumayo  v4.1
===========================================================================
Ubicacion: C:\\MAESTRO_TRANSCARGA\\PIPELINES\\pipeline_regional_sur_putumayo_v4.py

ESTRUCTURA DE CARPETAS:
  C:\\MAESTRO_TRANSCARGA\\
  ├─ INSUMOS\\
  │   ├─ ARCHIVOS_CARGUES\\   <- Archivos Cargues_*.xls y Planilla_Global_*.xls
  │   └─ ARCHIVOS_REFERENCIA\\ <- CARGUES_TRAYECTOS.xlsx
  └─ SALIDA\\
      ├─ CARGUES\\            <- Salida de este pipeline
      └─ UNIFICAR\\           <- Directorio del TRANSCARGA_MAESTRO

NOVEDADES v4.0:
  - DETECCION DE GUIAS SIN CARGUE: Cruza el Maestro para identificar guias cuyo
    destino es la Regional Sur (Florencia, Pitalito, Garzon) pero que NO tienen
    planilla de cargue de salida desde Bogota. Tambien incluye despachos historicos
    a ciudades del Putumayo (Abril-Junio 2026).
  - Logica de distincion: se elimina la restriccion sobre N_PLANILLA_CARGUE del
    Maestro y se usa exclusivamente la ausencia de la guia en las planillas fisicas
    de cargues. Esto permite capturar guias que llegaron a Bogota (con cargue en
    un tramo anterior, ej Cali-Bogota) pero sin despacho desde Bogota al Sur.
  - ESTRUCTURA DE 31 COLUMNAS FIJA para la hoja GUIAS DETALLE:
    CARGUE, NUMERO GUIA, FECHA CARGUE, HORA, MES, ORIGEN, DESTINO, CIUDAD,
    CATEGORIA, TIPO TRAYECTO, ESTADO CARGUE, ESTADO GUIA, MODALIDAD, REMITENTE,
    DESTINATARIO, FLETE ($), SEGURO ($), TOTAL ($), V. DECLARADO, CANT, PESO (kg),
    PLACA, CONDUCTOR, FECHA CREACION GUIA (M), ESTADO MAESTRO, PIEZAS (M),
    TOTAL VALOR (M), ESTADO LEGALIZACION, FECHA DEL ESTADO, CIUDAD ORIGEN,
    REGIONAL ORIGEN
  - ORIGEN siempre = BOGOTA; DESTINO = REG DESTINO CORRECTA del Maestro.
  - HORA y MES extraidos de la fecha de cargue (o fecha de creacion si sin cargue).
  - ESTADO GUIA, ESTADO MAESTRO, PIEZAS (M), ESTADO LEGALIZACION, FECHA DEL ESTADO,
    CIUDAD ORIGEN, REGIONAL ORIGEN: todos vienen del MAESTRO.
  - Guias sin cargue aparecen en GUIAS DETALLE (fondo rosa) Y en hoja separada
    GUIAS SIN CARGUE.
  - Se excluyen estados ANULADO/CANCELADO; se incluyen todos los demas.

NOVEDADES v4.1 (reglas de negocio actualizadas):
  - VENTANA DE CONTROL: la operacion de despachos hacia el Sur inicio el
    24-abr-2026. El reporte solo considera despachos desde esa fecha; el Maestro
    puede contener despachos anteriores (ej. 30-mar-2026) que quedan excluidos.
  - PUTUMAYO HISTORICO ABR-JUN 2026: los despachos a ciudades del Putumayo se
    generaron entre abril y junio 2026; a partir de julio 2026 Transcarga dejo de
    enviarlos, por lo que el reporte considera solo ABRIL-MAYO-JUNIO.
  - CIUDADES SUR EN CARGUES NEIVA: desde julio 2026 Transcarga empezo a enrutar
    ciudades de Florencia/Pitalito/Garzon dentro de cargues destino NEIVA
    (Bogota->NEIVA). Esas ciudades siguen siendo de la Regional Sur; el pipeline
    las identifica por CIUDAD y por 'REG DESTINO CORRECTA' del Maestro, las
    incluye como guias CON cargue y corrige su REGIONAL_DESTINO al Sur real.

HERENCIA v3 (multi-archivo + multi-hoja):
  - Lee TODOS los Cargues_*.xls* y Planilla_Global_*.xls* disponibles.
  - Hojas de continuacion sin cabecera manejadas correctamente.
  - Deduplicacion inteligente por guia, priorizando BOGOTA y fecha mas reciente.
"""

import re
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# =============================================================================
#  RUTAS
# =============================================================================

_SCRIPT_DIR        = Path(__file__).resolve().parent
_BASE_DIR          = _SCRIPT_DIR.parent
_ARCHIVOS_CARGUES  = _BASE_DIR / "INSUMOS" / "ARCHIVOS_CARGUES"
_ARCHIVOS_REF      = _BASE_DIR / "INSUMOS" / "ARCHIVOS_REFERENCIA"
_SALIDA_CARGUES    = _BASE_DIR / "SALIDA"  / "CARGUES"
_SALIDA_UNIFICAR   = _BASE_DIR / "SALIDA"  / "UNIFICAR"
_TRAYECTOS_DEFAULT = _ARCHIVOS_REF / "CARGUES_TRAYECTOS.xlsx"

# =============================================================================
#  SECCION 1 — CONSTANTES
# =============================================================================

PUTUMAYO_CITIES = [
    "MOCOA", "COLON", "COLON PUT", "ORITO", "PUERTO ASIS",
    "PUERTO CAICEDO", "PUERTO GUZMAN", "PUERTO LEGUIZAMO",
    "SAN FRANCISCO", "SAN MIGUEL", "SAN MIGUEL PUT", "LA DORADA PUT",
    "SANTIAGO", "SIBUNDOY", "VALLE DEL GUAMUEZ", "LA HORMIGA", "VILLAGARZON"
]
PUTUMAYO_CITIES_SET = frozenset(c.strip().upper() for c in PUTUMAYO_CITIES)

REGIONALES_SUR = [
    "FLORENCIA - PRINCIPAL",
    "PITALITO - PRINCIPAL",
    "GARZON - PRINCIPAL",
]

# Valores en el campo 'REG DESTINO CORRECTA' del Maestro
REGIONALES_SUR_MAESTRO = ["FLORENCIA", "PITALITO", "GARZON"]

# ─────────────────────────────────────────────────────────────────────────────
#  VENTANA DE CONTROL Y REGLAS DE NEGOCIO (Susenvios - Sur + Putumayo)
# ─────────────────────────────────────────────────────────────────────────────
# Transcarga inicio los despachos hacia la zona Sur el 24-abr-2026. Aunque el
# Maestro contenga despachos desde el 30-mar-2026, el control operativo (guias
# con cargue y sin cargue) inicia el 24-abr-2026.
FECHA_INICIO_CONTROL = "2026-04-25"

# Putumayo: los despachos se generaron solo de abril a junio 2026; a partir de
# julio 2026 Transcarga dejo de enviar despachos a los destinos del Putumayo.
FECHA_PUTUMAYO_FIN = "2026-06-30"
MESES_PUTUMAYO = [4, 5, 6]  # ABRIL, MAYO, JUNIO 2026

# Ciudades de la Regional Sur (Susenvios) agrupadas por zona. Desde ~julio 2026
# Transcarga enruta varias de estas ciudades dentro de cargues destino NEIVA
# (Bogota->NEIVA); esas ciudades siguen siendo nuestras. La lista se derivo del
# analisis de 'REG DESTINO CORRECTA' del Maestro y de los cargues de junio 2026.
CIUDADES_SUR_FLORENCIA = [
    "FLORENCIA CAQ", "SAN VICENTE DEL CAGUAN", "EL DONCELLO", "CARTAGENA CAQ",
    "PUERTO RICO CAQ", "EL PAUJIL", "SAN JOSE DEL FRAGUA", "CURILLO", "SOLITA",
    "BELEN DE LOS ANDAQUIES", "VALPARAISO CAQ", "LA MONTAÑITA", "LA MONTANITA",
    "MORELIA", "CARTAGENA DEL CHAIRA", "ALBANIA CAQ",
]
CIUDADES_SUR_PITALITO = [
    "PITALITO", "SAN AGUSTIN", "TIMANA", "ISNOS", "TARQUI",
    "SALADOBLANCO", "OPORAPA", "PALESTINA HLA",
]
CIUDADES_SUR_GARZON = [
    "GARZON", "GIGANTE", "GUADALUPE HUI", "SUAZA", "ACEVEDO",
    "AGRADO HUILA", "ZULUAGA", "PITAL",
]
CIUDADES_SUR_POR_ZONA = {
    "FLORENCIA": CIUDADES_SUR_FLORENCIA,
    "PITALITO":  CIUDADES_SUR_PITALITO,
    "GARZON":    CIUDADES_SUR_GARZON,
}
CIUDADES_SUR = frozenset(
    c.strip().upper() for zonas in CIUDADES_SUR_POR_ZONA.values() for c in zonas
)

ESTADOS_ANULADO = ["ANULADO", "ANULADA", "CANCELADO", "CANCELADA"]

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

# =============================================================================
#  SECCION 2 — ESTILOS EXCEL (definidos primero para evitar NameError)
# =============================================================================

C = {
    "hdr_dark":   "1F3864",
    "hdr_mid":    "2F5496",
    "hdr_light":  "D6E4F7",
    "urbano":     "E2EFDA",
    "nacional":   "FFF2CC",
    "reexp":      "FCE4D6",
    "sin_cargue": "FFD7D7",
    "total":      "D9D9D9",
    "white":      "FFFFFF",
    "green":      "00B050",
    "orange":     "FF9900",
    "red":        "C00000",
    "amarillo":   "FFEB84",
    "font_dark":  "1F1F1F",
}

CAT_BG = {
    "URBANO":       C["urbano"],
    "NACIONAL":     C["nacional"],
    "REEXPEDICION": C["reexp"],
}

_th = Side(style="thin",   color="BFBFBF")
_md = Side(style="medium", color="595959")
B_THIN = Border(left=_th, right=_th, top=_th, bottom=_th)
B_HDR  = Border(left=_md, right=_md, top=_md, bottom=_md)


def _pct(num, den):
    return np.where(den > 0, np.round(num / den, 4), 0.0)


def _fill(c):
    return PatternFill("solid", fgColor=c)


def _font(bold=False, color="1F1F1F", size=9):
    return Font(name="Arial", bold=bold, color=color, size=size)


def _aln(h="center", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)


def _col_width(ws, mn=8, mx=55):
    for col in ws.columns:
        cl = get_column_letter(col[0].column)
        lens = [len(str(c.value)) for c in col if c.value is not None]
        ws.column_dimensions[cl].width = min(max(max(lens) + 2 if lens else mn, mn), mx)


def _sem_ef(v: float) -> str:
    if v >= 0.75: return C["green"]
    if v >= 0.45: return C["orange"]
    return C["red"]


def _sem_pend(v: float) -> str:
    if v <= 0.25: return C["green"]
    if v <= 0.55: return C["orange"]
    return C["red"]


def _encabezado(ws, cols, row, color=None):
    color = color or C["hdr_dark"]
    for j, col in enumerate(cols, 1):
        c = ws.cell(row=row, column=j, value=col)
        c.font      = _font(bold=True, color=C["white"], size=10)
        c.fill      = _fill(color)
        c.alignment = _aln("center", wrap=True)
        c.border    = B_HDR
    ws.row_dimensions[row].height = 30


def _banner(ws, texto, ncols, row=1, color=None):
    color = color or C["hdr_dark"]
    ws.merge_cells(f"A{row}:{get_column_letter(ncols)}{row}")
    c = ws[f"A{row}"]
    c.value     = texto
    c.font      = _font(bold=True, color=C["white"], size=13)
    c.fill      = _fill(color)
    c.alignment = _aln("center")
    ws.row_dimensions[row].height = 32


def _escribir_df(ws, df: pd.DataFrame, start_row: int,
                 pct_ef_cols=None, pct_pend_cols=None,
                 cat_col=None, num_cols=None, fecha_cols=None,
                 sin_cargue_col=None):
    """
    Escribe filas de un DataFrame en la hoja ws.
    sin_cargue_col: columna cuyo valor 'NO' activa fondo rosa.
    """
    pct_ef_cols   = pct_ef_cols   or []
    pct_pend_cols = pct_pend_cols or []
    num_cols      = num_cols      or []
    fecha_cols    = fecha_cols    or []
    ci = {name: i + 1 for i, name in enumerate(df.columns)}
    last_row = start_row

    for i, tup in enumerate(df.itertuples(index=False), start_row):
        row_vals = list(tup)

        # Color de fila
        row_bg = None
        if sin_cargue_col and sin_cargue_col in ci:
            if str(row_vals[ci[sin_cargue_col] - 1]).strip().upper() == "NO":
                row_bg = C["sin_cargue"]
        if row_bg is None and cat_col and cat_col in ci:
            cat = str(row_vals[ci[cat_col] - 1])
            row_bg = CAT_BG.get(cat)

        for j, val in enumerate(row_vals, 1):
            try:
                if pd.isnull(val):
                    val = ""
            except Exception:
                pass
            c = ws.cell(row=i, column=j, value=val)
            c.font      = _font(size=9)
            c.alignment = _aln("center")
            c.border    = B_THIN
            if row_bg:
                c.fill = _fill(row_bg)

        for col in num_cols:
            if col in ci:
                ws.cell(row=i, column=ci[col]).number_format = "#,##0"
        for col in fecha_cols:
            if col in ci:
                ws.cell(row=i, column=ci[col]).number_format = "YYYY-MM-DD"

        for col in pct_ef_cols:
            if col in ci:
                try:
                    fv = float(row_vals[ci[col] - 1])
                    c  = ws.cell(row=i, column=ci[col])
                    c.number_format = "0.0%"
                    c.font = _font(bold=True, color=C["white"], size=9)
                    c.fill = _fill(_sem_ef(fv))
                except Exception:
                    pass

        for col in pct_pend_cols:
            if col in ci:
                try:
                    fv = float(row_vals[ci[col] - 1])
                    c  = ws.cell(row=i, column=ci[col])
                    c.number_format = "0.0%"
                    c.font = _font(bold=True, color=C["white"], size=9)
                    c.fill = _fill(_sem_pend(fv))
                except Exception:
                    pass

        last_row = i
    return last_row


# =============================================================================
#  SECCION 3 — LECTURA MULTIPESTAÑA ROBUSTA
# =============================================================================

def _fix_xml(content: str) -> str:
    content = content.replace("<xml version>", '<?xml version="1.0" encoding="UTF-8"?>', 1)
    content = re.sub(r"urn:schemas-\s+microsoft-com", "urn:schemas-microsoft-com", content)
    content = re.sub(r"&(?!amp;|lt;|gt;|quot;|apos;|#)", "&amp;", content)
    content = re.sub(r"<(?=\s|\d)", "&lt;", content)
    return content


def leer_xls_xml_raw(filepath: str) -> dict:
    """Lee un archivo Excel 2003 XML devolviendo DataFrames crudos por hoja."""
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
    except UnicodeDecodeError:
        with open(filepath, "r", encoding="latin-1") as f:
            content = f.read()
    content = _fix_xml(content)
    root = etree.fromstring(content.encode("utf-8"))
    NS = "urn:schemas-microsoft-com:office:spreadsheet"
    sheets = {}
    for ws in root.findall(f".//{{{NS}}}Worksheet"):
        name = ws.get(f"{{{NS}}}Name", "Sheet")
        rows_data = []
        for row in ws.findall(f".//{{{NS}}}Row"):
            cells = [
                (cell.find(f"{{{NS}}}Data").text
                 if cell.find(f"{{{NS}}}Data") is not None else None)
                for cell in row.findall(f"{{{NS}}}Cell")
            ]
            rows_data.append(cells)
        if rows_data:
            max_cols = max(len(r) for r in rows_data)
            padded = [r + [None] * (max_cols - len(r)) for r in rows_data]
            df = pd.DataFrame(padded)
            sheets[name] = df
    return sheets


def _leer_flexible(filepath: str) -> pd.DataFrame:
    """Lee todas las hojas de un archivo, manejando hojas de continuacion sin cabecera."""
    p = Path(filepath)
    raw_sheets = {}

    if p.suffix.lower() == ".xls":
        try:
            raw_sheets = leer_xls_xml_raw(filepath)
        except Exception:
            try:
                raw_sheets = pd.read_excel(
                    filepath, dtype=str, engine="xlrd",
                    sheet_name=None, header=None
                )
            except Exception as e2:
                print(f"      ERROR leyendo {p.name}: {e2}")
                return pd.DataFrame()
    else:
        raw_sheets = pd.read_excel(
            filepath, dtype=str, engine="openpyxl",
            sheet_name=None, header=None
        )

    if not raw_sheets:
        return pd.DataFrame()

    dfs = []
    global_columns = None

    for _sheet_name, df in raw_sheets.items():
        if df.empty or len(df.columns) == 0:
            continue
        df = df.dropna(how="all").reset_index(drop=True)
        if df.empty:
            continue

        if global_columns is None:
            # Primera hoja: fila 0 = cabecera
            headers = df.iloc[0].astype(str).str.strip().tolist()
            global_columns = [
                c if c.lower() not in ["nan", "none", ""] else f"Unnamed_{i}"
                for i, c in enumerate(headers)
            ]
            df = df.iloc[1:].reset_index(drop=True)
        else:
            # Hojas de continuacion: detectar si se repite cabecera
            first_row = df.iloc[0].astype(str).str.strip().tolist()
            match_count = sum(1 for a, b in zip(first_row, global_columns) if a == b)
            if len(global_columns) > 0 and match_count > (len(global_columns) / 2):
                df = df.iloc[1:].reset_index(drop=True)

        # Ajustar ancho de columnas
        if len(df.columns) < len(global_columns):
            for i in range(len(df.columns), len(global_columns)):
                df[i] = None
        elif len(df.columns) > len(global_columns):
            df = df.iloc[:, :len(global_columns)]

        df.columns = global_columns
        dfs.append(df)

    if not dfs:
        return pd.DataFrame()

    df_final = pd.concat(dfs, ignore_index=True)
    if global_columns:
        col_0 = global_columns[0]
        df_final = df_final[df_final[col_0] != col_0].copy()
    return df_final


def _unificar_archivos(lista_paths: list) -> pd.DataFrame:
    """Lee y concatena multiples archivos, cada uno con posibles multiples hojas."""
    dfs = []
    for path in lista_paths:
        try:
            df = _leer_flexible(path)
            if not df.empty:
                df["__origen"] = Path(path).name
                dfs.append(df)
                print(f"      OK: {Path(path).name} ({len(df):,} filas)")
            else:
                print(f"      AVISO Vacio: {Path(path).name}")
        except Exception as e:
            print(f"      ERROR leyendo {Path(path).name}: {e}")
    if not dfs:
        return pd.DataFrame()
    df_total = pd.concat(dfs, ignore_index=True)
    df_total.drop(columns=["__origen"], inplace=True, errors="ignore")
    return df_total


def cargar_archivos(lista_cargues: list, lista_planilla: list, path_trayectos: str):
    """Carga en paralelo los archivos de Cargues, Planilla y Trayectos."""
    def _c(): return _unificar_archivos(lista_cargues)
    def _p(): return _unificar_archivos(lista_planilla)
    def _t(): return pd.read_excel(path_trayectos, sheet_name="TRAYECTOS")
    with ThreadPoolExecutor(max_workers=3) as ex:
        fc, fp, ft = ex.submit(_c), ex.submit(_p), ex.submit(_t)
        return fc.result(), fp.result(), ft.result()


# =============================================================================
#  SECCION 4 — MAPAS Y CLASIFICADORES
# =============================================================================

_MAP_REG = {
    "BOGOTA":       "BOGOTA - PRINCIPAL",
    "MEDELLIN":     "MEDELLIN - PRINCIPAL",
    "CALI":         "CALI - PRINCIPAL",
    "BUCARAMANGA":  "BUCARAMANGA - PRINCIPAL",
    "EJE":          "EJE CAFETERO - PRINCIPAL",
    "CAFETERO":     "EJE CAFETERO - PRINCIPAL",
    "VILLAVICENCIO":"VILLAVICENCIO - PRINCIPAL",
    "IBAGUE":       "IBAGUE - PRINCIPAL",
    "CUCUTA":       "CUCUTA-PRINCIPAL",
    "CUCUTA-PRINCIPAL": "CUCUTA-PRINCIPAL",
    "NEIVA":        "NEIVA - PRINCIPAL",
    "YOPAL":        "YOPAL - PRINCIPAL",
    "PITALITO":     "PITALITO - PRINCIPAL",
    "GARZON":       "GARZON - PRINCIPAL",
    "FLORENCIA":    "FLORENCIA - PRINCIPAL",
    "FUSAGASUGA":   "FUSAGASUGA - PRINCIPAL",
    "BOYACA":       "BOYACA - PRINCIPAL",
    "PASTO":        "PASTO - PRINCIPAL",
    "PUTUMAYO":     "PUTUMAYO - PRINCIPAL",
}


def _nombre_completo(raw: str) -> str:
    u = str(raw).strip().upper()
    if u in _MAP_REG:
        return _MAP_REG[u]
    for k, v in _MAP_REG.items():
        if k in u:
            return v
    return str(raw).strip()


def _clasificar_categoria(tipo: str) -> str:
    if tipo == "URBANO":
        return "URBANO"
    if tipo.startswith("NAL DEST"):
        return "NACIONAL"
    if tipo.startswith("DEST REEXP"):
        return "REEXPEDICION"
    return "OTROS"


def _regional_responsable(row) -> str:
    tipo   = str(row.get("TIPO_TRAYECTO", ""))
    cat    = str(row.get("CATEGORIA_TRAYECTO", ""))
    nomreg = str(row.get("NOMREG", ""))
    if cat == "URBANO":
        return nomreg
    if cat == "NACIONAL":
        if tipo.startswith("NAL DEST "):
            return _nombre_completo(tipo.replace("NAL DEST ", "").strip())
        zona = str(row.get("NOMZONA", ""))
        if "/" in zona:
            return zona.split("/")[-1].strip()
        return nomreg
    if cat == "REEXPEDICION":
        partes = tipo.replace("DEST REEXP ", "").split()
        return _nombre_completo(partes[0]) if partes else nomreg
    return nomreg


# Mapa ciudad -> 'ZONA - PRINCIPAL' para las ciudades de la Regional Sur
_ZONA_CIUDAD = {
    c.strip().upper(): zona + " - PRINCIPAL"
    for zona, ciudades in CIUDADES_SUR_POR_ZONA.items()
    for c in ciudades
}


def _regional_sur_ciudad(ciudad: str, dept: str = "") -> str:
    """Devuelve 'ZONA - PRINCIPAL' si la ciudad/depto pertenece a Regional Sur o Putumayo; si no, ''."""
    c_up = str(ciudad).strip().upper()
    d_up = str(dept).strip().upper()
    if c_up in PUTUMAYO_CITIES_SET or d_up == "PUTUMAYO":
        return "PUTUMAYO - PRINCIPAL"
    return _ZONA_CIUDAD.get(c_up, "")


def _corregir_regional_destino(df: pd.DataFrame) -> pd.DataFrame:
    """
    Corrige REGIONAL_DESTINO de guias CON cargue:
    - Asigna 'PUTUMAYO - PRINCIPAL' a los destinos del departamento de Putumayo.
    - Corrige guias de Florencia/Pitalito/Garzon enrutadas por Transcarga en cargues de otras regionales.
    """
    df = df.copy()
    if "REGIONAL_DESTINO" not in df.columns:
        return df
    ciudad = df.get("CIUDAD", pd.Series("", index=df.index)).astype(str).str.strip().str.upper()
    dept = df.get("DEPARTAMENTO", pd.Series("", index=df.index)).astype(str).str.strip().str.upper()
    if "CIUDAD_DESTINO_M" in df.columns:
        ciudad_m = df["CIUDAD_DESTINO_M"].astype(str).str.strip().str.upper()
        ciudad = ciudad.where(ciudad != "", ciudad_m)

    es_putumayo = ciudad.isin(PUTUMAYO_CITIES_SET) | (dept == "PUTUMAYO")
    corr_ciudad = ciudad.map(lambda c: _regional_sur_ciudad(c, ""))

    corr_maestro = pd.Series("", index=df.index)
    if "REG_DESTINO_CORRECTA_M" in df.columns:
        reg_corr = df["REG_DESTINO_CORRECTA_M"].astype(str).str.strip().str.upper()
        sur_set = {r.upper() for r in REGIONALES_SUR_MAESTRO}
        corr_maestro = reg_corr.where(reg_corr.isin(sur_set), "")
        corr_maestro = corr_maestro.map(lambda r: _nombre_completo(r) if r else "")

    regional_correcta = corr_maestro.where(corr_maestro.str.strip() != "", corr_ciudad)

    actual = df["REGIONAL_DESTINO"].astype(str).str.strip()
    es_sur_actual = actual.str.upper().isin({r.upper() for r in REGIONALES_SUR})
    sin_correccion = regional_correcta.str.strip() == ""

    df["REGIONAL_DESTINO"] = actual.where((es_sur_actual | sin_correccion) & ~es_putumayo, regional_correcta)
    df.loc[es_putumayo, "REGIONAL_DESTINO"] = "PUTUMAYO - PRINCIPAL"
    return df


# =============================================================================
#  SECCION 5 — LIMPIEZA Y DEDUPLICACION
# =============================================================================

def limpiar_cargues(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    for col in ["CANTGUIAS", "CANTENTREGAS", "CANTPEDIDAS", "CANTNOVEDAD", "TOTALCARGUE"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "NOMEST" in df.columns:
        df = df[~df["NOMEST"].astype(str).str.upper().str.strip()
                .isin(ESTADOS_ANULADO)].copy()
    for col in ["CODCARGUE", "NOMREG", "NOMZONA", "NOMEST"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    if "FECREG" in df.columns:
        df["FECREG"] = pd.to_datetime(df["FECREG"], errors="coerce", format="mixed")
    for c in ["CANTGUIAS", "CANTENTREGAS", "CANTPEDIDAS", "CANTNOVEDAD"]:
        if c not in df.columns:
            df[c] = 0
    df["SIN_GESTION"] = (
        df["CANTGUIAS"] - df["CANTENTREGAS"] - df["CANTPEDIDAS"] - df["CANTNOVEDAD"]
    ).clip(lower=0)
    return df


def limpiar_planilla(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    if "NOMBRE_ESTADO" in df.columns:
        df = df[~df["NOMBRE_ESTADO"].astype(str).str.upper().str.strip().isin([e.upper() for e in ESTADOS_ANULADO])].copy()
    if "NUMERO_GUIA" in df.columns:
        df["NUMERO_GUIA"] = df["NUMERO_GUIA"].astype(str).str.strip()
    if "FECHA_REGISTRO" in df.columns:
        df["FECHA_REGISTRO"] = pd.to_datetime(
            df["FECHA_REGISTRO"], errors="coerce", format="mixed"
        )

    # Deduplicacion inteligente: prioriza registros de BOGOTA mas recientes
    if "NUMERO_GUIA" in df.columns:
        df["__score"] = 0
        if "NOMBRE_ZONA" in df.columns:
            df["__score"] = (
                df["NOMBRE_ZONA"].astype(str).str.upper()
                .str.startswith("BOGOTA -").astype(int)
            )
        sort_cols = ["NUMERO_GUIA", "__score"]
        asc_cols  = [True, False]
        if "FECHA_REGISTRO" in df.columns:
            sort_cols.append("FECHA_REGISTRO")
            asc_cols.append(False)
        df = df.sort_values(by=sort_cols, ascending=asc_cols)
        antes = len(df)
        df = df.drop_duplicates(subset=["NUMERO_GUIA"], keep="first").copy()
        df.drop(columns=["__score"], inplace=True, errors="ignore")
        print(f"      [DEDUP] {antes:,} -> {len(df):,} guias unicas en Planilla.")

    for col in ["FLETE", "SEGURO", "CARGOS_ANEXOS", "TOTAL", "VALOR_DECLARADO", "CANTIDAD", "PESO"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in ["CODIGO_CARGUE", "REGIONAL_ORIGEN", "NOMBRE_ZONA",
                "NOMBRE_ESTADO", "MODALIDAD", "CIUDAD"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    return df


def enriquecer(df_c, df_p, df_t):
    if df_c.empty or df_t.empty:
        return df_c, df_p
    df_t = df_t.copy()
    mask = df_t["TIPO TRAYECTO"] == "NOMZONA"
    df_t.loc[mask, "TIPO TRAYECTO"] = df_t.loc[mask, "NOMZONA"].apply(
        lambda z: "NAL DEST " + z.split("/")[-1].strip().upper()
                  .replace("- PRINCIPAL", "").replace("-PRINCIPAL", "").strip()
        if "/" in str(z) else "NAL DEST CUCUTA"
    )
    df_t = df_t.rename(columns={"TIPO TRAYECTO": "TIPO_TRAYECTO"})
    df_t["_NR"] = df_t["NOMREG"].str.strip().str.upper()
    df_t["_NZ"] = df_t["NOMZONA"].str.strip().str.upper()

    df_c = df_c.copy()
    if "NOMREG" in df_c.columns and "NOMZONA" in df_c.columns:
        df_c["_NR"] = df_c["NOMREG"].astype(str).str.strip().str.upper()
        df_c["_NZ"] = df_c["NOMZONA"].astype(str).str.strip().str.upper()
        df_c = df_c.merge(df_t[["_NR", "_NZ", "TIPO_TRAYECTO"]], on=["_NR", "_NZ"], how="left")
        df_c["TIPO_TRAYECTO"]        = df_c["TIPO_TRAYECTO"].fillna("SIN CLASIFICAR")
        df_c["CATEGORIA_TRAYECTO"]   = df_c["TIPO_TRAYECTO"].map(_clasificar_categoria)
        df_c["REGIONAL_RESPONSABLE"] = df_c.apply(_regional_responsable, axis=1)
        df_c.drop(columns=["_NR", "_NZ"], inplace=True)

    if not df_p.empty:
        df_p = df_p.copy()
        if "REGIONAL_ORIGEN" in df_p.columns and "NOMBRE_ZONA" in df_p.columns:
            df_p["_NR"] = df_p["REGIONAL_ORIGEN"].astype(str).str.strip().str.upper()
            df_p["_NZ"] = df_p["NOMBRE_ZONA"].astype(str).str.strip().str.upper()
            df_p = df_p.merge(df_t[["_NR", "_NZ", "TIPO_TRAYECTO"]], on=["_NR", "_NZ"], how="left")
            df_p["TIPO_TRAYECTO"]      = df_p["TIPO_TRAYECTO"].fillna("SIN CLASIFICAR")
            df_p["CATEGORIA_TRAYECTO"] = df_p["TIPO_TRAYECTO"].map(_clasificar_categoria)
            df_p.drop(columns=["_NR", "_NZ"], inplace=True)
    return df_c, df_p


# =============================================================================
#  SECCION 6 — CRUCE CON MAESTRO
# =============================================================================

def _cargar_maestro(maestro_dir: Path):
    """Carga el TRANSCARGA_MAESTRO mas reciente. Devuelve (df, nombre) o (vacio, None)."""
    archivos_m = list(maestro_dir.glob("TRANSCARGA_MAESTRO*.xlsx"))
    usar_excel = True
    if not archivos_m:
        archivos_m = list(maestro_dir.glob("TRANSCARGA_MAESTRO*.csv"))
        usar_excel = False
    if not archivos_m:
        return pd.DataFrame(), None

    archivos_m.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    latest = archivos_m[0]

    cols_needed = [
        "GUIA", "FECHA_DE_CREACION", "NOMEST", "FECREGEST",
        "PIEZA", "PESOS_FISICO", "FLETE", "SEGURO", "TOTAL", "VALOR_DECLARADO",
        "REGIONAL_ORIGEN", "REGIONAL_DESTINO", "REG DESTINO CORRECTA",
        "CIUDAD_ORIGEN", "CIUDAD_DESTINO", "DEPARTAMENTO",
        "REMITENTE_CLIENTE", "DESTINATARIO_CLIENTE", "MODALIDAD",
        "CLIENTE", "COMERCIAL", "LINEA COMERCIAL",
        "N_PLANILLA_CARGUE", "FECHA_CARGUE",
        "ESTADO OPERATIVO", "TIPO TRAYECTO", "TIPO_COBERTURA_2026",
        "ESTADO_LEGALIZACION",
    ]
    try:
        if usar_excel:
            df_m = pd.read_excel(
                latest, usecols=lambda c: c in cols_needed,
                dtype=str, engine="openpyxl"
            )
        else:
            df_m = pd.read_csv(latest, usecols=lambda c: c in cols_needed, dtype=str)
    except Exception as e:
        print(f"      ERROR leyendo Maestro: {e}")
        return pd.DataFrame(), None

    return df_m, latest.name


def cruzar_con_maestro(df_p: pd.DataFrame, maestro_dir: Path) -> pd.DataFrame:
    """
    Enriquece la planilla (guias CON cargue) con datos del Maestro:
    estado, fecha creacion, piezas, legalizacion, ciudad origen, regional origen, etc.
    """
    print("\n  [+] Cruzando Planilla con MAESTRO TRANSCARGA...")
    if df_p.empty:
        return df_p

    df_m, nombre = _cargar_maestro(maestro_dir)
    if df_m.empty:
        print("      AVISO: NO SE ENCONTRO NINGUN TRANSCARGA MAESTRO.")
        print("             Ejecute primero consolidar_transcarga_v8_0.py")
        for col, val in [("FECHA CREACION GUIA (M)", "SIN DATOS"),
                          ("ESTADO GUIA MAESTRO", "SIN DATOS"),
                          ("PIEZAS_M", 0), ("TOTAL_MAESTRO", 0)]:
            df_p[col] = val
        return df_p

    print(f"      Maestro: {nombre}")

    # Renombrar columnas del Maestro para el merge
    renames = {
        "FECHA_DE_CREACION":    "FECHA CREACION GUIA (M)",
        "NOMEST":               "ESTADO GUIA MAESTRO",
        "PIEZA":                "PIEZAS_M",
        "TOTAL":                "TOTAL_MAESTRO",
        "VALOR_DECLARADO":      "VALOR_DECLARADO_M",
        "FLETE":                "FLETE_M",
        "SEGURO":               "SEGURO_M",
        "ESTADO_LEGALIZACION":  "ESTADO LEGALIZACION",
        "FECREGEST":            "FECHA DEL ESTADO",
        "CIUDAD_ORIGEN":        "CIUDAD ORIGEN (M)",
        "REGIONAL_ORIGEN":      "REGIONAL ORIGEN (M)",
        "CIUDAD_DESTINO":       "CIUDAD_DESTINO_M",
        "REGIONAL_DESTINO":     "REGIONAL_DESTINO_M",
        "REG DESTINO CORRECTA": "REG_DESTINO_CORRECTA_M",
        "REMITENTE_CLIENTE":    "REMITENTE_M",
        "DESTINATARIO_CLIENTE": "DESTINATARIO_M",
        "MODALIDAD":            "MODALIDAD_M",
    }
    df_m2 = df_m.rename(columns={k: v for k, v in renames.items() if k in df_m.columns})
    df_m2["GUIA"] = df_m2["GUIA"].astype(str).str.strip()
    df_p["NUMERO_GUIA"] = df_p["NUMERO_GUIA"].astype(str).str.strip()

    df_p = df_p.merge(df_m2, left_on="NUMERO_GUIA", right_on="GUIA", how="left")

    # Estado Maestro = copia del estado guia maestro
    df_p["ESTADO MAESTRO"] = df_p["ESTADO GUIA MAESTRO"]

    # 1. Reconciliar campos numericos (VALOR_DECLARADO, FLETE, SEGURO): priorizar planilla si > 0, si no maestro
    for num_col in ["VALOR_DECLARADO", "FLETE", "SEGURO"]:
        col_m = f"{num_col}_M"
        if col_m in df_p.columns:
            v_p = pd.to_numeric(df_p.get(num_col, 0), errors="coerce").fillna(0)
            v_m = pd.to_numeric(df_p.get(col_m, 0), errors="coerce").fillna(0)
            df_p[num_col] = np.where(v_p > 0, v_p, v_m)
            df_p.drop(columns=[col_m], inplace=True, errors="ignore")

    # 2. Reconciliar MODALIDAD, REMITENTE, DESTINATARIO: priorizar planilla, completar con Maestro
    for str_col in ["MODALIDAD", "REMITENTE", "DESTINATARIO"]:
        col_m = f"{str_col}_M"
        if col_m in df_p.columns:
            s_p = df_p.get(str_col, pd.Series("", index=df_p.index)).fillna("").astype(str).str.strip()
            s_m = df_p.get(col_m, pd.Series("", index=df_p.index)).fillna("").astype(str).str.strip()
            df_p[str_col] = np.where(s_p != "", s_p, s_m)
            df_p.drop(columns=[col_m], inplace=True, errors="ignore")

    # 3. Reconciliar REGIONAL ORIGEN (Prioridad Maestro)
    ori_m = df_p.get("REGIONAL ORIGEN (M)", pd.Series("", index=df_p.index)).fillna("").astype(str).str.strip()
    ori_p = df_p.get("REGIONAL_ORIGEN", pd.Series("", index=df_p.index)).fillna("").astype(str).str.strip()
    res_ori = np.where(ori_m.isin(["", "nan", "None", "NO CRUZO"]), ori_p, ori_m)
    df_p["REGIONAL ORIGEN"] = [ _nombre_completo(x) if (x and str(x).upper() not in ["", "NAN", "NONE"]) else "BOGOTA - PRINCIPAL" for x in res_ori ]
    if "REGIONAL ORIGEN (M)" in df_p.columns:
        df_p.drop(columns=["REGIONAL ORIGEN (M)"], inplace=True)

    # 4. Reconciliar CIUDAD ORIGEN (Prioridad Maestro)
    ciu_ori_m = df_p.get("CIUDAD ORIGEN (M)", pd.Series("", index=df_p.index)).fillna("").astype(str).str.strip()
    ciu_ori_p = df_p.get("CIUDAD_ORIGEN", pd.Series("", index=df_p.index)).fillna("").astype(str).str.strip()
    res_ciu_ori = np.where(ciu_ori_m.isin(["", "nan", "None", "NO CRUZO"]), ciu_ori_p, ciu_ori_m)
    df_p["CIUDAD ORIGEN"] = np.where(pd.Series(res_ciu_ori).isin(["", "nan", "None"]), "BOGOTA", res_ciu_ori)
    if "CIUDAD ORIGEN (M)" in df_p.columns:
        df_p.drop(columns=["CIUDAD ORIGEN (M)"], inplace=True)

    # Rellenar vacios en campos clave del Maestro
    for col in ["FECHA CREACION GUIA (M)", "ESTADO GUIA MAESTRO", "ESTADO MAESTRO"]:
        if col in df_p.columns:
            df_p[col] = df_p[col].fillna("NO CRUZO")

    for col, default in [("PIEZAS_M", 0), ("TOTAL_MAESTRO", 0)]:
        df_p[col] = pd.to_numeric(df_p.get(col, pd.Series(dtype=float)),
                                   errors="coerce").fillna(default)

    # Si CIUDAD no esta en planilla, completar desde maestro
    if "CIUDAD" in df_p.columns and "CIUDAD_DESTINO_M" in df_p.columns:
        mask_sin_ciudad = df_p["CIUDAD"].astype(str).str.strip().isin(["", "nan", "None"])
        df_p.loc[mask_sin_ciudad, "CIUDAD"] = df_p.loc[mask_sin_ciudad, "CIUDAD_DESTINO_M"]

    # Si REGIONAL_DESTINO no esta en planilla, completar desde maestro
    if "REG_DESTINO_CORRECTA_M" in df_p.columns:
        if "REGIONAL_DESTINO" not in df_p.columns:
            df_p["REGIONAL_DESTINO"] = df_p["REG_DESTINO_CORRECTA_M"].apply(
                lambda x: _nombre_completo(str(x).strip()))
        else:
            mask_sin_dest = df_p["REGIONAL_DESTINO"].astype(str).str.strip().isin(["", "nan", "None"])
            df_p.loc[mask_sin_dest, "REGIONAL_DESTINO"] = df_p.loc[
                mask_sin_dest, "REG_DESTINO_CORRECTA_M"
            ].apply(lambda x: _nombre_completo(str(x).strip()))

    # Campos para escritura en la hoja de guias
    for col in ["ESTADO LEGALIZACION", "FECHA DEL ESTADO", "CIUDAD ORIGEN", "REGIONAL ORIGEN"]:
        if col not in df_p.columns:
            df_p[col] = ""
        df_p[col] = df_p[col].fillna("")

    if "GUIA" in df_p.columns:
        df_p.drop(columns=["GUIA"], inplace=True)

    cruzados = (df_p["ESTADO GUIA MAESTRO"] != "NO CRUZO").sum()
    print(f"      Cruce completado: {cruzados:,} guias encontradas en Maestro.")
    return df_p


# =============================================================================
#  SECCION 7 — DETECCION DE GUIAS SIN CARGUE
# =============================================================================

def detectar_guias_sin_cargue(maestro_dir: Path,
                               guias_ya_en_planilla: set) -> pd.DataFrame:
    """
    Identifica guias con destino a la Regional Sur (o Putumayo historico Abr-Jun)
    que NO tienen planilla de cargue de salida desde Bogota.
    """
    print("\n  [+] Detectando guias SIN CARGUE para Regional Sur / Putumayo...")
    df_m, nombre = _cargar_maestro(maestro_dir)
    if df_m.empty:
        print("      AVISO: Maestro no disponible. No se pueden detectar guias sin cargue.")
        return pd.DataFrame()

    df_m["GUIA"] = df_m["GUIA"].astype(str).str.strip()
    print(f"      Maestro: {nombre} ({len(df_m):,} registros)")

    if "REG DESTINO CORRECTA" not in df_m.columns:
        print("      AVISO: Columna 'REG DESTINO CORRECTA' no encontrada en Maestro.")
        return pd.DataFrame()

    # Ventana de control: despachos generados desde el 24-abr-2026 hasta el 30-jun-2026 para Putumayo
    fecha_inicio = pd.to_datetime(FECHA_INICIO_CONTROL)
    fecha_put_fin = pd.to_datetime(FECHA_PUTUMAYO_FIN)

    if "FECHA_DE_CREACION" in df_m.columns:
        fecha_crea = pd.to_datetime(df_m["FECHA_DE_CREACION"], errors="coerce", format="mixed")
    else:
        fecha_crea = pd.Series(pd.NaT, index=df_m.index)
    mask_ventana = fecha_crea >= fecha_inicio

    reg_dest = df_m["REG DESTINO CORRECTA"].astype(str).str.strip().str.upper()
    ciudad_dest = df_m.get(
        "CIUDAD_DESTINO", pd.Series("", index=df_m.index)
    ).astype(str).str.strip().str.upper()
    dept_dest = df_m.get(
        "DEPARTAMENTO", pd.Series("", index=df_m.index)
    ).astype(str).str.strip().str.upper()

    es_putumayo_m = ciudad_dest.isin(PUTUMAYO_CITIES_SET) | (dept_dest == "PUTUMAYO")

    # Regional Sur (Florencia, Pitalito, Garzon), EXCLUYENDO ciudades/dept de Putumayo
    mask_sur = (
        (reg_dest.isin([r.upper() for r in REGIONALES_SUR_MAESTRO]) | ciudad_dest.isin(CIUDADES_SUR))
        & ~es_putumayo_m
        & mask_ventana
    )

    # Putumayo historico: solo Abril-Junio 2026 (fecha <= 2026-06-30)
    ori_bogota = df_m.get("REGIONAL_ORIGEN", pd.Series(dtype=str)).astype(str).str.upper().str.contains("BOGOTA")
    mask_putumayo = es_putumayo_m & mask_ventana & (fecha_crea <= fecha_put_fin) & ori_bogota

    # Excluir anuladas/canceladas
    if "NOMEST" in df_m.columns:
        estado_norm = df_m["NOMEST"].astype(str).str.upper().str.strip()
        mask_no_anulado = ~estado_norm.isin([e.upper() for e in ESTADOS_ANULADO])
    else:
        mask_no_anulado = pd.Series([True] * len(df_m), index=df_m.index)

    # Excluir guias que ya tienen planilla de cargue
    mask_no_en_planilla = ~df_m["GUIA"].isin(guias_ya_en_planilla)

    df_sin = df_m[(mask_sur | mask_putumayo) & mask_no_anulado & mask_no_en_planilla].copy()

    if df_sin.empty:
        print("      No se encontraron guias sin cargue para la Regional Sur / Putumayo.")
        return pd.DataFrame()

    print(f"      Guias SIN CARGUE detectadas: {len(df_sin):,}")
    for reg in REGIONALES_SUR_MAESTRO:
        n = ((df_sin["REG DESTINO CORRECTA"].str.upper() == reg.upper()) & ~es_putumayo_m[df_sin.index]).sum()
        if n:
            print(f"        {reg:<12}: {n:,}")
    put_n = es_putumayo_m[df_sin.index].sum()
    if put_n:
        print(f"        PUTUMAYO  : {put_n:,}")

    def _g(col, default=""):
        return (df_sin[col].astype(str).str.strip()
                if col in df_sin.columns
                else pd.Series([default] * len(df_sin), index=df_sin.index))

    reg_dest_sin = df_sin["REG DESTINO CORRECTA"].astype(str).str.strip().str.upper()
    ciudad_dest_sin = df_sin.get("CIUDAD_DESTINO", pd.Series([""] * len(df_sin), index=df_sin.index))

    reg_destino_final = reg_dest_sin.map(lambda x: _nombre_completo(x))
    reg_destino_final = np.where(es_putumayo_m[df_sin.index], "PUTUMAYO - PRINCIPAL", reg_destino_final)

    df_resultado = pd.DataFrame({
        "CODIGO_CARGUE":          "SIN CARGUE",
        "FECHA_REGISTRO":         pd.to_datetime(_g("FECHA_DE_CREACION"), errors="coerce"),
        "REGIONAL_ORIGEN":        _g("REGIONAL_ORIGEN", "BOGOTA - PRINCIPAL"),
        "NOMBRE_ZONA":            "SIN CARGUE",
        "NUMERO_GUIA":            df_sin["GUIA"],
        "CIUDAD":                 ciudad_dest_sin.astype(str).str.strip(),
        "MODALIDAD":              _g("MODALIDAD"),
        "REMITENTE":              _g("REMITENTE_CLIENTE"),
        "DESTINATARIO":           _g("DESTINATARIO_CLIENTE"),
        "FLETE":                  pd.to_numeric(_g("FLETE"),           errors="coerce").fillna(0),
        "SEGURO":                 pd.to_numeric(_g("SEGURO"),          errors="coerce").fillna(0),
        "TOTAL":                  pd.to_numeric(_g("TOTAL"),           errors="coerce").fillna(0),
        "VALOR_DECLARADO":        pd.to_numeric(_g("VALOR_DECLARADO"), errors="coerce").fillna(0),
        "CANTIDAD":               0,
        "PESO":                   pd.to_numeric(_g("PESOS_FISICO"),    errors="coerce").fillna(0),
        "NOMBRE_ESTADO":          _g("NOMEST"),
        "PLACA":                  "SIN CARGUE",
        "CONDUCTOR_ASIGNADO":     "SIN CARGUE",
        "FECHA CREACION GUIA (M)": _g("FECHA_DE_CREACION"),
        "ESTADO GUIA MAESTRO":    _g("NOMEST"),
        "ESTADO MAESTRO":         _g("NOMEST"),
        "PIEZAS_M":               pd.to_numeric(_g("PIEZA"), errors="coerce").fillna(0),
        "TOTAL_MAESTRO":          pd.to_numeric(_g("TOTAL"), errors="coerce").fillna(0),
        "REGIONAL_DESTINO":       reg_destino_final,
        "TIENE_CARGUE":           "NO",
        "TIPO_TRAYECTO":          _g("TIPO TRAYECTO", "SIN CLASIFICAR"),
        "CATEGORIA_TRAYECTO":     "NACIONAL",
        "ESTADO LEGALIZACION":    _g("ESTADO_LEGALIZACION"),
        "FECHA DEL ESTADO":       _g("FECREGEST"),
        "CIUDAD ORIGEN":          _g("CIUDAD_ORIGEN", "BOGOTA"),
        "REGIONAL ORIGEN":        _g("REGIONAL_ORIGEN", "BOGOTA - PRINCIPAL").map(lambda x: _nombre_completo(x) if (x and str(x).upper() not in ["", "NAN", "NONE"]) else "BOGOTA - PRINCIPAL"),
        "DEPARTAMENTO":           _g("DEPARTAMENTO"),
        "CLIENTE":                _g("CLIENTE"),
        "COMERCIAL":              _g("COMERCIAL"),
    }, index=df_sin.index)

    return df_resultado.reset_index(drop=True)


# =============================================================================
#  SECCION 8 — FILTRADO REGIONAL SUR Y PUTUMAYO
# =============================================================================

def filtrar_datos_especificos(df_c, df_p):
    """
    Filtra cargues y guias de la Regional Sur y Bogota->Putumayo.
    """
    print("\n[3/4] Aplicando filtros especificos (SUR + PUTUMAYO)...")

    if df_c.empty or df_p.empty:
        print("      AVISO: Faltan datos para filtrar.")
        return df_c, df_p

    fecha_inicio = pd.to_datetime(FECHA_INICIO_CONTROL)
    fecha_put_fin = pd.to_datetime(FECHA_PUTUMAYO_FIN)

    # ── Cargues de la Regional Sur (Florencia, Pitalito, Garzon) ──
    mask_c = df_c.get("REGIONAL_RESPONSABLE", pd.Series(dtype=str)).isin(REGIONALES_SUR)
    df_c_filtered = df_c[mask_c].copy()
    if "FECREG" in df_c_filtered.columns:
        fc = pd.to_datetime(df_c_filtered["FECREG"], errors="coerce", format="mixed")
        df_c_filtered = df_c_filtered[fc.isna() | (fc >= fecha_inicio)].copy()
    print(f"      Cargues Regional Sur (Florencia, Pitalito, Garzon): {len(df_c_filtered)}")

    cargues_sur_ids = (
        df_c_filtered["CODCARGUE"].astype(str).str.strip().unique()
        if "CODCARGUE" in df_c_filtered.columns else np.array([], dtype=str)
    )

    # Excluir guias en estado ANULADO/CANCELADO
    if "NOMBRE_ESTADO" in df_p.columns:
        mask_anul = df_p["NOMBRE_ESTADO"].astype(str).str.upper().str.strip().isin([e.upper() for e in ESTADOS_ANULADO])
        df_p = df_p[~mask_anul].copy()
    if "ESTADO GUIA MAESTRO" in df_p.columns:
        mask_anul_m = df_p["ESTADO GUIA MAESTRO"].astype(str).str.upper().str.strip().isin([e.upper() for e in ESTADOS_ANULADO])
        df_p = df_p[~mask_anul_m].copy()

    # ── Guias CON cargue ──
    reg_ori = df_p.get("REGIONAL_ORIGEN", pd.Series("", index=df_p.index)).astype(str).str.strip().str.upper()
    ori_bogota = reg_ori.str.contains("BOGOTA", na=False)

    ciudad_p = df_p.get("CIUDAD", pd.Series("", index=df_p.index)).astype(str).str.strip().str.upper()
    dept_p = df_p.get("DEPARTAMENTO", pd.Series("", index=df_p.index)).astype(str).str.strip().str.upper()
    if "CIUDAD_DESTINO_M" in df_p.columns:
        ciudad_m = df_p["CIUDAD_DESTINO_M"].astype(str).str.strip().str.upper()
        ciudad_p = ciudad_p.where(ciudad_p != "", ciudad_m)

    es_putumayo_p = ciudad_p.isin(PUTUMAYO_CITIES_SET) | (dept_p == "PUTUMAYO")

    # 1) Putumayo (origen Bogota, ciudades o departamento del Putumayo)
    mask_p_putumayo = ori_bogota & es_putumayo_p
    # 2) Guias en cargues Sur (excluyendo Putumayo de la mascara general Sur)
    mask_p_sur = df_p.get("CODIGO_CARGUE", pd.Series(dtype=str)).astype(str).str.strip().isin(cargues_sur_ids) & ~es_putumayo_p
    # 3) Ciudades de la Regional Sur dentro de cargues de otras regionales (ej. NEIVA)
    es_nuestra_ciudad = ciudad_p.isin(CIUDADES_SUR) & ~es_putumayo_p
    if "REG_DESTINO_CORRECTA_M" in df_p.columns:
        reg_corr = df_p["REG_DESTINO_CORRECTA_M"].astype(str).str.strip().str.upper()
        es_nuestra_reg = reg_corr.isin([r.upper() for r in REGIONALES_SUR_MAESTRO]) & ~es_putumayo_p
    else:
        es_nuestra_reg = pd.Series(False, index=df_p.index)
    mask_p_ciudad = ori_bogota & (es_nuestra_ciudad | es_nuestra_reg)

    df_p_filtered = df_p[mask_p_putumayo | mask_p_sur | mask_p_ciudad].copy()

    # ── Ventana de control: desde el 24-abr-2026; Putumayo SOLO hasta el 30-jun-2026 ──
    if "FECHA_REGISTRO" in df_p_filtered.columns:
        fr = pd.to_datetime(df_p_filtered["FECHA_REGISTRO"], errors="coerce", format="mixed")
        ok_desde = fr.isna() | (fr >= fecha_inicio)
        es_put_filt = es_putumayo_p.reindex(df_p_filtered.index, fill_value=False)
        ok_put = fr.isna() | (fr <= fecha_put_fin)
        df_p_filtered = df_p_filtered[ok_desde & (ok_put | ~es_put_filt)].copy()

    df_p_filtered["TIENE_CARGUE"] = "SI"

    # ── Cargues extra vinculados (Putumayo + ciudades Sur en cargues NEIVA) ──
    if "CODIGO_CARGUE" in df_p_filtered.columns and "CODCARGUE" in df_c.columns:
        cargues_extra_ids = df_p_filtered[
            ~df_p_filtered["CODIGO_CARGUE"].astype(str).str.strip().isin(cargues_sur_ids)
        ]["CODIGO_CARGUE"].unique()
        if len(cargues_extra_ids) > 0:
            df_c_extra = df_c[df_c["CODCARGUE"].astype(str).str.strip().isin(cargues_extra_ids)].copy()
            df_c_filtered = pd.concat(
                [df_c_filtered, df_c_extra], ignore_index=True
            ).drop_duplicates(subset=["CODCARGUE"])
            print(f"      Cargues extra vinculados (Putumayo + ciudades Sur en cargues NEIVA): {len(df_c_extra)}")

    # ── Corregir REGIONAL_DESTINO para ciudades Sur enrutadas en cargues de
    #    otras regionales (ej. NEIVA): el destino correcto es la zona Sur propia ──
    df_p_filtered = _corregir_regional_destino(df_p_filtered)

    print(f"      Total guias CON cargue para el reporte: {len(df_p_filtered)}")
    return df_c_filtered, df_p_filtered


# =============================================================================
#  SECCION 9 — CONSTRUCCION DE HOJAS EXCEL
# =============================================================================

def _extraer_fecha_hora_mes(serie_fechas):
    """
    Dada una serie de fechas/strings, devuelve (fecha_limpia, hora, mes).
    - fecha_limpia: YYYY-MM-DD  (string)
    - hora:         HH:MM:SS    (string)
    - mes:          ENERO, FEBRERO, etc.  (string)
    Para valores invalidos devuelve "", "00:00:00", "".
    """
    dt = pd.to_datetime(serie_fechas, errors="coerce", format="mixed")
    fecha_limpia = dt.dt.strftime("%Y-%m-%d").fillna("")
    hora         = dt.dt.strftime("%H:%M:%S").where(dt.notna(), "00:00:00")
    mes          = dt.dt.month.map(MESES_ES).fillna("")
    return fecha_limpia, hora, mes


def _seccion_cargues(ws, df_c: pd.DataFrame, row_ini: int) -> int:
    """Escribe la seccion de detalle de cargues en la hoja dada."""
    ws.cell(row=row_ini, column=1, value="DETALLE DE CARGUES").font = _font(bold=True, size=11)
    ws.cell(row=row_ini, column=1).fill = _fill(C["hdr_light"])

    if df_c.empty:
        ws.cell(row=row_ini + 1, column=1, value="Sin cargues operativos.")
        return row_ini + 2

    col_map = {
        "CODCARGUE":          "CARGUE",
        "FECREG":             "FECHA REG.",
        "NOMREG":             "REG. ORIGEN",
        "NOMZONA":            "ZONA",
        "TIPO_TRAYECTO":      "TIPO TRAYECTO",
        "CATEGORIA_TRAYECTO": "CATEGORIA",
        "NOMEST":             "ESTADO",
        "PLACA":              "PLACA",
        "CONDUCTOR":          "CONDUCTOR",
        "CANTGUIAS":          "GUIAS",
        "CANTENTREGAS":       "ENTREGADAS",
        "CANTPEDIDAS":        "PENDIENTES",
        "CANTNOVEDAD":        "NOVEDAD",
        "SIN_GESTION":        "SIN GESTION",
        "TOTALCARGUE":        "VALOR ($)",
    }
    df_out = df_c[[c for c in col_map if c in df_c.columns]].rename(columns=col_map).copy()
    if "ENTREGADAS" in df_out.columns and "GUIAS" in df_out.columns:
        df_out["% EFECTIVIDAD"] = _pct(df_out["ENTREGADAS"], df_out["GUIAS"])

    _encabezado(ws, list(df_out.columns), row=row_ini + 1, color=C["hdr_mid"])
    last = _escribir_df(
        ws, df_out, start_row=row_ini + 2,
        pct_ef_cols=["% EFECTIVIDAD"],
        cat_col="CATEGORIA",
        num_cols=["GUIAS", "ENTREGADAS", "PENDIENTES", "NOVEDAD", "SIN GESTION", "VALOR ($)"],
        fecha_cols=["FECHA REG."],
    )
    return last + 1


def _seccion_guias_detalle(ws, df_p: pd.DataFrame, df_sin: pd.DataFrame,
                            df_c: pd.DataFrame, row_ini: int):
    """
    Escribe la hoja GUIAS DETALLE con exactamente 31 columnas en el orden fijo:
    CARGUE, NUMERO GUIA, FECHA CARGUE, HORA, MES, ORIGEN, DESTINO, CIUDAD,
    CATEGORIA, TIPO TRAYECTO, ESTADO CARGUE, ESTADO GUIA, MODALIDAD, REMITENTE,
    DESTINATARIO, FLETE ($), SEGURO ($), TOTAL ($), V. DECLARADO, CANT, PESO (kg),
    PLACA, CONDUCTOR, FECHA CREACION GUIA (M), ESTADO MAESTRO, PIEZAS (M),
    TOTAL VALOR (M), ESTADO LEGALIZACION, FECHA DEL ESTADO, CIUDAD ORIGEN,
    REGIONAL ORIGEN
    """
    ws.cell(row=row_ini, column=1,
            value="DETALLE DE GUIAS CONSOLIDADAS (CON Y SIN CARGUE)").font = _font(bold=True, size=11)
    ws.cell(row=row_ini, column=1).fill = _fill(C["hdr_light"])

    # ── Enriquecer guias CON cargue con datos del cargue (placa, conductor, estado, categoria) ──
    df_con = df_p.copy()
    if not df_c.empty and "CODCARGUE" in df_c.columns and "CODIGO_CARGUE" in df_p.columns:
        cols_cargue = [c for c in ["CODCARGUE", "NOMEST", "TIPO_TRAYECTO",
                                    "CATEGORIA_TRAYECTO", "PLACA", "CONDUCTOR"]
                       if c in df_c.columns]
        cargues_info = df_c[cols_cargue].rename(columns={
            "NOMEST":             "EST_CARGUE",
            "TIPO_TRAYECTO":      "TIP_CARGUE",
            "CATEGORIA_TRAYECTO": "CAT_CARGUE",
            "PLACA":              "PLA_CARGUE",
            "CONDUCTOR":          "CON_CARGUE",
        })
        df_con = df_p.merge(cargues_info, left_on="CODIGO_CARGUE",
                             right_on="CODCARGUE", how="left")

    def _safe(df_in, col, default=""):
        return (df_in[col] if col in df_in.columns
                else pd.Series([default] * len(df_in), index=df_in.index))

    def construir_fila(df_in: pd.DataFrame, es_sin: bool) -> pd.DataFrame:
        """Construye el DataFrame con las 31+1 columnas (la +1 es _TIENE_CARGUE para coloreado)."""
        if df_in.empty:
            return pd.DataFrame()

        fecha_cargue, hora, mes = _extraer_fecha_hora_mes(_safe(df_in, "FECHA_REGISTRO"))

        out = pd.DataFrame(index=df_in.index)
        out["CARGUE"]       = _safe(df_in, "CODIGO_CARGUE", "SIN CARGUE")
        out["NUMERO GUIA"]  = _safe(df_in, "NUMERO_GUIA")
        out["FECHA CARGUE"] = fecha_cargue
        out["HORA"]         = hora
        out["MES"]          = mes
        out["ORIGEN"]       = "BOGOTA"
        out["DESTINO"]      = _safe(df_in, "REGIONAL_DESTINO")

        # CIUDAD: usar del cruce con maestro si disponible; si no, de la planilla
        ciudad_m = _safe(df_in, "CIUDAD_DESTINO_M", "")
        ciudad_p = _safe(df_in, "CIUDAD", "")
        out["CIUDAD"] = np.where(
            ciudad_m.astype(str).str.strip().isin(["", "nan"]), ciudad_p, ciudad_m
        )

        if es_sin:
            out["CATEGORIA"]    = "NACIONAL"
            out["TIPO TRAYECTO"]= _safe(df_in, "TIPO_TRAYECTO", "SIN CLASIFICAR")
            out["ESTADO CARGUE"]= "SIN CARGUE"
        else:
            out["CATEGORIA"]    = _safe(df_in, "CAT_CARGUE", "SIN CLASIFICAR").fillna("SIN CLASIFICAR")
            out["TIPO TRAYECTO"]= _safe(df_in, "TIP_CARGUE", "SIN CLASIFICAR").fillna("SIN CLASIFICAR")
            out["ESTADO CARGUE"]= _safe(df_in, "EST_CARGUE", "").fillna("")

        # Estado Guia: del Maestro (NOMEST), fallback a Planilla si no cruzo
        est_m = _safe(df_in, "ESTADO GUIA MAESTRO", "")
        est_p = _safe(df_in, "NOMBRE_ESTADO", "")
        out["ESTADO GUIA"] = np.where(est_m.astype(str).str.strip().isin(["", "nan", "None", "NO CRUZO"]), est_p, est_m)

        out["MODALIDAD"]    = _safe(df_in, "MODALIDAD")
        out["REMITENTE"]    = _safe(df_in, "REMITENTE")
        out["DESTINATARIO"] = _safe(df_in, "DESTINATARIO")

        out["FLETE ($)"]    = pd.to_numeric(_safe(df_in, "FLETE",   0), errors="coerce").fillna(0)
        out["SEGURO ($)"]   = pd.to_numeric(_safe(df_in, "SEGURO",  0), errors="coerce").fillna(0)
        out["TOTAL ($)"]    = pd.to_numeric(_safe(df_in, "TOTAL",   0), errors="coerce").fillna(0)
        out["V. DECLARADO"] = pd.to_numeric(_safe(df_in, "VALOR_DECLARADO", 0), errors="coerce").fillna(0)

        if es_sin:
            out["CANT"]       = 0
            out["PLACA"]      = "SIN CARGUE"
            out["CONDUCTOR"]  = "SIN CARGUE"
        else:
            out["CANT"]       = pd.to_numeric(_safe(df_in, "CANTIDAD", 0), errors="coerce").fillna(0)
            placa_c = _safe(df_in, "PLA_CARGUE", "").fillna("").astype(str)
            placa_p = _safe(df_in, "PLACA", "").fillna("").astype(str)
            out["PLACA"]      = np.where(placa_c.str.strip() != "", placa_c, placa_p)
            cond_c  = _safe(df_in, "CON_CARGUE", "").fillna("").astype(str)
            cond_p  = _safe(df_in, "CONDUCTOR_ASIGNADO", "").fillna("").astype(str)
            out["CONDUCTOR"]  = np.where(cond_c.str.strip() != "", cond_c, cond_p)

        out["PESO (kg)"]    = pd.to_numeric(_safe(df_in, "PESO", 0), errors="coerce").fillna(0)

        # Del Maestro
        out["FECHA CREACION GUIA (M)"] = _safe(df_in, "FECHA CREACION GUIA (M)")
        out["ESTADO MAESTRO"]          = _safe(df_in, "ESTADO MAESTRO")
        out["PIEZAS (M)"]              = pd.to_numeric(_safe(df_in, "PIEZAS_M",    0), errors="coerce").fillna(0)
        out["TOTAL VALOR (M)"]         = pd.to_numeric(_safe(df_in, "TOTAL_MAESTRO", 0), errors="coerce").fillna(0)
        out["ESTADO LEGALIZACION"]     = _safe(df_in, "ESTADO LEGALIZACION")
        out["FECHA DEL ESTADO"]        = _safe(df_in, "FECHA DEL ESTADO")
        out["CIUDAD ORIGEN"]           = _safe(df_in, "CIUDAD ORIGEN", "BOGOTA")
        out["REGIONAL ORIGEN"]         = _safe(df_in, "REGIONAL ORIGEN", "BOGOTA - PRINCIPAL")

        # Columna interna para coloreo (no se escribe en Excel)
        out["_TIENE_CARGUE"] = "NO" if es_sin else "SI"
        return out

    df_con_out = construir_fila(df_con, es_sin=False)
    df_sin_out = construir_fila(df_sin, es_sin=True) if not df_sin.empty else pd.DataFrame()

    if not df_sin_out.empty:
        # Alinear columnas
        todas_cols = list(df_con_out.columns)
        for col in df_sin_out.columns:
            if col not in todas_cols:
                todas_cols.append(col)
        for df_tmp in [df_con_out, df_sin_out]:
            for col in todas_cols:
                if col not in df_tmp.columns:
                    df_tmp[col] = ""
        df_con_out = df_con_out[todas_cols]
        df_sin_out = df_sin_out[todas_cols]
        df_total = pd.concat([df_con_out, df_sin_out], ignore_index=True)
    else:
        df_total = df_con_out

    if df_total.empty:
        ws.cell(row=row_ini + 1, column=1, value="Sin guias consolidadas.")
        return

    # Columnas para Excel (excluyendo la interna _TIENE_CARGUE)
    cols_excel = [c for c in df_total.columns if c != "_TIENE_CARGUE"]
    ci_total   = {name: i + 1 for i, name in enumerate(df_total.columns)}

    _encabezado(ws, cols_excel, row=row_ini + 1, color=C["hdr_mid"])

    num_cols_g  = ["FLETE ($)", "SEGURO ($)", "TOTAL ($)", "V. DECLARADO",
                   "CANT", "PESO (kg)", "PIEZAS (M)", "TOTAL VALOR (M)"]
    fecha_cols_g = ["FECHA CARGUE", "FECHA CREACION GUIA (M)", "FECHA DEL ESTADO"]

    for i, tup in enumerate(df_total.itertuples(index=False), row_ini + 2):
        row_vals = list(tup)
        tiene_cargue = str(row_vals[ci_total["_TIENE_CARGUE"] - 1]).strip().upper()

        row_bg = None
        if tiene_cargue == "NO":
            row_bg = C["sin_cargue"]
        else:
            cat = str(row_vals[ci_total.get("CATEGORIA", 0) - 1] if "CATEGORIA" in ci_total else "")
            row_bg = CAT_BG.get(cat)

        for j, col_name in enumerate(cols_excel, 1):
            val = row_vals[ci_total[col_name] - 1]
            try:
                if pd.isnull(val):
                    val = ""
            except Exception:
                pass
            c = ws.cell(row=i, column=j, value=val)
            c.font      = _font(size=9)
            c.alignment = _aln("center")
            c.border    = B_THIN
            if row_bg:
                c.fill = _fill(row_bg)
            if col_name in num_cols_g:
                c.number_format = "#,##0"
            elif col_name in fecha_cols_g:
                c.number_format = "YYYY-MM-DD"


def _hoja_guias_sin_cargue(ws, df_sin: pd.DataFrame, fecha: str):
    """Hoja exclusiva con las guias sin planilla de cargue al Sur o Putumayo."""
    ws.sheet_view.showGridLines = False
    n = len(df_sin) if not df_sin.empty else 0
    _banner(ws,
            f"GUIAS SIN CARGUE - REGIONAL SUR Y PUTUMAYO  -  {fecha}  [{n} guias]",
            ncols=21, row=1, color="C00000")

    if df_sin.empty:
        ws.cell(row=3, column=1, value="No se detectaron guias sin cargue.")
        return

    col_map = {
        "NUMERO_GUIA":            "NUMERO GUIA",
        "FECHA CREACION GUIA (M)":"FECHA CREACION",
        "REGIONAL_DESTINO":       "REGIONAL DESTINO",
        "CIUDAD":                 "CIUDAD DESTINO",
        "DEPARTAMENTO":           "DEPARTAMENTO",
        "REGIONAL ORIGEN":        "REGIONAL ORIGEN",
        "ESTADO MAESTRO":         "ESTADO GUIA",
        "MODALIDAD":              "MODALIDAD",
        "REMITENTE":              "REMITENTE",
        "DESTINATARIO":           "DESTINATARIO",
        "CLIENTE":                "CLIENTE",
        "COMERCIAL":              "COMERCIAL",
        "PIEZAS_M":               "PIEZAS",
        "PESO":                   "PESO (kg)",
        "FLETE":                  "FLETE ($)",
        "SEGURO":                 "SEGURO ($)",
        "TOTAL":                  "TOTAL ($)",
        "VALOR_DECLARADO":        "V. DECLARADO",
        "ESTADO LEGALIZACION":    "ESTADO LEGALIZACION",
        "FECHA DEL ESTADO":       "FECHA ESTADO",
        "CIUDAD ORIGEN":          "CIUDAD ORIGEN",
    }
    df_out = df_sin[[c for c in col_map if c in df_sin.columns]].rename(columns=col_map).copy()

    # Nota explicativa
    ncols_out = max(len(df_out.columns), 1)
    ws.merge_cells(f"A2:{get_column_letter(ncols_out)}2")
    c2 = ws["A2"]
    c2.value = ("ATENCION: Estas guias tienen destino a la Regional Sur o Putumayo (hasta Junio 2026) "
                "pero NO tienen planilla de cargue de salida desde Bogota. "
                "Campos de cargue (N Cargue, Placa, Conductor) son genericos. "
                "Fecha de referencia = Fecha de Creacion de la guia en el Maestro.")
    c2.font      = _font(bold=True, color="7B2D00", size=9)
    c2.fill      = _fill(C["amarillo"])
    c2.alignment = _aln("left", wrap=True)
    ws.row_dimensions[2].height = 32

    _encabezado(ws, list(df_out.columns), row=3, color="C00000")
    _escribir_df(
        ws, df_out, start_row=4,
        num_cols=["PIEZAS", "PESO (kg)", "FLETE ($)", "SEGURO ($)", "TOTAL ($)", "V. DECLARADO"],
        fecha_cols=["FECHA CREACION", "FECHA ESTADO"],
    )
    _col_width(ws)


def _hoja_resumen(ws, df_c: pd.DataFrame, df_p: pd.DataFrame,
                  df_sin: pd.DataFrame, fecha: str):
    """Hoja RESUMEN con: resumen por categoria, totales con/sin cargue, tabla por regional."""
    ws.sheet_view.showGridLines = False
    n_sin = len(df_sin) if not df_sin.empty else 0
    n_con = len(df_p)

    _banner(ws, f"REPORTE CONSOLIDADO SUR Y PUTUMAYO  v4.1  -  {fecha}", ncols=12, row=1)

    # --- Bloque 1: cargues por categoria ---
    row = 3
    ws.cell(row=row, column=1, value="RESUMEN DE CARGUES POR CATEGORIA").font = _font(bold=True, size=11)
    ws.cell(row=row, column=1).fill = _fill(C["hdr_light"])
    row += 1

    if not df_c.empty and "CATEGORIA_TRAYECTO" in df_c.columns:
        g = (df_c.groupby("CATEGORIA_TRAYECTO", observed=True)
             .agg(
                 CARGUES=("CODCARGUE", "count"),
                 GUIAS=("CANTGUIAS", "sum"),
                 ENTREGADAS=("CANTENTREGAS", "sum"),
                 PENDIENTES=("CANTPEDIDAS", "sum"),
                 NOVEDAD=("CANTNOVEDAD", "sum"),
                 SIN_GESTION=("SIN_GESTION", "sum"),
             )
             .reset_index())
        g["PCT_E"] = _pct(g["ENTREGADAS"], g["GUIAS"])
        g["PCT_P"] = _pct(g["PENDIENTES"], g["GUIAS"])
        g = g.rename(columns={
            "CATEGORIA_TRAYECTO": "CATEGORIA",
            "SIN_GESTION":        "SIN GESTION",
            "PCT_E":              "% EFECTIVIDAD",
            "PCT_P":              "% PENDIENTE",
        })
        _encabezado(ws, list(g.columns), row=row, color=C["hdr_mid"])
        row = _escribir_df(
            ws, g, start_row=row + 1,
            pct_ef_cols=["% EFECTIVIDAD"], pct_pend_cols=["% PENDIENTE"],
            cat_col="CATEGORIA",
            num_cols=["CARGUES", "GUIAS", "ENTREGADAS", "PENDIENTES", "NOVEDAD", "SIN GESTION"],
        ) + 2
    else:
        ws.cell(row=row, column=1, value="Sin cargues con categoria disponible.")
        row += 2

    # --- Bloque 2: guias con/sin cargue ---
    ws.cell(row=row, column=1, value="RESUMEN GUIAS CON Y SIN CARGUE").font = _font(bold=True, size=11)
    ws.cell(row=row, column=1).fill = _fill(C["hdr_light"])
    row += 1

    df_res = pd.DataFrame([
        {"TIPO": "Guias CON Cargue (Regional Sur + Putumayo)",
         "CANTIDAD": n_con,
         "NOTA": "Registradas en Planilla_Global con cargue asignado"},
        {"TIPO": "Guias SIN Cargue (Sur + Putumayo historico)",
         "CANTIDAD": n_sin,
         "NOTA": "Detectadas en Maestro sin planilla fisica de cargue al Sur"},
        {"TIPO": "TOTAL GUIAS EN REPORTE",
         "CANTIDAD": n_con + n_sin,
         "NOTA": "Panorama completo de despachos hacia la Regional Sur y Putumayo"},
    ])
    _encabezado(ws, list(df_res.columns), row=row, color=C["hdr_mid"])
    last = _escribir_df(ws, df_res, start_row=row + 1, num_cols=["CANTIDAD"])
    row = last + 2

    # --- Bloque 3: sin cargue por regional ---
    if not df_sin.empty and "REGIONAL_DESTINO" in df_sin.columns:
        ws.cell(row=row, column=1,
                value="GUIAS SIN CARGUE POR REGIONAL DESTINO").font = _font(bold=True, size=11)
        ws.cell(row=row, column=1).fill = _fill(C["sin_cargue"])
        row += 1

        g_sin = (df_sin.groupby("REGIONAL_DESTINO", observed=True)
                 .agg(
                     GUIAS=("NUMERO_GUIA", "count"),
                     FLETE=("FLETE", "sum"),
                     TOTAL=("TOTAL", "sum"),
                     PIEZAS=("PIEZAS_M", "sum"),
                 )
                 .reset_index()
                 .rename(columns={
                     "REGIONAL_DESTINO": "REGIONAL",
                     "FLETE": "FLETE ($)",
                     "TOTAL": "TOTAL ($)",
                 }))
        _encabezado(ws, list(g_sin.columns), row=row, color="C00000")
        _escribir_df(ws, g_sin, start_row=row + 1,
                     num_cols=["GUIAS", "FLETE ($)", "TOTAL ($)", "PIEZAS"])

    _col_width(ws)


# =============================================================================
#  SECCION 10 — GENERACION DEL EXCEL FINAL
# =============================================================================

def generar_reporte_unificado(df_c: pd.DataFrame, df_p: pd.DataFrame,
                               df_sin: pd.DataFrame, fecha: str,
                               output_path: Path):
    """Genera el archivo Excel final con 4 hojas."""
    print(f"\n[4/4] Escribiendo archivo consolidado final...")
    wb = Workbook()
    wb.remove(wb.active)

    n_sin = len(df_sin) if not df_sin.empty else 0

    # Hoja 1: RESUMEN
    ws_res = wb.create_sheet("RESUMEN")
    _hoja_resumen(ws_res, df_c, df_p, df_sin, fecha)

    # Hoja 2: CARGUES
    ws_c = wb.create_sheet("CARGUES")
    ws_c.sheet_view.showGridLines = False
    _banner(ws_c, f"DETALLE DE CARGUES CONSOLIDADOS  v4.1  -  {fecha}", ncols=16)
    _seccion_cargues(ws_c, df_c, row_ini=3)
    _col_width(ws_c)

    # Hoja 3: GUIAS DETALLE (31 columnas fijas, con y sin cargue)
    ws_g = wb.create_sheet("GUIAS DETALLE")
    ws_g.sheet_view.showGridLines = False
    _banner(ws_g,
            f"GUIAS CON CARGUE + SIN CARGUE  |  SUR + PUTUMAYO  v4.1  -  {fecha}  "
            f"[{len(df_p)} con cargue  |  {n_sin} sin cargue]",
            ncols=31)
    _seccion_guias_detalle(ws_g, df_p, df_sin, df_c, row_ini=3)
    _col_width(ws_g)

    # Hoja 4: GUIAS SIN CARGUE (solo huerfanas)
    ws_sc = wb.create_sheet("GUIAS SIN CARGUE")
    _hoja_guias_sin_cargue(ws_sc, df_sin, fecha)

    wb.save(str(output_path))
    print(f"      OK -> {output_path.name}")
    print(f"      Hojas: RESUMEN | CARGUES | GUIAS DETALLE | GUIAS SIN CARGUE")


# =============================================================================
#  SECCION 11 — MAIN
# =============================================================================

def main():
    fecha = datetime.today().strftime("%Y-%m-%d")

    cargues_files  = []
    planilla_files = []

    if _ARCHIVOS_CARGUES.exists():
        cargues_files = sorted(
            p for p in _ARCHIVOS_CARGUES.glob("Cargues_*.xls*")
            if p.suffix.lower() not in (".crdownload", ".tmp", ".part")
        )
        planilla_files = sorted(
            p for p in _ARCHIVOS_CARGUES.glob("Planilla_Global_*.xls*")
            if p.suffix.lower() not in (".crdownload", ".tmp", ".part")
        )

    if not cargues_files or not planilla_files:
        print("ERROR: No se encontraron los archivos requeridos en:")
        print(f"  {_ARCHIVOS_CARGUES}")
        print("  Patrones esperados:  Cargues_*.xls   y   Planilla_Global_*.xls")
        input("\nPresiona ENTER para cerrar...")
        return

    if not _TRAYECTOS_DEFAULT.exists():
        print(f"ERROR: No se encontro la matriz de trayectos en: {_TRAYECTOS_DEFAULT}")
        input("\nPresiona ENTER para cerrar...")
        return

    _SALIDA_CARGUES.mkdir(parents=True, exist_ok=True)

    print(f"\n{'=' * 68}")
    print(f"  PIPELINE SUR + PUTUMAYO  v4.1  -  {fecha}")
    print(f"  Archivos Cargues  : {len(cargues_files)}")
    for f in cargues_files:
        print(f"    . {f.name}")
    print(f"  Archivos Planilla : {len(planilla_files)}")
    for f in planilla_files:
        print(f"    . {f.name}")
    print(f"  Maestro en        : {_SALIDA_UNIFICAR}")
    print(f"  Salida ->          {_SALIDA_CARGUES}")
    print(f"{'=' * 68}")

    # Paso 1 — Carga paralela
    print("\n[1/4] Leyendo y consolidando archivos (todas las hojas)...")
    t0 = datetime.now()
    lista_cargues  = [str(p) for p in cargues_files]
    lista_planilla = [str(p) for p in planilla_files]
    df_c, df_p, df_t = cargar_archivos(lista_cargues, lista_planilla, str(_TRAYECTOS_DEFAULT))
    t_carga = (datetime.now() - t0).total_seconds()
    print(f"\n      Carga en {t_carga:.1f}s | Cargues brutos: {len(df_c):,} | Planilla bruta: {len(df_p):,}")

    # Paso 2 — Limpieza y enriquecimiento
    print("\n[2/4] Limpiando, deduplicando y enriqueciendo...")
    df_c = limpiar_cargues(df_c)
    df_p = limpiar_planilla(df_p)
    df_p = cruzar_con_maestro(df_p, _SALIDA_UNIFICAR)
    df_c, df_p = enriquecer(df_c, df_p, df_t)

    # Paso 3 — Filtros regionales
    df_c, df_p = filtrar_datos_especificos(df_c, df_p)

    # Paso 4 — Deteccion de guias sin cargue al Sur
    print("\n[3b/4] Detectando guias sin planilla de cargue al Sur / Putumayo...")
    guias_ya_con_cargue = (
        set(df_p["NUMERO_GUIA"].astype(str).str.strip().tolist())
        if "NUMERO_GUIA" in df_p.columns else set()
    )
    df_sin_cargue = detectar_guias_sin_cargue(_SALIDA_UNIFICAR, guias_ya_con_cargue)

    # Paso 5 — Generar reporte
    output_file = _SALIDA_CARGUES / f"Reporte_SUR_PUTUMAYO_v4_{fecha}.xlsx"
    generar_reporte_unificado(df_c, df_p, df_sin_cargue, fecha, output_file)

    n_sin = len(df_sin_cargue) if not df_sin_cargue.empty else 0
    print(f"\n{'=' * 68}")
    print(f"  PROCESO COMPLETADO  v4.1")
    print(f"  Guias CON cargue  : {len(df_p):,}")
    print(f"  Guias SIN cargue  : {n_sin:,}  <- NUEVO en v4")
    print(f"  Total panorama    : {len(df_p) + n_sin:,}")
    print(f"  Archivo final     : {output_file}")
    print(f"{'=' * 68}\n")
    input("\nPresiona ENTER para cerrar...")


if __name__ == "__main__":
    main()
