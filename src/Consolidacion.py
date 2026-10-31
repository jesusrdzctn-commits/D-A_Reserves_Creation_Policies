"""RDA Pólizas D&A — Consolidación: builds the 'Variaciones' sheet (five blocks grouped by CHANNEL, the 'CT' of 'Base').
Optionally pastes the 'RESERVAS MANUALES' templates into 'Base' first (Reservas_Manuales.py), in memory so the matrices see them.
Amounts in B and D are VALUES summed in pandas; everything else is a LIVE FORMULA, like the hand-made sheet.
Then hands the file to Comparativos.py for the PivotTables (needs Excel; optional, so a failure keeps the saved sheet).
"""

import os
import re
import shutil
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from utils import carpeta_output_desde_input, claves_mes_actual_y_anterior, etiqueta_mes_es, parse_clave_mes

# --- Sheets and source columns ---
HOJA_BASE        = "Base"
HOJA_VARIACIONES = "Variaciones"

# Columns read from 'Base'. Matched by NORMALISED name (trimmed, case-folded),
# because the real file ships 'Importe ' with a trailing space.
COL_MES       = "Mes"
COL_AGRUPADOR = "Agrupador"
COL_IMPORTE   = "Importe"
# Channel of each row = the column the five blocks are grouped by: 'CT' (ECOMM/HD/...), not 'Canal Venta' (the CANAL cell as it came).
COL_CANAL     = "CT"
# Read only to notice rows of an OLDER paste (channel in 'Canal Venta', 'CT' blank).
COL_CANAL_VENTA = "Canal Venta"

# The 'Base' header row is detected by content (padding rows above it have changed); this is the fallback.
FILA_ENCABEZADO_BASE_DEFAULT = 7


# --- The five blocks ---
# Grouped by CHANNEL ('CT' of 'Base'), also the title of each 'RESERVAS MANUALES <CANAL>' template; this is the whole
# configuration point. titulo: block title in 'Variaciones'. canal: name written into 'CT' on paste.
# canales: every spelling that belongs to the block, matched by words (accent/case/plural-insensitive, see resolver_grupo).
GRUPOS = [
    {"titulo": "Smkt Regionales", "canal": "REGIONALES", "canales": ["REGIONALES"]},
    {"titulo": "Proximity",       "canal": "PROXIMITY",  "canales": ["PROXIMITY"]},
    {"titulo": "Smkt Nacionales", "canal": "NACIONALES", "canales": ["NACIONALES"]},
    {"titulo": "ECOMMERCE",       "canal": "ECOMM",      "canales": ["ECOMM", "ECOMMERCE", "E-COMMERCE"]},
    {"titulo": "HD",              "canal": "HD",         "canales": ["HD"]},
]


# --- Concept catalogue (row labels of every matrix) ---
# The stakeholder's fixed list; it does NOT match 'Agrupador' word for word, so emparejar_concepto() matches them. Row order follows it.
CONCEPTOS_VARIACIONES = [
    "RENTA DE ESPACIOS",
    "DESCUENTO LOGÍSTICO",
    "STALES",
    "PLAN DE CRECIMIENTO",
    "DESCUENTO POR NO DEVOLUCIÓN",
    "DESCUENTO PROMOCIONAL",
    "APERTURA / APOYOS",
    "DESCUENTO FIJO",
    "DESCUENTO ESPECIAL",
    "PUBLICIDAD",
    "DIFERENCIA EN PRECIO",
]

# Words carrying no meaning for matching ('RENTA DE ESPACIOS' vs 'RENTAS DE
# ESPACIO' must match; 'DE' must not be what makes them match).
PALABRAS_VACIAS = {"DE", "DEL", "LA", "LAS", "EL", "LOS", "Y", "POR", "EN",
                   "A", "AL", "CON", "PARA"}

# Similarity threshold for the THIRD pass only (typos). The only true match reaching it scores 0.95 and the best false
# candidate ('DESCUENTO DE LÍNEA' vs 'DESCUENTO FIJO') 0.72, so 0.86 sits in the gap; lower it and 'DESCUENTO ...' concepts collide.
UMBRAL_SIMILITUD = 0.86

# An 'Agrupador' matching nothing becomes its own bottom row (raw name) and is reported in 'conceptos_no_mapeados': no money leaves a Grand Total.
AGREGAR_NO_MAPEADOS = True

# False (the stakeholder's call) prints only concepts with activity in the block; True prints all 11 so the five blocks line up.
MOSTRAR_CONCEPTOS_SIN_MOVIMIENTO = False


# --- Layout of one block (1-based column indexes) ---
COL_CONCEPTO = 1   # A
COL_MES_ANT  = 2   # B
COL_SEP      = 3   # C  (narrow visual separator, always empty)
COL_MES_ACT  = 4   # D
COL_PCT      = 5   # E
COL_BW       = 6   # F
COL_VAR      = 7   # G
COL_INDEX    = 8   # H

FILA_INICIAL   = 1   # first block starts here
FILAS_ENTRE_BLOQUES = 1   # blank rows between one block's total and the next

ANCHOS_COLUMNA = {
    "A": 31.33, "B": 14.55, "C": 2.66, "D": 14.55,
    "E": 7.44,  "F": 14.55, "G": 11.55, "H": 8.0,
}

# Add the variance / index formulas to the 'Grand Total' row too. The hand-made
# sheet leaves G and H empty on that row; set to False to match it exactly.
TOTAL_INCLUYE_VARIACION = True

# Also build the two PivotTable sheets once 'Variaciones' is saved (needs Excel + pywin32; False to skip).
CONSTRUIR_COMPARATIVOS = True


# --- Styles — lifted from the stakeholder's sheet (theme colours resolved to RGB) ---
GRIS_CREACIONES = "BFBFBF"   # 'Creaciones 2026' banner
AZUL_TITULO     = "DCE6F2"   # group title
TEAL_HEADER     = "31859C"   # headers 'Concepto' and previous month  (A, B)
VERDE_HEADER    = "00B050"   # headers current month, '%', 'B/W'      (D, E, F)
NARANJA_HEADER  = "E46C0A"   # headers '% VS <mes>' and 'INDEX'       (G, H)

FUENTE = "Calibri"
TAM    = 10

CONTABLE_2  = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'
CONTABLE_1  = '_-* #,##0.0_-;\\-* #,##0.0_-;_-* "-"??_-;_-@_-'
CONTABLE_0  = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'
PORCENTAJE  = "0.0%"
PORCENTAJE_TOTAL = "0%"

_BLANCO_NEGRITA = Font(name=FUENTE, size=TAM, bold=True, color="FFFFFF")
_NEGRO_NEGRITA  = Font(name=FUENTE, size=TAM, bold=True, color="000000")
_NEGRO_NORMAL   = Font(name=FUENTE, size=TAM, color="000000")
_CENTRADO       = Alignment(horizontal="center")
_IZQUIERDA      = Alignment(horizontal="left")

_BORDE_TOTAL = Border(top=Side(style="thin"), bottom=Side(style="double"))
_BORDE_TOTAL_CONCEPTO = Border(top=Side(style="double"))


def _relleno(rgb):
    return PatternFill(fill_type="solid", start_color=rgb, end_color=rgb)


# --- Helpers ---
def _es_nulo(valor):
    """True for None and pandas' NaN/NA: str(NaN) is 'nan', so an empty 'Mes' or 'Agrupador' would slip past the filters."""
    if valor is None:
        return True
    try:
        return bool(pd.isna(valor))
    except (TypeError, ValueError):
        return False   # not a scalar (cannot happen for a cell value)


def _normalizar(texto):
    """Trim and upper-case, tolerating None/NaN. Used to match headers and values."""
    return "" if _es_nulo(texto) else str(texto).strip().upper()


def _parse_importe(valor):
    """'Importe ' cell -> float. A hand-maintained file also produces text ('1,234.56', '1,234.56-', '(1,234.56)', '').
    A non-empty token that is still not numeric raises, so a corrupt amount stops the run instead of understating a total."""
    if _es_nulo(valor):
        return 0.0
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor)

    s = str(valor).strip()
    if s in ("", "-", "+"):
        return 0.0

    negativo = False
    if s.startswith("(") and s.endswith(")"):
        negativo = True
        s = s[1:-1].strip()
    if s.endswith("-"):
        negativo = True
        s = s[:-1].strip()
    elif s.endswith("+"):
        s = s[:-1].strip()
    if s.startswith("-"):
        negativo = True
        s = s[1:].strip()
    elif s.startswith("+"):
        s = s[1:].strip()

    s = s.replace(",", "").replace(" ", "").replace("$", "")
    if s == "":
        return 0.0

    try:
        numero = float(s)
    except ValueError:
        raise ValueError(
            f"No pude convertir el importe '{valor}' a número en la columna "
            f"'{COL_IMPORTE}' de la pestaña '{HOJA_BASE}'. ¿Formato inesperado?"
        )

    return -numero if negativo else numero


def _localizar_fila_encabezado(ws):
    """1-based row of the 'Base' header: scans the first 30 rows for 'Mes' + 'Agrupador' (padding rows above it have changed)."""
    buscados = {_normalizar(COL_MES), _normalizar(COL_AGRUPADOR)}
    for fila in range(1, min(ws.max_row, 30) + 1):
        valores = {
            _normalizar(ws.cell(row=fila, column=c).value)
            for c in range(1, ws.max_column + 1)
        }
        if buscados.issubset(valores):
            return fila
    return FILA_ENCABEZADO_BASE_DEFAULT


def _resolver_columna(df, nombre):
    """Column by NORMALISED name; no positional fallback on purpose (it would silently sum the wrong data): fails loud."""
    objetivo = _normalizar(nombre)
    for columna in df.columns:
        if _normalizar(columna) == objetivo:
            return columna
    raise ValueError(
        f"No encontré la columna '{nombre}' en la pestaña '{HOJA_BASE}'. "
        f"Columnas disponibles: {[str(c) for c in df.columns]}"
    )


# --- Matching 'Agrupador' -> catalogue concept ---
def _sin_acentos(texto):
    """'DESCUENTO LOGÍSTICO' -> 'DESCUENTO LOGISTICO'."""
    return (
        unicodedata.normalize("NFKD", str(texto))
        .encode("ascii", "ignore")
        .decode("ascii")
    )


def _canonico(texto):
    """Comparable form of a label: no accents, upper case, punctuation as spaces, single spaces."""
    limpio = re.sub(r"[^A-Za-z0-9\s]", " ", _sin_acentos(texto).upper())
    return re.sub(r"\s+", " ", limpio).strip()


def _singular(palabra):
    """Crude Spanish de-pluralisation ('ESPACIOS'->'ESPACIO', 'PLANES'->'PLAN').
    Leaves gender alone so 'RENTA DE ESPACIO FIJA' is never mistaken for 'DESCUENTO FIJO'."""
    if len(palabra) > 4 and palabra.endswith("ES"):
        return palabra[:-2]
    if len(palabra) > 3 and palabra.endswith("S"):
        return palabra[:-1]
    return palabra


def _tokens_significativos(texto):
    """Set of meaningful, de-pluralised words of a label."""
    return {
        _singular(p)
        for p in _canonico(texto).split()
        if p not in PALABRAS_VACIAS
    }


def emparejar_concepto(agrupador, catalogo=None, umbral=None):
    """Match one 'Agrupador' against the catalogue in three passes, strictest first: (concepto|None, regla, score).
    1 'exacto': identical once canonicalised. 2 'tokens': every meaningful word of one label is in the other (keeps
    the four 'DESCUENTO ...' concepts apart). 3 'fuzzy': similarity above `umbral`, for typos only.
    Unmatched rows return (None, 'sin match', best score) and stay under their raw name."""
    catalogo = catalogo if catalogo is not None else CONCEPTOS_VARIACIONES
    umbral   = umbral if umbral is not None else UMBRAL_SIMILITUD

    canon_agr  = _canonico(agrupador)
    tokens_agr = _tokens_significativos(agrupador)
    if not canon_agr:
        return None, "vacío", 0.0

    # --- 1. exact -----------------------------------------------------
    for concepto in catalogo:
        if _canonico(concepto) == canon_agr:
            return concepto, "exacto", 1.0

    # --- 2. token containment (the most specific candidate wins; ties fall back to fuzzy) ---
    candidatos = []
    for concepto in catalogo:
        tokens_con = _tokens_significativos(concepto)
        if tokens_con and tokens_agr and (
            tokens_con <= tokens_agr or tokens_agr <= tokens_con
        ):
            comunes = len(tokens_con & tokens_agr)
            similitud = SequenceMatcher(None, _canonico(concepto), canon_agr).ratio()
            candidatos.append((comunes, similitud, concepto))
    if candidatos:
        comunes, similitud, concepto = max(candidatos)
        return concepto, "tokens", round(similitud, 3)

    # --- 3. fuzzy (typos) ---------------------------------------------
    mejor, mejor_score = None, 0.0
    for concepto in catalogo:
        score = SequenceMatcher(None, _canonico(concepto), canon_agr).ratio()
        if score > mejor_score:
            mejor, mejor_score = concepto, score

    if mejor_score >= umbral:
        return mejor, "fuzzy", round(mejor_score, 3)
    return None, "sin match", round(mejor_score, 3)


def construir_mapeo_conceptos(agrupadores, catalogo=None, umbral=None):
    """Match every distinct 'Agrupador' ONCE: (mapeo {agrupador: concepto|None}, reporte), the audit trail of label -> row and rule."""
    mapeo, reporte = {}, []
    for agrupador in sorted(set(agrupadores)):
        concepto, regla, score = emparejar_concepto(agrupador, catalogo, umbral)
        mapeo[agrupador] = concepto
        reporte.append({
            "agrupador": agrupador,
            "concepto":  concepto,
            "regla":     regla,
            "score":     score,
        })
    return mapeo, reporte


# --- Matching 'CT' (channel) -> block ---
def resolver_grupo(valor, grupos=None):
    """Index (into `grupos`) of the block a channel label belongs to, or None if it matches none or MORE than one.
    Word-level rule as in the concept matching, so one function reads 'HD', 'smkt nacionales', a template title
    and a file name. Ambiguous labels ('HD NACIONALES') return None on purpose: guessing would move money silently."""
    grupos = grupos if grupos is not None else GRUPOS
    tokens = _tokens_significativos(valor)
    if not tokens:
        return None

    aciertos = [
        i for i, grupo in enumerate(grupos)
        if any(
            (alias_tokens := _tokens_significativos(alias))
            and alias_tokens <= tokens
            for alias in grupo["canales"]
        )
    ]
    return aciertos[0] if len(aciertos) == 1 else None


def asignar_grupos(df, grupos=None):
    """Add '_grupo' to `df`: the block index of each row's '_canal' (-1 if none); each DISTINCT channel is resolved once."""
    grupos = grupos if grupos is not None else GRUPOS
    resueltos = {c: resolver_grupo(c, grupos) for c in df["_canal"].unique()}

    df = df.copy()
    df["_grupo"] = df["_canal"].map(
        lambda c: -1 if resueltos[c] is None else resueltos[c]
    )
    return df


# --- Reading 'Base' ---
def leer_base(ruta_libro):
    """Read 'Base' into a tidy DataFrame: '_nfila' (sheet row), '_mes', '_agrupador', '_importe', '_canal' ('CT')
    and '_canal_venta' ('' without that column). Returns (df, fila_encabezado, ultima_fila), where ultima_fila is
    the last row holding ANY value: Reservas_Manuales pastes below it and refuses to paste under a footer."""
    if not os.path.exists(ruta_libro):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta_libro}")

    wb = load_workbook(ruta_libro, data_only=True, read_only=True)
    if HOJA_BASE not in wb.sheetnames:
        disponibles = wb.sheetnames
        wb.close()
        raise ValueError(
            f"El archivo no tiene una pestaña llamada '{HOJA_BASE}'. "
            f"Pestañas encontradas: {disponibles}"
        )

    ws = wb[HOJA_BASE]
    fila_encabezado = _localizar_fila_encabezado(ws)

    filas = list(
        ws.iter_rows(min_row=fila_encabezado, max_row=ws.max_row, values_only=True)
    )
    wb.close()

    if len(filas) < 2:
        raise ValueError(f"La pestaña '{HOJA_BASE}' no tiene filas de datos.")

    encabezados = [
        str(h) if h is not None else f"_sin_nombre_{i}"
        for i, h in enumerate(filas[0])
    ]
    # Last row holding any value, found BEFORE any filtering, from the raw rows.
    ultima_fila = fila_encabezado
    for desplazamiento in range(len(filas) - 1, 0, -1):
        if any(v is not None and str(v).strip() != "" for v in filas[desplazamiento]):
            ultima_fila = fila_encabezado + desplazamiento
            break

    df = pd.DataFrame(filas[1:], columns=encabezados)
    df = df.dropna(how="all")
    # dropna keeps the 0..n-1 index labels, so a label maps straight back to the sheet row ('_nfila'; '_fila' is taken).
    df["_nfila"] = df.index + fila_encabezado + 1

    c_mes  = _resolver_columna(df, COL_MES)
    c_agr  = _resolver_columna(df, COL_AGRUPADOR)
    c_imp  = _resolver_columna(df, COL_IMPORTE)
    c_can  = _resolver_columna(df, COL_CANAL)

    df["_mes"]        = df[c_mes].map(_normalizar)
    df["_agrupador"]  = df[c_agr].map(lambda v: "" if _es_nulo(v) else str(v).strip())
    df["_importe"]    = df[c_imp].map(_parse_importe)
    df["_canal"]      = df[c_can].map(_normalizar)
    # Not used for grouping any more: only so the paste can spot rows of an older
    # paste (channel in 'Canal Venta', 'CT' still blank) before it doubles them.
    try:
        df["_canal_venta"] = df[_resolver_columna(df, COL_CANAL_VENTA)].map(_normalizar)
    except ValueError:
        df["_canal_venta"] = ""

    # Rows with no month or no concept cannot land in any matrix.
    df = df[(df["_mes"] != "") & (df["_agrupador"] != "")]

    return df, fila_encabezado, ultima_fila


# --- Calculation ---
def calcular_matrices(df, clave_anterior, clave_actual, grupos=None,
                      catalogo=None, umbral=None):
    """Build the data behind the five blocks: every 'Agrupador' is matched ONCE for the whole file (so blocks agree), then per
    group its channel's rows and the two months are summed by concept. Catalogue order, only concepts with activity
    (MOSTRAR_CONCEPTOS_SIN_MOVIMIENTO flips it); unmatched ones are appended under their raw name so a Grand Total is
    the group's real total. Returns (matrices, mapeo, reporte)."""
    grupos   = grupos if grupos is not None else GRUPOS
    catalogo = catalogo if catalogo is not None else CONCEPTOS_VARIACIONES

    mapeo, reporte = construir_mapeo_conceptos(df["_agrupador"], catalogo, umbral)

    # '_grupo' is normally added by the caller (it also needs it for the
    # unassigned-channel report); resolve it here only if it is missing.
    if "_grupo" not in df.columns:
        df = asignar_grupos(df, grupos)
    else:
        df = df.copy()
    df["_concepto"] = df["_agrupador"].map(mapeo)

    # Row label: catalogue concept when matched, else the first spelling seen of its canonical form (no split rows).
    representantes, etiquetas = {}, {}
    for agrupador, concepto in mapeo.items():
        if concepto is not None:
            etiquetas[agrupador] = concepto
        else:
            etiquetas[agrupador] = representantes.setdefault(
                _canonico(agrupador), agrupador
            )
    df["_fila"] = df["_agrupador"].map(etiquetas)

    matrices = []
    for indice, grupo in enumerate(grupos):
        sub = df[
            (df["_grupo"] == indice)
            & df["_mes"].isin([clave_anterior, clave_actual])
        ]

        # Sum once, then read the two months out of the result.
        if sub.empty:
            sumas = {}
        else:
            agrupado = sub.groupby(["_fila", "_mes"])["_importe"].sum()
            sumas = {clave: round(float(v), 2) for clave, v in agrupado.items()}

        con_movimiento = {fila for fila, _ in sumas}

        # Catalogue rows first, in catalogue order...
        conceptos = [
            c for c in catalogo
            if MOSTRAR_CONCEPTOS_SIN_MOVIMIENTO or c in con_movimiento
        ]
        # ...then anything that matched nothing, in order of appearance.
        if AGREGAR_NO_MAPEADOS:
            conceptos += [
                fila for fila in dict.fromkeys(
                    sub.loc[sub["_concepto"].isna(), "_fila"].tolist()
                )
                if fila not in conceptos
            ]

        anterior = {c: sumas.get((c, clave_anterior), 0.0) for c in conceptos}
        actual   = {c: sumas.get((c, clave_actual),   0.0) for c in conceptos}

        matrices.append({
            "titulo":         grupo["titulo"],
            "canales":        grupo["canales"],
            "conceptos":      conceptos,
            "anterior":       anterior,
            "actual":         actual,
            "total_anterior": round(sum(anterior.values()), 2),
            "total_actual":   round(sum(actual.values()), 2),
            "no_mapeados":    sorted(
                set(sub.loc[sub["_concepto"].isna(), "_agrupador"])
            ),
        })

    return matrices, mapeo, reporte


# --- Writing the sheet ---
def _escribir_bloque(ws, fila, matriz, clave_anterior, clave_actual, anio):
    """Write one block starting at `fila`; returns the row where the NEXT block starts."""
    f_banner   = fila
    f_titulo   = fila + 1
    f_encabeza = fila + 2
    f_primera  = fila + 3
    n          = len(matriz["conceptos"])
    f_total    = f_primera + n

    # --- 'Creaciones <año>' banner -----------------------------------
    celda = ws.cell(row=f_banner, column=COL_CONCEPTO, value=f"Creaciones {anio}")
    celda.fill = _relleno(GRIS_CREACIONES)
    celda.font = _NEGRO_NEGRITA

    # --- Group title --------------------------------------------------
    celda = ws.cell(row=f_titulo, column=COL_CONCEPTO, value=matriz["titulo"])
    celda.fill = _relleno(AZUL_TITULO)
    celda.font = _NEGRO_NEGRITA

    # --- Header row ---------------------------------------------------
    encabezados = [
        (COL_CONCEPTO, "Concepto",                 TEAL_HEADER,    None),
        (COL_MES_ANT,  clave_anterior,             TEAL_HEADER,    CONTABLE_2),
        (COL_MES_ACT,  clave_actual,               VERDE_HEADER,   CONTABLE_2),
        (COL_PCT,      "%",                        VERDE_HEADER,   None),
        (COL_BW,       "B/W",                      VERDE_HEADER,   CONTABLE_2),
        (COL_VAR,      f"% VS {clave_anterior}",   NARANJA_HEADER, PORCENTAJE),
        (COL_INDEX,    "INDEX",                    NARANJA_HEADER, CONTABLE_1),
    ]
    for columna, texto, color, formato in encabezados:
        celda = ws.cell(row=f_encabeza, column=columna, value=texto)
        celda.fill = _relleno(color)
        celda.font = _BLANCO_NEGRITA
        celda.alignment = _CENTRADO
        if formato:
            celda.number_format = formato

    # --- Body ---------------------------------------------------------
    letra_ant, letra_act = get_column_letter(COL_MES_ANT), get_column_letter(COL_MES_ACT)
    letra_bw = get_column_letter(COL_BW)

    for i, concepto in enumerate(matriz["conceptos"]):
        f = f_primera + i

        c = ws.cell(row=f, column=COL_CONCEPTO, value=concepto)
        c.font, c.alignment = _NEGRO_NORMAL, _IZQUIERDA

        c = ws.cell(row=f, column=COL_MES_ANT, value=matriz["anterior"][concepto])
        c.font, c.number_format = _NEGRO_NORMAL, CONTABLE_2

        c = ws.cell(row=f, column=COL_SEP)
        c.font, c.number_format = _NEGRO_NORMAL, CONTABLE_0

        c = ws.cell(row=f, column=COL_MES_ACT, value=matriz["actual"][concepto])
        c.font, c.number_format = _NEGRO_NORMAL, CONTABLE_2

        # IFERROR keeps an empty block from filling the column with #DIV/0!.
        c = ws.cell(row=f, column=COL_PCT,
                    value=f"=IFERROR({letra_act}{f}/${letra_act}${f_total},0)")
        c.font, c.number_format, c.alignment = _NEGRO_NORMAL, PORCENTAJE, _CENTRADO

        # B/W keeps the stakeholder's sign: previous minus current.
        c = ws.cell(row=f, column=COL_BW,
                    value=f"={letra_ant}{f}-{letra_act}{f}")
        c.font, c.number_format = _NEGRO_NORMAL, CONTABLE_2

        c = ws.cell(row=f, column=COL_VAR,
                    value=f"=IFERROR(-{letra_bw}{f}/{letra_ant}{f},0)")
        c.font, c.number_format, c.alignment = _NEGRO_NORMAL, PORCENTAJE, _CENTRADO

        c = ws.cell(row=f, column=COL_INDEX,
                    value=f"=IFERROR(((-{letra_bw}{f}/{letra_ant}{f}*100)+100),0)")
        c.font, c.number_format, c.alignment = _NEGRO_NORMAL, CONTABLE_1, _CENTRADO

    # --- 'Grand Total' row --------------------------------------------
    c = ws.cell(row=f_total, column=COL_CONCEPTO, value="Grand Total")
    c.font, c.number_format, c.border = _NEGRO_NEGRITA, CONTABLE_2, _BORDE_TOTAL_CONCEPTO

    if n > 0:
        rango = lambda letra: f"{letra}{f_primera}:{letra}{f_total - 1}"
        totales = [
            (COL_MES_ANT, f"=SUM({rango(letra_ant)})", CONTABLE_2),
            (COL_MES_ACT, f"=SUM({rango(letra_act)})", CONTABLE_2),
            (COL_PCT,     f"=SUM({rango(get_column_letter(COL_PCT))})", PORCENTAJE_TOTAL),
            (COL_BW,      f"={letra_ant}{f_total}-{letra_act}{f_total}", CONTABLE_2),
        ]
        if TOTAL_INCLUYE_VARIACION:
            totales += [
                (COL_VAR,   f"=IFERROR(-{letra_bw}{f_total}/{letra_ant}{f_total},0)", PORCENTAJE),
                (COL_INDEX, f"=IFERROR(((-{letra_bw}{f_total}/{letra_ant}{f_total}*100)+100),0)", CONTABLE_1),
            ]
        for columna, formula, formato in totales:
            c = ws.cell(row=f_total, column=columna, value=formula)
            c.font, c.number_format, c.border = _NEGRO_NEGRITA, formato, _BORDE_TOTAL
            if columna in (COL_PCT, COL_VAR, COL_INDEX):
                c.alignment = _CENTRADO
    else:
        # Empty group: leave the total row visible at zero so the block does not look missing.
        for columna in (COL_MES_ANT, COL_MES_ACT, COL_BW):
            c = ws.cell(row=f_total, column=columna, value=0)
            c.font, c.number_format, c.border = _NEGRO_NEGRITA, CONTABLE_2, _BORDE_TOTAL

    return f_total + 1 + FILAS_ENTRE_BLOQUES


def escribir_variaciones(wb, matrices, clave_anterior, clave_actual):
    """Drop any existing 'Variaciones' sheet and rebuild it in its original tab position."""
    posicion = None
    if HOJA_VARIACIONES in wb.sheetnames:
        posicion = wb.sheetnames.index(HOJA_VARIACIONES)
        del wb[HOJA_VARIACIONES]

    ws = wb.create_sheet(HOJA_VARIACIONES, posicion)

    anio = parse_clave_mes(clave_actual).year
    fila = FILA_INICIAL
    for matriz in matrices:
        fila = _escribir_bloque(ws, fila, matriz, clave_anterior, clave_actual, anio)

    for letra, ancho in ANCHOS_COLUMNA.items():
        ws.column_dimensions[letra].width = ancho

    return ws


# --- Main entry point ---
def construir_variaciones(
    ruta_libro,
    ruta_salida=None,
    mes=None,
    grupos=None,
    sobrescribir_original=False,
    con_comparativos=None,
    callback_status=None,
    templates=None,
):
    """Build 'Variaciones', optionally pasting the 'RESERVAS MANUALES' `templates` into 'Base' first (None/empty = plain
    consolidation; if any template fails validation nothing is written). `mes`: None = current month, 'M07-26' re-runs a past one.
    Writes a COPY unless sobrescribir_original; `con_comparativos` None -> CONSTRUIR_COMPARATIVOS; `grupos` overrides GRUPOS (tests).
    Returns a dict: ruta, mes_anterior, mes_actual, filas_base, bloques, totales... ('pegado': paste summary or None)."""
    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    if con_comparativos is None:
        con_comparativos = CONSTRUIR_COMPARATIVOS

    if templates and sobrescribir_original:
        # A failed paste cannot be rolled back in the user's own file, so it only ever runs against a COPY.
        raise ValueError(
            "Pegar los templates de Reservas Manuales sólo se permite sobre una "
            "COPIA (sobrescribir_original=False)."
        )

    clave_anterior, clave_actual = claves_mes_actual_y_anterior(mes)
    update_status(
        f"📅 Comparando {etiqueta_mes_es(clave_anterior)} vs "
        f"{etiqueta_mes_es(clave_actual)}  ({clave_anterior} → {clave_actual})"
    )

    update_status(f"📄 Leyendo la pestaña '{HOJA_BASE}'...")
    df, fila_encabezado, ultima_fila = leer_base(ruta_libro)

    # --- Optional: plan the paste of the templates ---
    # Planned and validated COMPLETELY before anything is written; applied to the in-memory DataFrame here and to the copy below.
    # Imported HERE because Reservas_Manuales imports this module (a top-level import is circular; PyInstaller needs --hidden-import).
    plan = None
    if templates:
        from Reservas_Manuales import aplicar_plan_a_df, planificar_pegado

        update_status(f"📥 Revisando {len(templates)} template(s) de Reservas Manuales...")
        plan = planificar_pegado(
            df, fila_encabezado, ultima_fila, templates, clave_actual, grupos,
        )
        df = aplicar_plan_a_df(df, plan)
        for aviso in plan["advertencias"]:
            update_status(f"⚠️ {aviso}")

    df = asignar_grupos(df, grupos)

    update_status("🧮 Calculando las cinco matrices...")
    matrices, mapeo, reporte = calcular_matrices(df, clave_anterior, clave_actual, grupos)

    # Money whose channel belongs to no block would silently miss every matrix: report it with its amount.
    huerfanas = df[(df["_grupo"] == -1) & df["_mes"].isin([clave_anterior, clave_actual])]
    canales_sin_bloque = [
        f"{canal or '(vacío)'} — {int(n)} fila(s), Σ {suma:,.2f}"
        for canal, (n, suma) in huerfanas.groupby("_canal")["_importe"].agg(["size", "sum"]).iterrows()
    ]

    no_mapeados = sorted({r["agrupador"] for r in reporte if r["concepto"] is None})
    if no_mapeados:
        update_status(
            f"⚠️ {len(no_mapeados)} valor(es) de 'Agrupador' no empataron con el "
            f"catálogo y salen con su nombre original: {', '.join(no_mapeados)}"
        )

    # --- Where the result goes ----------------------------------------
    if sobrescribir_original:
        ruta_final = ruta_libro
    else:
        if not ruta_salida:
            ruta_salida = carpeta_output_desde_input(os.path.dirname(os.path.abspath(ruta_libro)))
        os.makedirs(ruta_salida, exist_ok=True)
        base = os.path.splitext(os.path.basename(ruta_libro))[0]
        ruta_final = os.path.join(ruta_salida, f"{base}_Variaciones_{clave_actual}.xlsx")
        shutil.copy2(ruta_libro, ruta_final)

    pegado = None
    try:
        wb = load_workbook(ruta_final)
        try:
            info_escritura = None
            if plan:
                from Reservas_Manuales import aplicar_plan_a_libro

                update_status(f"📥 Pegando los templates en la pestaña '{HOJA_BASE}'...")
                info_escritura = aplicar_plan_a_libro(wb, plan)

            update_status(f"🎨 Escribiendo la pestaña '{HOJA_VARIACIONES}'...")
            escribir_variaciones(wb, matrices, clave_anterior, clave_actual)
            wb.save(ruta_final)
        finally:
            wb.close()

        if plan:
            # Re-open what is ON DISK and compare it cell by cell: a paste that looked fine in memory must not reach the stakeholder.
            from Reservas_Manuales import resumir_pegado, verificar_pegado

            update_status("🔎 Verificando lo pegado contra los templates...")
            verificar_pegado(ruta_final, plan, info_escritura)
            pegado = resumir_pegado(plan, info_escritura)
    except Exception:
        # The copy in 'Output' would look like a real result: remove it (the original was never touched).
        if not sobrescribir_original:
            try:
                os.remove(ruta_final)
            except OSError:
                pass
        raise

    # --- Optional second pass: the two PivotTable sheets ---
    # Runs AFTER openpyxl has saved and closed: it cannot round-trip a brand new pivot cache.
    comparativos, comparativos_error = None, None
    if con_comparativos:
        try:
            # Imported HERE so this module still imports on a machine with no Excel / pywin32.
            from Comparativos import construir_comparativos

            comparativos = construir_comparativos(
                ruta_final, mes=mes, callback_status=callback_status,
            )
        except Exception as e:
            # Deliberately broad: 'Variaciones' is already saved, so a failure here must downgrade to a warning, not discard it.
            comparativos_error = str(e)
            update_status(f"⚠️ No pude crear las tablas dinámicas: {e}")

    update_status("✅ Consolidación completada")

    return {
        "ruta":             ruta_final,
        "mes_anterior":     clave_anterior,
        "mes_actual":       clave_actual,
        "filas_base":       len(df),
        "fila_encabezado":  fila_encabezado,
        "bloques":          len(matrices),
        "conceptos_por_bloque": {m["titulo"]: len(m["conceptos"]) for m in matrices},
        "total_anterior":   round(sum(m["total_anterior"] for m in matrices), 2),
        "total_actual":     round(sum(m["total_actual"] for m in matrices), 2),
        # Audit trail of the Agrupador -> Concepto matching, so a wrong match is visible.
        "mapeo_conceptos":  reporte,
        "conceptos_no_mapeados": no_mapeados,
        # PivotTable pass summary (None if skipped) and the reason it did not run, if it failed.
        "comparativos":         comparativos,
        "comparativos_error":   comparativos_error,
        # Channels of the two months that belong to no block: their money is NOT in any matrix.
        "canales_sin_bloque":   canales_sin_bloque,
        # Summary of the template paste (None when no templates were chosen).
        "pegado":               pegado,
    }


# --- Standalone run: python Consolidacion.py ---
if __name__ == "__main__":
    print("=== Prueba de consolidación 'Variaciones' (standalone) ===")
    ruta = input("Ruta del archivo de Reservas (.xlsx): ").strip().strip('"')
    mes_txt = input("Mes a correr [Enter = hoy, o p.ej. M07-26]: ").strip() or None
    rutas_txt = input("Templates de Reservas Manuales, separados por ';' [Enter = ninguno]: ")
    templates_txt = [r.strip().strip('"') for r in rutas_txt.split(";") if r.strip()]

    try:
        resumen = construir_variaciones(
            ruta, mes=mes_txt, callback_status=print, templates=templates_txt or None,
        )
        print("\n🎉 LISTO 🎉")
        for clave, valor in resumen.items():
            print(f"  {clave:<22}: {valor}")
    except Exception as e:
        print(f"❌ Error durante la consolidación: {e}")
        raise
