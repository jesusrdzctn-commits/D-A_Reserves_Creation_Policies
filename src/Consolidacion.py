"""
RDA Pólizas D&A — Consolidación
================================

Builds the 'Variaciones' sheet of '07 Archivo Creacion de Reservas PMF OT.xlsx'
out of the 'Base' sheet.

This is the mirror image of the Validación Factura Global flow: there we
downloaded from SAP and then consolidated; here we consolidate first and the
upload to SAP comes afterwards.

What the sheet looks like (one block per group, stacked vertically, one blank
row between blocks):

    A1   Creaciones 2026                      <- grey banner
    A2   Regionales David                     <- light blue title (group + analyst)
    A3   Concepto | M06-26 |   | M07-26 | % | B/W | % VS M06-26 | INDEX
    A4   <agrupador>   1,000.00      500.00  ...
    ...
    A17  Grand Total
    (blank row)
    A19  Creaciones 2026   ... next block

Columns (fixed, taken from the stakeholder's hand-made sheet):

    A  Concepto      unique 'Agrupador' values present in that group
    B  <mes pasado>  sum of 'Importe ' for the previous month        [value]
    C  (empty)       narrow visual separator
    D  <mes actual>  sum of 'Importe ' for the current month         [value]
    E  %             share of the block total       =IFERROR(D/$D$total,0)
    F  B/W           =B-D   (previous minus current, the stakeholder's sign)
    G  % VS <mes>    =IFERROR(-F/B,0)
    H  INDEX         =IFERROR(((-F/B*100)+100),0)   i.e. current as base-100

Amounts in B and D are pasted as VALUES (already summed in pandas). Everything
else is written as a LIVE FORMULA, exactly like the sheet the stakeholder
maintains by hand, so that editing B or D keeps the block consistent.

This module does NOT depend on SAP or tkinter: it can be tested on its own with
`python Consolidacion.py`.
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

from utils import claves_mes_actual_y_anterior, etiqueta_mes_es, parse_clave_mes

# ----------------------------------------------------------------------
# Sheets and source columns
# ----------------------------------------------------------------------
HOJA_BASE        = "Base"
HOJA_VARIACIONES = "Variaciones"

# Columns read from 'Base'. Matched by NORMALISED name (trimmed, case-folded),
# because the real file ships 'Importe ' with a trailing space.
COL_MES       = "Mes"
COL_AGRUPADOR = "Agrupador"
COL_IMPORTE   = "Importe"
COL_ANALISTA  = "Solicitante / KAM Resp"

# The header row of 'Base' is not row 1 (there are empty rows plus a
# 'Seleccionar / Escritura / Formula' hint row above it), so it is detected by
# looking for these headers. This is the fallback if detection fails.
FILA_ENCABEZADO_BASE_DEFAULT = 7


# ----------------------------------------------------------------------
# The five blocks
# ----------------------------------------------------------------------
# PENDING: the real rule that maps an analyst to a group is still with the
# stakeholder. This list is the whole configuration point — when the rule
# arrives, only this constant changes.
#
#   titulo            : text of the block title.
#   analistas         : values of 'Solicitante / KAM Resp' that feed the block.
#   mostrar_analista  : True appends the analyst name to the title
#                       ('Regionales' + 'DAVID' -> 'Regionales David').
#                       False leaves the title alone ('ECOMMERCE', 'HD KAREN').
GRUPOS = [
    {"titulo": "Smkt Regionales", "analistas": ["DAVID"],  "mostrar_analista": True},
    {"titulo": "Proximity",       "analistas": ["JIMENA"], "mostrar_analista": True},
    {"titulo": "Smkt Nacionales", "analistas": ["CARLOS"], "mostrar_analista": True},
    {"titulo": "ECOMMERCE",       "analistas": ["PABLO"],  "mostrar_analista": False},
    {"titulo": "HD",              "analistas": ["KAREN"],  "mostrar_analista": False},
]


# ----------------------------------------------------------------------
# Concept catalogue (the row labels of every matrix)
# ----------------------------------------------------------------------
# The stakeholder's fixed list. It does NOT match 'Agrupador' word for word
# (the list says 'RENTA DE ESPACIOS', 'Base' says 'RENTA DE ESPACIO FIJA'), so
# every 'Agrupador' is matched against this catalogue by emparejar_concepto().
# Row order in the sheet follows this list.
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

# Similarity threshold for the THIRD pass only (typos). The first two passes are
# exact/token rules and do not use it. Tuned against the real naming variants:
# the only true match that reaches this pass ('DIFERENENCIA EN PRECIO') scores
# 0.95, while the best false candidate ('DESCUENTO DE LÍNEA' against 'DESCUENTO
# FIJO') scores 0.72, so 0.86 sits in the gap. Lower it and the four different
# 'DESCUENTO ...' concepts start colliding.
UMBRAL_SIMILITUD = 0.86

# An 'Agrupador' that matches nothing in the catalogue is added as its own row
# at the bottom of its block. Money never disappears from a Grand Total just
# because a concept was not on the list; it shows up under its raw name and is
# also reported in 'conceptos_no_mapeados'.
AGREGAR_NO_MAPEADOS = True

# False (the stakeholder's call) prints only the concepts with activity in that
# block. True prints all 11 catalogue rows in every matrix, zeros included, so
# the five blocks line up row by row.
MOSTRAR_CONCEPTOS_SIN_MOVIMIENTO = False


# ----------------------------------------------------------------------
# Layout of one block (1-based column indexes)
# ----------------------------------------------------------------------
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


# ----------------------------------------------------------------------
# Styles — lifted from the stakeholder's sheet (theme colours already resolved
# to their RGB, so the output looks identical without depending on the theme)
# ----------------------------------------------------------------------
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


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def carpeta_output_desde_input(ruta_input):
    """
    Return the 'Output' folder that sits next to the input folder.

    Same convention as Validación Factura Global: if the chosen folder is named
    'Input' the Output lands BESIDE it (src/Output); for any other folder the
    Output is created INSIDE it, so nothing is scattered outside the user's
    choice.
    """
    ruta = os.path.normpath(ruta_input)
    if os.path.basename(ruta).lower() == "input":
        return os.path.join(os.path.dirname(ruta), "Output")
    return os.path.join(ruta, "Output")


def _normalizar(texto):
    """Trim and upper-case, tolerating None. Used to match headers and values."""
    return str(texto).strip().upper() if texto is not None else ""


def _parse_importe(valor):
    """
    Convert an 'Importe ' cell into a float.

    The column normally arrives as a real number, but a hand-maintained file
    also produces text ('1,234.56', '1,234.56-', '(1,234.56)', '') so the same
    tolerant parser used in Validación Factura Global applies here. A non-empty
    token that is still not numeric raises, so a corrupt amount stops the run
    instead of quietly understating a block total.
    """
    if valor is None:
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
    """
    Find the 1-based row of the 'Base' header.

    The sheet carries empty rows and a 'Seleccionar / Escritura / Formula' hint
    row above the real headers, and that padding has changed before. Scanning
    the first 30 rows for 'Mes' + 'Agrupador' survives someone inserting a row.
    """
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
    """
    Resolve a column strictly by NORMALISED name (trimmed, case-insensitive).

    No positional fallback on purpose: guessing a column index here would mean
    silently summing the wrong data. If the header is missing we fail loud with
    the list of what was actually found.
    """
    objetivo = _normalizar(nombre)
    for columna in df.columns:
        if _normalizar(columna) == objetivo:
            return columna
    raise ValueError(
        f"No encontré la columna '{nombre}' en la pestaña '{HOJA_BASE}'. "
        f"Columnas disponibles: {[str(c) for c in df.columns]}"
    )


# ----------------------------------------------------------------------
# Matching 'Agrupador' -> catalogue concept
# ----------------------------------------------------------------------
def _sin_acentos(texto):
    """'DESCUENTO LOGÍSTICO' -> 'DESCUENTO LOGISTICO'."""
    return (
        unicodedata.normalize("NFKD", str(texto))
        .encode("ascii", "ignore")
        .decode("ascii")
    )


def _canonico(texto):
    """
    Reduce a label to its comparable form: no accents, upper case, punctuation
    turned into spaces, single spaces. This alone already absorbs the accent,
    case and slash-spacing variants ('Aperturas /Apoyos' vs 'APERTURA / APOYOS').
    """
    limpio = re.sub(r"[^A-Za-z0-9\s]", " ", _sin_acentos(texto).upper())
    return re.sub(r"\s+", " ", limpio).strip()


def _singular(palabra):
    """
    Crude Spanish de-pluralisation, enough for these labels:
    'ESPACIOS'->'ESPACIO', 'PLANES'->'PLAN', 'APOYOS'->'APOYO'.

    Deliberately does NOT touch gender, so 'FIJA' never collapses into 'FIJO'
    and 'RENTA DE ESPACIO FIJA' cannot be mistaken for 'DESCUENTO FIJO'.
    """
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
    """
    Match one 'Agrupador' value against the catalogue, in three passes from
    strictest to loosest. Returns (concepto|None, regla, score).

        1. 'exacto'  — identical once canonicalised (accents/case/punctuation).
        2. 'tokens'  — every meaningful word of one label is present in the
                       other. This is what makes 'RENTA DE ESPACIO FIJA' land on
                       'RENTA DE ESPACIOS' while keeping the four different
                       'DESCUENTO ...' concepts apart, because each of them
                       needs its own second word to be present.
        3. 'fuzzy'   — character similarity above `umbral`, for typos only
                       ('DIFERENENCIA EN PRECIO' -> 'DIFERENCIA EN PRECIO').

    Returns (None, 'sin match', mejor_score) when nothing clears the bar. The
    caller keeps those rows under their raw name instead of dropping them.
    """
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

    # --- 2. token containment ----------------------------------------
    # Several concepts could contain each other in theory, so the most specific
    # one wins (the one sharing the most words) and ties fall back to fuzzy.
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
    """
    Match every distinct 'Agrupador' ONCE and return (mapeo, reporte).

    `mapeo` is {agrupador: concepto|None}; `reporte` is the audit trail the
    summary carries, so the stakeholder can see exactly which raw label landed
    on which catalogue row and by which rule.
    """
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


# ----------------------------------------------------------------------
# Reading 'Base'
# ----------------------------------------------------------------------
def leer_base(ruta_libro):
    """
    Read the 'Base' sheet into a tidy DataFrame with four working columns:
    '_mes', '_agrupador', '_importe', '_analista'.

    Returns:
        (df, fila_encabezado)
    """
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
    df = pd.DataFrame(filas[1:], columns=encabezados)
    df = df.dropna(how="all")

    c_mes  = _resolver_columna(df, COL_MES)
    c_agr  = _resolver_columna(df, COL_AGRUPADOR)
    c_imp  = _resolver_columna(df, COL_IMPORTE)
    c_ana  = _resolver_columna(df, COL_ANALISTA)

    df["_mes"]        = df[c_mes].map(_normalizar)
    df["_agrupador"]  = df[c_agr].map(lambda v: str(v).strip() if v is not None else "")
    df["_importe"]    = df[c_imp].map(_parse_importe)
    df["_analista"]   = df[c_ana].map(_normalizar)

    # Rows with no month or no concept cannot land in any matrix.
    df = df[(df["_mes"] != "") & (df["_agrupador"] != "")]

    return df, fila_encabezado


# ----------------------------------------------------------------------
# Calculation
# ----------------------------------------------------------------------
def calcular_matrices(df, clave_anterior, clave_actual, grupos=None,
                      catalogo=None, umbral=None):
    """
    Build the data behind the five blocks.

    Every 'Agrupador' is first matched to a catalogue concept (once for the whole
    file, so the five blocks always agree). Then, per group: keep only its
    analysts and the two months in play, and sum 'Importe ' by concept.

    Row order is the catalogue's order, showing only the concepts with activity
    in that block (the stakeholder's call; flip
    MOSTRAR_CONCEPTOS_SIN_MOVIMIENTO to print all 11 every time). Any
    'Agrupador' that matched nothing is appended at the bottom under its raw
    name, so a block's Grand Total is always the group's real total.

    Returns:
        (matrices, mapeo, reporte)
    """
    grupos   = grupos if grupos is not None else GRUPOS
    catalogo = catalogo if catalogo is not None else CONCEPTOS_VARIACIONES

    mapeo, reporte = construir_mapeo_conceptos(df["_agrupador"], catalogo, umbral)

    df = df.copy()
    df["_concepto"] = df["_agrupador"].map(mapeo)

    # Row label: the catalogue concept when matched, the raw value otherwise.
    # Two spellings of the SAME unmatched concept ('DESCUENTO DE LÍNEA' and
    # 'DESCUENTO DE LINEA') must not split into two rows, so they share the
    # first spelling seen for their canonical form.
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
    for grupo in grupos:
        analistas = [_normalizar(a) for a in grupo["analistas"]]
        sub = df[
            df["_analista"].isin(analistas)
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

        titulo = grupo["titulo"]
        if grupo.get("mostrar_analista") and grupo["analistas"]:
            titulo = f"{titulo} {' / '.join(a.title() for a in grupo['analistas'])}"

        matrices.append({
            "titulo":         titulo,
            "analistas":      grupo["analistas"],
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


# ----------------------------------------------------------------------
# Writing the sheet
# ----------------------------------------------------------------------
def _escribir_bloque(ws, fila, matriz, clave_anterior, clave_actual, anio):
    """
    Write one block starting at `fila` and return the row where the NEXT block
    should start (total row + blank separators).
    """
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

        # Share of the block's current-month total. IFERROR keeps an empty block
        # from filling the column with #DIV/0!.
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
        # Empty group: leave the total row visible at zero so the stakeholder
        # sees the block exists rather than wondering where it went.
        for columna in (COL_MES_ANT, COL_MES_ACT, COL_BW):
            c = ws.cell(row=f_total, column=columna, value=0)
            c.font, c.number_format, c.border = _NEGRO_NEGRITA, CONTABLE_2, _BORDE_TOTAL

    return f_total + 1 + FILAS_ENTRE_BLOQUES


def escribir_variaciones(wb, matrices, clave_anterior, clave_actual):
    """
    Drop any existing 'Variaciones' sheet and rebuild it, keeping its original
    tab position so the workbook's tab order does not shuffle.
    """
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


# ----------------------------------------------------------------------
# Main entry point
# ----------------------------------------------------------------------
def construir_variaciones(
    ruta_libro,
    ruta_salida=None,
    mes=None,
    grupos=None,
    sobrescribir_original=False,
    callback_status=None,
):
    """
    Build the 'Variaciones' sheet of the Reservas workbook.

    Args:
        ruta_libro (str): path to '07 Archivo Creacion de Reservas PMF OT.xlsx'.
        ruta_salida (str|None): folder for the output copy. None -> the 'Output'
            folder resolved by carpeta_output_desde_input().
        mes (datetime|str|None): reference month. None -> today (so 'current
            month' is the month the code runs, per the spec). A month key
            ('M07-26') re-runs a past month.
        grupos (list|None): override the GRUPOS mapping (handy for tests).
        sobrescribir_original (bool): True writes back into the source file.
            Default False — it writes a copy, because openpyxl cannot preserve
            everything a hand-maintained workbook carries (see the note in the
            project README).
        callback_status (function|None): progress reporting for the GUI.

    Returns:
        dict: ruta, mes_anterior, mes_actual, filas_base, bloques, totales...
    """
    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    clave_anterior, clave_actual = claves_mes_actual_y_anterior(mes)
    update_status(
        f"📅 Comparando {etiqueta_mes_es(clave_anterior)} vs "
        f"{etiqueta_mes_es(clave_actual)}  ({clave_anterior} → {clave_actual})"
    )

    update_status(f"📄 Leyendo la pestaña '{HOJA_BASE}'...")
    df, fila_encabezado = leer_base(ruta_libro)

    update_status("🧮 Calculando las cinco matrices...")
    matrices, mapeo, reporte = calcular_matrices(df, clave_anterior, clave_actual, grupos)

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

    update_status(f"🎨 Escribiendo la pestaña '{HOJA_VARIACIONES}'...")
    wb = load_workbook(ruta_final)
    escribir_variaciones(wb, matrices, clave_anterior, clave_actual)
    wb.save(ruta_final)
    wb.close()

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
        # Audit trail of the Agrupador -> Concepto matching, so a wrong match is
        # visible instead of buried inside a total.
        "mapeo_conceptos":  reporte,
        "conceptos_no_mapeados": no_mapeados,
        "analistas_sin_bloque": sorted(
            set(df["_analista"].unique())
            - {_normalizar(a) for g in (grupos or GRUPOS) for a in g["analistas"]}
        ),
    }


# ----------------------------------------------------------------------
# Standalone run (no GUI, no controller):  python Consolidacion.py
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=== Prueba de consolidación 'Variaciones' (standalone) ===")
    ruta = input("Ruta del archivo de Reservas (.xlsx): ").strip().strip('"')
    mes_txt = input("Mes a correr [Enter = hoy, o p.ej. M07-26]: ").strip() or None

    try:
        resumen = construir_variaciones(ruta, mes=mes_txt, callback_status=print)
        print("\n🎉 LISTO 🎉")
        for clave, valor in resumen.items():
            print(f"  {clave:<22}: {valor}")
    except Exception as e:
        print(f"❌ Error durante la consolidación: {e}")
        raise
