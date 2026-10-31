"""RDA Pólizas D&A — Reservas_Manuales: pastes the five 'RESERVAS MANUALES <CANAL> <MES>' templates (both tabs:
'Reservas Manuales' required, 'Reservas Rentas' optional) into 'Base' of a COPY of the Reservas workbook.
The channel comes from the template TITLE (-> 'CT'); re-pasting a (month, channel) REPLACES its rows.
Validates everything first, rejects templates of another month and verifies the paste on disk.
"""

import math
import os
import re
import shutil
from copy import copy
from datetime import date, datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.datetime import from_excel

# Consolidacion imports THIS module lazily (inside construir_variaciones), so this top-level import is not circular;
# run as a script, Consolidacion loads twice (as __main__ and 'Consolidacion'), which is harmless.
from Consolidacion import (
    COL_CANAL,
    GRUPOS,
    HOJA_BASE,
    _canonico,
    _localizar_fila_encabezado,
    _normalizar,
    _parse_importe,
    asignar_grupos,
    leer_base,
    resolver_grupo,
)
from utils import (
    MESES_ES,
    carpeta_output_desde_input,
    clave_mes,
    claves_mes_actual_y_anterior,
    norm_texto,
    parse_clave_mes,
)

# --- Template layout ---
# First tab (required); without that name the only visible tab that is not the rents one is used.
HOJA_TEMPLATE = "Reservas Manuales"
# Second tab (optional): same columns, pasted into the same 'Base'.
HOJA_RENTAS = "Reservas Rentas"

# (column of 'Base', column of the template). ORDER = order of the summary. The file's channel goes to
# 'CT' (COL_BASE_CANAL below), which has no template column, so it is not listed.
MAPEO_COLUMNAS = (
    ("Mes",                         "MES"),
    ("Cadena",                      "CLIENTE"),
    ("Canal Venta",                 "CANAL"),
    ("BU",                          "BU"),
    ("Agrupador",                   "AGRUPADOR ABIERTO"),
    ("Motivo",                      "MOTIVO"),
    ("Descripción Motivo",          "AGRUPADOR ABIERTO"),
    ("Descripción de la Provisión", "DESCRIPCIÓN DE LA PROVISIÓN"),
    ("Importe",                     "RESERVA ($)"),
    ("Inicio Vigencia",             "INICIO"),
    ("Término Vigencia",            "FIN"),
)

# 'Base' column that receives the CHANNEL OF THE FILE (read from the template TITLE, never a cell);
# the 'Variaciones' blocks are grouped by it (Consolidacion.COL_CANAL).
COL_BASE_CANAL = COL_CANAL

# Template columns a data row cannot do without. Anything else may be blank.
OBLIGATORIAS_TEMPLATE = ("MES", "AGRUPADOR ABIERTO", "RESERVA ($)")

# Month key the templates must carry: the month being run. Set True only to
# back-fill a closed month on purpose.
PERMITIR_OTROS_MESES = False

MAX_FILAS_ENCABEZADO = 30
MAX_FILAS_EN_BLANCO  = 500
MAX_COLUMNAS_BASE    = 200
MAX_ERRORES_LISTADOS = 10
MAX_LARGO_CELDA      = 32767

FORMATOS_FECHA = ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y")

# Month names/abbreviations accepted in a 'MES' cell that is not already a key.
_MESES_TEXTO = {}
for _n, _nombre in MESES_ES.items():
    _MESES_TEXTO[_nombre.upper()] = _n
    _MESES_TEXTO[_nombre[:3].upper()] = _n
_MESES_TEXTO.update({"SEPT": 9, "JAN": 1, "APR": 4, "AUG": 8, "DEC": 12})


# --- Cell-level helpers ---
# Zero-width/direction marks, BOM and soft hyphen: invisible, NOT removed by str.strip(), and a cell of only these LOOKS empty.
_INVISIBLES = re.compile("[\u00ad\u200b-\u200f\u2060\ufeff]")


def _es_vacio(valor):
    """True for a cell with nothing visible: None, NaN, whitespace-only text or zero-width characters."""
    if valor is None:
        return True
    if isinstance(valor, float) and math.isnan(valor):
        return True
    return isinstance(valor, str) and _INVISIBLES.sub("", valor).strip() == ""


def _como_esta_vacio(valor):
    """How an empty cell is empty, for the messages: 'vacía' or 'solo espacios'."""
    return "vacía" if valor is None or not isinstance(valor, str) else "solo espacios en blanco"


def _limpiar(valor):
    """Blank -> None; text -> trimmed and free of invisible characters; else untouched."""
    if _es_vacio(valor):
        return None
    return _INVISIBLES.sub("", valor).strip() if isinstance(valor, str) else valor


def clave_de_mes(valor):
    """'MES' cell -> month key as 'Base' writes it ('M10-26'); raises ValueError.
    Accepts the key, a real date, 'OCT 26' / 'octubre 2026' / 'Oct-26' and '2026-10'; anything else is refused, never guessed."""
    if _es_vacio(valor):
        raise ValueError("está vacío")

    if isinstance(valor, (datetime, date)):
        return clave_mes(datetime(valor.year, valor.month, 1))

    texto = _limpiar(valor) if isinstance(valor, str) else str(valor).strip()
    try:
        return clave_mes(parse_clave_mes(texto))
    except ValueError:
        pass

    partes = _canonico(texto).split()
    if len(partes) == 2 and partes[0] in _MESES_TEXTO and partes[1].isdigit() \
            and len(partes[1]) in (2, 4):
        anio = int(partes[1])
        anio = anio + 2000 if anio < 100 else anio
        return clave_mes(datetime(anio, _MESES_TEXTO[partes[0]], 1))

    coincidencia = re.fullmatch(r"(\d{4})[-/](\d{1,2})", texto)
    if coincidencia and 1 <= int(coincidencia.group(2)) <= 12:
        return clave_mes(datetime(int(coincidencia.group(1)), int(coincidencia.group(2)), 1))

    raise ValueError(f"'{texto}' no es un mes reconocible (se espera algo como M10-26 u OCT 26)")


def fecha_de(valor):
    """'INICIO'/'FIN' cell -> (datetime|None, aviso|None). Unparseable text is kept as it came with a warning: never fatal."""
    if _es_vacio(valor):
        return None, None
    if isinstance(valor, datetime):
        return valor, None
    if isinstance(valor, date):
        return datetime(valor.year, valor.month, valor.day), None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        try:
            return from_excel(valor), None
        except (ValueError, OverflowError, TypeError):
            return valor, f"fecha numérica fuera de rango ({valor})"

    texto = _limpiar(valor) if isinstance(valor, str) else str(valor).strip()
    for formato in FORMATOS_FECHA:
        try:
            return datetime.strptime(texto, formato), None
        except ValueError:
            continue
    return texto, f"fecha no reconocida '{texto}' (se pegó como texto)"


def importe_de(valor):
    """'RESERVA ($)' cell -> finite float; raises ValueError. Accepts ' 1,234.50 ', '$1,234.50', '(500.00)', '500.00-'.
    Refuses blanks (reported by cell), a bare '-'/'+', Excel errors and non-numeric text."""
    if _es_vacio(valor):
        raise ValueError(f"está {_como_esta_vacio(valor)}")
    if isinstance(valor, bool):
        raise ValueError(f"'{valor}' no es un importe")

    if isinstance(valor, str):
        texto = re.sub(r"\s+", "", _INVISIBLES.sub("", valor))
        if texto.startswith("$"):
            texto = texto[1:]
        if texto == "-":
            return 0.0   # accounting dash = zero
        if texto in ("", "+"):
            raise ValueError(
                f"'{valor.strip()}' no es un importe (escribe 0 si de verdad es cero)"
            )
        mostrado = valor.strip()
    else:
        texto, mostrado = valor, valor

    try:
        numero = _parse_importe(texto)
    except ValueError:
        raise ValueError(f"'{mostrado}' no es un número")
    if not math.isfinite(numero):
        raise ValueError(f"'{mostrado}' no es un número finito")
    return numero


def _problema_de_texto(valor):
    """Why a text value cannot be written into a cell, or None if it can."""
    if isinstance(valor, str):
        if ILLEGAL_CHARACTERS_RE.search(valor):
            return "trae caracteres de control que Excel no admite"
        if len(valor) > MAX_LARGO_CELDA:
            return f"es demasiado largo ({len(valor):,} caracteres)"
    return None


def _listar(errores, maximo=MAX_ERRORES_LISTADOS):
    """First `maximo` problems, plus a '...y N más' line."""
    texto = "\n".join(f"    • {e}" for e in errores[:maximo])
    if len(errores) > maximo:
        texto += f"\n    • ...y {len(errores) - maximo} más"
    return texto


# --- Reading ONE template ---
def _hoja_unica(wb, nombre, titulo):
    """The tab called `titulo` (case/accent-insensitive), or None. Two of them is an error."""
    objetivo = norm_texto(titulo)
    halladas = [ws for ws in wb.worksheets if norm_texto(ws.title) == objetivo]
    if len(halladas) > 1:
        raise ValueError(
            f"'{nombre}' tiene más de una pestaña llamada '{titulo}' "
            f"({[ws.title for ws in halladas]}). No sé cuál leer."
        )
    return halladas[0] if halladas else None


def _hojas_del_template(wb, nombre):
    """([(ws, etiqueta)], [advertencias]): the tabs of a template to paste, in paste order, into the SAME 'Base'.
    'Reservas Manuales' is required (older templates: the only visible tab, never one that looks like rents);
    'Reservas Rentas' may be missing (reported) but a look-alike name is an ERROR; a hidden tab is read, with a warning."""
    avisos = []
    manuales = _hoja_unica(wb, nombre, HOJA_TEMPLATE)
    rentas = _hoja_unica(wb, nombre, HOJA_RENTAS)

    if manuales is None:
        candidatas = [
            ws for ws in wb.worksheets
            if ws.sheet_state == "visible" and ws is not rentas
            and "RENTA" not in norm_texto(ws.title)
        ]
        if len(candidatas) != 1:
            raise ValueError(
                f"No sé qué pestaña leer de '{nombre}': no hay una llamada "
                f"'{HOJA_TEMPLATE}' y hay {len(candidatas)} candidata(s) visible(s) "
                f"({[ws.title for ws in candidatas]})."
            )
        manuales = candidatas[0]

    if rentas is None:
        parecidas = [
            ws.title for ws in wb.worksheets
            if ws is not manuales and "RENTA" in norm_texto(ws.title)
        ]
        if parecidas:
            raise ValueError(
                f"En '{nombre}' hay una pestaña {parecidas} que parece la de rentas, "
                f"pero debe llamarse '{HOJA_RENTAS}' para traspasarla. Renómbrala: "
                f"no adivino si es la correcta."
            )
        avisos.append(
            f"{nombre}: no trae la pestaña '{HOJA_RENTAS}'; sólo se traspasó "
            f"'{manuales.title}'."
        )

    hojas = [(manuales, "Manuales")]
    if rentas is not None:
        hojas.append((rentas, "Rentas"))

    for ws, _ in hojas:
        if ws.sheet_state != "visible":
            avisos.append(f"{nombre}: la pestaña '{ws.title}' está oculta; se leyó de todos modos.")
    return hojas, avisos


def _hoja_en_blanco(ws):
    """True when nothing visible sits in the first MAX_FILAS_EN_BLANCO rows/columns (bounded: formatting can span the sheet)."""
    return not any(
        not _es_vacio(valor)
        for fila in ws.iter_rows(min_row=1, max_row=min(ws.max_row, MAX_FILAS_EN_BLANCO),
                                 max_col=min(ws.max_column, MAX_FILAS_EN_BLANCO),
                                 values_only=True)
        for valor in fila
    )


def _localizar_encabezado_template(ws, nombre):
    """(fila_encabezado, {normalized header: column}) of ONE tab, found by content (all mapped headers), not position."""
    requeridos = {norm_texto(t) for _, t in MAPEO_COLUMNAS}

    for fila, valores in enumerate(
        ws.iter_rows(min_row=1, max_row=MAX_FILAS_ENCABEZADO, values_only=True), start=1
    ):
        columnas = {}
        for columna, valor in enumerate(valores, start=1):
            clave = norm_texto(valor)
            if clave:
                columnas.setdefault(clave, []).append(columna)
        if requeridos <= set(columnas):
            repetidas = sorted(c for c in requeridos if len(columnas[c]) > 1)
            if repetidas:
                raise ValueError(
                    f"En '{nombre}' (pestaña '{ws.title}') estos encabezados aparecen más "
                    f"de una vez en la fila {fila}: {repetidas}. No sé cuál columna usar."
                )
            return fila, {c: columnas[c][0] for c in requeridos}

    raise ValueError(
        f"No encontré la fila de encabezados en '{nombre}' (pestaña '{ws.title}', primeras "
        f"{MAX_FILAS_ENCABEZADO} filas). Debe traer: "
        f"{sorted({t for _, t in MAPEO_COLUMNAS})}"
    )


def _grupo_del_titulo(ws, fila_encabezado, nombre, grupos):
    """Index of the channel block named by a tab's TITLE (text above its header), or None. Two channels is an error."""
    if fila_encabezado <= 1:
        return None
    del_titulo = set()
    for valores in ws.iter_rows(min_row=1, max_row=fila_encabezado - 1, values_only=True):
        for valor in valores:
            if isinstance(valor, str):
                indice = resolver_grupo(valor, grupos)
                if indice is not None:
                    del_titulo.add(indice)

    if len(del_titulo) > 1:
        raise ValueError(
            f"El título de la pestaña '{ws.title}' de '{nombre}' menciona más de un canal "
            f"({sorted(grupos[i]['canal'] for i in del_titulo)}); no sé cuál es."
        )
    return next(iter(del_titulo)) if del_titulo else None


def _grupo_del_archivo(titulos, nombre, grupos):
    """(indice, fuente) of the channel block a template FILE belongs to: its tab TITLES cross-checked with the FILE NAME.
    `titulos` is [(tab title, channel index | None)]. Raises ValueError when the channel cannot be established or sources disagree."""
    distintos = {i for _, i in titulos if i is not None}
    if len(distintos) > 1:
        detalle = ", ".join(
            f"'{hoja}': {grupos[i]['canal']}" for hoja, i in titulos if i is not None
        )
        raise ValueError(
            f"Las pestañas de '{nombre}' dicen canales distintos ({detalle}). "
            f"¿Es el archivo correcto?"
        )

    titulo = next(iter(distintos)) if distintos else None
    del_nombre = resolver_grupo(os.path.splitext(nombre)[0], grupos)

    if titulo is not None and del_nombre is not None and titulo != del_nombre:
        raise ValueError(
            f"'{nombre}' dice {grupos[del_nombre]['canal']} en el nombre del archivo pero "
            f"{grupos[titulo]['canal']} en su título. ¿Es el archivo correcto?"
        )
    if titulo is not None:
        return titulo, "título"
    if del_nombre is not None:
        return del_nombre, "nombre del archivo"
    raise ValueError(
        f"No pude saber de qué canal es '{nombre}': ni su título ni su nombre "
        f"mencionan {[g['canal'] for g in grupos]}."
    )


def _leer_hoja(ws, etiqueta, fila_enc, columnas, canal, indice_grupo, grupos):
    """Read the data rows of ONE tab. Never raises for a bad cell: every problem comes back in 'errores'
    (prefixed with the tab name) so all tabs of all files are reported together.
    'CT' holds the channel of the FILE, not a cell."""
    col = {tpl: columnas[norm_texto(tpl)] for _, tpl in MAPEO_COLUMNAS}
    ancho = max(col.values())
    prefijo = f"[{ws.title}] "

    filas, errores = [], []
    ocultas, meses_vistos, avisos_fecha = [], {}, []
    reserva_vacia = []           # [(celda, 'vacía' | 'solo espacios en blanco')]

    for n, valores in enumerate(
        ws.iter_rows(min_row=fila_enc + 1, max_col=ancho, values_only=True),
        start=fila_enc + 1,
    ):
        leidos = {tpl: valores[idx - 1] for tpl, idx in col.items()}
        if all(_es_vacio(v) for v in leidos.values()):
            continue  # a really empty row

        if ws.row_dimensions[n].hidden:
            ocultas.append(n)

        # --- required cells -------------------------------------
        faltan = [t for t in OBLIGATORIAS_TEMPLATE if _es_vacio(leidos[t])]
        if "RESERVA ($)" in faltan:
            # Reported apart, by cell and all together: a lone space looks like an empty cell and vice versa.
            faltan.remove("RESERVA ($)")
            reserva_vacia.append(
                (f"{ws.title}!{get_column_letter(col['RESERVA ($)'])}{n}",
                 _como_esta_vacio(leidos["RESERVA ($)"]))
            )
            fila_incompleta = True
        else:
            fila_incompleta = False
        if faltan:
            errores.append(f"{prefijo}fila {n}: falta {', '.join(faltan)}")
            fila_incompleta = True
        if fila_incompleta:
            continue

        fila_ok = True
        try:
            mes_fila = clave_de_mes(leidos["MES"])
        except ValueError as e:
            errores.append(f"{prefijo}fila {n}: MES {e}")
            fila_ok = False
            mes_fila = None
        try:
            importe = importe_de(leidos["RESERVA ($)"])
        except ValueError as e:
            errores.append(f"{prefijo}fila {n}: RESERVA ($) {e}")
            fila_ok = False
            importe = None

        # --- channel of the row itself ---
        # CANAL is copied as it comes; it matters only when it names ANOTHER of the five channels (a row of a different file).
        celda_canal = leidos["CANAL"]
        if not _es_vacio(celda_canal):
            propio = resolver_grupo(str(celda_canal), grupos)
            if propio is not None and propio != indice_grupo:
                errores.append(
                    f"{prefijo}fila {n}: CANAL dice '{str(celda_canal).strip()}' "
                    f"({grupos[propio]['canal']}) pero el archivo es de {canal}"
                )
                fila_ok = False

        # --- text sanity ----------------------------------------
        valores_base, revisados = {}, set()
        for base_col, tpl in MAPEO_COLUMNAS:
            if base_col in ("Mes", "Importe", "Inicio Vigencia", "Término Vigencia"):
                continue
            valor = _limpiar(leidos[tpl])
            problema = _problema_de_texto(valor)
            # 'AGRUPADOR ABIERTO' feeds two Base columns: report it once.
            if problema and tpl not in revisados:
                errores.append(f"{prefijo}fila {n}: {tpl} {problema}")
                fila_ok = False
            revisados.add(tpl)
            valores_base[base_col] = valor

        if not fila_ok:
            continue

        # --- dates (informative columns, never fatal) -----------
        inicio, aviso_i = fecha_de(leidos["INICIO"])
        fin, aviso_f = fecha_de(leidos["FIN"])
        for aviso in (aviso_i, aviso_f):
            if aviso:
                avisos_fecha.append(f"{prefijo}fila {n}: {aviso}")

        valores_base.update({
            "Mes":              mes_fila,
            COL_BASE_CANAL:     canal,      # 'CT': the channel of the file
            "Importe":          importe,
            "Inicio Vigencia":  inicio,
            "Término Vigencia": fin,
        })
        # Agrupador/Descripción Motivo share a source and it was validated above.
        meses_vistos[mes_fila] = meses_vistos.get(mes_fila, 0) + 1
        filas.append({"fila": n, "hoja": ws.title, "valores": valores_base})

    return {
        "hoja":            ws.title,
        "etiqueta":        etiqueta,
        "fila_encabezado": fila_enc,
        "filas":           filas,
        "errores":         errores,
        "ocultas":         ocultas,
        "reserva_vacia":   reserva_vacia,
        "meses":           meses_vistos,
        "avisos_fecha":    avisos_fecha,
    }


def leer_template(ruta, clave_esperada, grupos=None):
    """Read and validate ONE template (all its tabs); raises ValueError with every problem found, not just the first.
    Returns {ruta, nombre, hoja, hojas, fila_encabezado, grupo, canal, fuente_canal, total, advertencias,
    filas: [{'fila', 'hoja', 'valores': {Base column: value}}]}; 'hojas' is [{'hoja', 'etiqueta', 'filas', 'total'}]."""
    grupos = grupos if grupos is not None else GRUPOS
    ruta = os.path.abspath(ruta)
    nombre = os.path.basename(ruta)

    if nombre.startswith("~$"):
        raise ValueError(
            f"'{nombre}' es el archivo temporal que Excel crea mientras el libro "
            f"está abierto, no el template. Elige el archivo sin '~$'."
        )
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta}")

    try:
        wb = load_workbook(ruta, data_only=True)
    except PermissionError:
        raise ValueError(f"No pude abrir '{nombre}': el sistema lo tiene bloqueado.")
    except Exception as e:
        raise ValueError(f"No pude abrir '{nombre}' como libro de Excel (.xlsx): {e}")

    try:
        hojas, advertencias = _hojas_del_template(wb, nombre)
        advertencias = list(advertencias)

        # A completely blank tab is skipped (and reported); one with a header but no rows is read and gives 0 rows.
        con_datos = []
        for ws, etiqueta in hojas:
            if _hoja_en_blanco(ws):
                advertencias.append(
                    f"{nombre}: la pestaña '{ws.title}' está en blanco; no se traspasó nada de ella."
                )
                continue
            fila_enc, columnas = _localizar_encabezado_template(ws, nombre)
            con_datos.append((ws, etiqueta, fila_enc, columnas))
        if not con_datos:
            raise ValueError(
                f"'{nombre}' no trae datos: sus pestañas "
                f"{[ws.title for ws, _ in hojas]} están en blanco."
            )

        titulos = [
            (ws.title, _grupo_del_titulo(ws, fila_enc, nombre, grupos))
            for ws, _, fila_enc, _ in con_datos
        ]
        indice_grupo, fuente = _grupo_del_archivo(titulos, nombre, grupos)
        canal = grupos[indice_grupo]["canal"]

        leidas = [
            _leer_hoja(ws, etiqueta, fila_enc, columnas, canal, indice_grupo, grupos)
            for ws, etiqueta, fila_enc, columnas in con_datos
        ]

        # --- file-level checks (across tabs) --------------------------
        errores = [e for h in leidas for e in h["errores"]]
        filas = [f for h in leidas for f in h["filas"]]

        reserva_vacia = [c for h in leidas for c in h["reserva_vacia"]]
        if reserva_vacia:
            MAX_CELDAS = 30
            detalle = ", ".join(f"{c} ({como})" for c, como in reserva_vacia[:MAX_CELDAS])
            if len(reserva_vacia) > MAX_CELDAS:
                detalle += f", ...y {len(reserva_vacia) - MAX_CELDAS} más"
            errores.insert(0,
                f"RESERVA ($) sin importe en {len(reserva_vacia)} celda(s): {detalle}. "
                f"Escribe el importe (0 si de verdad es cero) o borra esa fila. Si la celda "
                f"es una fórmula, abre el archivo en Excel y guárdalo para que se calcule."
            )
        for h in leidas:
            if h["ocultas"]:
                ocultas = h["ocultas"]
                errores.append(
                    f"[{h['hoja']}] hay {len(ocultas)} fila(s) con datos OCULTAS (¿filtro aplicado?): "
                    f"{', '.join(map(str, ocultas[:MAX_ERRORES_LISTADOS]))}"
                    f"{'...' if len(ocultas) > MAX_ERRORES_LISTADOS else ''}. "
                    f"Muéstralas o elimínalas; no adivino cuáles querías pegar."
                )
        if not filas and not errores:
            errores.append(
                f"no trae filas de datos debajo de los encabezados "
                f"(pestañas leídas: {[h['hoja'] for h in leidas]})"
            )
        if not PERMITIR_OTROS_MESES:
            meses_vistos = {}
            for h in leidas:
                for mes, k in h["meses"].items():
                    meses_vistos[mes] = meses_vistos.get(mes, 0) + k
            otros = {m: k for m, k in meses_vistos.items() if m != clave_esperada}
            if otros:
                detalle = ", ".join(f"{m} ({k} fila(s))" for m, k in sorted(otros.items()))
                errores.append(
                    f"trae meses distintos al que se está procesando ({clave_esperada}): "
                    f"{detalle}. Si es el mes correcto, escríbelo en el campo 'Mes' del "
                    f"botón."
                )

        if errores:
            raise ValueError(
                f"'{nombre}' ({len(errores)} problema(s)):\n{_listar(errores)}"
            )

        if fuente != "título":
            advertencias.append(
                f"{nombre}: su título no menciona el canal; se tomó del {fuente} ({canal})."
            )
        avisos_fecha = [a for h in leidas for a in h["avisos_fecha"]]
        if avisos_fecha:
            advertencias.append(
                f"{nombre}: {len(avisos_fecha)} fecha(s) de vigencia sin formato "
                f"reconocido, p. ej. {avisos_fecha[0]}"
            )
        for h in leidas:
            if not h["filas"]:
                advertencias.append(
                    f"{nombre}: la pestaña '{h['hoja']}' no trae filas de datos."
                )

        return {
            "ruta":            ruta,
            "nombre":          nombre,
            "hoja":            leidas[0]["hoja"],
            "hojas": [
                {"hoja": h["hoja"], "etiqueta": h["etiqueta"], "filas": len(h["filas"]),
                 "total": round(sum(f["valores"]["Importe"] for f in h["filas"]), 2)}
                for h in leidas
            ],
            "fila_encabezado": leidas[0]["fila_encabezado"],
            "grupo":           indice_grupo,
            "canal":           canal,
            "fuente_canal":    fuente,
            "filas":           filas,
            "total":           round(sum(f["valores"]["Importe"] for f in filas), 2),
            "advertencias":    advertencias,
        }
    finally:
        wb.close()


# --- Planning the paste (pure: reads, validates, writes nothing) ---
def planificar_pegado(df_base, fila_encabezado, ultima_fila, rutas, clave_actual, grupos=None):
    """Read every template, validate the lot and work out what the paste will do; writes nothing.
    df_base, fila_encabezado and ultima_fila come from Consolidacion.leer_base(); clave_actual is the month being run.
    Returns the plan dict used by aplicar_plan_a_df / aplicar_plan_a_libro / verificar_pegado / resumir_pegado."""
    grupos = grupos if grupos is not None else GRUPOS

    # --- the selection itself ------------------------------------------
    limpias, vistas = [], {}
    for ruta in rutas or []:
        if _es_vacio(ruta):
            continue
        absoluta = os.path.abspath(str(ruta).strip().strip('"'))
        clave = os.path.normcase(absoluta)
        if clave in vistas:
            raise ValueError(f"El template '{os.path.basename(absoluta)}' está repetido en la selección.")
        vistas[clave] = absoluta
        limpias.append(absoluta)
    if not limpias:
        raise ValueError("No se eligió ningún template.")

    # --- 'Base' must be pasteable --------------------------------------
    if ultima_fila > fila_encabezado and ultima_fila not in set(df_base["_nfila"]):
        raise ValueError(
            f"La última fila con contenido de '{HOJA_BASE}' (fila {ultima_fila}) no tiene "
            f"Mes/Agrupador: parece un total o una nota. No pego debajo de ella para no "
            f"taparla ni dejarla en medio de los datos."
        )

    # --- read all the templates, collect ALL the problems --------------
    archivos, problemas = [], []
    for ruta in limpias:
        try:
            archivos.append(leer_template(ruta, clave_actual, grupos))
        except ValueError as e:
            problemas.append(str(e))
    if problemas:
        raise ValueError(
            "Los templates no pasaron la validación; no se escribió nada.\n\n"
            + "\n\n".join(problemas)
        )

    por_grupo = {}
    for archivo in archivos:
        por_grupo.setdefault(archivo["grupo"], []).append(archivo["nombre"])
    repetidos = {grupos[g]["canal"]: n for g, n in por_grupo.items() if len(n) > 1}
    if repetidos:
        detalle = "; ".join(f"{c}: {', '.join(n)}" for c, n in repetidos.items())
        raise ValueError(
            f"Hay más de un template del mismo canal ({detalle}). Pegarlos juntos "
            f"duplicaría el importe; elige uno por canal."
        )

    archivos.sort(key=lambda a: a["grupo"])           # block order, deterministic
    pegados = {a["grupo"] for a in archivos}

    # --- rows of an OLDER paste: channel in 'Canal Venta', 'CT' still blank ---
    # Those rows are invisible to the (month, channel) lookup below, so a new paste would double the money: refuse instead.
    if "_canal_venta" in df_base.columns:
        sin_ct = df_base[(df_base["_mes"] == clave_actual) & (df_base["_canal"] == "")]
        viejas = {}
        for canal_venta, importe in zip(sin_ct["_canal_venta"], sin_ct["_importe"]):
            g = resolver_grupo(canal_venta, grupos)
            if g in pegados:
                n_, suma_ = viejas.get(g, (0, 0.0))
                viejas[g] = (n_ + 1, suma_ + importe)
        if viejas:
            detalle = "; ".join(
                f"{grupos[g]['canal']}: {n_} fila(s), Σ {suma_:,.2f}"
                for g, (n_, suma_) in sorted(viejas.items())
            )
            raise ValueError(
                f"'{HOJA_BASE}' ya trae {clave_actual} de estos canales con '{COL_BASE_CANAL}' "
                f"vacío ({detalle}); parecen de un traspaso anterior, que escribía el canal en "
                f"'Canal Venta'. Si sigo, el importe se duplicaría. Llena '{COL_BASE_CANAL}' con el "
                f"canal en esas filas (o bórralas) y vuelve a correr; no se escribió nada."
            )

    # --- what already sits in 'Base' for those (month, channel)? -------
    df = asignar_grupos(df_base, grupos)
    a_reemplazar = df[(df["_mes"] == clave_actual) & (df["_grupo"].isin(pegados))]
    reemplazo_por_canal = {
        grupos[g]["canal"]: int((a_reemplazar["_grupo"] == g).sum()) for g in sorted(pegados)
    }

    filas_nuevas, origen = [], []
    for archivo in archivos:
        for f in archivo["filas"]:
            filas_nuevas.append(f["valores"])
            origen.append((archivo["nombre"], f["hoja"], f["fila"]))

    advertencias = [a for archivo in archivos for a in archivo["advertencias"]]
    faltan = [g["canal"] for i, g in enumerate(grupos) if i not in pegados]
    if faltan:
        advertencias.append(
            f"Se pegaron {len(pegados)} de {len(grupos)} canales; no se tocó nada de: "
            f"{', '.join(faltan)}."
        )

    return {
        "mes":                 clave_actual,
        "fila_encabezado":     fila_encabezado,
        "ultima_fila":         ultima_fila,
        "archivos":            archivos,
        "filas_nuevas":        filas_nuevas,
        "origen_filas":        origen,
        "filas_a_reemplazar":  sorted(int(x) for x in a_reemplazar["_nfila"]),
        "reemplazo_por_canal": reemplazo_por_canal,
        "total_nuevo":         round(sum(v["Importe"] for v in filas_nuevas), 2),
        "total_reemplazado":   round(float(a_reemplazar["_importe"].sum()), 2),
        "advertencias":        advertencias,
    }


def previsualizar_pegado(ruta_libro, rutas_templates, mes=None, grupos=None):
    """Plan the paste against the real workbook WITHOUT writing: what the confirmation box shows. Returns the plan."""
    _, clave_actual = claves_mes_actual_y_anterior(mes)
    df, fila_encabezado, ultima_fila = leer_base(ruta_libro)
    return planificar_pegado(df, fila_encabezado, ultima_fila, rutas_templates,
                             clave_actual, grupos)


def aplicar_plan_a_df(df, plan):
    """The working DataFrame as it will look after the paste (replaced rows gone, new ones appended): the matrices' source."""
    quitar = set(plan["filas_a_reemplazar"])
    conservadas = df[~df["_nfila"].isin(quitar)]

    nuevas = pd.DataFrame({
        "_nfila":     -1,   # not in the sheet yet
        "_mes":       [v["Mes"] for v in plan["filas_nuevas"]],
        "_agrupador": [str(v["Agrupador"]).strip() for v in plan["filas_nuevas"]],
        "_importe":   [v["Importe"] for v in plan["filas_nuevas"]],
        "_canal":     [_normalizar(v[COL_BASE_CANAL]) for v in plan["filas_nuevas"]],
        "_canal_venta": [_normalizar(v["Canal Venta"]) for v in plan["filas_nuevas"]],
    })
    return pd.concat([conservadas, nuevas], ignore_index=True)


# --- Applying the plan to the workbook (openpyxl, on the Output copy) ---
def _eliminar_filas(ws, filas, ultima_col, ultima_fila):
    """Remove sheet rows `filas`, closing the gaps; returns the new last row. ONE top-down pass moves each kept run once
    (move_range translate=True re-points formulas; delete_rows would not, and a per-block tail move was quadratic)."""
    borradas = sorted(set(filas))
    if not borradas:
        return ultima_fila

    letra = get_column_letter(ultima_col)
    for i, fila in enumerate(borradas):
        inicio = fila + 1
        fin = borradas[i + 1] - 1 if i + 1 < len(borradas) else ultima_fila
        if fin >= inicio:
            ws.move_range(f"A{inicio}:{letra}{fin}", rows=-(i + 1), translate=True)

    nueva_ultima = ultima_fila - len(borradas)

    # Blank the rows below the new last row: they may still hold deleted rows that nothing moved onto.
    for r in range(nueva_ultima + 1, ultima_fila + 1):
        for c in range(1, ultima_col + 1):
            celda = ws.cell(row=r, column=c)
            celda.value = None
            celda.style = "Normal"
    return nueva_ultima


def _ajustar_rangos(ws, fila_encabezado, ultima_fila):
    """Stretch Excel Tables / the sheet autofilter that start at the header row."""
    def estirar(ref):
        min_col, min_row, max_col, _ = range_boundaries(ref)
        if min_row != fila_encabezado:
            return None
        return (f"{get_column_letter(min_col)}{min_row}:"
                f"{get_column_letter(max_col)}{max(ultima_fila, min_row + 1)}")

    for tabla in list(ws.tables.values()):
        nuevo = estirar(tabla.ref)
        if nuevo:
            tabla.ref = nuevo
            if tabla.autoFilter is not None:
                tabla.autoFilter.ref = nuevo
    if ws.auto_filter.ref:
        nuevo = estirar(ws.auto_filter.ref)
        if nuevo:
            ws.auto_filter.ref = nuevo


def aplicar_plan_a_libro(wb, plan):
    """Replace the planned rows of 'Base' and append the new ones in the open openpyxl workbook `wb` (a COPY; the caller saves).
    Returns the info dict verificar_pegado() and resumir_pegado() need."""
    if HOJA_BASE not in wb.sheetnames:
        raise ValueError(f"El archivo no tiene una pestaña llamada '{HOJA_BASE}'.")
    ws = wb[HOJA_BASE]

    fila_enc = _localizar_fila_encabezado(ws)
    if fila_enc != plan["fila_encabezado"]:
        raise RuntimeError(
            f"La fila de encabezados de '{HOJA_BASE}' cambió entre la lectura "
            f"({plan['fila_encabezado']}) y la escritura ({fila_enc}). No se escribió nada."
        )

    ultima_col = ws.max_column
    if ultima_col > MAX_COLUMNAS_BASE:
        raise ValueError(f"'{HOJA_BASE}' tiene {ultima_col} columnas; no es la hoja esperada.")

    # --- where do the mapped columns live? ------------------------------
    por_nombre = {}
    for c in range(1, ultima_col + 1):
        clave = norm_texto(ws.cell(row=fila_enc, column=c).value)
        if clave:
            por_nombre.setdefault(clave, []).append(c)

    columnas = {}
    for base_col in [b for b, _ in MAPEO_COLUMNAS] + [COL_BASE_CANAL]:
        encontradas = por_nombre.get(norm_texto(base_col), [])
        if len(encontradas) != 1:
            raise ValueError(
                f"Necesito exactamente una columna '{base_col}' en '{HOJA_BASE}' y encontré "
                f"{len(encontradas)}. Columnas: {[k for k in por_nombre]}"
            )
        columnas[base_col] = encontradas[0]

    # A column the workbook marks as 'Formula' must not be overwritten with values.
    if fila_enc > 1:
        con_formula = [
            b for b, c in columnas.items()
            if norm_texto(ws.cell(row=fila_enc - 1, column=c).value) == "FORMULA"
        ]
        if con_formula:
            raise ValueError(
                f"En '{HOJA_BASE}' las columnas {con_formula} están marcadas como 'Fórmula'. "
                f"Pegar valores encima las rompería; no se escribió nada."
            )

    # --- 1. remove the rows being replaced -----------------------------
    ultima = _eliminar_filas(ws, plan["filas_a_reemplazar"], ultima_col, plan["ultima_fila"])

    # --- 2. append the new ones ----------------------------------------
    inicio = ultima + 1
    n = len(plan["filas_nuevas"])
    fin = ultima + n

    for r in range(inicio, fin + 1):
        for base_col, c in columnas.items():
            if ws.cell(row=r, column=c).value is not None:
                raise ValueError(
                    f"La celda {get_column_letter(c)}{r} de '{HOJA_BASE}' ya tiene contenido; "
                    f"no la sobrescribo. Revisa que no haya datos debajo de la tabla."
                )

    fila_modelo = ultima if ultima > fila_enc else None   # last data row = style/formula model
    advertencias = []
    if fila_modelo:
        con_formula_modelo = [
            b for b, c in columnas.items()
            if isinstance(ws.cell(row=fila_modelo, column=c).value, str)
            and ws.cell(row=fila_modelo, column=c).value.startswith("=")
        ]
        if con_formula_modelo:
            advertencias.append(
                f"En la última fila de '{HOJA_BASE}' las columnas {con_formula_modelo} "
                f"traen fórmula; los renglones pegados llevan valores."
            )

    destino = {c: b for b, c in columnas.items()}
    for k, valores in enumerate(plan["filas_nuevas"]):
        r = inicio + k
        for c in range(1, ultima_col + 1):
            celda = ws.cell(row=r, column=c)
            modelo = ws.cell(row=fila_modelo, column=c) if fila_modelo else None
            if modelo is not None and modelo.has_style:
                celda._style = copy(modelo._style)

            if c in destino:
                valor = valores[destino[c]]
                celda.value = valor
                if isinstance(valor, str):
                    # A description starting with '=' must stay TEXT (openpyxl would store a formula).
                    celda.data_type = "s"
            elif (celda.value is None and modelo is not None
                  and isinstance(modelo.value, str) and modelo.value.startswith("=")):
                # Unmapped column with a formula in the model row (e.g. the untitled column M): carry it down, re-pointed.
                celda.value = Translator(
                    modelo.value, origin=modelo.coordinate
                ).translate_formula(celda.coordinate)

    _ajustar_rangos(ws, fila_enc, fin)

    return {
        "fila_inicio":  inicio,
        "fila_fin":     fin,
        "borradas":     len(plan["filas_a_reemplazar"]),
        "columnas":     columnas,
        "advertencias": advertencias,
    }


# --- Verification and summary ---
def _coincide(esperado, leido):
    """Does the value read back from disk match what the plan said to write?"""
    if esperado is None:
        return _es_vacio(leido)
    if isinstance(esperado, bool):
        return leido is esperado
    if isinstance(esperado, (int, float)):
        return (isinstance(leido, (int, float)) and not isinstance(leido, bool)
                and math.isclose(leido, esperado, rel_tol=1e-12, abs_tol=1e-9))
    if isinstance(esperado, datetime):
        return isinstance(leido, datetime) and leido == esperado
    return leido == esperado


def verificar_pegado(ruta, plan, info):
    """Re-open the SAVED file and compare the pasted block with the plan cell by cell, plus the control total (RuntimeError on any difference)."""
    wb = load_workbook(ruta, read_only=True, data_only=False)
    try:
        ws = wb[HOJA_BASE]
        filas = list(ws.iter_rows(min_row=info["fila_inicio"], max_row=info["fila_fin"],
                                  values_only=True))
    finally:
        wb.close()

    esperadas = plan["filas_nuevas"]
    if len(filas) != len(esperadas):
        raise RuntimeError(
            f"Verificación fallida: esperaba {len(esperadas)} fila(s) pegadas y el archivo "
            f"trae {len(filas)} en las filas {info['fila_inicio']}–{info['fila_fin']}."
        )

    diferencias, total_leido = [], 0.0
    for k, (esperado, leida) in enumerate(zip(esperadas, filas)):
        for base_col, c in info["columnas"].items():
            leido = leida[c - 1] if c - 1 < len(leida) else None
            if not _coincide(esperado[base_col], leido):
                archivo, hoja_t, fila_t = plan["origen_filas"][k]
                diferencias.append(
                    f"{base_col} (fila {info['fila_inicio'] + k} de Base, origen "
                    f"{archivo} [{hoja_t}] fila {fila_t}): esperaba {esperado[base_col]!r}, leí {leido!r}"
                )
        c_importe = info["columnas"]["Importe"]
        valor = leida[c_importe - 1] if c_importe - 1 < len(leida) else None
        if isinstance(valor, (int, float)) and not isinstance(valor, bool):
            total_leido += valor

    if diferencias:
        raise RuntimeError(
            f"Verificación fallida: {len(diferencias)} celda(s) no coinciden con los "
            f"templates.\n{_listar(diferencias)}"
        )
    if abs(total_leido - plan["total_nuevo"]) > 0.01:
        raise RuntimeError(
            f"Verificación fallida: Σ Importe pegado = {total_leido:,.2f} y los templates "
            f"suman {plan['total_nuevo']:,.2f}."
        )


def resumir_pegado(plan, info=None):
    """Plain-data summary of a paste (info=None: of a PLANNED paste) for the popups; holds no workbook objects."""
    resumen = {
        "mes":                 plan["mes"],
        "archivos": [
            {"nombre": a["nombre"], "canal": a["canal"], "filas": len(a["filas"]),
             "total": a["total"], "fuente_canal": a["fuente_canal"],
             "hojas": [dict(h) for h in a["hojas"]]}
            for a in plan["archivos"]
        ],
        "filas_pegadas":       len(plan["filas_nuevas"]),
        "total_pegado":        plan["total_nuevo"],
        "filas_reemplazadas":  len(plan["filas_a_reemplazar"]),
        "total_reemplazado":   plan["total_reemplazado"],
        "reemplazo_por_canal": dict(plan["reemplazo_por_canal"]),
        "advertencias":        list(plan["advertencias"]),
        "verificado":          False,
    }
    if info:
        resumen["fila_inicio"] = info["fila_inicio"]
        resumen["fila_fin"] = info["fila_fin"]
        resumen["advertencias"] += info["advertencias"]
        resumen["verificado"] = True
    return resumen


# --- The paste as a step of its own (the 'Traspasar templates' button) ---
_SUFIJO_TRASPASO = re.compile(r"_Reservas_M\d{2}-\d{2}(_\d+)?$")


def _ruta_copia_traspaso(ruta_libro, ruta_salida, clave_actual):
    """Where the pasted copy goes: '<name>_Reservas_M10-26.xlsx' in 'Output' (suffix never doubled).
    If it would land ON the source, '_2', '_3'... is used: the source is never overwritten."""
    if not ruta_salida:
        ruta_salida = carpeta_output_desde_input(os.path.dirname(os.path.abspath(ruta_libro)))
    os.makedirs(ruta_salida, exist_ok=True)

    base = _SUFIJO_TRASPASO.sub("", os.path.splitext(os.path.basename(ruta_libro))[0])
    origen = os.path.normcase(os.path.abspath(ruta_libro))
    candidata = os.path.join(ruta_salida, f"{base}_Reservas_{clave_actual}.xlsx")
    n = 2
    while os.path.normcase(os.path.abspath(candidata)) == origen:
        candidata = os.path.join(ruta_salida, f"{base}_Reservas_{clave_actual}_{n}.xlsx")
        n += 1
    return candidata


def traspasar_templates(ruta_libro, rutas_templates, mes=None, ruta_salida=None,
                        grupos=None, callback_status=None):
    """Paste the 'RESERVAS MANUALES' templates into 'Base' of a COPY of the Reservas workbook and verify it on disk
    (no 'Variaciones', no PivotTables: that is the other button). Returns resumir_pegado() plus 'ruta', the copy
    to feed the consolidation. Raises ValueError/FileNotFoundError/PermissionError; on ANY failure the copy is
    removed and the original untouched."""
    def estado(mensaje):
        if callback_status:
            callback_status(mensaje)

    if not rutas_templates:
        raise ValueError("No hay templates que traspasar.")

    _, clave_actual = claves_mes_actual_y_anterior(mes)

    estado(f"📄 Leyendo la pestaña '{HOJA_BASE}'...")
    df, fila_encabezado, ultima_fila = leer_base(ruta_libro)

    estado(f"📥 Revisando {len(rutas_templates)} template(s) (Reservas Manuales y Reservas Rentas)...")
    plan = planificar_pegado(df, fila_encabezado, ultima_fila, rutas_templates,
                             clave_actual, grupos)
    for aviso in plan["advertencias"]:
        estado(f"⚠️ {aviso}")

    ruta_final = _ruta_copia_traspaso(ruta_libro, ruta_salida, clave_actual)
    shutil.copy2(ruta_libro, ruta_final)
    try:
        wb = load_workbook(ruta_final)
        try:
            estado(f"📥 Pegando los templates en la pestaña '{HOJA_BASE}'...")
            info = aplicar_plan_a_libro(wb, plan)
            wb.save(ruta_final)
        finally:
            wb.close()

        estado("🔎 Verificando lo pegado contra los templates...")
        verificar_pegado(ruta_final, plan, info)
        resumen = resumir_pegado(plan, info)
    except Exception:
        try:
            os.remove(ruta_final)
        except OSError:
            pass
        raise

    resumen["ruta"] = ruta_final
    estado("✅ Templates traspasados")
    return resumen


# --- Standalone run: python Reservas_Manuales.py (plans the paste and prints it; writes nothing) ---
if __name__ == "__main__":
    print("=== Prueba de pegado de Reservas Manuales (standalone, NO escribe) ===")
    ruta = input("Ruta del archivo de Reservas (.xlsx): ").strip().strip('"')
    mes_txt = input("Mes a correr [Enter = hoy, o p.ej. M10-26]: ").strip() or None
    rutas_txt = input("Templates, separados por ';': ")
    rutas = [r.strip().strip('"') for r in rutas_txt.split(";") if r.strip()]

    try:
        plan = previsualizar_pegado(ruta, rutas, mes=mes_txt)
        for clave, valor in resumir_pegado(plan).items():
            print(f"  {clave:<20}: {valor}")
    except Exception as e:
        print(f"❌ {e}")
        raise
