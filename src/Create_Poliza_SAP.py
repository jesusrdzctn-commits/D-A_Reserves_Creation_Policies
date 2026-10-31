"""RDA Pólizas D&A — fills 'PÓLIZA SAP' (R:Y) on BOTH tabs of the OT layout through Excel COM.
Columns are found by HEADER (never by letter); lookups use the layout's own 'Catalogo'.
Validates everything first, writes on a temp copy and renames at the end.
Formula spec and decisions: claude/hallazgos-layout-ot.md
"""

import math
import os
import shutil
import time
from datetime import datetime

from utils import abrir_excel as _abrir_excel
from utils import carpeta_output_desde_input
from utils import cerrar_excel as _cerrar_excel
from utils import error_de_excel as _error_de
from utils import es_vacio_excel as _vacio
from utils import hoja_por_nombre as _hoja
from utils import norm_texto as _norm

# Excel COM constants as plain numbers: win32com.client.constants breaks inside a PyInstaller .exe.
XL_UP                = -4162
XL_CELLTYPE_FORMULAS = -4123
XL_ERRORS            = 16
XL_CALC_MANUAL       = -4135
XL_CALC_AUTOMATIC    = -4105
XL_PASTE_VALUES      = -4163
XL_EXCEL_LINKS       = 1


# (tab name, short label, required). ORDER = processing order = row order in the 'Póliza Carga Batch'.
HOJAS_LAYOUT = (
    ("Reservas Manuales", "Manuales", True),
    ("Reservas Rentas",   "Rentas",   False),
)
NOMBRES_HOJAS_LAYOUT = tuple(nombre for nombre, _, _ in HOJAS_LAYOUT)

# Lookup tab; hoja_por_nombre() matches the exact name first ('Catalogo' vs 'Catálogo').
HOJA_CATALOGO = "Catalogo"

# How far down to look for the header row of a layout tab / of the catalogue.
FILAS_ENCABEZADO_MAX = 25
FILAS_ENCABEZADO_CATALOGO_MAX = 15
MAX_COLUMNAS = 200

# Safety net: going over it is an ERROR, never a silent truncation (a cut policy would look complete).
MAX_FILAS = 50_000

# Input columns read from each tab (key -> header): matched ignoring accents/case, must appear ONCE on the row.
ENTRADAS = {
    "bu":        "BU",
    "agrupador": "AGRUPADOR ABIERTO",
    "canal":     "CANAL",
    "id_cadena": "ID GPO CADENA",
    "reserva":   "RESERVA ($)",
}

# The eight 'PÓLIZA SAP' columns, left to right. 'prefijo': the header only has to START with 'encabezado'.
COLUMNAS_POLIZA = (
    {"clave": "cuenta",     "encabezado": "CUENTA SAP",                 "tipo": "formula"},
    {"clave": "cebe",       "encabezado": "CeBe",                       "tipo": "formula"},
    {"clave": "segmento",   "encabezado": "SEGMENTO",                   "tipo": "valor", "valor": "MXC0000"},
    {"clave": "asignacion", "encabezado": "ASIGNACION",                 "tipo": "formula"},
    {"clave": "ct",         "encabezado": "CT",                         "tipo": "valor", "valor": 40},
    {"clave": "mxn",        "encabezado": "MXN",                        "tipo": "formula"},
    {"clave": "redondeo",   "encabezado": "REDONDEO",                   "tipo": "formula"},
    {"clave": "concepto",   "encabezado": "CONCEPTO (DESCRIPCION SAP)", "tipo": "formula",
     "prefijo": "CONCEPTO"},
)

# Column that carries the AMOUNT (control total, and Debit/Credit source in the batch). REDONDEO is the
# last link of MXN -> REDONDEO; the batch warns when it differs from 'mxn'. Change here to use 'mxn'.
CLAVE_MONTO = "redondeo"

# Columns that decide how far the section goes: the HIGHEST last row wins, so one blank cell
# cannot cut the policy short.
COLUMNAS_PILOTO = ("agrupador", "bu", "reserva")

# Text between the two halves of CONCEPTO (account text + channel text).
SEPARADOR_CONCEPTO = " "

# name -> (SEARCH header, RESULT header); VLOOKUP only looks right, so RESULT must sit to the right.
BUSQUEDAS_CATALOGO = {
    "cuenta": ("AGRUPADOR D&A", "CUENTA POLIZA"),
    "texto":  ("AGRUPADOR D&A", "TEXTO POLIZA"),
    "cebe":   ("BU",            "CeBe"),
    "canal":  ("CANAL",         "CANAL POLIZA"),
}
# Headers that identify the catalogue's header row. 'BU' / 'CeBe' are left out
# on purpose: their absence must produce its own, specific message.
ANCLA_CATALOGO = ("AGRUPADOR D&A", "CUENTA POLIZA", "TEXTO POLIZA", "CANAL", "CANAL POLIZA")

# True: write formulas, let Excel calculate, then freeze them as values (auditable, and the SAP
# upload has no #N/A nor dependency on 'Catalogo').
CONVERTIR_A_VALORES = True

# True stops the run (the copy keeps live formulas) when a lookup fails; False writes anyway and reports.
# Either way the batch step refuses a layout that still has error cells.
ABORTAR_SI_HAY_ERRORES = False

# True writes CT as the text "40" instead of the number 40 (the workbook stores a number).
CT_COMO_TEXTO = False

# Clear the eight columns first so a shorter month leaves no stale rows (contents only, never above row 1 of data).
LIMPIAR_ANTES = True

# How many offending cells to list per problem in the summary.
MAX_ERRORES_REPORTADOS = 10

# Transient file-system failures (antivirus scan, a dying EXCEL.EXE still holding the lock).
INTENTOS = 4
ESPERA_SEG = 0.6


# --- Small helpers ---
def titulo_de(clave):
    """Header of one of the eight 'PÓLIZA SAP' columns, by key."""
    for spec in COLUMNAS_POLIZA:
        if spec["clave"] == clave:
            return spec["encabezado"]
    raise KeyError(clave)


def _letra(idx):
    """1-based column index -> letters (1 -> 'A', 27 -> 'AA')."""
    texto = ""
    while idx > 0:
        idx, resto = divmod(idx - 1, 26)
        texto = chr(65 + resto) + texto
    return texto


def _ref_hoja(nombre):
    """Sheet name as a formula reference. Always quoted: valid for ANY name."""
    return "'" + nombre.replace("'", "''") + "'"


def _leer_columna(ws, idx, fila_ini, fila_fin):
    """Value2 of ONE column as a flat list (a single cell returns a scalar: the classic COM trap)."""
    n = fila_fin - fila_ini + 1
    valor = ws.Range(ws.Cells(fila_ini, idx), ws.Cells(fila_fin, idx)).Value2
    if n == 1:
        return [valor]
    return [fila[0] for fila in valor]


def _leer_bloque(ws, filas):
    """Value2 of the first `filas` rows as a tuple of row tuples (always 2-D since filas > 1)."""
    usado = ws.UsedRange
    ultima_col = int(usado.Column + usado.Columns.Count - 1)
    ultima_col = max(1, min(ultima_col, MAX_COLUMNAS))
    return ws.Range(ws.Cells(1, 1), ws.Cells(filas, ultima_col)).Value2


def _columnas_de(fila, esperado, prefijo=False):
    """1-based columns of `fila` whose header equals (or starts with) `esperado`."""
    objetivo = _norm(esperado)
    halladas = []
    for i, celda in enumerate(fila, start=1):
        texto = _norm(celda)
        if texto and (texto.startswith(objetivo) if prefijo else texto == objetivo):
            halladas.append(i)
    return halladas


def _hoja_en_blanco(ws):
    """True when the tab holds nothing at all (a used range of one empty cell)."""
    usado = ws.UsedRange
    if int(usado.Rows.Count) * int(usado.Columns.Count) != 1:
        return False
    return _vacio(usado.Value2)


# --- Layout tabs: headers, columns, rows ---
def _encabezados_requeridos():
    """[(key, header, starts_with)] for every column a tab must carry."""
    requeridos = [(clave, enc, False) for clave, enc in ENTRADAS.items()]
    requeridos += [
        (spec["clave"], spec.get("prefijo") or spec["encabezado"], "prefijo" in spec)
        for spec in COLUMNAS_POLIZA
    ]
    return requeridos


def _encontrar_encabezados(ws, nombre):
    """(header row, {key: {'idx', 'letra', 'encabezado'}}) of one tab.
    The header row is FOUND (first row carrying every required header), not assumed.
    A header repeated on that row is an error: there is no safe way to guess."""
    requeridos = _encabezados_requeridos()
    bloque = _leer_bloque(ws, FILAS_ENCABEZADO_MAX)

    menos_faltantes = None
    for fila_idx, fila in enumerate(bloque, start=1):
        halladas, faltan = {}, []
        for clave, enc, prefijo in requeridos:
            cols = _columnas_de(fila, enc, prefijo)
            if cols:
                halladas[clave] = cols
            else:
                faltan.append(enc)

        if not faltan:
            repetidos = [enc for clave, enc, _ in requeridos if len(halladas[clave]) > 1]
            if repetidos:
                raise ValueError(
                    f"En la pestaña '{nombre}' estos encabezados aparecen más de una "
                    f"vez en la fila {fila_idx}: {repetidos}. No sé cuál columna usar."
                )
            return fila_idx, {
                clave: {
                    "idx": halladas[clave][0],
                    "letra": _letra(halladas[clave][0]),
                    "encabezado": enc,
                }
                for clave, enc, _ in requeridos
            }

        if menos_faltantes is None or len(faltan) < len(menos_faltantes):
            menos_faltantes = faltan

    raise ValueError(
        f"No encontré la fila de encabezados de la pestaña '{nombre}' en las primeras "
        f"{FILAS_ENCABEZADO_MAX} filas.\n\nMe faltan: {menos_faltantes}\n\n"
        f"¿Alguien cambió o movió una columna?"
    )


def _ultima_fila(ws, fila_inicio, indices):
    """Last row with data, or None: walks up from End(xlUp) over THREE pilot columns, keeping the highest.
    Content below MAX_FILAS is an error, never walked through or cut off."""
    mejor = None
    for idx in indices:
        fila = int(ws.Cells(ws.Rows.Count, idx).End(XL_UP).Row)
        if fila - fila_inicio + 1 > MAX_FILAS:
            raise ValueError(
                f"La pestaña '{ws.Name}' trae contenido hasta la fila {fila} (columna "
                f"{_letra(idx)}), más de {MAX_FILAS:,} renglones debajo de los "
                f"encabezados. Revisa que no haya basura al fondo de la hoja."
            )
        while fila >= fila_inicio:
            if not _vacio(ws.Cells(fila, idx).Value2):
                break
            fila -= 1
        if fila >= fila_inicio:
            mejor = fila if mejor is None else max(mejor, fila)
    return mejor


def localizar_seccion(ws, etiqueta=""):
    """Where the 'PÓLIZA SAP' section of ONE tab lives (dict with hoja, nombre, etiqueta, fila_encabezado,
    fila_inicio, fila_fin (None without data), filas, columnas).
    Raises ValueError on an active filter: Copy would take only VISIBLE rows and the policy would come out short."""
    nombre = str(ws.Name)
    if bool(ws.FilterMode):
        raise ValueError(
            f"La pestaña '{nombre}' tiene un filtro activo (hay filas ocultas por el "
            f"filtro).\n\nQuita el filtro y vuelve a intentar."
        )

    fila_enc, columnas = _encontrar_encabezados(ws, nombre)
    fila_inicio = fila_enc + 1
    fila_fin = _ultima_fila(ws, fila_inicio, [columnas[k]["idx"] for k in COLUMNAS_PILOTO])
    return {
        "hoja": ws,
        "nombre": nombre,
        "etiqueta": etiqueta or nombre,
        "fila_encabezado": fila_enc,
        "fila_inicio": fila_inicio,
        "fila_fin": fila_fin,
        "filas": 0 if fila_fin is None else fila_fin - fila_inicio + 1,
        "columnas": columnas,
    }


def localizar_secciones(wb):
    """Tabs of the layout that have data, in processing order: ([sections], [warnings]).
    'Reservas Manuales' must exist; a missing 'Reservas Rentas' is reported, a look-alike name is an ERROR.
    Blank tabs are skipped with a warning; no tab with rows at all is an error."""
    avisos, secciones = [], []
    nombres = [str(ws.Name) for ws in wb.Worksheets]

    for nombre, etiqueta, obligatoria in HOJAS_LAYOUT:
        try:
            ws = _hoja(wb, nombre)
        except ValueError:
            if obligatoria:
                raise
            parecidas = [n for n in nombres if "RENTA" in _norm(n)]
            if parecidas:
                raise ValueError(
                    f"Hay una pestaña {parecidas} que parece la de rentas, pero debe "
                    f"llamarse '{nombre}'. Renómbrala: no adivino si es la correcta."
                )
            avisos.append(f"El libro no trae la pestaña '{nombre}'; sólo se procesó el resto.")
            continue

        if _hoja_en_blanco(ws):
            avisos.append(f"La pestaña '{nombre}' está en blanco; no se procesó.")
            continue

        seccion = localizar_seccion(ws, etiqueta)
        if seccion["fila_fin"] is None:
            avisos.append(
                f"La pestaña '{nombre}' no trae renglones de datos debajo de los "
                f"encabezados (fila {seccion['fila_encabezado']}); no se procesó."
            )
            continue
        secciones.append(seccion)

    if not secciones:
        raise ValueError(
            "No encontré renglones de datos en el layout:\n\n"
            + "\n".join(f"     • {a}" for a in avisos)
            + "\n\n¿Es el archivo correcto?"
        )
    return secciones, avisos


# --- 'Catalogo' tab ---
def localizar_catalogo(wb):
    """Lookups of the 'Catalogo' tab found by header: {name: {'rango', 'indice', ...}} plus '_hoja'/'_fila_encabezado'.
    Raises ValueError when a header is missing or repeated, or RESULT is not to the right of SEARCH."""
    ws = _hoja(wb, HOJA_CATALOGO)
    nombre = str(ws.Name)
    bloque = _leer_bloque(ws, FILAS_ENCABEZADO_CATALOGO_MAX)

    fila_enc, fila_valores = None, None
    for i, fila in enumerate(bloque, start=1):
        if all(_columnas_de(fila, h) for h in ANCLA_CATALOGO):
            fila_enc, fila_valores = i, fila
            break
    if fila_enc is None:
        raise ValueError(
            f"No encontré la fila de encabezados de la pestaña '{nombre}' (primeras "
            f"{FILAS_ENCABEZADO_CATALOGO_MAX} filas). Debe traer: {list(ANCLA_CATALOGO)}"
        )

    busquedas = {}
    for clave, (h_busca, h_resultado) in BUSQUEDAS_CATALOGO.items():
        cols_busca = _columnas_de(fila_valores, h_busca)
        cols_result = _columnas_de(fila_valores, h_resultado)

        if not cols_busca or not cols_result:
            faltan = [h for h, c in ((h_busca, cols_busca), (h_resultado, cols_result)) if not c]
            if clave == "cebe":
                raise ValueError(
                    f"La pestaña '{nombre}' no trae la tabla BU → CeBe: en su fila de "
                    f"encabezados ({fila_enc}) me faltan {faltan}.\n\n"
                    f"La columna 'CeBe' del layout necesita esa tabla en el MISMO libro "
                    f"(el ejemplo de la fórmula apuntaba a un libro externo, y un vínculo "
                    f"así no se puede auditar). Agrega a '{nombre}' dos columnas con los "
                    f"encabezados 'BU' y 'CeBe' (el CeBe a la derecha del BU) en la fila "
                    f"{fila_enc} y vuelve a correr."
                )
            raise ValueError(
                f"La pestaña '{nombre}' no trae el/los encabezado(s) {faltan} en la fila "
                f"{fila_enc}. ¿Se renombró o se quitó una columna?"
            )
        if len(cols_busca) > 1 or len(cols_result) > 1:
            raise ValueError(
                f"En la pestaña '{nombre}' (fila {fila_enc}) el encabezado "
                f"'{h_busca if len(cols_busca) > 1 else h_resultado}' aparece más de una "
                f"vez. No sé cuál columna usar."
            )

        c_busca, c_result = cols_busca[0], cols_result[0]
        if c_result <= c_busca:
            raise ValueError(
                f"En la pestaña '{nombre}' la columna '{h_resultado}' ({_letra(c_result)}) "
                f"debe quedar a la DERECHA de '{h_busca}' ({_letra(c_busca)}): VLOOKUP sólo "
                f"busca hacia la derecha."
            )
        busquedas[clave] = {
            "rango":     f"{_ref_hoja(nombre)}!${_letra(c_busca)}:${_letra(c_result)}",
            "indice":    c_result - c_busca + 1,
            "clave":     h_busca,
            "valor":     h_resultado,
            "col_clave": c_busca,
            "col_valor": c_result,
        }

    busquedas["_hoja"] = ws
    busquedas["_fila_encabezado"] = fila_enc
    return busquedas


def validar_catalogo(catalogo):
    """Refuse a catalogue where one key returns two different results (VLOOKUP silently takes the first).
    Case is ignored like VLOOKUP; identical duplicates are harmless. Raises ValueError listing the keys."""
    ws = catalogo["_hoja"]
    fila_ini = catalogo["_fila_encabezado"] + 1
    problemas = []

    for clave, b in catalogo.items():
        if clave.startswith("_"):
            continue
        fila_fin = int(ws.Cells(ws.Rows.Count, b["col_clave"]).End(XL_UP).Row)
        if fila_fin < fila_ini:
            problemas.append(f"'{b['clave']}' no trae ningún renglón debajo de su encabezado")
            continue
        claves = _leer_columna(ws, b["col_clave"], fila_ini, fila_fin)
        valores = _leer_columna(ws, b["col_valor"], fila_ini, fila_fin)

        vistos = {}
        for k, v in zip(claves, valores):
            if _vacio(k):
                continue
            vistos.setdefault(str(k).upper(), set()).add(repr(v))
        ambiguos = sorted(k for k, vs in vistos.items() if len(vs) > 1)
        if ambiguos:
            lista = ", ".join(repr(a) for a in ambiguos[:MAX_ERRORES_REPORTADOS])
            problemas.append(
                f"'{b['clave']}' → '{b['valor']}': {len(ambiguos)} valor(es) repetido(s) "
                f"con resultados distintos ({lista}"
                f"{', ...' if len(ambiguos) > MAX_ERRORES_REPORTADOS else ''})"
            )

    if problemas:
        raise ValueError(
            f"La pestaña '{ws.Name}' no es confiable para los VLOOKUP; no se escribió "
            f"nada:\n\n" + "\n".join(f"     • {p}" for p in problemas)
        )


# --- Input validation (nothing is written until this passes) ---
def validar_entradas(seccion):
    """Check one tab's input cells BEFORE writing: all present, free of Excel errors, RESERVA ($) a finite number.
    A blank ID GPO CADENA still builds a valid-looking ASIGNACION, and a text amount becomes 0 via IFERROR.
    Returns a list of problems: {'hoja', 'titulo', 'tipo', 'total', 'celdas'}."""
    ws, nombre = seccion["hoja"], seccion["nombre"]
    ini, fin = seccion["fila_inicio"], seccion["fila_fin"]
    leidas = {
        clave: _leer_columna(ws, seccion["columnas"][clave]["idx"], ini, fin)
        for clave in ENTRADAS
    }

    grupos, filas_vacias = {}, []
    for i in range(fin - ini + 1):
        fila = ini + i
        if all(_vacio(leidas[clave][i]) for clave in ENTRADAS):
            filas_vacias.append(f"fila {fila}")
            continue
        for clave in ENTRADAS:
            valor = leidas[clave][i]
            col = seccion["columnas"][clave]
            celda = f"{nombre}!{col['letra']}{fila}"

            error = _error_de(valor)
            if error:
                tipo = f"error {error}"
            elif _vacio(valor):
                tipo = "vacío"
            elif isinstance(valor, bool):
                tipo = "tipo inesperado"
            elif clave == "reserva" and (
                not isinstance(valor, (int, float)) or not math.isfinite(valor)
            ):
                tipo = "no numérico"
            elif not isinstance(valor, (str, int, float)):
                tipo = "tipo inesperado"
            else:
                continue
            grupos.setdefault((col["encabezado"], tipo), []).append(celda)

    problemas = []
    for (titulo, tipo), celdas in grupos.items():
        problemas.append({"hoja": nombre, "titulo": titulo, "tipo": tipo,
                          "total": len(celdas), "celdas": celdas[:MAX_ERRORES_REPORTADOS]})
    if filas_vacias:
        problemas.append({"hoja": nombre, "titulo": "(renglón completo)",
                          "tipo": "vacío entre los datos (bórralo)",
                          "total": len(filas_vacias),
                          "celdas": filas_vacias[:MAX_ERRORES_REPORTADOS]})
    return problemas


def _texto_problemas_entradas(problemas):
    """Readable message for the input problems found in the layout."""
    total = sum(p["total"] for p in problemas)
    lineas = []
    for p in problemas:
        resto = p["total"] - len(p["celdas"])
        lineas.append(
            f"     • {p['hoja']} — {p['titulo']}: {p['total']} celda(s) con {p['tipo']}\n"
            f"       " + ", ".join(p["celdas"]) + (f", ...y {resto} más" if resto > 0 else "")
        )
    return (
        f"El layout tiene {total} dato(s) de entrada con problemas que producirían una "
        f"póliza incorrecta. No se escribió nada.\n\n"
        + "\n".join(lineas)
        + "\n\nCorrígelos en el layout y vuelve a correr."
    )


# --- Formulas ---
def construir_formulas(fila, letras, catalogo):
    """Formulas of the FIRST data row (`letras`: column key -> letter), assigned to the WHOLE column at once.
    .Formula is always US English (VLOOKUP, commas); .FormulaLocal would break on other languages."""
    def buscar(nombre, columna_clave):
        b = catalogo[nombre]
        return f"VLOOKUP(${letras[columna_clave]}{fila},{b['rango']},{b['indice']},0)"

    barra = '"\\"'                                  # the Excel text literal "\"
    sep = '"' + SEPARADOR_CONCEPTO + '"'
    bu = f"${letras['bu']}{fila}"
    canal = f"${letras['canal']}{fila}"
    id_cadena = f"${letras['id_cadena']}{fila}"
    reserva = f"${letras['reserva']}{fila}"
    mxn = f"${letras['mxn']}{fila}"

    return {
        "cuenta":     "=" + buscar("cuenta", "agrupador"),
        "cebe":       "=" + buscar("cebe", "bu"),
        "asignacion": f"=UPPER(LEFT({bu},3)&{barra}&LEFT({canal},2)&{barra}&{id_cadena})",
        "mxn":        f"=IFERROR(ROUND({reserva},2),0)",
        "redondeo":   f"=ROUND({mxn},2)",
        "concepto":   "=" + buscar("texto", "agrupador") + "&" + sep + "&" + buscar("canal", "canal"),
    }


# --- Excel operations ---
def _detectar_errores(ws, letra, fila_inicio, fila_fin):
    """(detalle, total) of one column's error cells: the first MAX_ERRORES_REPORTADOS offenders and the real count.
    SpecialCells(Formulas, Errors) is one call instead of N, and raises com_error when none: run BEFORE freezing."""
    rango = ws.Range(f"{letra}{fila_inicio}:{letra}{fila_fin}")
    try:
        erroneas = rango.SpecialCells(XL_CELLTYPE_FORMULAS, XL_ERRORS)
    except Exception:
        return [], 0

    detalle = []
    for celda in erroneas:
        detalle.append({
            "celda": str(celda.Address(False, False)),
            "error": str(celda.Text),
            "fila":  int(celda.Row),
        })
        if len(detalle) >= MAX_ERRORES_REPORTADOS:
            break

    return detalle, int(erroneas.Count)


def _congelar_a_valores(app, ws, letra, fila_inicio, fila_fin):
    """Replace one column's formulas with their results via Copy + PasteSpecial(values).
    NOT `rango.Value = rango.Value`: re-parsing '0004209900' as typed text drops the leading zeros."""
    rango = ws.Range(f"{letra}{fila_inicio}:{letra}{fila_fin}")
    try:
        rango.Copy()
        rango.PasteSpecial(Paste=XL_PASTE_VALUES)
    finally:
        # Always clear the marching ants, even when the paste failed.
        app.CutCopyMode = False


def _activar(ws):
    """PasteSpecial wants the sheet to be the active one (matters with two tabs)."""
    try:
        ws.Parent.Activate()
        ws.Activate()
    except Exception:
        pass


def _vinculos_externos(wb):
    """Sources of the external links the workbook still carries ([] if none)."""
    try:
        fuentes = wb.LinkSources(XL_EXCEL_LINKS)
    except Exception:
        return []
    return [str(f) for f in fuentes] if fuentes else []


def _procesar_libro(ruta, convertir_a_valores, update_status):
    """Everything that needs Excel, on the working copy at `ruta`; returns the summary (without 'ruta').
    Raises ValueError for anything the person can fix; with ABORTAR_SI_HAY_ERRORES the summary carries 'bloqueado'."""
    app = wb = pythoncom = None
    calculo_previo = None
    try:
        update_status("📂 Abriendo Excel...")
        app, pythoncom = _abrir_excel()

        wb = app.Workbooks.Open(ruta, UpdateLinks=0, ReadOnly=False)
        if bool(wb.ReadOnly):
            raise PermissionError(f"Excel abrió el libro en solo lectura: {ruta}")
        calculo_previo = app.Calculation
        # RESERVA ($) may be a formula (rents tab): validate what Excel shows, not a stale cache.
        app.CalculateFull()

        # --- 1. Read and validate EVERYTHING (nothing is written yet) --------
        update_status("🔎 Revisando las pestañas del layout...")
        secciones, avisos = localizar_secciones(wb)

        update_status(f"📚 Revisando la pestaña '{HOJA_CATALOGO}'...")
        catalogo = localizar_catalogo(wb)
        validar_catalogo(catalogo)

        problemas = []
        for s in secciones:
            if bool(s["hoja"].ProtectContents):
                raise ValueError(
                    f"La pestaña '{s['nombre']}' está protegida: no puedo escribir en "
                    f"ella.\n\nQuítale la protección y vuelve a intentar."
                )
            problemas += validar_entradas(s)
        if problemas:
            raise ValueError(_texto_problemas_entradas(problemas))

        for s in secciones:
            update_status(
                f"📊 '{s['nombre']}': encabezado en la fila {s['fila_encabezado']}, "
                f"datos {s['fila_inicio']}–{s['fila_fin']} ({s['filas']:,} renglones)"
            )

        # --- 2. Write (manual calculation: a recalculation per write is minutes, not seconds) ---
        app.Calculation = XL_CALC_MANUAL

        for s in secciones:
            ws = s["hoja"]
            cols = s["columnas"]
            ini, fin = s["fila_inicio"], s["fila_fin"]

            # Never above fila_inicio: rows 1-5 hold the title, banner and totals.
            if LIMPIAR_ANTES:
                fila_usada = max(int(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1), fin)
                for spec in COLUMNAS_POLIZA:
                    letra = cols[spec["clave"]]["letra"]
                    ws.Range(f"{letra}{ini}:{letra}{fila_usada}").ClearContents()

            letras = {clave: c["letra"] for clave, c in cols.items()}
            formulas = construir_formulas(ini, letras, catalogo)

            for spec in COLUMNAS_POLIZA:
                clave, letra = spec["clave"], cols[spec["clave"]]["letra"]
                update_status(f"✍️  {s['nombre']} — columna {letra} {spec['encabezado']}...")
                rango = ws.Range(f"{letra}{ini}:{letra}{fin}")
                if spec["tipo"] == "formula":
                    rango.Formula = formulas[clave]
                else:
                    valor = spec["valor"]
                    if clave == "ct" and CT_COMO_TEXTO:
                        rango.NumberFormat = "@"
                        valor = str(valor)
                    rango.Value = valor

        # --- 3. Calculate and look for failed lookups ---------------------
        update_status("🧮 Recalculando el libro...")
        app.Calculation = XL_CALC_AUTOMATIC
        app.CalculateFullRebuild()

        update_status("🔎 Revisando #N/A en los VLOOKUP...")
        errores, total_errores = [], 0
        for s in secciones:
            for spec in COLUMNAS_POLIZA:
                if spec["tipo"] != "formula":
                    continue
                letra = s["columnas"][spec["clave"]]["letra"]
                detalle, cuantos = _detectar_errores(
                    s["hoja"], letra, s["fila_inicio"], s["fila_fin"]
                )
                if not cuantos:
                    continue
                total_errores += cuantos
                errores.append({
                    "hoja":    s["nombre"],
                    "columna": letra,
                    "titulo":  spec["encabezado"],
                    "total":   cuantos,
                    "celdas":  detalle,
                })

        bloqueado = None
        if errores and ABORTAR_SI_HAY_ERRORES:
            resumen_err = ", ".join(
                f"{e['hoja']}!{e['columna']} ({e['total']})" for e in errores
            )
            bloqueado = (
                f"{total_errores} celda(s) con error en la póliza: {resumen_err}.\n\n"
                f"Se dejó el archivo con las fórmulas vivas para que puedas ver qué no "
                f"empató en '{HOJA_CATALOGO}'."
            )

        # --- 4. Control totals ---------------------------------------------
        letra_ct = lambda s: s["columnas"][CLAVE_MONTO]["letra"]   # noqa: E731
        hojas_resumen, total_control = [], 0.0
        for s in secciones:
            rango_monto = s["hoja"].Range(
                f"{letra_ct(s)}{s['fila_inicio']}:{letra_ct(s)}{s['fila_fin']}"
            )
            total_hoja = float(app.WorksheetFunction.Sum(rango_monto))
            total_control += total_hoja
            hojas_resumen.append({
                "hoja":            s["nombre"],
                "etiqueta":        s["etiqueta"],
                "fila_encabezado": s["fila_encabezado"],
                "fila_inicio":     s["fila_inicio"],
                "fila_fin":        s["fila_fin"],
                "filas":           s["filas"],
                "total_control":   round(total_hoja, 2),
                "columnas":        [
                    f"{s['columnas'][spec['clave']]['letra']} — {spec['encabezado']}"
                    for spec in COLUMNAS_POLIZA
                ],
            })

        # --- 5. Freeze the formulas into values --------------------------
        if convertir_a_valores and not bloqueado:
            update_status("🧊 Congelando fórmulas a valores...")
            for s in secciones:
                _activar(s["hoja"])
                for spec in COLUMNAS_POLIZA:
                    if spec["tipo"] != "formula":
                        continue
                    letra = s["columnas"][spec["clave"]]["letra"]
                    _congelar_a_valores(app, s["hoja"], letra, s["fila_inicio"], s["fila_fin"])

        # The example row's external link is overwritten but stays in the workbook and
        # makes Excel prompt for updates on every open.
        vinculos = _vinculos_externos(wb)
        if vinculos:
            avisos.append(
                "El libro todavía tiene vínculos externos (ya sin uso en la póliza): "
                + "; ".join(vinculos[:3])
                + ". Ábrelo y usa Datos → Editar vínculos → Interrumpir vínculo para que "
                  "Excel deje de preguntar por ellos."
            )

        update_status("💾 Guardando...")
        wb.Save()

        return {
            "hojas":                hojas_resumen,
            "hoja":                 " + ".join(s["nombre"] for s in secciones),
            "filas":                sum(s["filas"] for s in secciones),
            "columnas":             hojas_resumen[0]["columnas"],
            "total_control":        round(total_control, 2),
            "columna_control":      titulo_de(CLAVE_MONTO),
            "errores":              errores,
            "total_errores":        total_errores,
            "advertencias":         avisos,
            "convertido_a_valores": bool(convertir_a_valores and not bloqueado),
            "bloqueado":            bloqueado,
        }

    finally:
        try:
            if app is not None and calculo_previo is not None:
                app.Calculation = calculo_previo
        except Exception:
            pass
        _cerrar_excel(app, wb, pythoncom)


def _promover(ruta_tmp, ruta_final):
    """Rename the finished working copy to its final name, retrying while Excel lets go."""
    for intento in range(INTENTOS):
        try:
            os.replace(ruta_tmp, ruta_final)
            return
        except PermissionError:
            if intento == INTENTOS - 1:
                raise
            time.sleep(ESPERA_SEG * (intento + 1))


# --- Main entry point ---
def generar_poliza_sap(
    ruta_libro,
    ruta_salida=None,
    sobrescribir_original=False,
    convertir_a_valores=None,
    callback_status=None,
):
    """Fill 'PÓLIZA SAP' on both tabs of the OT layout; ruta_salida=None -> 'Output' beside the input.
    Returns: ruta, hojas, hoja, filas, columnas, total_control, columna_control, errores, total_errores,
    advertencias, convertido_a_valores. Raises FileNotFoundError/ValueError for fixable problems,
    PermissionError if the destination is open in Excel, RuntimeError without pywin32/Excel."""
    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    if convertir_a_valores is None:
        convertir_a_valores = CONVERTIR_A_VALORES

    if not ruta_libro or not os.path.isfile(ruta_libro):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta_libro}")
    if os.path.basename(ruta_libro).startswith("~$"):
        raise ValueError(
            f"'{os.path.basename(ruta_libro)}' es el archivo temporal que Excel crea "
            f"mientras el libro está abierto, no el layout. Elige el archivo sin '~$'."
        )

    # --- Where the result goes ----------------------------------------
    if sobrescribir_original:
        ruta_final = os.path.abspath(ruta_libro)
        ruta_trabajo = ruta_final
    else:
        if not ruta_salida:
            ruta_salida = carpeta_output_desde_input(
                os.path.dirname(os.path.abspath(ruta_libro))
            )
        os.makedirs(ruta_salida, exist_ok=True)
        base, ext = os.path.splitext(os.path.basename(ruta_libro))
        sello = datetime.today().strftime("%Y%m%d")
        ruta_final = os.path.abspath(
            os.path.join(ruta_salida, f"{base}_PolizaSAP_{sello}{ext}")
        )
        # Built under a temporary name and renamed at the end: a failed run never leaves a half file nor overwrites a good one.
        ruta_trabajo = os.path.join(
            os.path.dirname(ruta_final), f"_tmp{os.getpid()}_{os.path.basename(ruta_final)}"
        )
        update_status("📑 Copiando el libro (el original no se toca)...")
        # copy2 keeps the bytes, so Save() writes the same format back (no SaveAs / FileFormat).
        shutil.copy2(ruta_libro, ruta_trabajo)

    try:
        resumen = _procesar_libro(ruta_trabajo, convertir_a_valores, update_status)
        if ruta_trabajo != ruta_final:
            update_status("📦 Dejando el archivo final...")
            _promover(ruta_trabajo, ruta_final)
    finally:
        # Whatever happened, no temporary file stays behind.
        if ruta_trabajo != ruta_final and os.path.exists(ruta_trabajo):
            try:
                os.remove(ruta_trabajo)
            except OSError:
                pass

    resumen["ruta"] = ruta_final
    if resumen["bloqueado"]:
        raise ValueError(f"{resumen['bloqueado']}\n\n{ruta_final}")
    return resumen


# --- Standalone run: python Create_Poliza_SAP.py ---
if __name__ == "__main__":
    print("=== Prueba de 'PÓLIZA SAP' (standalone) ===")
    ruta = input("Ruta del Layout OT (.xlsx): ").strip().strip('"')

    try:
        resumen = generar_poliza_sap(ruta, callback_status=print)
        print("\n🎉 LISTO 🎉")
        for clave, valor in resumen.items():
            if clave == "errores":
                print(f"  {clave:<22}: {len(valor)} columna(s) con error")
                for e in valor:
                    celdas = ", ".join(c["celda"] for c in e["celdas"])
                    print(f"      {e['hoja']} {e['columna']} ({e['total']}): {celdas}")
            elif clave == "hojas":
                for h in valor:
                    print(f"  {h['hoja']:<22}: filas {h['fila_inicio']}–{h['fila_fin']}, "
                          f"Σ {h['total_control']:,.2f}")
            elif clave in ("advertencias", "columnas"):
                print(f"  {clave:<22}:")
                for v in valor:
                    print(f"      • {v}")
            else:
                print(f"  {clave:<22}: {valor}")
    except Exception as e:
        print(f"❌ Error generando la póliza: {e}")
        raise
