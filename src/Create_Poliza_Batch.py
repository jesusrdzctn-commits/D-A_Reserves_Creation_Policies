"""
RDA Pólizas D&A — Create_Poliza_Batch.py
========================================

Third consolidation of the project: copies the finished 'PÓLIZA SAP' section of
'Layout Creación Reservas' into the template 'Póliza Carga Batch Creaciones'
(sheet 'GL Posting Upload'), which is the file the stakeholder uploads to SAP.

    Layout Creación Reservas (.xlsb, sheet 'Creaciones')  --values-->  Póliza Carga Batch (.xlsx)

It runs AFTER Create_Poliza_SAP.generar_poliza_sap(): that function fills and
freezes columns R..Z of the layout; this one only READS them. The layout is
opened read-only and is never saved.


WHAT GOES WHERE  (yellow cells of the template)
-----------------------------------------------
    Layout 'Creaciones'            Template 'GL Posting Upload'     Block
    ---------------------------    ----------------------------     -------------
    R  CUENTA SAP                  GL Account                       debit only
    T  CeBe                        Profit Center                    debit + credit
    U  SEGMENTO                    Segment                          debit + credit
    V  Asignación                  Assignment                       debit + credit
    X  MONTO                       Debit Amount                     debit block
    X  MONTO                       Credit Amount                    credit block
    Z  CONCEPTO (descripción SAP)  Line item Text                   debit + credit

The data is written TWICE, one block under the other, with no gap:

    row 1                header
    rows 2 .. n+1        debit block   (GL Account, Profit Center, Segment,
                                        Assignment, Debit Amount, Line item Text)
    rows n+2 .. 2n+1    credit block  (Profit Center, Segment, Assignment,
                                        Credit Amount, Line item Text)
                                        -> GL Account is intentionally left EMPTY
                                           in the credit block.

Everything is pasted with Copy + PasteSpecial(Values), i.e. Excel's own
'Pegado especial -> Valores'. NOT `dest.Value = src.Value`: that would make Excel
re-parse '0004209900' and drop the leading zeros (see Create_Poliza_SAP.py).

The orange cells of the template (Document Sequence #, Company Code, dates...)
are NOT filled here yet. They are left untouched until their rules are defined.


SAFETY NET
----------
The run refuses to produce the file when the layout still has error cells
(#N/A from a failed VLOOKUP), empty cells or a non-numeric amount: those rows
would be rejected by SAP. After pasting, every pasted column is read back and
compared with the layout; the file is only saved if they are identical, and it
is written under a temporary name and renamed at the very end, so a failed run
can never leave a half-built (or empty) batch file next to the good ones.
"""

import math
import os
import re
import shutil
import time
from datetime import datetime

# The layout helpers live in Create_Poliza_SAP.py. They are reused instead of
# copied so both modules always agree on where the section starts and ends.
from Create_Poliza_SAP import (
    COLUMNAS_POLIZA,
    HOJA_CREACIONES,
    XL_PASTE_VALUES,
    _localizar_fila_encabezado,
    _ultima_fila,
    _validar_encabezados,
)
from utils import abrir_excel as _abrir_excel
from utils import cerrar_excel as _cerrar_excel
from utils import hoja_por_nombre as _hoja
from utils import norm_texto as _norm

# ----------------------------------------------------------------------
# Template layout
# ----------------------------------------------------------------------
HOJA_PLANTILLA = "GL Posting Upload"
EXTENSIONES_PLANTILLA = (".xlsx", ".xlsm")

# The header row is searched in the first rows instead of hard-coding row 1.
FILA_ENCABEZADO_MAX = 10

# Headers of the template that this module writes under. Matched on the whole
# normalised text, so 'Profit Center' never grabs 'partner_profit_center'.
DEST_CUENTA, DEST_DEBITO, DEST_CREDITO = "GL Account", "Debit Amount", "Credit Amount"
DEST_CEBE, DEST_SEGMENTO = "Profit Center", "Segment"
DEST_ASIGNACION, DEST_TEXTO = "Assignment", "Line item Text"

# Title (in COLUMNAS_POLIZA) of the layout column that feeds each header.
# The letters (R, T, U...) are NOT repeated here: they are looked up in
# COLUMNAS_POLIZA, so if the layout columns ever move there is one place to fix.
TITULO_MONTO = "MONTO"
TITULO_CUENTA = "CUENTA SAP"
TRANSFERENCIAS = (
    {"origen": TITULO_CUENTA,                "destino": DEST_CUENTA,
     "en_debito": True,  "en_credito": False},
    {"origen": "CeBe",                       "destino": DEST_CEBE,
     "en_debito": True,  "en_credito": True},
    {"origen": "SEGMENTO",                   "destino": DEST_SEGMENTO,
     "en_debito": True,  "en_credito": True},
    {"origen": "Asignación",                 "destino": DEST_ASIGNACION,
     "en_debito": True,  "en_credito": True},
    {"origen": TITULO_MONTO,                 "destino": DEST_DEBITO,
     "en_debito": True,  "en_credito": False},
    {"origen": TITULO_MONTO,                 "destino": DEST_CREDITO,
     "en_debito": False, "en_credito": True},
    {"origen": "CONCEPTO (descripción SAP)", "destino": DEST_TEXTO,
     "en_debito": True,  "en_credito": True},
)

# Standard SAP field lengths. Used ONLY to warn (SAP truncates or rejects
# longer values on upload); the run is never blocked by them.
LONGITUD_MAX_SAP = {
    DEST_CUENTA: 10,
    DEST_CEBE: 10,
    DEST_SEGMENTO: 10,
    DEST_ASIGNACION: 18,
    DEST_TEXTO: 50,
}

# ----------------------------------------------------------------------
# Behaviour flags
# ----------------------------------------------------------------------
# Clear the target columns of the template (below the header) before pasting,
# so leftovers of an old run in the template cannot survive under the new data.
LIMPIAR_ANTES = True

# How many offending cells / rows to list per problem.
MAX_DETALLE = 10

# Excel's clipboard and the file system both fail transiently now and then
# (another app grabbing the clipboard, antivirus scanning the new file).
INTENTOS = 4
ESPERA_SEG = 0.6

# Excel error values as pywin32 returns them (CVErr codes as signed ints).
CODIGOS_ERROR_EXCEL = {
    -2146826288: "#NULL!",
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
    -2146826245: "#GETTING_DATA",
}

# Tolerance when comparing the debit / credit / layout totals (cents).
TOLERANCIA_TOTAL = 0.005


# ======================================================================
# Small helpers
# ======================================================================
def _misma_ruta(a, b):
    """True when two paths point to the same file (Windows: case-insensitive)."""
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def _nombre_salida(ruta_plantilla):
    """
    Name of the generated file, derived from the template's name.

        '07 Póliza Carga Batch Creaciones_OT 26 template_IA_v2.xlsx'
            -> '07 Póliza Carga Batch Creaciones_OT 26_20261027.xlsx'

    Everything from the word 'template' on is dropped, and the date stamp is
    ALWAYS appended, so the output can never be named like the template itself.
    """
    base, ext = os.path.splitext(os.path.basename(ruta_plantilla))
    limpio = re.sub(r"[\s_\-]*template.*$", "", base, flags=re.IGNORECASE).strip(" _-")
    sello = datetime.today().strftime("%Y%m%d")
    return f"{limpio or base}_{sello}{ext}"


def _letra_origen(titulo):
    """Layout column letter of a COLUMNAS_POLIZA entry, looked up by its title."""
    for spec in COLUMNAS_POLIZA:
        if spec["titulo"] == titulo:
            return spec["col"]
    raise ValueError(
        f"La columna '{titulo}' ya no existe en COLUMNAS_POLIZA "
        f"(Create_Poliza_SAP.py). ¿Se renombró o se quitó?"
    )


def _titulos_origen():
    """Layout columns that feed the batch, in order, without repeats."""
    vistos = []
    for t in TRANSFERENCIAS:
        if t["origen"] not in vistos:
            vistos.append(t["origen"])
    return vistos


def _leer_rango(rango, n):
    """
    Value2 of a ONE-column range as a flat list of n items.

    A single-cell range returns a scalar (not a tuple of tuples), which is the
    classic COM trap for a policy with one line.
    Value2, not Value: it hands back raw numbers and text, never dates/currency.
    """
    valor = rango.Value2
    if n == 1:
        return [valor]
    return [fila[0] for fila in valor]


def _vacio(valor):
    return valor is None or (isinstance(valor, str) and not valor.strip())


def _error_de(valor):
    """
    Excel error text ('#N/A', '#REF!'...) of a cell value, or None if it is not an error.

    pywin32 hands errors over as CVErr ints; the literal text is accepted too so
    a different pywin32/Excel build cannot sneak an error past the check.
    """
    if isinstance(valor, bool):          # bool is an int subclass: exclude it first
        return None
    if isinstance(valor, int):
        return CODIGOS_ERROR_EXCEL.get(valor)
    if isinstance(valor, str) and valor.strip() in CODIGOS_ERROR_EXCEL.values():
        return valor.strip()
    return None


def _abrir_libro(app, ruta, solo_lectura):
    """Workbooks.Open with a few retries (file still locked by a dying Excel)."""
    ultimo = None
    for intento in range(INTENTOS):
        try:
            return app.Workbooks.Open(ruta, UpdateLinks=0, ReadOnly=solo_lectura)
        except Exception as e:  # com_error
            ultimo = e
            time.sleep(ESPERA_SEG * (intento + 1))
    raise RuntimeError(
        f"Excel no pudo abrir '{os.path.basename(ruta)}':\n\n{ultimo}\n\n"
        f"¿Está abierto en otro programa?"
    )


def _pegar_valores(app, origen, destino):
    """
    Copy -> Paste Special -> Values, retrying when the clipboard is busy.

    The marching ants are always cleared, so the user's clipboard is not left
    holding a chunk of the workbook.
    """
    ultimo = None
    for intento in range(INTENTOS):
        try:
            origen.Copy()
            destino.PasteSpecial(Paste=XL_PASTE_VALUES)
            return
        except Exception as e:  # com_error
            ultimo = e
            time.sleep(ESPERA_SEG * (intento + 1))
        finally:
            try:
                app.CutCopyMode = False
            except Exception:
                pass
    raise RuntimeError(
        f"No pude pegar los valores en la plantilla (¿otro programa está usando "
        f"el portapapeles?):\n\n{ultimo}"
    )


# ======================================================================
# Layout (source)
# ======================================================================
def _localizar_origen(ws):
    """
    Header row, first and last data row of the layout's 'PÓLIZA SAP' section.

    Same detection as the module that fills it, on purpose.
    """
    if bool(ws.FilterMode):
        # Copy on a filtered range only takes the VISIBLE rows: the batch would
        # silently come out short. Better to stop and ask for the filter off.
        raise ValueError(
            f"La pestaña '{HOJA_CREACIONES}' del Layout tiene un filtro activo "
            f"(hay filas ocultas por el filtro).\n\n"
            f"Quita el filtro y vuelve a intentar."
        )

    fila_encabezado = _localizar_fila_encabezado(ws)
    _validar_encabezados(ws, fila_encabezado)
    fila_inicio = fila_encabezado + 1
    fila_fin = _ultima_fila(ws, fila_inicio)
    if fila_fin is None:
        raise ValueError(
            f"No encontré datos en la pestaña '{HOJA_CREACIONES}' del Layout "
            f"(a partir de la fila {fila_inicio}).\n\n¿Es el archivo correcto?"
        )
    return fila_inicio, fila_fin


def _leer_origen(ws, fila_inicio, fila_fin):
    """{title: (column letter, [values])} for every layout column that is copied."""
    n = fila_fin - fila_inicio + 1
    datos = {}
    for titulo in _titulos_origen():
        letra = _letra_origen(titulo)
        rango = ws.Range(f"{letra}{fila_inicio}:{letra}{fila_fin}")
        datos[titulo] = (letra, _leer_rango(rango, n))
    return datos


def _validar_origen(datos, fila_inicio):
    """
    Check what is about to go up to SAP.

    Returns (problemas, advertencias):
      problemas    -> BLOCK the run: error cells, empty cells, non-numeric amount.
      advertencias -> reported in the summary, do not block: zero / negative
                      amounts, numeric GL accounts, values longer than SAP allows.
    """
    problemas = []
    advertencias = []

    for titulo, (letra, valores) in datos.items():
        es_monto = titulo == TITULO_MONTO
        grupos = {"error": [], "vacío": [], "no numérico": [], "tipo inesperado": []}

        for i, valor in enumerate(valores):
            celda = f"{letra}{fila_inicio + i}"
            if _error_de(valor):
                grupos["error"].append(f"{celda} {_error_de(valor)}")
            elif _vacio(valor):
                grupos["vacío"].append(celda)
            elif isinstance(valor, bool):
                grupos["tipo inesperado"].append(celda)
            elif es_monto and not isinstance(valor, (int, float)):
                grupos["no numérico"].append(celda)
            elif not isinstance(valor, (str, int, float)):
                grupos["tipo inesperado"].append(celda)

        for tipo, celdas in grupos.items():
            if celdas:
                problemas.append({
                    "columna": letra, "titulo": titulo, "tipo": tipo,
                    "total": len(celdas), "celdas": celdas[:MAX_DETALLE],
                })

    # --- Warnings (only meaningful when the cells themselves are usable) ---
    if not problemas:
        letra_m, montos = datos[TITULO_MONTO]
        ceros = [f"{letra_m}{fila_inicio + i}" for i, v in enumerate(montos) if v == 0]
        negativos = [f"{letra_m}{fila_inicio + i}" for i, v in enumerate(montos) if v < 0]
        if ceros:
            advertencias.append(
                f"{len(ceros)} renglón(es) con MONTO = 0 (se pasan tal cual): "
                + ", ".join(ceros[:MAX_DETALLE]) + ("..." if len(ceros) > MAX_DETALLE else "")
            )
        if negativos:
            advertencias.append(
                f"{len(negativos)} renglón(es) con MONTO negativo (irían con signo "
                f"negativo en Debit y Credit Amount): "
                + ", ".join(negativos[:MAX_DETALLE]) + ("..." if len(negativos) > MAX_DETALLE else "")
            )

        letra_c, cuentas = datos[TITULO_CUENTA]
        numericas = [f"{letra_c}{fila_inicio + i}" for i, v in enumerate(cuentas)
                     if not isinstance(v, str)]
        if numericas:
            advertencias.append(
                f"{len(numericas)} CUENTA SAP vienen como NÚMERO y no como texto "
                f"(¿se perdieron ceros a la izquierda?): "
                + ", ".join(numericas[:MAX_DETALLE]) + ("..." if len(numericas) > MAX_DETALLE else "")
            )

        for t in TRANSFERENCIAS:
            limite = LONGITUD_MAX_SAP.get(t["destino"])
            if not limite:
                continue
            letra_o, valores = datos[t["origen"]]
            largos = [f"{letra_o}{fila_inicio + i}" for i, v in enumerate(valores)
                      if isinstance(v, str) and len(v) > limite]
            if largos:
                advertencias.append(
                    f"{len(largos)} valor(es) de '{t['destino']}' exceden los {limite} "
                    f"caracteres que SAP acepta: "
                    + ", ".join(largos[:MAX_DETALLE]) + ("..." if len(largos) > MAX_DETALLE else "")
                )

    return problemas, advertencias


def _texto_problemas(problemas, ruta_layout):
    """Readable message for the blocking problems found in the layout."""
    total = sum(p["total"] for p in problemas)
    lineas = []
    for p in problemas:
        resto = p["total"] - len(p["celdas"])
        lineas.append(
            f"     • {p['columna']} ({p['titulo']}): {p['total']} celda(s) con "
            f"{p['tipo']}\n       " + ", ".join(p["celdas"])
            + (f", ...y {resto} más" if resto > 0 else "")
        )
    return (
        f"No se generó la Póliza Carga Batch: el Layout todavía tiene {total} "
        f"celda(s) con problemas que SAP rechazaría.\n\n"
        + "\n".join(lineas)
        + "\n\nCorrígelas (normalmente es un valor que falta en la pestaña "
          "'catalogo') y vuelve a correr 'Generar PÓLIZA SAP'.\n\n"
          f"Layout revisado:\n{ruta_layout}"
    )


# ======================================================================
# Template (destination)
# ======================================================================
def _hoja_plantilla(wb):
    """The template's sheet by name; a one-sheet workbook is accepted whatever it is called."""
    try:
        return _hoja(wb, HOJA_PLANTILLA)
    except ValueError:
        if int(wb.Worksheets.Count) == 1:
            return wb.Worksheets(1)
        raise


def _localizar_destino(ws):
    """
    Find the header row of the template and the column of every target header.

    Returns (fila_encabezado, {header: column index}). Looking the columns up by
    NAME (not by letter) means a template where somebody inserted a column still
    lands every value under the right header — or stops with a readable message.
    """
    requeridos = {_norm(h): h for h in (
        DEST_CUENTA, DEST_DEBITO, DEST_CREDITO, DEST_CEBE,
        DEST_SEGMENTO, DEST_ASIGNACION, DEST_TEXTO,
    )}

    usado = ws.UsedRange
    ultima_col = max(int(usado.Column + usado.Columns.Count - 1), 1)
    bloque = ws.Range(ws.Cells(1, 1), ws.Cells(FILA_ENCABEZADO_MAX, ultima_col)).Value2

    mejor_faltantes = None
    for fila_idx, fila in enumerate(bloque, start=1):
        encontrados = {}
        for col_idx, celda in enumerate(fila, start=1):
            clave = _norm(celda)
            if clave in requeridos:
                encontrados.setdefault(clave, []).append(col_idx)

        faltantes = [requeridos[k] for k in requeridos if k not in encontrados]
        if not faltantes:
            repetidos = [requeridos[k] for k, cols in encontrados.items() if len(cols) > 1]
            if repetidos:
                raise ValueError(
                    f"La plantilla tiene encabezados repetidos en la fila {fila_idx}: "
                    f"{repetidos}. No sé en cuál pegar."
                )
            return fila_idx, {requeridos[k]: cols[0] for k, cols in encontrados.items()}

        if mejor_faltantes is None or len(faltantes) < len(mejor_faltantes):
            mejor_faltantes = faltantes

    raise ValueError(
        f"La plantilla no tiene los encabezados que necesito en las primeras "
        f"{FILA_ENCABEZADO_MAX} filas. Me faltan: {mejor_faltantes}.\n\n"
        f"¿Es la plantilla correcta de 'Póliza Carga Batch'?"
    )


def _rango_destino(ws, col_idx, fila_ini, fila_fin):
    return ws.Range(ws.Cells(fila_ini, col_idx), ws.Cells(fila_fin, col_idx))


def _limpiar_destino(ws, fila_encabezado, columnas):
    """ClearContents (never formats) of the target columns, below the header only."""
    usado = ws.UsedRange
    fila_usada = int(usado.Row + usado.Rows.Count - 1)
    if fila_usada <= fila_encabezado:
        return
    for col_idx in columnas.values():
        _rango_destino(ws, col_idx, fila_encabezado + 1, fila_usada).ClearContents()


# ======================================================================
# Paste + verification
# ======================================================================
def _bloques(fila_encabezado, n):
    """Row spans of the two blocks, the credit one starting right under the debit one."""
    d_ini = fila_encabezado + 1
    d_fin = d_ini + n - 1
    c_ini = d_fin + 1
    c_fin = c_ini + n - 1
    return {"debito": (d_ini, d_fin), "credito": (c_ini, c_fin)}


def _pegar_todo(app, ws_origen, ws_destino, columnas, bloques, fila_inicio, fila_fin):
    """Paste every (column, block) of TRANSFERENCIAS as values."""
    # PasteSpecial wants the destination sheet to be the active one.
    try:
        ws_destino.Parent.Activate()
        ws_destino.Activate()
    except Exception:
        pass

    for t in TRANSFERENCIAS:
        letra = _letra_origen(t["origen"])
        origen = ws_origen.Range(f"{letra}{fila_inicio}:{letra}{fila_fin}")
        for bloque, bandera in (("debito", "en_debito"), ("credito", "en_credito")):
            if not t[bandera]:
                continue
            r1, r2 = bloques[bloque]
            _pegar_valores(app, origen, _rango_destino(ws_destino, columnas[t["destino"]], r1, r2))


def _verificar(ws_destino, columnas, bloques, datos, n):
    """
    Read the template back and compare it with the layout.

    Returns (diferencias, sumas, total_origen):
      diferencias  readable differences, empty list = identical
      sumas        {'Debit Amount': x, 'Credit Amount': y} as read from the template
      total_origen sum of MONTO in the layout

    Three things are checked: every pasted column equals its source cell by cell
    (this is what catches a lost leading zero), every cell that must stay empty
    IS empty, and debit total = credit total = layout total.
    """
    diferencias = []
    sumas = {}

    for t in TRANSFERENCIAS:
        _, esperado = datos[t["origen"]]
        col_idx = columnas[t["destino"]]
        for bloque, bandera in (("debito", "en_debito"), ("credito", "en_credito")):
            r1, r2 = bloques[bloque]
            real = _leer_rango(_rango_destino(ws_destino, col_idx, r1, r2), n)

            if t[bandera]:
                malos = [(r1 + i, e, r) for i, (e, r) in enumerate(zip(esperado, real)) if e != r]
                for fila, e, r in malos[:MAX_DETALLE]:
                    diferencias.append(
                        f"{t['destino']} fila {fila} ({bloque}): esperaba {e!r} y quedó {r!r}"
                    )
                if len(malos) > MAX_DETALLE:
                    diferencias.append(f"{t['destino']} ({bloque}): ...y {len(malos) - MAX_DETALLE} más")
                if t["destino"] in (DEST_DEBITO, DEST_CREDITO):
                    sumas[t["destino"]] = math.fsum(v for v in real if isinstance(v, (int, float)))
            else:
                sucios = [r1 + i for i, v in enumerate(real) if not _vacio(v)]
                if sucios:
                    diferencias.append(
                        f"{t['destino']} ({bloque}) debía quedar vacío y trae datos en las "
                        f"filas {sucios[:MAX_DETALLE]}"
                    )

    total_origen = math.fsum(v for v in datos[TITULO_MONTO][1] if isinstance(v, (int, float)))
    for destino, suma in sumas.items():
        if abs(suma - total_origen) > TOLERANCIA_TOTAL:
            diferencias.append(
                f"Σ {destino} = {suma:,.2f} y la Σ MONTO del Layout = {total_origen:,.2f}"
            )
    return diferencias, sumas, total_origen


# ======================================================================
# Main entry point
# ======================================================================
def generar_poliza_batch(
    ruta_layout,
    ruta_plantilla,
    ruta_salida=None,
    callback_status=None,
):
    """
    Fill 'Póliza Carga Batch Creaciones' from the finished layout.

    Args:
        ruta_layout (str): the layout ALREADY processed by generar_poliza_sap()
            (its 'PÓLIZA SAP' columns filled). Opened read-only, never saved.
        ruta_plantilla (str): the empty 'Póliza Carga Batch ... template' (.xlsx).
            Never modified: the work happens on a copy.
        ruta_salida (str|None): folder for the result. None -> the folder of
            ruta_layout (which is already the 'Output' folder when it is the copy
            produced by generar_poliza_sap()).
        callback_status (function|None): progress reporting for the GUI.

    Returns:
        dict: ruta, origen, hoja, filas_origen, filas_totales, bloque_debito,
              bloque_credito, total_debito, total_credito, columnas,
              advertencias.

    Raises:
        FileNotFoundError / ValueError: expected, actionable problems (missing
            file, filtered layout, error cells in the layout, wrong template).
        RuntimeError: Excel / clipboard failures, or the read-back check failing.
        PermissionError: the destination file is open in Excel.
    """
    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    # --- Inputs ----------------------------------------------------------
    for ruta, etiqueta in ((ruta_layout, "el Layout Creación Reservas"),
                           (ruta_plantilla, "la plantilla de Póliza Carga Batch")):
        if not ruta or not os.path.isfile(ruta):
            raise FileNotFoundError(f"No encontré {etiqueta}: {ruta}")

    if os.path.splitext(ruta_plantilla)[1].lower() not in EXTENSIONES_PLANTILLA:
        raise ValueError(
            f"La plantilla debe ser {' o '.join(EXTENSIONES_PLANTILLA)} y recibí "
            f"'{os.path.splitext(ruta_plantilla)[1]}'."
        )

    carpeta = os.path.abspath(ruta_salida) if ruta_salida \
        else os.path.dirname(os.path.abspath(ruta_layout))
    os.makedirs(carpeta, exist_ok=True)

    ruta_final = os.path.join(carpeta, _nombre_salida(ruta_plantilla))
    if _misma_ruta(ruta_final, ruta_plantilla) or _misma_ruta(ruta_final, ruta_layout):
        # Never overwrite one of the inputs.
        raise ValueError(f"El archivo de salida coincidiría con un archivo de entrada:\n{ruta_final}")

    # Built under a temporary name; renamed to the final one only after success.
    ruta_tmp = os.path.join(carpeta, f"_tmp{os.getpid()}_{os.path.basename(ruta_final)}")

    update_status("📑 Copiando la plantilla (la original no se toca)...")
    shutil.copy2(ruta_plantilla, ruta_tmp)

    resumen = None
    try:
        app = wb_origen = wb_destino = pythoncom = None
        try:
            update_status("📂 Abriendo Excel...")
            app, pythoncom = _abrir_excel()

            wb_origen = _abrir_libro(app, os.path.abspath(ruta_layout), solo_lectura=True)
            wb_destino = _abrir_libro(app, ruta_tmp, solo_lectura=False)

            ws_origen = _hoja(wb_origen, HOJA_CREACIONES)
            ws_destino = _hoja_plantilla(wb_destino)

            # --- Read and validate the layout ----------------------------
            fila_inicio, fila_fin = _localizar_origen(ws_origen)
            n = fila_fin - fila_inicio + 1
            update_status(f"🔎 Revisando {n:,} renglones del Layout (filas {fila_inicio}–{fila_fin})...")

            datos = _leer_origen(ws_origen, fila_inicio, fila_fin)
            problemas, advertencias = _validar_origen(datos, fila_inicio)
            if problemas:
                raise ValueError(_texto_problemas(problemas, ruta_layout))

            # --- Locate the template columns -----------------------------
            fila_encabezado, columnas = _localizar_destino(ws_destino)
            bloques = _bloques(fila_encabezado, n)
            if bloques["credito"][1] > int(ws_destino.Rows.Count):
                raise ValueError(
                    f"{n:,} renglones x 2 bloques no caben en la hoja de la plantilla."
                )

            if LIMPIAR_ANTES:
                _limpiar_destino(ws_destino, fila_encabezado, columnas)

            # --- Paste ---------------------------------------------------
            update_status("✍️  Pegando valores en la plantilla (Débito)...")
            _pegar_todo(app, ws_origen, ws_destino, columnas, bloques, fila_inicio, fila_fin)

            # --- Read back and compare ------------------------------------
            update_status("🧮 Verificando lo pegado contra el Layout...")
            diferencias, sumas, total_origen = _verificar(
                ws_destino, columnas, bloques, datos, n
            )
            if diferencias:
                raise RuntimeError(
                    "La verificación posterior al pegado encontró diferencias, así que "
                    "NO se generó la Póliza Carga Batch:\n\n"
                    + "\n".join(f"     • {d}" for d in diferencias)
                )

            update_status("💾 Guardando...")
            wb_destino.Save()

            resumen = {
                "ruta":           ruta_final,
                "origen":         os.path.abspath(ruta_layout),
                "hoja":           str(ws_destino.Name),
                "filas_origen":   n,
                "filas_totales":  2 * n,
                "bloque_debito":  bloques["debito"],
                "bloque_credito": bloques["credito"],
                "total_debito":   round(sumas.get(DEST_DEBITO, 0.0), 2),
                "total_credito":  round(sumas.get(DEST_CREDITO, 0.0), 2),
                "total_layout":   round(total_origen, 2),
                "columnas":       [f"{t['origen']} → {t['destino']}" for t in TRANSFERENCIAS],
                "advertencias":   advertencias,
            }

        finally:
            # Layout first (read-only, nothing to save), then the template, then Excel.
            try:
                if wb_origen is not None:
                    wb_origen.Close(SaveChanges=False)
            except Exception:
                pass
            _cerrar_excel(app, wb_destino, pythoncom)

        # Excel is gone, so the temp file is no longer locked. Retry the rename:
        # the EXCEL.EXE process can take a moment to release it.
        update_status("📦 Dejando el archivo final...")
        for intento in range(INTENTOS):
            try:
                os.replace(ruta_tmp, ruta_final)
                break
            except PermissionError:
                if intento == INTENTOS - 1:
                    raise
                time.sleep(ESPERA_SEG * (intento + 1))

    finally:
        # Whatever happened, no temp file stays behind.
        if os.path.exists(ruta_tmp):
            try:
                os.remove(ruta_tmp)
            except OSError:
                pass

    return resumen


# ----------------------------------------------------------------------
# Standalone run (no GUI, no controller):  python Create_Poliza_Batch.py
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=== Prueba de 'Póliza Carga Batch' (standalone) ===")
    layout = input("Ruta del Layout YA procesado (.xlsb): ").strip().strip('"')
    plantilla = input("Ruta de la plantilla de Póliza Carga Batch (.xlsx): ").strip().strip('"')

    try:
        res = generar_poliza_batch(layout, plantilla, callback_status=print)
        print("\n🎉 LISTO 🎉")
        for clave, valor in res.items():
            if clave in ("columnas", "advertencias"):
                print(f"  {clave:<16}:")
                for v in valor:
                    print(f"      • {v}")
            else:
                print(f"  {clave:<16}: {valor}")
    except Exception as e:
        print(f"❌ Error generando la póliza batch: {e}")
        raise
