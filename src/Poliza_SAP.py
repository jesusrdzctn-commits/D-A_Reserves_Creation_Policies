"""
RDA Pólizas D&A — Poliza_SAP.py
================================

Fills the 'PÓLIZA SAP' section of 'Layout Creación Reservas DA WHSL_<mes>.xlsb'.

This is the SECOND consolidation of the project. It is deliberately a separate
module from Consolidacion.py because it is a different animal:

    Consolidacion.py  ->  '...Creacion de Reservas PMF OT.xlsx'   openpyxl + pandas
    Poliza_SAP.py     ->  'Layout Creación Reservas DA WHSL.xlsb'  Excel COM (win32com)

Why Excel COM and not openpyxl: the file is .xlsb (Excel Binary Workbook).
openpyxl cannot even OPEN a .xlsb — it is a completely different container
(binary records instead of XML). The options were to convert the file to .xlsx
(which changes the format the team actually uses) or to drive the Excel that is
already installed on the machine. We drive Excel.

Side benefit: Excel COM is a perfect round-trip. It does not lose the Purview
sensitivity label, the printer settings or the custom properties that the
openpyxl round-trip drops (see claude/hallazgos-variaciones.md).

Cost: this module only runs on Windows with Excel installed. That is fine — it
is the same machine that already runs the SAP GUI automations.


LAYOUT OF THE 'Creaciones' SHEET (verified against the real workbook)
---------------------------------------------------------------------
    rows 1-2    legend off to the side (P:U): cargo = CT 40, abono = CT 50
    row   4     'PÓLIZA SAP' banner over R, and 'Formulated' markers
    row   5     SAP field names: Cuenta | Centro de Costo | Centro de Beneficio |
                Segmento | Asignación | Clave Contab. | Importe | Importe Conv. | Texto
    row   6     REAL header row
    row   7+    data

    Input block    A MES | B BU | C AGRUPADOR ABIERTO | D MOTIVO | E VALIDACION |
                   F AGRUPADOR CERRADO | G CANAL | H ID GPO CADENA |
                   I DESCR GPO CADENA | J CLIENTE | K DESCRIPCIÓN DE LA PROVISIÓN |
                   L Venta | M D&A | N Reserva | O VLOOK UP
    PÓLIZA SAP     R CUENTA SAP | S Centro de Costo (vacía) | T CeBe | U SEGMENTO |
                   V Asignación | W CT | X MONTO | Y Importe Conv. (vacía) |
                   Z CONCEPTO (descripción SAP)

The seven columns this module writes, exactly as the analyst types them today:

    R  CUENTA SAP   =VLOOKUP(C7,catalogo!$E:$H,4,0)      -> catalogo.H 'CUENTA POLIZA'
    T  CeBe         =VLOOKUP(B7,catalogo!$B:$D,3,0)      -> catalogo.D 'Cebe'
    U  SEGMENTO     MXC0000                              (constant)
    V  Asignación   =UPPER(LEFT(B7,3)&"\\"&LEFT(G7,2)&"\\"&H7)
    W  CT           40                                   (constant)
    X  MONTO        =IFERROR(ROUND(N7,2),0)
    Z  CONCEPTO     =VLOOKUP(C7,catalogo!$E:$I,5,0)&" "&VLOOKUP(Creaciones!G7,catalogo!$J:$K,2,0)
                                                         -> catalogo.I + catalogo.K

Flow: copy the workbook -> locate the header row -> write the formulas -> let
Excel calculate -> report every #N/A -> freeze the results as values -> save.
"""

import os
import shutil
from datetime import datetime

# The COM plumbing lives in utils.py because 'Comparativos' needs the very same
# four helpers. Aliased to the private names this module already used, so every
# call site below stays exactly as it was.
from utils import abrir_excel as _abrir_excel
from utils import carpeta_output_desde_input
from utils import cerrar_excel as _cerrar_excel
from utils import hoja_por_nombre as _hoja
from utils import norm_texto as _norm

# ----------------------------------------------------------------------
# Excel COM constants
# ----------------------------------------------------------------------
# Hard-coded instead of using win32com.client.constants on purpose: those only
# exist after gencache.EnsureDispatch(), which is exactly the call that breaks
# inside a PyInstaller .exe. Plain numbers always work.
XL_UP                = -4162   # xlUp
XL_CELLTYPE_FORMULAS = -4123   # xlCellTypeFormulas
XL_ERRORS            = 16      # xlErrors
XL_CALC_MANUAL       = -4135   # xlCalculationManual
XL_CALC_AUTOMATIC    = -4105   # xlCalculationAutomatic
XL_PASTE_VALUES      = -4163   # xlPasteValues


# ----------------------------------------------------------------------
# Sheets and layout
# ----------------------------------------------------------------------
HOJA_CREACIONES = "Creaciones"
# CAREFUL: the workbook also carries a sheet named 'Catálogo' (capital C, with
# an accent) which is a DIFFERENT table. The formulas point at 'catalogo', so
# _hoja() matches the exact name first before falling back to a loose match.
HOJA_CATALOGO   = "catalogo"

# Header row of the 'PÓLIZA SAP' section. Detected at run time by looking for
# these two headers; this is only the fallback if detection fails.
FILA_ENCABEZADO_DEFAULT = 6

# Headers used to FIND the header row. Both must be present on the same row.
ANCLA_ENCABEZADO = {"C": "AGRUPADOR ABIERTO", "R": "CUENTA SAP"}

# Sanity check once the row is found: these must be where we think they are.
# Matched as "the real header STARTS WITH this", so 'CONCEPTO (descripción SAP)'
# matches 'CONCEPTO' and a renamed column still fails loudly.
ENCABEZADOS_ESPERADOS = {
    "R": "CUENTA SAP",
    "T": "CEBE",
    "U": "SEGMENTO",
    "W": "CT",
    "X": "MONTO",
    "Z": "CONCEPTO",
}

# Columns that decide how far down the section goes. The last row is the
# HIGHEST of the three, so a blank 'AGRUPADOR ABIERTO' on the final row cannot
# silently cut the policy short.
#   C = AGRUPADOR ABIERTO   B = BU   N = Reserva (the amount)
COLUMNAS_PILOTO = ["C", "B", "N"]

# Safety net so a stray value 40,000 rows down cannot make the run write a
# million formulas.
MAX_FILAS = 50_000


# ----------------------------------------------------------------------
# The seven columns of the 'PÓLIZA SAP' section
# ----------------------------------------------------------------------
# 'plantilla' is written with {f} where the first data row goes. It is applied
# to the WHOLE column range in a single call and Excel shifts the relative
# references down by itself — the same thing that happens when you grab the
# little square at the corner of a cell and drag it down.
#
# IMPORTANT: .Formula always speaks US English, whatever the language of the
# installed Excel. So VLOOKUP and commas here are correct even on a Spanish
# Excel; writing BUSCARV would fail. (The localized twin is .FormulaLocal, and
# using it would make the module break on an English machine — so we don't.)
COLUMNAS_POLIZA = [
    {
        "col": "R",
        "titulo": "CUENTA SAP",
        "tipo": "formula",
        "plantilla": "=VLOOKUP(C{f},{cat}!$E:$H,4,0)",
    },
    {
        "col": "T",
        "titulo": "CeBe",
        "tipo": "formula",
        "plantilla": "=VLOOKUP(B{f},{cat}!$B:$D,3,0)",
    },
    {
        "col": "U",
        "titulo": "SEGMENTO",
        "tipo": "valor",
        "valor": "MXC0000",
    },
    {
        "col": "V",
        "titulo": "Asignación",
        "tipo": "formula",
        "plantilla": '=UPPER(LEFT(B{f},3)&"\\"&LEFT(G{f},2)&"\\"&H{f})',
    },
    {
        "col": "W",
        "titulo": "CT",
        "tipo": "valor",
        "valor": 40,          # posting key: 40 = cargo, 50 = abono (see rows 1-2)
    },
    {
        "col": "X",
        "titulo": "MONTO",
        "tipo": "formula",
        "plantilla": "=IFERROR(ROUND(N{f},2),0)",
    },
    {
        "col": "Z",
        "titulo": "CONCEPTO (descripción SAP)",
        "tipo": "formula",
        "plantilla": (
            '=VLOOKUP(C{f},{cat}!$E:$I,5,0)&" "'
            '&VLOOKUP({hoja}!G{f},{cat}!$J:$K,2,0)'
        ),
    },
]

# Column whose total is reported back as the control figure of the run.
COLUMNA_CONTROL = "X"   # MONTO


# ----------------------------------------------------------------------
# Behaviour flags
# ----------------------------------------------------------------------
# True writes the formulas, lets Excel calculate, and then replaces them with
# their results. Auditable while reviewing, bullet-proof for the SAP upload
# (no #N/A, no dependency on the 'catalogo' sheet surviving the trip).
# Set to False to leave the formulas alive.
CONVERTIR_A_VALORES = True

# True stops the run (and leaves the copy with its formulas intact, so the
# errors are visible) when any lookup fails. False writes anyway and only
# reports. False is the default: the analyst usually wants to SEE which rows
# failed rather than get nothing.
ABORTAR_SI_HAY_ERRORES = False

# True writes CT as the text "40" instead of the number 40. The real file
# stores it as a number, so False matches what is there today.
CT_COMO_TEXTO = False

# Clear the seven columns before writing, so a shorter month never leaves rows
# from a longer previous run hanging below the data. Only contents are cleared,
# never formats, and never above the first data row (the legend in rows 1-5
# must survive).
LIMPIAR_ANTES = True

# How many offending cells to list per column in the summary.
MAX_ERRORES_REPORTADOS = 10


# ======================================================================
# Layout helpers
# ======================================================================
# _norm, _abrir_excel, _cerrar_excel and _hoja used to live here; they now
# come from utils.py so 'Comparativos' can reuse them (see the imports at
# the top). Everything below is specific to the 'PÓLIZA SAP' section.


def _localizar_fila_encabezado(ws):
    """
    Find the 1-based header row of the 'PÓLIZA SAP' section.

    It is row 6 in today's file, not row 1: above it sit a legend, the
    'PÓLIZA SAP' banner and a row of SAP field names. That padding has a habit
    of growing, so instead of hard-coding 6 we scan the first rows for a row
    that carries 'AGRUPADOR ABIERTO' in C *and* 'CUENTA SAP' in R. Same trick
    the 'Variaciones' module uses on the 'Base' sheet.
    """
    for fila in range(1, 26):
        if all(
            _norm(ws.Range(f"{col}{fila}").Value).startswith(_norm(texto))
            for col, texto in ANCLA_ENCABEZADO.items()
        ):
            return fila
    return FILA_ENCABEZADO_DEFAULT


def _validar_encabezados(ws, fila_encabezado):
    """
    Confirm the seven target columns are still where we think they are.

    Cheap insurance: if somebody inserts a column, every formula would land one
    column off and the policy would go up to SAP with the cost centre in the
    account field. Better to stop here with a readable message.
    """
    errores = []
    for col, esperado in ENCABEZADOS_ESPERADOS.items():
        real = _norm(ws.Range(f"{col}{fila_encabezado}").Value)
        if not real.startswith(_norm(esperado)):
            errores.append(f"     • {col}{fila_encabezado}: esperaba "
                           f"'{esperado}' y encontré '{real or '(vacío)'}'")
    if errores:
        raise ValueError(
            f"Los encabezados de la sección 'PÓLIZA SAP' no están donde los "
            f"esperaba (fila {fila_encabezado}):\n\n"
            + "\n".join(errores)
            + "\n\n¿Alguien insertó o movió una columna?"
        )


def _ultima_fila(ws, fila_inicio):
    """
    Last row of the section that actually carries data.

    Ctrl+Shift+Up from the bottom of the sheet (End(xlUp)) finds the last USED
    cell, which can overshoot when a formula below returns "" or a cell only
    carries formatting. So we walk back up while the value is empty, and we do
    it over THREE pilot columns and keep the highest: a blank 'AGRUPADOR
    ABIERTO' on the last row must not chop the policy short.
    """
    mejor = None
    for columna in COLUMNAS_PILOTO:
        fila = ws.Cells(ws.Rows.Count, columna).End(XL_UP).Row
        while fila >= fila_inicio:
            valor = ws.Range(f"{columna}{fila}").Value
            if valor is not None and str(valor).strip() != "":
                break
            fila -= 1
        if fila >= fila_inicio:
            mejor = fila if mejor is None else max(mejor, fila)

    if mejor is None:
        return None
    return min(mejor, fila_inicio + MAX_FILAS - 1)


def _detectar_errores(ws, columna, fila_inicio, fila_fin):
    """
    List the error cells of one column, asking Excel instead of scanning.

    SpecialCells(xlCellTypeFormulas, xlErrors) is Excel's own 'Go To Special ->
    Formulas -> Errors'. It returns only the offending cells, so a 5,000-row
    column costs one call instead of 5,000. When there are none it raises a
    com_error, which is why the whole thing sits inside a try.

    Has to run BEFORE the formulas are turned into values — afterwards there are
    no formulas left to inspect.

    Returns (detalle, total): the first MAX_ERRORES_REPORTADOS offenders, and how
    many there really are.
    """
    rango = ws.Range(f"{columna}{fila_inicio}:{columna}{fila_fin}")
    try:
        erroneas = rango.SpecialCells(XL_CELLTYPE_FORMULAS, XL_ERRORS)
    except Exception:
        return [], 0   # no error cells in the range

    detalle = []
    for celda in erroneas:
        detalle.append({
            "celda": str(celda.Address(False, False)),   # 'R123'
            "error": str(celda.Text),                    # '#N/A'
            "fila":  int(celda.Row),
        })
        if len(detalle) >= MAX_ERRORES_REPORTADOS:
            break

    # .Count is the real total even when we stopped listing early.
    return detalle, int(erroneas.Count)


def _congelar_a_valores(app, ws, columna, fila_inicio, fila_fin):
    """
    Replace the formulas of one column with their results.

    Copy + PasteSpecial(xlPasteValues), NOT `rango.Value = rango.Value`.

    The obvious one-liner has a nasty bug here: reading .Value gives Python the
    string '0004209900' (CUENTA SAP keeps its leading zeros), and writing a
    string back into a General-formatted cell makes Excel re-parse it exactly as
    if it had been typed — so it becomes the number 4209900 and the four leading
    zeros are gone. SAP would reject the whole policy.

    PasteSpecial pastes the value Excel already computed, without re-parsing
    anything. It is literally 'Copiar → Pegado especial → Valores'.
    """
    rango = ws.Range(f"{columna}{fila_inicio}:{columna}{fila_fin}")
    try:
        rango.Copy()
        rango.PasteSpecial(Paste=XL_PASTE_VALUES)
    finally:
        # Always clear the marching ants, even if the paste failed, so the
        # user's clipboard is not left holding a chunk of the workbook.
        app.CutCopyMode = False


# ======================================================================
# Main entry point
# ======================================================================
def generar_poliza_sap(
    ruta_libro,
    ruta_salida=None,
    sobrescribir_original=False,
    convertir_a_valores=None,
    callback_status=None,
):
    """
    Fill the 'PÓLIZA SAP' section of the WHSL layout workbook.

    Args:
        ruta_libro (str): path to 'Layout Creación Reservas DA WHSL_<mes>.xlsb'.
        ruta_salida (str|None): folder for the output copy. None -> the 'Output'
            folder resolved by carpeta_output_desde_input().
        sobrescribir_original (bool): True writes into the source file. Default
            False, same policy as the 'Variaciones' module.
        convertir_a_valores (bool|None): None -> the CONVERTIR_A_VALORES flag.
        callback_status (function|None): progress reporting for the GUI.

    Returns:
        dict: ruta, hoja, fila_encabezado, fila_inicio, fila_fin, filas,
              columnas, total_control, errores, total_errores,
              convertido_a_valores.
    """
    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    if convertir_a_valores is None:
        convertir_a_valores = CONVERTIR_A_VALORES

    if not os.path.exists(ruta_libro):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta_libro}")

    # --- Where the result goes ----------------------------------------
    if sobrescribir_original:
        ruta_final = os.path.abspath(ruta_libro)
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
        update_status("📑 Copiando el libro (el original no se toca)...")
        # copy2 keeps the .xlsb bytes untouched, so Save() later writes the very
        # same format back. No SaveAs, no FileFormat number to get wrong.
        shutil.copy2(ruta_libro, ruta_final)

    app = wb = pythoncom = None
    calculo_previo = None
    try:
        update_status("📂 Abriendo Excel...")
        app, pythoncom = _abrir_excel()

        wb = app.Workbooks.Open(ruta_final, UpdateLinks=0, ReadOnly=False)
        ws = _hoja(wb, HOJA_CREACIONES)
        _hoja(wb, HOJA_CATALOGO)   # fail early if the lookups have nothing to hit

        # --- Where does the section start? ----------------------------
        fila_encabezado = _localizar_fila_encabezado(ws)
        _validar_encabezados(ws, fila_encabezado)
        fila_inicio = fila_encabezado + 1

        # --- How far down does it go? ---------------------------------
        fila_fin = _ultima_fila(ws, fila_inicio)
        if fila_fin is None:
            raise ValueError(
                f"No encontré datos en la pestaña '{HOJA_CREACIONES}': las "
                f"columnas piloto {COLUMNAS_PILOTO} están vacías a partir de la "
                f"fila {fila_inicio}.\n\n¿Es el archivo correcto?"
            )
        n_filas = fila_fin - fila_inicio + 1
        update_status(
            f"📊 Sección 'PÓLIZA SAP': encabezado en la fila {fila_encabezado}, "
            f"datos {fila_inicio}–{fila_fin} ({n_filas:,} renglones)"
        )

        # Manual calculation while writing: with 7 columns x thousands of rows,
        # letting Excel recalculate after every single write is the difference
        # between seconds and minutes.
        calculo_previo = app.Calculation
        app.Calculation = XL_CALC_MANUAL

        # --- Clear old contents ---------------------------------------
        # Never above fila_inicio: rows 1-5 carry the legend, the 'PÓLIZA SAP'
        # banner and the SAP field names, and none of that is ours to touch.
        if LIMPIAR_ANTES:
            fila_usada = max(
                int(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1), fila_fin
            )
            if fila_usada >= fila_inicio:
                for spec in COLUMNAS_POLIZA:
                    ws.Range(
                        f"{spec['col']}{fila_inicio}:{spec['col']}{fila_usada}"
                    ).ClearContents()

        # --- Write the seven columns ----------------------------------
        for spec in COLUMNAS_POLIZA:
            col = spec["col"]
            rango = ws.Range(f"{col}{fila_inicio}:{col}{fila_fin}")
            update_status(f"✍️  Columna {col} — {spec['titulo']}...")

            if spec["tipo"] == "formula":
                # One assignment fills the entire column. Excel shifts the
                # relative references itself, exactly like dragging the fill
                # handle down.
                rango.Formula = spec["plantilla"].format(
                    f=fila_inicio,
                    cat=HOJA_CATALOGO,
                    hoja=HOJA_CREACIONES,
                )
            else:
                valor = spec["valor"]
                if col == "W" and CT_COMO_TEXTO:
                    rango.NumberFormat = "@"
                    valor = str(valor)
                rango.Value = valor

        # --- Calculate --------------------------------------------------
        update_status("🧮 Recalculando el libro...")
        app.Calculation = XL_CALC_AUTOMATIC
        app.CalculateFullRebuild()

        # --- Look for failed lookups -----------------------------------
        update_status("🔎 Revisando #N/A en los VLOOKUP...")
        errores, total_errores = [], 0
        for spec in COLUMNAS_POLIZA:
            if spec["tipo"] != "formula":
                continue
            detalle, cuantos = _detectar_errores(
                ws, spec["col"], fila_inicio, fila_fin
            )
            if not cuantos:
                continue
            total_errores += cuantos
            errores.append({
                "columna": spec["col"],
                "titulo":  spec["titulo"],
                "total":   cuantos,
                "celdas":  detalle,
            })

        if errores and ABORTAR_SI_HAY_ERRORES:
            resumen_err = ", ".join(
                f"{e['columna']} ({e['total']})" for e in errores
            )
            raise ValueError(
                f"{total_errores} celda(s) con error en la póliza: {resumen_err}.\n\n"
                f"Se dejó el archivo con las fórmulas vivas para que puedas ver "
                f"qué no empató en '{HOJA_CATALOGO}':\n{ruta_final}"
            )

        # --- Control total ---------------------------------------------
        total_control = float(app.WorksheetFunction.Sum(
            ws.Range(f"{COLUMNA_CONTROL}{fila_inicio}:"
                     f"{COLUMNA_CONTROL}{fila_fin}")
        ))

        # --- Freeze the formulas into values ---------------------------
        if convertir_a_valores:
            update_status("🧊 Congelando fórmulas a valores...")
            for spec in COLUMNAS_POLIZA:
                if spec["tipo"] != "formula":
                    continue
                _congelar_a_valores(app, ws, spec["col"], fila_inicio, fila_fin)

        update_status("💾 Guardando...")
        wb.Save()

        return {
            "ruta":                 ruta_final,
            "hoja":                 HOJA_CREACIONES,
            "fila_encabezado":      fila_encabezado,
            "fila_inicio":          fila_inicio,
            "fila_fin":             fila_fin,
            "filas":                n_filas,
            "columnas":             [f"{s['col']} — {s['titulo']}"
                                     for s in COLUMNAS_POLIZA],
            "total_control":        round(total_control, 2),
            "columna_control":      COLUMNA_CONTROL,
            "errores":              errores,
            "total_errores":        total_errores,
            "convertido_a_valores": bool(convertir_a_valores),
        }

    finally:
        try:
            if app is not None and calculo_previo is not None:
                app.Calculation = calculo_previo
        except Exception:
            pass
        _cerrar_excel(app, wb, pythoncom)


# ----------------------------------------------------------------------
# Standalone run (no GUI, no controller):  python Poliza_SAP.py
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=== Prueba de 'PÓLIZA SAP' (standalone) ===")
    ruta = input("Ruta del Layout WHSL (.xlsb): ").strip().strip('"')

    try:
        resumen = generar_poliza_sap(ruta, callback_status=print)
        print("\n🎉 LISTO 🎉")
        for clave, valor in resumen.items():
            if clave == "errores":
                print(f"  {clave:<22}: {len(valor)} columna(s) con error")
                for e in valor:
                    celdas = ", ".join(c["celda"] for c in e["celdas"])
                    print(f"      {e['columna']} ({e['total']}): {celdas}")
            else:
                print(f"  {clave:<22}: {valor}")
    except Exception as e:
        print(f"❌ Error generando la póliza: {e}")
        raise
