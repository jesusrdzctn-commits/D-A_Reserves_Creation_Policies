"""
RDA Pólizas D&A — Comparativos
===============================

Builds two extra sheets of '07 Archivo Creacion de Reservas PMF OT.xlsx', each
holding TWO REAL Excel PivotTables stacked vertically:

    'Comparativo BU Mes'          Rows: BU         Columns: Mes   Values: Σ Importe
    'Comparativo Agrupador Mes'   Rows: Agrupador  Columns: Mes   Values: Σ Importe

In both sheets the TOP table shows LAST year's months and the BOTTOM one the
CURRENT year's months, so the two years sit one above the other:

    A1   Comparativo BU vs Mes                  <- sheet heading (blue)
    A3   Creaciones 2025                        <- grey banner
    A4   [PivotTable, columns filtered to M01-25 ... M12-25]
    ...
    A??  Creaciones 2026                        <- grey banner
    A??  [PivotTable, columns filtered to M01-26 ... M12-26]

Only the months that actually exist in 'Base' become columns (same as a pivot
you would build by hand), and both Grand Totals are on: the column on the right
is the year total per BU/Agrupador, the row at the bottom is the total per month.

Why win32com and not openpyxl
-----------------------------
openpyxl can READ and PRESERVE a PivotTable that already exists, but it cannot
CREATE one: a pivot needs its own cache (xl/pivotCache/*.xml) wired to the
source range, and openpyxl ships no writer for that. Excel COM asks Excel itself
to build it, so what lands in the file is a real PivotTable — with its filter
dropdowns and its 'Actualizar' button — exactly like the 16 pivots the workbook
already carries.

The trade-off, stated out loud: this module needs Excel installed on whatever
machine runs the .exe. So it is called as an OPTIONAL last step of
construir_variaciones(): if Excel or pywin32 is missing, 'Variaciones' is
already written and saved and the run reports the comparativos as skipped
instead of failing.

ORDER MATTERS: openpyxl writes and saves FIRST, this module opens the saved file
SECOND. Never the other way around — an openpyxl round-trip over a freshly
created pivot cache is precisely the kind of thing that loses it.

This module does NOT depend on pandas, openpyxl, SAP or tkinter: it can be
tested on its own with `python Comparativos.py`.
"""

import os
from datetime import datetime

from utils import (
    abrir_excel,
    carpeta_output_desde_input,
    cerrar_excel,
    hoja_por_nombre,
    norm_texto,
    parse_clave_mes,
)

# ----------------------------------------------------------------------
# Excel COM constants
# ----------------------------------------------------------------------
# Hard-coded instead of using win32com.client.constants, for the same reason as
# in Poliza_SAP.py: those only exist after gencache.EnsureDispatch(), which is
# exactly the call that breaks inside a PyInstaller .exe. Plain numbers always
# work.
XL_UP           = -4162   # xlUp
XL_DATABASE     = 1       # xlDatabase          (source is a worksheet range)
XL_ROW_FIELD    = 1       # xlRowField
XL_COLUMN_FIELD = 2       # xlColumnField
XL_SUM          = -4157   # xlSum
XL_ASCENDING    = 1       # xlAscending
XL_TABULAR_ROW  = 1       # xlTabularRow        (row field shown in its own column)


# ----------------------------------------------------------------------
# Sheets and source columns
# ----------------------------------------------------------------------
HOJA_BASE = "Base"

HOJA_COMPARATIVO_BU        = "Comparativo BU Mes"
HOJA_COMPARATIVO_AGRUPADOR = "Comparativo Agrupador Mes"

# Headers as they read in 'Base'. Matched by NORMALISED name (trimmed,
# upper-cased, accent-free), because the real file ships 'Importe ' with a
# trailing space.
COL_MES       = "Mes"
COL_BU        = "BU"
COL_AGRUPADOR = "Agrupador"
COL_IMPORTE   = "Importe"

# The header row of 'Base' is not row 1 (empty rows plus a 'Seleccionar /
# Escritura / Formula' hint row sit above it), so it is detected by looking for
# 'Mes' + 'Agrupador'. This is only the fallback if detection fails.
FILA_ENCABEZADO_BASE_DEFAULT = 7

# How far to look when hunting for the header row / reading the header itself.
MAX_FILAS_BUSQUEDA   = 30
MAX_COLUMNAS_LECTURA = 40

# Upper bound on the source range, so a corrupt sheet cannot make us build a
# pivot cache over a million rows.
MAX_FILAS = 200_000

# The two sheets to build: (sheet name, heading, row field, pivot name prefix).
COMPARATIVOS = [
    (HOJA_COMPARATIVO_BU,        "Comparativo BU vs Mes",        COL_BU,        "ptComparativoBU"),
    (HOJA_COMPARATIVO_AGRUPADOR, "Comparativo Agrupador vs Mes", COL_AGRUPADOR, "ptComparativoAgrupador"),
]


# ----------------------------------------------------------------------
# Layout
# ----------------------------------------------------------------------
FILA_ENCABEZADO_HOJA = 1   # sheet heading
FILA_PRIMERA_TABLA   = 4   # first pivot lands here; its banner goes one row above
FILAS_ENTRE_TABLAS   = 3   # blank rows between one pivot's last row and the next banner

ANCHO_COLUMNA_A = 31.33    # same as the 'Concepto' column of 'Variaciones'

# Caption of the value field. It must NOT be identical to the source field name
# ('Importe '), or Excel rejects it as a duplicate.
NOMBRE_CAMPO_VALOR = "Suma de Importe"

# Banner above each pivot. Same wording as the 'Variaciones' sheet.
TITULO_TABLA = "Creaciones {anio}"


# ----------------------------------------------------------------------
# Styles — same palette as the 'Variaciones' sheet
# ----------------------------------------------------------------------
GRIS_CREACIONES = "BFBFBF"   # banner above each pivot
AZUL_TITULO     = "DCE6F2"   # sheet heading

FUENTE = "Calibri"
TAM    = 10

CONTABLE_2 = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'

# Built-in PivotTable style. Set to None to leave Excel's default. An unknown
# name is ignored rather than fatal — a cosmetic detail must never kill a run.
ESTILO_TABLA = "PivotStyleMedium9"


def _color_excel(rgb_hex):
    """
    'DCE6F2' -> the integer Excel's Interior.Color expects.

    Excel stores colours BACKWARDS: the low byte is RED and the high byte is
    BLUE (BGR), not the RGB order every designer writes. Hand it 0xDCE6F2 raw
    and a pale blue comes out a pale orange. Same three cans of paint, opposite
    order on the shelf.
    """
    r = int(rgb_hex[0:2], 16)
    g = int(rgb_hex[2:4], 16)
    b = int(rgb_hex[4:6], 16)
    return r + (g << 8) + (b << 16)


# ----------------------------------------------------------------------
# Reading the 'Base' sheet through COM
# ----------------------------------------------------------------------
def _bloque(ws, fila1, col1, fila2, col2):
    """
    Read a rectangle in ONE COM call and always return a list of rows.

    Every attribute hop across the COM boundary costs about a millisecond, so
    reading 30x40 cells one at a time is 1,200 round trips. `Range.Value` over
    the whole rectangle is a single trip. Excel collapses a 1x1 range to a
    scalar and a 1xN range to a flat tuple, so both shapes are normalised here.
    """
    valores = ws.Range(ws.Cells(fila1, col1), ws.Cells(fila2, col2)).Value

    if valores is None:
        return [[None] * (col2 - col1 + 1)]
    if not isinstance(valores, tuple):          # single cell
        return [[valores]]
    if valores and not isinstance(valores[0], tuple):   # single row
        return [list(valores)]
    return [list(fila) for fila in valores]


def _localizar_fila_encabezado(ws):
    """
    Find the 1-based header row of 'Base'.

    The sheet carries empty rows and a hint row above the real headers, and that
    padding has changed before. Scanning the top of the sheet for 'Mes' +
    'Agrupador' survives somebody inserting a row — same trick the openpyxl side
    uses, so both modules agree on where the table starts.
    """
    buscados = {norm_texto(COL_MES), norm_texto(COL_AGRUPADOR)}
    filas = _bloque(ws, 1, 1, MAX_FILAS_BUSQUEDA, MAX_COLUMNAS_LECTURA)

    for desplazamiento, fila in enumerate(filas):
        if buscados <= {norm_texto(v) for v in fila}:
            return desplazamiento + 1
    return FILA_ENCABEZADO_BASE_DEFAULT


def _encabezados(ws, fila_encabezado):
    """
    Return the CONTIGUOUS header names of 'Base', left to right, stopping at the
    first empty one.

    Stopping is not cosmetic: the real file has a column M with data but NO
    header, and a pivot cache refuses to be built over a range with a blank
    header ("The PivotTable field name is not valid"). Cutting the source range
    at the last named column is what keeps that from happening.
    """
    fila = _bloque(ws, fila_encabezado, 1, fila_encabezado, MAX_COLUMNAS_LECTURA)[0]

    encabezados = []
    for valor in fila:
        if valor is None or str(valor).strip() == "":
            break
        encabezados.append(str(valor))
    return encabezados


def _indice_columna(encabezados, nombre):
    """
    1-based position of a column inside `encabezados`, matched by normalised
    name. Fails loud with the list of what was actually found: guessing an index
    here would mean building a pivot over the wrong column, which looks right
    and adds up wrong.
    """
    objetivo = norm_texto(nombre)
    for i, encabezado in enumerate(encabezados, start=1):
        if norm_texto(encabezado) == objetivo:
            return i
    raise ValueError(
        f"No encontré la columna '{nombre}' en la pestaña '{HOJA_BASE}'.\n\n"
        f"Columnas encontradas: {encabezados}"
    )


def _ultima_fila(ws, fila_inicio, columnas_piloto):
    """
    Last row of 'Base' that actually carries data.

    End(xlUp) from the bottom finds the last USED cell, which overshoots when a
    cell only carries formatting, so we walk back up while it is empty. Done
    over three pilot columns keeping the highest, because a blank 'BU' on the
    last row must not chop the table short.
    """
    mejor = None
    for columna in columnas_piloto:
        fila = ws.Cells(ws.Rows.Count, columna).End(XL_UP).Row
        while fila >= fila_inicio:
            valor = ws.Cells(fila, columna).Value
            if valor is not None and str(valor).strip() != "":
                break
            fila -= 1
        if fila >= fila_inicio:
            mejor = fila if mejor is None else max(mejor, fila)

    if mejor is None:
        return None
    return min(mejor, fila_inicio + MAX_FILAS - 1)


def _meses_por_anio(ws, fila_inicio, fila_fin, columna_mes):
    """
    Read the 'Mes' column once and return {año: [claves de mes, ordenadas]}.

    Keys that do not parse as 'M06-26' (blanks, typos) land under None and are
    simply never made visible in either pivot; they are reported so nobody has
    to wonder where an amount went.
    """
    valores = _bloque(ws, fila_inicio, columna_mes, fila_fin, columna_mes)

    por_anio = {}
    for fila in valores:
        valor = fila[0]
        if valor is None or str(valor).strip() == "":
            continue
        clave = str(valor).strip().upper()
        try:
            anio = parse_clave_mes(clave).year
        except ValueError:
            anio = None
        por_anio.setdefault(anio, set()).add(clave)

    return {anio: sorted(claves) for anio, claves in por_anio.items()}


# ----------------------------------------------------------------------
# Building one PivotTable
# ----------------------------------------------------------------------
def _nombres_de_tablas(wb):
    """Every PivotTable name already in the workbook (they must be unique)."""
    nombres = set()
    for ws in wb.Worksheets:
        try:
            for pt in ws.PivotTables():
                nombres.add(str(pt.Name))
        except Exception:
            # A sheet with no pivots (or a chart sheet) has nothing to add.
            pass
    return nombres


def _nombre_libre(base, ocupados):
    """'ptComparativoBU_2026' or the first free '..._2', '..._3' if it is taken."""
    nombre, i = base, 1
    while nombre in ocupados:
        i += 1
        nombre = f"{base}_{i}"
    ocupados.add(nombre)
    return nombre


def _campos(pt):
    """{normalised field name: real field name} for one PivotTable."""
    return {norm_texto(campo.Name): str(campo.Name) for campo in pt.PivotFields()}


def _filtrar_meses(campo, meses_visibles):
    """
    Leave only `meses_visibles` showing on the 'Mes' column field.

    Two rules, both learned the hard way:

      1. Show the wanted items FIRST, hide the rest afterwards. Excel refuses to
         leave a field with zero visible items, so hiding first can blow up
         halfway through with a field that is already empty.
      2. Drive the PivotItem OBJECTS, not their names. `PivotItems("M06-26")`
         is a name lookup that Excel can misread; the object is unambiguous.
    """
    items = list(campo.PivotItems())
    visibles = set(meses_visibles)

    for item in items:
        if str(item.Name) in visibles and not item.Visible:
            item.Visible = True
    for item in items:
        if str(item.Name) not in visibles and item.Visible:
            item.Visible = False


def _crear_tabla(cache, ws, fila, nombre_tabla, campo_fila, meses_visibles):
    """
    Create one PivotTable at row `fila`, column A, and return it.

    ManualUpdate is on for the whole build: with it off, Excel recalculates the
    whole pivot after every single property we touch — one field, one data
    field, then one redraw per hidden month. On a dozen months that is the
    difference between a blink and several seconds.
    """
    destino = f"'{ws.Name}'!R{fila}C1"
    pt = cache.CreatePivotTable(TableDestination=destino, TableName=nombre_tabla)

    pt.ManualUpdate = True
    try:
        campos = _campos(pt)

        c_fila = pt.PivotFields(campos[norm_texto(campo_fila)])
        c_fila.Orientation = XL_ROW_FIELD
        c_fila.Position = 1

        c_mes = pt.PivotFields(campos[norm_texto(COL_MES)])
        c_mes.Orientation = XL_COLUMN_FIELD
        c_mes.Position = 1

        valor = pt.AddDataField(
            pt.PivotFields(campos[norm_texto(COL_IMPORTE)]),
            NOMBRE_CAMPO_VALOR,
            XL_SUM,
        )
        valor.NumberFormat = CONTABLE_2

        _filtrar_meses(c_mes, meses_visibles)

        # 'Mes' is text ('M06-26'), so an A-Z sort of ONE year's keys is also
        # the chronological order: M01 < M02 < ... < M12. It only works because
        # each table shows a single year — mixing years would sort M01-25 next
        # to M01-26.
        c_mes.AutoSort(XL_ASCENDING, c_mes.Name)

        pt.RowGrand    = True   # 'Grand Total' COLUMN on the right  (year per row)
        pt.ColumnGrand = True   # 'Grand Total' ROW at the bottom    (total per month)
        pt.RowAxisLayout(XL_TABULAR_ROW)

        if ESTILO_TABLA:
            try:
                pt.TableStyle2 = ESTILO_TABLA
            except Exception:
                pass   # unknown style name in this Excel build — cosmetic only
    finally:
        # Always hand the pivot back in auto-update mode, even if something
        # above failed: a workbook saved with ManualUpdate still on opens with a
        # pivot that refuses to refresh.
        pt.ManualUpdate = False

    return pt


def _banner(ws, fila, texto, color, tam=TAM):
    """Write one of the coloured title cells above a pivot."""
    celda = ws.Cells(fila, 1)
    celda.Value = texto
    celda.Font.Name = FUENTE
    celda.Font.Size = tam
    celda.Font.Bold = True
    celda.Interior.Color = _color_excel(color)
    return celda


def _preparar_hoja(wb, nombre):
    """
    Drop the sheet if it is already there and create it again in the SAME tab
    position, so re-running the process does not shuffle the workbook's tabs.
    """
    posicion = None
    for i in range(1, wb.Worksheets.Count + 1):
        if str(wb.Worksheets(i).Name) == nombre:
            posicion = i
            wb.Worksheets(i).Delete()   # DisplayAlerts is already off
            break

    if posicion is not None and posicion <= wb.Worksheets.Count:
        ws = wb.Worksheets.Add(Before=wb.Worksheets(posicion))
    else:
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))

    ws.Name = nombre
    return ws


def _construir_hoja(wb, cache, nombre_hoja, encabezado_hoja, campo_fila,
                    prefijo_tabla, anios, meses_por_anio, ocupados, update_status):
    """
    Build one comparison sheet: heading, then one banner + one pivot per year,
    stacked top to bottom. Returns (ws, [info de cada tabla], [advertencias]).
    """
    ws = _preparar_hoja(wb, nombre_hoja)
    _banner(ws, FILA_ENCABEZADO_HOJA, encabezado_hoja, AZUL_TITULO, tam=TAM + 1)

    tablas, advertencias = [], []
    fila = FILA_PRIMERA_TABLA

    for anio in anios:
        meses = meses_por_anio.get(anio, [])
        _banner(ws, fila - 1, TITULO_TABLA.format(anio=anio), GRIS_CREACIONES)

        if not meses:
            # No data for that year. Writing the note instead of the pivot is
            # deliberate: a PivotTable cannot exist with every column hidden, so
            # the alternative would be either a crash or a table silently
            # showing the OTHER year's months under this year's banner.
            aviso = (f"Sin movimientos de {anio} en la pestaña '{HOJA_BASE}' "
                     f"— no hay nada que tabular.")
            ws.Cells(fila, 1).Value = aviso
            ws.Cells(fila, 1).Font.Italic = True
            advertencias.append(f"{nombre_hoja}: {aviso}")
            tablas.append({"hoja": nombre_hoja, "tabla": None, "anio": anio,
                           "meses": [], "total": None, "celda": None})
            fila += 1 + FILAS_ENTRE_TABLAS + 1
            continue

        update_status(
            f"   • {nombre_hoja} — {anio}: {len(meses)} mes(es) "
            f"({meses[0]} → {meses[-1]})"
        )

        nombre_tabla = _nombre_libre(f"{prefijo_tabla}_{anio}", ocupados)
        pt = _crear_tabla(cache, ws, fila, nombre_tabla, campo_fila, meses)

        rango = pt.TableRange2
        fila_fin = rango.Row + rango.Rows.Count - 1

        # Bottom-right cell of the pivot = Grand Total of Grand Totals. Cheap
        # control figure: it has to match the sum of that year's 'Importe'.
        try:
            total = rango.Cells(rango.Rows.Count, rango.Columns.Count).Value
            total = round(float(total), 2) if total is not None else None
        except Exception:
            total = None

        tablas.append({
            "hoja":   nombre_hoja,
            "tabla":  nombre_tabla,
            "anio":   anio,
            "meses":  meses,
            "total":  total,
            "celda":  f"A{fila}",
        })

        fila = fila_fin + FILAS_ENTRE_TABLAS + 2   # +1 blank, +1 for the next banner

    # Widths last, once everything is on the sheet.
    try:
        ws.UsedRange.Columns.AutoFit()
        ws.Columns(1).ColumnWidth = ANCHO_COLUMNA_A
    except Exception:
        pass

    ws.Cells(1, 1).Select()   # leave the sheet scrolled to the top
    return ws, tablas, advertencias


# ----------------------------------------------------------------------
# Main entry point
# ----------------------------------------------------------------------
def construir_comparativos(
    ruta_libro,
    ruta_salida=None,
    mes=None,
    sobrescribir_original=True,
    callback_status=None,
):
    """
    Build the 'Comparativo BU Mes' and 'Comparativo Agrupador Mes' sheets.

    Args:
        ruta_libro (str): path to the Reservas workbook. Normally the COPY that
            construir_variaciones() just wrote in 'Output'.
        ruta_salida (str|None): only used when sobrescribir_original is False.
            None -> the 'Output' folder resolved by carpeta_output_desde_input().
        mes (datetime|str|None): reference month, deciding which year is
            'current' and which is 'last'. None -> today. A month key ('M07-26')
            re-runs a past period.
        sobrescribir_original (bool): default TRUE here, unlike the other two
            modules, because this one is normally handed the output copy that
            'Variaciones' just produced — copying a copy would leave the pivots
            in a second file and the matrices in the first.
        callback_status (function|None): progress reporting for the GUI.

    Returns:
        dict: ruta, anio_actual, anio_anterior, fila_encabezado, filas_base,
              hojas, tablas, meses_por_anio, advertencias.
    """
    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    if not os.path.exists(ruta_libro):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta_libro}")

    # --- Which two years are we comparing? ----------------------------
    if mes is None:
        fecha = datetime.today()
    elif isinstance(mes, str):
        fecha = parse_clave_mes(mes)
    else:
        fecha = mes
    anio_actual   = fecha.year
    anio_anterior = anio_actual - 1

    # --- Where the result goes ----------------------------------------
    if sobrescribir_original:
        ruta_final = os.path.abspath(ruta_libro)
    else:
        import shutil
        if not ruta_salida:
            ruta_salida = carpeta_output_desde_input(
                os.path.dirname(os.path.abspath(ruta_libro))
            )
        os.makedirs(ruta_salida, exist_ok=True)
        base, ext = os.path.splitext(os.path.basename(ruta_libro))
        ruta_final = os.path.abspath(
            os.path.join(ruta_salida, f"{base}_Comparativos{ext}")
        )
        shutil.copy2(ruta_libro, ruta_final)

    app = wb = pythoncom = None
    try:
        update_status("📂 Abriendo Excel para las tablas dinámicas...")
        app, pythoncom = abrir_excel()

        wb = app.Workbooks.Open(ruta_final, UpdateLinks=0, ReadOnly=False)
        ws_base = hoja_por_nombre(wb, HOJA_BASE)

        # --- Where does the source table start and end? ---------------
        fila_encabezado = _localizar_fila_encabezado(ws_base)
        encabezados     = _encabezados(ws_base, fila_encabezado)
        if not encabezados:
            raise ValueError(
                f"La fila {fila_encabezado} de la pestaña '{HOJA_BASE}' no tiene "
                f"encabezados. ¿Es el archivo correcto?"
            )

        c_mes = _indice_columna(encabezados, COL_MES)
        c_bu  = _indice_columna(encabezados, COL_BU)
        c_agr = _indice_columna(encabezados, COL_AGRUPADOR)
        _indice_columna(encabezados, COL_IMPORTE)   # fail early if it is missing

        fila_inicio = fila_encabezado + 1
        fila_fin    = _ultima_fila(ws_base, fila_inicio, (c_mes, c_bu, c_agr))
        if fila_fin is None:
            raise ValueError(
                f"No encontré datos en la pestaña '{HOJA_BASE}' a partir de la "
                f"fila {fila_inicio}.\n\n¿Es el archivo correcto?"
            )

        n_filas = fila_fin - fila_inicio + 1
        update_status(
            f"📊 Origen: '{HOJA_BASE}' filas {fila_encabezado}–{fila_fin} "
            f"({n_filas:,} renglones, {len(encabezados)} columnas)"
        )

        # --- Which months exist, and under which year? ----------------
        meses_por_anio = _meses_por_anio(ws_base, fila_inicio, fila_fin, c_mes)
        advertencias = []
        if None in meses_por_anio:
            raros = sorted(meses_por_anio.pop(None))[:8]
            advertencias.append(
                f"Valores de '{COL_MES}' con formato inesperado (se ignoran): "
                f"{', '.join(raros)}"
            )

        update_status(
            f"📅 Comparando {anio_anterior} (arriba) vs {anio_actual} (abajo) — "
            f"{len(meses_por_anio.get(anio_anterior, []))} y "
            f"{len(meses_por_anio.get(anio_actual, []))} mes(es) con datos"
        )

        # --- One cache, four pivots -----------------------------------
        # All four tables read the same range, so they share a single
        # PivotCache: Excel stores the data once instead of four times, and one
        # 'Actualizar' refreshes the lot.
        direccion = (
            f"'{ws_base.Name}'!R{fila_encabezado}C1:R{fila_fin}C{len(encabezados)}"
        )
        cache = wb.PivotCaches().Create(SourceType=XL_DATABASE, SourceData=direccion)

        ocupados = _nombres_de_tablas(wb)
        hojas, tablas = [], []
        anios = [anio_anterior, anio_actual]     # last year on top, current below

        for nombre_hoja, encabezado_hoja, campo_fila, prefijo in COMPARATIVOS:
            update_status(f"🧮 Construyendo '{nombre_hoja}'...")
            _ws, info, avisos = _construir_hoja(
                wb, cache, nombre_hoja, encabezado_hoja, campo_fila, prefijo,
                anios, meses_por_anio, ocupados, update_status,
            )
            hojas.append(nombre_hoja)
            tablas.extend(info)
            advertencias.extend(avisos)

        update_status("💾 Guardando...")
        wb.Save()

        return {
            "ruta":            ruta_final,
            "anio_actual":     anio_actual,
            "anio_anterior":   anio_anterior,
            "fila_encabezado": fila_encabezado,
            "filas_base":      n_filas,
            "hojas":           hojas,
            "tablas":          tablas,
            "meses_por_anio":  {a: list(m) for a, m in sorted(meses_por_anio.items())},
            "advertencias":    advertencias,
        }

    finally:
        cerrar_excel(app, wb, pythoncom)


# ----------------------------------------------------------------------
# Standalone run (no GUI, no controller):  python Comparativos.py
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=== Prueba de tablas dinámicas comparativas (standalone) ===")
    ruta = input("Ruta del archivo de Reservas (.xlsx): ").strip().strip('"')
    mes_txt = input("Mes de referencia [Enter = hoy, o p.ej. M07-26]: ").strip() or None

    try:
        resumen = construir_comparativos(ruta, mes=mes_txt, callback_status=print)
        print("\n🎉 LISTO 🎉")
        for clave, valor in resumen.items():
            print(f"  {clave:<18}: {valor}")
    except Exception as e:
        print(f"❌ Error construyendo los comparativos: {e}")
        raise
