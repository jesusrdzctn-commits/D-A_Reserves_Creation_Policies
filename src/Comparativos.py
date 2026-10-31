"""RDA Pólizas D&A — Comparativos: builds 'Comparativo BU Mes' and 'Comparativo Agrupador Mes', each with TWO REAL Excel
PivotTables stacked (top: last year's months, bottom: current year's): Σ Importe by BU / Agrupador x Mes, grand totals on.
Uses win32com because openpyxl cannot CREATE a pivot cache; it runs LAST and OPTIONALLY on the file openpyxl already saved
(never the other way round: an openpyxl round-trip loses a fresh pivot cache). No Excel/pywin32 -> reported as skipped.
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

# Excel COM constants as plain numbers: win32com.client.constants breaks inside a PyInstaller .exe.
XL_UP           = -4162
XL_DATABASE     = 1       # xlDatabase          (source is a worksheet range)
XL_ROW_FIELD    = 1
XL_COLUMN_FIELD = 2
XL_SUM          = -4157
XL_ASCENDING    = 1
XL_TABULAR_ROW  = 1       # xlTabularRow        (row field shown in its own column)


# --- Sheets and source columns ---
HOJA_BASE = "Base"

HOJA_COMPARATIVO_BU        = "Comparativo BU Mes"
HOJA_COMPARATIVO_AGRUPADOR = "Comparativo Agrupador Mes"

# Headers as they read in 'Base', matched by NORMALISED name ('Importe ' ships with a trailing space).
COL_MES       = "Mes"
COL_BU        = "BU"
COL_AGRUPADOR = "Agrupador"
COL_IMPORTE   = "Importe"

# The 'Base' header row is detected by 'Mes' + 'Agrupador' (padding rows above it have changed); this is only the fallback.
FILA_ENCABEZADO_BASE_DEFAULT = 7

# How far to look when hunting for the header row / reading the header itself.
MAX_FILAS_BUSQUEDA   = 30
MAX_COLUMNAS_LECTURA = 40

# Upper bound on the source range, so a corrupt sheet cannot build a pivot cache over a million rows.
MAX_FILAS = 200_000

# The two sheets to build: (sheet name, heading, row field, pivot name prefix).
COMPARATIVOS = [
    (HOJA_COMPARATIVO_BU,        "Comparativo BU vs Mes",        COL_BU,        "ptComparativoBU"),
    (HOJA_COMPARATIVO_AGRUPADOR, "Comparativo Agrupador vs Mes", COL_AGRUPADOR, "ptComparativoAgrupador"),
]


# --- Layout ---
FILA_ENCABEZADO_HOJA = 1   # sheet heading
FILA_PRIMERA_TABLA   = 4   # first pivot lands here; its banner goes one row above
FILAS_ENTRE_TABLAS   = 3   # blank rows between one pivot's last row and the next banner

ANCHO_COLUMNA_A = 31.33    # same as the 'Concepto' column of 'Variaciones'

# Caption of the value field: must differ from the source field name ('Importe '), or Excel rejects it.
NOMBRE_CAMPO_VALOR = "Suma de Importe"

# Banner above each pivot. Same wording as the 'Variaciones' sheet.
TITULO_TABLA = "Creaciones {anio}"


# --- Styles — same palette as the 'Variaciones' sheet ---
GRIS_CREACIONES = "BFBFBF"   # banner above each pivot
AZUL_TITULO     = "DCE6F2"   # sheet heading

FUENTE = "Calibri"
TAM    = 10

CONTABLE_2 = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'

# Built-in PivotTable style (None = Excel's default); an unknown name is ignored: cosmetics must never kill a run.
ESTILO_TABLA = "PivotStyleMedium9"


def _color_excel(rgb_hex):
    """'DCE6F2' -> the integer Interior.Color expects: Excel stores colours as BGR, not RGB."""
    r = int(rgb_hex[0:2], 16)
    g = int(rgb_hex[2:4], 16)
    b = int(rgb_hex[4:6], 16)
    return r + (g << 8) + (b << 16)


# --- Reading the 'Base' sheet through COM ---
def _bloque(ws, fila1, col1, fila2, col2):
    """Read a rectangle in ONE COM call as a list of rows (1x1 and 1xN ranges come back collapsed, so shapes are normalised)."""
    valores = ws.Range(ws.Cells(fila1, col1), ws.Cells(fila2, col2)).Value

    if valores is None:
        return [[None] * (col2 - col1 + 1)]
    if not isinstance(valores, tuple):          # single cell
        return [[valores]]
    if valores and not isinstance(valores[0], tuple):   # single row
        return [list(valores)]
    return [list(fila) for fila in valores]


def _localizar_fila_encabezado(ws):
    """1-based header row of 'Base': scans the top of the sheet for 'Mes' + 'Agrupador' (same trick as the openpyxl side)."""
    buscados = {norm_texto(COL_MES), norm_texto(COL_AGRUPADOR)}
    filas = _bloque(ws, 1, 1, MAX_FILAS_BUSQUEDA, MAX_COLUMNAS_LECTURA)

    for desplazamiento, fila in enumerate(filas):
        if buscados <= {norm_texto(v) for v in fila}:
            return desplazamiento + 1
    return FILA_ENCABEZADO_BASE_DEFAULT


def _encabezados(ws, fila_encabezado):
    """CONTIGUOUS header names of 'Base', left to right, stopping at the first empty one: the real file has a data column
    with NO header, and a pivot cache refuses a range with a blank header."""
    fila = _bloque(ws, fila_encabezado, 1, fila_encabezado, MAX_COLUMNAS_LECTURA)[0]

    encabezados = []
    for valor in fila:
        if valor is None or str(valor).strip() == "":
            break
        encabezados.append(str(valor))
    return encabezados


def _indice_columna(encabezados, nombre):
    """1-based position of a column in `encabezados` (normalised name); fails loud: a guessed index builds a pivot over the wrong column."""
    objetivo = norm_texto(nombre)
    for i, encabezado in enumerate(encabezados, start=1):
        if norm_texto(encabezado) == objetivo:
            return i
    raise ValueError(
        f"No encontré la columna '{nombre}' en la pestaña '{HOJA_BASE}'.\n\n"
        f"Columnas encontradas: {encabezados}"
    )


def _ultima_fila(ws, fila_inicio, columnas_piloto):
    """Last row of 'Base' with data: walks up from End(xlUp) over three pilot columns, keeping the highest (a blank 'BU' must not chop the table)."""
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
    """{año: [month keys, sorted]} from the 'Mes' column; keys that do not parse land under None, are never shown in a pivot, and are reported."""
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


# --- Building one PivotTable ---
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
    """Leave only `meses_visibles` showing on 'Mes'. Show the wanted items FIRST, hide the rest after (Excel refuses a field
    with zero visible items), and drive PivotItem OBJECTS, not names (a name lookup can be misread)."""
    items = list(campo.PivotItems())
    visibles = set(meses_visibles)

    for item in items:
        if str(item.Name) in visibles and not item.Visible:
            item.Visible = True
    for item in items:
        if str(item.Name) not in visibles and item.Visible:
            item.Visible = False


def _crear_tabla(cache, ws, fila, nombre_tabla, campo_fila, meses_visibles):
    """Create one PivotTable at row `fila`, column A, and return it. ManualUpdate is on for the whole build:
    otherwise Excel recalculates after every property touched, the difference between a blink and several seconds."""
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

        # 'Mes' is text, so an A-Z sort of ONE year's keys is chronological (M01 < ... < M12); mixing years would break it.
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
        # Always hand the pivot back in auto-update mode: a workbook saved with ManualUpdate on opens with a pivot that never refreshes.
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
    """Drop the sheet if present and recreate it in the SAME tab position, so re-runs do not shuffle the tabs."""
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
    """Build one comparison sheet: heading, then one banner + one pivot per year. Returns (ws, [info de cada tabla], [advertencias])."""
    ws = _preparar_hoja(wb, nombre_hoja)
    _banner(ws, FILA_ENCABEZADO_HOJA, encabezado_hoja, AZUL_TITULO, tam=TAM + 1)

    tablas, advertencias = [], []
    fila = FILA_PRIMERA_TABLA

    for anio in anios:
        meses = meses_por_anio.get(anio, [])
        _banner(ws, fila - 1, TITULO_TABLA.format(anio=anio), GRIS_CREACIONES)

        if not meses:
            # No data for that year: write a note, not a pivot (a pivot cannot have every column hidden, and the
            # alternative would show the OTHER year's months under this banner).
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

        # Bottom-right cell = Grand Total of Grand Totals: a cheap control figure that must match the year's 'Importe'.
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


# --- Main entry point ---
def construir_comparativos(
    ruta_libro,
    ruta_salida=None,
    mes=None,
    sobrescribir_original=True,
    callback_status=None,
):
    """Build the 'Comparativo BU Mes' and 'Comparativo Agrupador Mes' sheets in the Reservas workbook (normally the COPY
    construir_variaciones() just wrote). `mes` decides which year is 'current'. sobrescribir_original defaults to TRUE here,
    unlike the other modules: copying a copy would split the pivots and the matrices across two files.
    Returns a dict: ruta, anio_actual, anio_anterior, fila_encabezado, filas_base, hojas, tablas, meses_por_anio, advertencias."""
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

        # --- One cache, four pivots ---
        # All four tables read the same range, so they share one PivotCache: data stored once, one 'Actualizar' refreshes all.
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


# --- Standalone run: python Comparativos.py ---
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
