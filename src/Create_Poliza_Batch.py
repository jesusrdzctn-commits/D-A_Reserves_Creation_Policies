"""RDA Pólizas D&A — builds ONE 'Póliza Carga Batch' ('GL Posting Upload') from the finished 'PÓLIZA SAP' of the channel
layouts (each with its 'Reservas Manuales' and 'Reservas Rentas' tabs): one debit block, then one credit block.
Rows without an amount (0, blank, '-') are left out. Fills the orange fields and the credit GL Account, adds the
'Validación - Cifras Control' sheet (two PivotTables), reads everything back and saves only if it matches.
"""

import math
import os
import re
import shutil
import time
from datetime import datetime
from decimal import Decimal

# Layout helpers are reused from Create_Poliza_SAP so both modules agree on where the section starts and ends.
from Consolidacion import GRUPOS, resolver_grupo
from Create_Poliza_SAP import (
    ENTRADAS,
    XL_PASTE_VALUES,
    localizar_secciones,
    titulo_de,
)
from Comparativos import _campos, _nombre_libre, _nombres_de_tablas
from Reservas_Manuales import _grupo_del_archivo as _canal_del_archivo
from utils import abrir_excel as _abrir_excel
from utils import cerrar_excel as _cerrar_excel
from utils import error_de_excel as _error_de
from utils import es_vacio_excel as _vacio
from utils import hoja_por_nombre as _hoja
from utils import mes_de_cierre, motivo_reserva_omitida
from utils import norm_texto as _norm
from utils import ultimo_dia_mes

# Excel COM constants as plain numbers (win32com.client.constants breaks inside a PyInstaller .exe).
XL_DATABASE, XL_ROW_FIELD, XL_SUM, XL_COMPACT_ROW = 1, 1, -4157, 0

# --- Template layout ---
HOJA_PLANTILLA = "GL Posting Upload"
EXTENSIONES_PLANTILLA = (".xlsx", ".xlsm")

# The header row is searched in the first rows instead of hard-coding row 1.
FILA_ENCABEZADO_MAX = 10

# Headers of the template that this module writes under. Matched on the whole
# normalised text, so 'Profit Center' never grabs 'partner_profit_center'.
DEST_CUENTA, DEST_DEBITO, DEST_CREDITO = "GL Account", "Debit Amount", "Credit Amount"
DEST_CEBE, DEST_SEGMENTO = "Profit Center", "Segment"
DEST_ASIGNACION, DEST_TEXTO = "Assignment", "Line item Text"

# Orange header fields: the same value on every row of both blocks.
DEST_SECUENCIA, DEST_SOCIEDAD = "Document Sequence #", "Posting Company Code"
DEST_FECHA_CONTABLE, DEST_FECHA_DOC = "Posting Date", "Document Date"
DEST_PERIODO, DEST_MONEDA = "Period", "Currency"
DEST_TIPO_DOC, DEST_CABECERA, DEST_REFERENCIA = "Document Type", "Document Header", "Reference"

SECUENCIA, SOCIEDAD, MONEDA = 1, "MX21", "MXN"
TIPO_DOC, CABECERA = "SA", "CREACION RESERVA"
# Dates are written as TEXT in this format, so Excel never re-parses them by the machine's locale.
FORMATO_FECHA = "%d.%m.%Y"

# GL Account of the credit block, only for cells left empty: 'PROMOCIONAL' in the Line item Text -> first account.
PALABRA_PROMOCIONAL = "PROMOCIONAL"
CUENTA_PROMOCIONAL, CUENTA_RESTO = "0001011003", "0002031130"

ENCABEZADOS_DESTINO = (
    DEST_CUENTA, DEST_DEBITO, DEST_CREDITO, DEST_CEBE, DEST_SEGMENTO, DEST_ASIGNACION, DEST_TEXTO,
    DEST_SECUENCIA, DEST_SOCIEDAD, DEST_FECHA_CONTABLE, DEST_FECHA_DOC, DEST_PERIODO,
    DEST_MONEDA, DEST_TIPO_DOC, DEST_CABECERA, DEST_REFERENCIA,
)

# Key (in Create_Poliza_SAP.COLUMNAS_POLIZA) of the layout column feeding each header; letters are never repeated here.
CLAVE_CUENTA = "cuenta"
CLAVE_CONCEPTO = "concepto"
# SAP takes two decimals only: Debit and Credit Amount come from REDONDEO (= ROUND(MXN, 2)); MXN is only compared.
CLAVE_MONTO = "redondeo"
CLAVE_COMPARACION = "mxn"
# RESERVA ($) decides which rows are left out of the policy (key of Create_Poliza_SAP.ENTRADAS).
CLAVE_RESERVA = "reserva"
TRANSFERENCIAS = (
    {"origen": CLAVE_CUENTA,   "destino": DEST_CUENTA,
     "en_debito": True,  "en_credito": False},
    {"origen": "cebe",         "destino": DEST_CEBE,
     "en_debito": True,  "en_credito": True},
    {"origen": "segmento",     "destino": DEST_SEGMENTO,
     "en_debito": True,  "en_credito": True},
    {"origen": "asignacion",   "destino": DEST_ASIGNACION,
     "en_debito": True,  "en_credito": True},
    {"origen": CLAVE_MONTO,    "destino": DEST_DEBITO,
     "en_debito": True,  "en_credito": False},
    {"origen": CLAVE_MONTO,    "destino": DEST_CREDITO,
     "en_debito": False, "en_credito": True},
    {"origen": CLAVE_CONCEPTO, "destino": DEST_TEXTO,
     "en_debito": True,  "en_credito": True},
)

# Standard SAP field lengths; only used to WARN (SAP truncates or rejects longer values).
LONGITUD_MAX_SAP = {
    DEST_CUENTA: 10,
    DEST_CEBE: 10,
    DEST_SEGMENTO: 10,
    DEST_ASIGNACION: 18,
    DEST_TEXTO: 50,
    DEST_REFERENCIA: 16,
    DEST_CABECERA: 25,
}

# --- Behaviour flags ---
# Clear the template's target columns (below the header) first so an old run's leftovers cannot survive.
LIMPIAR_ANTES = True

# How many offending cells / rows to list per problem.
MAX_DETALLE = 10

# Excel's clipboard and the file system fail transiently (another app on the clipboard, antivirus).
INTENTOS = 4
ESPERA_SEG = 0.6

# Tolerance when comparing the debit / credit / layout totals (cents).
TOLERANCIA_TOTAL = 0.005

# SAP takes at most two decimals in an amount.
DECIMALES_SAP = 2

# --- Control sheet: two PivotTables (GL Account, Assignment) x Σ Debit / Σ Credit, as in the reference file ---
HOJA_CIFRAS = "Validación - Cifras Control"
TITULO_CIFRAS = "CREACIONES OT"
FILA_TITULO_CIFRAS, FILA_PIVOTES_CIFRAS = 4, 5
# (key, PivotTable name, row field, caption of the row labels, first column)
PIVOTES_CIFRAS = (
    ("cuenta",     "ptCifrasCuenta",     DEST_CUENTA,     "No Cuenta",  1),
    ("asignacion", "ptCifrasAsignacion", DEST_ASIGNACION, "Assignment", 6),
)
CAPTION_DEBITO, CAPTION_CREDITO = "Sum of Debit Amount", "Sum of Credit Amount"
FORMATO_CIFRAS = "#,##0.00"
ESTILO_PIVOTE = "PivotStyleLight16"

# Rows above a layout header searched for the channel name (title), and how many columns of them.
COLUMNAS_TITULO = 30


# --- Small helpers ---
def _ref_hoja(nombre):
    """Sheet name as a formula / pivot reference ('' doubles any quote)."""
    return "'" + str(nombre).replace("'", "''") + "'"


def _misma_ruta(a, b):
    """True when two paths point to the same file (Windows: case-insensitive)."""
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def _nombre_salida(ruta_plantilla):
    """Output name from the template's: everything from 'template' on is dropped and a date stamp ALWAYS appended
    ('...OT 26 template_IA_v2.xlsx' -> '...OT 26_20261027.xlsx'), so it can never equal the template's name."""
    base, ext = os.path.splitext(os.path.basename(ruta_plantilla))
    limpio = re.sub(r"[\s_\-]*template.*$", "", base, flags=re.IGNORECASE).strip(" _-")
    sello = datetime.today().strftime("%Y%m%d")
    return f"{limpio or base}_{sello}{ext}"


def referencia_de_usuario(ruta_perfil=None):
    """User number after 'Users\\' in the profile folder (C:\\Users\\81113276 -> '81113276')."""
    partes = [p for p in re.split(r"[\\/]+", ruta_perfil or os.path.expanduser("~")) if p]
    for i, parte in enumerate(partes[:-1]):
        if parte.lower() == "users":
            return partes[i + 1]
    raise ValueError(
        f"No pude sacar el número de usuario de la carpeta '{ruta_perfil or os.path.expanduser('~')}': "
        f"esperaba algo como C:\\Users\\<número>."
    )


def fecha_contable(mes=None):
    """Posting Date: last day of the closing month as text ('31.10.2026' for 'M10-26'; None = current month)."""
    return ultimo_dia_mes(mes_de_cierre(mes)).strftime(FORMATO_FECHA)


def _campos_naranjas(mes, referencia):
    """(header, value, as_text) of every orange field; dates and Period come from the closing month."""
    fecha = fecha_contable(mes)
    return (
        (DEST_SECUENCIA, SECUENCIA, False),
        (DEST_SOCIEDAD, SOCIEDAD, True),
        (DEST_FECHA_CONTABLE, fecha, True),
        (DEST_FECHA_DOC, fecha, True),
        (DEST_PERIODO, str(mes_de_cierre(mes).month), True),
        (DEST_MONEDA, MONEDA, True),
        (DEST_TIPO_DOC, TIPO_DOC, True),
        (DEST_CABECERA, CABECERA, True),
        (DEST_REFERENCIA, referencia, True),
    )


def _cuenta_contrapartida(texto):
    """'PROMOCIONAL' anywhere in the text (ignoring case and accents) -> promotional account, else the general one."""
    return CUENTA_PROMOCIONAL if PALABRA_PROMOCIONAL in _norm(texto) else CUENTA_RESTO


def _claves_origen():
    """Layout columns that feed the batch, in order, without repeats."""
    vistas = []
    for t in TRANSFERENCIAS:
        if t["origen"] not in vistas:
            vistas.append(t["origen"])
    return vistas


def _leer_rango(rango, n):
    """Value2 of a ONE-column range as a flat list (a single cell returns a scalar: the COM trap for one-line policies)."""
    valor = rango.Value2
    if n == 1:
        return [valor]
    return [fila[0] for fila in valor]


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
    """Copy -> Paste Special -> Values, retrying while the clipboard is busy.
    Always clears the marching ants so the user's clipboard is not left holding part of the workbook."""
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


# --- Layouts (source) ---
def _titulo(clave):
    """Header text of a layout column: the eight 'PÓLIZA SAP' ones, or an input one such as 'RESERVA ($)'."""
    return ENTRADAS[clave] if clave in ENTRADAS else titulo_de(clave)


def _canal_de_titulo(s):
    """Index of the channel named by the text above a tab's header (None = no channel word there); two channels is an error."""
    fila_enc = s["fila_encabezado"]
    if fila_enc <= 1:
        return None
    ws = s["hoja"]
    bloque = ws.Range(ws.Cells(1, 1), ws.Cells(fila_enc - 1, COLUMNAS_TITULO)).Value2
    if not isinstance(bloque, tuple):
        bloque = ((bloque,),)
    hallados = {resolver_grupo(v) for fila in bloque for v in fila if isinstance(v, str)} - {None}
    if len(hallados) > 1:
        raise ValueError(
            f"El título de la pestaña '{s['nombre']}' menciona más de un canal "
            f"({sorted(GRUPOS[i]['canal'] for i in hallados)}); no sé cuál es."
        )
    return next(iter(hallados)) if hallados else None


def _leer_layout(app, ruta, abiertos):
    """Open ONE processed layout read-only (it joins `abiertos` right away, so the caller always closes it) and
    return {'ruta', 'nombre', 'indice', 'canal', 'secciones', 'avisos'}. Its channel comes from the tab titles
    cross-checked with the file name (the same rule as step 1); tabs without rows are skipped."""
    nombre = os.path.basename(ruta)
    wb = _abrir_libro(app, os.path.abspath(ruta), solo_lectura=True)
    abiertos.append(wb)
    try:
        secciones, avisos = localizar_secciones(wb)
        secciones = [s for s in secciones if s["filas"] > 0]
        if not secciones:
            raise ValueError("no tiene renglones con datos en ninguna pestaña.")
        titulos = [(s["nombre"], _canal_de_titulo(s)) for s in secciones]
    except ValueError as e:
        raise ValueError(f"En '{nombre}': {e}") from e

    indice, _ = _canal_del_archivo(titulos, nombre, GRUPOS)
    canal = GRUPOS[indice]["canal"]
    for s in secciones:
        s["canal"] = canal
        s["rotulo"] = f"{canal}/{s['nombre']}"
    return {
        "ruta": os.path.abspath(ruta), "nombre": nombre, "indice": indice, "canal": canal,
        "secciones": secciones, "avisos": [f"[{canal}] {a}" for a in avisos],
    }


def _ordenar_layouts(layouts):
    """(layouts in the channel order of GRUPOS, names of the channels that were not given).
    One file per channel: two files of the same channel are an error, since adding both would double its money."""
    por_canal = {}
    for lay in layouts:
        por_canal.setdefault(lay["indice"], []).append(lay)

    repetidos = [f"{GRUPOS[i]['canal']}: " + ", ".join(f"'{l['nombre']}'" for l in ls)
                 for i, ls in sorted(por_canal.items()) if len(ls) > 1]
    if repetidos:
        raise ValueError(
            "Hay más de un archivo para el mismo canal (sólo se admite uno por canal):\n\n"
            + "\n".join(f"     • {r}" for r in repetidos)
        )
    faltantes = [g["canal"] for i, g in enumerate(GRUPOS) if i not in por_canal]
    return [por_canal[i][0] for i in sorted(por_canal)], faltantes


def _leer_origen(secciones):
    """{key: {'valores', 'celdas', 'titulo'}} for every copied layout column, the amount's comparison column and RESERVA ($).
    Rows of all tabs of all layouts follow one another in section order; 'celdas' holds each address ('HD/Reservas Rentas!K8')."""
    claves = _claves_origen()
    for extra in (CLAVE_COMPARACION, CLAVE_RESERVA):
        if extra not in claves:
            claves.append(extra)

    datos = {}
    for clave in claves:
        valores, celdas = [], []
        for s in secciones:
            col = s["columnas"][clave]
            rango = s["hoja"].Range(
                f"{col['letra']}{s['fila_inicio']}:{col['letra']}{s['fila_fin']}"
            )
            valores += _leer_rango(rango, s["filas"])
            celdas += [f"{s['rotulo']}!{col['letra']}{s['fila_inicio'] + i}"
                       for i in range(s["filas"])]
        datos[clave] = {"valores": valores, "celdas": celdas, "titulo": _titulo(clave)}
    return datos


def _posiciones(secciones):
    """(section index, row offset inside it) of every row read, in the same order as `_leer_origen`."""
    return [(i, k) for i, s in enumerate(secciones) for k in range(s["filas"])]


def _filtrar(datos, conservar):
    """The same dict keeping only the rows flagged in `conservar`."""
    return {
        clave: {
            "valores": [v for v, c in zip(d["valores"], conservar) if c],
            "celdas":  [x for x, c in zip(d["celdas"], conservar) if c],
            "titulo":  d["titulo"],
        }
        for clave, d in datos.items()
    }


def _componer(primera, segunda):
    """Mask over the original rows from a first mask and a second one made over the rows that survived the first."""
    final = list(primera)
    vivas = [i for i, c in enumerate(primera) if c]
    for i, c in zip(vivas, segunda):
        final[i] = c
    return final


def _omitir_por_reserva(datos):
    """(conservar, omitidas): rows whose RESERVA ($) holds no amount (0, blank or '-') are left out of the policy.
    An Excel error or any other text is NOT left out: it stays so that the validation blocks it."""
    omitidas = {"cero": [], "blanco": [], "guion": []}
    conservar = []
    reserva = datos[CLAVE_RESERVA]
    for celda, valor in zip(reserva["celdas"], reserva["valores"]):
        motivo = motivo_reserva_omitida(valor)
        conservar.append(motivo is None)
        if motivo:
            omitidas[motivo].append(celda)
    return conservar, omitidas


def _omitir_redondeo_cero(datos):
    """(conservar, celdas): rows that survived but whose REDONDEO is 0 (RESERVA below half a cent). SAP would
    take a line of 0.00, so it is left out too and reported. Call it only after the amounts were validated."""
    monto = datos[CLAVE_MONTO]
    conservar = [v != 0 for v in monto["valores"]]
    celdas = [c for c, k in zip(monto["celdas"], conservar) if not k]
    return conservar, celdas


def _enumerar(celdas):
    """First MAX_DETALLE cells, plus an ellipsis when there are more."""
    return ", ".join(celdas[:MAX_DETALLE]) + ("..." if len(celdas) > MAX_DETALLE else "")


def _mas_de_dos_decimales(valor):
    """True when a finite amount needs more than the two decimals SAP accepts (judged on its shortest text form)."""
    return Decimal(repr(float(valor))).as_tuple().exponent < -DECIMALES_SAP


def _problemas_origen(datos):
    """Blocking problems of the rows that WILL go up: error cells, empty cells, non-numeric or non-finite amounts,
    more than two decimals in the amount. Returns [{'titulo', 'tipo', 'total', 'celdas'}]."""
    problemas = []
    for clave in _claves_origen() + [CLAVE_RESERVA]:
        d = datos[clave]
        es_monto = clave in (CLAVE_MONTO, CLAVE_RESERVA)
        grupos = {"error": [], "vacío": [], "no numérico": [], "tipo inesperado": [],
                  "más de 2 decimales": []}

        for celda, valor in zip(d["celdas"], d["valores"]):
            if _error_de(valor):
                grupos["error"].append(f"{celda} {_error_de(valor)}")
            elif _vacio(valor):
                grupos["vacío"].append(celda)
            elif isinstance(valor, bool):
                grupos["tipo inesperado"].append(celda)
            elif es_monto and (not isinstance(valor, (int, float)) or not math.isfinite(valor)):
                grupos["no numérico"].append(celda)
            elif clave == CLAVE_MONTO and _mas_de_dos_decimales(valor):
                grupos["más de 2 decimales"].append(celda)
            elif not isinstance(valor, (str, int, float)):
                grupos["tipo inesperado"].append(celda)

        for tipo, celdas in grupos.items():
            if celdas:
                problemas.append({
                    "titulo": d["titulo"], "tipo": tipo,
                    "total": len(celdas), "celdas": celdas[:MAX_DETALLE],
                })
    return problemas


def _advertencias_origen(datos):
    """Non-blocking warnings about the rows that go up: negative amounts, REDONDEO vs MXN, numeric GL accounts,
    values over the SAP field lengths. Only call it with problem-free data."""
    advertencias = []
    monto = datos[CLAVE_MONTO]
    negativos = [c for c, v in zip(monto["celdas"], monto["valores"]) if v < 0]
    if negativos:
        advertencias.append(
            f"{len(negativos)} renglón(es) con {monto['titulo']} negativo (irían con "
            f"signo negativo en Debit y Credit Amount): " + _enumerar(negativos)
        )

    comparado = datos[CLAVE_COMPARACION]
    difieren = [
        c for c, a, b in zip(monto["celdas"], monto["valores"], comparado["valores"])
        if isinstance(b, (int, float)) and abs(a - b) > TOLERANCIA_TOTAL
    ]
    if difieren:
        advertencias.append(
            f"{len(difieren)} renglón(es) donde '{monto['titulo']}' (el que se manda a "
            f"SAP) y '{comparado['titulo']}' no coinciden: " + _enumerar(difieren)
        )

    cuentas = datos[CLAVE_CUENTA]
    numericas = [c for c, v in zip(cuentas["celdas"], cuentas["valores"])
                 if not isinstance(v, str)]
    if numericas:
        advertencias.append(
            f"{len(numericas)} {cuentas['titulo']} vienen como NÚMERO y no como texto "
            f"(¿se perdieron ceros a la izquierda?): " + _enumerar(numericas)
        )

    for t in TRANSFERENCIAS:
        limite = LONGITUD_MAX_SAP.get(t["destino"])
        if not limite:
            continue
        d = datos[t["origen"]]
        largos = [c for c, v in zip(d["celdas"], d["valores"])
                  if isinstance(v, str) and len(v) > limite]
        if largos:
            advertencias.append(
                f"{len(largos)} valor(es) de '{t['destino']}' exceden los {limite} "
                f"caracteres que SAP acepta: " + _enumerar(largos)
            )
    return advertencias


def _avisos_omitidas(omitidas, redondeo, suma_omitida):
    """Warnings about the rows left out. Zeros are the expected case (counted in the summary); blank or '-' cells
    and rows lost to rounding are listed cell by cell, because a forgotten amount looks exactly like them."""
    avisos = []
    textos = {
        "blanco": "con RESERVA ($) en blanco",
        "guion":  "con RESERVA ($) en '-'",
    }
    for motivo, texto in textos.items():
        if omitidas[motivo]:
            avisos.append(f"{len(omitidas[motivo])} renglón(es) {texto} se dejaron fuera de la póliza: "
                          + _enumerar(omitidas[motivo]))
    if redondeo:
        avisos.append(f"{len(redondeo)} renglón(es) con RESERVA ($) distinta de 0 pero REDONDEO = 0.00 "
                      f"se dejaron fuera de la póliza: " + _enumerar(redondeo))
    if abs(suma_omitida) > TOLERANCIA_TOTAL:
        avisos.append(f"Los renglones dejados fuera suman {suma_omitida:,.2f} en REDONDEO (debían sumar 0): "
                      f"revisa que no se esté perdiendo dinero.")
    return avisos


def _texto_problemas(problemas, layouts):
    """Readable message for the blocking problems found in the layouts."""
    total = sum(p["total"] for p in problemas)
    lineas = []
    for p in problemas:
        resto = p["total"] - len(p["celdas"])
        lineas.append(
            f"     • {p['titulo']}: {p['total']} celda(s) con {p['tipo']}\n       "
            + ", ".join(p["celdas"])
            + (f", ...y {resto} más" if resto > 0 else "")
        )
    return (
        f"No se generó la Póliza Carga Batch: los Layouts todavía tienen {total} "
        f"celda(s) con problemas que SAP rechazaría.\n\n"
        + "\n".join(lineas)
        + "\n\nCorrígelas (normalmente es un valor que falta en la pestaña "
          "'Catalogo') y vuelve a correr el paso 3.\n\n"
          "Layouts revisados:\n" + "\n".join(f"     • {lay['nombre']}" for lay in layouts)
    )


# --- Template (destination) ---
def _hoja_plantilla(wb):
    """The template's sheet by name; a one-sheet workbook is accepted whatever it is called."""
    try:
        return _hoja(wb, HOJA_PLANTILLA)
    except ValueError:
        if int(wb.Worksheets.Count) == 1:
            return wb.Worksheets(1)
        raise


def _localizar_destino(ws):
    """(fila_encabezado, {header: column index}) of the template, found by NAME so an inserted column
    still lands every value under the right header, or stops with a readable message."""
    requeridos = {_norm(h): h for h in ENCABEZADOS_DESTINO}

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


# --- Paste + verification ---
def _bloques(fila_encabezado, n):
    """Row spans of the two blocks, the credit one starting right under the debit one."""
    d_ini = fila_encabezado + 1
    d_fin = d_ini + n - 1
    c_ini = d_fin + 1
    c_fin = c_ini + n - 1
    return {"debito": (d_ini, d_fin), "credito": (c_ini, c_fin)}


def _plan_pegado(secciones, posiciones, conservar):
    """Runs of consecutive KEPT rows as (section, first offset, count, offset inside each block). Rows left out are
    never pasted, so the two blocks come out compact with no delete step that could shift the template's rows."""
    plan = []
    destino = 0
    previa = None
    for (i, k), c in zip(posiciones, conservar):
        if not c:
            previa = None
            continue
        if previa is not None and previa[0] == i and previa[1] == k - 1:
            plan[-1][2] += 1
        else:
            plan.append([i, k, 1, destino])
        previa = (i, k)
        destino += 1
    return [(secciones[i], k, cuenta, dest) for i, k, cuenta, dest in plan]


def _pegar_todo(app, ws_destino, columnas, bloques, plan):
    """Paste every (column, block) of TRANSFERENCIAS as values, run by run, each at its offset inside the block."""
    # PasteSpecial wants the destination sheet to be the active one.
    try:
        ws_destino.Parent.Activate()
        ws_destino.Activate()
    except Exception:
        pass

    for t in TRANSFERENCIAS:
        for bloque, bandera in (("debito", "en_debito"), ("credito", "en_credito")):
            if not t[bandera]:
                continue
            r1, _ = bloques[bloque]
            for s, primera, cuenta, dest in plan:
                letra = s["columnas"][t["origen"]]["letra"]
                desde_origen = s["fila_inicio"] + primera
                origen = s["hoja"].Range(f"{letra}{desde_origen}:{letra}{desde_origen + cuenta - 1}")
                desde = r1 + dest
                _pegar_valores(
                    app, origen,
                    _rango_destino(ws_destino, columnas[t["destino"]], desde, desde + cuenta - 1),
                )


def _fijar_formato(rango, formato):
    """Apply `formato` unless the range already has it. False when Excel refuses it (typically a
    protected sheet that allows typing in the cells but not formatting them)."""
    try:
        if rango.NumberFormat == formato:
            return True
        rango.NumberFormat = formato
        return True
    except Exception:  # com_error
        return False


def _con_apostrofe(valor):
    """Text with a leading apostrophe: a scalar, or a column tuple of 1-tuples."""
    if isinstance(valor, tuple):
        return tuple(("'" + fila[0],) for fila in valor)
    return "'" + valor


def _escribir(ws, rango, titulo, valor, como_texto, avisos):
    """Write `valor` under the wanted format. If Excel refuses the format, text goes in with a leading
    apostrophe (what was typed by hand): it stays text and the read-back checks what was really stored."""
    if not _fijar_formato(rango, "@" if como_texto else "General") and como_texto:
        valor = _con_apostrofe(valor)
        aviso = (f"Excel no dejó dar formato Texto a '{titulo}' (¿hoja protegida?): "
                 f"se escribió con apóstrofo y la verificación lo confirmó.")
        if avisos is not None and aviso not in avisos:
            avisos.append(aviso)
    try:
        rango.Value2 = valor
    except Exception as e:  # com_error
        try:
            protegida = bool(ws.ProtectContents)
        except Exception:
            protegida = False
        causa = (f"La hoja '{ws.Name}' está PROTEGIDA: quita la protección (o desbloquea esas celdas) "
                 f"y vuelve a intentar." if protegida else "¿La plantilla está en solo lectura o protegida?")
        raise RuntimeError(
            f"Excel no me dejó escribir en la columna '{titulo}' de la plantilla:\n\n{e}\n\n{causa}"
        ) from e


def _rellenar_naranjas(ws, columnas, desde, hasta, campos, avisos=None):
    """Write each orange field down every row. Text fields get the Text format first, so Excel keeps
    them as typed ('81113276', '31.10.2026') instead of re-parsing them."""
    for destino, valor, como_texto in campos:
        rango = _rango_destino(ws, columnas[destino], desde, hasta)
        _escribir(ws, rango, destino, valor, como_texto, avisos)


def _completar_cuentas(ws, columnas, desde, hasta, avisos=None):
    """Fill ONLY the empty GL Account cells (the credit block) by the PROMOCIONAL rule on Line item Text.
    Returns {account: count}. Text format first: the leading zeros of '0001011003' must survive."""
    n = hasta - desde + 1
    cuentas = _leer_rango(_rango_destino(ws, columnas[DEST_CUENTA], desde, hasta), n)
    textos = _leer_rango(_rango_destino(ws, columnas[DEST_TEXTO], desde, hasta), n)
    puestas = {}
    i = 0
    while i < n:
        if not _vacio(cuentas[i]):
            i += 1
            continue
        j = i
        while j + 1 < n and _vacio(cuentas[j + 1]):
            j += 1
        valores = [_cuenta_contrapartida(t) for t in textos[i:j + 1]]
        rango = _rango_destino(ws, columnas[DEST_CUENTA], desde + i, desde + j)
        _escribir(ws, rango, DEST_CUENTA, tuple((v,) for v in valores), True, avisos)
        for v in valores:
            puestas[v] = puestas.get(v, 0) + 1
        i = j + 1
    return puestas


def _comparar(destino, bloque, fila_ini, esperado, real):
    """Readable differences between what a column should hold and what the template holds."""
    malos = [(fila_ini + i, e, r) for i, (e, r) in enumerate(zip(esperado, real)) if e != r]
    dif = [f"{destino} fila {f} ({bloque}): esperaba {e!r} y quedó {r!r}"
           for f, e, r in malos[:MAX_DETALLE]]
    if len(malos) > MAX_DETALLE:
        dif.append(f"{destino} ({bloque}): ...y {len(malos) - MAX_DETALLE} más")
    return dif


def _verificar(ws_destino, columnas, bloques, datos, n):
    """Read the template back: returns (diferencias, sumas, total_origen); empty diferencias = identical.
    Checks every pasted column cell by cell (catches lost leading zeros), the credit GL Account rule, that the
    other cells stay empty, and that debit total = credit total = layout total (all tabs)."""
    diferencias = []
    sumas = {}
    cuentas_credito = [_cuenta_contrapartida(v) for v in datos[CLAVE_CONCEPTO]["valores"]]

    for t in TRANSFERENCIAS:
        col_idx = columnas[t["destino"]]
        for bloque, bandera in (("debito", "en_debito"), ("credito", "en_credito")):
            r1, r2 = bloques[bloque]
            real = _leer_rango(_rango_destino(ws_destino, col_idx, r1, r2), n)

            if t[bandera]:
                diferencias += _comparar(t["destino"], bloque, r1, datos[t["origen"]]["valores"], real)
                if t["destino"] in (DEST_DEBITO, DEST_CREDITO):
                    sumas[t["destino"]] = math.fsum(v for v in real if isinstance(v, (int, float)))
            elif t["destino"] == DEST_CUENTA:
                diferencias += _comparar(t["destino"], bloque, r1, cuentas_credito, real)
            else:
                sucios = [r1 + i for i, v in enumerate(real) if not _vacio(v)]
                if sucios:
                    diferencias.append(
                        f"{t['destino']} ({bloque}) debía quedar vacío y trae datos en las "
                        f"filas {sucios[:MAX_DETALLE]}"
                    )

    total_origen = math.fsum(
        v for v in datos[CLAVE_MONTO]["valores"] if isinstance(v, (int, float))
    )
    for destino, suma in sumas.items():
        if abs(suma - total_origen) > TOLERANCIA_TOTAL:
            diferencias.append(
                f"Σ {destino} = {suma:,.2f} y la Σ MONTO de los renglones incluidos = {total_origen:,.2f}"
            )
    return diferencias, sumas, total_origen


def _verificar_naranjas(ws_destino, columnas, desde, hasta, campos):
    """Every orange field must read back as written on ALL rows (both blocks)."""
    n = hasta - desde + 1
    diferencias = []
    for destino, valor, _ in campos:
        real = _leer_rango(_rango_destino(ws_destino, columnas[destino], desde, hasta), n)
        diferencias += _comparar(destino, "ambos bloques", desde, [valor] * n, real)
    return diferencias


def _avisos_referencia(referencia):
    """Warnings (never blocking) about the user number used as 'Reference'."""
    avisos = []
    if not referencia.isdigit():
        avisos.append(f"La referencia '{referencia}' (carpeta de usuario) no son sólo dígitos: revisa 'Reference'.")
    limite = LONGITUD_MAX_SAP[DEST_REFERENCIA]
    if len(referencia) > limite:
        avisos.append(f"La referencia '{referencia}' excede los {limite} caracteres que SAP acepta.")
    return avisos


# --- Control sheet ---
def _etiqueta_pivote(valor):
    """Row label as a PivotTable shows it, compared case-insensitively (a pivot merges 'abc' and 'ABC')."""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip().casefold()


def _esperado_cifras(datos, cuentas_credito, clave):
    """{label: [debit, credit]} the pivot of `clave` ('cuenta' or 'asignacion') must show, from the rows that go up."""
    monto = datos[CLAVE_MONTO]["valores"]
    asignacion = datos["asignacion"]["valores"]
    cuenta_debito = datos[CLAVE_CUENTA]["valores"]
    esperado = {}
    for i, importe in enumerate(monto):
        for lado, cuenta in ((0, cuenta_debito[i]), (1, cuentas_credito[i])):
            etiqueta = _etiqueta_pivote(cuenta if clave == "cuenta" else asignacion[i])
            esperado.setdefault(etiqueta, [[], []])[lado].append(importe)
    return {k: (math.fsum(d), math.fsum(c)) for k, (d, c) in esperado.items()}


def _leer_pivote(pt):
    """({label: (debit, credit)}, (total debit, total credit)) read back from a PivotTable's body.
    Raises ValueError when it is not the three-column shape built here (label, Σ debit, Σ credit, then a Grand Total row)."""
    bloque = pt.TableRange1.Value2
    if (not isinstance(bloque, tuple) or len(bloque) < 3
            or any(not isinstance(f, tuple) or len(f) != 3 for f in bloque)):
        raise ValueError("no tiene la forma esperada (etiqueta, Σ débito, Σ crédito y Gran Total)")
    if (_norm(bloque[0][1]), _norm(bloque[0][2])) != (_norm(CAPTION_DEBITO), _norm(CAPTION_CREDITO)):
        raise ValueError(f"sus columnas de valores son {bloque[0][1:]!r}, no débito y crédito")

    def num(v):
        return 0.0 if v is None else float(v)

    filas = {}
    for etiqueta, deb, cre in bloque[1:-1]:
        filas[_etiqueta_pivote(etiqueta)] = (num(deb), num(cre))
    return filas, (num(bloque[-1][1]), num(bloque[-1][2]))


def _verificar_pivote(pt, esperado, total_debito, total_credito):
    """Differences between a PivotTable read back and what the template rows say (empty list = identical)."""
    try:
        filas, totales = _leer_pivote(pt)
    except (ValueError, TypeError) as e:
        return [f"la tabla '{pt.Name}' {e}"]

    dif = []
    faltan = sorted(set(esperado) - set(filas))
    sobran = sorted(set(filas) - set(esperado))
    if faltan:
        dif.append(f"'{pt.Name}': faltan etiquetas {faltan[:MAX_DETALLE]}")
    if sobran:
        dif.append(f"'{pt.Name}': sobran etiquetas {sobran[:MAX_DETALLE]}")
    malas = [k for k in esperado if k in filas
             and any(abs(a - b) > TOLERANCIA_TOTAL for a, b in zip(esperado[k], filas[k]))]
    for k in malas[:MAX_DETALLE]:
        dif.append(f"'{pt.Name}' {k!r}: esperaba {esperado[k]!r} y quedó {filas[k]!r}")
    if len(malas) > MAX_DETALLE:
        dif.append(f"'{pt.Name}': ...y {len(malas) - MAX_DETALLE} etiquetas más con importes distintos")
    for nombre, real, esp in (("Gran Total débito", totales[0], total_debito),
                              ("Gran Total crédito", totales[1], total_credito)):
        if abs(real - esp) > TOLERANCIA_TOTAL:
            dif.append(f"'{pt.Name}' {nombre} = {real:,.2f} y la póliza suma {esp:,.2f}")
    return dif


def _crear_pivote(cache, ws, columna, nombre_tabla, campo_fila, rotulo_fila):
    """One PivotTable at row FILA_PIVOTES_CIFRAS: `campo_fila` down the side, Σ Debit / Σ Credit across (compact layout,
    Grand Totals, as in the reference file). ManualUpdate is on while it is built and always handed back off."""
    pt = cache.CreatePivotTable(
        TableDestination=f"{_ref_hoja(ws.Name)}!R{FILA_PIVOTES_CIFRAS}C{columna}", TableName=nombre_tabla
    )
    pt.ManualUpdate = True
    try:
        campos = _campos(pt)
        c_fila = pt.PivotFields(campos[_norm(campo_fila)])
        c_fila.Orientation = XL_ROW_FIELD
        c_fila.Position = 1

        for campo, caption in ((DEST_DEBITO, CAPTION_DEBITO), (DEST_CREDITO, CAPTION_CREDITO)):
            valor = pt.AddDataField(pt.PivotFields(campos[_norm(campo)]), caption, XL_SUM)
            valor.NumberFormat = FORMATO_CIFRAS

        pt.RowGrand = True
        pt.ColumnGrand = True
        pt.RowAxisLayout(XL_COMPACT_ROW)
        for propiedad, valor in (("CompactLayoutRowHeader", rotulo_fila), ("TableStyle2", ESTILO_PIVOTE)):
            try:
                setattr(pt, propiedad, valor)
            except Exception:
                pass   # header text and style are cosmetic: an Excel build that refuses them still gets the numbers
    finally:
        pt.ManualUpdate = False
    return pt


def _quitar_hoja(wb, nombre):
    """Delete a sheet by name, never raising (cleanup after a failed build)."""
    try:
        wb.Worksheets(nombre).Delete()
    except Exception:
        pass


def _hoja_cifras(wb, ws_gl):
    """A fresh control sheet right after the policy's: a stale one from the template is deleted first."""
    if str(ws_gl.Name).casefold() == HOJA_CIFRAS.casefold():
        raise ValueError(f"la hoja de la póliza se llama igual que la de control ('{HOJA_CIFRAS}')")
    for ws in [w for w in wb.Worksheets]:
        if str(ws.Name).casefold() == HOJA_CIFRAS.casefold():
            _quitar_hoja(wb, str(ws.Name))
    nueva = wb.Worksheets.Add(After=ws_gl)
    try:
        nueva.Name = HOJA_CIFRAS
    except Exception:
        nueva.Delete()
        raise
    return nueva


def _crear_cifras_control(wb, ws_gl, columnas, fila_enc, bloques, datos, cuentas_credito, totales):
    """Add 'Validación - Cifras Control': title + two PivotTables over the policy (by GL Account and by Assignment),
    each read back against the rows that went up. NEVER raises: any failure removes the sheet and comes back as a
    warning, because the policy itself is already verified and must not be lost to a cosmetic control sheet.
    Returns (info, avisos); info is None when the sheet was not built."""
    nombre_hoja = None
    try:
        ws = _hoja_cifras(wb, ws_gl)
        nombre_hoja = str(ws.Name)
        celda = ws.Cells(FILA_TITULO_CIFRAS, 1)
        celda.Value2 = TITULO_CIFRAS
        celda.Font.Bold = True

        cols = [columnas[d] for d in (DEST_CUENTA, DEST_ASIGNACION, DEST_DEBITO, DEST_CREDITO)]
        fuente = (f"{_ref_hoja(ws_gl.Name)}!R{fila_enc}C{min(cols)}:R{bloques['credito'][1]}C{max(cols)}")
        cache = wb.PivotCaches().Create(SourceType=XL_DATABASE, SourceData=fuente)

        ocupados = _nombres_de_tablas(wb)
        pivotes, diferencias = [], []
        for clave, nombre, campo, rotulo, columna in PIVOTES_CIFRAS:
            pt = _crear_pivote(cache, ws, columna, _nombre_libre(nombre, ocupados), campo, rotulo)
            esperado = _esperado_cifras(datos, cuentas_credito, clave)
            diferencias += _verificar_pivote(pt, esperado, totales["debito"], totales["credito"])
            pivotes.append({"tabla": str(pt.Name), "campo": campo, "filas": len(esperado)})

        if diferencias:
            raise RuntimeError("las tablas no cuadran con la póliza:\n" + "\n".join(f"     • {d}" for d in diferencias))

        try:
            ws.Columns("A:H").AutoFit()
        except Exception:
            pass
        ws_gl.Activate()   # the policy stays the sheet that opens
        return {"hoja": nombre_hoja, "pivotes": pivotes}, []
    except Exception as e:
        if nombre_hoja:
            _quitar_hoja(wb, nombre_hoja)
        try:
            ws_gl.Activate()
        except Exception:
            pass
        return None, [
            f"⚠️ NO se creó la hoja '{HOJA_CIFRAS}' (la póliza sí quedó completa y verificada): {e}"
        ]


# --- Main entry point ---
def generar_poliza_batch(
    rutas_layout,
    ruta_plantilla,
    ruta_salida=None,
    callback_status=None,
    *,
    mes=None,
    referencia=None,
):
    """Fill ONE 'Póliza Carga Batch Creaciones' from the layouts ALREADY processed by generar_poliza_sap(): one file
    per channel (REGIONALES, PROXIMITY, NACIONALES, ECOMM, HD), in that order. A channel not given is warned, a
    channel given twice is an error. Rows with RESERVA ($) = 0, blank or '-' are left out. Layouts and template are
    never modified; ruta_salida=None -> the first layout's folder ('Output').
    mes: closing month key ('M10-26', None = current month) -> Posting/Document Date and Period.
    referencia: 'Reference' value; None = the user number from the profile folder (C:\\Users\\<number>).
    Returns: ruta, origenes, hojas_origen, canales, canales_faltantes, hoja, filas_leidas, filas_origen (rows that went
    up), filas_totales, omitidas, bloque_debito, bloque_credito, total_debito, total_credito, total_layout, naranjas,
    cuentas_credito, columnas, cifras_control, advertencias. Raises FileNotFoundError/ValueError (fixable),
    RuntimeError (Excel/clipboard/read-back failures), PermissionError (destination open)."""
    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    # --- Inputs ----------------------------------------------------------
    if isinstance(rutas_layout, (str, os.PathLike)):
        rutas_layout = [rutas_layout]
    rutas_layout = [os.fspath(r) for r in (rutas_layout or [])]
    if not rutas_layout:
        raise ValueError("No recibí ningún Layout procesado para armar la póliza.")

    for ruta, etiqueta in [(r, "el Layout procesado") for r in rutas_layout] + [
            (ruta_plantilla, "la plantilla de Póliza Carga Batch")]:
        if not ruta or not os.path.isfile(ruta):
            raise FileNotFoundError(f"No encontré {etiqueta}: {ruta}")

    vistas = set()
    for ruta in rutas_layout:
        clave = os.path.normcase(os.path.abspath(ruta))
        if clave in vistas:
            raise ValueError(f"El mismo Layout viene repetido:\n{ruta}")
        vistas.add(clave)

    if os.path.splitext(ruta_plantilla)[1].lower() not in EXTENSIONES_PLANTILLA:
        raise ValueError(
            f"La plantilla debe ser {' o '.join(EXTENSIONES_PLANTILLA)} y recibí "
            f"'{os.path.splitext(ruta_plantilla)[1]}'."
        )

    # Resolved before touching any file: a bad month key or user folder fails here.
    referencia = (referencia or "").strip() or referencia_de_usuario()
    campos = _campos_naranjas(mes, referencia)

    carpeta = os.path.abspath(ruta_salida) if ruta_salida \
        else os.path.dirname(os.path.abspath(rutas_layout[0]))
    os.makedirs(carpeta, exist_ok=True)

    ruta_final = os.path.join(carpeta, _nombre_salida(ruta_plantilla))
    if _misma_ruta(ruta_final, ruta_plantilla) or any(_misma_ruta(ruta_final, r) for r in rutas_layout):
        # Never overwrite one of the inputs.
        raise ValueError(f"El archivo de salida coincidiría con un archivo de entrada:\n{ruta_final}")

    # Built under a temporary name; renamed to the final one only after success.
    ruta_tmp = os.path.join(carpeta, f"_tmp{os.getpid()}_{os.path.basename(ruta_final)}")

    update_status("📑 Copiando la plantilla (la original no se toca)...")
    shutil.copy2(ruta_plantilla, ruta_tmp)

    resumen = None
    try:
        app = wb_destino = pythoncom = None
        abiertos = []
        try:
            update_status("📂 Abriendo Excel...")
            app, pythoncom = _abrir_excel()

            # --- Read every layout, one per channel -----------------------
            layouts = []
            for ruta in rutas_layout:
                update_status(f"📂 Leyendo '{os.path.basename(ruta)}'...")
                layouts.append(_leer_layout(app, ruta, abiertos))
            layouts, faltantes = _ordenar_layouts(layouts)

            wb_destino = _abrir_libro(app, ruta_tmp, solo_lectura=False)
            ws_destino = _hoja_plantilla(wb_destino)

            secciones = [s for lay in layouts for s in lay["secciones"]]
            avisos_layout = [a for lay in layouts for a in lay["avisos"]]
            if faltantes:
                avisos_layout.append(
                    f"Faltan los canales {', '.join(faltantes)}: la póliza sale SIN ellos. "
                    f"Si eso no es lo esperado, agrega sus templates en la tarjeta 1 y vuelve a correr el paso 3."
                )

            # --- Which rows go up, and are they valid? -------------------
            leidas = sum(s["filas"] for s in secciones)
            desglose = " + ".join(f"{lay['canal']} {sum(s['filas'] for s in lay['secciones']):,}" for lay in layouts)
            update_status(f"🔎 Revisando {leidas:,} renglones de {len(layouts)} canal(es) ({desglose})...")

            posiciones = _posiciones(secciones)
            todos = _leer_origen(secciones)
            por_reserva, omitidas = _omitir_por_reserva(todos)
            vivos = _filtrar(todos, por_reserva)
            if not vivos[CLAVE_MONTO]["valores"]:
                raise ValueError(
                    "Todos los renglones tienen RESERVA ($) en 0, en blanco o en '-': no hay nada que subir a SAP."
                )

            problemas = _problemas_origen(vivos)
            if problemas:
                raise ValueError(_texto_problemas(problemas, layouts))

            por_redondeo, redondeo_cero = _omitir_redondeo_cero(vivos)
            conservar = _componer(por_reserva, por_redondeo)
            datos = _filtrar(todos, conservar)
            n = len(datos[CLAVE_MONTO]["valores"])
            if n == 0:
                raise ValueError("Todos los renglones quedaron en 0.00 al redondear: no hay nada que subir a SAP.")

            omitidas_cuenta = {
                "cero": len(omitidas["cero"]), "blanco": len(omitidas["blanco"]),
                "guion": len(omitidas["guion"]), "redondeo": len(redondeo_cero),
            }
            omitidas_cuenta["total"] = sum(omitidas_cuenta.values())
            suma_omitida = math.fsum(
                v for v, c in zip(todos[CLAVE_MONTO]["valores"], conservar)
                if not c and isinstance(v, (int, float)) and not isinstance(v, bool)
                and not _error_de(v) and math.isfinite(v)
            )
            advertencias = (list(avisos_layout) + _avisos_omitidas(omitidas, redondeo_cero, suma_omitida)
                            + _advertencias_origen(datos) + _avisos_referencia(referencia))
            update_status(f"✂️  {n:,} renglón(es) van a la póliza; {omitidas_cuenta['total']:,} se dejan fuera (sin importe).")

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
            update_status("✍️  Pegando valores en la plantilla (Débito y Crédito)...")
            plan = _plan_pegado(secciones, posiciones, conservar)
            _pegar_todo(app, ws_destino, columnas, bloques, plan)

            # --- Orange fields + credit GL Account -------------------------
            update_status("🟧 Llenando campos fijos y cuentas del bloque de crédito...")
            desde, hasta = bloques["debito"][0], bloques["credito"][1]
            avisos_formato = []
            _rellenar_naranjas(ws_destino, columnas, desde, hasta, campos, avisos_formato)
            cuentas_credito = _completar_cuentas(ws_destino, columnas, desde, hasta, avisos_formato)

            # --- Read back and compare ------------------------------------
            update_status("🧮 Verificando lo escrito contra los Layouts...")
            diferencias, sumas, total_origen = _verificar(
                ws_destino, columnas, bloques, datos, n
            )
            diferencias += _verificar_naranjas(ws_destino, columnas, desde, hasta, campos)
            if diferencias:
                raise RuntimeError(
                    "La verificación posterior al pegado encontró diferencias, así que "
                    "NO se generó la Póliza Carga Batch:\n\n"
                    + "\n".join(f"     • {d}" for d in diferencias)
                )

            # --- Control sheet (never blocks the policy) -------------------
            update_status(f"📊 Armando '{HOJA_CIFRAS}' (dos tablas dinámicas)...")
            cuentas_por_fila = [_cuenta_contrapartida(v) for v in datos[CLAVE_CONCEPTO]["valores"]]
            cifras, avisos_cifras = _crear_cifras_control(
                wb_destino, ws_destino, columnas, fila_encabezado, bloques, datos, cuentas_por_fila,
                {"debito": sumas.get(DEST_DEBITO, 0.0), "credito": sumas.get(DEST_CREDITO, 0.0)},
            )
            advertencias = advertencias + avisos_formato + avisos_cifras

            update_status("💾 Guardando...")
            wb_destino.Save()

            resumen = {
                "ruta":              ruta_final,
                "origenes":          [{"ruta": lay["ruta"], "canal": lay["canal"],
                                       "hojas": [s["nombre"] for s in lay["secciones"]]} for lay in layouts],
                "hojas_origen":      [{"canal": s["canal"], "hoja": s["nombre"], "filas": s["filas"]}
                                      for s in secciones],
                "canales":           [lay["canal"] for lay in layouts],
                "canales_faltantes": faltantes,
                "hoja":              str(ws_destino.Name),
                "filas_leidas":      leidas,
                "filas_origen":      n,
                "filas_totales":     2 * n,
                "omitidas":          omitidas_cuenta,
                "bloque_debito":     bloques["debito"],
                "bloque_credito":    bloques["credito"],
                "total_debito":      round(sumas.get(DEST_DEBITO, 0.0), 2),
                "total_credito":     round(sumas.get(DEST_CREDITO, 0.0), 2),
                "total_layout":      round(total_origen, 2),
                "naranjas":          {destino: valor for destino, valor, _ in campos},
                "cuentas_credito":   cuentas_credito,
                "columnas":          [f"{_titulo(t['origen'])} → {t['destino']}"
                                      for t in TRANSFERENCIAS],
                "cifras_control":    cifras,
                "advertencias":      advertencias,
            }

        finally:
            # Layouts first (read-only, nothing to save), then the template, then Excel.
            for wb in abiertos:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            _cerrar_excel(app, wb_destino, pythoncom)

        # Excel is gone; retry the rename while EXCEL.EXE releases the temp file.
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


# --- Standalone run: python Create_Poliza_Batch.py ---
if __name__ == "__main__":
    print("=== Prueba de 'Póliza Carga Batch' (standalone) ===")
    entrada = input("Rutas de los Layouts YA procesados, uno por canal (separadas por ';'): ")
    layouts_ = [r.strip().strip('"') for r in entrada.split(";") if r.strip()]
    plantilla = input("Ruta de la plantilla de Póliza Carga Batch (.xlsx): ").strip().strip('"')
    mes_cierre = input("Mes de cierre (ej. M10-26, vacío = el mes en curso): ").strip().upper() or None

    try:
        res = generar_poliza_batch(layouts_, plantilla, callback_status=print, mes=mes_cierre)
        print("\n🎉 LISTO 🎉")
        for clave, valor in res.items():
            if clave in ("columnas", "advertencias", "canales", "canales_faltantes"):
                print(f"  {clave:<16}:")
                for v in valor:
                    print(f"      • {v}")
            elif clave in ("naranjas", "cuentas_credito", "omitidas"):
                print(f"  {clave:<16}:")
                for k, v in valor.items():
                    print(f"      • {k}: {v}")
            else:
                print(f"  {clave:<16}: {valor}")
    except Exception as e:
        print(f"❌ Error generando la póliza batch: {e}")
        raise
