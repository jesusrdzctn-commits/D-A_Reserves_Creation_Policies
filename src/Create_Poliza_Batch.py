"""RDA Pólizas D&A — copies the finished 'PÓLIZA SAP' of the OT layout (both tabs) into the template
'Póliza Carga Batch' ('GL Posting Upload'): one debit block, then one credit block.
Runs after generar_poliza_sap(). Pastes values, fills the orange header fields on every row and the
credit block's GL Account, reads everything back and saves only if it matches.
"""

import math
import os
import re
import shutil
import time
from datetime import datetime

# Layout helpers are reused from Create_Poliza_SAP so both modules agree on where the section starts and ends.
from Create_Poliza_SAP import (
    CLAVE_MONTO,
    XL_PASTE_VALUES,
    localizar_secciones,
    titulo_de,
)
from utils import abrir_excel as _abrir_excel
from utils import cerrar_excel as _cerrar_excel
from utils import error_de_excel as _error_de
from utils import es_vacio_excel as _vacio
from utils import hoja_por_nombre as _hoja
from utils import mes_de_cierre
from utils import norm_texto as _norm
from utils import ultimo_dia_mes

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
# The column the amount is compared against, to warn when they differ.
CLAVE_COMPARACION = "mxn" if CLAVE_MONTO != "mxn" else "redondeo"
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


# --- Small helpers ---
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
        (DEST_PERIODO, mes_de_cierre(mes).month, False),
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


# --- Layout (source) ---
def _leer_origen(secciones):
    """{key: {'valores', 'celdas', 'titulo'}} for every copied layout column (plus the amount's comparison column).
    Rows of all tabs follow one another in tab order, as pasted; 'celdas' holds each address ('Reservas Rentas!K8')."""
    claves = _claves_origen()
    if CLAVE_COMPARACION not in claves:
        claves.append(CLAVE_COMPARACION)

    datos = {}
    for clave in claves:
        valores, celdas = [], []
        for s in secciones:
            col = s["columnas"][clave]
            rango = s["hoja"].Range(
                f"{col['letra']}{s['fila_inicio']}:{col['letra']}{s['fila_fin']}"
            )
            valores += _leer_rango(rango, s["filas"])
            celdas += [f"{s['nombre']}!{col['letra']}{s['fila_inicio'] + i}"
                       for i in range(s["filas"])]
        datos[clave] = {"valores": valores, "celdas": celdas, "titulo": titulo_de(clave)}
    return datos


def _enumerar(celdas):
    """First MAX_DETALLE cells, plus an ellipsis when there are more."""
    return ", ".join(celdas[:MAX_DETALLE]) + ("..." if len(celdas) > MAX_DETALLE else "")


def _validar_origen(datos):
    """Check what is about to go up to SAP: returns (problemas, advertencias).
    problemas BLOCK the run (error cells, empty cells, non-numeric amount); advertencias only report
    (zero/negative amounts, numeric GL accounts, values over SAP lengths, REDONDEO != MXN)."""
    problemas = []
    advertencias = []

    for clave in _claves_origen():
        d = datos[clave]
        es_monto = clave == CLAVE_MONTO
        grupos = {"error": [], "vacío": [], "no numérico": [], "tipo inesperado": []}

        for celda, valor in zip(d["celdas"], d["valores"]):
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
                    "titulo": d["titulo"], "tipo": tipo,
                    "total": len(celdas), "celdas": celdas[:MAX_DETALLE],
                })

    # --- Warnings (only meaningful when the cells themselves are usable) ---
    if not problemas:
        monto = datos[CLAVE_MONTO]
        ceros = [c for c, v in zip(monto["celdas"], monto["valores"]) if v == 0]
        negativos = [c for c, v in zip(monto["celdas"], monto["valores"]) if v < 0]
        if ceros:
            advertencias.append(
                f"{len(ceros)} renglón(es) con {monto['titulo']} = 0 (se pasan tal cual): "
                + _enumerar(ceros)
            )
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

    return problemas, advertencias


def _texto_problemas(problemas, ruta_layout):
    """Readable message for the blocking problems found in the layout."""
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
        f"No se generó la Póliza Carga Batch: el Layout todavía tiene {total} "
        f"celda(s) con problemas que SAP rechazaría.\n\n"
        + "\n".join(lineas)
        + "\n\nCorrígelas (normalmente es un valor que falta en la pestaña "
          "'Catalogo') y vuelve a correr 'Generar PÓLIZA SAP'.\n\n"
          f"Layout revisado:\n{ruta_layout}"
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


def _pegar_todo(app, ws_destino, columnas, bloques, secciones):
    """Paste every (column, block) of TRANSFERENCIAS as values; each tab lands at its 'desplazamiento' offset."""
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
            for s in secciones:
                letra = s["columnas"][t["origen"]]["letra"]
                origen = s["hoja"].Range(f"{letra}{s['fila_inicio']}:{letra}{s['fila_fin']}")
                desde = r1 + s["desplazamiento"]
                hasta = desde + s["filas"] - 1
                _pegar_valores(
                    app, origen, _rango_destino(ws_destino, columnas[t["destino"]], desde, hasta)
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
                f"Σ {destino} = {suma:,.2f} y la Σ MONTO del Layout = {total_origen:,.2f}"
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


# --- Main entry point ---
def generar_poliza_batch(
    ruta_layout,
    ruta_plantilla,
    ruta_salida=None,
    callback_status=None,
    *,
    mes=None,
    referencia=None,
):
    """Fill 'Póliza Carga Batch Creaciones' from the layout ALREADY processed by generar_poliza_sap().
    Layout (read-only) and template are never modified; ruta_salida=None -> the layout's folder ('Output').
    mes: closing month key ('M10-26', None = current month) -> Posting/Document Date and Period.
    referencia: 'Reference' value; None = the user number from the profile folder (C:\\Users\\<number>).
    Returns: ruta, origen, hojas_origen, hoja, filas_origen, filas_totales, bloque_debito, bloque_credito,
    total_debito, total_credito, naranjas, cuentas_credito, columnas, advertencias. Raises FileNotFoundError/
    ValueError (fixable), RuntimeError (Excel/clipboard/read-back failures), PermissionError (destination open)."""
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

    # Resolved before touching any file: a bad month key or user folder fails here.
    referencia = (referencia or "").strip() or referencia_de_usuario()
    campos = _campos_naranjas(mes, referencia)

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

            ws_destino = _hoja_plantilla(wb_destino)

            # --- Read and validate the layout ----------------------------
            secciones, avisos_layout = localizar_secciones(wb_origen)
            # A tab without data has filas == 0 and fila_fin None: nothing to copy from it.
            secciones = [s for s in secciones if s["filas"] > 0]
            if not secciones:
                raise ValueError("El Layout no tiene renglones con datos en ninguna pestaña.")
            n = 0
            for s in secciones:
                s["desplazamiento"] = n      # where this tab starts inside each block
                n += s["filas"]
            desglose = " + ".join(f"{s['etiqueta']} {s['filas']}" for s in secciones)
            update_status(f"🔎 Revisando {n:,} renglones del Layout ({desglose})...")

            datos = _leer_origen(secciones)
            problemas, advertencias = _validar_origen(datos)
            if problemas:
                raise ValueError(_texto_problemas(problemas, ruta_layout))
            advertencias = list(avisos_layout) + advertencias + _avisos_referencia(referencia)

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
            _pegar_todo(app, ws_destino, columnas, bloques, secciones)

            # --- Orange fields + credit GL Account -------------------------
            update_status("🟧 Llenando campos fijos y cuentas del bloque de crédito...")
            desde, hasta = bloques["debito"][0], bloques["credito"][1]
            avisos_formato = []
            _rellenar_naranjas(ws_destino, columnas, desde, hasta, campos, avisos_formato)
            cuentas_credito = _completar_cuentas(ws_destino, columnas, desde, hasta, avisos_formato)

            # --- Read back and compare ------------------------------------
            update_status("🧮 Verificando lo escrito contra el Layout...")
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

            advertencias = advertencias + avisos_formato

            update_status("💾 Guardando...")
            wb_destino.Save()

            resumen = {
                "ruta":           ruta_final,
                "origen":         os.path.abspath(ruta_layout),
                "hojas_origen":   [{"hoja": s["nombre"], "filas": s["filas"]} for s in secciones],
                "hoja":           str(ws_destino.Name),
                "filas_origen":   n,
                "filas_totales":  2 * n,
                "bloque_debito":  bloques["debito"],
                "bloque_credito": bloques["credito"],
                "total_debito":   round(sumas.get(DEST_DEBITO, 0.0), 2),
                "total_credito":  round(sumas.get(DEST_CREDITO, 0.0), 2),
                "total_layout":   round(total_origen, 2),
                "naranjas":       {destino: valor for destino, valor, _ in campos},
                "cuentas_credito": cuentas_credito,
                "columnas":       [f"{titulo_de(t['origen'])} → {t['destino']}"
                                   for t in TRANSFERENCIAS],
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
    layout = input("Ruta del Layout OT YA procesado (.xlsx): ").strip().strip('"')
    plantilla = input("Ruta de la plantilla de Póliza Carga Batch (.xlsx): ").strip().strip('"')
    mes_cierre = input("Mes de cierre (ej. M10-26, vacío = el mes en curso): ").strip().upper() or None

    try:
        res = generar_poliza_batch(layout, plantilla, callback_status=print, mes=mes_cierre)
        print("\n🎉 LISTO 🎉")
        for clave, valor in res.items():
            if clave in ("columnas", "advertencias"):
                print(f"  {clave:<16}:")
                for v in valor:
                    print(f"      • {v}")
            elif clave in ("naranjas", "cuentas_credito"):
                print(f"  {clave:<16}:")
                for k, v in valor.items():
                    print(f"      • {k}: {v}")
            else:
                print(f"  {clave:<16}: {valor}")
    except Exception as e:
        print(f"❌ Error generando la póliza batch: {e}")
        raise
