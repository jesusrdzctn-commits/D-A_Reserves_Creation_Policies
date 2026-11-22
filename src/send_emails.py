"""RDA Pólizas D&A — send_emails.py: one e-mail per channel to its stakeholder, built from the 'Variaciones' sheet.
Text and screenshot come from the SAME saved sheet, so they cannot disagree: the bullets are read from the cells,
the picture is Excel's own snapshot of that block. Excel renders the pictures, Outlook sends (both through COM).
Recipients are confidential: they live in a per-user file OUTSIDE the project (see ruta_destinatarios), never in code.
Default mode is 'borrador' (saved in Outlook's Drafts for a last look); 'enviar' sends for real.
Everything that can fail runs BEFORE the first e-mail leaves: pictures first, then all the messages are built, then sent.
"""

import html
import json
import os
import re
import shutil
import tempfile
import time
import zlib
from urllib.parse import unquote

from openpyxl import load_workbook

from Consolidacion import GRUPOS, HOJA_VARIACIONES
from utils import abrir_excel, cerrar_excel, etiqueta_mes_es, hoja_por_nombre

# --- Message content ---
# 'MMPs' = millones de pesos: the sheet holds pesos, the stakeholder reads millions.
DIVISOR_MMP = 1_000_000
DECIMALES_MMP = 2

MODO_BORRADOR = "borrador"
MODO_ENVIAR   = "enviar"

# --- Recipients file ---
MAX_CORREOS_POR_CAMPO = 20
# Deliberately strict: no display names, quotes or angle brackets; one plain address per entry.
_CORREO = re.compile(r"^[A-Za-z0-9._%+'\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)+$")

# --- Excel / Outlook constants as plain numbers (win32com.client.constants breaks inside a PyInstaller .exe) ---
XL_SCREEN = 1
XL_PRINTER = 2                           # CopyPicture 'Appearance': as printed, does not depend on what the screen shows
XL_BITMAP = 2
XL_PICTURE = -4147
OL_MAIL_ITEM = 0
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

# Snapshot strategies, tried in order until one gives a picture that is not blank: (appearance, format, excel_visible).
# An invisible Excel sometimes copies an empty picture, so the last one lets Excel show itself for a moment.
ESTRATEGIAS_IMAGEN = (
    (XL_SCREEN,  XL_PICTURE, False),
    (XL_PRINTER, XL_PICTURE, False),
    (XL_SCREEN,  XL_BITMAP,  False),
    (XL_PRINTER, XL_BITMAP,  False),
    (XL_SCREEN,  XL_BITMAP,  True),
)
PAUSA_PORTAPAPELES = 0.3                 # seconds: let the clipboard settle between copy and paste
_PNG_FIRMA = b"\x89PNG\r\n\x1a\n"

# Layout of the 'Variaciones' blocks (see Consolidacion): A..H, banner row first, 'Grand Total' row last.
PRIMERA_COL, ULTIMA_COL = 1, 8
COL_ANTERIOR, COL_ACTUAL = 2, 4          # B and D (C is a spacer)
TEXTO_TOTAL = "Grand Total"


# ----------------------------------------------------------------------
# Recipients (stored per user, outside the project)
# ----------------------------------------------------------------------
def ruta_destinatarios():
    """%APPDATA%\\RDA_Polizas_DA\\destinatarios.json (RDA_DESTINATARIOS overrides it, for tests). Not next to the code:
    the project is versioned and shared, the addresses are not."""
    forzada = os.environ.get("RDA_DESTINATARIOS", "").strip()
    if forzada:
        return forzada
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "RDA_Polizas_DA", "destinatarios.json")


def parsear_correos(texto, campo="correos"):
    """'a@x.com; b@x.com' (also commas, spaces, line breaks) -> ['a@x.com', 'b@x.com']; no repeats (case-insensitive).
    Raises ValueError naming EVERY invalid entry: a typo must stop here, not bounce after sending."""
    crudos = [p for p in re.split(r"[;,\s]+", str(texto or "")) if p]
    invalidos = [p for p in crudos if not _CORREO.match(p) or len(p) > 254]
    if invalidos:
        raise ValueError(
            f"{campo}: correo(s) no válido(s): {', '.join(invalidos)}. "
            f"Usa solo direcciones sueltas (nombre@dominio.com), separadas por ';'."
        )
    vistos, correos = set(), []
    for correo in crudos:
        if correo.lower() not in vistos:
            vistos.add(correo.lower())
            correos.append(correo)
    if len(correos) > MAX_CORREOS_POR_CAMPO:
        raise ValueError(f"{campo}: máximo {MAX_CORREOS_POR_CAMPO} correos por campo (hay {len(correos)}).")
    return correos


def normalizar_destinatarios(datos, grupos=None):
    """{canal: {'para': text|list, 'cc': text|list}} -> {canal: {'para': [..], 'cc': [..]}}, validated.
    Unknown channels raise; a channel with empty 'para' is kept (means: send nothing for it); CC never replaces 'para'."""
    grupos = grupos if grupos is not None else GRUPOS
    canales = {g["canal"] for g in grupos}
    limpio = {}
    for canal, campos in (datos or {}).items():
        if canal not in canales:
            raise ValueError(f"Canal desconocido en los destinatarios: '{canal}'.")
        campos = campos if isinstance(campos, dict) else {}
        listas = {}
        for clave in ("para", "cc"):
            valor = campos.get(clave, "")
            texto = "; ".join(valor) if isinstance(valor, (list, tuple)) else valor
            listas[clave] = parsear_correos(texto, f"{canal} ({'Para' if clave == 'para' else 'CC'})")
        para = {c.lower() for c in listas["para"]}
        listas["cc"] = [c for c in listas["cc"] if c.lower() not in para]     # one copy per person
        limpio[canal] = listas
    return limpio


def cargar_destinatarios(ruta=None, grupos=None):
    """Saved recipients, or {} when there is no file. A damaged file is NOT silently trusted: it raises ValueError."""
    ruta = ruta or ruta_destinatarios()
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            datos = json.load(f)
    except FileNotFoundError:
        return {}
    except (OSError, ValueError) as e:
        raise ValueError(f"No pude leer los destinatarios guardados ({ruta}): {e}")
    if not isinstance(datos, dict):
        raise ValueError(f"El archivo de destinatarios ({ruta}) no tiene el formato esperado.")
    return normalizar_destinatarios(datos, grupos)


def guardar_destinatarios(datos, ruta=None, grupos=None):
    """Validate and save atomically (temp file + replace: a crash cannot leave half a file). Returns the clean dict."""
    limpio = normalizar_destinatarios(datos, grupos)
    ruta = ruta or ruta_destinatarios()
    carpeta = os.path.dirname(os.path.abspath(ruta))
    os.makedirs(carpeta, exist_ok=True)
    fd, temporal = tempfile.mkstemp(dir=carpeta, prefix=".dest_", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(limpio, f, ensure_ascii=False, indent=2)
        os.replace(temporal, ruta)
    except BaseException:
        try:
            os.remove(temporal)
        except OSError:
            pass
        raise
    return limpio


# ----------------------------------------------------------------------
# Reading the blocks of 'Variaciones' (pure: no Excel needed)
# ----------------------------------------------------------------------
def _numero(valor, celda):
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        raise ValueError(f"'{HOJA_VARIACIONES}'!{celda}: esperaba un número y encontré {valor!r}.")
    return float(valor)


def _texto(valor):
    return str(valor).strip() if valor is not None else ""


def bloques_desde_filas(filas, grupos=None):
    """Blocks of a 'Variaciones' sheet given as rows (tuples A..H, row 1 first). Returns
    {canal: {titulo, canal, mes_anterior, mes_actual, conceptos: [(name, previous, current)], rango, fila_inicio, fila_total}}.
    Fails loud on a malformed block, an unknown title or a repeated channel: a wrong e-mail is worse than none."""
    grupos = grupos if grupos is not None else GRUPOS
    por_titulo = {g["titulo"].strip().upper(): g["canal"] for g in grupos}
    bloques, i, n = {}, 0, len(filas)

    while i < n:
        if not _texto(filas[i][0]).startswith("Creaciones"):
            i += 1
            continue
        if i + 3 >= n:
            raise ValueError(f"'{HOJA_VARIACIONES}': el bloque de la fila {i + 1} está incompleto.")
        titulo, encabezado = _texto(filas[i + 1][0]), filas[i + 2]
        canal = por_titulo.get(titulo.upper())
        if canal is None:
            raise ValueError(f"'{HOJA_VARIACIONES}'!A{i + 2}: el título '{titulo}' no es de ningún canal conocido.")
        if canal in bloques:
            raise ValueError(f"'{HOJA_VARIACIONES}': el canal '{titulo}' aparece dos veces.")
        if _texto(encabezado[0]) != "Concepto":
            raise ValueError(f"'{HOJA_VARIACIONES}'!A{i + 3}: esperaba el encabezado 'Concepto'.")

        conceptos, j = [], i + 3
        while j < n and _texto(filas[j][0]) != TEXTO_TOTAL:
            nombre = _texto(filas[j][0])
            if not nombre:
                raise ValueError(f"'{HOJA_VARIACIONES}'!A{j + 1}: renglón sin concepto dentro del bloque '{titulo}'.")
            conceptos.append((
                nombre,
                _numero(filas[j][COL_ANTERIOR - 1], f"B{j + 1}"),
                _numero(filas[j][COL_ACTUAL - 1], f"D{j + 1}"),
            ))
            j += 1
        if j >= n:
            raise ValueError(f"'{HOJA_VARIACIONES}': el bloque '{titulo}' no termina en '{TEXTO_TOTAL}'.")

        bloques[canal] = {
            "titulo": titulo, "canal": canal,
            "mes_anterior": _texto(encabezado[COL_ANTERIOR - 1]), "mes_actual": _texto(encabezado[COL_ACTUAL - 1]),
            "conceptos": conceptos, "fila_inicio": i + 1, "fila_total": j + 1,
            "rango": f"A{i + 1}:H{j + 1}",
        }
        i = j + 1

    meses = {(b["mes_anterior"], b["mes_actual"]) for b in bloques.values()}
    if len(meses) > 1:
        raise ValueError(f"'{HOJA_VARIACIONES}': los bloques comparan meses distintos: {sorted(meses)}.")
    return bloques


def leer_bloques(ruta_libro, grupos=None):
    """Read the blocks from the saved workbook (read-only; the values of B and D are literal numbers, no recalculation needed)."""
    if not os.path.exists(ruta_libro):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta_libro}")
    wb = load_workbook(ruta_libro, read_only=True, data_only=False)
    try:
        if HOJA_VARIACIONES not in wb.sheetnames:
            raise ValueError(f"El archivo no tiene la pestaña '{HOJA_VARIACIONES}'. ¿Corriste 'Consolidar Variaciones'?")
        filas = list(wb[HOJA_VARIACIONES].iter_rows(
            min_col=PRIMERA_COL, max_col=ULTIMA_COL, values_only=True))
    finally:
        wb.close()
    filas = [tuple(f) + (None,) * (ULTIMA_COL - len(f)) for f in filas]
    bloques = bloques_desde_filas(filas, grupos)
    if not bloques:
        raise ValueError(f"No encontré ningún bloque en la pestaña '{HOJA_VARIACIONES}'. ¿Corriste 'Consolidar Variaciones'?")
    return bloques


# ----------------------------------------------------------------------
# The message (pure)
# ----------------------------------------------------------------------
def describir_variacion(anterior, actual):
    """'Incremento, representa 1.25 MMPs vs mes anterior.' — current minus previous (positive = increase)."""
    delta = round(actual - anterior, 2)
    if delta == 0:
        return "Sin variación vs mes anterior."
    tipo = "Incremento" if delta > 0 else "Disminución"
    millones = abs(delta) / DIVISOR_MMP
    if round(millones, DECIMALES_MMP) == 0:
        cantidad = f"menos de {10 ** -DECIMALES_MMP:.{DECIMALES_MMP}f}"
    else:
        cantidad = f"{millones:,.{DECIMALES_MMP}f}"
    return f"{tipo}, representa {cantidad} MMPs vs mes anterior."


def asunto_de(bloque):
    return f"Variaciones creación de reservas – {bloque['titulo']} – {etiqueta_mes_es(bloque['mes_actual'])}"


def cuerpo_html(bloque, cid):
    """Message body (HTML, everything escaped): greeting, the ask, one bullet per agrupador, the picture of the block."""
    items = "".join(
        f"<li><b>{html.escape(nombre)}</b>: {html.escape(describir_variacion(ant, act))}</li>"
        for nombre, ant, act in bloque["conceptos"]
    )
    return (
        '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">'
        "<p>Buen día,</p>"
        f"<p>Te comparto las variaciones que revisé en las creaciones de reservas del canal "
        f"<b>{html.escape(bloque['titulo'])}</b>, me apoyas por favor a validar si son correctas "
        f"o existe alguna modificación.</p>"
        f"<ul>{items}</ul>"
        f'<p><img src="cid:{html.escape(cid)}" alt="Variaciones {html.escape(bloque["titulo"])}"></p>'
        "</div>"
    )


def insertar_en_cuerpo(html_original, contenido):
    """Put `contenido` at the top of Outlook's own HTML body (which already carries the user's signature), if it has one."""
    coincidencia = re.search(r"<body[^>]*>", html_original or "", flags=re.IGNORECASE)
    if not coincidencia:
        return contenido
    return html_original[:coincidencia.end()] + contenido + html_original[coincidencia.end():]


def _cid(canal):
    return "variaciones_" + re.sub(r"[^a-z0-9]+", "_", canal.lower()).strip("_")


def planear_correos(bloques, destinatarios, canales=None, grupos=None):
    """(correos, omitidos): one dict per e-mail to send, and [(canal, motivo)] for the channels that get none.
    `canales` limits the run to those channels (None = all). Nothing is sent or written here."""
    grupos = grupos if grupos is not None else GRUPOS
    destinatarios = normalizar_destinatarios(destinatarios, grupos)
    correos, omitidos = [], []
    for g in grupos:
        canal = g["canal"]
        if canales is not None and canal not in canales:
            continue
        bloque = bloques.get(canal)
        para = destinatarios.get(canal, {}).get("para", [])
        if bloque is None:
            omitidos.append((canal, f"no hay bloque de '{g['titulo']}' en '{HOJA_VARIACIONES}'"))
        elif not bloque["conceptos"]:
            omitidos.append((canal, "el bloque no trae conceptos (nada que reportar)"))
        elif not para:
            omitidos.append((canal, "no tiene destinatario"))
        else:
            cid = _cid(canal)
            correos.append({
                "canal": canal, "titulo": bloque["titulo"], "para": para,
                "cc": destinatarios[canal]["cc"], "asunto": asunto_de(bloque), "cid": cid,
                "rango": bloque["rango"], "html": cuerpo_html(bloque, cid),
                "conceptos": len(bloque["conceptos"]),
            })
    return correos, omitidos


def previsualizar_correos(ruta_libro, destinatarios=None, canales=None):
    """What a run WOULD do, without Excel or Outlook: (bloques, correos, omitidos). Used by the confirmation box."""
    bloques = leer_bloques(ruta_libro)
    destinatarios = cargar_destinatarios() if destinatarios is None else destinatarios
    correos, omitidos = planear_correos(bloques, destinatarios, canales)
    return bloques, correos, omitidos


# ----------------------------------------------------------------------
# Excel pictures and Outlook messages (COM, Windows only)
# ----------------------------------------------------------------------
def _describir_error(e):
    """A COM error as readable text: 'The parameter is incorrect. (0x80070057)' instead of a raw tuple."""
    args = getattr(e, "args", ())
    if len(args) > 2 and isinstance(args[2], tuple) and len(args[2]) > 2 and args[2][2]:
        return str(args[2][2]).strip()
    if len(args) > 1 and isinstance(args[0], int) and isinstance(args[1], str):
        return f"{args[1].strip()} (0x{args[0] & 0xFFFFFFFF:08X})"
    return str(e)


def _es_png(ruta):
    try:
        with open(ruta, "rb") as f:
            return os.path.getsize(ruta) >= 67 and f.read(8) == _PNG_FIRMA     # 67 bytes = the smallest valid PNG
    except OSError:
        return False


def _paeth(a, b, c):
    p = a + b - c
    pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
    if pa <= pb and pa <= pc:
        return a
    return b if pb <= pc else c


def _png_en_blanco(ruta):
    """True when the PNG is one flat colour (what Excel exports when the paste into the chart silently failed).
    Reads 8-bit, non-interlaced PNGs (Excel's output) and stops at the first row that is not flat. Anything it cannot
    read counts as 'has content': the check may miss a blank picture, it never rejects a good one."""
    try:
        with open(ruta, "rb") as f:
            datos = f.read()
        if datos[:8] != _PNG_FIRMA:
            return False
        pos, ihdr, idat = 8, None, []
        while pos + 8 <= len(datos):
            largo = int.from_bytes(datos[pos:pos + 4], "big")
            tipo, cuerpo = datos[pos + 4:pos + 8], datos[pos + 8:pos + 8 + largo]
            if tipo == b"IHDR":
                ihdr = cuerpo
            elif tipo == b"IDAT":
                idat.append(cuerpo)
            elif tipo == b"IEND":
                break
            pos += 12 + largo
        if ihdr is None or len(ihdr) < 13 or not idat:
            return False
        ancho, alto = int.from_bytes(ihdr[0:4], "big"), int.from_bytes(ihdr[4:8], "big")
        profundidad, tipo_color, entrelazado = ihdr[8], ihdr[9], ihdr[12]
        bpp = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}.get(tipo_color)
        if bpp is None or profundidad != 8 or entrelazado != 0 or ancho <= 0 or alto <= 0:
            return False
        largo_fila = ancho * bpp
        crudo = zlib.decompress(b"".join(idat))
        if len(crudo) != alto * (largo_fila + 1):
            return False

        previa, plana = bytearray(largo_fila), None
        for y in range(alto):
            base = y * (largo_fila + 1)
            filtro = crudo[base]
            fila = bytearray(crudo[base + 1:base + 1 + largo_fila])
            if filtro not in (0, 1, 2, 3, 4):
                return False
            if filtro:
                for i in range(largo_fila):
                    a = fila[i - bpp] if i >= bpp else 0
                    b = previa[i]
                    c = previa[i - bpp] if i >= bpp else 0
                    if filtro == 1:
                        fila[i] = (fila[i] + a) & 255
                    elif filtro == 2:
                        fila[i] = (fila[i] + b) & 255
                    elif filtro == 3:
                        fila[i] = (fila[i] + ((a + b) >> 1)) & 255
                    else:
                        fila[i] = (fila[i] + _paeth(a, b, c)) & 255
            if plana is None:
                plana = bytes(fila[:bpp]) * ancho         # the first pixel, repeated, is what every row must equal
            if fila != plana:
                return False
            previa = fila
        return True
    except Exception:
        return False


def _tomar_foto(app, pythoncom, ws, direccion, destino, apariencia, formato):
    """One snapshot attempt: copy the range as a picture, paste it into a throw-away chart and export that chart."""
    rango = ws.Range(direccion)
    ws.Activate()
    rango.CopyPicture(apariencia, formato)
    time.sleep(PAUSA_PORTAPAPELES)
    try:
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass
    # A chart is the only object that can export what the clipboard holds as a picture.
    objeto = ws.ChartObjects().Add(rango.Left, rango.Top, rango.Width, rango.Height)
    try:
        try:
            objeto.ShapeRange.Line.Visible = 0
        except Exception:
            pass
        objeto.Activate()                 # without it Chart.Paste() pastes nothing on some Excel 2016+ builds
        objeto.Chart.Paste()
        objeto.Chart.Export(destino, "PNG")
    finally:
        try:
            objeto.Delete()
        except Exception:
            pass
        try:
            app.CutCopyMode = False
        except Exception:
            pass


def exportar_imagenes(ruta_libro, rangos, carpeta):
    """{canal: path of a PNG} with Excel's snapshot of each range of 'Variaciones'. `rangos` = {canal: 'A1:H20'}.
    The workbook is opened READ-ONLY and closed without saving: the file is never touched.
    A picture that comes out blank is never accepted: the next strategy of ESTRATEGIAS_IMAGEN is tried, and if all
    fail this raises, before anything has been mailed."""
    app, pythoncom = abrir_excel()
    wb = None
    try:
        wb = app.Workbooks.Open(os.path.abspath(ruta_libro), UpdateLinks=0, ReadOnly=True)
        ws = hoja_por_nombre(wb, HOJA_VARIACIONES)
        try:
            app.ScreenUpdating = True      # abrir_excel() turns it off; a snapshot of a frozen screen can be empty
        except Exception:
            pass
        ws.Activate()
        try:
            app.ActiveWindow.DisplayGridlines = False       # cosmetic: a cleaner picture
        except Exception:
            pass

        rutas = {}
        for canal, direccion in rangos.items():
            destino = os.path.join(carpeta, f"{_cid(canal)}.png")
            motivos = []
            for apariencia, formato, visible in ESTRATEGIAS_IMAGEN:
                try:
                    if os.path.exists(destino):
                        os.remove(destino)
                    app.Visible = visible
                    _tomar_foto(app, pythoncom, ws, direccion, destino, apariencia, formato)
                    if not _es_png(destino):
                        motivos.append("Excel no dejó una imagen PNG válida")
                    elif _png_en_blanco(destino):
                        motivos.append("la imagen salió en blanco")
                    else:
                        rutas[canal] = destino
                        break
                except Exception as e:
                    # The clipboard is shared with whatever the user is doing: just try the next strategy.
                    motivos.append(_describir_error(e))
                time.sleep(0.5)
            if canal not in rutas:
                raise RuntimeError(f"No pude sacar una imagen válida de {canal} ({direccion}): {'; '.join(motivos)}")
        return rutas
    finally:
        try:
            app.Visible = False
        except Exception:
            pass
        cerrar_excel(app, wb, pythoncom)


def _abrir_outlook():
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        raise RuntimeError("Este proceso necesita 'pywin32' y Outlook instalado.\n\nInstálalo con:   pip install pywin32")
    pythoncom.CoInitialize()
    try:
        return win32.Dispatch("Outlook.Application"), pythoncom
    except Exception as e:
        pythoncom.CoUninitialize()
        raise RuntimeError(f"No pude abrir Outlook: {e}\n\nÁbrelo y vuelve a intentar.")


def _carpeta_firmas():
    base = os.environ.get("APPDATA")
    return os.path.join(base, "Microsoft", "Signatures") if base else None


def _nombre_firma_predeterminada():
    """Name of the default signature for NEW messages (Outlook keeps it in the registry), or None."""
    try:
        import winreg
    except ImportError:
        return None
    for version in ("16.0", "15.0", "14.0"):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                rf"Software\Microsoft\Office\{version}\Common\MailSettings") as clave:
                nombre = str(winreg.QueryValueEx(clave, "NewSignature")[0]).strip()
                if nombre:
                    return nombre
        except OSError:
            continue
    return None


def leer_firma(carpeta=None, nombre=None):
    """(html, [(cid, path)]): the user's default signature read from Outlook's signature folder, with its local pictures
    swapped for cid: references. Reading the file replaces the old GetInspector trick, which breaks MailItem.Send on
    recent Outlook builds. It never raises: no signature or an unreadable one gives ('', [])."""
    try:
        carpeta = carpeta or _carpeta_firmas()
        nombre = nombre or _nombre_firma_predeterminada()
        if not carpeta or not nombre or os.path.basename(nombre) != nombre:
            return "", []
        ruta = os.path.join(carpeta, nombre + ".htm")
        with open(ruta, "rb") as f:
            crudo = f.read()
        charset = re.search(rb"charset=[\"']?([A-Za-z0-9_\-]+)", crudo[:4096], flags=re.IGNORECASE)
        try:
            texto = crudo.decode(charset.group(1).decode("ascii") if charset else "cp1252", errors="replace")
        except LookupError:
            texto = crudo.decode("cp1252", errors="replace")
        cuerpo = re.search(r"<body[^>]*>(.*)</body>", texto, flags=re.IGNORECASE | re.DOTALL)
        if not cuerpo:
            return "", []

        raiz, imagenes = os.path.realpath(carpeta), []

        def cambiar(m):
            fuente = unquote(m.group(2))
            if re.match(r"^[a-z][a-z0-9+.\-]*:", fuente, flags=re.IGNORECASE):
                return m.group(0)                    # http:, cid:, data:... is left as it is
            destino = os.path.realpath(os.path.join(carpeta, fuente))
            if not destino.startswith(raiz + os.sep) or not os.path.isfile(destino):
                return m.group(0)
            cid = f"firma_{len(imagenes) + 1}"
            imagenes.append((cid, destino))
            return f'{m.group(1)}cid:{cid}{m.group(3)}'

        html_firma = re.sub(r"""(<img\b[^>]*?\bsrc\s*=\s*["'])([^"']+)(["'])""", cambiar, cuerpo.group(1),
                            flags=re.IGNORECASE)
        return html_firma, imagenes
    except Exception:
        return "", []


def _tiene_contenido(html_original):
    """True when Outlook's own HTML body already holds something (a signature it inserted itself)."""
    cuerpo = re.search(r"<body[^>]*>(.*)</body>", html_original or "", flags=re.IGNORECASE | re.DOTALL)
    interior = cuerpo.group(1) if cuerpo else (html_original or "")
    if re.search(r"<img\b", interior, flags=re.IGNORECASE):
        return True
    texto = re.sub(r"<[^>]+>", "", interior)
    return bool(html.unescape(texto).replace("\xa0", " ").strip())


def _adjuntar_inline(mensaje, ruta, cid):
    adjunto = mensaje.Attachments.Add(os.path.abspath(ruta))
    adjunto.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)
    try:
        adjunto.PropertyAccessor.SetProperty(PR_ATTACHMENT_HIDDEN, True)   # shown inline, not as a paperclip
    except Exception:
        pass


def _armar_mensaje(outlook, correo, ruta_imagen):
    """One Outlook MailItem, fully built and NOT sent. The default signature goes below our text.
    The item is never displayed nor inspected (no GetInspector): that makes MailItem.Send fail on recent Outlook."""
    mensaje = outlook.CreateItem(OL_MAIL_ITEM)
    mensaje.To = "; ".join(correo["para"])
    if correo["cc"]:
        mensaje.CC = "; ".join(correo["cc"])
    mensaje.Subject = correo["asunto"]

    if not mensaje.Recipients.ResolveAll():
        sin_resolver = [r.Name for r in mensaje.Recipients if not r.Resolved]
        raise RuntimeError(f"Outlook no reconoce estos destinatarios: {', '.join(sin_resolver) or '(sin detalle)'}")

    original = mensaje.HTMLBody or ""
    firma, imagenes_firma = ("", []) if _tiene_contenido(original) else leer_firma()
    _adjuntar_inline(mensaje, ruta_imagen, correo["cid"])
    for cid, ruta in imagenes_firma:
        _adjuntar_inline(mensaje, ruta, cid)
    contenido = correo["html"] + (f"<br>{firma}" if firma else "")
    mensaje.HTMLBody = insertar_en_cuerpo(original, contenido)
    return mensaje


def enviar_correos(ruta_libro, destinatarios=None, modo=MODO_BORRADOR, canales=None, callback_status=None):
    """Build the pictures, then ALL the messages, and only then save (modo 'borrador') or send (modo 'enviar') them.
    Returns {'modo', 'mes_anterior', 'mes_actual', 'resultados': [{'canal','titulo','estado','detalle'}], 'omitidos': [...]};
    estado is 'enviado' | 'borrador' | 'error'. Raises (and sends nothing) if the sheet, a picture or a message fails."""
    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    if modo not in (MODO_BORRADOR, MODO_ENVIAR):
        raise ValueError(f"Modo desconocido: '{modo}' (usa '{MODO_BORRADOR}' o '{MODO_ENVIAR}').")

    update_status(f"📄 Leyendo la pestaña '{HOJA_VARIACIONES}'...")
    bloques, correos, omitidos = previsualizar_correos(ruta_libro, destinatarios, canales)
    if not correos:
        motivos = "\n".join(f"  • {c}: {m}" for c, m in omitidos) or "  (ningún canal seleccionado)"
        raise ValueError(f"No hay ningún correo que enviar:\n{motivos}")

    mes_ant, mes_act = next(iter(bloques.values()))["mes_anterior"], next(iter(bloques.values()))["mes_actual"]
    carpeta = tempfile.mkdtemp(prefix="rda_correos_")
    outlook = pythoncom = None
    try:
        update_status(f"📸 Sacando {len(correos)} imagen(es) de las tablas con Excel...")
        imagenes = exportar_imagenes(ruta_libro, {c["canal"]: c["rango"] for c in correos}, carpeta)

        update_status("✉️ Armando los mensajes en Outlook...")
        outlook, pythoncom = _abrir_outlook()
        mensajes = []
        try:
            for correo in correos:
                mensajes.append((correo, _armar_mensaje(outlook, correo, imagenes[correo["canal"]])))
        except Exception as e:
            raise RuntimeError(f"No pude armar el mensaje de {correo['canal']}: {_describir_error(e)}\n\nNo se envió ningún correo.")

        resultados = []
        for correo, mensaje in mensajes:
            try:
                if modo == MODO_ENVIAR:
                    update_status(f"📤 Enviando {correo['titulo']}...")
                    mensaje.Send()
                    estado = "enviado"
                else:
                    update_status(f"📝 Guardando borrador de {correo['titulo']}...")
                    mensaje.Save()
                    estado = "borrador"
                resultados.append({"canal": correo["canal"], "titulo": correo["titulo"], "estado": estado,
                                   "detalle": "; ".join(correo["para"])})
            except Exception as e:
                detalle = _describir_error(e)
                if modo == MODO_ENVIAR:
                    # Not sent: keep the finished message in Drafts so the work is not lost (Save never sends).
                    try:
                        mensaje.Save()
                        detalle += " — el mensaje quedó guardado en Borradores"
                    except Exception:
                        pass
                resultados.append({"canal": correo["canal"], "titulo": correo["titulo"], "estado": "error",
                                   "detalle": detalle})
        return {"modo": modo, "mes_anterior": mes_ant, "mes_actual": mes_act,
                "resultados": resultados, "omitidos": omitidos}
    finally:
        try:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        except Exception:
            pass
        shutil.rmtree(carpeta, ignore_errors=True)


# --- Standalone run: python send_emails.py ---
if __name__ == "__main__":
    print("=== Correos de variaciones a los stakeholders (standalone) ===")
    ruta = input("Ruta del archivo con la pestaña 'Variaciones' (.xlsx): ").strip().strip('"')
    modo_txt = input("¿Enviar de verdad? [s = enviar / Enter = solo guardar borradores]: ").strip().lower()
    try:
        _, planeados, saltados = previsualizar_correos(ruta)
        print(f"\nSe prepararán {len(planeados)} correo(s):")
        for c in planeados:
            print(f"  • {c['titulo']:<18} → {'; '.join(c['para'])}" + (f"  (CC: {'; '.join(c['cc'])})" if c["cc"] else ""))
        for canal, motivo in saltados:
            print(f"  ⚠️ {canal}: {motivo}")
        resumen = enviar_correos(ruta, modo=MODO_ENVIAR if modo_txt == "s" else MODO_BORRADOR, callback_status=print)
        print("\n🎉 LISTO 🎉")
        for r in resumen["resultados"]:
            print(f"  {r['estado']:<9} {r['titulo']}: {r['detalle']}")
    except Exception as e:
        print(f"❌ Error: {e}")
        raise
