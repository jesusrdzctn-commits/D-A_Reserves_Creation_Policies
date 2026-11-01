"""RDA Pólizas D&A — shared helpers (month keys, output folders, Excel COM).
'M06-26' (M + month + '-' + year) is the single source of truth for the 'Mes' column.
win32com is imported inside abrir_excel(), so this file imports cleanly without Excel.
"""

import calendar
import os
import unicodedata
from datetime import datetime

# Month key as written in the 'Mes' column of the 'Base' sheet: 'M06-26'.
FORMATO_CLAVE_MES = "M%m-%y"

MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre",
    12: "diciembre",
}


# --- Month keys ---
def clave_mes(fecha):
    """datetime(2026, 6, 15) -> 'M06-26'."""
    return fecha.strftime(FORMATO_CLAVE_MES)


def parse_clave_mes(clave):
    """'M06-26' -> datetime(2026, 6, 1). Raises ValueError on a malformed key
    so a typo fails loudly instead of producing an empty matrix."""
    texto = str(clave).strip().upper()
    try:
        return datetime.strptime(texto, "M%m-%y")
    except ValueError:
        raise ValueError(
            f"La clave de mes '{clave}' no tiene el formato esperado 'M06-26' "
            f"(M + mes de 2 dígitos + '-' + año de 2 dígitos)."
        )


def mes_anterior(fecha):
    """First day of the PREVIOUS month (day 1 first, so March 31 never becomes Feb 31)."""
    primero = fecha.replace(day=1)
    if primero.month == 1:
        return primero.replace(year=primero.year - 1, month=12)
    return primero.replace(month=primero.month - 1)


def mes_de_cierre(fecha=None):
    """First day of the closing month: None = today's month, a month key ('M10-26') or a datetime."""
    if fecha is None:
        fecha = datetime.today()
    elif isinstance(fecha, str):
        fecha = parse_clave_mes(fecha)
    return fecha.replace(day=1)


def ultimo_dia_mes(fecha):
    """Last day of the month of `fecha` (any day of October 2026 -> 31.10.2026)."""
    return fecha.replace(day=calendar.monthrange(fecha.year, fecha.month)[1])


def claves_mes_actual_y_anterior(fecha=None):
    """Pair (clave_anterior, clave_actual) for 'Variaciones'. `fecha` is a datetime,
    a month key ('M07-26') to re-run a past month, or None for the current month."""
    if fecha is None:
        fecha = datetime.today()
    elif isinstance(fecha, str):
        fecha = parse_clave_mes(fecha)

    return clave_mes(mes_anterior(fecha)), clave_mes(fecha)


def etiqueta_mes_es(clave):
    """'M06-26' -> 'junio 2026' (for status messages)."""
    fecha = parse_clave_mes(clave)
    return f"{MESES_ES[fecha.month]} {fecha.year}"


# --- Output folder ---
def carpeta_output_desde_input(ruta_input):
    """'Output' folder beside an 'Input' folder, or inside any other chosen folder.
    A folder already named 'Output' is returned as is, so steps never nest Output/Output."""
    ruta = os.path.normpath(ruta_input)
    nombre = os.path.basename(ruta).lower()
    if nombre == "output":
        return ruta
    if nombre == "input":
        return os.path.join(os.path.dirname(ruta), "Output")
    return os.path.join(ruta, "Output")


# --- Excel COM plumbing (shared by Poliza_SAP and Comparativos) ---
def norm_texto(texto):
    """Trim, upper-case and drop accents, tolerating None. 'Cebe' -> 'CEBE'."""
    if texto is None:
        return ""
    limpio = unicodedata.normalize("NFKD", str(texto))
    limpio = limpio.encode("ascii", "ignore").decode("ascii")
    return " ".join(limpio.upper().split())


# Excel error values as pywin32 returns them from Range.Value2 (CVErr ints).
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


def error_de_excel(valor):
    """Excel error text ('#N/A', '#REF!'...) of a cell value, or None. Accepts CVErr ints and literal text."""
    if isinstance(valor, bool):          # bool is an int subclass: exclude it first
        return None
    if isinstance(valor, int):
        return CODIGOS_ERROR_EXCEL.get(valor)
    if isinstance(valor, str) and valor.strip() in CODIGOS_ERROR_EXCEL.values():
        return valor.strip()
    return None


def es_vacio_excel(valor):
    """True for a cell that holds nothing: None, or text made only of whitespace."""
    return valor is None or (isinstance(valor, str) and not valor.strip())


def abrir_excel():
    """Start a PRIVATE Excel instance and return (app, pythoncom).
    DispatchEx, not Dispatch: Dispatch would attach to the user's open Excel and
    our app.Quit() would close THEIR workbooks."""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        raise RuntimeError(
            "Este proceso necesita 'pywin32' y Excel instalado.\n\n"
            "Instálalo con:   pip install pywin32"
        )

    pythoncom.CoInitialize()

    app = win32.DispatchEx("Excel.Application")
    app.Visible          = False
    app.DisplayAlerts    = False
    app.ScreenUpdating   = False
    app.AskToUpdateLinks = False
    app.EnableEvents     = False

    return app, pythoncom


def cerrar_excel(app, wb, pythoncom):
    """Close everything in order and never raise (called from `finally`).
    A leaked EXCEL.EXE keeps the file locked and the next run fails confusingly."""
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        if app is not None:
            app.CutCopyMode = False
            app.ScreenUpdating = True
            app.EnableEvents = True
            app.Quit()
    except Exception:
        pass
    try:
        if pythoncom is not None:
            pythoncom.CoUninitialize()
    except Exception:
        pass


def hoja_por_nombre(wb, nombre):
    """Sheet by name: exact match first, loose (case/accent-insensitive) as fallback.
    The exact pass matters: two sheets may differ only by case or accents."""
    for ws in wb.Worksheets:
        if str(ws.Name) == nombre:
            return ws

    objetivo = norm_texto(nombre)
    for ws in wb.Worksheets:
        if norm_texto(ws.Name) == objetivo:
            return ws

    disponibles = [str(ws.Name) for ws in wb.Worksheets]
    raise ValueError(
        f"El libro no tiene una pestaña llamada '{nombre}'.\n\n"
        f"Pestañas encontradas: {disponibles}"
    )


# --- Quick check: python utils.py ---
if __name__ == "__main__":
    print("Hoy            :", claves_mes_actual_y_anterior())
    print("Corriendo jul26:", claves_mes_actual_y_anterior("M07-26"))
    print("Cruce de año   :", claves_mes_actual_y_anterior("M01-26"))
    print("Etiqueta       :", etiqueta_mes_es("M06-26"))
    print("Output (Input) :", carpeta_output_desde_input(r"C:\RDA\Input"))
    print("Output (otra)  :", carpeta_output_desde_input(r"C:\RDA\Reservas"))
