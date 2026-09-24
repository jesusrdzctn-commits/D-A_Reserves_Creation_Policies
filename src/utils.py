"""
RDA Pólizas D&A — utils.py
===========================

Shared helpers with NO dependency on SAP or tkinter, so that:
  - the GUI can import them and still be openable/testable on its own, and
  - the consolidation layers can import them without duplicating logic.

This is the SINGLE SOURCE OF TRUTH for the month key used across the whole
process: the 'Mes' column of the 'Base' sheet is written as 'M06-26'
(M + 2-digit month + '-' + 2-digit year).

Golden rule: if the month key format ever changes, this file is touched and
nothing else.

It is also where the Excel COM plumbing lives (abrir_excel / cerrar_excel /
hoja_por_nombre / norm_texto), for the same reason
carpeta_output_desde_input() moved here: TWO processes need it now
('PÓLIZA SAP' and 'Comparativos'), and importing one of them from the other
would drag a whole module along just to reuse four helpers. win32com is
imported INSIDE abrir_excel(), so this file still imports cleanly on a machine
with no Excel — only actually using it fails, and with a readable message.
"""

import os
import unicodedata
from datetime import datetime

# Month key as written in the 'Mes' column of the 'Base' sheet: 'M06-26'.
FORMATO_CLAVE_MES = "M%m-%y"

# Spanish month names, for status messages and popups the stakeholder reads.
MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre",
    12: "diciembre",
}


# ----------------------------------------------------------------------
# Month keys
# ----------------------------------------------------------------------
def clave_mes(fecha):
    """
    datetime -> month key as written in the 'Base' sheet.

        datetime(2026, 6, 15) -> 'M06-26'
    """
    return fecha.strftime(FORMATO_CLAVE_MES)


def parse_clave_mes(clave):
    """
    Month key -> datetime at the first day of that month. Raises ValueError on
    a malformed key, so a typo in an override fails loudly instead of silently
    producing an empty matrix.

        'M06-26' -> datetime(2026, 6, 1)
    """
    texto = str(clave).strip().upper()
    try:
        return datetime.strptime(texto, "M%m-%y")
    except ValueError:
        raise ValueError(
            f"La clave de mes '{clave}' no tiene el formato esperado 'M06-26' "
            f"(M + mes de 2 dígitos + '-' + año de 2 dígitos)."
        )


def mes_anterior(fecha):
    """
    datetime -> datetime at the first day of the PREVIOUS month.

    Walking back to day 1 first avoids the classic 31-day bug (March 31 minus
    one month is not February 31).

        datetime(2026, 1, 15) -> datetime(2025, 12, 1)
    """
    primero = fecha.replace(day=1)
    if primero.month == 1:
        return primero.replace(year=primero.year - 1, month=12)
    return primero.replace(month=primero.month - 1)


def claves_mes_actual_y_anterior(fecha=None):
    """
    Resolve the pair of month keys the 'Variaciones' sheet compares.

    Args:
        fecha (datetime|str|None): reference date. None -> today (the month the
            code is being run, which is the spec's default). A string is read as
            a month key ('M07-26'), which is how a past month is re-run.

    Returns:
        (clave_anterior, clave_actual) e.g. ('M06-26', 'M07-26')
    """
    if fecha is None:
        fecha = datetime.today()
    elif isinstance(fecha, str):
        fecha = parse_clave_mes(fecha)

    return clave_mes(mes_anterior(fecha)), clave_mes(fecha)


def etiqueta_mes_es(clave):
    """
    Month key -> human label for status messages. 'M06-26' -> 'junio 2026'.
    """
    fecha = parse_clave_mes(clave)
    return f"{MESES_ES[fecha.month]} {fecha.year}"


# ----------------------------------------------------------------------
# Output folder
# ----------------------------------------------------------------------
def carpeta_output_desde_input(ruta_input):
    """
    Return the 'Output' folder that sits next to the input folder.

    Same convention as Validación Factura Global: if the chosen folder is named
    'Input' the Output lands BESIDE it (src/Output); for any other folder the
    Output is created INSIDE it, so nothing is scattered outside the user's
    choice.

    Lives here rather than in a consolidation module because BOTH processes
    ('Variaciones' and 'PÓLIZA SAP') need it, and importing one consolidation
    module from the other would drag pandas and openpyxl into a process that
    only talks to Excel COM.
    """
    ruta = os.path.normpath(ruta_input)
    if os.path.basename(ruta).lower() == "input":
        return os.path.join(os.path.dirname(ruta), "Output")
    return os.path.join(ruta, "Output")


# ----------------------------------------------------------------------
# Excel COM plumbing (shared by Poliza_SAP.py and Comparativos.py)
# ----------------------------------------------------------------------
def norm_texto(texto):
    """Trim, upper-case and drop accents, tolerating None. 'Cebe' -> 'CEBE'."""
    if texto is None:
        return ""
    limpio = unicodedata.normalize("NFKD", str(texto))
    limpio = limpio.encode("ascii", "ignore").decode("ascii")
    return " ".join(limpio.upper().split())


def abrir_excel():
    """
    Start a PRIVATE Excel instance and return (app, pythoncom).

    DispatchEx, not Dispatch, and this is the single most important line in the
    COM layer. Dispatch attaches to the Excel the user already has open; when we
    later call app.Quit() we would be closing THEIR workbooks with our own.
    DispatchEx always starts a brand new, invisible instance that is ours to
    kill. Different kitchen, same recipe.
    """
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
    app.DisplayAlerts    = False   # no 'do you want to save?' pop-ups
    app.ScreenUpdating   = False   # much faster without redrawing
    app.AskToUpdateLinks = False   # external links never block the run
    app.EnableEvents     = False   # the workbook's own macros stay quiet

    return app, pythoncom


def cerrar_excel(app, wb, pythoncom):
    """
    Close everything, in order, never raising. Called from a `finally`, so if it
    blew up it must not blow up AGAIN and hide the real error. A leaked EXCEL.EXE
    keeps the file locked and the next run fails with a confusing message.
    """
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
    """
    Find a sheet by name: EXACT match first, loose match (case/accent
    insensitive) only as a fallback.

    The exact pass is not paranoia — the WHSL workbook really does carry both
    'catalogo' and 'Catálogo', and they are different tables. A loose-first
    match could grab the wrong one and every VLOOKUP would quietly return
    garbage instead of failing.
    """
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


# ----------------------------------------------------------------------
# Quick check:  python utils.py
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("Hoy            :", claves_mes_actual_y_anterior())
    print("Corriendo jul26:", claves_mes_actual_y_anterior("M07-26"))
    print("Cruce de año   :", claves_mes_actual_y_anterior("M01-26"))
    print("Etiqueta       :", etiqueta_mes_es("M06-26"))
    print("Output (Input) :", carpeta_output_desde_input(r"C:\RDA\Input"))
    print("Output (otra)  :", carpeta_output_desde_input(r"C:\RDA\Reservas"))
