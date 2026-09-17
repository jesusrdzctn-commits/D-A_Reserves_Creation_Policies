"""
RDA Pólizas D&A — utils.py
===========================

Shared helpers with NO dependency on SAP or tkinter, so that:
  - the GUI can import them and still be openable/testable on its own, and
  - the consolidation layer can import them without duplicating logic.

This is the SINGLE SOURCE OF TRUTH for the month key used across the whole
process: the 'Mes' column of the 'Base' sheet is written as 'M06-26'
(M + 2-digit month + '-' + 2-digit year).

Golden rule: if the month key format ever changes, this file is touched and
nothing else.
"""

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
# Quick check:  python utils.py
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("Hoy            :", claves_mes_actual_y_anterior())
    print("Corriendo jul26:", claves_mes_actual_y_anterior("M07-26"))
    print("Cruce de año   :", claves_mes_actual_y_anterior("M01-26"))
    print("Etiqueta       :", etiqueta_mes_es("M06-26"))
