"""RDA Pólizas D&A — Upload_SAP.py: opens the journal-upload transaction in SAP GUI and loads the 'Póliza Carga
Batch' file that step 3 generated (recorded script Z_G_FIN_TRN_JOURN). It stops BEFORE executing (F8): for now
the user reviews the screen and runs it by hand. win32com is imported inside the function, so this module
imports cleanly on a machine without SAP GUI.
"""

import os
import time

# One source of truth for the values the batch file was built with (Posting Company Code and Document Type).
from Create_Poliza_Batch import SOCIEDAD, TIPO_DOC

TRANSACCION = "Z_G_FIN_TRN_JOURN"
EXTENSIONES_BATCH = (".xlsx", ".xlsm")

# Component ids copied from the recording; they are system-specific.
ID_VENTANA, ID_DIALOGO = "wnd[0]", "wnd[1]"
ID_OKCODE = "wnd[0]/tbar[0]/okcd"
ID_RUTA = "wnd[1]/usr/ctxtDY_PATH"
ID_ACEPTAR = "wnd[1]/tbar[0]/btn[0]"
ID_SOCIEDAD = "wnd[0]/usr/ctxtP_BUKRS"
ID_CLASE_DOC = "wnd[0]/usr/ctxtP_BLART"
ID_PROCESAR = "wnd[0]/usr/radRB_PROC"
ID_BARRA = "wnd[0]/sbar"

VKEY_ENTER, VKEY_F4 = 0, 4

ESPERA_PANTALLA_SEG = 20
ESPERA_DIALOGO_SEG = 15
INTERVALO_SEG = 0.25

MSG_SIN_SAP = (
    "No fue posible conectarse a SAP GUI. "
    "Verifica que SAP esté abierto y con una sesión activa."
)


# --- Input ---
def _validar_archivo(ruta):
    """Absolute path of the batch file, or ValueError / FileNotFoundError saying what to fix."""
    if not ruta or not str(ruta).strip():
        raise ValueError("No hay una Póliza Carga Batch para subir: genérala primero en el paso 3.")
    ruta = os.path.abspath(str(ruta).strip())
    if not os.path.isfile(ruta):
        raise FileNotFoundError(
            f"No encontré la Póliza Carga Batch:\n{ruta}\n\n¿La movieron o la borraron?"
        )
    extension = os.path.splitext(ruta)[1].lower()
    if extension not in EXTENSIONES_BATCH:
        raise ValueError(
            f"La Póliza Carga Batch debe ser {' o '.join(EXTENSIONES_BATCH)} y recibí '{extension}'."
        )
    # Excel holds an open workbook with a write lock: if we cannot open it for writing, it is in use.
    try:
        with open(ruta, "r+b"):
            pass
    except PermissionError:
        raise ValueError(
            f"No puedo abrir '{os.path.basename(ruta)}' con permiso de escritura.\n\n"
            f"¿Sigue abierto en Excel? Ciérralo para que SAP lea el archivo ya guardado."
        )
    return ruta


# --- SAP GUI helpers ---
def _existe(sesion, id_componente):
    """True if the component exists (findById raises when it does not)."""
    try:
        sesion.findById(id_componente)
        return True
    except Exception:
        return False


def _esperar(sesion, id_componente, timeout):
    """Wait until the component exists; False if it still does not after `timeout` seconds."""
    limite = time.monotonic() + timeout
    while not _existe(sesion, id_componente):
        if time.monotonic() >= limite:
            return False
        time.sleep(INTERVALO_SEG)
    return True


def _esperar_cierre(sesion, id_componente, timeout):
    """Wait until the component is gone; False if it is still there after `timeout` seconds."""
    limite = time.monotonic() + timeout
    while _existe(sesion, id_componente):
        if time.monotonic() >= limite:
            return False
        time.sleep(INTERVALO_SEG)
    return True


def _esperar_libre(sesion, timeout):
    """Wait while SAP is busy with this window. Never raises."""
    limite = time.monotonic() + timeout
    try:
        while sesion.Busy and time.monotonic() < limite:
            time.sleep(INTERVALO_SEG)
    except Exception:
        pass


def _barra(sesion):
    """(message type, text) of the status bar; ('', '') when it cannot be read. Types: E error, A abort..."""
    try:
        barra = sesion.findById(ID_BARRA)
        return str(barra.MessageType or "").strip().upper(), str(barra.Text or "").strip()
    except Exception:
        return "", ""


def _describir_ventana(sesion, ventana=ID_DIALOGO, maximo=300):
    """'<title>: text | text' of a popup, for error messages. Never raises."""
    try:
        titulo = str(sesion.findById(ventana).Text).strip()
    except Exception:
        titulo = ""
    textos = []
    try:
        area = sesion.findById(f"{ventana}/usr")
        for i in range(min(int(area.Children.Count), 60)):
            try:
                texto = str(area.Children(i).Text).strip()
            except Exception:
                continue
            if texto:
                textos.append(texto)
    except Exception:
        pass
    detalle = " | ".join(textos)
    descripcion = f"{titulo}: {detalle}" if titulo and detalle else (titulo or detalle)
    return descripcion[:maximo] if descripcion else "(sin texto)"


def _estado_sap(sesion):
    """Status bar and popup (if any) as one line for an error message."""
    _, texto = _barra(sesion)
    partes = [f"Barra de estado: {texto}"] if texto else []
    if _existe(sesion, ID_DIALOGO):
        partes.append(f"Ventana abierta: {_describir_ventana(sesion)}")
    return "\n".join(partes) if partes else "SAP no mostró ningún mensaje."


def _mismo_texto(a, b):
    """SAP may change case or pad spaces on a field, so compare ignoring both."""
    return str(a).strip().lower() == str(b).strip().lower()


def _sesion_sap(win32):
    """First window of the first SAP connection (the one the user works in), or RuntimeError."""
    try:
        aplicacion = win32.GetObject("SAPGUI").GetScriptingEngine
        return aplicacion.Children(0).Children(0)
    except Exception as e:
        raise RuntimeError(MSG_SIN_SAP) from e


# --- The recorded script ---
def _cargar_en_sap(sesion, ruta, update_status):
    """Run the recorded steps on `sesion` and return the status-bar text. Every field is read back, and
    anything unexpected stops with a RuntimeError that names the step. F8 (execute) is NOT sent."""
    paso = "revisar la sesión de SAP"
    try:
        if _existe(sesion, ID_DIALOGO):
            raise RuntimeError(
                "SAP tiene una ventana emergente abierta; ciérrala y vuelve a intentar.\n\n"
                + _estado_sap(sesion)
            )

        paso = "abrir la transacción"
        update_status("🔌 Abriendo la transacción en SAP...")
        sesion.findById(ID_VENTANA).maximize()
        # '/n' ends any transaction left open; from the main menu it is the same as typing the code.
        sesion.findById(ID_OKCODE).text = "/n" + TRANSACCION
        sesion.findById(ID_VENTANA).sendVKey(VKEY_ENTER)
        if _existe(sesion, ID_DIALOGO):
            raise RuntimeError(
                f"SAP mostró una ventana al abrir {TRANSACCION}; resuélvela en SAP y vuelve a intentar.\n\n"
                f"{_estado_sap(sesion)}"
            )
        if not _esperar(sesion, ID_SOCIEDAD, ESPERA_PANTALLA_SEG):
            raise RuntimeError(
                f"La transacción {TRANSACCION} no abrió su pantalla de selección.\n\n{_estado_sap(sesion)}"
            )

        paso = "abrir el selector de archivo (F4)"
        update_status("📂 Eligiendo el archivo en SAP...")
        sesion.findById(ID_VENTANA).sendVKey(VKEY_F4)
        if not _esperar(sesion, ID_RUTA, ESPERA_DIALOGO_SEG):
            raise RuntimeError(
                "SAP no abrió el selector de archivo con F4 (el cursor debe quedar en el campo del "
                f"archivo al abrir la transacción).\n\n{_estado_sap(sesion)}"
            )

        paso = "escribir la ruta del archivo"
        campo = sesion.findById(ID_RUTA)
        campo.text = ruta
        campo.setFocus()
        if not _mismo_texto(campo.text, ruta):
            raise RuntimeError(
                f"SAP no aceptó la ruta completa del archivo.\n\nEscribí:  {ruta}\nQuedó:    {campo.text}\n\n"
                f"Prueba con una carpeta de ruta más corta."
            )
        sesion.findById(ID_ACEPTAR).press()
        if not _esperar_cierre(sesion, ID_DIALOGO, ESPERA_DIALOGO_SEG):
            raise RuntimeError(f"SAP no aceptó el archivo.\n\n{_estado_sap(sesion)}")
        _esperar_libre(sesion, ESPERA_PANTALLA_SEG)
        tipo, texto = _barra(sesion)
        if tipo in ("E", "A"):
            raise RuntimeError(f"SAP rechazó el archivo: {texto}")

        paso = "llenar la pantalla de selección"
        update_status("📝 Llenando sociedad y clase de documento...")
        for id_campo, valor, nombre in ((ID_SOCIEDAD, SOCIEDAD, "la sociedad"),
                                        (ID_CLASE_DOC, TIPO_DOC, "la clase de documento")):
            campo = sesion.findById(id_campo)
            campo.text = valor
            if not _mismo_texto(campo.text, valor):
                raise RuntimeError(f"SAP no dejó {nombre} en '{valor}' (quedó '{campo.text}').")
        radio = sesion.findById(ID_PROCESAR)
        radio.setFocus()
        radio.select()
        if not radio.Selected:
            raise RuntimeError("No pude seleccionar la opción de procesar en SAP.")

        # <- the F8 (execute) click goes here once the recording of that last step is shared.
        _esperar_libre(sesion, ESPERA_PANTALLA_SEG)
        return _barra(sesion)[1]

    except RuntimeError:
        raise
    except Exception as e:
        # Typically a com_error: a component that does not exist or SAP closing under us.
        raise RuntimeError(
            f"SAP falló al {paso}:\n\n{e}\n\n{_estado_sap(sesion)}"
        ) from e


# --- Main entry point ---
def preparar_subida_sap(ruta_batch, callback_status=None):
    """Open the upload transaction in the user's SAP session and load `ruta_batch` (the 'Póliza Carga
    Batch' of step 3), filling company code, document type and the process option. Does NOT execute.
    Returns: ruta, transaccion, sociedad, clase_documento, ejecutado (always False for now), aviso_sap.
    Raises FileNotFoundError / ValueError (fixable input) or RuntimeError (SAP GUI / pywin32 problems)."""
    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    ruta = _validar_archivo(ruta_batch)

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        raise RuntimeError(
            "Este proceso necesita 'pywin32' y SAP GUI instalado.\n\n"
            "Instálalo con:   pip install pywin32"
        )

    update_status("🔌 Conectando con SAP GUI...")
    pythoncom.CoInitialize()
    sesion = None
    try:
        sesion = _sesion_sap(win32)
        aviso = _cargar_en_sap(sesion, ruta, update_status)
    finally:
        # Release the COM proxies before COM goes away.
        sesion = None
        pythoncom.CoUninitialize()

    return {
        "ruta":            ruta,
        "transaccion":     TRANSACCION,
        "sociedad":        SOCIEDAD,
        "clase_documento": TIPO_DOC,
        "ejecutado":       False,
        "aviso_sap":       aviso,
    }


# --- Standalone run: python Upload_SAP.py ---
if __name__ == "__main__":
    print("=== Prueba de la subida a SAP (standalone, sin F8) ===")
    archivo = input("Ruta de la Póliza Carga Batch ya generada (.xlsx): ").strip().strip('"')
    try:
        res = preparar_subida_sap(archivo, callback_status=print)
        print("\n🎉 LISTO (falta ejecutar F8 en SAP) 🎉")
        for clave, valor in res.items():
            print(f"  {clave:<16}: {valor}")
    except Exception as e:
        print(f"❌ Error preparando la subida a SAP: {e}")
        raise
