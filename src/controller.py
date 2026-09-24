"""
RDA Pólizas D&A — controller.py
================================

Wires the GUI to the two consolidation modules. Same shape as Validación
Factura Global's controller: confirm, run, summarise in a popup, always
re-enable the buttons in a `finally`.

Three actions:

    execute_consolidacion()  archivo 1 -> pestaña 'Variaciones'   (openpyxl)
    execute_poliza_sap()     archivo 2 -> sección 'PÓLIZA SAP'    (Excel COM)
    execute_todo()           both, in order, with a combined report

The two processes never share state: different files, different engines,
different failures. 'Ejecutar todo' just calls them one after the other and
keeps the score.

The SAP upload will hang off this same class once that half of the process is
defined.
"""

import os
from tkinter import messagebox

from Consolidacion import HOJA_VARIACIONES, construir_variaciones
from Poliza_SAP import HOJA_CREACIONES, generar_poliza_sap
from utils import claves_mes_actual_y_anterior, etiqueta_mes_es

# How many unmatched concepts to list inside the popup before trimming, so the
# message box never grows taller than the screen.
MAX_NO_MAPEADOS_EN_POPUP = 8


class ReservasController:
    """Connects the GUI with the two consolidations."""

    def __init__(self, gui):
        self.gui = gui
        self.gui.on_consolidar = self.execute_consolidacion
        self.gui.on_poliza     = self.execute_poliza_sap
        self.gui.on_todo       = self.execute_todo

    # =================================================================
    # Shared runner
    # =================================================================
    def _ejecutar(self, funcion):
        """
        Run one consolidation and translate its exceptions into (ok, payload).

        Returns (True, resumen) or (False, mensaje_de_error). It does NOT touch
        the buttons or show popups: the caller decides whether a failure means
        'stop everything' or 'note it and carry on with step 2'. That is the
        whole reason this is factored out — 'Ejecutar todo' needs to survive a
        failed step.
        """
        try:
            return True, funcion()

        except (FileNotFoundError, ValueError) as e:
            # Expected, actionable problems: missing file, missing sheet,
            # missing column, unparseable amount. The message already explains
            # what to fix, so it is passed through as-is.
            return False, str(e)

        except PermissionError:
            return False, (
                "No pude escribir el resultado porque el archivo está abierto "
                "en Excel.\n\nCiérralo y vuelve a intentar."
            )

        except RuntimeError as e:
            # Raised by Poliza_SAP when pywin32 / Excel are not available.
            return False, str(e)

        except Exception as e:
            return False, f"Ocurrió un error inesperado:\n\n{e}"

    def _finalizar(self):
        """Re-enable the buttons and reset the status line if nothing succeeded."""
        self.gui.enable_buttons()
        if "completad" not in self.gui.status_var.get():
            self.gui.set_status("✓ Listo para comenzar")

    # =================================================================
    # PROCESO 1 — pestaña 'Variaciones'
    # =================================================================
    def execute_consolidacion(self):
        if not self.gui.validate_variaciones():
            return

        ruta_libro = self.gui.get_archivo()
        mes        = self.gui.get_mes()
        anterior, actual = claves_mes_actual_y_anterior(mes)

        confirm = messagebox.askyesno(
            "Confirmar consolidación",
            f"¿Generar la pestaña '{HOJA_VARIACIONES}'?\n\n"
            f"  📄 Archivo : {os.path.basename(ruta_libro)}\n"
            f"  📅 Compara : {anterior} vs {actual}\n"
            f"               ({etiqueta_mes_es(anterior)} vs {etiqueta_mes_es(actual)})\n\n"
            f"Se trabajará sobre una COPIA en la subcarpeta 'Output',\n"
            f"así que el archivo original no se toca."
        )
        if not confirm:
            return

        self.gui.disable_buttons()
        try:
            ok, payload = self._ejecutar(lambda: construir_variaciones(
                ruta_libro, mes=mes, callback_status=self.gui.set_status,
            ))
            if ok:
                self.gui.set_status("✅ ¡Consolidación completada!")
                messagebox.showinfo(
                    "Consolidación completada",
                    self._texto_variaciones(payload),
                )
            else:
                self.gui.set_status("❌ Error en la consolidación")
                messagebox.showerror("Error en la consolidación", payload)
        finally:
            self._finalizar()

    # =================================================================
    # PROCESO 2 — sección 'PÓLIZA SAP'
    # =================================================================
    def execute_poliza_sap(self):
        if not self.gui.validate_poliza():
            return

        ruta_libro = self.gui.get_archivo_poliza()

        confirm = messagebox.askyesno(
            "Confirmar póliza SAP",
            f"¿Llenar la sección 'PÓLIZA SAP'?\n\n"
            f"  📄 Archivo : {os.path.basename(ruta_libro)}\n"
            f"  📑 Pestaña : {HOJA_CREACIONES}\n"
            f"  ✍️  Columnas: R, T, U, V, W, X, Z\n\n"
            f"Se abrirá Excel en segundo plano y se trabajará sobre una\n"
            f"COPIA en la subcarpeta 'Output'. El original no se toca."
        )
        if not confirm:
            return

        self.gui.disable_buttons()
        try:
            ok, payload = self._ejecutar(lambda: generar_poliza_sap(
                ruta_libro, callback_status=self.gui.set_status,
            ))
            if ok:
                self.gui.set_status("✅ ¡Póliza SAP completada!")
                messagebox.showinfo(
                    "Póliza SAP completada", self._texto_poliza(payload)
                )
            else:
                self.gui.set_status("❌ Error generando la póliza SAP")
                messagebox.showerror("Error generando la póliza SAP", payload)
        finally:
            self._finalizar()

    # =================================================================
    # Both, in order
    # =================================================================
    def execute_todo(self):
        if not self.gui.validate_todo():
            return

        ruta_var = self.gui.get_archivo()
        ruta_pol = self.gui.get_archivo_poliza()
        mes      = self.gui.get_mes()
        anterior, actual = claves_mes_actual_y_anterior(mes)

        confirm = messagebox.askyesno(
            "Confirmar ejecución completa",
            f"¿Correr los DOS procesos?\n\n"
            f"  1️⃣  '{HOJA_VARIACIONES}'\n"
            f"      📄 {os.path.basename(ruta_var)}\n"
            f"      📅 {anterior} vs {actual}\n\n"
            f"  2️⃣  'PÓLIZA SAP'\n"
            f"      📄 {os.path.basename(ruta_pol)}\n\n"
            f"Ambos escriben sobre COPIAS en la subcarpeta 'Output'."
        )
        if not confirm:
            return

        self.gui.disable_buttons()
        try:
            # --- Paso 1 ---------------------------------------------------
            self.gui.set_status("1️⃣  Consolidando 'Variaciones'...")
            ok1, payload1 = self._ejecutar(lambda: construir_variaciones(
                ruta_var, mes=mes, callback_status=self.gui.set_status,
            ))

            # A failed step 1 does not have to kill step 2: the two files are
            # independent. But it is the user's call, not the code's — silently
            # carrying on would hide half a month-end.
            if not ok1:
                seguir = messagebox.askyesno(
                    "Falló el paso 1",
                    f"La consolidación de '{HOJA_VARIACIONES}' falló:\n\n"
                    f"{payload1}\n\n"
                    f"¿Continúo con el paso 2 ('PÓLIZA SAP')?",
                )
                if not seguir:
                    self.gui.set_status("❌ Ejecución detenida en el paso 1")
                    return

            # --- Paso 2 ---------------------------------------------------
            self.gui.set_status("2️⃣  Generando la 'PÓLIZA SAP'...")
            ok2, payload2 = self._ejecutar(lambda: generar_poliza_sap(
                ruta_pol, callback_status=self.gui.set_status,
            ))

            # --- Combined report ------------------------------------------
            partes = []
            partes.append(
                f"1️⃣  {HOJA_VARIACIONES}  —  {'✅ OK' if ok1 else '❌ FALLÓ'}\n"
                + (self._texto_variaciones(payload1) if ok1 else payload1)
            )
            partes.append(
                f"2️⃣  PÓLIZA SAP  —  {'✅ OK' if ok2 else '❌ FALLÓ'}\n"
                + (self._texto_poliza(payload2) if ok2 else payload2)
            )
            texto = ("\n" + "─" * 48 + "\n").join(partes)

            if ok1 and ok2:
                self.gui.set_status("✅ ¡Ejecución completa completada!")
                messagebox.showinfo("Ejecución completa", texto)
            elif ok1 or ok2:
                self.gui.set_status("⚠️ Ejecución completada a medias")
                messagebox.showwarning("Ejecución parcial", texto)
            else:
                self.gui.set_status("❌ Fallaron los dos procesos")
                messagebox.showerror("Ejecución fallida", texto)
        finally:
            self._finalizar()

    # =================================================================
    # Summary popups
    # =================================================================
    @staticmethod
    def _texto_variaciones(resumen):
        """
        Build the 'Variaciones' popup text.

        The unmatched-concepts warning is deliberately loud: those rows DID make
        it into the sheet under their raw name, so the Grand Total is right, but
        somebody has to decide whether they belong in the catalogue.
        """
        bloques = "\n".join(
            f"     • {titulo}: {n} concepto(s)"
            for titulo, n in resumen["conceptos_por_bloque"].items()
        )

        texto = (
            f"Archivo generado:\n{resumen['ruta']}\n\n"
            f"  • Meses comparados : {resumen['mes_anterior']} vs {resumen['mes_actual']}\n"
            f"  • Filas leídas     : {resumen['filas_base']:,}\n"
            f"  • Matrices         : {resumen['bloques']}\n"
            f"{bloques}\n"
            f"  • Σ {resumen['mes_anterior']}        : {resumen['total_anterior']:,.2f}\n"
            f"  • Σ {resumen['mes_actual']}        : {resumen['total_actual']:,.2f}"
        )

        no_mapeados = resumen.get("conceptos_no_mapeados") or []
        if no_mapeados:
            mostrados = no_mapeados[:MAX_NO_MAPEADOS_EN_POPUP]
            resto = len(no_mapeados) - len(mostrados)
            texto += (
                f"\n\n⚠️ {len(no_mapeados)} valor(es) de 'Agrupador' no empataron "
                f"con el catálogo de conceptos.\n"
                f"Salieron como renglón propio al final de su matriz, con su "
                f"nombre original (los totales sí cuadran):\n"
                + "\n".join(f"     • {c}" for c in mostrados)
                + (f"\n     • ...y {resto} más" if resto > 0 else "")
            )

        sin_bloque = resumen.get("analistas_sin_bloque") or []
        if sin_bloque:
            texto += (
                f"\n\n⚠️ Analista(s) en la Base sin matriz asignada: "
                f"{', '.join(sin_bloque)}.\n"
                f"Sus importes NO aparecen en ninguna matriz."
            )

        return texto

    @staticmethod
    def _texto_poliza(resumen):
        """
        Build the 'PÓLIZA SAP' popup text.

        The #N/A warning is the important half. Those rows are going up to SAP:
        a failed VLOOKUP means an account, a cost centre or a description that
        the 'catalogo' sheet does not know about, and it has to be fixed BEFORE
        the upload, not discovered by SAP.
        """
        columnas = "\n".join(f"     • {c}" for c in resumen["columnas"])

        texto = (
            f"Archivo generado:\n{resumen['ruta']}\n\n"
            f"  • Pestaña          : {resumen['hoja']}\n"
            f"  • Encabezado en    : fila {resumen['fila_encabezado']}\n"
            f"  • Filas escritas   : {resumen['filas']:,} "
            f"(filas {resumen['fila_inicio']}–{resumen['fila_fin']})\n"
            f"  • Columnas         :\n{columnas}\n"
            f"  • Σ columna {resumen['columna_control']} (MONTO) : "
            f"{resumen['total_control']:,.2f}\n"
            f"  • Fórmulas         : "
            + ("congeladas a valores" if resumen["convertido_a_valores"]
               else "vivas (se recalculan solas)")
        )

        errores = resumen.get("errores") or []
        if errores:
            texto += (
                f"\n\n⚠️ {resumen['total_errores']} celda(s) con error en los "
                f"VLOOKUP — hay que corregirlas ANTES de subir a SAP:\n"
            )
            for e in errores:
                celdas = ", ".join(f"{c['celda']} {c['error']}" for c in e["celdas"])
                resto = e["total"] - len(e["celdas"])
                texto += (
                    f"     • {e['columna']} ({e['titulo']}): {e['total']} error(es)\n"
                    f"       {celdas}"
                    + (f", ...y {resto} más" if resto > 0 else "")
                    + "\n"
                )
            texto += (
                "\nNormalmente significa que el valor no existe en la pestaña "
                "'catalogo'."
            )
        else:
            texto += "\n\n✅ Sin errores: todos los VLOOKUP empataron."

        return texto
