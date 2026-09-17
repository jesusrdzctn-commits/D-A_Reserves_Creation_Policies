"""
RDA Pólizas D&A — controller.py
================================

Wires the GUI to the consolidation logic. Same shape as Validación Factura
Global's controller: confirm, run, summarise in a popup, always re-enable the
buttons in a `finally`.

Today it owns a single action ('Consolidar'). The SAP upload will hang off the
same class once that half of the process is defined.
"""

import os
from tkinter import messagebox

from Consolidacion import HOJA_VARIACIONES, construir_variaciones
from utils import claves_mes_actual_y_anterior, etiqueta_mes_es

# How many unmatched concepts to list inside the popup before trimming, so the
# message box never grows taller than the screen.
MAX_NO_MAPEADOS_EN_POPUP = 8


class ReservasController:
    """Connects the GUI with the 'Variaciones' consolidation."""

    def __init__(self, gui):
        self.gui = gui
        self.gui.on_consolidar = self.execute_consolidacion

    # =================================================================
    # CONSOLIDACIÓN: builds the 'Variaciones' sheet
    # =================================================================
    def execute_consolidacion(self):
        if not self.gui.validate():
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
            resumen = construir_variaciones(
                ruta_libro,
                mes=mes,
                callback_status=self.gui.set_status,
            )

            self.gui.set_status("✅ ¡Consolidación completada!")
            messagebox.showinfo(
                "Consolidación completada", self._texto_resumen(resumen)
            )

        except (FileNotFoundError, ValueError) as e:
            # Expected, actionable problems: missing file, missing 'Base' sheet,
            # missing column, unparseable amount. The message already explains
            # what to fix, so it is shown as-is.
            self.gui.set_status("❌ Error en la consolidación")
            messagebox.showerror("Error en la consolidación", str(e))
        except PermissionError:
            self.gui.set_status("❌ El archivo está abierto en Excel")
            messagebox.showerror(
                "Archivo en uso",
                "No pude escribir el resultado porque el archivo está abierto "
                "en Excel.\n\nCiérralo y vuelve a intentar.",
            )
        except Exception as e:
            self.gui.set_status("❌ Error en la consolidación")
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n\n{e}")
        finally:
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")

    # =================================================================
    # Summary popup
    # =================================================================
    @staticmethod
    def _texto_resumen(resumen):
        """
        Build the popup text.

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
