"""
RDA Pólizas D&A — interfaz_GUI.py
=================================

Minimal tkinter GUI for the first demo to the stakeholder: pick the Reservas
workbook, press Consolidar, get the 'Variaciones' sheet.

Same skeleton and colour scheme as Validación Factura Global, so growing this
into the full app later is adding widgets, not rewriting the file.

No SAP and no pandas imports here on purpose: this file can be opened and
clicked through on any machine with `python interfaz_GUI.py`, even one that
cannot reach SAP.
"""

import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

from utils import claves_mes_actual_y_anterior, etiqueta_mes_es


def _ruta_config():
    """
    Where config.json lives, so the app remembers the last workbook used.

    Next to the executable when frozen by PyInstaller, next to this file when
    run as a script. That way each machine keeps its own path without anybody
    touching the code.
    """
    if getattr(sys, "frozen", False):        # running as .exe (PyInstaller)
        base = os.path.dirname(sys.executable)
    else:                                    # running as a .py script
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "config.json")


class ReservasGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("RDA Pólizas D&A")
        self.root.geometry("640x420")
        self.root.resizable(False, False)

        # Colour scheme (same as Validación Factura Global)
        self.bg_color        = "#FFFFFF"
        self.primary_color   = "#00094F"
        self.secondary_color = "#333333"
        self.light_gray      = "#F5F5F5"
        self.border_color    = "#E0E0E0"
        self.success_color   = "#28A745"

        self.root.configure(bg=self.bg_color)

        # Callback the controller plugs in.
        self.on_consolidar = None

        self._config_file = _ruta_config()
        cfg = self._cargar_config()

        self.archivo_var = tk.StringVar(value=(cfg.get("ultimo_archivo") or "").strip())
        # Empty = the month the code is being run (the spec's default). Typing
        # a key such as M07-26 re-runs a past month, which is also how the demo
        # is driven when the test file holds older months.
        self.mes_var = tk.StringVar(value="")

        self._build_ui()
        self._actualizar_preview_mes()

    # ------------------------------------------------------------------
    # config.json
    # ------------------------------------------------------------------
    def _cargar_config(self):
        try:
            with open(self._config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (FileNotFoundError, json.JSONDecodeError, OSError, ValueError):
            return {}

    def _guardar_config(self):
        try:
            with open(self._config_file, "w", encoding="utf-8") as f:
                json.dump(
                    {"ultimo_archivo": self.archivo_var.get().strip()},
                    f, ensure_ascii=False, indent=2,
                )
        except OSError as e:
            # Not critical: the app keeps working without a saved path.
            print(f"[CONFIG] No se pudo guardar la configuración: {e}")

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        main = tk.Frame(self.root, bg=self.bg_color, padx=24, pady=18)
        main.pack(fill="both", expand=True)

        # --- Header ---
        tk.Label(
            main, text="RDA Pólizas D&A",
            font=("Segoe UI", 16, "bold"),
            bg=self.bg_color, fg=self.primary_color,
        ).pack(pady=(0, 4))
        tk.Label(
            main, text="Consolidación de Creación de Reservas PMF OT — pestaña 'Variaciones'",
            font=("Segoe UI", 10),
            bg=self.bg_color, fg=self.secondary_color,
        ).pack(pady=(0, 16))

        # --- Workbook picker ---
        archivo_frame = tk.LabelFrame(
            main, text="  Archivo de Creación de Reservas PMF OT  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=12,
        )
        archivo_frame.pack(fill="x", pady=(0, 12))

        tk.Entry(
            archivo_frame, textvariable=self.archivo_var,
            font=("Segoe UI", 9), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        ).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 8), ipady=4)

        tk.Button(
            archivo_frame, text="Examinar...", font=("Segoe UI", 9, "bold"),
            bg=self.primary_color, fg=self.bg_color, relief=tk.FLAT,
            cursor="hand2", padx=12, command=self._elegir_archivo,
        ).pack(side=tk.LEFT)

        # --- Month ---
        # Leaving it empty is the normal case. The field exists because the
        # matrices compare the previous month against the current one, and a
        # demo (or a re-run of a closed month) needs to point at a past month.
        mes_frame = tk.LabelFrame(
            main, text="  Mes a procesar  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=12,
        )
        mes_frame.pack(fill="x", pady=(0, 12))

        tk.Entry(
            mes_frame, textvariable=self.mes_var, width=12,
            font=("Segoe UI", 10), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        ).pack(side=tk.LEFT, padx=(0, 10), ipady=3)

        self.preview_mes_var = tk.StringVar()
        tk.Label(
            mes_frame, textvariable=self.preview_mes_var,
            font=("Segoe UI", 9), bg=self.bg_color, fg=self.secondary_color,
            anchor="w", justify="left",
        ).pack(side=tk.LEFT, fill="x", expand=True)

        self.mes_var.trace_add("write", lambda *_: self._actualizar_preview_mes())

        # --- Action button ---
        self.consolidar_btn = tk.Button(
            main, text="🧩 Consolidar...!",
            font=("Segoe UI", 11, "bold"),
            bg=self.success_color, fg=self.bg_color,
            relief=tk.RAISED, bd=2, pady=12, cursor="hand2",
            activebackground=self.secondary_color, activeforeground=self.bg_color,
            command=self._handle_consolidar,
        )
        self.consolidar_btn.pack(fill="x", pady=(6, 0))

        # --- Status bar ---
        self.status_var = tk.StringVar(value="✓ Listo para comenzar")
        tk.Label(
            main, textvariable=self.status_var, font=("Segoe UI", 10),
            bg=self.bg_color, fg="#666666", wraplength=580,
            anchor="w", justify="left",
        ).pack(fill="x", pady=(14, 0))

    # ------------------------------------------------------------------
    # Handlers
    # ------------------------------------------------------------------
    def _elegir_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo de Creación de Reservas PMF OT",
            initialdir=(
                os.path.dirname(self.archivo_var.get())
                or os.path.expanduser("~")
            ),
            filetypes=[("Libros de Excel", "*.xlsx *.xlsm"),
                       ("Todos los archivos", "*.*")],
        )
        if ruta:  # cancel leaves the current selection untouched
            self.archivo_var.set(os.path.normpath(ruta))
            self._guardar_config()

    def _actualizar_preview_mes(self):
        """Live hint of which two months the run will compare."""
        try:
            anterior, actual = claves_mes_actual_y_anterior(self.get_mes())
            self.preview_mes_var.set(
                f"Comparará {anterior} vs {actual}\n"
                f"({etiqueta_mes_es(anterior)} vs {etiqueta_mes_es(actual)})"
            )
        except ValueError:
            self.preview_mes_var.set("⚠️ Formato inválido — se espera M07-26")

    def _handle_consolidar(self):
        if not self.validate():
            return
        if self.on_consolidar:
            self.on_consolidar()
        else:
            messagebox.showinfo(
                "Info",
                f"Funcionalidad no conectada\n\nArchivo: {self.get_archivo()}",
            )

    # ------------------------------------------------------------------
    # Validation and accessors used by the controller
    # ------------------------------------------------------------------
    def validate(self):
        ruta = self.archivo_var.get().strip()
        if not ruta:
            messagebox.showerror(
                "Falta el archivo",
                "Selecciona el archivo de Creación de Reservas PMF OT con el "
                "botón 'Examinar...'.",
            )
            return False
        if not os.path.exists(ruta):
            messagebox.showerror(
                "Archivo no encontrado",
                f"No encontré el archivo:\n\n{ruta}\n\n"
                f"¿Lo movieron o le cambiaron el nombre?",
            )
            return False

        try:
            claves_mes_actual_y_anterior(self.get_mes())
        except ValueError as e:
            messagebox.showerror("Mes inválido", str(e))
            return False

        self._guardar_config()
        return True

    def get_archivo(self):
        return self.archivo_var.get().strip()

    def get_mes(self):
        """Month key typed by the user, or None to use today's month."""
        return self.mes_var.get().strip().upper() or None

    def set_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()

    def disable_buttons(self):
        self.consolidar_btn.config(state="disabled", bg="#666666")

    def enable_buttons(self):
        self.consolidar_btn.config(state="normal", bg=self.success_color)


def main():
    # Standalone entry point, to click through the GUI without the controller.
    root = tk.Tk()
    ReservasGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
