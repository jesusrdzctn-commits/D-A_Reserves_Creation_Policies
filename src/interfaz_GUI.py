"""
RDA Pólizas D&A — interfaz_GUI.py
=================================

tkinter GUI for the two consolidations of the project:

    Archivo 1  '...Creacion de Reservas PMF OT.xlsx'   -> pestaña 'Variaciones'
    Archivo 2  'Layout Creación Reservas DA WHSL.xlsb' -> sección 'PÓLIZA SAP'

Each file owns its picker and its own button, because they are separate inputs
with separate failures: if the .xlsb is open in Excel, the 'Variaciones' run
must still work. A third button runs both in sequence for the normal month-end.

Same skeleton and colour scheme as Validación Factura Global.

No SAP and no pandas imports here on purpose: this file can be opened and
clicked through on any machine with `python interfaz_GUI.py`, even one that
cannot reach SAP — or, now, one without Excel installed.
"""

import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

from utils import claves_mes_actual_y_anterior, etiqueta_mes_es

# Which extensions each picker offers. The WHSL layout is .xlsb (binary
# workbook), so it needs its own filter — it would not show up under '*.xlsx'.
FILTRO_VARIACIONES = [("Libros de Excel", "*.xlsx *.xlsm"),
                      ("Todos los archivos", "*.*")]
FILTRO_POLIZA      = [("Libros de Excel", "*.xlsb *.xlsx *.xlsm"),
                      ("Todos los archivos", "*.*")]


def _ruta_config():
    """
    Where config.json lives, so the app remembers the last workbooks used.

    Next to the executable when frozen by PyInstaller, next to this file when
    run as a script. That way each machine keeps its own paths without anybody
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
        self.root.geometry("680x600")
        self.root.resizable(False, False)

        # Colour scheme (same as Validación Factura Global)
        self.bg_color        = "#FFFFFF"
        self.primary_color   = "#00094F"
        self.secondary_color = "#333333"
        self.light_gray      = "#F5F5F5"
        self.border_color    = "#E0E0E0"
        self.success_color   = "#28A745"
        self.accent_color    = "#0067B1"   # second process
        self.todo_color      = "#00094F"   # run-everything button

        self.root.configure(bg=self.bg_color)

        # Callbacks the controller plugs in.
        self.on_consolidar = None   # archivo 1 -> 'Variaciones'
        self.on_poliza     = None   # archivo 2 -> 'PÓLIZA SAP'
        self.on_todo       = None   # both, in order

        self._config_file = _ruta_config()
        cfg = self._cargar_config()

        # 'ultimo_archivo' is the key the single-file version wrote. Reading it
        # as a fallback means an existing config.json out in the wild keeps
        # working instead of coming back empty.
        self.archivo_var = tk.StringVar(
            value=(cfg.get("ultimo_archivo_variaciones")
                   or cfg.get("ultimo_archivo") or "").strip()
        )
        self.archivo_poliza_var = tk.StringVar(
            value=(cfg.get("ultimo_archivo_poliza") or "").strip()
        )
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
                    {
                        "ultimo_archivo_variaciones": self.archivo_var.get().strip(),
                        "ultimo_archivo_poliza": self.archivo_poliza_var.get().strip(),
                    },
                    f, ensure_ascii=False, indent=2,
                )
        except OSError as e:
            # Not critical: the app keeps working without saved paths.
            print(f"[CONFIG] No se pudo guardar la configuración: {e}")

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        main = tk.Frame(self.root, bg=self.bg_color, padx=24, pady=16)
        main.pack(fill="both", expand=True)

        # --- Header ---
        tk.Label(
            main, text="RDA Pólizas D&A",
            font=("Segoe UI", 16, "bold"),
            bg=self.bg_color, fg=self.primary_color,
        ).pack(pady=(0, 2))
        tk.Label(
            main, text="Consolidación de los dos archivos de Creación de Reservas",
            font=("Segoe UI", 10),
            bg=self.bg_color, fg=self.secondary_color,
        ).pack(pady=(0, 14))

        # ==============================================================
        # PROCESO 1 — 'Variaciones'
        # ==============================================================
        frame1 = tk.LabelFrame(
            main, text="  1. Creación de Reservas PMF OT  →  pestaña 'Variaciones'  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=14, pady=10,
        )
        frame1.pack(fill="x", pady=(0, 10))

        self._fila_archivo(frame1, self.archivo_var, self._elegir_archivo)

        # Month only belongs to this process: the matrices compare the previous
        # month against the current one. The 'PÓLIZA SAP' section does not.
        mes_row = tk.Frame(frame1, bg=self.bg_color)
        mes_row.pack(fill="x", pady=(8, 0))

        tk.Label(
            mes_row, text="Mes:", font=("Segoe UI", 9, "bold"),
            bg=self.bg_color, fg=self.secondary_color,
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Entry(
            mes_row, textvariable=self.mes_var, width=10,
            font=("Segoe UI", 9), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        ).pack(side=tk.LEFT, padx=(0, 10), ipady=3)

        self.preview_mes_var = tk.StringVar()
        tk.Label(
            mes_row, textvariable=self.preview_mes_var,
            font=("Segoe UI", 8), bg=self.bg_color, fg=self.secondary_color,
            anchor="w", justify="left",
        ).pack(side=tk.LEFT, fill="x", expand=True)

        self.mes_var.trace_add("write", lambda *_: self._actualizar_preview_mes())

        self.consolidar_btn = tk.Button(
            frame1, text="🧩 Consolidar 'Variaciones'",
            font=("Segoe UI", 10, "bold"),
            bg=self.success_color, fg=self.bg_color,
            relief=tk.RAISED, bd=2, pady=8, cursor="hand2",
            activebackground=self.secondary_color, activeforeground=self.bg_color,
            command=self._handle_consolidar,
        )
        self.consolidar_btn.pack(fill="x", pady=(10, 0))

        # ==============================================================
        # PROCESO 2 — 'PÓLIZA SAP'
        # ==============================================================
        frame2 = tk.LabelFrame(
            main, text="  2. Layout Creación Reservas DA WHSL  →  sección 'PÓLIZA SAP'  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=14, pady=10,
        )
        frame2.pack(fill="x", pady=(0, 10))

        self._fila_archivo(frame2, self.archivo_poliza_var, self._elegir_archivo_poliza)

        tk.Label(
            frame2,
            font=("Segoe UI", 8), bg=self.bg_color, fg=self.secondary_color,
            anchor="w", justify="left",
        ).pack(fill="x", pady=(8, 0))

        self.poliza_btn = tk.Button(
            frame2, text="📋 Generar 'PÓLIZA SAP'",
            font=("Segoe UI", 10, "bold"),
            bg=self.accent_color, fg=self.bg_color,
            relief=tk.RAISED, bd=2, pady=8, cursor="hand2",
            activebackground=self.secondary_color, activeforeground=self.bg_color,
            command=self._handle_poliza,
        )
        self.poliza_btn.pack(fill="x", pady=(8, 0))

        # ==============================================================
        # Run everything
        # ==============================================================
        self.todo_btn = tk.Button(
            main, text="⚡ Ejecutar todo  (1 → 2)",
            font=("Segoe UI", 11, "bold"),
            bg=self.todo_color, fg=self.bg_color,
            relief=tk.RAISED, bd=2, pady=10, cursor="hand2",
            activebackground=self.secondary_color, activeforeground=self.bg_color,
            command=self._handle_todo,
        )
        self.todo_btn.pack(fill="x", pady=(2, 0))

        # --- Status bar ---
        self.status_var = tk.StringVar(value="✓ Listo para comenzar")
        tk.Label(
            main, textvariable=self.status_var, font=("Segoe UI", 9),
            bg=self.bg_color, fg="#666666", wraplength=610,
            anchor="w", justify="left",
        ).pack(fill="x", pady=(12, 0))

    def _fila_archivo(self, parent, variable, comando):
        """The 'entry + Examinar...' row, identical for both files."""
        fila = tk.Frame(parent, bg=self.bg_color)
        fila.pack(fill="x")

        tk.Entry(
            fila, textvariable=variable,
            font=("Segoe UI", 9), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        ).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 8), ipady=4)

        tk.Button(
            fila, text="Examinar...", font=("Segoe UI", 9, "bold"),
            bg=self.primary_color, fg=self.bg_color, relief=tk.FLAT,
            cursor="hand2", padx=12, command=comando,
        ).pack(side=tk.LEFT)
        return fila

    # ------------------------------------------------------------------
    # Pickers
    # ------------------------------------------------------------------
    def _elegir(self, variable, titulo, filtros):
        ruta = filedialog.askopenfilename(
            title=titulo,
            initialdir=(os.path.dirname(variable.get()) or os.path.expanduser("~")),
            filetypes=filtros,
        )
        if ruta:  # cancel leaves the current selection untouched
            variable.set(os.path.normpath(ruta))
            self._guardar_config()

    def _elegir_archivo(self):
        self._elegir(
            self.archivo_var,
            "Selecciona el archivo de Creación de Reservas PMF OT",
            FILTRO_VARIACIONES,
        )

    def _elegir_archivo_poliza(self):
        self._elegir(
            self.archivo_poliza_var,
            "Selecciona el Layout Creación Reservas DA WHSL",
            FILTRO_POLIZA,
        )

    def _actualizar_preview_mes(self):
        """Live hint of which two months the run will compare."""
        try:
            anterior, actual = claves_mes_actual_y_anterior(self.get_mes())
            self.preview_mes_var.set(
                f"Comparará {anterior} vs {actual} "
                f"({etiqueta_mes_es(anterior)} vs {etiqueta_mes_es(actual)})"
            )
        except ValueError:
            self.preview_mes_var.set("⚠️ Formato inválido — se espera M07-26")

    # ------------------------------------------------------------------
    # Handlers
    # ------------------------------------------------------------------
    def _handle_consolidar(self):
        self._disparar(self.on_consolidar, self.validate_variaciones,
                       self.get_archivo)

    def _handle_poliza(self):
        self._disparar(self.on_poliza, self.validate_poliza,
                       self.get_archivo_poliza)

    def _handle_todo(self):
        self._disparar(self.on_todo, self.validate_todo, self.get_archivo)

    def _disparar(self, callback, validador, accesor):
        """Validate, then call the controller — or say the wiring is missing."""
        if not validador():
            return
        if callback:
            callback()
        else:
            messagebox.showinfo(
                "Info",
                f"Funcionalidad no conectada\n\nArchivo: {accesor()}",
            )

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------
    def _validar_ruta(self, ruta, etiqueta, extensiones=None):
        """Shared checks: something was chosen, it exists, and it looks right."""
        if not ruta:
            messagebox.showerror(
                "Falta el archivo",
                f"Selecciona el archivo de {etiqueta} con el botón 'Examinar...'.",
            )
            return False
        if not os.path.exists(ruta):
            messagebox.showerror(
                "Archivo no encontrado",
                f"No encontré el archivo:\n\n{ruta}\n\n"
                f"¿Lo movieron o le cambiaron el nombre?",
            )
            return False
        if extensiones and os.path.splitext(ruta)[1].lower() not in extensiones:
            # A warning, not a blocker: the file could legitimately be saved in
            # another Excel format one month.
            seguir = messagebox.askyesno(
                "Extensión inesperada",
                f"El archivo de {etiqueta} suele ser "
                f"{' o '.join(extensiones)}, y este es "
                f"'{os.path.splitext(ruta)[1]}'.\n\n¿Continúo de todas formas?",
            )
            if not seguir:
                return False
        return True

    def validate_variaciones(self):
        if not self._validar_ruta(
            self.get_archivo(), "Creación de Reservas PMF OT", {".xlsx", ".xlsm"}
        ):
            return False
        try:
            claves_mes_actual_y_anterior(self.get_mes())
        except ValueError as e:
            messagebox.showerror("Mes inválido", str(e))
            return False
        self._guardar_config()
        return True

    def validate_poliza(self):
        if not self._validar_ruta(
            self.get_archivo_poliza(), "Layout Creación Reservas DA WHSL",
            {".xlsb", ".xlsx", ".xlsm"}
        ):
            return False
        self._guardar_config()
        return True

    def validate_todo(self):
        return self.validate_variaciones() and self.validate_poliza()

    # Kept so any older call site (or a quick test script) still works.
    def validate(self):
        return self.validate_variaciones()

    # ------------------------------------------------------------------
    # Accessors used by the controller
    # ------------------------------------------------------------------
    def get_archivo(self):
        return self.archivo_var.get().strip()

    def get_archivo_poliza(self):
        return self.archivo_poliza_var.get().strip()

    def get_mes(self):
        """Month key typed by the user, or None to use today's month."""
        return self.mes_var.get().strip().upper() or None

    def set_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()

    # ------------------------------------------------------------------
    # Button state
    # ------------------------------------------------------------------
    def _botones(self):
        return (
            (self.consolidar_btn, self.success_color),
            (self.poliza_btn,     self.accent_color),
            (self.todo_btn,       self.todo_color),
        )

    def disable_buttons(self):
        for boton, _ in self._botones():
            boton.config(state="disabled", bg="#666666")

    def enable_buttons(self):
        for boton, color in self._botones():
            boton.config(state="normal", bg=color)


def main():
    # Standalone entry point, to click through the GUI without the controller.
    root = tk.Tk()
    ReservasGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
