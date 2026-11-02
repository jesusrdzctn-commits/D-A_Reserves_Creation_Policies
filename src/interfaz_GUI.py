"""
RDA Pólizas D&A — interfaz_GUI.py
=================================

tkinter GUI for the month-end close of the two Reservas workbooks. Four steps,
top to bottom, the way the stakeholder runs them every month:

    1  Traspasar templates   the five 'RESERVAS <CANAL>' templates (tabs
                             'Reservas Manuales' + 'Reservas Rentas') -> 'Base'
                             of a COPY of the Creación de Reservas workbook
    2  Consolidar Variaciones  pestaña 'Variaciones' + tablas dinámicas, on that copy
    3  Generar Póliza SAP    Layout OT (.xlsx, tabs 'Reservas Manuales' + 'Reservas Rentas')
                             -> sección 'PÓLIZA SAP' of each tab
                             + plantilla 'Póliza Carga Batch' (.xlsx), the file SAP gets
    4  Subir a SAP           opens the upload transaction in SAP GUI and loads the 'Póliza Carga
                             Batch' that step 3 generated (it does not execute it yet)

Design: the look of a trade-spend / Discounts & Allowances workbench. A navy side
rail holds the context (period) and the progress of the close, one coloured dot
per step (pending / next / running / done / partial / error); the work area holds
one card per step with its inputs and ONE primary button; a status strip at the
bottom says what is happening. Colour carries meaning (green done, amber
attention, red error), never decoration.

Each step owns its pickers and its button because they are separate inputs with
separate failures: if the layout is open in Excel, step 2 must still work.

No SAP, no pandas and no Excel imports here on purpose: this file can be opened
and clicked through on any machine with `python interfaz_GUI.py`. The controller
plugs in the callbacks, the channel detector for the file names, and reports
the state of each step with marcar_paso().
"""

import json
import os
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox

from utils import claves_mes_actual_y_anterior, etiqueta_mes_es

# Which extensions each picker offers. The OT layout is a plain .xlsx, like the
# rest of the files (the old WHSL layout was a binary .xlsb).
FILTRO_VARIACIONES = [("Libros de Excel", "*.xlsx *.xlsm"),
                      ("Todos los archivos", "*.*")]
FILTRO_POLIZA      = [("Libros de Excel", "*.xlsx *.xlsm"),
                      ("Todos los archivos", "*.*")]
# The 'Póliza Carga Batch' template is a plain .xlsx (the file SAP receives).
FILTRO_BATCH       = [("Libros de Excel", "*.xlsx *.xlsm"),
                      ("Todos los archivos", "*.*")]
# The five 'RESERVAS MANUALES <CANAL>' templates (several can be picked at once).
FILTRO_TEMPLATES   = [("Libros de Excel", "*.xlsx *.xlsm"),
                      ("Todos los archivos", "*.*")]

FUENTE = "Segoe UI"


# ----------------------------------------------------------------------
# Palette and step states
# ----------------------------------------------------------------------
class Paleta:
    fondo     = "#F3F5F9"   # work area
    tarjeta   = "#FFFFFF"
    borde     = "#DDE3EE"
    campo     = "#F5F7FB"   # entries and lists
    navy      = "#00094F"   # side rail, primary
    navy_suave = "#1A2370"  # lines on the rail
    navy_brillo = "#2C3894" # light edge of the rail (catches the "light")
    sombra    = "#000521"   # tint of the shadow the rail casts
    texto     = "#1B2540"
    suave     = "#6B7690"   # secondary text
    riel_texto = "#AEB8E6"  # secondary text on the rail
    acento    = "#080808"
    negro     = "#000000"   # 'Examinar...' buttons
    verde     = "#034E26"
    ambar     = "#E59A0B"
    rojo      = "#D64545"


# state -> (label, text colour, chip background, symbol)
ESTADOS = {
    "pendiente": ("Pendiente", "#6B7690", "#E9EDF5", "○"),
    "siguiente": ("Siguiente", "#0067B1", "#DCEBF7", "▶"),
    "en_curso":  ("En curso",  "#0067B1", "#DCEBF7", "●"),
    "listo":     ("Listo",     "#17794A", "#DDF3E7", "✓"),
    "parcial":   ("A medias",  "#9A6400", "#FCEFD2", "!"),
    "error":     ("Con error", "#B83232", "#FBE0E0", "✕"),
}

# step key -> (number, short title for the rail, title of the card)
PASOS = (
    ("traspaso",      "1", "Traspasar layouts",   "Traspasar layouts a 'Base' (opcional)"),
    ("consolidacion", "2", "Consolidar Variaciones", "Consolidar Variaciones"),
    ("poliza",        "3", "Póliza SAP",             "Generar Póliza SAP"),
    ("subida",        "4", "Subir a SAP",            "Subir a SAP"),
)


def _oscurecer(color, factor=0.12):
    """'#RRGGBB' darkened by `factor` (hover / pressed states)."""
    r, g, b = (int(color[i:i + 2], 16) for i in (1, 3, 5))
    return "#%02X%02X%02X" % tuple(int(c * (1 - factor)) for c in (r, g, b))


def _resaltar(color, factor=0.12):
    """Hover / pressed shade of `color`: darker, or lighter when it is near black (black cannot get darker)."""
    r, g, b = (int(color[i:i + 2], 16) for i in (1, 3, 5))
    if max(r, g, b) >= 48:
        return _oscurecer(color, factor)
    return "#%02X%02X%02X" % tuple(int(c + (255 - c) * factor * 1.5) for c in (r, g, b))


def _degradado_sombra(fondo, tinte, ancho, opacidad):
    """`ancho` hex colours: `tinte` blended over `fondo` at `opacidad` next to the rail, fading to plain `fondo`.
    Tk has no alpha channel, so the blend is computed against the known background."""
    base = [int(fondo[i:i + 2], 16) for i in (1, 3, 5)]
    sombra = [int(tinte[i:i + 2], 16) for i in (1, 3, 5)]
    colores = []
    for i in range(ancho):
        alfa = opacidad * (1 - i / ancho) ** 2      # quadratic falloff = soft, diffused edge
        colores.append("#%02X%02X%02X" % tuple(round(b + (s - b) * alfa) for b, s in zip(base, sombra)))
    return colores


def activar_dpi_alto():
    """
    Ask Windows not to bitmap-stretch the window on high-DPI screens (blurry
    text). Must run BEFORE tk.Tk() is created. No-op elsewhere; never raises.
    """
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


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


# ----------------------------------------------------------------------
# Rounded button (tk.Button cannot round its corners)
# ----------------------------------------------------------------------
class BotonRedondeado(tk.Canvas):
    """
    A flat, rounded, hover-aware button drawn on a Canvas. Works with the mouse
    and with the keyboard (Tab to reach it, Enter / Space to press it).
    """
    DESHABILITADO = "#C9CFDC"

    def __init__(self, parent, texto, comando, color, fondo_padre,
                 fg="#FFFFFF", fuente=(FUENTE, 10, "bold"), alto=42, radio=10, ancho=None):
        super().__init__(parent, height=alto, bg=fondo_padre, highlightthickness=0,
                         bd=0, takefocus=1, cursor="hand2")
        # Minimum width = the text plus margins. A Canvas asks for ~394 px by default, which would
        # force every column that wide; buttons packed with fill="x" stretch to their card anyway.
        self.configure(width=ancho or tkfont.Font(root=parent, font=fuente).measure(texto) + 32)
        self._texto, self._comando = texto, comando
        self._color, self._fg, self._fuente, self._radio = color, fg, fuente, radio
        self._habilitado = True
        self._hover = self._presionado = self._foco = False

        self.bind("<Configure>",       lambda _e: self._dibujar())
        self.bind("<Enter>",           lambda _e: self._estado(hover=True))
        self.bind("<Leave>",           lambda _e: self._estado(hover=False, presionado=False))
        self.bind("<ButtonPress-1>",   self._al_presionar)
        self.bind("<ButtonRelease-1>", self._al_soltar)
        self.bind("<FocusIn>",         lambda _e: self._estado(foco=True))
        self.bind("<FocusOut>",        lambda _e: self._estado(foco=False))
        self.bind("<Return>",          lambda _e: self._invocar())
        self.bind("<space>",           lambda _e: self._invocar())

    # --- drawing ------------------------------------------------------
    def _estado(self, **cambios):
        for clave, valor in cambios.items():
            setattr(self, f"_{clave}", valor)
        self._dibujar()

    def _dibujar(self):
        self.delete("all")
        w, h, r = self.winfo_width(), self.winfo_height(), self._radio
        if w <= 2 or h <= 2:
            return
        if not self._habilitado:
            relleno = self.DESHABILITADO
        elif self._presionado:
            relleno = _resaltar(self._color, 0.24)
        elif self._hover:
            relleno = _resaltar(self._color, 0.10)
        else:
            relleno = self._color
        borde = Paleta.ambar if (self._foco and self._habilitado) else relleno
        x1, y1, x2, y2 = 1, 1, w - 1, h - 1
        puntos = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
                  x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        self.create_polygon(puntos, smooth=True, fill=relleno, outline=borde, width=2)
        self.create_text(w / 2, h / 2, text=self._texto, fill=self._fg, font=self._fuente)

    # --- behaviour ----------------------------------------------------
    def _al_presionar(self, _evento):
        if self._habilitado:
            self._estado(presionado=True)

    def _al_soltar(self, evento):
        dentro = 0 <= evento.x <= self.winfo_width() and 0 <= evento.y <= self.winfo_height()
        fue_presionado = self._presionado
        self._estado(presionado=False)
        if self._habilitado and dentro and fue_presionado:
            self._invocar()

    def _invocar(self):
        if self._habilitado and self._comando:
            self._comando()

    def habilitar(self, habilitado):
        self._habilitado = bool(habilitado)
        self.configure(cursor="hand2" if self._habilitado else "arrow")
        self._dibujar()

    def poner_texto(self, texto):
        self._texto = texto
        self._dibujar()


# ----------------------------------------------------------------------
# Progress bar (hand-drawn: ttk themes disagree on how to paint one)
# ----------------------------------------------------------------------
class BarraProgreso(tk.Canvas):
    """A thin 0-100 progress bar. `barra.valor = 40` moves it."""

    def __init__(self, parent, fondo, color=Paleta.acento, pista="#E3E8F2", alto=6):
        super().__init__(parent, height=alto, bg=fondo, highlightthickness=0, bd=0)
        self._color, self._pista, self._valor = color, pista, 0.0
        self.bind("<Configure>", lambda _e: self._dibujar())

    @property
    def valor(self):
        return self._valor

    @valor.setter
    def valor(self, nuevo):
        self._valor = max(0.0, min(100.0, float(nuevo)))
        self._dibujar()

    def _dibujar(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 2:
            return
        self.create_rectangle(0, 0, w, h, fill=self._pista, outline=self._pista)
        ancho = int(w * self._valor / 100.0)
        if ancho > 0:
            self.create_rectangle(0, 0, ancho, h, fill=self._color, outline=self._color)


# ----------------------------------------------------------------------
# The window
# ----------------------------------------------------------------------
class ReservasGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Robotic Desktop Automation Pólizas D&A")
        self.root.configure(bg=Paleta.fondo)
        self._ajustar_ventana()

        # Callbacks the controller plugs in.
        self.on_traspasar  = None   # templates -> 'Base' (copy of archivo 1)
        self.on_consolidar = None   # archivo 1 -> 'Variaciones'
        self.on_poliza     = None   # archivo 2 -> 'PÓLIZA SAP'
        self.on_subir      = None   # Carga Batch -> SAP (opens the transaction, no F8 yet)

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
        # Empty 'Póliza Carga Batch' template that step 3 fills at the end.
        self.archivo_batch_var = tk.StringVar(
            value=(cfg.get("ultimo_archivo_batch") or "").strip()
        )
        # Empty = the month the code is being run (the spec's default). Typing
        # a key such as M07-26 re-runs a past month, which is also how the demo
        # is driven when the test file holds older months.
        self.mes_var = tk.StringVar(value="")

        # The month's 'RESERVAS MANUALES' templates. Deliberately NOT restored
        # from config.json: last month's files must never sit pre-selected and
        # get pasted again by a forgetful click. Only the FOLDER is remembered,
        # so the file dialog opens where the templates usually live.
        self.templates = []
        self.carpeta_templates = (cfg.get("ultima_carpeta_templates") or "").strip()

        # The 'Póliza Carga Batch' step 3 produced in THIS session, the file step 4 uploads. Deliberately
        # NOT restored from config.json: last month's file must never sit armed and get uploaded by a click.
        self.ruta_subida = ""

        # Channel checklist: the controller says which channels to expect and how
        # to read one from a file name (the GUI knows nothing about the rules).
        self._canales_esperados = []
        self._detector_canal = None
        self._chips_canal = {}

        self._estado_paso = {clave: "pendiente" for clave, *_ in PASOS}
        self._chips_paso = {}       # card chips
        self._riel_paso = {}        # rail rows: (circle canvas, state label)
        self._ocupado = False

        self._build_ui()
        self._refrescar_pasos()
        self._actualizar_preview_mes()
        self.archivo_var.trace_add("write", lambda *_: self._actualizar_archivo_trabajo())
        self._actualizar_archivo_trabajo()
        # Last: the window is measured once every chip and label already has its text.
        self._ajustar_al_contenido()

    # ------------------------------------------------------------------
    # Window geometry
    # ------------------------------------------------------------------
    def _ajustar_ventana(self):
        """Centered, scaled for high-DPI screens and never taller than the screen."""
        escala = max(1.0, self.root.winfo_fpixels("1i") / 96.0)
        pantalla_w, pantalla_h = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        ancho = min(int(1040 * escala), pantalla_w - 40)
        alto = min(int(640 * escala), pantalla_h - 90)
        x = max(0, (pantalla_w - ancho) // 2)
        y = max(0, (pantalla_h - alto) // 3)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.root.minsize(min(int(980 * escala), ancho), min(int(600 * escala), alto))

    def _ajustar_al_contenido(self):
        """Fit the window to what the layout really needs (fonts and DPI change it), within the screen."""
        self.root.update_idletasks()
        pantalla_w, pantalla_h = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        ancho = min(max(self.root.winfo_width(), self.root.winfo_reqwidth()), pantalla_w - 20)
        alto = min(self.root.winfo_reqheight(), pantalla_h - 60)
        x = max(0, (pantalla_w - ancho) // 2)
        y = max(0, (pantalla_h - alto) // 3)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.root.minsize(min(self.root.winfo_reqwidth(), ancho), min(self.root.winfo_reqheight(), alto))

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
                        "ultimo_archivo_batch": self.archivo_batch_var.get().strip(),
                        # Only the folder: the templates themselves change every month.
                        "ultima_carpeta_templates": self.carpeta_templates,
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
        riel = tk.Frame(self.root, bg=Paleta.navy, width=270)
        riel.pack(side=tk.LEFT, fill="y")
        riel.pack_propagate(False)
        self._build_riel(riel)

        principal = tk.Frame(self.root, bg=Paleta.fondo, padx=20, pady=16)
        principal.pack(side=tk.LEFT, fill="both", expand=True)
        self._dibujar_sombra_riel(riel)

        # Footer first (side=bottom) so it is never pushed out when the window is short.
        self._build_pie(principal)

        # Step 1 spans the top row; steps 2, 3 and 4 share the second one in equal columns.
        # Step 4 is just a button: "new" keeps its natural height instead of stretching to its neighbours.
        cuadricula = tk.Frame(principal, bg=Paleta.fondo)
        cuadricula.pack(fill="both", expand=True)
        for columna in range(3):
            cuadricula.columnconfigure(columna, weight=1, uniform="tarjetas")
        cuadricula.rowconfigure(1, weight=1)
        self._build_tarjeta_traspaso(cuadricula).grid(
            row=0, column=0, columnspan=3, sticky="nsew", pady=(0, 12))
        # Margins (0,8) / (4,4) / (8,0): 12 px between cards and the three come out exactly as wide.
        self._build_tarjeta_consolidacion(cuadricula).grid(
            row=1, column=0, sticky="nsew", padx=(0, 8))
        self._build_tarjeta_poliza(cuadricula).grid(row=1, column=1, sticky="nsew", padx=4)
        self._build_tarjeta_subida(cuadricula).grid(row=1, column=2, sticky="new", padx=(8, 0))

    # --- side rail ----------------------------------------------------
    def _dibujar_sombra_riel(self, riel):
        """3D look: a soft shadow cast on the work area plus a light 1 px edge on the rail itself."""
        ancho = 18      # narrower than the 20 px gutter next to the rail, so it never touches a card
        sombra = tk.Canvas(self.root, width=ancho, bg=Paleta.fondo, highlightthickness=0, bd=0)
        alto = self.root.winfo_screenheight()   # taller than any window: the canvas clips the lines
        for x, color in enumerate(_degradado_sombra(Paleta.fondo, Paleta.sombra, ancho, 0.34)):
            sombra.create_line(x, 0, x, alto, fill=color)
        # Overlay glued to the rail's right edge: it takes no layout space.
        sombra.place(in_=riel, relx=1.0, y=0, relheight=1)
        tk.Frame(riel, bg=Paleta.navy_brillo, width=1).place(relx=1.0, x=-1, y=0, relheight=1)

    def _build_riel(self, riel):
        marca = tk.Frame(riel, bg=Paleta.navy, padx=22, pady=20)
        marca.pack(fill="x")
        tk.Label(marca, text="RDA", font=(FUENTE, 24, "bold"), fg="#FFFFFF",
                 bg=Paleta.navy).pack(anchor="w")
        tk.Label(marca, text="Pólizas SAP", font=(FUENTE, 12, "bold"), fg="#FFFFFF",
                 bg=Paleta.navy).pack(anchor="w")
        tk.Label(marca, text="Cierre mensual de reservas\nDiscounts & Allowances",
                 font=(FUENTE, 9), fg=Paleta.riel_texto, bg=Paleta.navy,
                 justify="left").pack(anchor="w", pady=(4, 0))

        self._separador_riel(riel)

        # Period: belongs to the whole close (the paste checks the templates against
        # it, the matrices compare it with the previous month).
        periodo = tk.Frame(riel, bg=Paleta.navy, padx=22, pady=6)
        periodo.pack(fill="x")
        self._titulo_riel(periodo, "PERIODO DE CIERRE")
        fila = tk.Frame(periodo, bg=Paleta.navy)
        fila.pack(fill="x", pady=(6, 0))
        tk.Label(fila, text="Mes", font=(FUENTE, 10, "bold"), fg="#FFFFFF",
                 bg=Paleta.navy).pack(side=tk.LEFT, padx=(0, 8))
        tk.Entry(
            fila, textvariable=self.mes_var, width=9, font=(FUENTE, 11, "bold"),
            bg="#FFFFFF", fg=Paleta.navy, relief=tk.FLAT, bd=0, justify="center",
            insertbackground=Paleta.navy, highlightthickness=2,
            highlightbackground=Paleta.navy, highlightcolor=Paleta.ambar,
        ).pack(side=tk.LEFT, ipady=4)
        tk.Label(fila, text="  ej. M10-26", font=(FUENTE, 8), fg=Paleta.riel_texto,
                 bg=Paleta.navy).pack(side=tk.LEFT)

        self.preview_mes_var = tk.StringVar()
        self.preview_mes_lbl = tk.Label(
            periodo, textvariable=self.preview_mes_var, font=(FUENTE, 9),
            fg="#FFFFFF", bg=Paleta.navy, justify="left", anchor="w", wraplength=220,
        )
        self.preview_mes_lbl.pack(fill="x", pady=(8, 0))
        tk.Label(periodo, text="Vacío = el mes en curso.", font=(FUENTE, 8),
                 fg=Paleta.riel_texto, bg=Paleta.navy, anchor="w").pack(fill="x", pady=(2, 0))
        self.mes_var.trace_add("write", lambda *_: self._actualizar_preview_mes())

        self._separador_riel(riel)

        # Progress of the close
        avance = tk.Frame(riel, bg=Paleta.navy, padx=22, pady=6)
        avance.pack(fill="x")
        self._titulo_riel(avance, "AVANCE DEL CIERRE")
        for clave, numero, titulo, _ in PASOS:
            fila = tk.Frame(avance, bg=Paleta.navy)
            fila.pack(fill="x", pady=(10, 0))
            circulo = tk.Canvas(fila, width=30, height=30, bg=Paleta.navy, highlightthickness=0)
            circulo.pack(side=tk.LEFT)
            textos = tk.Frame(fila, bg=Paleta.navy)
            textos.pack(side=tk.LEFT, padx=(10, 0))
            tk.Label(textos, text=titulo, font=(FUENTE, 10, "bold"), fg="#FFFFFF",
                     bg=Paleta.navy, anchor="w").pack(anchor="w")
            estado = tk.Label(textos, text="", font=(FUENTE, 8), fg=Paleta.riel_texto,
                              bg=Paleta.navy, anchor="w")
            estado.pack(anchor="w")
            self._riel_paso[clave] = (circulo, estado, numero)

        # Footer of the rail: reassurance, in plain words.
        tk.Label(
            riel,
            text="Cada paso pide confirmación\nantes de escribir. Los archivos\n"
                 "originales no se tocan: todo se\nguarda como copia en 'Output'.",
            font=(FUENTE, 8), fg=Paleta.riel_texto, bg=Paleta.navy, justify="left",
        ).pack(side=tk.BOTTOM, anchor="w", padx=22, pady=18)

    @staticmethod
    def _titulo_riel(parent, texto):
        tk.Label(parent, text=texto, font=(FUENTE, 8, "bold"), fg=Paleta.riel_texto,
                 bg=Paleta.navy, anchor="w").pack(fill="x")

    @staticmethod
    def _separador_riel(riel):
        tk.Frame(riel, bg=Paleta.navy_suave, height=1).pack(fill="x", padx=22, pady=10)

    # --- cards --------------------------------------------------------
    def _tarjeta(self, parent, paso, descripcion=None):
        """A white card with a numbered header, a state chip and (optionally) a description."""
        numero, titulo = next((n, t) for c, n, _, t in PASOS if c == paso)
        marco = tk.Frame(parent, bg=Paleta.tarjeta, highlightthickness=1,
                         highlightbackground=Paleta.borde, highlightcolor=Paleta.borde)
        interior = tk.Frame(marco, bg=Paleta.tarjeta, padx=16, pady=10)
        interior.pack(fill="both", expand=True)

        cabecera = tk.Frame(interior, bg=Paleta.tarjeta)
        cabecera.pack(fill="x")
        insignia = tk.Canvas(cabecera, width=30, height=30, bg=Paleta.tarjeta, highlightthickness=0)
        insignia.create_oval(1, 1, 29, 29, fill=Paleta.navy, outline=Paleta.navy)
        insignia.create_text(15, 15, text=numero, fill="#FFFFFF", font=(FUENTE, 11, "bold"))
        insignia.pack(side=tk.LEFT)
        tk.Label(cabecera, text=titulo, font=(FUENTE, 11, "bold"), fg=Paleta.texto,
                 bg=Paleta.tarjeta).pack(side=tk.LEFT, padx=(10, 0))
        chip = tk.Label(cabecera, font=(FUENTE, 8, "bold"), padx=9, pady=2)
        chip.pack(side=tk.RIGHT)
        self._chips_paso[paso] = chip

        if descripcion:
            texto = tk.Label(interior, text=descripcion, font=(FUENTE, 9), fg=Paleta.suave,
                             bg=Paleta.tarjeta, justify="left", anchor="w", wraplength=300)
            texto.pack(fill="x", pady=(4, 6))
            # The description wraps to whatever width the card ends up with.
            texto.bind("<Configure>", lambda e, t=texto: t.configure(wraplength=max(e.width - 4, 120)))
        return marco, interior

    def _build_tarjeta_traspaso(self, parent):
        marco, interior = self._tarjeta(
            parent, "traspaso",
            "Pega las pestañas 'Reservas Manuales' y 'Reservas Rentas' de los 5 canales en "
            "'Base' y marca cada renglón con su canal en la columna CT. Si el mes ya estaba, "
            "lo reemplaza en vez de duplicarlo.",
        )

        self._etiqueta(interior, "Archivo de Creación de Reservas (destino):")
        self._fila_archivo(interior, self.archivo_var, self._elegir_archivo)

        encabezado = tk.Frame(interior, bg=Paleta.tarjeta)
        encabezado.pack(fill="x", pady=(8, 3))
        tk.Label(encabezado, text="Templates del mes:", font=(FUENTE, 8, "bold"),
                 fg=Paleta.texto, bg=Paleta.tarjeta).pack(side=tk.LEFT)
        self.canales_fila = tk.Frame(encabezado, bg=Paleta.tarjeta)
        self.canales_fila.pack(side=tk.RIGHT)

        fila = tk.Frame(interior, bg=Paleta.tarjeta)
        fila.pack(fill="x")
        self.templates_lista = tk.Listbox(
            fila, height=3, selectmode=tk.EXTENDED, activestyle="none",
            font=(FUENTE, 9), bg=Paleta.campo, fg=Paleta.navy, relief=tk.FLAT, bd=0,
            highlightthickness=1, exportselection=False, selectbackground=Paleta.acento,
            selectforeground="#FFFFFF", highlightbackground=Paleta.borde,
            highlightcolor=Paleta.navy,
        )
        self.templates_lista.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 8))

        # Small helpers of the list: smaller than the action buttons and in italics.
        ayudas = tk.Frame(fila, bg=Paleta.tarjeta)
        ayudas.pack(side=tk.LEFT, fill="y")
        for texto, comando in (("Agregar...", self._agregar_templates),
                               ("Quitar", self._quitar_templates),
                               ("Limpiar", self._limpiar_templates)):
            BotonRedondeado(
                ayudas, texto, comando, Paleta.suave, Paleta.tarjeta,
                fuente=(FUENTE, 8, "bold italic"), alto=26, radio=8, ancho=92,
            ).pack(pady=(0, 3))

        self.templates_resumen_var = tk.StringVar()
        tk.Label(interior, textvariable=self.templates_resumen_var, font=(FUENTE, 8),
                 fg=Paleta.suave, bg=Paleta.tarjeta, anchor="w",
                 justify="left").pack(fill="x", pady=(2, 6))

        self.traspasar_btn = BotonRedondeado(
            interior, "📥  Traspasar layouts...", self._handle_traspasar,
            Paleta.navy, Paleta.tarjeta, alto=38,
        )
        self.traspasar_btn.pack(fill="x")
        self._refrescar_templates()
        return marco

    def _build_tarjeta_consolidacion(self, parent):
        marco, interior = self._tarjeta(
            parent, "consolidacion",
            "Arma 'Variaciones' (5 matrices por canal, mes anterior vs actual) y las dos "
            "tablas dinámicas. Necesita Excel instalado.",
        )
        self._etiqueta(interior, "Trabajará sobre (copia del paso 1):")
        self.archivo_trabajo_var = tk.StringVar()
        trabajo = tk.Label(interior, textvariable=self.archivo_trabajo_var, font=(FUENTE, 9, "bold"),
                           fg=Paleta.navy, bg=Paleta.campo, anchor="w", justify="left", padx=8,
                           pady=6, wraplength=240)
        trabajo.pack(fill="x")
        # The file name wraps to the card's width (a third of the window), never wider than it.
        trabajo.bind("<Configure>", lambda e, t=trabajo: t.configure(wraplength=max(e.width - 16, 120)))

        self.consolidar_btn = BotonRedondeado(
            interior, "🧩  Consolidar 'Variaciones'", self._handle_consolidar,
            Paleta.navy, Paleta.tarjeta, alto=38,
        )
        self.consolidar_btn.pack(side=tk.BOTTOM, fill="x", pady=(8, 0))
        return marco

    def _build_tarjeta_poliza(self, parent):
        marco, interior = self._tarjeta(
            parent, "poliza",
            "Llena 'PÓLIZA SAP' en las pestañas 'Reservas Manuales' y 'Reservas Rentas' "
            "del Layout OT y de ahí la plantilla 'Póliza Carga Batch', el archivo que se "
            "sube a SAP. Usa el periodo para las fechas y el Period.",
        )
        # Captions above the fields: a third of the window is too narrow for an inline caption.
        self._etiqueta(interior, "Layout OT (.xlsx)")
        self._fila_archivo(interior, self.archivo_poliza_var, self._elegir_archivo_poliza)
        self._etiqueta(interior, "Carga Batch (.xlsx)", pady=(6, 3))
        self._fila_archivo(interior, self.archivo_batch_var, self._elegir_plantilla_batch)

        self.poliza_btn = BotonRedondeado(
            interior, "📋  Generar 'PÓLIZA SAP' + Carga Batch!", self._handle_poliza,
            Paleta.navy, Paleta.tarjeta, alto=38,
        )
        self.poliza_btn.pack(side=tk.BOTTOM, fill="x", pady=(8, 0))
        return marco

    def _build_tarjeta_subida(self, parent):
        # Compact on purpose: no description, just the file step 3 left ready and one button.
        marco, interior = self._tarjeta(parent, "subida")
        self.archivo_subida_var = tk.StringVar()
        archivo = tk.Label(interior, textvariable=self.archivo_subida_var, font=(FUENTE, 9, "bold"),
                           fg=Paleta.navy, bg=Paleta.campo, anchor="w", justify="left", padx=8,
                           pady=6, wraplength=240)
        archivo.pack(fill="x", pady=(8, 0))
        # A long file name wraps to the card's width instead of being cut.
        archivo.bind("<Configure>", lambda e, t=archivo: t.configure(wraplength=max(e.width - 16, 120)))

        self.subir_btn = BotonRedondeado(
            interior, "🚀  Subir a SAP", self._handle_subida, Paleta.navy, Paleta.tarjeta, alto=38,
        )
        self.subir_btn.pack(fill="x", pady=(8, 0))
        self._actualizar_archivo_subida()
        return marco

    # --- footer -------------------------------------------------------
    def _build_pie(self, parent):
        pie = tk.Frame(parent, bg=Paleta.fondo)
        pie.pack(side=tk.BOTTOM, fill="x", pady=(12, 0))

        izquierda = tk.Frame(pie, bg=Paleta.fondo)
        izquierda.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 16))
        fila = tk.Frame(izquierda, bg=Paleta.fondo)
        fila.pack(fill="x")
        self.punto = tk.Canvas(fila, width=12, height=12, bg=Paleta.fondo, highlightthickness=0)
        self.punto.pack(side=tk.LEFT, padx=(0, 8))
        self.status_var = tk.StringVar(value="✓ Listo para comenzar")
        self.status_lbl = tk.Label(
            fila, textvariable=self.status_var, font=(FUENTE, 9), fg=Paleta.texto,
            bg=Paleta.fondo, anchor="w", justify="left", wraplength=420,
        )
        self.status_lbl.pack(side=tk.LEFT, fill="x", expand=True)
        izquierda.bind("<Configure>", lambda e: self.status_lbl.configure(
            wraplength=max(e.width - 40, 160)))
        # The bar only shows while a step runs; its slot is always reserved.
        self.progreso_marco = tk.Frame(izquierda, bg=Paleta.fondo, height=8)
        self.progreso_marco.pack(fill="x", pady=(6, 0))
        self.progreso_marco.pack_propagate(False)
        self.progreso = BarraProgreso(self.progreso_marco, Paleta.fondo)
        self._pintar_punto(Paleta.verde)

    # --- small widgets ------------------------------------------------
    def _etiqueta(self, parent, texto, pady=(0, 3)):
        """Small caption above a field, so each picker says what it expects."""
        tk.Label(parent, text=texto, font=(FUENTE, 8, "bold"), fg=Paleta.texto,
                 bg=Paleta.tarjeta, anchor="w").pack(fill="x", pady=pady)

    def _fila_archivo(self, parent, variable, comando, etiqueta=None):
        """The 'entry + Examinar...' row, identical for every file (optional inline label)."""
        fila = tk.Frame(parent, bg=Paleta.tarjeta)
        fila.pack(fill="x")
        if etiqueta:
            tk.Label(fila, text=etiqueta, font=(FUENTE, 8, "bold"), fg=Paleta.texto,
                     bg=Paleta.tarjeta, width=15, anchor="w").pack(side=tk.LEFT)
        tk.Entry(
            fila, textvariable=variable, width=10, font=(FUENTE, 9), bg=Paleta.campo,
            fg=Paleta.navy, relief=tk.FLAT, bd=0, highlightthickness=1,
            highlightbackground=Paleta.borde, highlightcolor=Paleta.navy,
        ).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 8), ipady=5)
        BotonRedondeado(
            fila, "Examinar...", comando, Paleta.negro, Paleta.tarjeta,
            fuente=(FUENTE, 9, "bold"), alto=30, radio=8, ancho=96,
        ).pack(side=tk.LEFT)
        return fila

    @staticmethod
    def _pintar_chip(chip, estado):
        etiqueta, texto, fondo, simbolo = ESTADOS[estado]
        chip.configure(text=f"{simbolo}  {etiqueta}", fg=texto, bg=fondo)

    def _pintar_punto(self, color):
        self.punto.delete("all")
        self.punto.create_oval(1, 1, 11, 11, fill=color, outline=color)

    # ------------------------------------------------------------------
    # Step states (the controller reports them)
    # ------------------------------------------------------------------
    def marcar_paso(self, paso, estado):
        """
        Set the state of one step: paso in ('traspaso', 'consolidacion', 'poliza', 'subida'),
        estado in ('pendiente', 'en_curso', 'listo', 'parcial', 'error').
        """
        if paso not in self._estado_paso:
            raise ValueError(f"Paso desconocido: {paso!r}")
        if estado not in ESTADOS or estado == "siguiente":
            raise ValueError(f"Estado desconocido: {estado!r}")
        self._estado_paso[paso] = estado
        self._refrescar_pasos()
        self.root.update_idletasks()

    def _estados_mostrados(self):
        """
        State of each step as shown. 'Siguiente' is not stored: it is the first
        pending step after the last one that was already run (or the first step,
        if none was), and only while nothing is running.
        """
        estados = dict(self._estado_paso)
        orden = [clave for clave, *_ in PASOS]
        if "en_curso" in estados.values():
            return estados
        ultimo = max((i for i, c in enumerate(orden)
                      if estados[c] in ("listo", "parcial", "error")), default=-1)
        for clave in orden[ultimo + 1:]:
            if estados[clave] == "pendiente":
                estados[clave] = "siguiente"
                break
        return estados

    def _refrescar_pasos(self):
        colores = {"pendiente": ("#6F7BB8", None, "#AEB8E6"),
                   "siguiente": (Paleta.acento, Paleta.acento, "#FFFFFF"),
                   "en_curso":  (Paleta.acento, Paleta.acento, "#FFFFFF"),
                   "listo":     (Paleta.verde, Paleta.verde, "#FFFFFF"),
                   "parcial":   (Paleta.ambar, Paleta.ambar, "#FFFFFF"),
                   "error":     (Paleta.rojo, Paleta.rojo, "#FFFFFF")}
        for clave, estado in self._estados_mostrados().items():
            self._pintar_chip(self._chips_paso[clave], estado)
            circulo, etiqueta, numero = self._riel_paso[clave]
            borde, relleno, texto = colores[estado]
            circulo.delete("all")
            circulo.create_oval(2, 2, 28, 28, outline=borde, width=2, fill=relleno or Paleta.navy)
            simbolo = {"listo": "✓", "error": "✕", "parcial": "!"}.get(estado, numero)
            circulo.create_text(15, 15, text=simbolo, fill=texto, font=(FUENTE, 10, "bold"))
            etiqueta.configure(text=ESTADOS[estado][0])

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
            "Selecciona el archivo de Creación de Reservas",
            FILTRO_VARIACIONES,
        )

    def _elegir_archivo_poliza(self):
        self._elegir(
            self.archivo_poliza_var,
            "Selecciona el Layout OT (Reservas Manuales / Reservas Rentas)",
            FILTRO_POLIZA,
        )

    def _elegir_plantilla_batch(self):
        self._elegir(
            self.archivo_batch_var,
            "Selecciona la plantilla de Póliza Carga Batch",
            FILTRO_BATCH,
        )

    # --- RESERVAS MANUALES templates --------------------------------------
    @staticmethod
    def _clave_ruta(ruta):
        """Identity of a path for duplicate checks (case-insensitive on Windows)."""
        return os.path.normcase(os.path.abspath(ruta))

    def _agregar_templates(self):
        rutas = filedialog.askopenfilenames(
            title="Selecciona los templates RESERVAS MANUALES del mes",
            initialdir=self.carpeta_templates or os.path.dirname(self.get_archivo())
                       or os.path.expanduser("~"),
            filetypes=FILTRO_TEMPLATES,
        )
        if not rutas:          # cancel -> '' or an empty tuple
            return

        existentes = {self._clave_ruta(r) for r in self.templates}
        for ruta in rutas:
            ruta = os.path.normpath(ruta)
            if self._clave_ruta(ruta) not in existentes:
                self.templates.append(ruta)
                existentes.add(self._clave_ruta(ruta))

        self.carpeta_templates = os.path.dirname(os.path.normpath(rutas[-1]))
        self._refrescar_templates()
        self._guardar_config()

    def _quitar_templates(self):
        seleccion = set(self.templates_lista.curselection())
        if not seleccion:
            return
        self.templates = [r for i, r in enumerate(self.templates) if i not in seleccion]
        self._refrescar_templates()

    def _limpiar_templates(self):
        self.templates = []
        self._refrescar_templates()

    def configurar_canales(self, canales_esperados, detector):
        """
        Show the channel checklist of step 1. `canales_esperados` is the list of
        channel names (['ECOMM', 'HD', ...]); `detector(ruta)` returns the channel
        a file NAME points at, or None. Only a hint: the real channel comes from
        the template title and is validated by the paste.
        """
        self._canales_esperados = list(canales_esperados)
        self._detector_canal = detector
        for hijo in self.canales_fila.winfo_children():
            hijo.destroy()
        self._chips_canal = {}
        for canal in self._canales_esperados:
            chip = tk.Label(self.canales_fila, font=(FUENTE, 8, "bold"), padx=6, pady=1)
            chip.pack(side=tk.LEFT, padx=(4, 0))
            self._chips_canal[canal] = chip
        self._refrescar_templates()

    def _refrescar_templates(self):
        """Redraw the list (file names only; full paths are kept in self.templates)."""
        self.templates_lista.delete(0, tk.END)
        for ruta in self.templates:
            self.templates_lista.insert(tk.END, os.path.basename(ruta))

        detectados, sin_canal = {}, 0
        if self._detector_canal:
            for ruta in self.templates:
                try:
                    canal = self._detector_canal(ruta)
                except Exception:
                    canal = None
                if canal is None:
                    sin_canal += 1
                else:
                    detectados[canal] = detectados.get(canal, 0) + 1

        for canal, chip in self._chips_canal.items():
            n = detectados.get(canal, 0)
            if n == 1:
                chip.configure(text=f"✓ {canal}", fg=ESTADOS["listo"][1], bg=ESTADOS["listo"][2])
            elif n > 1:
                chip.configure(text=f"! {canal} ×{n}", fg=ESTADOS["parcial"][1],
                               bg=ESTADOS["parcial"][2])
            else:
                chip.configure(text=f"○ {canal}", fg=ESTADOS["pendiente"][1],
                               bg=ESTADOS["pendiente"][2])

        n = len(self.templates)
        if n == 0:
            resumen = "Sin templates: agrega los archivos del mes (uno por canal) para poder traspasarlos."
        else:
            resumen = f"{n} template(s) seleccionado(s): se validan todos antes de pegar nada."
            if sin_canal:
                resumen += f" {sin_canal} sin el canal en su nombre (se lee de su título)."
            faltan = [c for c in self._canales_esperados if c not in detectados]
            if self._detector_canal and faltan and not sin_canal:
                resumen += f" Faltan: {', '.join(faltan)}."
        self.templates_resumen_var.set(resumen)

    def _actualizar_preview_mes(self):
        """Live hint of which two months the run will compare."""
        try:
            anterior, actual = claves_mes_actual_y_anterior(self.get_mes())
            self.preview_mes_var.set(
                f"Compara {anterior} vs {actual}\n"
                f"({etiqueta_mes_es(anterior)} vs {etiqueta_mes_es(actual)})"
            )
            self.preview_mes_lbl.configure(fg="#FFFFFF")
        except ValueError:
            self.preview_mes_var.set("⚠️ Formato inválido — se espera M07-26")
            self.preview_mes_lbl.configure(fg="#FFB4B4")

    def _actualizar_archivo_trabajo(self):
        ruta = self.get_archivo()
        self.archivo_trabajo_var.set(os.path.basename(ruta) if ruta else "(elige el archivo en el paso 1)")

    def _actualizar_archivo_subida(self):
        ruta = self.get_archivo_subida()
        self.archivo_subida_var.set(os.path.basename(ruta) if ruta else "(se llena al terminar el paso 3)")

    # ------------------------------------------------------------------
    # Handlers
    # ------------------------------------------------------------------
    def _handle_traspasar(self):
        self._disparar(self.on_traspasar, self.validate_traspaso,
                       self.get_archivo)

    def _handle_consolidar(self):
        self._disparar(self.on_consolidar, self.validate_variaciones,
                       self.get_archivo)

    def _handle_poliza(self):
        self._disparar(self.on_poliza, self.validate_poliza,
                       self.get_archivo_poliza)

    def _handle_subida(self):
        self._disparar(self.on_subir, self.validate_subida,
                       self.get_archivo_subida)

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

    def _validar_mes(self):
        """The closing month must be empty (current month) or a valid key such as M10-26."""
        try:
            claves_mes_actual_y_anterior(self.get_mes())
        except ValueError as e:
            messagebox.showerror("Mes inválido", str(e))
            return False
        return True

    def _validar_archivo_y_mes(self):
        """Checks shared by the paste and the consolidation (same file, same month)."""
        if not self._validar_ruta(
            self.get_archivo(), "Creación de Reservas PMF", {".xlsx", ".xlsm"}
        ):
            return False
        return self._validar_mes()

    def validate_variaciones(self):
        if not self._validar_archivo_y_mes():
            return False
        self._guardar_config()
        return True

    def validate_traspaso(self):
        if not self._validar_archivo_y_mes():
            return False

        templates = self.get_templates()
        if not templates:
            messagebox.showerror(
                "Faltan los templates",
                "Agrega los templates RESERVAS MANUALES del mes con el botón "
                "'Agregar...'.",
            )
            return False

        principal = self._clave_ruta(self.get_archivo())
        for ruta in templates:
            if not self._validar_ruta(ruta, "un template RESERVAS MANUALES",
                                      {".xlsx", ".xlsm"}):
                return False
            # '~$name.xlsx' is the lock file Excel leaves next to a workbook that
            # is open: not a workbook at all.
            if os.path.basename(ruta).startswith("~$"):
                messagebox.showerror(
                    "Archivo temporal de Excel",
                    f"'{os.path.basename(ruta)}' es el archivo temporal que Excel crea "
                    f"mientras el libro está abierto.\n\nQuítalo de la lista y elige el "
                    f"template de verdad.",
                )
                return False
            # The same file as the workbook being filled would paste a file into itself.
            if self._clave_ruta(ruta) == principal:
                messagebox.showerror(
                    "Mismo archivo",
                    "Uno de los templates es el mismo archivo que el de Creación de "
                    "Reservas.\n\nQuítalo de la lista de templates.",
                )
                return False
        self._guardar_config()
        return True

    def validate_poliza(self):
        if not self._validar_mes():
            return False
        if not self._validar_ruta(
            self.get_archivo_poliza(), "Layout OT",
            {".xlsx", ".xlsm"}
        ):
            return False
        # Step 3 ends by filling the batch template, so it is mandatory.
        if not self._validar_ruta(
            self.get_archivo_batch(), "la plantilla de Póliza Carga Batch",
            {".xlsx", ".xlsm"}
        ):
            return False
        # Same file in both pickers would make the run overwrite its own input.
        if os.path.normcase(os.path.abspath(self.get_archivo_poliza())) == \
                os.path.normcase(os.path.abspath(self.get_archivo_batch())):
            messagebox.showerror(
                "Mismo archivo",
                "El Layout y la plantilla de Póliza Carga Batch son el mismo "
                "archivo.\n\nElige cada uno con su botón 'Examinar...'.",
            )
            return False
        self._guardar_config()
        return True

    def validate_subida(self):
        ruta = self.get_archivo_subida()
        if not ruta:
            messagebox.showerror(
                "Falta la Póliza Carga Batch",
                "Primero genera la póliza en el paso 3: el archivo que se sube a SAP es la "
                "Póliza Carga Batch que ese paso deja lista.",
            )
            return False
        return self._validar_ruta(ruta, "la Póliza Carga Batch", {".xlsx", ".xlsm"})

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

    def get_archivo_batch(self):
        return self.archivo_batch_var.get().strip()

    def get_archivo_subida(self):
        """The 'Póliza Carga Batch' step 4 would upload ('' = step 3 has not produced one)."""
        return self.ruta_subida

    def set_archivo_subida(self, ruta):
        """Arm step 4 with the batch file step 3 just produced; None or '' disarms it."""
        self.ruta_subida = os.path.normpath(ruta) if ruta else ""
        self._actualizar_archivo_subida()

    def set_archivo(self, ruta):
        """Point the Creación de Reservas field at `ruta` (and remember it)."""
        self.archivo_var.set(os.path.normpath(ruta))
        self._guardar_config()

    def get_templates(self):
        """The chosen 'RESERVAS MANUALES' templates (a copy; [] = none)."""
        return list(self.templates)

    def get_mes(self):
        """Month key typed by the user, or None to use today's month."""
        return self.mes_var.get().strip().upper() or None

    def set_status(self, message):
        """
        Show a progress / result line. The colour of the dot follows the leading
        symbol of the message (✅ green, ❌ red, ⚠ amber, anything else blue).
        While a step is running every message also advances the progress bar a
        fraction of what is left (it never reaches 100%: the length of a step is
        not known), since the steps run on the UI thread and a timer-driven
        animation would freeze.
        """
        self.status_var.set(message)
        inicio = message.lstrip()[:1]
        if inicio in ("✅", "✓"):
            color = Paleta.verde
        elif inicio == "❌":
            color = Paleta.rojo
        elif inicio == "⚠":
            color = Paleta.ambar
        else:
            color = Paleta.acento
        self._pintar_punto(color)
        if self._ocupado:
            self.progreso.valor = self.progreso.valor + (95.0 - self.progreso.valor) * 0.12
        self.root.update_idletasks()

    # ------------------------------------------------------------------
    # Button state
    # ------------------------------------------------------------------
    def _botones(self):
        return (self.traspasar_btn, self.consolidar_btn, self.poliza_btn, self.subir_btn)

    def disable_buttons(self):
        self._ocupado = True
        for boton in self._botones():
            boton.habilitar(False)
        self.progreso.valor = 0
        self.progreso.pack(fill="x")
        self.root.configure(cursor="watch")
        self.root.update_idletasks()

    def enable_buttons(self):
        self._ocupado = False
        for boton in self._botones():
            boton.habilitar(True)
        self.progreso.valor = 0
        self.progreso.pack_forget()
        self.root.configure(cursor="")
        self.root.update_idletasks()


def main():
    # Standalone entry point, to click through the GUI without the controller.
    activar_dpi_alto()
    root = tk.Tk()
    ReservasGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
