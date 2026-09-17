"""
RDA Pólizas D&A
===============

Application entry point. Wires the GUI to the controller.

Reverse of Validación Factura Global: here the consolidation runs first and the
SAP upload comes afterwards. Today only the consolidation half exists.

Fecha: 2026
"""

import tkinter as tk

from controller import ReservasController
from interfaz_GUI import ReservasGUI


def main():
    root = tk.Tk()
    gui = ReservasGUI(root)
    ReservasController(gui)   # plugs the GUI callbacks in
    root.mainloop()


if __name__ == "__main__":
    main()
