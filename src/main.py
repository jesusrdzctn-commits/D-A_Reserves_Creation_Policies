"""RDA Pólizas D&A — application entry point: wires the GUI to the controller."""

import tkinter as tk

from controller import ReservasController
from interfaz_GUI import ReservasGUI, activar_dpi_alto


def main():
    activar_dpi_alto()   # before tk.Tk(): sharp text on high-DPI screens
    root = tk.Tk()
    gui = ReservasGUI(root)
    ReservasController(gui)   # plugs the GUI callbacks in
    root.mainloop()


if __name__ == "__main__":
    main()
