import threading
import ttkbootstrap as ttk
from gui import TransporteGUI


def run_gui():
    root = ttk.Window(title="Reportes de Transporte MHC", themename="morph")
    TransporteGUI(root)
    root.mainloop()

if __name__ == "__main__":
    run_gui()  # Solo ejecuta la interfaz gráfica