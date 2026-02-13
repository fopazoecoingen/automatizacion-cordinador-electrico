import tkinter as tk

from app.gui.informe import InterfazInforme


def main() -> None:
    """Punto de entrada principal de la aplicación GUI."""
    root = tk.Tk()
    app = InterfazInforme(root)
    root.mainloop()


if __name__ == "__main__":
    main()

