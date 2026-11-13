"""GUI entry point: python -m markdown_to_word.gui"""
import os
import tkinter as tk

from .view import TkView


def _dpi_awareness() -> None:
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass


def main() -> None:
    _dpi_awareness()
    root = tk.Tk()
    TkView(root)
    root.mainloop()


if __name__ == "__main__":
    main()
