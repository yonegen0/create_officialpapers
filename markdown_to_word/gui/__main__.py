"""Windows GUI entry point: python -m markdown_to_word.gui."""
from __future__ import annotations

import ctypes
import logging
import os
import tkinter as tk
from pathlib import Path

from .view import TkView

logger = logging.getLogger(__name__)


def _dpi_awareness() -> None:
    """Select one Windows DPI mode before creating the first Tk window."""
    if os.name != "nt":
        return
    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        if user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass
    try:
        shcore = ctypes.WinDLL("shcore", use_last_error=True)
        if shcore.SetProcessDpiAwareness(2) == 0:
            return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.WinDLL("user32", use_last_error=True).SetProcessDPIAware()
    except (AttributeError, OSError):
        pass


def _work_area(root: tk.Tk) -> tuple[int, int, int, int]:
    if os.name == "nt":
        class Rect(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        rect = Rect()
        try:
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            if user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
                return rect.left, rect.top, rect.right, rect.bottom
        except (AttributeError, OSError):
            pass
    return 0, 0, root.winfo_screenwidth(), root.winfo_screenheight()


def _fit_to_work_area(root: tk.Tk) -> None:
    """Keep the initial window and its minimum inside the available desktop."""
    root.update_idletasks()
    left, top, right, bottom = _work_area(root)
    area_width = max(1, right - left)
    area_height = max(1, bottom - top)
    width = min(960, area_width)
    height = min(700, area_height)
    root.minsize(min(760, area_width), min(600, area_height))
    x = left + max(0, (area_width - width) // 2)
    y = top + max(0, (area_height - height) // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")


def _set_window_icon(root: tk.Tk) -> None:
    """Set the bundled Windows icon without preventing startup on failure."""
    if os.name != "nt":
        return
    try:
        icon_path = Path(__file__).resolve().parents[1] / "assets" / "MarkdownToWord.ico"
        root.iconbitmap(default=str(icon_path))
    except (tk.TclError, OSError) as exc:
        logger.warning("Application icon could not be loaded: %s", exc)


def main() -> None:
    _dpi_awareness()
    root = tk.Tk()
    root.withdraw()
    _set_window_icon(root)
    TkView(root)
    _fit_to_work_area(root)
    root.deiconify()
    root.mainloop()


if __name__ == "__main__":
    main()
