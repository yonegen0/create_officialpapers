"""Lightweight blue visual system for the Windows Markdown-to-Word interface."""
from __future__ import annotations

import ctypes
import os
import tkinter as tk
from tkinter import font as tkfont, ttk

BG = "#F3F6FA"
HEADER = "#EFF6FF"
WHITE = "#FFFFFF"
TEXT = "#172033"
MUTED = "#5B6575"
ACCENT = "#2563EB"
ACCENT_DARK = "#1D4ED8"
ACCENT_SOFT = "#DBEAFE"
BORDER = "#D6DEE8"
SUCCESS = "#15803D"
WARNING = "#B45309"
ERROR = "#B42318"


def high_contrast_enabled() -> bool:
    """Use native Windows colors if the high-contrast accessibility mode is on."""
    if os.name != "nt":
        return False
    try:
        from ctypes import wintypes

        class HighContrast(ctypes.Structure):
            _fields_ = (
                ("cbSize", wintypes.UINT),
                ("dwFlags", wintypes.DWORD),
                ("lpszDefaultScheme", wintypes.LPWSTR),
            )

        value = HighContrast()
        value.cbSize = ctypes.sizeof(value)
        return bool(
            ctypes.windll.user32.SystemParametersInfoW(0x0042, value.cbSize, ctypes.byref(value), 0)
            and value.dwFlags & 1
        )
    except (AttributeError, OSError, TypeError, ValueError):
        return False


def px(root: tk.Misc, value: int) -> int:
    """Scale drawn pixel ornaments/spacing, not fonts (Tk scales points itself)."""
    return max(1, round(value * float(root.tk.call("tk", "scaling")) / (96 / 72)))


def setup(root: tk.Tk) -> None:
    native = high_contrast_enabled()
    root._word_high_contrast = native
    colors = {
        "bg": "SystemButtonFace" if native else BG,
        "header": "SystemButtonFace" if native else HEADER,
        "card": "SystemWindow" if native else WHITE,
        "text": "SystemWindowText" if native else TEXT,
        "muted": "SystemWindowText" if native else MUTED,
        "border": "SystemWindowFrame" if native else BORDER,
    }
    root._word_colors = colors
    root.configure(bg=colors["bg"])
    style = ttk.Style(root)
    available = set(tkfont.families(root))
    family = "Yu Gothic UI" if "Yu Gothic UI" in available else "TkDefaultFont"
    root._word_font_family = family
    if native:
        for theme in ("vista", "xpnative", "default"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
    else:
        style.theme_use("clam")

    style.configure("App.TFrame", background=colors["bg"])
    style.configure("Header.TFrame", background=colors["header"])
    style.configure("Card.TFrame", background=colors["card"])
    style.configure("Metric.TFrame", background=colors["card"] if native else ACCENT_SOFT)
    style.configure("HeaderRule.TFrame", background=colors["border"] if native else ACCENT)
    style.configure("FooterRule.TFrame", background=colors["border"])
    for name, background, foreground, size, weight in (
        ("App.TLabel", "bg", "text", 10, "normal"),
        ("Muted.TLabel", "bg", "muted", 9, "normal"),
        ("HeaderTitle.TLabel", "header", "text", 21, "bold"),
        ("HeaderSub.TLabel", "header", "muted", 10, "normal"),
        ("Card.TLabel", "card", "text", 10, "normal"),
        ("CardMuted.TLabel", "card", "muted", 9, "normal"),
        ("Section.TLabel", "card", "text", 12, "bold"),
        ("MetricTitle.TLabel", "card", "muted", 9, "normal"),
        ("MetricValue.TLabel", "card", "text", 16, "bold"),
        ("Success.TLabel", "card", "text", 10, "normal"),
    ):
        style.configure(
            name, background=colors[background], foreground=colors[foreground],
            font=(family, size, weight),
        )
    if not native:
        style.configure("Success.TLabel", foreground=SUCCESS)
        style.configure("MetricTitle.TLabel", background=ACCENT_SOFT)
        style.configure("MetricValue.TLabel", background=ACCENT_SOFT)
    style.configure("App.TButton", font=(family, 10), padding=(px(root, 12), px(root, 7)))
    style.configure(
        "Primary.TButton", font=(family, 10, "bold"),
        background=ACCENT if not native else colors["card"],
        foreground=WHITE if not native else colors["text"],
        borderwidth=1, padding=(px(root, 18), px(root, 10)),
    )
    if not native:
        style.map(
            "Primary.TButton",
            background=[("disabled", "#BEC9DA"), ("pressed", ACCENT_DARK), ("active", ACCENT_DARK)],
            foreground=[("disabled", "#4F5B6C"), ("!disabled", WHITE)],
            bordercolor=[("focus", ACCENT_DARK), ("!focus", ACCENT)],
        )
        style.map(
            "App.TButton",
            background=[("disabled", "#ECEFF3"), ("active", ACCENT_SOFT)],
            foreground=[("disabled", "#85909F"), ("!disabled", TEXT)],
            bordercolor=[("focus", ACCENT)],
        )
    style.configure("App.TEntry", padding=px(root, 6), font=(family, 10))
    style.configure("Card.TCombobox", padding=px(root, 5), font=(family, 10))
    style.configure(
        "Card.TRadiobutton", background=colors["card"], foreground=colors["text"],
        font=(family, 10), padding=(px(root, 9), px(root, 6)),
    )
    if not native:
        style.map(
            "Card.TRadiobutton",
            background=[("selected", ACCENT_SOFT), ("active", HEADER), ("!selected", WHITE)],
            foreground=[("selected", ACCENT_DARK), ("!selected", TEXT)],
        )
    style.configure(
        "Horizontal.TProgressbar",
        background=colors["text"] if native else ACCENT,
        troughcolor=colors["bg"] if native else ACCENT_SOFT,
    )


class Card(tk.Canvas):
    """Rounded decoration wrapping a single native Frame; no bitmap dependencies."""

    def __init__(self, parent, padding: int = 18):
        root = self._root_of(parent)
        colors = getattr(root, "_word_colors", {"bg": BG, "card": WHITE, "border": BORDER})
        self._colors = colors
        self._native = getattr(root, "_word_high_contrast", False)
        self._radius = px(root, 12)
        self._inset = px(root, 5)
        super().__init__(
            parent, background=colors["bg"], borderwidth=0,
            highlightthickness=0, takefocus=0, height=px(root, 80),
        )
        self.content = ttk.Frame(
            self, style="Card.TFrame", padding=px(root, padding) - self._inset,
        )
        self.window = self.create_window(
            self._inset, self._inset, anchor="nw", window=self.content,
        )
        self._pending = None
        self._window_width = 0
        self._painted_size = (0, 0)
        self.content.bind("<Configure>", self._schedule)
        self.bind("<Configure>", self._schedule)
        self.bind("<Destroy>", self._destroyed)

    @staticmethod
    def _root_of(widget):
        return widget.winfo_toplevel()

    def _destroyed(self, event):
        if event.widget is self and self._pending is not None:
            try:
                self.after_cancel(self._pending)
            except tk.TclError:
                pass
            self._pending = None

    def _schedule(self, event=None):
        if self._pending is None:
            self._pending = self.after_idle(self._layout)

    def _layout(self):
        self._pending = None
        if not self.winfo_exists():
            return
        width = max(1, self.winfo_width())
        inner_width = max(1, width - 2 * self._inset)
        if inner_width != self._window_width:
            self.itemconfigure(self.window, width=inner_width)
            self._window_width = inner_width
        height = max(px(self, 56), self.content.winfo_reqheight() + 2 * self._inset)
        if self.winfo_reqheight() != height:
            self.configure(height=height)
        self._paint(width, max(1, self.winfo_height()))

    def _paint(self, width: int, height: int):
        if (width, height) == self._painted_size:
            return
        self._painted_size = (width, height)
        self.delete("art")
        if self._native:
            self.create_rectangle(
                1, 1, width - 2, height - 2, fill=self._colors["card"],
                outline=self._colors["border"], tags="art",
            )
        else:
            r = min(self._radius, max(1, width // 4), max(1, height // 4))
            self.create_polygon(
                r, 1, width-r, 1, width-1, r, width-1, height-r,
                width-r, height-1, r, height-1, 1, height-r, 1, r,
                smooth=True, splinesteps=20,
                fill=self._colors["card"], outline=self._colors["border"],
                width=1, tags="art",
            )
        self.tag_lower("art")
