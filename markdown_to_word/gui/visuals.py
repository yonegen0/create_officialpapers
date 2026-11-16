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
    high_contrast = high_contrast_enabled()
    root._word_high_contrast = high_contrast
    colors = {
        "bg": "SystemButtonFace" if high_contrast else BG,
        "header": "SystemButtonFace" if high_contrast else HEADER,
        "card": "SystemWindow" if high_contrast else WHITE,
        "text": "SystemWindowText" if high_contrast else TEXT,
        "muted": "SystemWindowText" if high_contrast else MUTED,
        "border": "SystemWindowFrame" if high_contrast else BORDER,
    }
    root._word_colors = colors
    root.configure(bg=colors["bg"])
    style = ttk.Style(root)
    for theme in ("vista", "xpnative", "clam", "default", *style.theme_names()):
        if theme not in style.theme_names():
            continue
        try:
            style.theme_use(theme)
            if style.theme_use() == theme:
                break
        except tk.TclError:
            continue
    theme = style.theme_use()
    root._word_theme = theme
    native_controls = theme in ("vista", "xpnative")
    available = set(tkfont.families(root))
    family = "Yu Gothic UI" if "Yu Gothic UI" in available else tkfont.nametofont("TkDefaultFont").cget("family")
    root._word_font_family = family

    style.configure("App.TFrame", background=colors["bg"])
    style.configure("App.TNotebook", background=colors["bg"], borderwidth=0)
    style.configure("App.TNotebook.Tab", font=(family, 10, "bold"), padding=(px(root, 16), px(root, 8)))
    style.configure("Header.TFrame", background=colors["header"])
    style.configure("Card.TFrame", background=colors["card"])
    style.configure("Metric.TFrame", background=colors["card"] if high_contrast else ACCENT_SOFT)
    style.configure("HeaderRule.TFrame", background=colors["border"] if high_contrast else ACCENT)
    style.configure("FooterRule.TFrame", background=colors["border"])
    style.configure("SectionAccent.TFrame", background=colors["text"] if high_contrast else ACCENT)
    style.configure("MetricAccent.TFrame", background=colors["text"] if high_contrast else ACCENT)
    for name, background, foreground, size, weight in (
        ("App.TLabel", "bg", "text", 10, "normal"),
        ("Muted.TLabel", "bg", "muted", 9, "normal"),
        ("HeaderTitle.TLabel", "header", "text", 22, "bold"),
        ("HeaderSub.TLabel", "header", "muted", 10, "normal"),
        ("HeaderCaption.TLabel", "header", "muted", 9, "normal"),
        ("Card.TLabel", "card", "text", 10, "normal"),
        ("CardMuted.TLabel", "card", "muted", 9, "normal"),
        ("Section.TLabel", "card", "text", 11, "bold"),
        ("MetricTitle.TLabel", "card", "muted", 9, "normal"),
        ("MetricValue.TLabel", "card", "text", 16, "bold"),
        ("Success.TLabel", "card", "text", 10, "normal"),
        ("Warning.TLabel", "card", "text", 10, "normal"),
        ("Error.TLabel", "card", "text", 10, "normal"),
    ):
        style.configure(
            name, background=colors[background], foreground=colors[foreground],
            font=(family, size, weight),
        )
    if not high_contrast:
        for name, foreground in (("Success.TLabel", SUCCESS), ("Warning.TLabel", WARNING), ("Error.TLabel", ERROR)):
            style.configure(name, foreground=foreground)
        for name in ("MetricTitle.TLabel", "MetricValue.TLabel"):
            style.configure(name, background=ACCENT_SOFT)

    style.configure("App.TButton", font=(family, 10), padding=(px(root, 12), px(root, 7)))
    style.configure("Primary.TButton", font=(family, 10, "bold"), padding=(px(root, 18), px(root, 9)))
    if not high_contrast and theme == "clam":
        style.configure("Primary.TButton", background=ACCENT, foreground=WHITE, borderwidth=1)
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
    elif not high_contrast and native_controls:
        disabled = style.lookup("TButton", "foreground", ("disabled",)) or "SystemGrayText"
        # Native themes ignore arbitrary backgrounds; never pair a white label with one.
        style.configure("Primary.TButton", foreground=ACCENT)
        style.map(
            "Primary.TButton",
            foreground=[("disabled", disabled), ("pressed", ACCENT_DARK),
                        ("active", ACCENT_DARK), ("!disabled", ACCENT)],
        )
    if not high_contrast and native_controls:
        disabled_tab = style.lookup("TNotebook.Tab", "foreground", ("disabled",)) or "SystemGrayText"
        style.configure("App.TNotebook.Tab", foreground=TEXT)
        style.map(
            "App.TNotebook.Tab",
            foreground=[("disabled", disabled_tab), ("selected", ACCENT),
                        ("active", ACCENT_DARK), ("!disabled", TEXT)],
        )
    elif not high_contrast and theme == "clam":
        style.configure("App.TNotebook.Tab", foreground=TEXT)
        style.map(
            "App.TNotebook.Tab",
            foreground=[("disabled", MUTED), ("selected", ACCENT),
                        ("active", ACCENT_DARK), ("!disabled", TEXT)],
        )
    style.configure("App.TEntry", padding=px(root, 6), font=(family, 10))
    style.configure("Card.TCombobox", padding=px(root, 5), font=(family, 10))
    style.configure(
        "Card.TRadiobutton", background=colors["card"], foreground=colors["text"],
        font=(family, 10), padding=(px(root, 9), px(root, 6)),
    )
    if not high_contrast and theme == "clam":
        style.map(
            "Card.TRadiobutton",
            background=[("selected", ACCENT_SOFT), ("active", HEADER), ("!selected", WHITE)],
            foreground=[("selected", ACCENT_DARK), ("!selected", TEXT)],
        )
    style.configure(
        "Horizontal.TProgressbar",
        background=colors["text"] if high_contrast else ACCENT,
        troughcolor=colors["bg"] if high_contrast else ACCENT_SOFT,
    )


def step_heading(parent: ttk.Frame, root: tk.Misc, title: str) -> None:
    """Shared compact STEP header for both conversion tabs."""
    row = ttk.Frame(parent, style="Card.TFrame")
    row.pack(fill="x", pady=(0, px(root, 9)))
    ttk.Frame(
        row, style="SectionAccent.TFrame", width=px(root, 4), height=px(root, 19),
    ).pack(side="left", padx=(0, px(root, 9)))
    ttk.Label(row, text=title, style="Section.TLabel").pack(side="left", anchor="w")


def _card_corners(root: tk.Misc, colors: dict, radius: int, border: int):
    """Small, shared antialiased corner images; no Pillow or full-card bitmaps."""
    key = (radius, border, colors["bg"], colors["card"], colors["border"])
    cache = getattr(root, "_word_card_corners", None)
    if cache is None:
        cache = root._word_card_corners = {}
    if key in cache:
        return cache[key]

    def rgb(color: str):
        return tuple(int(color[index:index + 2], 16) for index in (1, 3, 5))

    def inside(x: float, y: float, side: int, inset: int, curve: int) -> bool:
        low, high = inset, side - inset
        if not (low <= x < high and low <= y < high):
            return False
        r = max(0, min(curve, (high - low) / 2))
        cx = min(max(x, low + r), high - r)
        cy = min(max(y, low + r), high - r)
        return (x - cx) ** 2 + (y - cy) ** 2 <= r ** 2

    cut = radius + border
    side = cut * 2
    outside, edge, surface = (rgb(colors[name]) for name in ("bg", "border", "card"))
    samples = 4
    rows = []
    for y in range(side):
        row = []
        for x in range(side):
            channels = [0, 0, 0]
            for iy in range(samples):
                for ix in range(samples):
                    sy = y + (iy + 0.5) / samples
                    sx = x + (ix + 0.5) / samples
                    if inside(sx, sy, side, border, radius - border):
                        color = surface
                    elif inside(sx, sy, side, 0, radius):
                        color = edge
                    else:
                        color = outside
                    for channel in range(3):
                        channels[channel] += color[channel]
            row.append("#%02X%02X%02X" % tuple(round(v / samples ** 2) for v in channels))
        rows.append(tuple(row))
    image = tk.PhotoImage(master=root, width=side, height=side)
    image.put(tuple(rows))
    corners = []
    for x, y in ((0, 0), (cut, 0), (0, cut), (cut, cut)):
        corner = tk.PhotoImage(master=root, width=cut, height=cut)
        corner.tk.call(str(corner), "copy", str(image), "-from", x, y, x + cut, y + cut)
        corners.append(corner)
    cache[key] = (cut, tuple(corners))
    return cache[key]


class Card(tk.Canvas):
    """Rounded canvas decoration around a compatible native ttk content Frame."""

    def __init__(self, parent, padding: int = 18):
        root = self._root_of(parent)
        colors = getattr(root, "_word_colors", {"bg": BG, "card": WHITE, "border": BORDER})
        self._colors = colors
        self._high_contrast = getattr(root, "_word_high_contrast", False)
        self._radius = px(root, 12)
        self._border = px(root, 1)
        self._inset = px(root, 5)
        self._corners = None
        if not self._high_contrast:
            try:
                self._corners = _card_corners(root, colors, self._radius, self._border)
            except (tk.TclError, ValueError):
                pass
        super().__init__(
            parent, background=colors["bg"], borderwidth=0,
            highlightthickness=0, takefocus=0, height=px(root, 80),
        )
        self.content = ttk.Frame(
            self, style="Card.TFrame", padding=max(0, px(root, padding) - self._inset),
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
        try:
            if self._corners is None or width < 2 * self._corners[0] or height < 2 * self._corners[0]:
                self._draw_rectangle(width, height)
            else:
                cut, images = self._corners
                card, border = self._colors["card"], self._colors["border"]
                self.create_rectangle(cut, 0, width-cut, height, fill=card, outline="", tags="art")
                self.create_rectangle(0, cut, width, height-cut, fill=card, outline="", tags="art")
                for xy in (
                    (cut, 0, width-cut, self._border),
                    (cut, height-self._border, width-cut, height),
                    (0, cut, self._border, height-cut),
                    (width-self._border, cut, width, height-cut),
                ):
                    self.create_rectangle(*xy, fill=border, outline="", tags="art")
                for img, (x, y) in zip(images, (
                    (0, 0), (width-cut, 0), (0, height-cut), (width-cut, height-cut),
                )):
                    self.create_image(x, y, image=img, anchor="nw", tags="art")
        except tk.TclError:
            self.delete("art")
            self._draw_rectangle(width, height)
        self.tag_lower("art")

    def _draw_rectangle(self, width: int, height: int) -> None:
        if width > 2 and height > 2:
            self.create_rectangle(
                1, 1, width - 2, height - 2, fill=self._colors["card"],
                outline=self._colors["border"], tags="art",
            )
