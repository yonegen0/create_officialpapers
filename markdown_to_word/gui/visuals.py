"""Small, standalone adaptation of the job tool's soft card visual language."""
import tkinter as tk
from tkinter import ttk

BG = "#F3F6F8"
WHITE = "#FFFFFF"
TEXT = "#243442"
MUTED = "#586979"
ACCENT = "#347B77"
BORDER = "#DAE4E8"


def setup(root: tk.Tk) -> None:
    root.configure(bg=BG)
    s = ttk.Style(root)
    s.theme_use("clam")
    s.configure("App.TFrame", background=BG)
    s.configure("Card.TFrame", background=WHITE)
    s.configure("App.TLabel", background=BG, foreground=TEXT, font=("Yu Gothic UI", 10))
    s.configure("Muted.TLabel", background=BG, foreground=MUTED, font=("Yu Gothic UI", 9))
    s.configure("Card.TLabel", background=WHITE, foreground=TEXT, font=("Yu Gothic UI", 10))
    s.configure("Section.TLabel", background=WHITE, foreground=TEXT, font=("Yu Gothic UI", 12, "bold"))
    s.configure("App.TButton", font=("Yu Gothic UI", 10), padding=(14, 9))
    s.configure("Primary.TButton", font=("Yu Gothic UI", 10, "bold"), padding=(17, 10), background=ACCENT, foreground=WHITE, borderwidth=0)
    s.map("Primary.TButton", background=[("active", "#246762"), ("disabled", "#B6CBCA")])
    s.configure("App.TEntry", padding=6)
    s.configure("Card.TRadiobutton", background=WHITE, foreground=TEXT, font=("Yu Gothic UI", 10))
    s.configure("Card.TCombobox", padding=5)
    s.configure("Horizontal.TProgressbar", background=ACCENT, troughcolor=BORDER)


class Card(tk.Canvas):
    """Rounded background with a content Frame as the single window item."""
    def __init__(self, parent, padding: int = 18):
        super().__init__(parent, background=BG, highlightthickness=0, borderwidth=0)
        self.padding = padding
        self.content = ttk.Frame(self, style="Card.TFrame", padding=padding)
        self.window = self.create_window(0, 0, anchor="nw", window=self.content)
        self.content.bind("<Configure>", self._size)
        self.bind("<Configure>", self._draw)

    def _size(self, event=None):
        req = self.content.winfo_reqheight()
        self.configure(height=req + 2)

    def _draw(self, event=None):
        w = max(1, self.winfo_width())
        h = max(1, self.winfo_height())
        self.delete("art")
        # A visually round card without large image resources/dependencies.
        r = min(14, h // 3, w // 3)
        self.create_polygon(
            r, 1, w-r, 1, w-1, r, w-1, h-r, w-r, h-1,
            r, h-1, 1, h-r, 1, r,
            smooth=True, splinesteps=24, fill=WHITE, outline=BORDER,
            width=1, tags="art",
        )
        self.tag_lower("art")
        self.coords(self.window, 1, 1)
        self.itemconfigure(self.window, width=max(w-2, 1))
