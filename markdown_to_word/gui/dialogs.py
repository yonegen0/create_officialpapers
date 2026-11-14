"""User-facing dialogs, including the safe three-way overwrite decision."""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from .visuals import px


def error(title: str, message: str, parent) -> None:
    messagebox.showerror(title, message, parent=parent)


def info(title: str, message: str, parent) -> None:
    messagebox.showinfo(title, message, parent=parent)


def conflict(parent, filename: str | Path) -> str:
    """Return 'replace', 'rename' or 'cancel'; the initial/safe choice is cancel."""
    window = tk.Toplevel(parent)
    window.withdraw()
    window.title("同名ファイルの確認")
    window.transient(parent)
    window.resizable(True, False)
    choice = tk.StringVar(master=window, value="cancel")
    colors = getattr(parent, "_word_colors", {})
    background = colors.get("card", "#FFFFFF")
    foreground = colors.get("text", "#172033")
    window.configure(bg=background)

    frame = ttk.Frame(window, style="Card.TFrame", padding=px(parent, 18))
    frame.pack(fill="both", expand=True)
    ttk.Label(
        frame, text="同名のWordファイルが存在します。",
        style="Section.TLabel",
    ).pack(anchor="w")
    ttk.Label(frame, text="保存先（長い場合は左右にスクロールできます）", style="CardMuted.TLabel").pack(
        anchor="w", pady=(px(parent, 9), px(parent, 5))
    )
    path_var = tk.StringVar(master=window, value=str(filename))
    path_entry = ttk.Entry(
        frame, textvariable=path_var, state="readonly", style="App.TEntry",
    )
    path_entry.pack(fill="x")
    ttk.Label(frame, text="保存方法を選択してください。", style="Card.TLabel").pack(
        anchor="w", pady=(px(parent, 12), px(parent, 12)),
    )
    controls = ttk.Frame(frame, style="Card.TFrame")
    controls.pack(fill="x")

    def finish(value: str = "cancel") -> None:
        choice.set(value)
        window.destroy()

    for label, value in (
        ("置換する", "replace"), ("別名で保存", "rename"), ("キャンセル", "cancel"),
    ):
        button = ttk.Button(
            controls, text=label, style="App.TButton",
            command=lambda v=value: finish(v),
        )
        button.pack(side="left", padx=(0, px(parent, 8)))
        if value == "cancel":
            cancel_button = button

    window.protocol("WM_DELETE_WINDOW", lambda: finish("cancel"))
    window.bind("<Escape>", lambda e: finish("cancel"))
    window.update_idletasks()

    left = parent.winfo_rootx()
    top = parent.winfo_rooty()
    work_width = parent.winfo_screenwidth()
    work_height = parent.winfo_screenheight()
    width = min(max(window.winfo_reqwidth(), px(parent, 460)), max(1, work_width - px(parent, 24)))
    height = min(window.winfo_reqheight(), max(1, work_height - px(parent, 48)))
    x = max(0, min(left + (parent.winfo_width() - width) // 2, work_width - width))
    y = max(0, min(top + (parent.winfo_height() - height) // 2, work_height - height))
    window.geometry(f"{width}x{height}+{x}+{y}")
    window.deiconify()
    window.lift(parent)
    window.grab_set()
    cancel_button.focus_set()
    window.wait_window()
    return choice.get()
