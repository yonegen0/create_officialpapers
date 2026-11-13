"""User-facing dialogs, including explicitly named overwrite choices."""
import tkinter as tk
from tkinter import messagebox, ttk


def error(title: str, message: str, parent) -> None:
    messagebox.showerror(title, message, parent=parent)


def info(title: str, message: str, parent) -> None:
    messagebox.showinfo(title, message, parent=parent)


def conflict(parent, filename: str) -> str:
    """Return replace, rename, or cancel."""
    window = tk.Toplevel(parent)
    window.title("同名ファイルの確認")
    window.transient(parent)
    window.resizable(False, False)
    window.geometry("400x155")
    choice = tk.StringVar(value="cancel")
    frame = ttk.Frame(window, padding=18)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text=f"{filename}\n同名のWordがあります。保存方法を選択してください。", wraplength=360).pack(anchor="w", pady=(0, 18))
    controls = ttk.Frame(frame)
    controls.pack(fill="x")
    for label, value in (("置換する", "replace"), ("別名で保存", "rename"), ("キャンセル", "cancel")):
        ttk.Button(controls, text=label, command=lambda v=value: (choice.set(v), window.destroy())).pack(side="left", padx=(0, 6))
    window.protocol("WM_DELETE_WINDOW", window.destroy)
    window.grab_set()
    window.wait_window()
    return choice.get()
