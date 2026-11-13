"""Compact, scrollable, two-input-mode Windows Tk interface."""
from __future__ import annotations

import os
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk

from . import dialogs
from .controller import Controller
from .visuals import BG, WHITE, TEXT, MUTED, ACCENT, Card, setup


class TkView:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        setup(root)
        self.input_mode = tk.StringVar(value="file")
        self.file_path = tk.StringVar()
        self.style = tk.StringVar(value="職務経歴書")
        self.output_name = tk.StringVar(value=f"文書_{datetime.now():%Y%m%d}.docx")
        self.output_dir = tk.StringVar(value=str(Path.home() / "Documents" if (Path.home() / "Documents").is_dir() else Path.home()))
        self.status = tk.StringVar(value="Markdownを入力して事前確認してください")
        self.info = tk.StringVar(value="文字数：—   見出し：—   表：—   箇条書き：—")
        self.warning = tk.StringVar(value="")
        self.success = tk.StringVar(value="")
        self._build()
        self.controller = Controller(self)
        for var in (self.file_path, self.style, self.input_mode):
            var.trace_add("write", self.controller.invalidate)
        self.input_mode.trace_add("write", self._show_mode)
        self.input_text.bind("<<Modified>>", self._on_text_changed)
        self.input_text.edit_modified(False)
        self._show_mode()
        self.root.protocol("WM_DELETE_WINDOW", self.controller.close)

    def _build(self) -> None:
        self.root.title("Markdown → Word 変換ツール")
        self.root.geometry("960x700")
        self.root.minsize(760, 600)
        self.root.configure(bg=BG)
        body = ttk.Frame(self.root, style="App.TFrame")
        body.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(body, background=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = ttk.Frame(self.canvas, style="App.TFrame", padding=(22, 16))
        window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda _: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(window, width=e.width))
        self.canvas.bind_all("<MouseWheel>", self._wheel, add="+")
        title = tk.Label(self.inner, text="Markdown → Word 変換ツール", bg=BG, fg=TEXT, font=("Yu Gothic UI", 20, "bold"), anchor="w")
        title.pack(fill="x", pady=(2, 3))
        ttk.Label(self.inner, text="職務経歴書を読みやすく、編集できるWordへ変換します。", style="Muted.TLabel").pack(anchor="w")
        ttk.Label(self.inner, text="✓ ファイルはこのPC内だけで処理します", style="Muted.TLabel").pack(anchor="w", pady=(3, 12))

        card = self._card("Markdown入力")
        selector = ttk.Frame(card, style="Card.TFrame")
        selector.pack(fill="x", pady=(4, 12))
        self.mode_widgets = []
        for label, mode in (("ファイルを選択", "file"), ("直接貼り付け", "paste")):
            w = ttk.Radiobutton(selector, text=label, variable=self.input_mode, value=mode, style="Card.TRadiobutton")
            w.pack(side="left", padx=(0, 20))
            self.mode_widgets.append(w)
        self.file_section = ttk.Frame(card, style="Card.TFrame")
        self.file_button = ttk.Button(self.file_section, text="Markdownファイルを参照…", style="App.TButton")
        self.file_button.pack(anchor="w", pady=(0, 7))
        self.file_name = ttk.Label(self.file_section, textvariable=self.file_path, style="Card.TLabel", wraplength=720)
        self.file_name.pack(anchor="w", fill="x")
        self.paste_section = ttk.Frame(card, style="Card.TFrame")
        ttk.Label(self.paste_section, text="ここにMarkdownを直接貼り付けてください", style="Card.TLabel").pack(anchor="w", pady=(0, 5))
        paste_box = ttk.Frame(self.paste_section, style="Card.TFrame")
        paste_box.pack(fill="x")
        self.input_text = tk.Text(paste_box, height=11, wrap="word", font=("Yu Gothic UI", 10), undo=True, relief="solid", borderwidth=1)
        text_scroll = ttk.Scrollbar(paste_box, orient="vertical", command=self.input_text.yview)
        self.input_text.configure(yscrollcommand=text_scroll.set)
        self.input_text.pack(side="left", fill="both", expand=True)
        text_scroll.pack(side="right", fill="y")
        self.input_text.bind("<Control-a>", lambda e: (self.input_text.tag_add("sel", "1.0", "end-1c"), "break")[-1])
        self.input_text.bind("<Control-A>", lambda e: (self.input_text.tag_add("sel", "1.0", "end-1c"), "break")[-1])

        card = self._card("Word設定")
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x", pady=(4, 8))
        ttk.Label(row, text="文書スタイル", style="Card.TLabel", width=15).pack(side="left")
        self.style_combo = ttk.Combobox(row, textvariable=self.style, values=("職務経歴書", "標準文書"), state="readonly", width=22)
        self.style_combo.pack(side="left")
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x", pady=(0, 8))
        ttk.Label(row, text="出力ファイル名", style="Card.TLabel", width=15).pack(side="left")
        self.name_entry = ttk.Entry(row, textvariable=self.output_name)
        self.name_entry.pack(side="left", fill="x", expand=True)
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x")
        ttk.Label(row, text="保存先", style="Card.TLabel", width=15).pack(side="left")
        self.dir_entry = ttk.Entry(row, textvariable=self.output_dir)
        self.dir_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.folder_button = ttk.Button(row, text="参照…", style="App.TButton")
        self.folder_button.pack(side="left")

        card = self._card("事前確認")
        ttk.Label(card, textvariable=self.info, style="Card.TLabel").pack(anchor="w", pady=(3, 5))
        self.warn_label = tk.Label(card, textvariable=self.warning, bg=WHITE, fg="#95621C", justify="left", anchor="w", wraplength=720)
        self.warn_label.pack(fill="x")

        card = self._card("変換を実行")
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(anchor="w", pady=(4, 0))
        self.preview_button = ttk.Button(row, text="事前確認", style="App.TButton")
        self.preview_button.pack(side="left", padx=(0, 8))
        self.generate_button = ttk.Button(row, text="Wordを生成", style="Primary.TButton")
        self.generate_button.pack(side="left")
        ttk.Label(card, textvariable=self.success, style="Card.TLabel", wraplength=710).pack(anchor="w", pady=(10, 0))
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(anchor="w", pady=(6, 0))
        self.open_button = ttk.Button(row, text="Wordを開く", style="App.TButton", state="disabled")
        self.open_button.pack(side="left", padx=(0, 8))
        self.open_folder_button = ttk.Button(row, text="保存先を開く", style="App.TButton", state="disabled")
        self.open_folder_button.pack(side="left")

        footer = ttk.Frame(self.root, style="App.TFrame", padding=(20, 8))
        footer.pack(side="bottom", fill="x")
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=100)
        self.progress.pack(side="left", padx=(0, 12))
        ttk.Label(footer, textvariable=self.status, style="Muted.TLabel").pack(side="left")
        self.exit_button = ttk.Button(footer, text="終了", style="App.TButton")
        self.exit_button.pack(side="right")
        self.help_button = ttk.Button(footer, text="使い方", style="App.TButton", command=self._help)
        self.help_button.pack(side="right", padx=(0, 8))

        self.file_button.configure(command=lambda: self.controller.choose_file())
        self.folder_button.configure(command=lambda: self.controller.choose_folder())
        self.preview_button.configure(command=lambda: self.controller.confirm())
        self.generate_button.configure(command=lambda: self.controller.generate())
        self.open_button.configure(command=lambda: self.controller.open_word())
        self.open_folder_button.configure(command=lambda: self.controller.open_folder())
        self.exit_button.configure(command=lambda: self.controller.close())

    def _card(self, heading: str):
        item = Card(self.inner)
        item.pack(fill="x", pady=(0, 13))
        ttk.Label(item.content, text=heading, style="Section.TLabel").pack(anchor="w", pady=(0, 9))
        return item.content

    def _show_mode(self, *args):
        self.file_section.pack_forget()
        self.paste_section.pack_forget()
        if self.input_mode.get() == "file":
            self.file_section.pack(fill="x")
        else:
            self.paste_section.pack(fill="x")

    def _on_text_changed(self, event):
        if self.input_text.edit_modified():
            self.input_text.edit_modified(False)
            self.controller.invalidate()

    def _wheel(self, event):
        if isinstance(event.widget, tk.Text) and event.widget == self.input_text and self.input_mode.get() == "paste":
            return None
        self.canvas.yview_scroll(-int(event.delta / 120), "units")
        return "break"

    def _help(self):
        dialogs.info("使い方", "① ファイルを選ぶかMarkdownを貼り付ける\n② 文書スタイル・保存先を選ぶ\n③ 事前確認を実行する\n④ Wordを生成する\n\n既存ファイルがある場合は、置換・別名保存・キャンセルを選べます。", self.root)

    def show_preview(self, preview):
        d = preview.document
        self.info.set(f"文字数：{d.characters:,}   見出し：{d.headings}   表：{d.tables}   箇条書き：{d.list_items}")
        self.warning.set("\n".join("注意: " + w for w in d.warnings) or "警告なし")

    def clear_preview(self):
        self.info.set("文字数：—   見出し：—   表：—   箇条書き：—")
        self.warning.set("")

    def set_status(self, message: str):
        self.status.set(message)

    def set_busy(self, busy: bool):
        for widget in (self.file_button, self.folder_button, self.preview_button, self.generate_button, self.help_button, self.exit_button, self.name_entry, self.dir_entry, *self.mode_widgets):
            widget.configure(state="disabled" if busy else "normal")
        self.style_combo.configure(state="disabled" if busy else "readonly")
        self.input_text.configure(state="disabled" if busy else "normal")
        if busy:
            self.progress.start(14)
        else:
            self.progress.stop()

    def show_success(self, path: Path):
        self.success.set(f"保存しました: {path}")
        self.open_button.configure(state="normal")
        self.open_folder_button.configure(state="normal")
