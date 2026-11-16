"""Independent three-step Word→PDF tab using the existing card styling."""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk

from .visuals import Card, px, step_heading


class PdfTab:
    def __init__(self, parent, root) -> None:
        self.root = root
        self.controller = None
        self.output_dir = tk.StringVar(value="")
        self.output_name = tk.StringVar(value="")
        self.file_count = tk.StringVar(value="Wordファイルは未選択です")
        self.result = tk.StringVar(value="")
        self.batch = False
        self._scroll_pending = None
        colors = root._word_colors
        self.canvas = tk.Canvas(
            parent, background=colors["bg"], borderwidth=0, highlightthickness=0, takefocus=0,
        )
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = ttk.Frame(self.canvas, style="App.TFrame", padding=(px(root, 20), px(root, 18)))
        self.window = self.canvas.create_window(0, 0, window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._schedule_scroll)
        self.canvas.bind("<Configure>", self._resize)

        card = self._card("STEP 1  Wordファイルを選択")
        ttk.Label(card, text=".docxファイルを1件または複数選択できます。", style="CardMuted.TLabel").pack(anchor="w")
        self.file_button = ttk.Button(
            card, text="Wordファイルを選択…", style="App.TButton",
            command=lambda: self.controller.choose_files(),
        )
        self.file_button.pack(anchor="w", pady=(px(root, 10), px(root, 8)))
        ttk.Label(card, textvariable=self.file_count, style="CardMuted.TLabel").pack(anchor="w")
        listing = ttk.Frame(card, style="Card.TFrame")
        listing.pack(fill="x", pady=(px(root, 6), 0))
        self.file_list = tk.Listbox(
            listing, height=4, font=(root._word_font_family, 9),
            bg=colors["card"], fg=colors["text"], activestyle="none",
            relief="solid", borderwidth=1,
        )
        scroll = ttk.Scrollbar(listing, orient="vertical", command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=scroll.set)
        self.file_list.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.file_list.bind("<MouseWheel>", self._list_wheel)

        card = self._card("STEP 2  PDFの保存先を設定")
        ttk.Label(card, text="PDFのファイル名（1件のみ編集可能）", style="CardMuted.TLabel").pack(anchor="w")
        self.name_entry = ttk.Entry(card, textvariable=self.output_name, style="App.TEntry")
        self.name_entry.pack(fill="x", pady=(px(root, 5), px(root, 12)))
        ttk.Label(card, text="保存先フォルダー", style="CardMuted.TLabel").pack(anchor="w")
        folder_row = ttk.Frame(card, style="Card.TFrame")
        folder_row.pack(fill="x", pady=(px(root, 5), 0))
        self.folder_button = ttk.Button(
            folder_row, text="参照…", style="App.TButton",
            command=lambda: self.controller.choose_folder(),
        )
        self.folder_button.pack(side="right", padx=(px(root, 8), 0))
        self.dir_entry = ttk.Entry(folder_row, textvariable=self.output_dir, style="App.TEntry")
        self.dir_entry.pack(side="left", expand=True, fill="x")

        card = self._card("STEP 3  PDFに変換")
        description = ttk.Label(
            card, text="Microsoft Wordを使用して印刷レイアウトをPDF化します。元のWordは変更しません。",
            style="CardMuted.TLabel", justify="left",
        )
        description.pack(anchor="w", fill="x", pady=(0, px(root, 10)))
        card.bind(
            "<Configure>",
            lambda e: description.configure(wraplength=max(100, e.width - px(root, 12))),
            add="+",
        )
        self.convert_button = ttk.Button(
            card, text="PDFに変換", style="Primary.TButton",
            command=lambda: self.controller.convert(),
        )
        self.convert_button.pack(anchor="w")
        self.result_label = ttk.Label(
            card, textvariable=self.result, style="Success.TLabel", justify="left", anchor="w",
        )
        self.result_label.pack(fill="x", pady=(px(root, 12), px(root, 6)))
        card.bind(
            "<Configure>",
            lambda e: self.result_label.configure(wraplength=max(100, e.width - px(root, 12))),
            add="+",
        )
        actions = ttk.Frame(card, style="Card.TFrame")
        actions.pack(fill="x")
        self.open_button = ttk.Button(
            actions, text="PDFを開く", style="App.TButton", state="disabled",
            command=lambda: self.controller.open_pdf(),
        )
        self.open_button.pack(side="left", padx=(0, px(root, 8)))
        self.open_folder_button = ttk.Button(
            actions, text="保存先を開く", style="App.TButton", state="disabled",
            command=lambda: self.controller.open_folder(),
        )
        self.open_folder_button.pack(side="left")

    def _card(self, title):
        item = Card(self.inner, padding=16)
        item.pack(fill="x", pady=(0, px(self.root, 10)))
        step_heading(item.content, self.root, title)
        return item.content

    def _resize(self, event):
        width = max(1, event.width)
        if self.canvas.itemcget(self.window, "width") != str(width):
            self.canvas.itemconfigure(self.window, width=width)
        self._schedule_scroll()

    def _schedule_scroll(self, event=None):
        if self._scroll_pending is None:
            self._scroll_pending = self.canvas.after_idle(self._finish_scroll)

    def _finish_scroll(self):
        self._scroll_pending = None
        if self.canvas.winfo_exists():
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _list_wheel(self, event):
        self.file_list.yview_scroll(-int(event.delta / 120) if event.delta else 0, "units")
        return "break"

    def set_selection(self, files: tuple[Path, ...]) -> None:
        self.batch = len(files) > 1
        self.file_count.set(f"{len(files)}件選択中")
        self.file_list.delete(0, "end")
        for file in files:
            self.file_list.insert("end", file.name)
        self.output_dir.set(str(files[0].parent))
        self.output_name.set("元ファイル名を使用" if self.batch else files[0].stem + ".pdf")
        self.name_entry.configure(state="disabled" if self.batch else "normal")
        self.convert_button.configure(text=f"PDFを{len(files)}件まとめて変換" if self.batch else "PDFに変換")

    def set_busy(self, busy: bool) -> None:
        for widget in (self.file_button, self.folder_button, self.convert_button, self.dir_entry, self.name_entry):
            widget.configure(state="disabled" if busy else "normal")
        if not busy and self.batch:
            self.name_entry.configure(state="disabled")
        self._result_actions(busy)

    def _result_actions(self, busy: bool = False):
        result = self.controller.last_result if self.controller else None
        can_open = result is not None and result.total == 1 and len(result.saved) == 1
        can_folder = result is not None and bool(result.saved)
        self.open_button.configure(state="normal" if can_open and not busy else "disabled")
        self.open_folder_button.configure(state="normal" if can_folder and not busy else "disabled")

    def mark_previous(self):
        if self.controller and self.controller.last_result:
            prior = self.controller.last_result
            summary = f"前回の変換結果：成功 {len(prior.saved)}件・失敗 {len(prior.failures)}件"
            if prior.warnings:
                summary += f"・警告 {len(prior.warnings)}件"
            self.result.set(summary)

    def show_result(self, result):
        if result.processed == 0 and result.stopped_reason:
            lines = ["PDF変換を開始できませんでした", result.stopped_reason]
            if "Microsoft Wordを起動できません" in result.stopped_reason:
                lines.append("デスクトップ版Microsoft Wordがインストールされているか確認してください。")
            lines.append("変換開始前に停止しました。選択したWordファイルは変更されていません。")
        else:
            lines = [
                f"PDF変換{'停止' if result.stopped_reason else '完了'}：{result.total}件中"
                f"{len(result.saved)}件成功・{len(result.failures)}件失敗"
            ]
            lines.extend(f"成功：{path.name}" for path in result.saved)
            lines.extend(f"失敗：{path.name} — {reason}" for path, reason in result.failures)
            if result.stopped_reason:
                lines.append(f"処理停止：{result.stopped_reason}")
        if result.processed < result.total:
            lines.append(f"未処理：{result.total - result.processed}件")
        if result.warnings:
            lines.append(f"一時PDFの後始末に関する警告：{len(result.warnings)}件")
            lines.extend(f"警告：{path.name} — {reason}" for path, reason in result.warnings)
        lines.append(f"保存先：{result.folder}")
        self.result.set("\n".join(lines))
        if result.stopped_reason and not result.saved:
            self.result_label.configure(style="Error.TLabel")
        elif result.failures or result.warnings or result.stopped_reason:
            self.result_label.configure(style="Warning.TLabel")
        else:
            self.result_label.configure(style="Success.TLabel")
        self._result_actions()
