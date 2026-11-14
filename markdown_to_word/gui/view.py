"""Single-page, three-step Windows Tk interface for Markdown-to-Word."""
from __future__ import annotations

import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk

from . import dialogs
from .controller import Controller
from .visuals import (
    ACCENT, ACCENT_SOFT, BG, BORDER, MUTED, TEXT, WARNING, WHITE,
    Card, px, setup,
)


class TkView:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        setup(root)
        self.input_mode = tk.StringVar(value="file")
        self.file_path = tk.StringVar()
        self.file_count = tk.StringVar(value="")
        self._batch_mode = False
        self._single_output_name = ""
        self.style = tk.StringVar(value="職務経歴書")
        self.output_name = tk.StringVar(value=f"文書_{datetime.now():%Y%m%d}.docx")
        home = Path.home()
        self.output_dir = tk.StringVar(value=str(home / "Documents" if (home / "Documents").is_dir() else home))
        self.status = tk.StringVar(value="Markdownを入力して事前確認してください")
        self.info = tk.StringVar(value="文字数：—   見出し：—   表：—   箇条書き：—")
        self.warning = tk.StringVar(value="")
        self.success = tk.StringVar(value="")
        self.metric_values = {key: tk.StringVar(value="—") for key in ("characters", "headings", "tables", "list_items")}
        self._settings_columns = None
        self._metric_columns = None
        self._scroll_pending = None
        self._build()
        self.controller = Controller(self)
        for var in (self.file_path, self.style, self.input_mode):
            var.trace_add("write", self.controller.invalidate)
        self.input_mode.trace_add("write", self._show_mode)
        self.input_text.bind("<<Modified>>", self._on_text_changed)
        self.input_text.edit_modified(False)
        self.input_text.bind("<MouseWheel>", self._text_wheel)
        self.root.bind("<MouseWheel>", self._wheel, add="+")
        self._show_mode()
        self.root.protocol("WM_DELETE_WINDOW", self.controller.close)

    def _build(self) -> None:
        self.root.title("Markdown → Word 変換ツール")
        colors = self.root._word_colors
        margin = px(self.root, 20)

        header = ttk.Frame(self.root, style="Header.TFrame", padding=(margin, px(self.root, 14)))
        header.pack(side="top", fill="x")
        ttk.Label(header, text="Markdown → Word 変換ツール", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(header, text="編集可能なWord文書をかんたん作成", style="HeaderSub.TLabel").pack(anchor="w")
        ttk.Label(header, text="ファイルはこのPC内だけで処理します", style="HeaderSub.TLabel").pack(anchor="w", pady=(px(self.root, 3), 0))
        ttk.Frame(self.root, height=px(self.root, 3), style="HeaderRule.TFrame").pack(side="top", fill="x")

        ttk.Frame(self.root, height=1, style="FooterRule.TFrame").pack(side="bottom", fill="x")
        footer = ttk.Frame(self.root, style="App.TFrame", padding=(margin, px(self.root, 8)))
        footer.pack(side="bottom", fill="x")
        footer.columnconfigure(1, weight=1)
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=px(self.root, 90))
        self.progress.grid(row=0, column=0, sticky="w", padx=(0, px(self.root, 10)))
        self.status_label = ttk.Label(footer, textvariable=self.status, style="Muted.TLabel")
        self.status_label.grid(row=0, column=1, sticky="ew")
        self.help_button = ttk.Button(footer, text="使い方", style="App.TButton", command=self._help)
        self.help_button.grid(row=0, column=2, padx=(px(self.root, 8), px(self.root, 6)))
        self.exit_button = ttk.Button(footer, text="終了", style="App.TButton")
        self.exit_button.grid(row=0, column=3)

        body = ttk.Frame(self.root, style="App.TFrame")
        body.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(
            body, background=colors["bg"], borderwidth=0, highlightthickness=0,
            takefocus=0,
        )
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = ttk.Frame(
            self.canvas, style="App.TFrame",
            padding=(margin, px(self.root, 18)),
        )
        self._canvas_window = self.canvas.create_window(0, 0, window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._update_scrollregion)
        self.canvas.bind("<Configure>", self._canvas_resized)

        card = self._card("STEP 1  Markdownを入力")
        ttk.Label(
            card, text="Markdownファイルを選ぶか、文章を直接貼り付けます。",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(0, px(self.root, 10)))
        selector = ttk.Frame(card, style="Card.TFrame")
        selector.pack(fill="x", pady=(0, px(self.root, 10)))
        self.mode_widgets = []
        for label, mode in (("ファイルを選択", "file"), ("直接貼り付け", "paste")):
            button = ttk.Radiobutton(
                selector, text=label, value=mode, variable=self.input_mode,
                style="Card.TRadiobutton",
            )
            button.pack(side="left", padx=(0, px(self.root, 8)))
            self.mode_widgets.append(button)
        self.file_section = ttk.Frame(card, style="Card.TFrame")
        self.file_button = ttk.Button(
            self.file_section, text="Markdownファイルを参照…", style="App.TButton",
        )
        self.file_button.pack(anchor="w", pady=(0, px(self.root, 8)))
        ttk.Label(self.file_section, text="選択したファイル（横スクロールで確認できます）", style="CardMuted.TLabel").pack(anchor="w")
        self.file_name = ttk.Entry(
            self.file_section, textvariable=self.file_path, state="readonly",
            style="App.TEntry",
        )
        self.file_name.pack(fill="x", pady=(px(self.root, 4), 0))
        ttk.Label(self.file_section, textvariable=self.file_count, style="CardMuted.TLabel").pack(
            anchor="w", pady=(px(self.root, 5), 0),
        )
        self.file_list_frame = ttk.Frame(self.file_section, style="Card.TFrame")
        self.file_list = tk.Listbox(
            self.file_list_frame, height=4, font=(self.root._word_font_family, 9),
            bg=colors["card"], fg=colors["text"], selectmode="browse",
            activestyle="none", relief="solid", borderwidth=1,
        )
        file_scroll = ttk.Scrollbar(self.file_list_frame, orient="vertical", command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=file_scroll.set)
        self.file_list.pack(side="left", fill="both", expand=True)
        file_scroll.pack(side="right", fill="y")
        self.file_list.bind("<MouseWheel>", self._file_list_wheel)

        self.paste_section = ttk.Frame(card, style="Card.TFrame")
        ttk.Label(
            self.paste_section, text="ここにMarkdownを直接貼り付けてください",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(0, px(self.root, 6)))
        paste_box = ttk.Frame(self.paste_section, style="Card.TFrame")
        paste_box.pack(fill="x")
        self.input_text = tk.Text(
            paste_box, height=9, wrap="word",
            font=(self.root._word_font_family, 10),
            foreground=colors["text"], background=colors["card"],
            insertbackground=colors["text"], undo=True, relief="solid",
            borderwidth=1,
        )
        paste_scroll = ttk.Scrollbar(paste_box, orient="vertical", command=self.input_text.yview)
        self.input_text.configure(yscrollcommand=paste_scroll.set)
        self.input_text.pack(side="left", fill="both", expand=True)
        paste_scroll.pack(side="right", fill="y")
        for sequence in ("<Control-a>", "<Control-A>"):
            self.input_text.bind(sequence, self._select_all_text)

        card = self._card("STEP 2  Wordの出力設定")
        self.settings = ttk.Frame(card, style="Card.TFrame")
        self.settings.pack(fill="x")
        self.setting_groups = []
        style_group = self._setting_group("文書スタイル")
        self.style_combo = ttk.Combobox(
            style_group, textvariable=self.style,
            values=("職務経歴書", "標準文書"), state="readonly",
            style="Card.TCombobox",
        )
        self.style_combo.pack(fill="x")
        name_group = self._setting_group("出力ファイル名")
        self.name_entry = ttk.Entry(
            name_group, textvariable=self.output_name, style="App.TEntry",
        )
        self.name_entry.pack(fill="x")
        folder_group = self._setting_group("保存先")
        folder_row = ttk.Frame(folder_group, style="Card.TFrame")
        folder_row.pack(fill="x")
        self.folder_button = ttk.Button(folder_row, text="参照…", style="App.TButton")
        self.folder_button.pack(side="right", padx=(px(self.root, 6), 0))
        self.dir_entry = ttk.Entry(
            folder_row, textvariable=self.output_dir, style="App.TEntry",
        )
        self.dir_entry.pack(side="left", fill="x", expand=True)
        self.settings.bind("<Configure>", self._adapt_settings)

        card = self._card("STEP 3  事前確認・Word生成")
        ttk.Label(
            card, text="内容を確認してからWordを生成します。",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(0, px(self.root, 8)))
        self.metrics = ttk.Frame(card, style="Card.TFrame")
        self.metrics.pack(fill="x", pady=(0, px(self.root, 10)))
        self.metric_tiles = []
        self.metric_titles = {}
        for key, label in (
            ("characters", "文字数"), ("headings", "見出し"),
            ("tables", "表"), ("list_items", "箇条書き"),
        ):
            tile = ttk.Frame(self.metrics, style="Metric.TFrame", padding=px(self.root, 10))
            title = ttk.Label(tile, text=label, style="MetricTitle.TLabel")
            title.pack(anchor="w")
            self.metric_titles[key] = title
            ttk.Label(tile, textvariable=self.metric_values[key], style="MetricValue.TLabel").pack(anchor="w")
            self.metric_tiles.append(tile)
        self.metrics.bind("<Configure>", self._adapt_metrics)
        self.warn_label = tk.Label(
            card, textvariable=self.warning, background=colors["card"],
            foreground=colors["text"] if self.root._word_high_contrast else WARNING,
            font=(self.root._word_font_family, 9), justify="left",
            anchor="w", wraplength=px(self.root, 690),
        )
        self.warn_label.pack(fill="x", pady=(0, px(self.root, 10)))
        card.bind(
            "<Configure>",
            lambda e: self.warn_label.configure(wraplength=max(100, e.width - px(self.root, 12))),
            add="+",
        )
        actions = ttk.Frame(card, style="Card.TFrame")
        actions.pack(fill="x")
        self.preview_button = ttk.Button(actions, text="事前確認", style="App.TButton")
        self.preview_button.pack(side="left", padx=(0, px(self.root, 8)))
        self.generate_button = ttk.Button(
            actions, text="Wordを生成", style="Primary.TButton",
        )
        self.generate_button.pack(side="left")
        self.result_heading = ttk.Label(
            card, textvariable=self.success, style="Success.TLabel",
            justify="left", anchor="w",
        )
        self.result_heading.pack(fill="x", pady=(px(self.root, 12), 0))
        card.bind(
            "<Configure>",
            lambda e: self.result_heading.configure(wraplength=max(100, e.width - px(self.root, 12))),
            add="+",
        )
        file_actions = ttk.Frame(card, style="Card.TFrame")
        file_actions.pack(fill="x", pady=(px(self.root, 6), 0))
        self.open_button = ttk.Button(
            file_actions, text="Wordを開く", style="App.TButton", state="disabled",
        )
        self.open_button.pack(side="left", padx=(0, px(self.root, 8)))
        self.open_folder_button = ttk.Button(
            file_actions, text="保存先を開く", style="App.TButton", state="disabled",
        )
        self.open_folder_button.pack(side="left")

        self.file_button.configure(command=lambda: self.controller.choose_file())
        self.folder_button.configure(command=lambda: self.controller.choose_folder())
        self.preview_button.configure(command=lambda: self.controller.confirm())
        self.generate_button.configure(command=lambda: self.controller.generate())
        self.open_button.configure(command=lambda: self.controller.open_word())
        self.open_folder_button.configure(command=lambda: self.controller.open_folder())
        self.exit_button.configure(command=lambda: self.controller.close())

    def _setting_group(self, caption: str):
        group = ttk.Frame(self.settings, style="Card.TFrame")
        ttk.Label(group, text=caption, style="CardMuted.TLabel").pack(anchor="w", pady=(0, px(self.root, 5)))
        self.setting_groups.append(group)
        return group

    def _card(self, heading: str):
        item = Card(self.inner)
        item.pack(fill="x", pady=(0, px(self.root, 14)))
        ttk.Label(
            item.content, text=heading, style="Section.TLabel",
        ).pack(anchor="w", pady=(0, px(self.root, 10)))
        return item.content

    def _adapt_settings(self, event=None):
        width = self.settings.winfo_width()
        if width <= 1:
            return
        columns = 3 if width >= px(self.root, 720) else 1
        if columns == self._settings_columns:
            return
        self._settings_columns = columns
        for group in self.setting_groups:
            group.grid_forget()
        for column in range(3):
            self.settings.columnconfigure(column, weight=1 if columns == 3 else (1 if column == 0 else 0))
        for index, group in enumerate(self.setting_groups):
            group.grid(
                row=0 if columns == 3 else index,
                column=index if columns == 3 else 0,
                sticky="ew",
                padx=(0, px(self.root, 10)) if columns == 3 and index < 2 else 0,
                pady=(0, px(self.root, 9)) if columns == 1 and index < 2 else 0,
            )

    def _adapt_metrics(self, event=None):
        width = self.metrics.winfo_width()
        if width <= 1:
            return
        columns = 4 if width >= px(self.root, 550) else 2
        if columns == self._metric_columns:
            return
        self._metric_columns = columns
        for tile in self.metric_tiles:
            tile.grid_forget()
        for col in range(4):
            self.metrics.columnconfigure(col, weight=1 if col < columns else 0)
        for index, tile in enumerate(self.metric_tiles):
            tile.grid(
                row=index // columns, column=index % columns, sticky="nsew",
                padx=(0, px(self.root, 8)) if index % columns != columns - 1 else 0,
                pady=(0, px(self.root, 8)) if index < 4 - columns else 0,
            )

    def _canvas_resized(self, event):
        if self.canvas.itemcget(self._canvas_window, "width") != str(event.width):
            self.canvas.itemconfigure(self._canvas_window, width=max(1, event.width))
        self._update_scrollregion()

    def _update_scrollregion(self, event=None):
        if self._scroll_pending is None:
            self._scroll_pending = self.canvas.after_idle(self._finish_scrollregion)

    def _finish_scrollregion(self):
        self._scroll_pending = None
        if self.canvas.winfo_exists():
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _show_mode(self, *args):
        self.file_section.pack_forget()
        self.paste_section.pack_forget()
        if self.input_mode.get() == "file":
            self.file_section.pack(fill="x")
        else:
            self.paste_section.pack(fill="x")
        if hasattr(self, "controller"):
            self.set_batch_mode(self.controller.is_batch())

    def _file_list_wheel(self, event):
        self.file_list.yview_scroll(-int(event.delta / 120) if event.delta else 0, "units")
        return "break"

    def show_selected_files(self, paths: tuple[Path, ...]) -> None:
        self.file_count.set(f"{len(paths)}件選択中" if len(paths) > 1 else "")
        self.file_list.delete(0, "end")
        for path in paths:
            self.file_list.insert("end", path.name)
        if len(paths) > 1:
            self.file_name.pack_forget()
            self.file_list_frame.pack(fill="x", pady=(px(self.root, 4), 0))
        else:
            self.file_list_frame.pack_forget()
            self.file_name.pack(fill="x", pady=(px(self.root, 4), 0))

    def set_batch_mode(self, enabled: bool) -> None:
        if self._batch_mode != enabled:
            if enabled:
                self._single_output_name = self.output_name.get()
                self.output_name.set("元ファイル名を使用")
            else:
                self.output_name.set(self._single_output_name)
            self._batch_mode = enabled
        if not self.controller.busy:
            self.name_entry.configure(state="disabled" if enabled else "normal")
        self.generate_button.configure(
            text=f"Wordを{len(self.controller.selected_files)}件まとめて生成" if enabled else "Wordを生成"
        )

    def _select_all_text(self, event):
        self.input_text.tag_add("sel", "1.0", "end-1c")
        return "break"

    def _on_text_changed(self, event):
        if self.input_text.edit_modified():
            self.input_text.edit_modified(False)
            self.controller.invalidate()

    def _text_wheel(self, event):
        if self.input_mode.get() == "paste":
            self.input_text.yview_scroll(-int(event.delta / 120) if event.delta else 0, "units")
            return "break"
        return None

    def _wheel(self, event):
        try:
            if event.widget.winfo_toplevel() is not self.root:
                return None
        except tk.TclError:
            return None
        if event.widget is self.input_text:
            return "break"
        self.canvas.yview_scroll(-int(event.delta / 120) if event.delta else 0, "units")
        return "break"

    def _help(self):
        dialogs.info(
            "使い方",
            "① Markdownファイルを選ぶか文章を貼り付ける\n"
            "② 文書スタイル、出力ファイル名、保存先を設定する\n"
            "③ 事前確認を実行し、Wordを生成する\n\n"
            "複数選択時は各ファイルを個別のWordへ生成し、同名なら自動連番します。\n"
            "単体生成では置換・別名保存・キャンセルを選べます。",
            self.root,
        )

    def show_preview(self, preview):
        self._set_metric_titles(("文字数", "見出し", "表", "箇条書き"))
        document = preview.document
        self.metric_values["characters"].set(f"{document.characters:,}")
        self.metric_values["headings"].set(str(document.headings))
        self.metric_values["tables"].set(str(document.tables))
        self.metric_values["list_items"].set(str(document.list_items))
        self.info.set(
            f"文字数：{document.characters:,}   見出し：{document.headings}   "
            f"表：{document.tables}   箇条書き：{document.list_items}"
        )
        self.warning.set("\n".join("注意: " + w for w in document.warnings) or "警告はありません")

    def _set_metric_titles(self, names: tuple[str, str, str, str]):
        for key, label in zip(self.metric_values, names):
            self.metric_titles[key].configure(text=label)

    def show_batch_preview(self, items):
        self._set_metric_titles(("選択ファイル", "生成可能", "エラー", "警告"))
        successes = [item for item in items if item.preview is not None]
        failures = [item for item in items if item.error is not None]
        warnings = sum(len(item.preview.document.warnings) for item in successes)
        for key, value in zip(self.metric_values, (len(items), len(successes), len(failures), warnings)):
            self.metric_values[key].set(str(value))
        self.info.set(f"選択：{len(items)}  生成可能：{len(successes)}  エラー：{len(failures)}  警告：{warnings}")
        problems = [f"{item.path.name}：{item.error}" for item in failures]
        self.warning.set(
            "\n".join(problems) if problems else
            (f"{warnings}件の警告があります。" if warnings else "全ファイルの事前確認が完了しました。")
        )

    def clear_preview(self):
        self._set_metric_titles(("文字数", "見出し", "表", "箇条書き"))
        for value in self.metric_values.values():
            value.set("—")
        self.info.set("文字数：—   見出し：—   表：—   箇条書き：—")
        self.warning.set("")
        self.mark_previous_output()

    def set_status(self, message: str):
        self.status.set(message)

    def set_busy(self, busy: bool):
        for widget in (
            self.file_button, self.folder_button, self.preview_button,
            self.generate_button, self.help_button, self.exit_button,
            self.name_entry, self.dir_entry, *self.mode_widgets,
        ):
            widget.configure(state="disabled" if busy else "normal")
        self.style_combo.configure(state="disabled" if busy else "readonly")
        self.input_text.configure(state="disabled" if busy else "normal")
        if not busy:
            self.name_entry.configure(state="disabled" if self._batch_mode else "normal")
        single = self.controller.last_result_kind != "batch" and self.controller.last_output is not None
        batch = (self.controller.last_result_kind == "batch" and
                 self.controller.last_batch_result is not None and
                 bool(self.controller.last_batch_result.saved))
        self.open_button.configure(state="normal" if not busy and single else "disabled")
        self.open_folder_button.configure(state="normal" if not busy and (single or batch) else "disabled")
        if busy:
            self.mark_previous_output()
            self.progress.start(14)
        else:
            self.progress.stop()

    def mark_previous_output(self):
        if self.controller.last_result_kind == "batch" and self.controller.last_batch_result:
            result = self.controller.last_batch_result
            self.success.set(
                f"前回の一括生成結果：成功 {len(result.saved)}件・失敗 {len(result.failures)}件"
                f"\n保存先：{result.folder}"
            )
        elif self.controller.last_output is not None:
            path = self.controller.last_output
            self.success.set(f"前回生成したWordファイル\nファイル名：{path.name}\n保存先：{path.parent}")
        else:
            self.success.set("")

    def show_batch_result(self, result):
        lines = [
            f"一括生成{'停止' if result.stopped_reason else '完了'}：{result.total}件中"
            f"{len(result.saved)}件成功・{len(result.failures)}件失敗"
        ]
        lines.extend(f"成功：{path.name}" for path in result.saved)
        lines.extend(f"失敗：{path.name} — {reason}" for path, reason in result.failures)
        if result.stopped_reason:
            lines.append(f"処理停止：{result.stopped_reason}")
        if result.processed < result.total:
            lines.append(f"未処理：{result.total - result.processed}件")
        if result.saved:
            lines.append(f"保存先：{result.folder}")
        self.success.set("\n".join(lines))
        self.open_button.configure(state="disabled")
        self.open_folder_button.configure(state="normal" if result.saved else "disabled")

    def show_success(self, path: Path):
        self.success.set(
            f"Wordの生成が完了しました\nファイル名：{path.name}\n保存先：{path.parent}"
        )
        self.open_button.configure(state="normal")
        self.open_folder_button.configure(state="normal")
