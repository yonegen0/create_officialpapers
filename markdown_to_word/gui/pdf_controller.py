"""Word→PDF UI actions; the owner Controller supplies the shared worker/busy."""
from __future__ import annotations

import os
from pathlib import Path
from tkinter import filedialog

from ..pdf_converter import PdfBatchResult, convert_pdf_files
from ..pdf_output import fingerprint, validate_folder, validate_pdf_name
from . import dialogs


class PdfController:
    def __init__(self, owner) -> None:
        self.owner = owner
        self.view = owner.view.pdf_tab
        self.view.controller = self
        self.selected_files: tuple[Path, ...] = ()
        self.last_result: PdfBatchResult | None = None

    def choose_files(self) -> None:
        if self.owner.busy:
            return
        names = filedialog.askopenfilenames(
            parent=self.owner.view.root, filetypes=[("Word文書", "*.docx")],
        )
        if not names:
            return
        selected = []
        seen = set()
        for name in names:
            path = Path(name)
            key = os.path.normcase(os.path.realpath(path)).casefold()
            if key not in seen:
                seen.add(key)
                selected.append(path)
        if selected:
            self.selected_files = tuple(selected)
            self.view.set_selection(self.selected_files)
            self.owner.view.set_status(f"Wordファイルを{len(selected)}件選択しました")

    def choose_folder(self) -> None:
        if self.owner.busy:
            return
        folder = filedialog.askdirectory(parent=self.owner.view.root)
        if folder:
            self.view.output_dir.set(folder)

    def convert(self) -> None:
        if self.owner.busy:
            return
        root = self.owner.view.root
        try:
            if not self.selected_files:
                raise ValueError("Wordファイルを選択してください。")
            folder_text = self.view.output_dir.get().strip()
            if not folder_text:
                raise ValueError("PDFの保存先フォルダーを選択してください。")
            folder = Path(folder_text)
            validate_folder(folder)
            name = None
            replace = False
            expected = None
            if len(self.selected_files) == 1:
                name = validate_pdf_name(self.view.output_name.get())
                destination = folder / name
                if destination.exists():
                    choice = dialogs.conflict(root, destination, kind="PDF")
                    if choice == "cancel":
                        self.owner.view.set_status("PDFの保存をキャンセルしました")
                        return
                    if choice == "replace":
                        replace = True
                        expected = fingerprint(destination)
                    else:
                        # User chose a numbered name. The worker rechecks races.
                        from ..pdf_output import unique_path
                        name = unique_path(destination).name
                        self.view.output_name.set(name)
            files = self.selected_files
            self.view.mark_previous()
            self.owner.submit_job(
                lambda: convert_pdf_files(
                    files, folder, output_name=name, replace=replace, expected=expected,
                    progress=self.owner._progress("PDF変換"),
                ),
                self._finish,
                f"{len(files)}件のWordをPDFに変換しています…",
            )
        except (ValueError, OSError) as exc:
            dialogs.error("PDF変換を開始できません", str(exc), root)

    def _finish(self, result: PdfBatchResult | None, error: Exception | None) -> None:
        if error is not None:
            dialogs.error("PDF変換できませんでした", str(error), self.owner.view.root)
            self.owner.view.set_status("PDF変換に失敗しました")
            return
        assert result is not None
        self.last_result = result
        self.view.show_result(result)
        self.owner.view.set_status(
            f"PDF変換{'停止' if result.stopped_reason else '完了'}：成功 {len(result.saved)}件・失敗 {len(result.failures)}件"
        )

    def open_pdf(self) -> None:
        if not self.last_result or len(self.last_result.saved) != 1 or self.last_result.total != 1:
            return
        self._open(self.last_result.saved[0])

    def open_folder(self) -> None:
        if self.last_result and self.last_result.saved:
            self._open(self.last_result.folder)

    def _open(self, path: Path) -> None:
        if os.name == "nt":
            try:
                os.startfile(path)  # type: ignore[attr-defined]
            except OSError as exc:
                dialogs.error("ファイルを開けませんでした", str(exc), self.owner.view.root)
        else:
            dialogs.info("Windows専用", str(path), self.owner.view.root)
