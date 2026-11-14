"""GUI orchestration for single-file and sequential batch generation."""
from __future__ import annotations

import os
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from tkinter import filedialog

from ..batch import BatchItem, BatchResult, convert_batch, prepare_batch
from ..errors import ConversionError
from ..models import Preview
from ..output import numbered_path
from ..service import prepare, convert, read_file
from .dispatcher import Dispatcher
from . import dialogs


class Controller:
    def __init__(self, view) -> None:
        self.view = view
        self.preview: Preview | None = None
        self.selected_files: tuple[Path, ...] = ()
        self.batch_previews: tuple[BatchItem, ...] | None = None
        self.busy = False
        self.worker = ThreadPoolExecutor(max_workers=1, thread_name_prefix="word-render")
        self.dispatcher = Dispatcher(view.root)
        self.last_output: Path | None = None
        self.last_batch_result: BatchResult | None = None
        self.last_result_kind: str | None = None

    def is_batch(self) -> bool:
        return self.view.input_mode.get() == "file" and len(self.selected_files) > 1

    def source(self) -> str:
        if self.view.input_mode.get() == "file":
            name = self.view.file_path.get().strip()
            if not name:
                raise ConversionError("Markdownファイルを選択してください。")
            return read_file(Path(name))
        return self.view.input_text.get("1.0", "end-1c")

    def _prepare(self) -> tuple[Preview, str]:
        content = self.source()
        result = prepare(content, self.view.style.get(), self.view.input_mode.get())
        return result, content

    def _error(self, exc: BaseException) -> None:
        self.view.set_status("エラーが発生しました")
        self.view.mark_previous_output()
        dialogs.error("処理できませんでした", str(exc), self.view.root)

    def invalidate(self, *args) -> None:
        if self.busy:
            return
        self.preview = None
        self.batch_previews = None
        self.view.clear_preview()

    def choose_file(self) -> None:
        if self.busy:
            return
        files = filedialog.askopenfilenames(
            parent=self.view.root,
            filetypes=[("Markdown", "*.md *.markdown *.txt"), ("すべて", "*.*")],
        )
        if not files:
            return
        selected: list[Path] = []
        seen: set[str] = set()
        for name in files:
            path = Path(name)
            canonical = os.path.normcase(os.path.realpath(path)).casefold()
            if canonical not in seen:
                selected.append(path)
                seen.add(canonical)
        if not selected:
            return
        self.selected_files = tuple(selected)
        self.view.file_path.set(str(selected[0]))
        self.view.show_selected_files(self.selected_files)
        self.view.set_batch_mode(self.is_batch())
        if len(selected) == 1:
            self.view.output_name.set(selected[0].stem + ".docx")
        self.view.output_dir.set(str(selected[0].parent))
        self.invalidate()

    def choose_folder(self) -> None:
        if self.busy:
            return
        folder = filedialog.askdirectory(parent=self.view.root)
        if folder:
            self.view.output_dir.set(folder)

    def _progress(self, action: str):
        def notify(done: int, total: int, name: str, ok: int, failed: int) -> None:
            message = f"{action} {done}/{total}件｜現在：{name}｜成功 {ok}・失敗 {failed}"
            self.dispatcher.post(lambda message=message: self.view.set_status(message))
        return notify

    def confirm(self) -> None:
        if self.busy:
            return
        if self.is_batch():
            self.preview = None
            self.batch_previews = None
            paths = self.selected_files
            style = self.view.style.get()
            self.busy = True
            self.view.clear_preview()
            self.view.set_busy(True)
            self.view.set_status(f"{len(paths)}件のMarkdownを確認しています…")
            future = self.worker.submit(prepare_batch, paths, style, self._progress("事前確認"))

            def done(f):
                try:
                    result = (f.result(), None)
                except Exception as exc:
                    result = (None, exc)
                self.dispatcher.post(lambda result=result: self._finish_batch_preview(*result))

            future.add_done_callback(done)
            return
        try:
            self.batch_previews = None
            self.preview, _ = self._prepare()
            self.view.show_preview(self.preview)
            self.view.set_status("事前確認が完了しました")
        except (ConversionError, OSError) as exc:
            self.preview = None
            self.view.clear_preview()
            self._error(exc)

    def _finish_batch_preview(self, items: tuple[BatchItem, ...] | None, exc: BaseException | None) -> None:
        self.busy = False
        self.view.set_busy(False)
        if exc is not None:
            self._error(exc)
            return
        assert items is not None
        self.batch_previews = items
        self.view.show_batch_preview(items)
        valid = sum(item.preview is not None for item in items)
        self.view.set_status(f"一括事前確認完了：{valid}/{len(items)}件生成可能")

    def generate(self) -> None:
        if self.busy:
            return
        self.view.mark_previous_output()
        if self.is_batch():
            self._generate_batch()
            return
        if self.preview is None:
            dialogs.info("事前確認", "先に「事前確認」を実行してください。", self.view.root)
            return
        try:
            content = self.source()
            if not self.view.output_name.get().strip():
                raise ConversionError("出力ファイル名を入力してください。")
            name = self.view.output_name.get().strip()
            path = Path(self.view.output_dir.get().strip()) / name
            if path.suffix.lower() != ".docx":
                raise ConversionError("出力ファイル名は .docx にしてください。")
            if not path.parent.is_dir():
                raise ConversionError("保存先フォルダーがありません。")
            if Path(name).name != name or name in (".", ".."):
                raise ConversionError("ファイル名にフォルダー区切り文字を使用しないでください。")
            replace = False
            if path.exists():
                choice = dialogs.conflict(self.view.root, path)
                if choice == "cancel":
                    self.view.set_status("保存をキャンセルしました")
                    self.view.mark_previous_output()
                    return
                if choice == "rename":
                    path = numbered_path(path)
                    self.view.output_name.set(path.name)
                else:
                    replace = True
            from ..service import _digest
            if self.preview.fingerprint != _digest(content, self.view.style.get(), self.view.input_mode.get()):
                raise ConversionError("入力が変わりました。事前確認をやり直してください。")
            preview = self.preview
        except (OSError, ConversionError) as exc:
            self._error(exc)
            return
        self.busy = True
        self.view.set_busy(True)
        self.view.set_status("Wordを生成しています…")
        future = self.worker.submit(convert, preview, content, self.view.style.get(), self.view.input_mode.get(), path, replace=replace)

        def done(f):
            try:
                output = f.result()
                outcome = (output, None)
            except Exception as exc:
                outcome = (None, exc)
            self.dispatcher.post(lambda outcome=outcome: self._finish(*outcome))
        future.add_done_callback(done)

    def _generate_batch(self) -> None:
        items = self.batch_previews
        if items is None:
            dialogs.info("事前確認", "先に「事前確認」を実行してください。", self.view.root)
            return
        if tuple(item.path for item in items) != self.selected_files or any(
            item.preview is not None and item.preview.style != self.view.style.get() for item in items
        ):
            self.invalidate()
            dialogs.info("事前確認", "入力またはスタイルが変わりました。事前確認をやり直してください。", self.view.root)
            return
        folder_text = self.view.output_dir.get().strip()
        folder = Path(folder_text) if folder_text else Path("__missing_output_directory__")
        if not folder.is_dir():
            self._error(ConversionError("保存先フォルダーが存在しません。"))
            return
        style = self.view.style.get()
        self.busy = True
        self.view.set_busy(True)
        self.view.set_status(f"{len(items)}件のWordを生成しています…")
        future = self.worker.submit(convert_batch, items, style, folder, self._progress("生成"))

        def done(f):
            try:
                outcome = (f.result(), None)
            except Exception as exc:
                outcome = (None, exc)
            self.dispatcher.post(lambda outcome=outcome: self._finish_batch(*outcome))

        future.add_done_callback(done)

    def _finish_batch(self, result: BatchResult | None, exc: BaseException | None) -> None:
        self.busy = False
        self.view.set_busy(False)
        if exc is not None:
            self._error(exc)
            return
        assert result is not None
        self.last_batch_result = result
        self.last_result_kind = "batch"
        self.view.show_batch_result(result)
        self.view.set_status(
            f"一括生成{ '停止' if result.stopped_reason else '完了' }："
            f"成功 {len(result.saved)}件・失敗 {len(result.failures)}件"
        )

    def _finish(self, result: Path | None, exc: BaseException | None) -> None:
        self.busy = False
        self.view.set_busy(False)
        if exc:
            self._error(exc)
        else:
            self.last_output = result
            self.last_result_kind = "single"
            self.view.show_success(result)
            self.view.set_status("Wordの生成が完了しました")

    def open_word(self) -> None:
        if self.last_result_kind != "single":
            return
        if self.last_output and self.last_output.exists():
            if os.name == "nt":
                try:
                    os.startfile(self.last_output)  # type: ignore[attr-defined]
                except OSError as exc:
                    dialogs.error("Wordを開けませんでした", str(exc), self.view.root)
            else:
                dialogs.info("Windows専用", "Windowsでは関連付けられたアプリで開きます。", self.view.root)

    def open_folder(self) -> None:
        if self.last_result_kind == "batch":
            folder = self.last_batch_result.folder if self.last_batch_result and self.last_batch_result.saved else None
        else:
            folder = self.last_output.parent if self.last_output else None
        if folder and folder.is_dir():
            if os.name == "nt":
                try:
                    os.startfile(folder)  # type: ignore[attr-defined]
                except OSError as exc:
                    dialogs.error("保存先を開けませんでした", str(exc), self.view.root)
            else:
                dialogs.info("Windows専用", str(folder), self.view.root)

    def close(self) -> None:
        if self.busy:
            dialogs.info("処理中", "処理が終了してから閉じてください。", self.view.root)
            return
        self.dispatcher.close()
        self.worker.shutdown(wait=False, cancel_futures=True)
        self.view.root.destroy()
