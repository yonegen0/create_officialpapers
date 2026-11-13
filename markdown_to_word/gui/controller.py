"""Preview on main thread, Word generation in one worker, UI updates via queue."""
from __future__ import annotations

import os
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from tkinter import filedialog

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
        self.busy = False
        self.worker = ThreadPoolExecutor(max_workers=1, thread_name_prefix="word-render")
        self.dispatcher = Dispatcher(view.root)
        self.last_output: Path | None = None

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
        dialogs.error("処理できませんでした", str(exc), self.view.root)

    def invalidate(self, *args) -> None:
        if self.busy:
            return
        self.preview = None
        self.view.clear_preview()

    def choose_file(self) -> None:
        if self.busy:
            return
        file = filedialog.askopenfilename(parent=self.view.root, filetypes=[("Markdown", "*.md *.markdown *.txt"), ("すべて", "*.*")])
        if not file:
            return
        self.view.file_path.set(file)
        self.view.output_name.set(Path(file).stem + ".docx")
        self.view.output_dir.set(str(Path(file).parent))
        self.invalidate()

    def choose_folder(self) -> None:
        if self.busy:
            return
        folder = filedialog.askdirectory(parent=self.view.root)
        if folder:
            self.view.output_dir.set(folder)

    def confirm(self) -> None:
        if self.busy:
            return
        try:
            self.preview, _ = self._prepare()
            self.view.show_preview(self.preview)
            self.view.set_status("事前確認が完了しました")
        except (ConversionError, OSError) as exc:
            self.preview = None
            self.view.clear_preview()
            self._error(exc)

    def generate(self) -> None:
        if self.busy:
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
            # Prevent accidental nested/relative or separator-containing filenames.
            if Path(name).name != name or name in (".", ".."):
                raise ConversionError("ファイル名にフォルダー区切り文字を使用しないでください。")
            replace = False
            if path.exists():
                choice = dialogs.conflict(self.view.root, path.name)
                if choice == "cancel":
                    self.view.set_status("保存をキャンセルしました")
                    return
                if choice == "rename":
                    path = numbered_path(path)
                    self.view.output_name.set(path.name)
                else:
                    replace = True
            # Ensure changes since preview cannot overwrite an existing file.
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
            self.dispatcher.post(lambda: self._finish(*outcome))
        future.add_done_callback(done)

    def _finish(self, result: Path | None, exc: BaseException | None) -> None:
        self.busy = False
        self.view.set_busy(False)
        if exc:
            self._error(exc)
        else:
            self.last_output = result
            self.view.show_success(result)
            self.view.set_status("Wordの生成が完了しました")

    def open_word(self) -> None:
        if self.last_output and self.last_output.exists():
            if os.name == "nt":
                os.startfile(self.last_output)  # type: ignore[attr-defined]
            else:
                dialogs.info("Windows専用", "Windowsでは関連付けられたアプリで開きます。", self.view.root)

    def open_folder(self) -> None:
        if self.last_output and self.last_output.parent.is_dir():
            if os.name == "nt":
                os.startfile(self.last_output.parent)  # type: ignore[attr-defined]
            else:
                dialogs.info("Windows専用", str(self.last_output.parent), self.view.root)

    def close(self) -> None:
        if self.busy:
            dialogs.info("処理中", "生成処理が終了してから閉じてください。", self.view.root)
            return
        self.dispatcher.close()
        self.worker.shutdown(wait=False, cancel_futures=True)
        self.view.root.destroy()
