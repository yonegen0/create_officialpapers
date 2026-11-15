"""Microsoft Word COM PDF export; COM objects never leave the worker thread."""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .pdf_output import (
    PdfOutputError, finalize_pdf, fingerprint, temporary_path, validate_folder,
    validate_pdf_name,
)


class WordConnectionError(RuntimeError):
    """The dedicated Word automation instance is no longer reliable."""


@dataclass(frozen=True)
class PdfBatchResult:
    folder: Path
    total: int
    processed: int
    saved: tuple[Path, ...]
    failures: tuple[tuple[Path, str], ...]
    stopped_reason: str | None = None
    warnings: tuple[tuple[Path, str], ...] = ()


_FATAL_HRESULTS = {0x80010108, 0x800706BA, 0x80010001, 0x80010105, 0x800401FD}


def _connection_failed(error: Exception) -> bool:
    if isinstance(error, WordConnectionError):
        return True
    hresult = getattr(error, "hresult", None)
    if isinstance(hresult, int) and (hresult & 0xffffffff) in _FATAL_HRESULTS:
        return True
    return False


def _reason(error: Exception) -> str:
    if isinstance(error, (PdfOutputError, WordConnectionError)):
        return str(error)
    if isinstance(error, (PermissionError, FileExistsError)):
        return "保存先にアクセスできないか、PDFが使用中です。"
    if isinstance(error, (FileNotFoundError, IsADirectoryError)):
        return "入力Wordファイルまたは保存先が見つかりません。"
    if getattr(error, "hresult", None) is not None:
        return "Wordで文書を開くかPDFを書き出せませんでした。破損・パスワード・使用中の可能性があります。"
    return str(error) or "PDF変換に失敗しました。"


class WordSession:
    """Own one DispatchEx instance on the current COM-initialized thread."""

    def __init__(self) -> None:
        self._pythoncom = None
        self.app = None

    def __enter__(self):
        if os.name != "nt":
            raise WordConnectionError("Word→PDF変換はWindows 11のデスクトップ版Microsoft Wordが必要です。")
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise WordConnectionError("PDF変換に必要なpywin32が利用できません。") from exc
        try:
            pythoncom.CoInitialize()
        except Exception as exc:
            raise WordConnectionError("WordのCOM接続を初期化できませんでした。") from exc
        self._pythoncom = pythoncom
        try:
            self.app = win32com.client.DispatchEx("Word.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = 0
            # Disable execution of macros in documents opened by automation.
            self.app.AutomationSecurity = 3
        except Exception as exc:
            self.__exit__(None, None, None)
            raise WordConnectionError("Microsoft Wordを起動できません。") from exc
        return self

    def export(self, source: Path, temp: Path) -> None:
        if self.app is None:
            raise WordConnectionError("Wordとの接続がありません。")
        document = None
        try:
            document = self.app.Documents.Open(
                FileName=str(source), ConfirmConversions=False, ReadOnly=True,
                AddToRecentFiles=False, UpdateLinks=0, Revert=False,
                PasswordDocument="", WritePasswordDocument="", Visible=False,
                NoEncodingDialog=True, OpenAndRepair=False,
            )
            document.ExportAsFixedFormat(
                OutputFileName=str(temp), ExportFormat=17, OpenAfterExport=False,
                OptimizeFor=0, Range=0, Item=0, IncludeDocProps=True,
                KeepIRM=True, CreateBookmarks=0, DocStructureTags=True,
                BitmapMissingFonts=True, UseISO19005_1=False,
            )
        finally:
            if document is not None:
                try:
                    document.Close(SaveChanges=0)
                except Exception as exc:
                    raise WordConnectionError("Word文書を安全に閉じられませんでした。一括変換を停止します。") from exc

    def is_alive(self) -> bool:
        try:
            return self.app is not None and bool(self.app.Version)
        except Exception:
            return False

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.app is not None:
                try:
                    self.app.Quit(SaveChanges=0)
                except Exception:
                    pass  # Do not terminate user-owned WINWORD.EXE processes.
        finally:
            self.app = None
            if self._pythoncom is not None:
                self._pythoncom.CoUninitialize()
                self._pythoncom = None


def convert_pdf_files(
    files: tuple[Path, ...],
    folder: Path,
    *,
    output_name: str | None = None,
    replace: bool = False,
    expected: tuple[int, int, int] | None = None,
    progress: Callable[[int, int, str, int, int], None] | None = None,
    session_factory=WordSession,
) -> PdfBatchResult:
    """Sequential batch in exactly one Word instance; each file is independent."""
    if not files:
        raise PdfOutputError("Wordファイルを選択してください。")
    validate_folder(folder)
    if len(files) == 1:
        validate_pdf_name(output_name or files[0].stem + ".pdf")
    saved: list[Path] = []
    failures: list[tuple[Path, str]] = []
    warnings: list[tuple[Path, str]] = []
    reserved: set[str] = set()
    processed = 0
    stopped_reason = None
    try:
        with session_factory() as session:
            for source in files:
                temp = None
                try:
                    if source.suffix.lower() != ".docx":
                        raise PdfOutputError("変換対象は .docx のみです。")
                    if not source.is_file():
                        raise FileNotFoundError(str(source))
                    name = validate_pdf_name(output_name if len(files) == 1 and output_name else source.stem + ".pdf")
                    destination = folder / name
                    temp = temporary_path(folder)
                    session.export(source, temp)
                    result = finalize_pdf(
                        temp, destination, replace=replace and len(files) == 1,
                        expected=expected, auto_rename=len(files) > 1 or not replace and destination.exists(),
                        reserved=reserved, cleanup_temp=False,
                    )
                    saved.append(result)
                except Exception as error:
                    failures.append((source, _reason(error)))
                    if _connection_failed(error) or not session.is_alive():
                        stopped_reason = "Wordとの接続が不安定なため、残りの変換を停止しました。"
                finally:
                    if temp is not None:
                        try:
                            temp.unlink(missing_ok=True)
                        except OSError as cleanup_error:
                            warnings.append((
                                source,
                                f"一時PDF「{temp.name}」を削除できませんでした（{cleanup_error}）。"
                                "保存済みPDFの結果には影響しません。",
                            ))
                processed += 1
                if progress is not None:
                    progress(processed, len(files), source.name, len(saved), len(failures))
                if stopped_reason:
                    break
    except WordConnectionError as error:
        stopped_reason = str(error)
    return PdfBatchResult(
        folder=folder, total=len(files), processed=processed,
        saved=tuple(saved), failures=tuple(failures), stopped_reason=stopped_reason,
        warnings=tuple(warnings),
    )
