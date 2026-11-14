"""Sequential batch preview/conversion reusing the single-file service."""
from __future__ import annotations

import tempfile
from dataclasses import dataclass
from pathlib import Path
from collections.abc import Callable

from .errors import ConversionError
from .models import Preview
from .service import convert, prepare, read_file


@dataclass(frozen=True)
class BatchItem:
    path: Path
    source: str | None = None
    preview: Preview | None = None
    error: str | None = None


@dataclass(frozen=True)
class BatchResult:
    saved: tuple[Path, ...]
    failures: tuple[tuple[Path, str], ...]
    processed: int
    total: int
    folder: Path
    stopped_reason: str | None = None


# Calls from this module run in the existing one-worker executor.
Progress = Callable[[int, int, str, int, int], None]


def prepare_batch(paths: tuple[Path, ...], style: str, progress: Progress | None = None) -> tuple[BatchItem, ...]:
    results: list[BatchItem] = []
    for path in paths:
        try:
            source = read_file(path)
            results.append(BatchItem(path, source, prepare(source, style, "file")))
        except (ConversionError, OSError) as exc:
            results.append(BatchItem(path, error=str(exc)))
        if progress:
            progress(len(results), len(paths), path.name,
                     sum(item.preview is not None for item in results),
                     sum(item.error is not None for item in results))
    return tuple(results)


def _preflight(folder: Path, style: str) -> None:
    if not folder.is_dir():
        raise ConversionError("保存先フォルダーが存在しません。")
    if style == "職務経歴書":
        # Check the shared template before committing any output.
        from .resume_template import _load_template
        _load_template()
    try:
        with tempfile.TemporaryFile(dir=folder):
            pass
    except OSError as exc:
        raise ConversionError(f"保存先フォルダーを利用できません: {exc}") from exc


def _available_path(folder: Path, stem: str, reserved: set[str]) -> Path:
    # Windows filenames are case insensitive; retain that protection in tests
    # and when writing to a case-sensitive shared folder.
    existing = {path.name.casefold() for path in folder.iterdir()}
    for index in range(10001):
        suffix = f" ({index})" if index else ""
        candidate = folder / f"{stem}{suffix}.docx"
        if candidate.name.casefold() not in existing | reserved:
            return candidate
    raise ConversionError("空いている出力ファイル名を確保できません。")


def convert_batch(
    items: tuple[BatchItem, ...],
    style: str,
    folder: Path,
    progress: Progress | None = None,
) -> BatchResult:
    folder = Path(folder)
    _preflight(folder, style)
    if any(item.preview is not None and item.preview.style != style for item in items):
        raise ConversionError("文書スタイルが変更されました。事前確認をやり直してください。")

    saved: list[Path] = []
    failures: list[tuple[Path, str]] = []
    reserved: set[str] = set()
    stopped_reason: str | None = None

    for item in items:
        try:
            if item.error:
                raise ConversionError(item.error)
            if item.preview is None:
                raise ConversionError("事前確認が完了していません。")
            # The original preview is authoritative; never generate from
            # changed input, even if the new Markdown still parses correctly.
            current = read_file(item.path)
            for attempt in range(3):
                target = _available_path(folder, item.path.stem, reserved)
                try:
                    output = convert(item.preview, current, style, "file", target, replace=False)
                    saved.append(output)
                    reserved.add(output.name.casefold())
                    break
                except (FileExistsError, ConversionError) as exc:
                    # An external writer may have won the create race.
                    # Only retry when a destination actually appeared.
                    if attempt == 2 or not target.exists():
                        raise
            else:
                raise ConversionError("同名ファイルの競合を解決できませんでした。")
        except (ConversionError, OSError) as exc:
            failures.append((item.path, str(exc)))
            if "テンプレート" in str(exc):
                stopped_reason = str(exc)
        except Exception as exc:
            failures.append((item.path, f"予期しないエラー: {exc}"))
            stopped_reason = "安全に続行できない内部エラーのため停止しました。"

        processed = len(saved) + len(failures)
        if progress:
            progress(processed, len(items), item.path.name, len(saved), len(failures))
        if stopped_reason:
            break

    return BatchResult(tuple(saved), tuple(failures),
                       len(saved) + len(failures), len(items), folder, stopped_reason)
