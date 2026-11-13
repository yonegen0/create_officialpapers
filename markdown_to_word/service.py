"""Shared conversion path for file input and pasted text."""
from __future__ import annotations

import hashlib
from pathlib import Path

from .errors import ConversionError, InputChangedError
from .models import Preview
from .parser import parse_markdown
from .output import save_document


def read_file(path: Path) -> str:
    if path.suffix.lower() not in (".md", ".markdown", ".txt"):
        raise ConversionError(".md / .markdown / .txt ファイルを選択してください。")
    try:
        return path.read_text(encoding="utf-8-sig")
    except FileNotFoundError as exc:
        raise ConversionError("選択したファイルが見つかりません。") from exc
    except UnicodeError as exc:
        raise ConversionError("UTF-8形式で読み取れませんでした。") from exc
    except OSError as exc:
        raise ConversionError(f"ファイルを読み取れませんでした: {exc}") from exc


def _digest(source: str, style: str, mode: str) -> str:
    return hashlib.sha256((style + "\x00" + mode + "\x00" + source).encode("utf-8")).hexdigest()


def prepare(source: str, style: str, mode: str) -> Preview:
    if mode not in ("file", "paste"):
        raise ConversionError("入力方式が正しくありません。")
    parsed = parse_markdown(source, style)
    return Preview(parsed, _digest(source, style, mode), style, mode)


def convert(preview: Preview, source: str, style: str, mode: str, target: Path, *, replace: bool) -> Path:
    if preview.fingerprint != _digest(source, style, mode):
        raise InputChangedError("事前確認後にMarkdownまたは設定が変更されました。事前確認をやり直してください。")
    return save_document(preview.document, target, replace=replace)
