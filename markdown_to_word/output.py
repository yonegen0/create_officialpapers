"""Atomic same-directory docx commit; never delete or truncate an existing file."""
from __future__ import annotations

import os
import tempfile
from pathlib import Path

from docx import Document

from .errors import ConversionError
from .models import ParsedDocument
from .docx_renderer import render_docx


def numbered_path(target: Path) -> Path:
    i = 1
    while True:
        candidate = target.with_name(f"{target.stem} ({i}){target.suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def validate_target(target: Path) -> Path:
    if target.suffix.lower() != ".docx" or not target.stem.strip():
        raise ConversionError("ファイル名は .docx で終わる名前にしてください。")
    if not target.parent.is_dir():
        raise ConversionError("保存先フォルダーが存在しません。")
    return target


def _signature(path: Path) -> tuple[int, int, int] | None:
    try:
        st = path.stat()
        return st.st_ino, st.st_size, st.st_mtime_ns
    except FileNotFoundError:
        return None


def save_document(parsed: ParsedDocument, target: Path, *, replace: bool = False) -> Path:
    """Render and verify a temp DOCX before atomic finalization.

    On ordinary NTFS, hardlink creation is atomic and fails if the destination
    exists. Overwrite uses os.replace, which preserves the prior destination
    if replacement fails (e.g., Word has the file open).
    """
    target = validate_target(Path(target))
    before = _signature(target)
    if before is not None and not replace:
        raise FileExistsError(f"同名のWordがあります: {target.name}")
    temp: Path | None = None
    try:
        fd, name = tempfile.mkstemp(prefix=".mtw-", suffix=".docx", dir=target.parent)
        temp = Path(name)
        os.close(fd)
        doc = render_docx(parsed)
        doc.save(temp)
        # Verify the DOCX is readable before the only destructive operation.
        reopened = Document(temp)
        if reopened.element.body is None:
            raise ConversionError("Wordファイルの検証に失敗しました。")
        if _signature(target) != before:
            raise ConversionError("保存中に同名ファイルが変更されました。再実行してください。")
        if replace:
            os.replace(temp, target)
        else:
            try:
                os.link(temp, target)
            except FileExistsError:
                raise
            except OSError:
                # Unsupported hardlinks (some network/share volumes): O_EXCL
                # guarantees we don't truncate an existing document.
                with temp.open("rb") as src:
                    with target.open("xb") as dst:
                        try:
                            while part := src.read(1024 * 1024):
                                dst.write(part)
                            dst.flush()
                            os.fsync(dst.fileno())
                        except BaseException:
                            dst.close()
                            target.unlink(missing_ok=True)
                            raise
        return target
    except (PermissionError, OSError) as exc:
        raise ConversionError(f"Wordを保存できませんでした。保存先やWordの使用状況を確認してください: {exc}") from exc
    finally:
        if temp is not None:
            temp.unlink(missing_ok=True)
