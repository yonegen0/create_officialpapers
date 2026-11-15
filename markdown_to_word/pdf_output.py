"""Safe, local-only PDF output naming and finalization."""
from __future__ import annotations

import errno
import os
import shutil
import uuid
from pathlib import Path


class PdfOutputError(ValueError):
    """Invalid output path or a failed safe save."""


def validate_pdf_name(name: str) -> str:
    name = name.strip()
    invalid = '<>:"/\\|?*'
    if not name or name in (".", "..") or any(c in name for c in invalid) or "\x00" in name:
        raise PdfOutputError("PDFファイル名に使用できない文字が含まれています。")
    if name.endswith((".", " ")) or Path(name).suffix.lower() != ".pdf":
        raise PdfOutputError("出力ファイル名は .pdf で終わる名前にしてください。")
    if name.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL", *("COM" + str(i) for i in range(1, 10)), *("LPT" + str(i) for i in range(1, 10))}:
        raise PdfOutputError("Windowsで予約されているファイル名は使用できません。")
    return name


def validate_folder(folder: Path) -> None:
    if not folder.is_dir():
        raise PdfOutputError("PDFの保存先フォルダーが存在しません。")


def fingerprint(path: Path) -> tuple[int, int, int] | None:
    try:
        s = path.stat()
    except FileNotFoundError:
        return None
    return (s.st_mtime_ns, s.st_size, s.st_ino)


def unique_path(candidate: Path, reserved: set[str] | None = None) -> Path:
    """Choose an available name; finalization checks again to close races."""
    reserved = reserved if reserved is not None else set()
    for n in range(10000):
        path = candidate if n == 0 else candidate.with_name(f"{candidate.stem} ({n}){candidate.suffix}")
        if str(path).casefold() not in reserved and not path.exists():
            return path
    raise PdfOutputError("保存できるファイル名を見つけられませんでした。")


def temporary_path(folder: Path) -> Path:
    """Unique name in the destination directory, not a held-open file handle."""
    for _ in range(10):
        path = folder / f".mtw-pdf-{uuid.uuid4().hex}.pdf"
        if not path.exists():
            return path
    raise PdfOutputError("一時PDFの保存先を確保できませんでした。")


def verify_pdf(path: Path) -> None:
    try:
        with path.open("rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise PdfOutputError("Wordの出力がPDF形式ではありません。")
            if path.stat().st_size == 0:
                raise PdfOutputError("出力されたPDFが空です。")
            stream.seek(max(0, path.stat().st_size - 1024))
            if b"%%EOF" not in stream.read():
                raise PdfOutputError("WordのPDF出力が完了していません。")
    except FileNotFoundError as exc:
        raise PdfOutputError("WordからPDFが出力されませんでした。") from exc


def _publish_new(temp: Path, target: Path) -> None:
    """Create destination exclusively; never overwrite an existing PDF."""
    try:
        os.link(temp, target)
    except FileExistsError:
        raise
    except OSError as exc:
        # Some Windows/network filesystems disallow hard links. Preserve O_EXCL.
        if exc.errno not in (errno.EPERM, errno.EACCES, errno.ENOTSUP, errno.EOPNOTSUPP, errno.EXDEV, errno.EINVAL):
            raise
        # The target is ours only after the exclusive open succeeded.
        # Flush and close errors must receive the same cleanup as copy errors.
        created = False
        try:
            with temp.open("rb") as source, target.open("xb") as output:
                created = True
                shutil.copyfileobj(source, output)
                output.flush()
        except BaseException as error:
            if created:
                try:
                    target.unlink(missing_ok=True)
                except OSError as cleanup_error:
                    raise PdfOutputError(
                        f"PDF保存に失敗しました（{error}）。"
                        f"未完成PDF「{target.name}」も削除できませんでした（{cleanup_error}）。"
                    ) from error
            raise


def finalize_pdf(
    temp: Path,
    destination: Path,
    *,
    replace: bool = False,
    expected: tuple[int, int, int] | None = None,
    auto_rename: bool = False,
    reserved: set[str] | None = None,
    cleanup_temp: bool = True,
) -> Path:
    verify_pdf(temp)
    reserved = reserved if reserved is not None else set()
    if replace:
        if expected is None or fingerprint(destination) != expected:
            raise PdfOutputError("既存PDFが確認後に変更されました。置換を中止しました。")
        os.replace(temp, destination)
        reserved.add(str(destination).casefold())
        return destination
    candidate = destination
    for _ in range(10000):
        candidate = unique_path(destination, reserved) if auto_rename else candidate
        try:
            _publish_new(temp, candidate)
        except FileExistsError:
            if not auto_rename:
                raise PdfOutputError("保存先に同名PDFが作成されました。既存ファイルは保護しました。") from None
            reserved.add(str(candidate).casefold())
            continue
        reserved.add(str(candidate).casefold())
        # Batch conversion owns cleanup so a locked temporary PDF cannot turn
        # an already-published result into a per-file failure.
        if cleanup_temp:
            temp.unlink(missing_ok=True)
        return candidate
    raise PdfOutputError("保存可能なPDFファイル名を見つけられませんでした。")
