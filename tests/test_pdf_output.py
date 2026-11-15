"""Collision safety and basic PDF verification: no Microsoft Word required."""
from __future__ import annotations

import pytest

from markdown_to_word.pdf_output import (
    PdfOutputError, finalize_pdf, fingerprint, unique_path, validate_pdf_name, verify_pdf,
)


PDF = b"%PDF-1.7\n1 0 obj\n<<>>\nendobj\n%%EOF\n"


def test_pdf_names_and_casefolded_collisions(tmp_path):
    assert validate_pdf_name("応募書類.PDF") == "応募書類.PDF"
    for name in ("", "bad.docx", "CON.pdf", "name?.pdf", "../x.pdf", "a.pdf."):
        with pytest.raises(PdfOutputError):
            validate_pdf_name(name)
    existing = tmp_path / "申請書.pdf"
    existing.write_bytes(b"existing")
    assert unique_path(existing).name == "申請書 (1).pdf"
    assert unique_path(tmp_path / "Case.pdf", {str(tmp_path / "case.pdf").casefold()}).name == "Case (1).pdf"


def test_new_file_race_does_not_overwrite(tmp_path):
    target = tmp_path / "file.pdf"
    target.write_bytes(b"original")
    temp = tmp_path / ".temp.pdf"
    temp.write_bytes(PDF)
    with pytest.raises(PdfOutputError):
        finalize_pdf(temp, target)
    assert target.read_bytes() == b"original"
    assert temp.read_bytes() == PDF
    renamed = finalize_pdf(temp, target, auto_rename=True)
    assert renamed.name == "file (1).pdf"
    assert renamed.read_bytes() == PDF
    assert target.read_bytes() == b"original"
    assert not temp.exists()


def test_replace_requires_unchanged_existing_file(tmp_path):
    destination = tmp_path / "existing.pdf"
    destination.write_bytes(b"previous")
    expected = fingerprint(destination)
    temp = tmp_path / ".temp.pdf"
    temp.write_bytes(PDF)
    destination.write_bytes(b"another document")
    with pytest.raises(PdfOutputError, match="変更"):
        finalize_pdf(temp, destination, replace=True, expected=expected)
    assert destination.read_bytes() == b"another document"
    assert temp.exists()
    actual = fingerprint(destination)
    assert finalize_pdf(temp, destination, replace=True, expected=actual) == destination
    assert destination.read_bytes() == PDF
    assert not temp.exists()


def test_invalid_export_cannot_destroy_prior_pdf(tmp_path):
    target = tmp_path / "original.pdf"
    target.write_bytes(b"safe")
    temp = tmp_path / ".temp.pdf"
    temp.write_bytes(b"partial")
    with pytest.raises(PdfOutputError):
        finalize_pdf(temp, target, replace=True, expected=fingerprint(target))
    assert target.read_bytes() == b"safe"
    with pytest.raises(PdfOutputError):
        verify_pdf(temp)



def test_copy_failure_removes_only_new_partial_pdf(tmp_path, monkeypatch):
    import errno
    from markdown_to_word import pdf_output

    temp = tmp_path / ".temp.pdf"
    temp.write_bytes(PDF)
    target = tmp_path / "new.pdf"
    original = tmp_path / "existing.pdf"
    original.write_bytes(b"keep")

    def break_copy(source, output):
        output.write(source.read(8))
        raise OSError("copy failed")

    monkeypatch.setattr(pdf_output.os, "link", lambda *_: (_ for _ in ()).throw(OSError(errno.EPERM, "no links")))
    monkeypatch.setattr(pdf_output.shutil, "copyfileobj", break_copy)
    with pytest.raises(OSError, match="copy failed"):
        finalize_pdf(temp, target)
    assert not target.exists()
    assert temp.read_bytes() == PDF
    assert original.read_bytes() == b"keep"


@pytest.mark.parametrize("failure", ["flush", "close"])
def test_flush_or_close_failure_cleans_partial_pdf(tmp_path, monkeypatch, failure):
    import errno
    from pathlib import Path
    from markdown_to_word import pdf_output

    temp, target = tmp_path / ".temp.pdf", tmp_path / "new.pdf"
    temp.write_bytes(PDF)
    real_open = Path.open

    class FailingOutput:
        def __init__(self, stream):
            self.stream = stream

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            self.stream.close()
            if failure == "close":
                raise OSError("close failed")

        def write(self, data):
            return self.stream.write(data)

        def flush(self):
            if failure == "flush":
                raise OSError("flush failed")
            return self.stream.flush()

    def open_with_fault(path, mode="r", *args, **kwargs):
        stream = real_open(path, mode, *args, **kwargs)
        if path == target and mode == "xb":
            return FailingOutput(stream)
        return stream

    monkeypatch.setattr(pdf_output.os, "link", lambda *_: (_ for _ in ()).throw(OSError(errno.EPERM, "no links")))
    monkeypatch.setattr(Path, "open", open_with_fault)
    with pytest.raises(OSError, match=f"{failure} failed"):
        finalize_pdf(temp, target)
    assert not target.exists()
    assert temp.read_bytes() == PDF


def test_failed_cleanup_reports_original_save_error(tmp_path, monkeypatch):
    import errno
    from pathlib import Path
    from markdown_to_word import pdf_output

    temp, target = tmp_path / ".temp.pdf", tmp_path / "partial.pdf"
    temp.write_bytes(PDF)
    real_unlink = Path.unlink

    def break_copy(source, output):
        output.write(source.read(5))
        raise OSError("write failed")

    def unlink_with_fault(path, *args, **kwargs):
        if path == target:
            raise PermissionError("cannot clean up")
        return real_unlink(path, *args, **kwargs)

    monkeypatch.setattr(pdf_output.os, "link", lambda *_: (_ for _ in ()).throw(OSError(errno.EPERM, "no links")))
    monkeypatch.setattr(pdf_output.shutil, "copyfileobj", break_copy)
    monkeypatch.setattr(Path, "unlink", unlink_with_fault)
    with pytest.raises(PdfOutputError, match="未完成PDF") as error:
        finalize_pdf(temp, target)
    assert "write failed" in str(error.value)
    assert "cannot clean up" in str(error.value)
    assert isinstance(error.value.__cause__, OSError)
    assert temp.read_bytes() == PDF
