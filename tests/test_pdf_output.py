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
