from pathlib import Path
import os
import pytest
from docx import Document

from markdown_to_word.errors import ConversionError, InputChangedError
from markdown_to_word.service import prepare, convert, read_file
from markdown_to_word.output import numbered_path


def test_file_encoding_and_changed_input(tmp_path):
    p = tmp_path / "入力.md"
    p.write_bytes(b"\xef\xbb\xbf# \xe8\xa6\x8b\xe5\x87\xba\xe3\x81\x97\r\n")
    source = read_file(p)
    assert "見出し" in source
    plan = prepare(source, "標準文書", "file")
    with pytest.raises(InputChangedError):
        convert(plan, source + "追加", "標準文書", "file", tmp_path / "word.docx", replace=False)
    assert not (tmp_path / "word.docx").exists()


def test_success_no_overwrite_and_rename(tmp_path):
    p = prepare("# 見出し\n\n文章", "標準文書", "paste")
    out = tmp_path / "架空.docx"
    convert(p, "# 見出し\n\n文章", "標準文書", "paste", out, replace=False)
    assert Document(out).paragraphs[1].text == "文章"
    before = out.read_bytes()
    with pytest.raises(FileExistsError):
        convert(p, "# 見出し\n\n文章", "標準文書", "paste", out, replace=False)
    assert out.read_bytes() == before
    next_path = numbered_path(out)
    convert(p, "# 見出し\n\n文章", "標準文書", "paste", next_path, replace=False)
    assert next_path.name == "架空 (1).docx"


def test_replacement_failure_protects_original(tmp_path, monkeypatch):
    import markdown_to_word.output as output
    p = prepare("変更後", "標準文書", "paste")
    path = tmp_path / "保存.docx"
    path.write_bytes(b"old-content")
    old = path.read_bytes()

    def fail_replace(*args, **kwargs):
        raise PermissionError("opened")

    monkeypatch.setattr(output.os, "replace", fail_replace)
    with pytest.raises(ConversionError):
        convert(p, "変更後", "標準文書", "paste", path, replace=True)
    assert path.read_bytes() == old
    assert not list(tmp_path.glob(".mtw-*"))


def test_replace_success(tmp_path):
    p = prepare("更新済み", "標準文書", "paste")
    path = tmp_path / "word.docx"
    path.write_bytes(b"old")
    convert(p, "更新済み", "標準文書", "paste", path, replace=True)
    assert Document(path).paragraphs[0].text == "更新済み"


def test_empty_input_and_invalid_encoding(tmp_path):
    with pytest.raises(ConversionError):
        prepare(" \n ", "職務経歴書", "paste")
    p = tmp_path / "input.md"
    p.write_bytes(b"\xff\xfe\xff")
    with pytest.raises(ConversionError):
        read_file(p)
