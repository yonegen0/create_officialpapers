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


def test_batch_four_files_and_same_stem_numbering(tmp_path):
    from markdown_to_word.batch import prepare_batch, convert_batch
    files = [tmp_path / name for name in ("企業A.md", "企業A.txt", "企業B.markdown", "企業C.md")]
    for path in files:
        path.write_text("# 見出し\n\n本文", encoding="utf-8")
    existing = tmp_path / "企業A.docx"
    existing.write_bytes(b"untouched")

    items = prepare_batch(tuple(files), "標準文書")
    assert len(items) == 4
    result = convert_batch(items, "標準文書", tmp_path)
    assert len(result.saved) == 4
    assert not result.failures
    assert [p.name for p in result.saved] == [
        "企業A (1).docx", "企業A (2).docx", "企業B.docx", "企業C.docx"
    ]
    assert existing.read_bytes() == b"untouched"
    assert all(Document(path).paragraphs[0].text == "見出し" for path in result.saved)


def test_batch_continues_after_invalid_and_changed_input(tmp_path):
    from markdown_to_word.batch import prepare_batch, convert_batch
    first, bad, changed, last = (tmp_path / name for name in
                                ("正常1.md", "読込失敗.md", "変更.md", "正常2.md"))
    for path in (first, changed, last):
        path.write_text("# 見出し", encoding="utf-8")
    bad.write_bytes(b"\xff\xfe\xff")
    items = prepare_batch((first, bad, changed, last), "標準文書")
    assert sum(item.error is not None for item in items) == 1
    changed.write_text("# 更新後", encoding="utf-8")
    result = convert_batch(items, "標準文書", tmp_path)
    assert len(result.saved) == 2
    assert len(result.failures) == 2
    assert "変更.md" in [path.name for path, _ in result.failures]
    assert not (tmp_path / "変更.docx").exists()
    assert (tmp_path / "正常1.docx").exists()
    assert (tmp_path / "正常2.docx").exists()


def test_batch_template_missing_stops_without_writing(tmp_path, monkeypatch):
    from markdown_to_word.batch import prepare_batch, convert_batch
    from markdown_to_word import resume_template
    path = tmp_path / "入力.md"
    path.write_text("# 職務経歴書", encoding="utf-8")
    items = prepare_batch((path,), "職務経歴書")
    monkeypatch.setattr(resume_template, "_TEMPLATE_PATH", tmp_path / "not-exist.docx")
    with pytest.raises(ConversionError, match="テンプレート"):
        convert_batch(items, "職務経歴書", tmp_path)
    assert not list(tmp_path.glob("*.docx"))


def test_batch_save_race_retries_numbered_path(tmp_path, monkeypatch):
    import markdown_to_word.batch as batch
    path = tmp_path / "競合.md"
    path.write_text("# 見出し", encoding="utf-8")
    items = batch.prepare_batch((path,), "標準文書")
    original = batch.convert
    calls = []

    def collide(*args, **kwargs):
        target = args[4]
        calls.append(target)
        if len(calls) == 1:
            target.write_bytes(b"other-writer")
        return original(*args, **kwargs)

    monkeypatch.setattr(batch, "convert", collide)
    result = batch.convert_batch(items, "標準文書", tmp_path)
    assert [p.name for p in result.saved] == ["競合 (1).docx"]
    assert (tmp_path / "競合.docx").read_bytes() == b"other-writer"
