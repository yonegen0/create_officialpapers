import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from markdown_to_word.docx_renderer import render_docx
from markdown_to_word.errors import ConversionError
from markdown_to_word.parser import parse_markdown
import markdown_to_word.resume_template as resume


@pytest.fixture
def sample_template(tmp_path, monkeypatch):
    """Synthetic contract fixture only; never shipped as the visual template."""
    doc = Document()
    doc.sections[0].page_width = Mm(210)
    doc.sections[0].page_height = Mm(297)
    normal = doc.styles["Normal"]
    normal.font.name = "Meiryo UI"
    rfonts = normal.element.get_or_add_rPr().find(qn("w:rFonts"))
    rfonts.set(qn("w:eastAsia"), "Meiryo UI")
    for key in ("TITLE", "DATE", "NAME", "HEADING1", "HEADING2", "BODY", "END"):
        doc.add_paragraph("{{MTW_" + key + "}}")
    for key in ("COMPANY", "ROLE", "SKILL"):
        table = doc.add_table(rows=2 if key != "COMPANY" else 1, cols=2)
        table.cell(0, 0).text = "{{MTW_" + key + "}}"
        table.cell(0, 1).text = "書式サンプル（出力不可）"
        if key != "COMPANY":
            table.cell(1, 0).text = "書式サンプル（出力不可）"
            table.cell(1, 1).text = "書式サンプル（出力不可）"
            if key == "ROLE":
                table.cell(1, 1).add_paragraph("{{MTW_ROLE_1}}")
                table.cell(1, 1).add_paragraph("{{MTW_ROLE_2}}")
    path = tmp_path / "resume_template.docx"
    doc.save(path)
    monkeypatch.setattr(resume, "_TEMPLATE_PATH", path)
    return path


def test_resume_company_and_tasks(sample_template):
    source = (
        "職務経歴書\n\n2026年11月13日現在\n\n氏名：架空 花子\n\n"
        "■ 職務経歴\n\n"
        "勤務先\t架空製作所\n在籍期間\t2023年～現在\n\n"
        "### 受付・一般事務\n\n連絡調整を担当しました。\n\n"
        "【主な業務】\n\n- 来客対応\n- 日程調整\n"
    )
    doc = render_docx(parse_markdown(source))
    assert abs(doc.sections[0].page_width - Mm(210)) < 10000
    assert len(doc.tables) == 2
    assert doc.tables[0].cell(0, 1).text == "架空製作所"
    assert doc.tables[1].cell(0, 0).text == "職種・期間"
    assert doc.tables[1].cell(1, 0).text == "受付・一般事務"
    details = [p.text for p in doc.tables[1].cell(1, 1).paragraphs]
    assert details == ["連絡調整を担当しました。", "【主な業務】", "・来客対応", "・日程調整"]
    assert "書式サンプル" not in "\n".join(p.text for p in doc.paragraphs) + "\n".join(
        cell.text for table in doc.tables for row in table.rows for cell in row.cells
    )


def test_multiple_employers_roles_skills_and_variable_rows(sample_template):
    text = (
        "■ 職務経歴\n\n"
        "勤務先\t架空A\n在籍期間\t2022年～現在\n\n"
        "### 営業\n\n営業職の本文です。**強調**\n\n"
        "### 企画\n\n企画職の本文です。\n\n"
        "勤務先\t架空B\n在籍期間\t2020年～2021年\n\n"
        "### 受付\n\n受付職の本文です。\n\n"
        "■ 活かせる経験・スキル\n\n"
        "経験・能力\t内容\nWord\t文書作成\nExcel\t集計\nPowerPoint\t資料編集\n\n"
        "以上"
    )
    doc = render_docx(parse_markdown(text))
    assert len(doc.tables) == 5
    assert [t.cell(0, 1).text for t in (doc.tables[0], doc.tables[2])] == ["架空A", "架空B"]
    assert [doc.tables[1].cell(i, 0).text for i in (1, 2)] == ["営業", "企画"]
    assert doc.tables[3].cell(1, 0).text == "受付"
    assert len(doc.tables[4].rows) == 4
    assert [doc.tables[4].cell(i, 0).text for i in (1, 2, 3)] == ["Word", "Excel", "PowerPoint"]
    assert doc.paragraphs[-1].text == "以上"
    assert doc.paragraphs[-1].alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert "強調" in doc.tables[1].cell(1, 1).text


def test_resume_unknown_role_remains_as_paragraph(sample_template):
    text = "■ 職務経歴\n\n勤務先\t架空A\n在籍期間\t2020年～現在\n\n【主な業務】\n\n- 来客対応"
    doc = render_docx(parse_markdown(text))
    assert len(doc.tables) == 1
    assert "・来客対応" in [p.text for p in doc.paragraphs]


def test_resume_template_absence_is_explicit(tmp_path, monkeypatch):
    monkeypatch.setattr(resume, "_TEMPLATE_PATH", tmp_path / "missing.docx")
    with pytest.raises(ConversionError, match="テンプレート"):
        render_docx(parse_markdown("職務経歴書\n\n■ 職務要約\n\n説明"))


def test_standard_body_remains_existing_renderer():
    doc = render_docx(parse_markdown("# 見出し\n\n" + "業務内容です。" * 1000, "標準文書"))
    assert doc.paragraphs[1].paragraph_format.page_break_before is None
    assert not doc.paragraphs[1].paragraph_format.keep_together
    assert doc.styles["Normal"].element.get_or_add_rPr().find(qn("w:rFonts")).get(qn("w:eastAsia")) == "游ゴシック"
    assert "PAGE" in doc.sections[0].footer._element.xml


def test_hyperlinks_remain_clickable():
    doc = render_docx(parse_markdown("[会社サイト](https://example.com)", "標準文書"))
    assert "w:hyperlink" in doc.paragraphs[0]._p.xml


def test_numbering_start_and_independence():
    src = "3. 一つ\n4. 二つ\n\n通常の文章\n\n1. 別リスト\n2. 別の項目"
    parsed = parse_markdown(src, "標準文書")
    lists = [x for x in parsed.blocks if x.kind == "list"]
    assert [x.start for x in lists] == [3, 3, 1, 1]
    assert lists[0].list_id == lists[1].list_id != lists[2].list_id
    doc = render_docx(parsed)
    numbered = [p for p in doc.paragraphs if p.style.name.startswith("List Number")]
    ids = [int(p._p.pPr.numPr.numId.val) for p in numbered]
    assert ids[0] == ids[1] and ids[2] == ids[3] and ids[0] != ids[2]
    numbering = doc.part.numbering_part.element
    for number_id, start in ((ids[0], 3), (ids[2], 1)):
        item = next(n for n in numbering.findall(qn("w:num")) if int(n.get(qn("w:numId"))) == number_id)
        assert item.find(qn("w:lvlOverride")).find(qn("w:startOverride")).get(qn("w:val")) == str(start)

