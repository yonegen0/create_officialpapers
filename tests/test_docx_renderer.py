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



def test_bundled_real_template_roundtrip(tmp_path):
    """Exercise the approved, anonymized DOCX instead of only the fake fixture."""
    source = (
        "職務経歴書\n\n2026年11月14日現在\n\n氏名：架空 花子\n\n"
        "■ 職務要約\n\n予約内容を確認し、関係者と調整してきました。\n\n"
        "■ 職務経歴\n\n"
        "勤務先\t架空A\n在籍期間\t2023年4月～現在\n雇用形態\t正社員\n\n"
        "### 受付（2023年4月～現在）\n\n来客対応を行いました。\n\n"
        "【主な業務】\n\n- 電話対応\n- 予約確認\n\n"
        "勤務先\t架空B\n在籍期間\t2021年4月～2023年3月\n\n"
        "### 一般事務\n\n資料作成を行いました。\n\n"
        "■ 活かせる経験・スキル\n\n"
        "経験・能力\t内容\nWord\t案内文作成\nExcel\t入力管理\n\n"
        "■ 自己PR\n\n相手の状況を確認しています。\n\n以上"
    )
    rendered = render_docx(parse_markdown(source))
    file = tmp_path / "resume.docx"
    rendered.save(file)
    reopened = Document(file)
    assert len(reopened.tables) == 5
    assert [reopened.tables[i].cell(0, 1).text for i in (0, 2)] == ["架空A", "架空B"]
    assert [reopened.tables[1].cell(1, 1).paragraphs[i].text for i in range(4)] == [
        "来客対応を行いました。", "【主な業務】", "・電話対応", "・予約確認"
    ]
    assert [reopened.tables[4].cell(i, 0).text for i in (1, 2)] == ["Word", "Excel"]
    assert [reopened.paragraphs[0].text, reopened.paragraphs[-1].text] == ["職務経歴書", "以上"]
    all_text = "\n".join(p.text for p in reopened.paragraphs)
    all_text += "\n".join(c.text for t in reopened.tables for row in t.rows for c in row.cells)
    assert "山崎製パン" not in all_text and "{{MTW_" not in all_text
    assert file.stat().st_size > 1000


def _has_row_flag(row, flag):
    trpr = row._tr.trPr
    return trpr is not None and trpr.find(qn("w:" + flag)) is not None


def test_role_subheadings_remain_in_right_cell(sample_template):
    source = (
        "■ 職務経歴\n\n勤務先\t架空A\n在籍期間\t2023年～現在\n\n"
        "### 受付・一般事務\n\n来客対応や予約管理を担当しました。\n\n"
        "### 主な成果\n\n予約内容の確認方法を改善しました。\n\n"
        "### 工夫した点\n\n関係者への事前確認を徹底しました。\n\n"
        "【主な業務】\n\n- 来客対応\n- 予約管理"
    )
    doc = render_docx(parse_markdown(source))
    assert len(doc.tables) == 2
    role_table = doc.tables[1]
    assert len(role_table.rows) == 2
    assert role_table.cell(1, 0).text == "受付・一般事務"
    paragraphs = role_table.cell(1, 1).paragraphs
    assert [p.text for p in paragraphs] == [
        "来客対応や予約管理を担当しました。",
        "主な成果", "予約内容の確認方法を改善しました。",
        "工夫した点", "関係者への事前確認を徹底しました。",
        "【主な業務】", "・来客対応", "・予約管理",
    ]
    assert all(paragraphs[i].paragraph_format.keep_with_next for i in (1, 3, 5))
    assert all(paragraphs[i].paragraph_format.keep_with_next is False for i in (0, 2, 4, 6, 7))


def test_circled_role_and_episode_headings(sample_template):
    source = (
        "■ 職務経歴\n\n勤務先\t架空A\n在籍期間\t2023年～現在\n\n"
        "### ① 受付（2023年4月～現在）\n\n来客対応を担当しました。\n\n"
        "① 来客対応で工夫したこと\n\n事前に確認しました。\n\n"
        "### ① 来客対応の事前準備\n\n手順を整理しました。\n\n"
        "① 受付業務\n\n内容を説明しました。\n\n"
        "### 職種：企画\n\n企画業務を担当しました。"
    )
    doc = render_docx(parse_markdown(source))
    assert len(doc.tables) == 2
    role_table = doc.tables[1]
    assert len(role_table.rows) == 3
    assert role_table.cell(1, 0).text == "① 受付\n（2023年4月～現在）"
    assert role_table.cell(2, 0).text == "企画"
    first_paragraphs = role_table.cell(1, 1).paragraphs
    assert [p.text for p in first_paragraphs] == [
        "来客対応を担当しました。",
        "① 来客対応で工夫したこと", "事前に確認しました。",
        "① 来客対応の事前準備", "手順を整理しました。",
        "① 受付業務", "内容を説明しました。",
    ]
    assert all(p._p.pPr is None or p._p.pPr.numPr is None for p in first_paragraphs)
    assert role_table.cell(2, 1).text == "企画業務を担当しました。"


def test_resume_row_break_controls(sample_template):
    source = (
        "■ 職務経歴\n\n勤務先\t架空A\n在籍期間\t" + "2023年～現在" + "\n\n"
        "### 受付\n\n短い説明です。\n\n"
        "### 営業\n\n" + ("長い説明です。" * 40) + "\n\n"
        "### 工夫した点\n\n" + ("改行を含む文章です。\n" * 6) + "\n\n"
        "■ 活かせる経験・スキル\n\n経験・能力\t内容\n"
        "Word\t文書作成\nExcel\t" + ("詳細な説明。" * 45)
    )
    doc = render_docx(parse_markdown(source))
    company, roles, skills = doc.tables
    assert all(_has_row_flag(row, "cantSplit") for row in company.rows)
    assert _has_row_flag(roles.rows[0], "tblHeader")
    assert _has_row_flag(roles.rows[0], "cantSplit")
    assert _has_row_flag(roles.rows[1], "cantSplit")
    assert not _has_row_flag(roles.rows[2], "cantSplit")
    assert _has_row_flag(skills.rows[0], "tblHeader")
    assert _has_row_flag(skills.rows[1], "cantSplit")
    assert not _has_row_flag(skills.rows[2], "cantSplit")
    heading = next(p for p in roles.cell(2, 1).paragraphs if p.text == "工夫した点")
    assert heading.paragraph_format.keep_with_next


def test_resume_purpose_specific_format_and_periods(tmp_path):
    """Roundtrip the real template, not just synthetic styles."""
    source = (
        "職務経歴書\n\n2026年11月14日現在\n\n氏名：架空 花子\n\n"
        "■ 職務要約\n\n最初の段落です。\n\n次の段落です。\n\n"
        "■ 職務経歴\n\n"
        "勤務先\t架空A\n在籍期間\t2023年4月～現在\n\n"
        "### **受付**・一般事務（2023年4月～現在）\n\n"
        "身につけた能力：日程管理と情報共有\n\n"
        "事前に担当者と確認しました。\n\n"
        "【主な業務】\n\n- 来客対応\n- ・予約管理\n\n"
        "■ 主な経験・強み\n\n"
        "① 事前準備\n\n初めの説明です。\n\n後続の説明です。\n\n"
        "■ 活かせる経験・スキル\n\n"
        "経験・能力\t内容\nWord\t文書作成\nExcel\t表計算\n\n以上"
    )
    result = render_docx(parse_markdown(source))
    path = tmp_path / "formatted.docx"
    result.save(path)
    doc = Document(path)

    title = doc.paragraphs[0]
    assert title.text == "職務経歴書"
    assert title.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert title.runs[0].font.size.pt == 20
    assert title.runs[0].bold

    summary = next(i for i, p in enumerate(doc.paragraphs) if p.text == "最初の段落です。")
    assert doc.paragraphs[summary].paragraph_format.first_line_indent.pt == 0
    assert doc.paragraphs[summary + 1].paragraph_format.first_line_indent.pt == 10.5

    company, roles, skills = doc.tables
    assert company.cell(0, 0).paragraphs[0].runs[0].font.size.pt == 10
    assert company.cell(0, 0).paragraphs[0].runs[0].bold
    assert company.cell(0, 1).paragraphs[0].runs[0].font.size.pt == 10
    assert roles.cell(0, 0).paragraphs[0].runs[0].font.size.pt == 9.5
    assert roles.cell(0, 0).paragraphs[0].runs[0].bold
    assert roles.cell(1, 0).text == "受付・一般事務\n（2023年4月～現在）"
    assert all(run.font.size.pt == 9.5 for run in roles.cell(1, 0).paragraphs[0].runs)
    descriptions = roles.cell(1, 1).paragraphs
    assert descriptions[0].text == "身につけた能力：日程管理と情報共有"
    assert descriptions[0].runs[0].bold
    assert all(r.font.size.pt == 9 for p in descriptions for r in p.runs if r.text)
    assert descriptions[1].paragraph_format.first_line_indent.pt == 9.9
    assert [p.text for p in descriptions[-2:]] == ["・来客対応", "・予約管理"]

    strong = next(i for i, p in enumerate(doc.paragraphs) if p.text == "初めの説明です。")
    assert all(r.bold for r in doc.paragraphs[strong].runs if r.text)
    assert all(not r.bold for r in doc.paragraphs[strong + 1].runs if r.text)
    assert doc.paragraphs[strong + 1].paragraph_format.first_line_indent.pt == 10.5

    assert skills.cell(0, 0).paragraphs[0].runs[0].font.size.pt == 9.5
    assert skills.cell(1, 0).paragraphs[0].runs[0].bold
    assert skills.cell(1, 1).paragraphs[0].runs[0].font.size.pt == 9
    assert "PAGE" in doc.sections[0].footer._element.xml


def test_template_rpr_selector_and_unwanted_role_labels(sample_template):
    """A first bold label must not contaminate the narrative run's style."""
    from docx.shared import Pt
    sample = Document(sample_template)
    paragraph = sample.tables[1].cell(1, 1).paragraphs[0]
    paragraph.clear()
    a = paragraph.add_run("見出し：")
    a.bold = True
    a.font.size = Pt(12)
    b = paragraph.add_run("{{MTW_ROLE_DETAIL}}")
    b.bold = False
    b.font.size = Pt(9)
    sample.save(sample_template)

    source = (
        "■ 職務経歴\n\n勤務先\t架空A\n在籍期間\t2023年～現在\n\n"
        "職種・期間と身につけた能力／経験・主な業務\n\n"
        "### 受付\n\n通常の説明です。\n\n"
        "### 主な成果\n\n改善しました。\n\n"
        "【主な業務】\n\n- ・電話対応"
    )
    doc = render_docx(parse_markdown(source))
    assert len(doc.tables) == 2
    roles = doc.tables[1]
    assert len(roles.rows) == 2
    assert roles.cell(1, 0).text == "受付"
    assert roles.cell(1, 1).paragraphs[0].runs[0].font.size.pt == 9
    assert [p.text for p in roles.cell(1, 1).paragraphs][-1] == "・電話対応"
    assert "職種・期間と身につけた能力" not in "\n".join(p.text for p in doc.paragraphs)
