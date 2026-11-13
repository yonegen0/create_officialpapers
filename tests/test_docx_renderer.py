from docx.oxml.ns import qn
from docx.shared import Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from markdown_to_word.docx_renderer import render_docx
from markdown_to_word.parser import parse_markdown


def test_page_font_header_and_table_are_editable():
    p = parse_markdown("職務経歴書\n\n2026年11月13日現在\n\n氏名：山田 花子\n\n■ 職務経歴\n\n勤務先\t架空製作所\n在籍期間\t2023年～現在\n\n- 連絡調整")
    doc = render_docx(p)
    sec = doc.sections[0]
    assert abs(sec.page_width - Mm(210)) < 10000
    assert abs(sec.top_margin - Mm(16)) < 10000
    assert doc.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert doc.paragraphs[1].alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert len(doc.tables) == 1
    assert doc.tables[0].cell(0, 1).text == "架空製作所"
    assert any(x.style.name == "List Bullet" for x in doc.paragraphs)
    assert doc.styles["Heading 1"].paragraph_format.keep_with_next
    assert doc.styles["Normal"].element.get_or_add_rPr().find(qn("w:rFonts")).get(qn("w:eastAsia")) == "游ゴシック"
    assert "PAGE" in sec.footer._element.xml


def test_long_body_does_not_force_page_breaks():
    doc = render_docx(parse_markdown("# 見出し\n\n" + "業務内容です。" * 1000, "標準文書"))
    assert doc.paragraphs[1].paragraph_format.page_break_before is None
    assert not doc.paragraphs[1].paragraph_format.keep_together


def test_hyperlinks_remain_clickable():
    doc = render_docx(parse_markdown("[会社サイト](https://example.com)", "標準文書"))
    assert 'w:hyperlink' in doc.paragraphs[0]._p.xml


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


def test_resume_end_alignment_and_tabular_docx():
    parsed = parse_markdown("■ 職務経歴\n\n一般事務・受付（2023年10月～現在）\n\n項目\t内容\n電話\t来客対応\n\n以上")
    doc = render_docx(parsed)
    assert doc.tables[0].cell(1, 1).text == "来客対応"
    assert doc.paragraphs[-1].alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert any(p.text == "一般事務・受付（2023年10月～現在）" and p.style.name == "Heading 2" for p in doc.paragraphs)
