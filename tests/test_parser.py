from markdown_to_word.parser import parse_markdown


def plain(block):
    return "".join(x.text for x in block.spans)


def test_standard_markdown_structure():
    data = "# タイトル\n\n通常 **太字** と *斜体* [サイト](https://example.com)\n\n- 一つ\n- 二つ\n\n1. 順番\n\n| 項目 | 内容 |\n|---|---|\n| A | B |\n\n---"
    d = parse_markdown(data, "標準文書")
    assert d.headings == 1 and d.tables == 1 and d.list_items == 3
    para = next(x for x in d.blocks if x.kind == "paragraph")
    assert any(s.bold and s.text == "太字" for s in para.spans)
    assert any(s.italic and s.text == "斜体" for s in para.spans)
    assert any(s.link == "https://example.com" for s in para.spans)
    assert any(x.ordered for x in d.blocks if x.kind == "list")
    assert any(x.kind == "rule" for x in d.blocks)


def test_resume_metadata_sections_company_table_and_exact_text():
    source = """職務経歴書

2026年11月13日現在

氏名：山田 花子

■ 職務経歴

勤務先　架空製作所
在籍期間　2023年10月～現在
雇用形態　正社員
事業内容　製造業
会社規模　100名

① 来客への対応

【主な業務】

- 予約情報の確認
"""
    d = parse_markdown(source)
    assert [b.kind for b in d.blocks[:3]] == ["meta_title", "meta_date", "meta_name"]
    assert len([b for b in d.blocks if b.kind == "heading"]) == 3
    table = next(b for b in d.blocks if b.kind == "table")
    assert len(table.rows) == 5 and table.table_type == "company"
    assert table.rows[0][1][0].text == "架空製作所"
    assert d.list_items == 1


def test_tabs_and_skill_table():
    source = "経験・能力\t内容\nWord\t文書作成\nExcel\t基本操作"
    d = parse_markdown(source)
    table = d.blocks[0]
    assert table.kind == "table" and table.table_type == "skill" and table.header
    assert len(table.rows) == 3


def test_generic_tabs_and_non_table_text():
    d = parse_markdown("受付\t予約調整\n営業\t調整", "職務経歴書")
    assert d.tables == 1 and len(d.blocks[0].rows) == 2
    assert d.blocks[0].rows[0][0][0].text == "受付"
    bad = parse_markdown("単独\t行\n\n文章\n\n- 箇条\t項目\n- 箇条\t項目\n\n> 引用\t値\n> 引用\t値\n\n```\nA\tB\nC\tD\n```")
    assert bad.tables == 0
    assert parse_markdown("A\tB\nC\tD\tE").tables == 0
    assert parse_markdown("A  B\nC  D").tables == 0


def test_style_separation_and_metadata_scope():
    src = "職務経歴書\n\n氏名：架空 太郎\n\n本文です。\n\n氏名：再掲"
    d = parse_markdown(src)
    assert d.blocks[-1].kind == "paragraph"
    standard = parse_markdown(src, "標準文書")
    assert all(x.kind not in ("meta_title", "meta_name") for x in standard.blocks)


def test_code_and_html_remain_literal():
    d = parse_markdown("```html\n<script>bad()</script>\n```\n\n<b>文字</b>")
    text = "".join(plain(b) for b in d.blocks)
    assert "<script>" in text and "<b>文字</b>" in text


def test_company_table_with_blank_lines_and_separate_employers():
    source = """項目　内容

勤務先　架空会社A

在籍期間　2021年～2022年

勤務先　架空会社B

在籍期間　2023年～現在
"""
    d = parse_markdown(source)
    tables = [b for b in d.blocks if b.kind == "table"]
    assert len(tables) == 2
    assert len(tables[0].rows) == 3 and tables[0].header
    assert tables[1].rows[0][1][0].text == "架空会社B"


def test_ambiguous_single_company_row_is_warned_and_preserved():
    d = parse_markdown("勤務先　架空会社\n\n説明")
    assert d.tables == 0
    assert d.warnings
    assert "勤務先" in plain(d.blocks[0])


def test_markdown_title_is_resume_title():
    d = parse_markdown("# 職務経歴書\n\n2026年11月13日現在\n\n氏名：山田 花子")
    assert d.blocks[0].kind == "meta_title"


def test_pipe_tables_use_resume_col_widths():
    company = parse_markdown("|項目|内容|\n|---|---|\n|勤務先|架空会社|")
    assert company.blocks[0].table_type == "company"
    skills = parse_markdown("|経験・能力|内容|\n|---|---|\n|Word|文章作成|")
    assert skills.blocks[0].table_type == "skill"


def test_images_and_resume_ending_and_role_period():
    image = '![会社ロゴ](images/logo.png "企業ロゴ")'
    d = parse_markdown("■ 職務経歴\n\n一般事務・受付（2023年10月～現在）\n\n" + image + "\n\n以上")
    assert d.blocks[1].kind == "heading" and d.blocks[1].level == 2
    assert any(image in plain(b) for b in d.blocks)
    assert d.blocks[-1].kind == "meta_end"
    assert any("画像" in w for w in d.warnings)
    assert parse_markdown("以上\n\n説明").blocks[0].kind == "paragraph"
    assert parse_markdown("以上", "標準文書").blocks[0].kind == "paragraph"
