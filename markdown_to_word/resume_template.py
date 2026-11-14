"""Résumé-only renderer that duplicates *formatting* from an anonymized DOCX.

The production template must be prepared from the approved ideal DOCX. This
module deliberately does not reconstruct or approximate missing styles.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
from pathlib import Path
import re
from zipfile import BadZipFile

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.shared import Cm, Pt, RGBColor

from .errors import ConversionError
from .models import Block, ParsedDocument, Span


_TEMPLATE_PATH = Path(__file__).resolve().parent / "assets" / "resume_template.docx"
_REQUIRED_PARAGRAPHS = ("TITLE", "DATE", "NAME", "HEADING1", "HEADING2", "BODY", "END")
_REQUIRED_TABLES = ("COMPANY", "ROLE", "SKILL")


def _text(spans: tuple[Span, ...]) -> str:
    return "".join(span.text for span in spans)


def _load_template():
    if not _TEMPLATE_PATH.is_file():
        raise ConversionError("職務経歴書のWordテンプレートが見つかりません。配布ファイルを確認してください。")
    try:
        doc = Document(_TEMPLATE_PATH)
        paragraphs = {}
        tables = {}
        for para in doc.paragraphs:
            for key in _REQUIRED_PARAGRAPHS + ("HEADING3",):
                if para.text.strip() == "{{MTW_" + key + "}}":
                    paragraphs[key] = deepcopy(para._p)
        for table in doc.tables:
            for key in _REQUIRED_TABLES:
                if table.cell(0, 0).text.strip() == "{{MTW_" + key + "}}":
                    tables[key] = deepcopy(table._tbl)
        absent = (
            ["段落:" + name for name in _REQUIRED_PARAGRAPHS if name not in paragraphs]
            + ["表:" + name for name in _REQUIRED_TABLES if name not in tables]
        )
        if absent:
            raise ConversionError("Wordテンプレートに必要な書式要素がありません: " + "、".join(absent))
    except ConversionError:
        raise
    except (OSError, BadZipFile, KeyError, ValueError, IndexError) as exc:
        raise ConversionError("職務経歴書のWordテンプレートを読み込めません。破損していないか確認してください。") from exc

    # Keep the approved document's styles, page setup, headers, and footers,
    # but never include sample paragraphs or original source tables.
    body = doc.element.body
    for node in list(body):
        if node.tag != qn("w:sectPr"):
            body.remove(node)
    return doc, paragraphs, tables


def _run_format(rpr, *, size: float | None = None, bold: bool | None = None) -> None:
    """Apply only purpose-specific overrides to a cloned run format."""
    if size is not None:
        for tag in ("sz", "szCs"):
            for old in list(rpr.findall(qn("w:" + tag))):
                rpr.remove(old)
            node = OxmlElement("w:" + tag)
            node.set(qn("w:val"), str(round(size * 2)))
            rpr.append(node)
    if bold is not None:
        for old in list(rpr.findall(qn("w:b"))):
            rpr.remove(old)
        node = OxmlElement("w:b")
        node.set(qn("w:val"), "1" if bold else "0")
        rpr.append(node)


def _format_run(run, prototype_rpr, span: Span, *,
                size: float | None = None, bold: bool | None = None) -> None:
    if prototype_rpr is not None:
        run._r.insert(0, deepcopy(prototype_rpr))
    if size is not None or bold is not None:
        _run_format(run._r.get_or_add_rPr(), size=size,
                    bold=(bold or span.bold) if bold is not None else None)
    if span.bold:
        run.bold = True
    if span.italic:
        run.italic = True
    if span.code:
        run.font.name = "Consolas"


def _append_spans(para: Paragraph, spans: tuple[Span, ...], prototype_rpr,
                  *, size: float | None = None, bold: bool | None = None) -> None:
    for span in spans:
        if span.link and span.link.lower().startswith(("https://", "http://", "mailto:")):
            for n, part in enumerate(span.text.split("\n")):
                if n:
                    para.add_run().add_break()
                if not part:
                    continue
                rel = para.part.relate_to(span.link, RT.HYPERLINK, is_external=True)
                hyperlink = OxmlElement("w:hyperlink")
                hyperlink.set(qn("r:id"), rel)
                r = OxmlElement("w:r")
                rpr = deepcopy(prototype_rpr) if prototype_rpr is not None else OxmlElement("w:rPr")
                _run_format(rpr, size=size,
                            bold=(bold or span.bold) if bold is not None else None)
                if span.bold:
                    for old in list(rpr.findall(qn("w:b"))):
                        rpr.remove(old)
                    rpr.append(OxmlElement("w:b"))
                if span.italic:
                    for old in list(rpr.findall(qn("w:i"))):
                        rpr.remove(old)
                    rpr.append(OxmlElement("w:i"))
                if span.code:
                    fonts = rpr.find(qn("w:rFonts"))
                    if fonts is None:
                        fonts = OxmlElement("w:rFonts")
                        rpr.insert(0, fonts)
                    fonts.set(qn("w:ascii"), "Consolas")
                    fonts.set(qn("w:hAnsi"), "Consolas")
                for old in list(rpr.findall(qn("w:u"))):
                    rpr.remove(old)
                underline = OxmlElement("w:u")
                underline.set(qn("w:val"), "single")
                rpr.append(underline)
                r.append(rpr)
                t = OxmlElement("w:t")
                t.text = part
                r.append(t)
                hyperlink.append(r)
                para._p.append(hyperlink)
        else:
            for n, part in enumerate(span.text.split("\n")):
                run = para.add_run()
                _format_run(run, prototype_rpr, span, size=size, bold=bold)
                if n:
                    run.add_break()
                if part:
                    run.add_text(part)


def _clean_paragraph(element, *, plain_body: bool = False) -> None:
    for child in list(element):
        if child.tag != qn("w:pPr"):
            element.remove(child)
    ppr = element.find(qn("w:pPr"))
    if ppr is not None:
        # The exemplar's numbering and manual breaks are not source text.
        # _native_list() reintroduces numbering for Markdown lists.
        for child in list(ppr):
            if child.tag in (qn("w:pageBreakBefore"), qn("w:numPr")) or (
                plain_body and child.tag in (qn("w:keepNext"), qn("w:keepLines"))
            ):
                ppr.remove(child)


def _sample_rpr(paragraph_element, *, normal: bool = False):
    """Choose the placeholder run or the normal-text run, not the first label."""
    runs = paragraph_element.xpath("./w:r")
    if normal:
        for node in runs:
            rpr = node.find(qn("w:rPr"))
            if rpr is None:
                continue
            bold = rpr.find(qn("w:b"))
            if bold is None or bold.get(qn("w:val")) in ("0", "false", "off"):
                return deepcopy(rpr)
    for node in runs:
        if any("{{MTW_" in (t.text or "") for t in node.xpath("./w:t")):
            rpr = node.find(qn("w:rPr"))
            if rpr is not None:
                return deepcopy(rpr)
    for node in runs:
        rpr = node.find(qn("w:rPr"))
        if rpr is not None:
            return deepcopy(rpr)
    return None


def _new_paragraph(doc, prototypes, key: str, spans: tuple[Span, ...], *,
                   size: float | None = None, bold: bool | None = None):
    source = prototypes.get(key)
    if source is None:
        source = prototypes["BODY"]
    element = deepcopy(source)
    rpr = _sample_rpr(element, normal=(key == "BODY"))
    _clean_paragraph(element, plain_body=(key == "BODY"))
    doc.element.body.insert_element_before(element, "w:sectPr")
    para = Paragraph(element, doc._body)
    _append_spans(para, spans, rpr, size=size, bold=bold)
    return para


def _fill_cell(cell, spans: tuple[Span, ...], *,
               size: float | None = None, bold: bool | None = None) -> None:
    paras = cell.paragraphs
    if not paras:
        cell.add_paragraph()
        paras = cell.paragraphs
    first = paras[0]
    rpr = _sample_rpr(first._p, normal=(bold is False))
    _clean_paragraph(first._p)
    _append_spans(first, spans, rpr, size=size, bold=bold)
    for para in paras[1:]:
        para._p.getparent().remove(para._p)
    # Remove any original nested table/content in the prototype cell.
    for node in list(cell._tc):
        if node.tag not in (qn("w:tcPr"), first._p.tag) or (node.tag == first._p.tag and node is not first._p):
            cell._tc.remove(node)


def _row_flag(row, name: str, enabled: bool) -> None:
    trpr = row._tr.get_or_add_trPr()
    for old in list(trpr):
        if old.tag == qn("w:" + name):
            trpr.remove(old)
    if enabled:
        marker = OxmlElement("w:" + name)
        if name == "tblHeader":
            marker.set(qn("w:val"), "true")
        trpr.append(marker)


def _keep_compact_table_row(cells: tuple[tuple[Span, ...], ...]) -> bool:
    """Only keep concise, unbroken table rows on one page."""
    values = [_text(cell).strip() for cell in cells]
    return (
        bool(values)
        and all("\n" not in value for value in values)
        and all(len(value) <= 100 for value in values)
        and sum(len(value) for value in values) <= 150
    )


def _keep_compact_company_table(rows: list[tuple[tuple[Span, ...], ...]]) -> bool:
    """Join only small company tables; avoid moving long tables as a unit."""
    return (
        1 < len(rows) <= 6
        and all(_keep_compact_table_row(cells) for cells in rows)
        and sum(len(_text(cell).strip()) for cells in rows for cell in cells) <= 360
    )


def _new_table(doc, prototypes, key: str, *, header: bool, count: int):
    source = deepcopy(prototypes[key])
    old_rows = list(source.tr_lst)
    if len(old_rows) < (2 if key in ("ROLE", "SKILL") else 1):
        raise ConversionError("Wordテンプレートの表に必要な見本行がありません: " + key)
    for row in old_rows:
        source.remove(row)
    if header:
        source.append(deepcopy(old_rows[0]))
    data_row = old_rows[1] if key in ("ROLE", "SKILL") else old_rows[0]
    for _ in range(count):
        source.append(deepcopy(data_row))
    doc.element.body.insert_element_before(source, "w:sectPr")
    table = Table(source, doc._body)
    if header:
        _row_flag(table.rows[0], "cantSplit", True)
        _row_flag(table.rows[0], "tblHeader", True)
    for row in table.rows[int(header):]:
        _row_flag(row, "cantSplit", False)
        _row_flag(row, "tblHeader", False)
    return table


def _style_history_table(table: Table, doc, *, company: bool) -> None:
    """Style only company/role tables, preserving the source template."""
    section = doc.sections[-1]
    available = section.page_width - section.left_margin - section.right_margin
    # Keep the intended 17.4 cm proportions without overflowing page margins.
    total = int(min(Cm(17.4), available) / 635) * 635  # EMUs per twip
    left = round(total * 3.5 / 17.4 / 635) * 635
    widths = (left, total - left)
    table.autofit = False
    # Both template tables must have the same two-column OOXML grid; Word
    # otherwise synthesizes a narrow third column when it joins adjacent tables.
    grid = table._tbl.tblGrid
    for old in list(grid):
        grid.remove(old)
    for width in widths:
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(width // 635))
        grid.append(column)
    for row in table.rows:
        if len(row._tr.tc_lst) != 2:
            raise ConversionError("職務経歴表のテンプレートは2列である必要があります。")
        for tc in row._tr.tc_lst:
            tc_pr = tc.get_or_add_tcPr()
            for old in tc_pr.findall(qn("w:gridSpan")):
                tc_pr.remove(old)

    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.insert(0, tbl_w)
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(total // 635))

    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for old in list(borders):
        borders.remove(old)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge = OxmlElement("w:" + side)
        edge.set(qn("w:val"), "single")
        edge.set(qn("w:sz"), "8")  # 1.0 pt (eighths of a point)
        edge.set(qn("w:color"), "64748B")
        borders.append(edge)

    margins = tbl_pr.find(qn("w:tblCellMar"))
    if margins is None:
        margins = OxmlElement("w:tblCellMar")
        tbl_pr.append(margins)
    for old in list(margins):
        margins.remove(old)
    for side, twips in (("top", 100), ("left", 120),
                        ("bottom", 100), ("right", 120)):
        margin = OxmlElement("w:" + side)
        margin.set(qn("w:w"), str(twips))
        margin.set(qn("w:type"), "dxa")
        margins.append(margin)

    for ci, width in enumerate(widths):
        table.columns[ci].width = width
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            cell.width = widths[ci]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            tc_pr = cell._tc.get_or_add_tcPr()
            # Prototype cell borders/margins would override the table-wide design.
            for old in list(tc_pr):
                if old.tag in (qn("w:tcBorders"), qn("w:tcMar")):
                    tc_pr.remove(old)
            shading = tc_pr.find(qn("w:shd"))
            if shading is None:
                shading = OxmlElement("w:shd")
                tc_pr.append(shading)
            shading.attrib.clear()
            shading.set(qn("w:val"), "clear")
            shade = ("E7F0F7" if ci == 0 else "FFFFFF") if company else (
                "244F77" if ri == 0 else "F1F6FA" if ci == 0 else "FFFFFF"
            )
            shading.set(qn("w:fill"), shade)
            if not company and ri == 0:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def _draw_company(doc, prototypes, block: Block) -> None:
    rows = [row for row in block.rows
            if not (len(row) == 2 and _text(row[0]).strip() == "項目"
                    and _text(row[1]).strip() == "内容")]
    if not rows:
        return
    table = _new_table(doc, prototypes, "COMPANY", header=False, count=len(rows))
    compact_table = _keep_compact_company_table(rows)
    for ri, cells in enumerate(rows):
        for ci in range(min(2, len(cells))):
            _fill_cell(table.cell(ri, ci), cells[ci], size=10, bold=(ci == 0))
        _row_flag(table.rows[ri], "cantSplit", _keep_compact_table_row(cells))
        # cantSplit keeps a short row together. Only the end-of-row paragraph
        # needs keepNext to link it to the next row; two copies are redundant.
        for ci, cell in enumerate(table.rows[ri].cells):
            paragraph = cell.paragraphs[-1]
            paragraph.paragraph_format.keep_with_next = (
                compact_table and ci == 1 and ri < len(rows) - 1
            )
            # A lengthy company value must not inherit keepLines from the
            # short template sample when its row is allowed to split.
            if ci < len(cells):
                value = _text(cells[ci])
                if len(value) > 110 or "\n" in value:
                    paragraph.paragraph_format.keep_together = False
    _style_history_table(table, doc, company=True)


def _draw_skill(doc, prototypes, block: Block) -> None:
    rows = list(block.rows)
    header = bool(rows and len(rows[0]) == 2
                  and (_text(rows[0][0]).strip(), _text(rows[0][1]).strip()) == ("経験・能力", "内容"))
    if header:
        rows.pop(0)
    table = _new_table(doc, prototypes, "SKILL", header=header, count=len(rows))
    if header:
        _fill_cell(table.cell(0, 0), block.rows[0][0], size=9.5, bold=True)
        _fill_cell(table.cell(0, 1), block.rows[0][1], size=9.5, bold=True)
    for ri, cells in enumerate(rows, start=int(header)):
        for ci in range(min(2, len(cells))):
            _fill_cell(table.cell(ri, ci), cells[ci], size=9, bold=(ci == 0))
        _row_flag(table.rows[ri], "cantSplit", _keep_compact_table_row(cells))


_CIRCLED_PREFIX = re.compile(r"^[①-⑳]\s*")
_ROLE_PERIOD = re.compile(
    r"^.+（\d{4}年\s*\d{1,2}月\s*[～〜－-]\s*(?:現在|\d{4}年\s*\d{1,2}月)）$"
)
_ROLE_HINT = re.compile(
    r"営業|企画|受付|事務|人事|総務|経理|開発|技術|設計|製造|品質|"
    r"販売|接客|広報|物流|購買|管理|サポート|エンジニア|デザイナー"
)
_SUBHEADINGS = frozenset((
    "主な成果", "主な実績", "担当業務", "業務内容", "主な業務",
    "工夫した点", "身につけた能力", "身につけたスキル",
    "取り組み", "成果", "実績", "課題", "改善内容",
))


def _in_tasks_heading(block: Block) -> bool:
    return block.kind == "heading" and _text(block.spans).strip() == "【主な業務】"


def _is_subheading(block: Block) -> bool:
    if block.kind != "heading" or block.level < 2:
        return False
    content = _text(block.spans).strip()
    label = _CIRCLED_PREFIX.sub("", content)
    if label.startswith("【"):
        return True
    if label in _SUBHEADINGS or any(
        label.startswith(title + sep)
        for title in _SUBHEADINGS for sep in ("：", ":")
    ):
        return True
    # A circled-number episode is never a role merely because it is a heading.
    return label != content and not (
        label.startswith("職種：") or _ROLE_PERIOD.fullmatch(label)
    )


def _is_role(block: Block, *, first: bool, following: Block | None) -> bool:
    if block.kind != "heading" or block.level < 2 or _is_subheading(block):
        return False
    content = _text(block.spans).strip()
    if not content or content.startswith(("■", "【")):
        return False
    if content.startswith("職種："):
        return bool(content.removeprefix("職種：").strip())
    label = _CIRCLED_PREFIX.sub("", content)
    if _ROLE_PERIOD.fullmatch(label):
        return True
    if label != content or following is None or following.kind in ("table", "rule"):
        return False
    # A short, unnumbered heading after company information is a legacy role.
    # Later roles also need role-like wording; ambiguous headings stay in the
    # preceding role's right cell instead of creating spurious rows.
    if len(label) > 24 or re.search(r"[。！？!?:：\n]", label):
        return False
    if label.endswith(("こと", "方法", "理由", "ポイント", "取り組み", "経験",
                       "成果", "実績", "能力", "スキル", "課題")):
        return False
    return first or bool(_ROLE_HINT.search(label))


def _display_role_spans(spans: tuple[Span, ...]) -> tuple[Span, ...]:
    """Remove only the optional role marker, retaining inline formatting."""
    if not _text(spans).startswith("職種："):
        return spans
    skip = len("職種：")
    converted: list[Span] = []
    for span in spans:
        if skip >= len(span.text):
            skip -= len(span.text)
            continue
        value = span.text[skip:] if skip else span.text
        skip = 0
        if not converted:
            value = value.lstrip()
        if value:
            converted.append(Span(value, span.bold, span.italic, span.link,
                                  span.code, span.break_line))
    return tuple(converted)


def _is_repeated_role_header(block: Block) -> bool:
    """Suppress only the duplicated fixed label, never candidate prose."""
    if block.kind not in ("heading", "paragraph"):
        return False
    label = re.sub(r"\s+", "", _text(block.spans))
    return label in (
        "職種・期間と身につけた能力／経験・主な業務",
        "職種・期間／身につけた能力／経験・主な業務",
    )


def _break_role_period(spans: tuple[Span, ...]) -> tuple[Span, ...]:
    """Split the trailing period without losing per-span Markdown formatting."""
    text = _text(spans)
    match = re.search(
        r"（\d{4}年\s*\d{1,2}月\s*[～〜－-]\s*(?:現在|\d{4}年\s*\d{1,2}月)）$",
        text,
    )
    if not match or match.start() == 0 or text[match.start() - 1] == "\n":
        return spans
    offset = match.start()
    output = []
    for span in spans:
        length = len(span.text)
        if 0 < offset < length:
            output.append(replace(span, text=span.text[:offset]))
            output.append(replace(span, text="\n" + span.text[offset:]))
        elif offset == 0:
            output.append(replace(span, text="\n" + span.text))
        else:
            output.append(span)
        offset -= length
    return tuple(output)


def _wrap_long_role_period(spans: tuple[Span, ...]) -> tuple[Span, ...]:
    """Break long year-month ranges after the separator, not inside a date."""
    text = _text(spans)
    match = re.search(
        r"（?(\d{4}年\s*(\d{1,2})月\s*[～〜－-])\s*(現在|\d{4}年\s*\d{1,2}月)）?$",
        text,
    )
    if not match or (
        match.group(3) == "現在"
        and (not match.group(0).startswith("（") or len(match.group(2)) < 2)
    ):
        return spans
    offset = match.end(1)
    if text[offset:offset + 1] == "\n":
        return spans
    output = []
    for span in spans:
        length = len(span.text)
        if 0 < offset < length:
            output.append(replace(span, text=span.text[:offset]))
            output.append(replace(span, text="\n" + span.text[offset:]))
        elif offset == 0:
            output.append(replace(span, text="\n" + span.text))
        else:
            output.append(span)
        offset -= length
    return tuple(output)


def _emphasize_ability_label(spans: tuple[Span, ...]) -> tuple[Span, ...]:
    """Emphasize only the opening label and preserve the original text."""
    match = re.match(r"^(?:身につけた能力|身につけたスキル)[：:]", _text(spans))
    if not match:
        return spans
    remaining = match.end()
    output = []
    for span in spans:
        take = min(remaining, len(span.text))
        if take:
            output.append(replace(span, text=span.text[:take], bold=True))
        if take < len(span.text):
            output.append(replace(span, text=span.text[take:]))
        remaining -= take
    return tuple(output)


def _keep_compact_role(left: tuple[Span, ...], entries: list[Block]) -> bool:
    """Keep a handful of short paragraphs; let Word split lengthy role rows."""
    values = [_text(b.spans) for b in entries]
    return (
        len(entries) <= 3
        and all("\n" not in value for value in values)
        and all(len(value) <= 110 for value in values)
        and len(_text(left)) + sum(len(value) for value in values) <= 240
    )


def _write_cell_job(cell, entries: list[Block], list_ids: dict[int, int], doc,
                    *, link_trailing_bullets: bool = False) -> None:
    # Separate paragraph exemplars for narrative text, task headings and bullets.
    source = [deepcopy(p._p) for p in cell.paragraphs]
    if len(source) < 3:
        raise ConversionError("Wordテンプレートの職務内容書式が不足しています。")
    for node in list(cell._tc):
        if node.tag != qn("w:tcPr"):
            cell._tc.remove(node)
    tasks = False
    narrative_index = 0
    for index, block in enumerate(entries):
        subheading = _is_subheading(block)
        if _in_tasks_heading(block):
            tasks = True
            template = source[1]
        elif subheading:
            tasks = False
            template = source[1]
        elif block.kind == "list" and tasks and not block.ordered:
            template = source[2]
        else:
            template = source[0]
            if block.kind != "list":
                tasks = False
        element = deepcopy(template)
        rpr = _sample_rpr(element, normal=not subheading)
        _clean_paragraph(element)
        cell._tc.append(element)
        para = Paragraph(element, cell)
        spans = block.spans
        # Body copy stays regular, even when the source Markdown has **bold**.
        # Purpose-specific labels/subheadings are applied separately below.
        if block.kind in ("paragraph", "list") and not subheading:
            spans = tuple(replace(span, bold=False) for span in spans)
        literal_bullet = block.kind == "list" and tasks and not block.ordered
        if literal_bullet:
            if not _text(spans).startswith("・"):
                spans = (Span("・"),) + spans
            if para.style.name.startswith("List "):
                para.style = doc.styles["Normal"]
            # The literal ・ starts at the margin; wrapped lines align to text.
            para.paragraph_format.left_indent = Pt(10)
            para.paragraph_format.first_line_indent = Pt(-10)
        elif block.kind == "list":
            _native_list(para, block, list_ids, doc)
        elif block.kind == "paragraph":
            spans = _emphasize_ability_label(spans)
            if narrative_index:
                para.paragraph_format.first_line_indent = Pt(9.9)
            narrative_index += 1
        if block.kind in ("paragraph", "list") and not subheading:
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
        # Never chain the whole list. Only the last two concise, flat items
        # in the same task list may travel together in a splittable role row.
        short_tail_pair = False
        if link_trailing_bullets and literal_bullet and index + 1 < len(entries):
            next_block = entries[index + 1]
            later = entries[index + 2] if index + 2 < len(entries) else None
            same_pair = (
                next_block.kind == "list" and not next_block.ordered
                and block.level == next_block.level == 1
                and block.list_id == next_block.list_id
                and all("\n" not in _text(item.spans)
                        and len(_text(item.spans)) <= 45
                        for item in (block, next_block))
            )
            ends_list = later is None or not (
                later.kind == "list" and later.list_id == block.list_id
                and later.level == block.level and later.ordered == block.ordered
            )
            short_tail_pair = same_pair and ends_list
        para.paragraph_format.keep_with_next = subheading or short_tail_pair
        para.paragraph_format.keep_together = subheading
        para.paragraph_format.widow_control = True
        _append_spans(para, spans, rpr, size=10.5,
                      bold=True if subheading else False if block.kind in ("paragraph", "list") else None)
    if not entries:
        element = deepcopy(source[0])
        _clean_paragraph(element)
        cell._tc.append(element)


def _draw_roles(doc, prototypes, roles: list[tuple[Block, list[Block]]], list_ids) -> None:
    if not roles:
        return
    # The ideal Word includes a genuine header; remove only duplicate labels.
    table = _new_table(doc, prototypes, "ROLE", header=True, count=len(roles))
    _fill_cell(table.cell(0, 0), (Span("職種・期間"),), size=9.5, bold=True)
    _fill_cell(table.cell(0, 1), (Span("身につけた能力／経験・主な業務"),), size=9.5, bold=True)
    for index, (role, items) in enumerate(roles, 1):
        row = table.rows[index]
        left_spans = _break_role_period(_display_role_spans(role.spans))
        detail_items = items
        if items and items[0].kind == "paragraph" and re.fullmatch(
            r"\d{4}年\s*\d{1,2}月\s*[～〜－-]\s*(?:現在|\d{4}年\s*\d{1,2}月)",
            _text(items[0].spans).strip(),
        ):
            left_spans = left_spans + (Span("\n"),) + items[0].spans
            detail_items = items[1:]
        left_spans = _wrap_long_role_period(left_spans)
        _fill_cell(row.cells[0], left_spans, size=9.5, bold=True)
        # A role label must not inherit keepNext from its template sample:
        # the row's cantSplit flag already protects compact role entries.
        row.cells[0].paragraphs[0].paragraph_format.keep_with_next = False
        compact_role = _keep_compact_role(left_spans, detail_items)
        _write_cell_job(row.cells[1], detail_items, list_ids, doc,
                        link_trailing_bullets=not compact_role)
        _row_flag(row, "cantSplit", compact_role)
    _style_history_table(table, doc, company=False)


def _native_list(para, block: Block, list_ids: dict[int, int], doc) -> None:
    name = "List Number" if block.ordered else "List Bullet"
    if block.level > 1:
        name += " " + str(min(block.level, 3))
    if name not in doc.styles:
        raise ConversionError("Wordテンプレートにリストの書式がありません: " + name)
    para.style = doc.styles[name]
    if block.ordered:
        from .docx_renderer import _ordered_numbering
        if block.list_id not in list_ids:
            list_ids[block.list_id] = _ordered_numbering(doc, name, block.start)
        num = para._p.get_or_add_pPr().get_or_add_numPr()
        num.get_or_add_ilvl().val = 0
        num.get_or_add_numId().val = list_ids[block.list_id]


def _draw_general_table(doc, block: Block) -> None:
    # Generic Markdown tables keep the existing implementation and mechanics.
    from .docx_renderer import _draw_table
    _draw_table(doc, block)
    # For generic Markdown tables use the existing table algorithm, but the
    # typography must inherit the approved template's Normal style.
    for cell in doc.tables[-1]._cells:
        for para in cell.paragraphs:
            for run in para.runs:
                rpr = run._r.get_or_add_rPr()
                fonts = rpr.find(qn("w:rFonts"))
                if fonts is None:
                    fonts = OxmlElement("w:rFonts")
                    rpr.insert(0, fonts)
                for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                    # The ideal Word explicitly uses Meiryo UI on its runs;
                    # its Normal style has a different fallback font.
                    fonts.set(qn("w:" + attr), "Meiryo UI")


def _draw_block(doc, paragraphs, tables, block: Block, list_ids, *, tasks=False,
                first_line_indent: float | None = None,
                episode_link_next: bool = False):
    if block.kind == "table":
        if block.table_type == "company":
            _draw_company(doc, tables, block)
        elif block.table_type == "skill":
            _draw_skill(doc, tables, block)
        else:
            _draw_general_table(doc, block)
        return
    keys = {
        "meta_title": "TITLE", "meta_date": "DATE", "meta_name": "NAME",
        "meta_end": "END", "heading": "HEADING" + str(min(block.level, 3)),
    }
    key = "BODY" if _in_tasks_heading(block) else keys.get(block.kind, "BODY")
    spans = block.spans
    body_text = block.kind in ("paragraph", "list") and not _in_tasks_heading(block)
    if body_text:
        spans = tuple(replace(span, bold=False) for span in spans)
    if block.kind == "list" and tasks and not block.ordered:
        raw = _text(spans)
        spans = (Span("・" if not raw.startswith("・") else ""),) + spans
    p = _new_paragraph(doc, paragraphs, key, spans,
                       size=20 if block.kind == "meta_title" else None,
                       bold=True if block.kind == "meta_title" else False if body_text else None)
    if first_line_indent is not None:
        p.paragraph_format.first_line_indent = Pt(first_line_indent)
    if block.kind == "meta_title":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        for run in p.runs:
            run.font.name = "Meiryo UI"
            fonts = run._r.get_or_add_rPr().find(qn("w:rFonts"))
            if fonts is not None:
                fonts.set(qn("w:eastAsia"), "Meiryo UI")
    if block.kind == "list" and not (tasks and not block.ordered):
        _native_list(p, block, list_ids, doc)
    if body_text:
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.keep_with_next = episode_link_next
        # A long paragraph must be free to split even if its template
        # exemplar happens to request keeping all lines on one page.
        if len(_text(spans)) > 110 or "\n" in _text(spans):
            p.paragraph_format.keep_together = False
    if block.kind == "heading" and block.level == 1 and _text(block.spans).lstrip().startswith(("■", "▪")):
        # Word displays a small black square in the editing margin for
        # paragraph pagination controls. Do not apply them to section titles.
        # An explicit False also overrides inherited heading-style settings.
        p.paragraph_format.keep_with_next = False
        p.paragraph_format.keep_together = False
        p.paragraph_format.page_break_before = False
        p.paragraph_format.widow_control = False
    elif block.kind == "heading" or _in_tasks_heading(block):
        p.paragraph_format.keep_with_next = True
    if block.kind in ("meta_date", "meta_name", "meta_end"):
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if block.kind == "meta_end":
        p.paragraph_format.keep_with_next = False
    if block.kind == "rule":
        from .docx_renderer import _rule
        _rule(p)
    if block.kind == "code":
        for run in p.runs:
            run.font.name = "Consolas"


def _work_after_company(doc, paragraphs, tables, blocks: tuple[Block, ...], list_ids) -> None:
    roles: list[tuple[Block, list[Block]]] = []
    tasks = False

    def flush() -> None:
        if roles:
            _draw_roles(doc, tables, roles, list_ids)
            roles.clear()

    for index, block in enumerate(blocks):
        following = blocks[index + 1] if index + 1 < len(blocks) else None
        if _is_repeated_role_header(block):
            continue
        if _is_role(block, first=not roles, following=following):
            roles.append((block, []))
            tasks = False
        elif roles and block.kind != "table":
            roles[-1][1].append(block)
        else:
            flush()
            if _in_tasks_heading(block):
                tasks = True
            elif block.kind != "list":
                tasks = False
            _draw_block(doc, paragraphs, tables, block, list_ids, tasks=tasks)
    flush()


def _format_page_fields(doc) -> None:
    """Match existing PAGE fields without adding or duplicating them."""
    for section in doc.sections:
        for para in section.footer.paragraphs:
            instructions = para._p.xpath(".//w:instrText")
            if not any(re.search(r"\bPAGE\b", node.text or "") for node in instructions):
                continue
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(0x6E, 0x6E, 0x6E)


def render_resume(parsed: ParsedDocument):
    """Build a résumé from Markdown while cloning original DOCX formatting.

    A missing/unapproved template is always an error, never a quiet fallback.
    """
    doc, paragraphs, tables = _load_template()
    list_ids: dict[int, int] = {}
    in_history = False
    tasks = False
    section = ""
    section_paragraph_count = 0
    episode_active = False
    episode_paragraph_count = 0
    blocks = parsed.blocks
    i = 0
    while i < len(blocks):
        block = blocks[i]
        if block.kind == "heading" and block.level == 1:
            section = _text(block.spans).strip().removeprefix("■").strip()
            in_history = section == "職務経歴"
            tasks = False
            section_paragraph_count = 0
            episode_active = False
            episode_paragraph_count = 0
        elif (section == "主な経験・強み" and block.kind == "heading"
              and block.level >= 2):
            episode_active = bool(_CIRCLED_PREFIX.match(_text(block.spans).strip()))
            episode_paragraph_count = 0
        if in_history and block.kind == "table" and block.table_type == "company":
            _draw_company(doc, tables, block)
            j = i + 1
            while j < len(blocks):
                candidate = blocks[j]
                if candidate.kind == "table" and candidate.table_type == "company":
                    break
                if candidate.kind == "heading" and candidate.level == 1:
                    break
                j += 1
            _work_after_company(doc, paragraphs, tables, blocks[i + 1:j], list_ids)
            i = j
            continue
        if _in_tasks_heading(block):
            tasks = True
        elif block.kind != "list":
            tasks = False
        indent = None
        link_episode_intro = False
        if block.kind == "paragraph":
            if section == "職務要約":
                indent = 0 if section_paragraph_count == 0 else 10.5
                section_paragraph_count += 1
            elif section == "主な経験・強み" and episode_active:
                indent = 0 if episode_paragraph_count == 0 else 10.5
                if episode_paragraph_count == 0 and i + 1 < len(blocks):
                    following = blocks[i + 1]
                    intro = _text(block.spans)
                    continuation = (_text(following.spans)
                                    if following.kind == "paragraph" else "")
                    link_episode_intro = (
                        following.kind == "paragraph"
                        and all(value and "\n" not in value and len(value) <= 65
                                for value in (intro, continuation))
                    )
                episode_paragraph_count += 1
        _draw_block(doc, paragraphs, tables, block, list_ids, tasks=tasks,
                    first_line_indent=indent, episode_link_next=link_episode_intro)
        i += 1
    _format_page_fields(doc)
    return doc
