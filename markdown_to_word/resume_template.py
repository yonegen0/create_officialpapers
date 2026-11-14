"""Résumé-only renderer that duplicates *formatting* from an anonymized DOCX.

The production template must be prepared from the approved ideal DOCX. This
module deliberately does not reconstruct or approximate missing styles.
"""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from zipfile import BadZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.table import Table
from docx.text.paragraph import Paragraph

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


def _format_run(run, prototype_rpr, span: Span) -> None:
    if prototype_rpr is not None:
        run._r.insert(0, deepcopy(prototype_rpr))
    if span.bold:
        run.bold = True
    if span.italic:
        run.italic = True
    if span.code:
        run.font.name = "Consolas"


def _append_spans(para: Paragraph, spans: tuple[Span, ...], prototype_rpr) -> None:
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
                if span.bold:
                    rpr.append(OxmlElement("w:b"))
                if span.italic:
                    rpr.append(OxmlElement("w:i"))
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
                _format_run(run, prototype_rpr, span)
                if n:
                    run.add_break()
                if part:
                    run.add_text(part)


def _clean_paragraph(element) -> None:
    for child in list(element):
        if child.tag != qn("w:pPr"):
            element.remove(child)
    ppr = element.find(qn("w:pPr"))
    if ppr is not None:
        # Never replicate intentional page breaks from an exemplar.
        for child in list(ppr):
            if child.tag == qn("w:pageBreakBefore"):
                ppr.remove(child)


def _first_rpr(paragraph_element):
    nodes = paragraph_element.xpath(".//w:rPr")
    return deepcopy(nodes[0]) if nodes else None


def _new_paragraph(doc, prototypes, key: str, spans: tuple[Span, ...]):
    source = prototypes.get(key)
    if source is None:
        source = prototypes["BODY"]
    element = deepcopy(source)
    rpr = _first_rpr(element)
    _clean_paragraph(element)
    doc.element.body.insert_element_before(element, "w:sectPr")
    para = Paragraph(element, doc._body)
    _append_spans(para, spans, rpr)
    return para


def _fill_cell(cell, spans: tuple[Span, ...]) -> None:
    paras = cell.paragraphs
    if not paras:
        cell.add_paragraph()
        paras = cell.paragraphs
    first = paras[0]
    rpr = _first_rpr(first._p)
    _clean_paragraph(first._p)
    _append_spans(first, spans, rpr)
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


def _draw_company(doc, prototypes, block: Block) -> None:
    rows = [row for row in block.rows
            if not (len(row) == 2 and _text(row[0]).strip() == "項目"
                    and _text(row[1]).strip() == "内容")]
    if not rows:
        return
    table = _new_table(doc, prototypes, "COMPANY", header=False, count=len(rows))
    for ri, cells in enumerate(rows):
        for ci in range(min(2, len(cells))):
            _fill_cell(table.cell(ri, ci), cells[ci])
        _row_flag(table.rows[ri], "cantSplit", True)


def _draw_skill(doc, prototypes, block: Block) -> None:
    rows = list(block.rows)
    header = bool(rows and len(rows[0]) == 2
                  and (_text(rows[0][0]).strip(), _text(rows[0][1]).strip()) == ("経験・能力", "内容"))
    if header:
        rows.pop(0)
    table = _new_table(doc, prototypes, "SKILL", header=header, count=len(rows))
    if header:
        _fill_cell(table.cell(0, 0), block.rows[0][0])
        _fill_cell(table.cell(0, 1), block.rows[0][1])
    for ri, cells in enumerate(rows, start=int(header)):
        for ci in range(min(2, len(cells))):
            _fill_cell(table.cell(ri, ci), cells[ci])
        _row_flag(table.rows[ri], "cantSplit", sum(len(s.text) for c in cells for s in c) < 180)


def _in_tasks_heading(block: Block) -> bool:
    return block.kind == "heading" and _text(block.spans).strip() == "【主な業務】"


def _is_role(block: Block) -> bool:
    if block.kind != "heading" or block.level < 2:
        return False
    content = _text(block.spans).strip()
    return bool(content and not content.startswith(("■", "【", "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩")))


def _write_cell_job(cell, entries: list[Block], list_ids: dict[int, int], doc) -> None:
    source = cell.paragraphs[0]._p
    body_rpr = _first_rpr(source)
    _fill_cell(cell, ())
    # The first description line retains the original template cell spacing.
    first = True
    tasks = False
    for block in entries:
        if block.kind == "table":
            # Complex nested tables are kept outside job rows by the caller.
            continue
        if _in_tasks_heading(block):
            tasks = True
        elif block.kind != "list":
            tasks = False
        if first:
            para = cell.paragraphs[0]
            first = False
        else:
            element = deepcopy(source)
            _clean_paragraph(element)
            cell._tc.append(element)
            para = Paragraph(element, cell)
        spans = block.spans
        if block.kind == "list" and tasks and not block.ordered:
            raw = _text(spans)
            spans = (Span("・" if not raw.startswith("・") else ""),) + spans
            # This is literal typography, never a native List Bullet.
        elif block.kind == "list":
            _native_list(para, block, list_ids, doc)
        elif _in_tasks_heading(block):
            para.paragraph_format.keep_with_next = True
        _append_spans(para, spans, body_rpr)
    if not entries:
        # Never fabricate résumé text to fill an empty description.
        _fill_cell(cell, ())


def _draw_roles(doc, prototypes, roles: list[tuple[Block, list[Block]]], list_ids) -> None:
    if not roles:
        return
    table = _new_table(doc, prototypes, "ROLE", header=True, count=len(roles))
    _fill_cell(table.cell(0, 0), (Span("職種・期間"),))
    _fill_cell(table.cell(0, 1), (Span("身につけた能力／経験・主な業務"),))
    for index, (role, items) in enumerate(roles, 1):
        row = table.rows[index]
        _fill_cell(row.cells[0], role.spans)
        _write_cell_job(row.cells[1], items, list_ids, doc)
        length = len(_text(role.spans)) + sum(len(_text(item.spans)) for item in items)
        # Split long descriptions naturally instead of moving the whole row.
        _row_flag(row, "cantSplit", length < 180)


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
    normal = doc.styles["Normal"]
    default_fonts = normal.element.get_or_add_rPr().find(qn("w:rFonts"))
    for cell in doc.tables[-1]._cells:
        for para in cell.paragraphs:
            for run in para.runs:
                rpr = run._r.get_or_add_rPr()
                fonts = rpr.find(qn("w:rFonts"))
                if fonts is None:
                    fonts = OxmlElement("w:rFonts")
                    rpr.insert(0, fonts)
                for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                    specified = default_fonts.get(qn("w:" + attr)) if default_fonts is not None else None
                    if specified:
                        fonts.set(qn("w:" + attr), specified)
                    else:
                        fonts.attrib.pop(qn("w:" + attr), None)


def _draw_block(doc, paragraphs, tables, block: Block, list_ids, *, tasks=False):
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
    key = keys.get(block.kind, "BODY")
    spans = block.spans
    if block.kind == "list" and tasks and not block.ordered:
        raw = _text(spans)
        spans = (Span("・" if not raw.startswith("・") else ""),) + spans
    p = _new_paragraph(doc, paragraphs, key, spans)
    if block.kind == "list" and not (tasks and not block.ordered):
        _native_list(p, block, list_ids, doc)
    if block.kind == "heading" or _in_tasks_heading(block):
        p.paragraph_format.keep_with_next = True
    if block.kind in ("meta_date", "meta_name", "meta_end"):
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if block.kind == "meta_end":
        p.paragraph_format.keep_with_next = False
    if block.kind == "rule":
        from .docx_renderer import _rule
        _rule(p)


def _work_after_company(doc, paragraphs, tables, blocks: tuple[Block, ...], list_ids) -> None:
    roles: list[tuple[Block, list[Block]]] = []
    tasks = False

    def flush() -> None:
        if roles:
            _draw_roles(doc, tables, roles, list_ids)
            roles.clear()

    for block in blocks:
        if _is_role(block):
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


def render_resume(parsed: ParsedDocument):
    """Build a résumé from Markdown while cloning original DOCX formatting.

    A missing/unapproved template is always an error, never a quiet fallback.
    """
    doc, paragraphs, tables = _load_template()
    list_ids: dict[int, int] = {}
    in_history = False
    tasks = False
    blocks = parsed.blocks
    i = 0
    while i < len(blocks):
        block = blocks[i]
        if block.kind == "heading" and block.level == 1:
            in_history = _text(block.spans).strip().removeprefix("■").strip() == "職務経歴"
            tasks = False
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
        _draw_block(doc, paragraphs, tables, block, list_ids, tasks=tasks)
        i += 1
    return doc
