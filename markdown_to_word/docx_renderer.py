"""Editable Word rendering with restrained Japanese résumé formatting."""
from __future__ import annotations

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from .models import Block, ParsedDocument, Span

_FONT = "游ゴシック"


def _run_east_asian(run) -> None:
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    fonts.set(qn("w:eastAsia"), _FONT)


def _hyperlink(p, text: str, url: str, *, bold: bool, italic: bool) -> None:
    # Only safe web/mail destinations become clickable; other references are preserved text.
    if not url.lower().startswith(("https://", "http://", "mailto:")):
        _plain_run(p, Span(text, bold, italic))
        return
    part = p.part
    rel = part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), rel)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    if bold:
        rpr.append(OxmlElement("w:b"))
    if italic:
        rpr.append(OxmlElement("w:i"))
    clr = OxmlElement("w:color")
    clr.set(qn("w:val"), "22689A")
    rpr.append(clr)
    under = OxmlElement("w:u")
    under.set(qn("w:val"), "single")
    rpr.append(under)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    h.append(r)
    p._p.append(h)


def _plain_run(p, span: Span) -> None:
    run = p.add_run()
    run.bold = span.bold or None
    run.italic = span.italic or None
    if span.code:
        run.font.name = "Consolas"
    _run_east_asian(run)
    parts = span.text.split("\n")
    for idx, part in enumerate(parts):
        if idx:
            run.add_break()
        if part:
            run.add_text(part)


def _spans(p, spans: tuple[Span, ...]) -> None:
    for span in spans:
        if span.link:
            for idx, part in enumerate(span.text.split("\n")):
                if idx:
                    p.add_run().add_break()
                if part:
                    _hyperlink(p, part, span.link, bold=span.bold, italic=span.italic)
        else:
            _plain_run(p, span)


def _shade(cell, color: str) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color)
    tcpr.append(shd)


def _border_table(table) -> None:
    tblpr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        item = OxmlElement("w:" + edge)
        item.set(qn("w:val"), "single")
        item.set(qn("w:sz"), "4")
        item.set(qn("w:color"), "DDE3E8")
        borders.append(item)
    tblpr.append(borders)


def _cell_margin(cell, margin: int = 100) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    m = OxmlElement("w:tcMar")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement("w:" + side)
        e.set(qn("w:w"), str(margin))
        e.set(qn("w:type"), "dxa")
        m.append(e)
    tcpr.append(m)


def _cant_split(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    trpr.append(OxmlElement("w:cantSplit"))


def _repeat_header(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    trpr.append(header)


def _draw_table(doc, b: Block) -> None:
    if not b.rows:
        return
    count = max(len(row) for row in b.rows)
    table = doc.add_table(rows=len(b.rows), cols=count)
    table.autofit = False
    _border_table(table)
    usable = 210 - 18 * 2
    fractions = ([.25, .75] if b.table_type == "company" else [.30, .70] if b.table_type == "skill" else [1 / count] * count)
    if len(fractions) != count:
        fractions = [1 / count] * count
    for col_idx, fraction in enumerate(fractions):
        table.columns[col_idx].width = Mm(usable * fraction)
    for ri, row in enumerate(b.rows):
        out_row = table.rows[ri]
        row_length = sum(len(s.text) for cell in row for s in cell)
        if row_length < 180:
            _cant_split(out_row)
        if ri == 0 and b.header:
            _repeat_header(out_row)
        for ci in range(count):
            cell = out_row.cells[ci]
            cell.width = Mm(usable * fractions[ci])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _cell_margin(cell)
            para = cell.paragraphs[0]
            para.style = doc.styles["Normal"]
            para.paragraph_format.space_after = Pt(1)
            if ci < len(row):
                _spans(para, row[ci])
            if (b.table_type in ("company", "skill") and ci == 0) or (b.header and ri == 0):
                _shade(cell, "F1F4F6")
                if b.header and ri == 0:
                    for run in para.runs:
                        run.bold = True


def _rule(p) -> None:
    ppr = p._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:color"), "CED6DD")
    b.set(qn("w:sz"), "6")
    borders.append(b)
    ppr.append(borders)


def _ordered_numbering(doc, style_name: str, start: int) -> int:
    """Create a separate Word numbering instance for one Markdown list."""
    style = doc.styles[style_name]
    num_pr = style.element.pPr.numPr
    base_id = int(num_pr.numId.val)
    numbering = doc.part.numbering_part.element
    base_num = next(n for n in numbering.findall(qn("w:num"))
                    if int(n.get(qn("w:numId"))) == base_id)
    abstract_id = int(base_num.find(qn("w:abstractNumId")).get(qn("w:val")))
    num_id = max((int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))), default=0) + 1
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract = OxmlElement("w:abstractNumId")
    abstract.set(qn("w:val"), str(abstract_id))
    num.append(abstract)
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), "0")
    start_override = OxmlElement("w:startOverride")
    start_override.set(qn("w:val"), str(start))
    override.append(start_override)
    num.append(override)
    numbering.append(num)
    return num_id


def render_docx(parsed: ParsedDocument) -> Document:
    from .resume_template import render_resume

    return render_resume(parsed)
