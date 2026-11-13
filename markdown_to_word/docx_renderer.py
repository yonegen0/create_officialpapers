"""Editable Word rendering with restrained Japanese résumé formatting."""
from __future__ import annotations

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.enum.style import WD_STYLE_TYPE

from .models import Block, ParsedDocument, Span

_FONT = "游ゴシック"


def _font(font, size: float, bold: bool | None = None) -> None:
    font.name = _FONT
    font.size = Pt(size)
    if bold is not None:
        font.bold = bold
    font.color.rgb = RGBColor(39, 48, 57)


def _east_asian(style) -> None:
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn("w:" + key), _FONT)


def _run_east_asian(run) -> None:
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    fonts.set(qn("w:eastAsia"), _FONT)


def _style(doc, name: str, size: float, before: float, after: float, *, bold: bool = False, keep: bool = False) -> None:
    styles = doc.styles
    s = styles[name] if name in styles else styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    _font(s.font, size, bold)
    _east_asian(s)
    f = s.paragraph_format
    f.space_before, f.space_after = Pt(before), Pt(after)
    f.line_spacing = 1.15
    f.keep_with_next = keep
    f.widow_control = True


def _setup(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.top_margin = section.bottom_margin = Mm(16)
    section.left_margin = section.right_margin = Mm(18)
    section.footer_distance = Mm(8)
    _style(doc, "Normal", 10.5, 0, 4)
    _style(doc, "Title", 20, 0, 10, bold=True, keep=True)
    title_ppr = doc.styles["Title"].element.get_or_add_pPr()
    old_border = title_ppr.find(qn("w:pBdr"))
    if old_border is not None:
        title_ppr.remove(old_border)
    _style(doc, "Heading 1", 13, 10, 5, bold=True, keep=True)
    _style(doc, "Heading 2", 11.5, 8, 3, bold=True, keep=True)
    _style(doc, "Heading 3", 10.5, 6, 3, bold=True, keep=True)
    for name in ("List Bullet", "List Number", "List Bullet 2", "List Number 2", "List Bullet 3", "List Number 3"):
        if name in doc.styles:
            _style(doc, name, 10.5, 0, 2)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.style = doc.styles["Normal"]
    run = footer.add_run()
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    run._r.append(fld)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    footer.add_run()._r.append(instr)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    footer.add_run()._r.append(end)


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


def render_docx(parsed: ParsedDocument) -> Document:
    doc = Document()
    _setup(doc)
    for b in parsed.blocks:
        if b.kind == "table":
            _draw_table(doc, b)
            continue
        if b.kind == "heading":
            p = doc.add_paragraph(style=f"Heading {min(b.level, 3)}")
        elif b.kind == "meta_title":
            p = doc.add_paragraph(style="Title")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif b.kind in ("meta_date", "meta_name"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.space_after = Pt(2)
        elif b.kind == "list":
            n = max(1, min(b.level, 3))
            name = "List Number" if b.ordered else "List Bullet"
            if n > 1:
                name += f" {n}"
            p = doc.add_paragraph(style=name)
        else:
            p = doc.add_paragraph()
        if b.kind == "rule":
            _rule(p)
        else:
            _spans(p, b.spans)
        if b.kind == "code":
            for run in p.runs:
                run.font.name = "Consolas"
    # Explicitly keep the Word file editable; no document protection is added.
    return doc
