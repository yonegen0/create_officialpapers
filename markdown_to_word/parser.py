"""Token-based Markdown conversion plus narrowly scoped résumé conventions.

No user-supplied HTML is evaluated, and unknown markup is rendered literally.
"""
from __future__ import annotations

import re
from collections.abc import Sequence
from markdown_it import MarkdownIt
from markdown_it.token import Token

from .errors import ConversionError
from .models import Block, ParsedDocument, Span

_ROLE_PERIOD = re.compile(r"^[^（）\n]{2,40}（\d{4}年\d{1,2}月\s*[～〜－-]\s*(?:現在|\d{4}年\d{1,2}月)）$")

_SECTION = re.compile(r"^■\s*\S")
_CIRCLED = re.compile(r"^[①-⑳]\s*\S")
_BRACKET = re.compile(r"^【[^】]+】$")
_DATE = re.compile(r"^\d{4}年\d{1,2}月\d{1,2}日現在$")
_NAME = re.compile(r"^氏名[：:]\s*\S")
_KV_KEYS = ("勤務先", "在籍期間", "雇用形態", "事業内容", "会社規模")
_HEADER_KEYS = (("項目", "内容"), ("経験・能力", "内容"))
_FENCE = re.compile(r"^\s{0,3}(`{3,}|~{3,})")


def _separated(line: str) -> tuple[str, str] | None:
    if line.startswith(("    ", "\t")) or "|" in line:
        return None
    # Tabs or U+3000 delimiters are unambiguous; never use the number of ASCII spaces.
    parts = re.split(r"\t+|　+", line.strip(), maxsplit=1)
    if len(parts) != 2 or not all(part.strip() for part in parts):
        return None
    return parts[0].strip(), parts[1].strip()


def _generic_tab_row(line: str) -> tuple[str, ...] | None:
    """Only explicit, nonindented tab-separated rows can become generic tables."""
    if "\t" not in line or line.startswith((" ", "\t")):
        return None
    cells = tuple(part.strip() for part in line.split("\t"))
    if len(cells) < 2 or any(not cell for cell in cells):
        return None
    return cells


def _protected_markdown_lines(source: str) -> set[int]:
    """Avoid interpreting Markdown block structures as generic tabular data."""
    md = MarkdownIt("commonmark", {"html": False}).enable("table")
    protected: set[int] = set()
    for token in md.parse(source):
        if token.map and token.type in (
            "blockquote_open", "bullet_list_open", "ordered_list_open",
            "fence", "code_block", "html_block", "heading_open", "table_open",
        ):
            protected.update(range(*token.map))
    return protected


def _reserved_resume_tables(source: str) -> tuple[str, dict[str, Block]]:
    """Replace only recognized, contiguous two-column data with opaque sentinels.

    Sentinels are expanded after markdown-it tokenization, avoiding heuristic
    re-interpretation of unrelated text as tables.
    """
    lines = source.split("\n")
    protected = _protected_markdown_lines(source)
    out: list[str] = []
    tables: dict[str, Block] = {}
    i = 0
    fence: tuple[str, int] | None = None
    while i < len(lines):
        line = lines[i]
        match = _FENCE.match(line)
        if match:
            mark = match.group(1)
            if fence is None:
                fence = (mark[0], len(mark))
            elif mark[0] == fence[0] and len(mark) >= fence[1]:
                fence = None
            out.append(line)
            i += 1
            continue
        if fence is not None:
            out.append(line)
            i += 1
            continue
        pair = _separated(line)
        is_header = pair in _HEADER_KEYS if pair else False
        is_company = bool(pair and pair[0] in _KV_KEYS)
        if not (is_header or is_company):
            first = _generic_tab_row(line) if i not in protected else None
            if first is not None:
                group_rows: list[tuple[str, ...]] = [first]
                j = i + 1
                while j < len(lines) and j not in protected:
                    row = _generic_tab_row(lines[j])
                    if row is None or len(row) != len(first):
                        break
                    group_rows.append(row)
                    j += 1
                if len(group_rows) >= 2:
                    marker = f"MTW_RESUME_TABLE_MARKER_{len(tables)}_END"
                    while marker in source:
                        marker += "X"
                    rows = tuple(tuple((Span(cell),) for cell in row) for row in group_rows)
                    tables[marker] = Block(kind="table", rows=rows, table_type="standard")
                    out.extend(("", marker, ""))
                    i = j
                    continue
            out.append(line)
            i += 1
            continue
        group: list[tuple[str, str]] = []
        j = i
        header = False
        if is_header:
            group.append(pair)  # type: ignore[arg-type]
            header = True
            j += 1
        seen_keys: set[str] = set()
        while j < len(lines):
            next_index = j
            # One blank separator is allowed between structured rows; a
            # repeated '勤務先' starts a new company table instead of merging.
            if lines[next_index].strip() == "":
                if next_index + 1 >= len(lines) or lines[next_index + 1].strip() == "":
                    break
                next_index += 1
            item = _separated(lines[next_index])
            if item is None:
                break
            if not header and item[0] not in _KV_KEYS:
                break
            if header and group[0][0] == "項目" and item[0] not in _KV_KEYS:
                break
            if item[0] in seen_keys or item[0] in ("項目", "経験・能力"):
                break
            seen_keys.add(item[0])
            group.append(item)
            j = next_index + 1
        count = len(group) - int(header)
        if count < 2:
            # A partly recognized company/skill table can still be an explicit
            # tabular block; fall back to generic rows when it is too short.
            first = _generic_tab_row(line) if i not in protected else None
            if first is not None:
                generic_rows: list[tuple[str, ...]] = [first]
                next_line = i + 1
                while next_line < len(lines) and next_line not in protected:
                    row = _generic_tab_row(lines[next_line])
                    if row is None or len(row) != len(first):
                        break
                    generic_rows.append(row)
                    next_line += 1
                if len(generic_rows) >= 2:
                    marker = f"MTW_RESUME_TABLE_MARKER_{len(tables)}_END"
                    while marker in source:
                        marker += "X"
                    rows = tuple(tuple((Span(cell),) for cell in row) for row in generic_rows)
                    tables[marker] = Block(kind="table", rows=rows, header=is_header, table_type="standard")
                    out.extend(("", marker, ""))
                    i = next_line
                    continue
            out.append(line)
            i += 1
            continue
        marker = f"MTW_RESUME_TABLE_MARKER_{len(tables)}_END"
        while marker in source:
            marker += "X"
        kind = "skill" if group[0][0] == "経験・能力" else "company"
        rows = tuple(tuple((Span(value),) for value in row) for row in group)
        tables[marker] = Block(kind="table", rows=rows, header=header, table_type=kind)
        out.extend(("", marker, ""))
        i = j
    return "\n".join(out), tables


def _inline(children: Sequence[Token] | None) -> tuple[Span, ...]:
    result: list[Span] = []
    strong = italic = False
    link: str | None = None
    for token in children or []:
        typ = token.type
        if typ == "strong_open":
            strong = True
        elif typ == "strong_close":
            strong = False
        elif typ == "em_open":
            italic = True
        elif typ == "em_close":
            italic = False
        elif typ == "link_open":
            link = token.attrGet("href")
        elif typ == "link_close":
            link = None
        elif typ in ("softbreak", "hardbreak"):
            # Preserve intentional line boundaries, including soft breaks.
            result.append(Span("\n", break_line=True))
        elif typ == "image":
            # Embedded images are intentionally unsupported; preserve the visible
            # Markdown reference rather than silently dropping its path/title.
            target = token.attrGet("src") or ""
            title = token.attrGet("title")
            suffix = (' "' + title.replace('"', '\\"') + '"') if title is not None else ""
            result.append(Span(f"![{token.content}]({target}{suffix})", strong, italic))
        elif typ in ("text", "code_inline", "html_inline"):
            result.append(Span(token.content, strong, italic, link, typ == "code_inline"))
        elif token.content:
            result.append(Span(token.content, strong, italic, link))
    return tuple(result)


def _plain(spans: Sequence[Span]) -> str:
    return "".join(s.text for s in spans)


def parse_markdown(source: str, style: str = "職務経歴書") -> ParsedDocument:
    if style not in ("職務経歴書", "標準文書"):
        raise ConversionError("文書スタイルを選択してください。")
    if not source or not source.strip():
        raise ConversionError("Markdownが空です。入力してください。")
    source = source.lstrip("\ufeff").replace("\r\n", "\n").replace("\r", "\n")
    reserved: dict[str, Block] = {}
    preprocessed = source
    if style == "職務経歴書":
        preprocessed, reserved = _reserved_resume_tables(source)
    md = MarkdownIt("commonmark", {"html": False}).enable("table")
    tokens = md.parse(preprocessed)
    blocks: list[Block] = []
    warnings: list[str] = []
    if style == "職務経歴書":
        leftover = [_separated(line) for line in preprocessed.splitlines()]
        if any(item and (item[0] in _KV_KEYS or item in _HEADER_KEYS) for item in leftover):
            warnings.append("一部の2列データを表と判定できませんでした。原文のまま保持します。")
    list_stack: list[tuple[bool, int, int]] = []
    next_list_id = 0
    in_work_history = False
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        typ = tok.type
        if typ in ("bullet_list_open", "ordered_list_open"):
            next_list_id += 1
            list_stack.append((typ == "ordered_list_open", int(tok.attrGet("start") or "1"), next_list_id))
            i += 1
            continue
        if typ in ("bullet_list_close", "ordered_list_close"):
            list_stack.pop()
            i += 1
            continue
        if typ == "list_item_open":
            i += 1
            continue
        if typ in ("paragraph_open", "heading_open"):
            inline = tokens[i + 1] if i + 1 < len(tokens) else None
            spans = _inline(inline.children if inline and inline.type == "inline" else None)
            text = _plain(spans).strip()
            if style == "職務経歴書" and not list_stack and (
                _SECTION.match(text) or (typ == "heading_open" and int(tok.tag[1:]) == 1)
            ):
                in_work_history = bool(re.fullmatch(r"(?:■\s*)?職務経歴", text))
            if text in reserved and typ == "paragraph_open":
                blocks.append(reserved[text])
            elif list_stack:
                ordered, start, list_id = list_stack[-1]
                blocks.append(Block("list", spans, level=min(len(list_stack), 3),
                                    ordered=ordered, start=start, list_id=list_id))
            elif typ == "heading_open":
                level = min(int(tok.tag[1:]), 3)
                if style == "職務経歴書" and text == "職務経歴書" and not blocks:
                    blocks.append(Block("meta_title", spans))
                    i += 3
                    continue
                if style == "職務経歴書" and _SECTION.match(text):
                    level = 1
                blocks.append(Block("heading", spans, level=level))
            elif style == "職務経歴書" and _SECTION.match(text) and "\n" not in text:
                blocks.append(Block("heading", spans, level=1))
            elif style == "職務経歴書" and (_CIRCLED.match(text) or _BRACKET.match(text)) and "\n" not in text:
                blocks.append(Block("heading", spans, level=2))
            elif style == "職務経歴書" and in_work_history and _ROLE_PERIOD.fullmatch(text) and "\n" not in text:
                blocks.append(Block("heading", spans, level=2))
            elif style == "職務経歴書" and text == "職務経歴書" and not blocks:
                blocks.append(Block("meta_title", spans))
            elif style == "職務経歴書" and len(blocks) <= 2 and all(b.kind in ("meta_title", "meta_date", "meta_name") for b in blocks) and _DATE.fullmatch(text):
                blocks.append(Block("meta_date", spans))
            elif style == "職務経歴書" and len(blocks) <= 2 and all(b.kind in ("meta_title", "meta_date", "meta_name") for b in blocks) and _NAME.match(text):
                blocks.append(Block("meta_name", spans))
            else:
                blocks.append(Block("paragraph", spans))
            i += 3
            continue
        if typ == "table_open":
            rows: list[tuple[tuple[Span, ...], ...]] = []
            header = False
            i += 1
            current_row: list[tuple[Span, ...]] | None = None
            while i < len(tokens) and tokens[i].type != "table_close":
                current = tokens[i]
                if current.type == "thead_open":
                    header = True
                elif current.type == "tr_open":
                    current_row = []
                elif current.type == "tr_close":
                    if current_row is not None:
                        rows.append(tuple(current_row))
                    current_row = None
                elif current.type in ("th_open", "td_open"):
                    inl = tokens[i + 1] if i + 1 < len(tokens) else None
                    if current_row is not None:
                        current_row.append(_inline(inl.children if inl and inl.type == "inline" else None))
                i += 1
            table_type = "standard"
            if style == "職務経歴書" and rows and len(rows[0]) == 2:
                labels = tuple("".join(s.text for s in cell).strip() for cell in rows[0])
                if labels == ("経験・能力", "内容"):
                    table_type = "skill"
                elif labels == ("項目", "内容") or labels[0] in _KV_KEYS:
                    table_type = "company"
            blocks.append(Block("table", rows=tuple(rows), header=header, table_type=table_type))
            i += 1
            continue
        if typ == "hr":
            blocks.append(Block("rule"))
        elif typ in ("fence", "code_block"):
            blocks.append(Block("code", (Span(tok.content),)))
        elif typ == "html_block":
            warnings.append("HTMLは実行せず、文字列として出力します。")
            blocks.append(Block("code", (Span(tok.content),)))
        i += 1
    if not blocks:
        raise ConversionError("変換できる本文がありません。")
    if style == "職務経歴書" and blocks[-1].kind == "paragraph" and _plain(blocks[-1].spans).strip() == "以上":
        blocks[-1] = Block("meta_end", blocks[-1].spans)
    # Reserved table rows are replaced by sentinels, so inspect original
    # Markdown tokens to detect images inside converted tab-separated cells.
    image_tokens = md.parse(source) if preprocessed != source else tokens
    if any(child.type == "image" for token in image_tokens for child in (token.children or [])):
        warnings.append("画像は埋め込まず、Markdown記法の文字列として保持します。")
    # The public counter tracks visible content characters; never count markup.
    char_count = sum(len(s.text) for b in blocks for s in b.spans)
    char_count += sum(len(s.text) for b in blocks for row in b.rows for cell in row for s in cell)
    for b in blocks:
        if b.kind == "table" and any(len(row) != len(b.rows[0]) for row in b.rows):
            warnings.append("列数が異なる表があります。")
    if re.search(r"<[/!A-Za-z]", source):
        warnings.append("HTML記法はWordの機能として実行されません。")
    return ParsedDocument(
        blocks=tuple(blocks), warnings=tuple(dict.fromkeys(warnings)),
        characters=char_count,
        headings=sum(b.kind in ("heading", "meta_title") for b in blocks),
        tables=sum(b.kind == "table" for b in blocks),
        list_items=sum(b.kind == "list" for b in blocks),
        style=style,
    )
