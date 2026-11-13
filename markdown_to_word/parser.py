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


def _reserved_resume_tables(source: str) -> tuple[str, dict[str, Block]]:
    """Replace only recognized, contiguous two-column data with opaque sentinels.

    Sentinels are expanded after markdown-it tokenization, avoiding heuristic
    re-interpretation of unrelated text as tables.
    """
    lines = source.split("\n")
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
            result.append(Span(token.content, strong, italic, link))
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
    list_stack: list[tuple[bool, int]] = []
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        typ = tok.type
        if typ in ("bullet_list_open", "ordered_list_open"):
            list_stack.append((typ == "ordered_list_open", int(tok.attrGet("start") or "1")))
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
            if text in reserved and typ == "paragraph_open":
                blocks.append(reserved[text])
            elif list_stack:
                ordered, start = list_stack[-1]
                blocks.append(Block("list", spans, level=min(len(list_stack), 3), ordered=ordered, start=start))
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
    )
