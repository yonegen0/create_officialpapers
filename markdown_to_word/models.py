from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal


@dataclass(frozen=True)
class Span:
    text: str
    bold: bool = False
    italic: bool = False
    link: str | None = None
    code: bool = False
    break_line: bool = False


@dataclass(frozen=True)
class Block:
    kind: Literal["paragraph", "heading", "meta_title", "meta_date", "meta_name", "table", "list", "rule", "code", "meta_end"]
    spans: tuple[Span, ...] = ()
    level: int = 0
    ordered: bool = False
    start: int = 1
    list_id: int = 0
    rows: tuple[tuple[tuple[Span, ...], ...], ...] = ()
    header: bool = False
    table_type: str = "standard"


@dataclass(frozen=True)
class ParsedDocument:
    blocks: tuple[Block, ...]
    warnings: tuple[str, ...] = ()
    characters: int = 0
    headings: int = 0
    tables: int = 0
    list_items: int = 0


@dataclass(frozen=True)
class Preview:
    document: ParsedDocument
    fingerprint: str
    style: str
    input_mode: str
