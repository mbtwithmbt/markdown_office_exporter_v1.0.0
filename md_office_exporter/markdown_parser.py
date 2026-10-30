"""Small, dependency-light Markdown block and table parser.

The parser intentionally focuses on common document-export constructs. Standard
pipe tables are supported, while merged cells are represented with embedded HTML
tables using ``rowspan`` and ``colspan``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import html
import re
from typing import Any

from lxml import etree
from lxml import html as lxml_html


_GFM_SEPARATOR = re.compile(r"^\s*:?-{3,}:?\s*$")
_HEADING = re.compile(r"^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$")
_LIST_ITEM = re.compile(r"^(\s*)([-+*]|\d+[.)])\s+(.+)$")
_FENCE = re.compile(r"^\s*(`{3,}|~{3,})\s*([\w+-]*)\s*$")
_HORIZONTAL_RULE = re.compile(r"^\s{0,3}((\*\s*){3,}|(-\s*){3,}|(_\s*){3,})$")


@dataclass(slots=True)
class TableCell:
    """Anchor cell and its rectangular span."""

    row: int
    col: int
    value: str
    rowspan: int = 1
    colspan: int = 1
    is_header: bool = False


@dataclass(slots=True)
class MarkdownTable:
    """Normalized table grid used by both DOCX and XLSX exporters."""

    rows: list[list[str]]
    cells: list[TableCell]
    header_rows: set[int] = field(default_factory=set)
    title: str = ""
    source: str = "gfm"

    @property
    def row_count(self) -> int:
        return len(self.rows)

    @property
    def column_count(self) -> int:
        return max((len(row) for row in self.rows), default=0)

    def is_anchor(self, row: int, col: int) -> bool:
        return any(cell.row == row and cell.col == col for cell in self.cells)

    def anchor_at(self, row: int, col: int) -> TableCell | None:
        for cell in self.cells:
            if cell.row == row and cell.col == col:
                return cell
        return None


def _split_pipe_row(line: str) -> list[str]:
    """Split a GFM row while respecting escaped pipes and inline-code spans."""

    text = line.strip()
    if text.startswith("|"):
        text = text[1:]
    if text.endswith("|") and not text.endswith(r"\|"):
        text = text[:-1]

    cells: list[str] = []
    current: list[str] = []
    escaped = False
    in_code = False
    for char in text:
        if escaped:
            current.append(char)
            escaped = False
            continue
        if char == "\\":
            escaped = True
            continue
        if char == "`":
            in_code = not in_code
            current.append(char)
            continue
        if char == "|" and not in_code:
            cells.append("".join(current).strip())
            current = []
        else:
            current.append(char)
    if escaped:
        current.append("\\")
    cells.append("".join(current).strip())
    return cells


def _is_gfm_table(lines: list[str], index: int) -> bool:
    if index + 1 >= len(lines) or "|" not in lines[index]:
        return False
    separators = _split_pipe_row(lines[index + 1])
    return bool(separators) and all(_GFM_SEPARATOR.match(cell) for cell in separators)


def _parse_gfm_table(lines: list[str], index: int) -> tuple[MarkdownTable, int]:
    header = _split_pipe_row(lines[index])
    index += 2  # skip header separator
    data_rows: list[list[str]] = [header]
    while index < len(lines) and "|" in lines[index] and lines[index].strip():
        if _HEADING.match(lines[index]) or _FENCE.match(lines[index]):
            break
        data_rows.append(_split_pipe_row(lines[index]))
        index += 1

    width = max(len(row) for row in data_rows)
    rows = [row + [""] * (width - len(row)) for row in data_rows]
    cells = [
        TableCell(row=r, col=c, value=value, is_header=(r == 0))
        for r, row in enumerate(rows)
        for c, value in enumerate(row)
    ]
    return MarkdownTable(rows=rows, cells=cells, header_rows={0}), index


def _direct_rows(table_element: etree._Element) -> list[tuple[etree._Element, bool]]:
    rows: list[tuple[etree._Element, bool]] = []
    for child in table_element:
        tag = etree.QName(child).localname.lower()
        if tag == "tr":
            rows.append((child, False))
        elif tag in {"thead", "tbody", "tfoot"}:
            in_head = tag == "thead"
            for row in child:
                if etree.QName(row).localname.lower() == "tr":
                    rows.append((row, in_head))
    return rows


def _cell_text(cell: etree._Element) -> str:
    for br in cell.xpath(".//br"):
        br.tail = "\n" + (br.tail or "")
    text = "".join(cell.itertext())
    lines = [" ".join(part.split()) for part in text.splitlines()]
    return "\n".join(part for part in lines if part).strip()


def _positive_int(value: str | None) -> int:
    try:
        return max(1, int(value or "1"))
    except ValueError:
        return 1


def parse_html_table(source: str) -> MarkdownTable:
    """Parse one HTML table and normalize rowspan/colspan into a dense grid."""

    fragment = lxml_html.fragment_fromstring(source, create_parent="div")
    tables = fragment.xpath(".//table")
    if not tables:
        raise ValueError("HTML 片段中未找到 <table>")
    table_element = tables[0]

    occupied: set[tuple[int, int]] = set()
    anchors: list[TableCell] = []
    header_rows: set[int] = set()
    max_row = 0
    max_col = 0

    for row_index, (row_element, in_thead) in enumerate(_direct_rows(table_element)):
        col_index = 0
        direct_cells = [
            child
            for child in row_element
            if etree.QName(child).localname.lower() in {"td", "th"}
        ]
        if in_thead or (direct_cells and all(
            etree.QName(child).localname.lower() == "th" for child in direct_cells
        )):
            header_rows.add(row_index)

        for cell_element in direct_cells:
            while (row_index, col_index) in occupied:
                col_index += 1
            rowspan = _positive_int(cell_element.get("rowspan"))
            colspan = _positive_int(cell_element.get("colspan"))
            is_header = etree.QName(cell_element).localname.lower() == "th" or in_thead
            anchor = TableCell(
                row=row_index,
                col=col_index,
                value=html.unescape(_cell_text(cell_element)),
                rowspan=rowspan,
                colspan=colspan,
                is_header=is_header,
            )
            anchors.append(anchor)
            for row_offset in range(rowspan):
                for col_offset in range(colspan):
                    occupied.add((row_index + row_offset, col_index + col_offset))
            max_row = max(max_row, row_index + rowspan)
            max_col = max(max_col, col_index + colspan)
            col_index += colspan

    rows = [["" for _ in range(max_col)] for _ in range(max_row)]
    for cell in anchors:
        for row_index in range(cell.row, cell.row + cell.rowspan):
            for col_index in range(cell.col, cell.col + cell.colspan):
                rows[row_index][col_index] = cell.value

    return MarkdownTable(
        rows=rows,
        cells=anchors,
        header_rows=header_rows,
        source="html",
    )


def _html_table_at(lines: list[str], index: int) -> tuple[str, int] | None:
    if not re.search(r"<table\b", lines[index], flags=re.IGNORECASE):
        return None
    collected: list[str] = []
    while index < len(lines):
        collected.append(lines[index])
        index += 1
        if re.search(r"</table\s*>", collected[-1], flags=re.IGNORECASE):
            return "\n".join(collected), index
    raise ValueError("HTML 表格缺少 </table> 结束标签")


def _starts_block(lines: list[str], index: int) -> bool:
    line = lines[index]
    return bool(
        not line.strip()
        or _HEADING.match(line)
        or _LIST_ITEM.match(line)
        or _FENCE.match(line)
        or _HORIZONTAL_RULE.match(line)
        or line.lstrip().startswith(">")
        or re.search(r"<table\b", line, flags=re.IGNORECASE)
        or _is_gfm_table(lines, index)
    )


def parse_markdown_blocks(markdown_text: str) -> list[dict[str, Any]]:
    """Parse common Markdown blocks into a small exporter-friendly AST."""

    lines = markdown_text.splitlines()
    blocks: list[dict[str, Any]] = []
    index = 0
    last_heading = ""
    table_number = 0

    while index < len(lines):
        line = lines[index]
        if not line.strip():
            index += 1
            continue

        html_table = _html_table_at(lines, index)
        if html_table:
            source, index = html_table
            table_number += 1
            table = parse_html_table(source)
            table.title = last_heading or f"Table_{table_number}"
            blocks.append({"type": "table", "table": table})
            continue

        if _is_gfm_table(lines, index):
            table, index = _parse_gfm_table(lines, index)
            table_number += 1
            table.title = last_heading or f"Table_{table_number}"
            blocks.append({"type": "table", "table": table})
            continue

        fence_match = _FENCE.match(line)
        if fence_match:
            fence, language = fence_match.groups()
            index += 1
            content: list[str] = []
            while index < len(lines) and not re.match(
                rf"^\s*{re.escape(fence[0])}{{{len(fence)},}}\s*$", lines[index]
            ):
                content.append(lines[index])
                index += 1
            if index < len(lines):
                index += 1
            blocks.append(
                {"type": "code", "text": "\n".join(content), "language": language}
            )
            continue

        heading_match = _HEADING.match(line)
        if heading_match:
            level = len(heading_match.group(1))
            text = heading_match.group(2).strip()
            last_heading = re.sub(r"[*_`]", "", text)
            blocks.append({"type": "heading", "level": level, "text": text})
            index += 1
            continue

        if _HORIZONTAL_RULE.match(line):
            blocks.append({"type": "horizontal_rule"})
            index += 1
            continue

        if line.lstrip().startswith(">"):
            quote_lines: list[str] = []
            while index < len(lines) and lines[index].lstrip().startswith(">"):
                quote_lines.append(re.sub(r"^\s*>\s?", "", lines[index]))
                index += 1
            blocks.append({"type": "blockquote", "text": "\n".join(quote_lines)})
            continue

        list_match = _LIST_ITEM.match(line)
        if list_match:
            ordered = list_match.group(2)[0].isdigit()
            items: list[dict[str, Any]] = []
            while index < len(lines):
                item_match = _LIST_ITEM.match(lines[index])
                if not item_match or item_match.group(2)[0].isdigit() != ordered:
                    break
                spaces, _, item_text = item_match.groups()
                items.append(
                    {
                        "text": item_text,
                        "level": max(0, len(spaces.expandtabs(4)) // 2),
                    }
                )
                index += 1
            blocks.append({"type": "list", "ordered": ordered, "items": items})
            continue

        paragraph_lines = [line.strip()]
        index += 1
        while index < len(lines) and not _starts_block(lines, index):
            paragraph_lines.append(lines[index].strip())
            index += 1
        blocks.append({"type": "paragraph", "text": " ".join(paragraph_lines)})

    return blocks


def extract_markdown_tables(markdown_text: str) -> list[MarkdownTable]:
    """Return all normalized tables in source order."""

    return [
        block["table"]
        for block in parse_markdown_blocks(markdown_text)
        if block["type"] == "table"
    ]
