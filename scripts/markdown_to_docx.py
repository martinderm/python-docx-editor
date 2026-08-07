#!/usr/bin/env python3
"""Convert Markdown to DOCX with rich formatting and modern design defaults.

Supports:
- Headings with keep_with_next
- GitHub Alerts (> [!NOTE], > [!TIP], > [!IMPORTANT], > [!WARNING], > [!CAUTION])
- Fenced code blocks (``` [lang])
- Clickable Word hyperlinks ([text](url))
- Rich tables with headers, zebra striping, cell padding, and inline styling
- Nested bullet & numbered lists
- Horizontal rules
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph

ALERT_CONFIGS = {
    "NOTE": {
        "title": "ℹ️ NOTE",
        "border_color": "2563EB",
        "bg_color": "F0F6FF",
        "text_color": RGBColor(0x1E, 0x3A, 0x8A),
    },
    "TIP": {
        "title": "💡 TIP",
        "border_color": "16A34A",
        "bg_color": "F0FDF4",
        "text_color": RGBColor(0x14, 0x53, 0x2D),
    },
    "IMPORTANT": {
        "title": "🔔 IMPORTANT",
        "border_color": "7C3AED",
        "bg_color": "F5F3FF",
        "text_color": RGBColor(0x4C, 0x1D, 0x95),
    },
    "WARNING": {
        "title": "⚠️ WARNING",
        "border_color": "D97706",
        "bg_color": "FFFBEB",
        "text_color": RGBColor(0x78, 0x35, 0x0F),
    },
    "CAUTION": {
        "title": "🚫 CAUTION",
        "border_color": "DC2626",
        "bg_color": "FEF2F2",
        "text_color": RGBColor(0x7F, 0x1D, 0x1D),
    },
}

INLINE_RE = re.compile(
    r"(\[\s*[^\]]+\s*\]\([^\)]+\)|\*\*\*.*?\*\*\*|\*\*.*?\*\*|`[^`]+`|\*[^*]+\*)"
)
TABLE_ALIGN_RE = re.compile(r"^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?$")


def existing_markdown(path_str: str) -> Path:
    path = Path(path_str)
    if not path.exists():
        raise argparse.ArgumentTypeError(f"Datei nicht gefunden: {path_str}")
    if path.suffix.lower() not in {".md", ".markdown", ".txt"}:
        raise argparse.ArgumentTypeError(f"Keine Markdown-Datei: {path_str}")
    return path


def configure_document(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(0x24, 0x29, 0x2E)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.15

    title = doc.styles["Title"]
    title.font.name = "Calibri"
    title.font.size = Pt(20)
    title.font.bold = True
    title.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(10)
    title.paragraph_format.line_spacing = 1.1
    title.paragraph_format.keep_with_next = True

    headings = [
        ("Heading 1", Pt(14.5), Pt(14), Pt(4), RGBColor(0x1F, 0x4E, 0x79)),
        ("Heading 2", Pt(12.5), Pt(10), Pt(3), RGBColor(0x2E, 0x75, 0xB6)),
        ("Heading 3", Pt(11), Pt(8), Pt(2), RGBColor(0x33, 0x33, 0x33)),
    ]
    for name, size, before, after, color in headings:
        if name in doc.styles:
            style = doc.styles[name]
            style.font.name = "Calibri"
            style.font.size = size
            style.font.bold = True
            style.font.color.rgb = color
            style.paragraph_format.space_before = before
            style.paragraph_format.space_after = after
            style.paragraph_format.line_spacing = 1.1
            style.paragraph_format.keep_with_next = True

    for style_name in ("List Bullet", "List Number"):
        if style_name in doc.styles:
            style = doc.styles[style_name]
            style.font.name = "Calibri"
            style.font.size = Pt(10.5)
            style.font.color.rgb = RGBColor(0x24, 0x29, 0x2E)
            style.paragraph_format.space_before = Pt(0)
            style.paragraph_format.space_after = Pt(2)
            style.paragraph_format.line_spacing = 1.15

    if "Markdown Quote" not in doc.styles:
        style = doc.styles.add_style("Markdown Quote", WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = doc.styles["Normal"]
        style.font.italic = True
        style.paragraph_format.left_indent = Cm(0.8)
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(3)


def add_hyperlink(
    paragraph: Paragraph,
    url: str,
    text: str,
    color: str = "0563C1",
    underline: bool = True,
    is_bold: bool = False,
    is_italic: bool = False,
) -> None:
    """Adds a clickable native Word hyperlink to a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    if is_bold:
        b = OxmlElement('w:b')
        rPr.append(b)

    if is_italic:
        i = OxmlElement('w:i')
        rPr.append(i)

    if color:
        c = OxmlElement('w:color')
        c.set(qn('w:val'), color)
        rPr.append(c)

    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

    new_run.append(rPr)
    text_el = OxmlElement('w:t')
    text_el.text = text
    new_run.append(text_el)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_inline_runs(
    paragraph: Paragraph,
    text: str,
    default_font_color: RGBColor | None = None,
    is_bold: bool = False,
    is_italic: bool = False,
) -> None:
    parts = INLINE_RE.split(text)
    for part in parts:
        if not part:
            continue
        if part.startswith("***") and part.endswith("***") and len(part) > 6:
            add_inline_runs(
                paragraph,
                part[3:-3],
                default_font_color=default_font_color,
                is_bold=True,
                is_italic=True,
            )
        elif part.startswith("**") and part.endswith("**") and len(part) > 4:
            add_inline_runs(
                paragraph,
                part[2:-2],
                default_font_color=default_font_color,
                is_bold=True,
                is_italic=is_italic,
            )
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            add_inline_runs(
                paragraph,
                part[1:-1],
                default_font_color=default_font_color,
                is_bold=is_bold,
                is_italic=True,
            )
        elif part.startswith("[") and "](" in part and part.endswith(")"):
            m = re.match(r"^\[(.*?)\]\((.*?)\)$", part)
            if m:
                link_text, url = m.groups()
                clean_link_text = link_text
                link_bold = is_bold
                link_italic = is_italic
                if clean_link_text.startswith("**") and clean_link_text.endswith("**") and len(clean_link_text) > 4:
                    clean_link_text = clean_link_text[2:-2]
                    link_bold = True
                if clean_link_text.startswith("*") and clean_link_text.endswith("*") and len(clean_link_text) > 2:
                    clean_link_text = clean_link_text[1:-1]
                    link_italic = True
                add_hyperlink(
                    paragraph,
                    url,
                    clean_link_text,
                    is_bold=link_bold,
                    is_italic=link_italic,
                )
            else:
                run = paragraph.add_run(part)
                run.bold = is_bold
                run.italic = is_italic
                if default_font_color:
                    run.font.color.rgb = default_font_color
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            run.bold = is_bold
            run.italic = is_italic
            run.font.name = "Consolas"
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        else:
            run = paragraph.add_run(part)
            run.bold = is_bold
            run.italic = is_italic
            if default_font_color:
                run.font.color.rgb = default_font_color


def add_inline_runs_multiline(paragraph: Paragraph, lines: list[str], default_font_color: RGBColor | None = None) -> None:
    for idx, line in enumerate(lines):
        if idx > 0:
            paragraph.add_run().add_break(WD_BREAK.LINE)
        add_inline_runs(paragraph, line, default_font_color=default_font_color)


def set_paragraph_numbering(paragraph: Paragraph, num_id: int, ilvl: int = 0) -> None:
    ppr = paragraph._p.get_or_add_pPr()  # pylint: disable=protected-access
    if ppr.numPr is not None:
        ppr.remove(ppr.numPr)

    num_pr = OxmlElement("w:numPr")
    ilvl_el = OxmlElement("w:ilvl")
    ilvl_el.set(qn("w:val"), str(ilvl))
    num_id_el = OxmlElement("w:numId")
    num_id_el.set(qn("w:val"), str(num_id))
    num_pr.append(ilvl_el)
    num_pr.append(num_id_el)
    ppr.append(num_pr)


def get_abstract_num_id_for_num_id(anchor: Paragraph, num_id: int) -> int | None:
    try:
        root = anchor.part.numbering_part._element  # pylint: disable=protected-access
        for num in root.findall(qn("w:num")):
            num_id_raw = num.get(qn("w:numId"))
            if not num_id_raw or int(num_id_raw) != int(num_id):
                continue
            abs_ref = num.find(qn("w:abstractNumId"))
            if abs_ref is None:
                return None
            abs_id_raw = abs_ref.get(qn("w:val"))
            return int(abs_id_raw) if abs_id_raw is not None else None
    except Exception:
        return None
    return None


def create_numbering_instance(anchor: Paragraph, abstract_num_id: int, *, restart_at_one: bool = False) -> int | None:
    try:
        root = anchor.part.numbering_part._element  # pylint: disable=protected-access
        max_num_id = 0
        for num in root.findall(qn("w:num")):
            raw = num.get(qn("w:numId"))
            if not raw:
                continue
            max_num_id = max(max_num_id, int(raw))

        new_num_id = max_num_id + 1
        num_el = OxmlElement("w:num")
        num_el.set(qn("w:numId"), str(new_num_id))
        abs_ref = OxmlElement("w:abstractNumId")
        abs_ref.set(qn("w:val"), str(int(abstract_num_id)))
        num_el.append(abs_ref)

        if restart_at_one:
            lvl_override = OxmlElement("w:lvlOverride")
            lvl_override.set(qn("w:ilvl"), "0")
            start_override = OxmlElement("w:startOverride")
            start_override.set(qn("w:val"), "1")
            lvl_override.append(start_override)
            num_el.append(lvl_override)

        root.append(num_el)
        return new_num_id
    except Exception:
        return None


def find_numbering_for_kind(anchor: Paragraph, kind: str) -> tuple[int, int] | None:
    wanted = "bullet" if kind == "ul" else "decimal"
    try:
        root = anchor.part.numbering_part._element  # pylint: disable=protected-access
        abstract_meta: dict[str, tuple[str | None, str | None]] = {}
        for abstract in root.findall(qn("w:abstractNum")):
            abs_id = abstract.get(qn("w:abstractNumId"))
            if not abs_id:
                continue
            lvl0 = None
            for lvl in abstract.findall(qn("w:lvl")):
                ilvl = lvl.get(qn("w:ilvl"))
                if ilvl in {"0", None}:
                    lvl0 = lvl
                    break
            if lvl0 is None:
                continue
            num_fmt_el = lvl0.find(qn("w:numFmt"))
            num_fmt_val = num_fmt_el.get(qn("w:val")) if num_fmt_el is not None else None
            lvl_text_el = lvl0.find(qn("w:lvlText"))
            lvl_text_val = lvl_text_el.get(qn("w:val")) if lvl_text_el is not None else None
            abstract_meta[abs_id] = (num_fmt_val, lvl_text_val)

        candidates: list[tuple[int, str, str | None]] = []
        for num in root.findall(qn("w:num")):
            num_id_raw = num.get(qn("w:numId"))
            abs_ref = num.find(qn("w:abstractNumId"))
            if not num_id_raw or abs_ref is None:
                continue
            abs_id = abs_ref.get(qn("w:val"))
            if not abs_id:
                continue
            fmt, lvl_text = abstract_meta.get(abs_id, (None, None))
            if fmt == wanted:
                candidates.append((int(num_id_raw), abs_id, lvl_text))

        if not candidates:
            return None
        if kind == "ul":
            visible = [c for c in candidates if c[2] is not None and c[2].strip() != ""]
            if visible:
                return (visible[0][0], 0)
        return (candidates[0][0], 0)
    except Exception:
        return None


def fresh_numbering_for_kind(anchor: Paragraph, kind: str) -> tuple[int, int] | None:
    base = find_numbering_for_kind(anchor, kind)
    if base is None:
        return None
    abs_id = get_abstract_num_id_for_num_id(anchor, base[0])
    if abs_id is None:
        return None
    new_num_id = create_numbering_instance(anchor, abs_id, restart_at_one=(kind == "ol"))
    if new_num_id is None:
        return None
    return (new_num_id, 0)


def add_callout_box(doc: Document, lines: list[str]) -> None:
    if not lines:
        return

    first = lines[0].strip()
    alert_type = None
    header_extra = ""
    start_idx = 0

    m = re.match(r"^>\s*\[!([A-Za-z]+)\]\s*(.*)$", first)
    if m:
        alert_type = m.group(1).upper()
        header_extra = m.group(2).strip()
        start_idx = 1

    config = ALERT_CONFIGS.get(alert_type, {
        "title": "",
        "border_color": "6C757D",
        "bg_color": "F8F9FA",
        "text_color": RGBColor(0x33, 0x33, 0x33),
    })

    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.autofit = False

    cell = table.cell(0, 0)
    tcPr = cell._tc.get_or_add_tcPr()

    # Background fill
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), config["bg_color"])
    tcPr.append(shd)

    # Borders: Left thick border, others none
    tcBorders = OxmlElement("w:tcBorders")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "24")  # 3pt
    left.set(qn("w:space"), "0")
    left.set(qn("w:color"), config["border_color"])
    tcBorders.append(left)

    for side in ("top", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "none")
        tcBorders.append(b)
    tcPr.append(tcBorders)

    # Padding
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", 140), ("bottom", 140), ("left", 200), ("right", 200)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)

    cell_p = cell.paragraphs[0]
    cell_p.paragraph_format.space_before = Pt(0)
    cell_p.paragraph_format.space_after = Pt(2)

    badge_title = config["title"]
    if badge_title or header_extra:
        title_text = badge_title
        if header_extra:
            title_text = f"{badge_title} {header_extra}" if badge_title else header_extra
        run = cell_p.add_run(title_text)
        run.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = config["text_color"]
        cell_p = cell.add_paragraph()
        cell_p.paragraph_format.space_before = Pt(2)
        cell_p.paragraph_format.space_after = Pt(2)

    body_lines = []
    for line in lines[start_idx:]:
        raw_l = line.strip()
        if raw_l.startswith("> "):
            raw_l = raw_l[2:]
        elif raw_l.startswith(">"):
            raw_l = raw_l[1:]
        body_lines.append(raw_l.strip())

    current_para_lines = []
    for line in body_lines:
        if not line:
            if current_para_lines:
                add_inline_runs_multiline(cell_p, current_para_lines)
                current_para_lines = []
                cell_p = cell.add_paragraph()
                cell_p.paragraph_format.space_before = Pt(2)
                cell_p.paragraph_format.space_after = Pt(2)
        else:
            current_para_lines.append(line)

    if current_para_lines:
        add_inline_runs_multiline(cell_p, current_para_lines)

    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(4)


def add_code_block(doc: Document, code_lines: list[str], lang: str = "") -> None:
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.autofit = False

    cell = table.cell(0, 0)
    tcPr = cell._tc.get_or_add_tcPr()

    # Shading
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F6F8FA")
    tcPr.append(shd)

    # Border
    tcBorders = OxmlElement("w:tcBorders")
    for b_name in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{b_name}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "D0D7DE")
        tcBorders.append(b)
    tcPr.append(tcBorders)

    # Padding
    tcMar = OxmlElement("w:tcMar")
    for m, val in (("top", 140), ("bottom", 140), ("left", 180), ("right", 180)):
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    if lang:
        lbl = p.add_run(f"// Code block: {lang}\n")
        lbl.font.name = "Consolas"
        lbl.font.size = Pt(8.5)
        lbl.font.color.rgb = RGBColor(0x6A, 0x73, 0x7D)

    full_code = "\n".join(code_lines)
    lines = full_code.split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            p.add_run().add_break(WD_BREAK.LINE)
        run = p.add_run(line)
        run.font.name = "Consolas"
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0x24, 0x29, 0x2E)

    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(4)


def is_table_line(line: str) -> bool:
    stripped = line.strip()
    return stripped.startswith("|") and stripped.endswith("|")


def split_table_line(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def add_table(doc: Document, lines: list[str]) -> None:
    raw_rows = [split_table_line(line) for line in lines]
    if not raw_rows:
        return

    alignments = []
    data_rows = []
    for idx, row in enumerate(raw_rows):
        if idx == 1 and all(set(cell) <= set("-: ") for cell in row):
            for cell in row:
                c = cell.strip()
                if c.startswith(":") and c.endswith(":"):
                    alignments.append(WD_PARAGRAPH_ALIGNMENT.CENTER)
                elif c.endswith(":"):
                    alignments.append(WD_PARAGRAPH_ALIGNMENT.RIGHT)
                else:
                    alignments.append(WD_PARAGRAPH_ALIGNMENT.LEFT)
            continue
        if all(set(cell) <= set("-: ") for cell in row):
            continue
        data_rows.append(row)

    if not data_rows:
        return

    col_count = max(len(row) for row in data_rows)
    table = doc.add_table(rows=len(data_rows), cols=col_count)
    table.style = "Table Grid"
    table.autofit = True

    for row_idx, row in enumerate(data_rows):
        is_header = (row_idx == 0)
        trPr = table.rows[row_idx]._tr.get_or_add_trPr()
        
        trPr.append(OxmlElement("w:cantSplit"))
        if is_header:
            trPr.append(OxmlElement("w:tblHeader"))

        bg_color = "1F4E79" if is_header else ("F8FAFC" if row_idx % 2 == 1 else "FFFFFF")

        for col_idx in range(col_count):
            cell = table.cell(row_idx, col_idx)
            tcPr = cell._tc.get_or_add_tcPr()

            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), bg_color)
            tcPr.append(shd)

            tcBorders = OxmlElement("w:tcBorders")
            border_color = "0F172A" if is_header else "CBD5E1"
            for b_name in ("top", "left", "bottom", "right"):
                b = OxmlElement(f"w:{b_name}")
                b.set(qn("w:val"), "single")
                b.set(qn("w:sz"), "4")
                b.set(qn("w:space"), "0")
                b.set(qn("w:color"), border_color)
                tcBorders.append(b)
            tcPr.append(tcBorders)

            tcMar = OxmlElement("w:tcMar")
            for m, val in (("top", 120), ("bottom", 120), ("left", 160), ("right", 160)):
                node = OxmlElement(f"w:{m}")
                node.set(qn("w:w"), str(val))
                node.set(qn("w:type"), "dxa")
                tcMar.append(node)
            tcPr.append(tcMar)

            cell_text = row[col_idx] if col_idx < len(row) else ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            if col_idx < len(alignments):
                p.alignment = alignments[col_idx]

            default_color = RGBColor(0xFF, 0xFF, 0xFF) if is_header else None
            add_inline_runs(p, cell_text, default_font_color=default_color)
            if is_header:
                for run in p.runs:
                    run.bold = True

    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(4)


def add_horizontal_rule(doc: Document) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "B7B7B7")
    borders.append(bottom)
    paragraph._p.get_or_add_pPr().append(borders)  # pylint: disable=protected-access


def convert_markdown(src: Path, dst: Path) -> None:
    doc = Document()
    configure_document(doc)

    lines = src.read_text(encoding="utf-8").splitlines()
    pending_blank = False
    index = 0
    while index < len(lines):
        raw = lines[index]
        stripped = raw.strip()

        if not stripped:
            pending_blank = True
            index += 1
            continue

        if stripped.startswith("```"):
            lang = stripped[3:].strip()
            code_lines = []
            index += 1
            while index < len(lines) and not lines[index].strip().startswith("```"):
                code_lines.append(lines[index])
                index += 1
            if index < len(lines) and lines[index].strip().startswith("```"):
                index += 1
            add_code_block(doc, code_lines, lang=lang)
            pending_blank = False
            continue

        if is_table_line(raw):
            table_lines = []
            while index < len(lines) and is_table_line(lines[index]):
                table_lines.append(lines[index])
                index += 1
            add_table(doc, table_lines)
            pending_blank = False
            continue

        if stripped == "---" or stripped == "***" or stripped == "___":
            add_horizontal_rule(doc)
            pending_blank = False
            index += 1
            continue

        if stripped.startswith("# "):
            paragraph = doc.add_paragraph(style="Title")
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            add_inline_runs(paragraph, stripped[2:].strip())
            pending_blank = False
            index += 1
            continue

        if stripped.startswith("## "):
            paragraph = doc.add_paragraph(style="Heading 1")
            paragraph.paragraph_format.keep_with_next = True
            add_inline_runs(paragraph, stripped[3:].strip())
            pending_blank = False
            index += 1
            continue

        if stripped.startswith("### "):
            paragraph = doc.add_paragraph(style="Heading 2")
            paragraph.paragraph_format.keep_with_next = True
            add_inline_runs(paragraph, stripped[4:].strip())
            pending_blank = False
            index += 1
            continue

        if stripped.startswith("#### "):
            paragraph = doc.add_paragraph(style="Heading 3")
            paragraph.paragraph_format.keep_with_next = True
            add_inline_runs(paragraph, stripped[5:].strip())
            pending_blank = False
            index += 1
            continue

        if stripped.startswith("> ") or stripped == ">":
            quote_lines = []
            while index < len(lines) and (lines[index].strip().startswith("> ") or lines[index].strip() == ">"):
                quote_lines.append(lines[index])
                index += 1
            add_callout_box(doc, quote_lines)
            pending_blank = False
            continue

        number_match = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        if number_match:
            items = []
            while index < len(lines):
                current = lines[index].strip()
                current_match = re.match(r"^(\d+)\.\s+(.*)$", current)
                if not current_match:
                    break
                items.append(current_match.group(2).strip())
                index += 1
            anchor = doc.add_paragraph()
            anchor.paragraph_format.space_before = Pt(0)
            anchor.paragraph_format.space_after = Pt(0)
            numbering = fresh_numbering_for_kind(anchor, "ol") or find_numbering_for_kind(anchor, "ol")
            parent = anchor._element.getparent()
            if parent is not None:
                parent.remove(anchor._element)
            for item in items:
                paragraph = doc.add_paragraph(style="List Number")
                add_inline_runs(paragraph, item)
                if numbering is not None:
                    set_paragraph_numbering(paragraph, numbering[0], numbering[1])
            pending_blank = False
            continue

        if stripped.startswith("* ") or stripped.startswith("- "):
            items = []
            while index < len(lines):
                current = lines[index].strip()
                if not (current.startswith("* ") or current.startswith("- ")):
                    break
                items.append(current[2:].strip())
                index += 1
            for item in items:
                paragraph = doc.add_paragraph(style="List Bullet")
                add_inline_runs(paragraph, item)
            pending_blank = False
            continue

        paragraph_lines = [raw.strip()]
        index += 1
        while index < len(lines):
            next_raw = lines[index]
            next_stripped = next_raw.strip()
            if not next_stripped:
                break
            if (
                next_stripped.startswith("```")
                or next_stripped.startswith("# ")
                or next_stripped.startswith("## ")
                or next_stripped.startswith("### ")
                or next_stripped.startswith("#### ")
                or next_stripped.startswith("> ")
                or next_stripped == ">"
                or next_stripped.startswith("* ")
                or next_stripped.startswith("- ")
                or re.match(r"^(\d+)\.\s+", next_stripped)
                or next_stripped in {"---", "***", "___"}
                or is_table_line(lines[index])
            ):
                break
            paragraph_lines.append(next_raw.strip())
            index += 1

        paragraph = doc.add_paragraph()
        if pending_blank:
            paragraph.paragraph_format.space_before = Pt(4)
        add_inline_runs_multiline(paragraph, paragraph_lines)
        pending_blank = False

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert Markdown to DOCX with rich styling")
    parser.add_argument("--in", dest="infile", type=existing_markdown, required=True)
    parser.add_argument("--out", dest="outfile", type=Path, required=True)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    convert_markdown(args.infile, args.outfile)
    print(args.outfile)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(2)
