#!/usr/bin/env python3
"""Convert a DOCX file to a clean, readable Markdown file.

This tool preserves run formatting (bold, italic), hyperlinks, and converts tables.
Layout tables (with only 1 column) are unwrapped into normal paragraphs,
while data tables (2+ columns) are preserved as markdown tables.

Tracked changes are resolved instead of dropped: ``--revisions accept`` (the
default) renders the proposed text, ``--revisions reject`` renders the original
text.  An unresolvable marker is never silently discarded, because the accepted
view is what a reviewer sees in Word.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


def existing_docx(path_str: str) -> Path:
    p = Path(path_str)
    if not p.exists():
        raise argparse.ArgumentTypeError(f"Datei nicht gefunden: {path_str}")
    if p.suffix.lower() != ".docx":
        raise argparse.ArgumentTypeError(f"Keine .docx-Datei: {path_str}")
    return p


REVISION_MODES = ("accept", "reject")
INSERTION_TAGS = (qn("w:ins"), qn("w:moveTo"))
DELETION_TAGS = (qn("w:del"), qn("w:moveFrom"))
CONTAINER_TAGS = (qn("w:smartTag"), qn("w:sdt"), qn("w:sdtContent"), qn("w:fldSimple"))


def iter_block_items(parent: DocumentObject, revisions: str = "accept"):
    yield from iter_block_children(parent.element.body, parent, revisions)


def iter_block_children(container, parent: DocumentObject, revisions: str, deleted: bool = False):
    """Body-level blocks as ``(block, deleted)``, descending into wrappers."""
    for child in container.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent), deleted
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent), deleted
        elif child.tag in INSERTION_TAGS:
            if revisions == "accept":
                yield from iter_block_children(child, parent, revisions, deleted)
        elif child.tag in DELETION_TAGS:
            if revisions == "reject":
                yield from iter_block_children(child, parent, revisions, True)
        elif child.tag in CONTAINER_TAGS:
            yield from iter_block_children(child, parent, revisions, deleted)


def iter_inline_runs(node, *, revisions: str, deleted: bool = False):
    """Yield ``(element, deleted)`` for runs and hyperlinks in document order.

    Revision wrappers are resolved according to ``revisions``: insertions are
    kept and deletions dropped for ``accept``, the other way round for
    ``reject``.  ``deleted`` carries whether the element sits inside a deletion,
    so only the matching text node (``w:t`` or ``w:delText``) is read.  Unknown
    inline containers are descended into so that no text is lost silently.
    """
    for child in node:
        tag = child.tag
        if tag == qn("w:r") or tag == qn("w:hyperlink"):
            yield child, deleted
        elif tag in INSERTION_TAGS:
            if revisions == "accept":
                yield from iter_inline_runs(child, revisions=revisions, deleted=deleted)
        elif tag in DELETION_TAGS:
            if revisions == "reject":
                yield from iter_inline_runs(child, revisions=revisions, deleted=True)
        elif tag in CONTAINER_TAGS:
            yield from iter_inline_runs(child, revisions=revisions, deleted=deleted)


def run_text(r_elm, *, deleted: bool = False) -> str:
    """Text of one run in its own view: kept text or deleted text, never both.

    ``w:del`` carries ``w:delText``, ``w:moveFrom`` carries regular ``w:t``;
    therefore the deleted view prefers ``w:delText`` and falls back to ``w:t``.
    """
    if deleted:
        deleted_parts = [
            el.text for el in r_elm.iter() if el.tag == qn("w:delText") and el.text
        ]
        if deleted_parts:
            return "".join(deleted_parts)
    return "".join(el.text for el in r_elm.iter() if el.tag == qn("w:t") and el.text)


def get_node_formatting(r_elm, *, deleted: bool = False) -> tuple[str, bool, bool]:
    rPr = r_elm.find(qn("w:rPr"))
    bold = False
    italic = False
    if rPr is not None:
        b_el = rPr.find(qn("w:b"))
        if b_el is not None and b_el.get(qn("w:val")) not in ("0", "false"):
            bold = True
        i_el = rPr.find(qn("w:i"))
        if i_el is not None and i_el.get(qn("w:val")) not in ("0", "false"):
            italic = True

    return run_text(r_elm, deleted=deleted), bold, italic


def paragraph_text(p: Paragraph, revisions: str = "accept", deleted: bool = False) -> str:
    """Plain text of a paragraph in the selected revision view."""
    parts = []
    for child, child_deleted in iter_inline_runs(p._p, revisions=revisions, deleted=deleted):
        if child.tag == qn("w:hyperlink"):
            parts.extend(
                run_text(run, deleted=run_deleted)
                for run, run_deleted in iter_inline_runs(
                    child, revisions=revisions, deleted=child_deleted
                )
            )
        else:
            parts.append(run_text(child, deleted=child_deleted))
    return "".join(parts).strip()


def paragraph_to_markdown(
    p: Paragraph, revisions: str = "accept", deleted: bool = False
) -> str:
    raw_text = paragraph_text(p, revisions, deleted)
    if not raw_text:
        return ""

    p_style = p.style.name if p.style else ""
    is_bullet = raw_text.startswith("\u2022") or raw_text.startswith("•") or "bullet" in p_style.lower()
    is_number = "number" in p_style.lower()

    ilvl = 0
    numPr = p._p.xpath(".//w:numPr")
    if numPr:
        ilvl_el = numPr[0].find(qn("w:ilvl"))
        if ilvl_el is not None and ilvl_el.get(qn("w:val")):
            try:
                ilvl = int(ilvl_el.get(qn("w:val")))
            except ValueError:
                pass
    elif p.paragraph_format.left_indent is not None:
        try:
            indent_pt = p.paragraph_format.left_indent.pt
            if indent_pt > 18:
                ilvl = max(0, int((indent_pt - 18) // 18) + 1)
        except Exception:
            pass

    run_texts = []
    for child, child_deleted in iter_inline_runs(p._p, revisions=revisions, deleted=deleted):
        if child.tag == qn("w:r"):
            text, bold, italic = get_node_formatting(child, deleted=child_deleted)
            if not text:
                continue
            if is_bullet and (text.strip() in {"\u2022", "•"}):
                continue
            stripped = text.strip()
            if not stripped:
                run_texts.append(text)
                continue
            leading_ws = text[: len(text) - len(text.lstrip())]
            trailing_ws = text[len(text.rstrip()) :]
            formatted = stripped
            if bold and italic:
                formatted = f"***{formatted}***"
            elif bold:
                formatted = f"**{formatted}**"
            elif italic:
                formatted = f"*{formatted}*"
            run_texts.append(leading_ws + formatted + trailing_ws)

        elif child.tag == qn("w:hyperlink"):
            r_id = child.get(qn("r:id"))
            url = ""
            if r_id and r_id in p.part.rels:
                url = p.part.rels[r_id].target_ref

            link_runs = []
            for r_child, r_deleted in iter_inline_runs(
                child, revisions=revisions, deleted=deleted
            ):
                if r_child.tag != qn("w:r"):
                    continue
                text, bold, italic = get_node_formatting(r_child, deleted=r_deleted)
                if text:
                    link_runs.append(text)
            link_text = "".join(link_runs)
            if link_text:
                if url:
                    run_texts.append(f"[{link_text}]({url})")
                else:
                    run_texts.append(link_text)

    para_md = "".join(run_texts).strip()
    if not para_md and raw_text:
        para_md = raw_text

    indent_str = "  " * ilvl
    if is_bullet:
        if para_md.startswith("\u2022") or para_md.startswith("•"):
            para_md = para_md[1:].strip()
        return f"{indent_str}- {para_md}"

    if is_number:
        return f"{indent_str}1. {para_md}"

    if p_style.startswith("Heading 1"):
        return f"# {para_md}"
    if p_style.startswith("Heading 2"):
        return f"## {para_md}"
    if p_style.startswith("Heading 3"):
        return f"### {para_md}"

    return para_md


def cell_to_markdown(cell, revisions: str = "accept", deleted: bool = False) -> str:
    md_paras = []
    for p in cell.paragraphs:
        md = paragraph_to_markdown(p, revisions, deleted)
        if md:
            md_paras.append(md)
    return "\n\n".join(md_paras)


def table_to_markdown(
    table: Table, revisions: str = "accept", deleted: bool = False
) -> str:
    max_cols = 0
    if table.rows:
        max_cols = len(table.rows[0].cells)

    if max_cols <= 1:
        paras = []
        for row in table.rows:
            for cell in row.cells:
                cell_md = cell_to_markdown(cell, revisions, deleted)
                if cell_md:
                    paras.append(cell_md)
        return "\n\n".join(paras)

    md_rows = []
    for r_idx, row in enumerate(table.rows):
        cells_md = []
        for cell in row.cells:
            cell_md = cell_to_markdown(cell, revisions, deleted).replace("\n", "<br>")
            cells_md.append(cell_md)
        md_rows.append("| " + " | ".join(cells_md) + " |")

        if r_idx == 0:
            md_rows.append("|" + "---|" * len(row.cells))

    return "\n".join(md_rows)


def docx_to_markdown(doc: DocumentObject, revisions: str = "accept") -> str:
    blocks = []
    for item, deleted in iter_block_items(doc, revisions):
        if isinstance(item, Paragraph):
            md = paragraph_to_markdown(item, revisions, deleted)
            if md:
                blocks.append(md)
        elif isinstance(item, Table):
            md = table_to_markdown(item, revisions, deleted)
            if md:
                blocks.append(md)
    return "\n\n".join(blocks)


def main():
    parser = argparse.ArgumentParser(description="Convert DOCX to Markdown (preserves bold/italic/hyperlinks/tables)")
    parser.add_argument("--in", required=True, type=existing_docx, help="Input DOCX file path", dest="infile")
    parser.add_argument("--out", required=True, help="Output Markdown file path", dest="outfile")
    parser.add_argument(
        "--revisions",
        choices=REVISION_MODES,
        default="accept",
        help="accept = proposed text (default, what Word shows as final), reject = original text",
    )
    args = parser.parse_args()

    doc = Document(args.infile)
    md_text = docx_to_markdown(doc, args.revisions)

    out_path = Path(args.outfile)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(md_text)

    print(f"Successfully converted DOCX to Markdown:\n- Input: {args.infile}\n- Output: {args.outfile}\n- Revisions: {args.revisions}")


if __name__ == "__main__":
    main()
