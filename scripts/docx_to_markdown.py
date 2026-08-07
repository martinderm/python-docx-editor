#!/usr/bin/env python3
"""Convert a DOCX file to a clean, readable Markdown file.

This tool preserves run formatting (bold, italic), hyperlinks, and converts tables.
Layout tables (with only 1 column) are unwrapped into normal paragraphs,
while data tables (2+ columns) are preserved as markdown tables.
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


def iter_block_items(parent: DocumentObject):
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def get_node_formatting(r_elm) -> tuple[str, bool, bool]:
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

    t_elms = r_elm.findall(qn("w:t"))
    text = "".join(t.text for t in t_elms if t.text)
    return text, bold, italic


def paragraph_to_markdown(p: Paragraph) -> str:
    raw_text = "".join(p._p.itertext()).strip()
    if not raw_text:
        return ""

    p_style = p.style.name if p.style else ""
    is_bullet = raw_text.startswith("\u2022") or raw_text.startswith("•") or "bullet" in p_style.lower()

    run_texts = []
    for child in p._p:
        if child.tag.endswith("r"):
            text, bold, italic = get_node_formatting(child)
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

        elif child.tag.endswith("hyperlink"):
            r_id = child.get(qn("r:id"))
            url = ""
            if r_id and r_id in p.part.rels:
                url = p.part.rels[r_id].target_ref

            link_runs = []
            for r_child in child.findall(qn("w:r")):
                text, bold, italic = get_node_formatting(r_child)
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

    if is_bullet:
        if para_md.startswith("\u2022") or para_md.startswith("•"):
            para_md = para_md[1:].strip()
        return f"- {para_md}"

    if p_style.startswith("Heading 1"):
        return f"# {para_md}"
    if p_style.startswith("Heading 2"):
        return f"## {para_md}"
    if p_style.startswith("Heading 3"):
        return f"### {para_md}"

    return para_md


def cell_to_markdown(cell) -> str:
    md_paras = []
    for p in cell.paragraphs:
        md = paragraph_to_markdown(p)
        if md:
            md_paras.append(md)
    return "\n\n".join(md_paras)


def table_to_markdown(table: Table) -> str:
    max_cols = 0
    if table.rows:
        max_cols = len(table.rows[0].cells)

    if max_cols <= 1:
        paras = []
        for row in table.rows:
            for cell in row.cells:
                cell_md = cell_to_markdown(cell)
                if cell_md:
                    paras.append(cell_md)
        return "\n\n".join(paras)

    md_rows = []
    for r_idx, row in enumerate(table.rows):
        cells_md = []
        for cell in row.cells:
            cell_md = cell_to_markdown(cell).replace("\n", "<br>")
            cells_md.append(cell_md)
        md_rows.append("| " + " | ".join(cells_md) + " |")

        if r_idx == 0:
            md_rows.append("|" + "---|" * len(row.cells))

    return "\n".join(md_rows)


def docx_to_markdown(doc: DocumentObject) -> str:
    blocks = []
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            md = paragraph_to_markdown(item)
            if md:
                blocks.append(md)
        elif isinstance(item, Table):
            md = table_to_markdown(item)
            if md:
                blocks.append(md)
    return "\n\n".join(blocks)


def main():
    parser = argparse.ArgumentParser(description="Convert DOCX to Markdown (preserves bold/italic/hyperlinks/tables)")
    parser.add_argument("--in", required=True, type=existing_docx, help="Input DOCX file path", dest="infile")
    parser.add_argument("--out", required=True, help="Output Markdown file path", dest="outfile")
    args = parser.parse_args()

    doc = Document(args.infile)
    md_text = docx_to_markdown(doc)

    out_path = Path(args.outfile)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(md_text)

    print(f"Successfully converted DOCX to Markdown:\n- Input: {args.infile}\n- Output: {args.outfile}")


if __name__ == "__main__":
    main()
