#!/usr/bin/env python3
"""Convert a DOCX file to a clean, readable Markdown file.

This tool preserves run formatting (bold, italic) and converts tables.
Layout tables (with only 1 column) are unwrapped into normal paragraphs,
while data tables (2+ columns) are preserved as markdown tables.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentObject
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

def paragraph_to_markdown(p: Paragraph) -> str:
    raw_text = p.text.strip()
    if not raw_text:
        return ""
        
    p_style = p.style.name if p.style else ""
    is_bullet = raw_text.startswith("\u2022") or raw_text.startswith("•") or "bullet" in p_style.lower()
    
    # Parse runs
    run_texts = []
    for r in p.runs:
        text = r.text
        if not text:
            continue
            
        # Skip bullet symbols to avoid duplication
        if is_bullet and (text.strip() == "\u2022" or text.strip() == "•"):
            continue
            
        stripped = text.strip()
        if not stripped:
            run_texts.append(text)
            continue
            
        # Keep leading/trailing spaces outside of format tags
        leading_ws = text[:len(text)-len(text.lstrip())]
        trailing_ws = text[len(text.rstrip()):]
        
        formatted = stripped
        if r.bold and r.italic:
            formatted = f"***{formatted}***"
        elif r.bold:
            formatted = f"**{formatted}**"
        elif r.italic:
            formatted = f"*{formatted}*"
            
        run_texts.append(leading_ws + formatted + trailing_ws)
        
    para_md = "".join(run_texts).strip()
    
    # Clean up leading bullet if present
    if is_bullet:
        if para_md.startswith("\u2022") or para_md.startswith("•"):
            para_md = para_md[1:].strip()
        return f"- {para_md}"
        
    # Check heading levels
    if p_style.startswith("Heading 1"):
        return f"# {para_md}"
    elif p_style.startswith("Heading 2"):
        return f"## {para_md}"
    elif p_style.startswith("Heading 3"):
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
    # Determine the actual number of columns
    max_cols = 0
    if table.rows:
        max_cols = len(table.rows[0].cells)
        
    if max_cols <= 1:
        # Unwrap 1-column layout tables
        paras = []
        for row in table.rows:
            for cell in row.cells:
                cell_md = cell_to_markdown(cell)
                if cell_md:
                    paras.append(cell_md)
        return "\n\n".join(paras)
        
    # Standard multi-column table
    md_rows = []
    for r_idx, row in enumerate(table.rows):
        cells_md = []
        for cell in row.cells:
            # Replace internal newlines with HTML breaks to keep table row valid
            cell_md = cell_to_markdown(cell).replace("\n", "<br>")
            cells_md.append(cell_md)
        md_rows.append("| " + " | ".join(cells_md) + " |")
        
        if r_idx == 0:
            md_rows.append("|" + "---|"*len(row.cells))
            
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
    parser = argparse.ArgumentParser(description="Convert DOCX to Markdown (preserves bold/italic/tables)")
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
