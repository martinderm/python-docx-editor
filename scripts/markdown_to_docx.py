#!/usr/bin/env python3
import argparse
import re
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

HEADING_RE = re.compile(r'^(#{1,6})\s+(.*)$')
BULLET_RE = re.compile(r'^\s*[-*]\s+(.*)$')
NUMBER_RE = re.compile(r'^\s*\d+\.\s+(.*)$')


def clean_inline(text: str) -> str:
    text = text.replace('**', '')
    text = text.replace('__', '')
    text = text.replace('`', '')
    text = re.sub(r'\[(.*?)\]\((.*?)\)', r'\1', text)
    return text.strip()


def is_table_line(line: str) -> bool:
    s = line.strip()
    return s.startswith('|') and s.endswith('|') and s.count('|') >= 2


def parse_table(lines, start):
    table_lines = []
    i = start
    while i < len(lines) and is_table_line(lines[i]):
        table_lines.append(lines[i].strip())
        i += 1
    if len(table_lines) < 2:
        return None, start + 1
    rows = []
    for idx, line in enumerate(table_lines):
        cells = [clean_inline(c) for c in line.strip('|').split('|')]
        if idx == 1 and all(re.fullmatch(r'\s*:?-{2,}:?\s*', c) for c in cells):
            continue
        rows.append(cells)
    if not rows:
        return None, i
    width = max(len(r) for r in rows)
    rows = [r + [''] * (width - len(r)) for r in rows]
    return rows, i


def add_runs(paragraph, text):
    paragraph.add_run(clean_inline(text))


def set_base_styles(doc):
    styles = doc.styles
    styles['Normal'].font.name = 'Aptos'
    styles['Normal'].font.size = Pt(11)
    for name, size in [('Title', 18), ('Heading 1', 15), ('Heading 2', 13), ('Heading 3', 12)]:
        if name in styles:
            styles[name].font.name = 'Aptos'
            styles[name].font.size = Pt(size)


def convert(md_path: Path, docx_path: Path, title: str | None = None):
    lines = md_path.read_text(encoding='utf-8').splitlines()
    doc = Document()
    set_base_styles(doc)

    first_heading_used = False
    i = 0
    while i < len(lines):
        raw = lines[i].rstrip('\n')
        line = raw.rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        table, next_i = parse_table(lines, i)
        if table:
            t = doc.add_table(rows=len(table), cols=len(table[0]))
            t.style = 'Table Grid'
            for r_idx, row in enumerate(table):
                for c_idx, cell_text in enumerate(row):
                    t.cell(r_idx, c_idx).text = cell_text
            i = next_i
            continue

        m = HEADING_RE.match(stripped)
        if m:
            level = len(m.group(1))
            text = clean_inline(m.group(2))
            if level == 1 and not first_heading_used:
                p = doc.add_paragraph(style='Title')
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                add_runs(p, title or text)
                first_heading_used = True
            else:
                doc.add_paragraph(clean_inline(text), style=f'Heading {min(level, 3)}')
            i += 1
            continue

        m = BULLET_RE.match(line)
        if m:
            p = doc.add_paragraph(style='List Bullet')
            add_runs(p, m.group(1))
            i += 1
            continue

        m = NUMBER_RE.match(line)
        if m:
            p = doc.add_paragraph(style='List Number')
            add_runs(p, m.group(1))
            i += 1
            continue

        if stripped.startswith('---') and set(stripped) == {'-'}:
            i += 1
            continue

        p = doc.add_paragraph(style='Normal')
        add_runs(p, stripped)
        i += 1

    docx_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(docx_path))


def main():
    ap = argparse.ArgumentParser(description='Convert a simple Markdown document to DOCX using python-docx.')
    ap.add_argument('--in', dest='input_path', required=True, help='Input markdown file')
    ap.add_argument('--out', dest='output_path', required=True, help='Output docx file')
    ap.add_argument('--title', help='Optional title override')
    args = ap.parse_args()
    convert(Path(args.input_path), Path(args.output_path), args.title)


if __name__ == '__main__':
    main()
