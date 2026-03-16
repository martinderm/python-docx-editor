#!/usr/bin/env python3
"""Small CLI helpers for .docx operations with python-docx.

Usage examples:
  py -3 scripts/docx_ops.py text --in input.docx
  py -3 scripts/docx_ops.py stats --in input.docx
  py -3 scripts/docx_ops.py replace --in in.docx --out out.docx --find "Alt" --replace "Neu"
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document


def cmd_text(infile: Path) -> None:
    doc = Document(str(infile))
    lines: list[str] = []

    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            lines.append(t)

    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                lines.append(" | ".join(cells))

    print("\n".join(lines))


def cmd_stats(infile: Path) -> None:
    doc = Document(str(infile))

    paragraph_count = len(doc.paragraphs)
    table_count = len(doc.tables)
    image_count = 0

    for rel in doc.part._rels.values():  # pylint: disable=protected-access
        if "image" in rel.reltype:
            image_count += 1

    data = {
        "file": str(infile),
        "paragraphs": paragraph_count,
        "tables": table_count,
        "images": image_count,
    }
    print(json.dumps(data, ensure_ascii=False, indent=2))


def cmd_replace(infile: Path, outfile: Path, old: str, new: str) -> None:
    doc = Document(str(infile))
    replaced = 0

    for p in doc.paragraphs:
        if old in p.text:
            p.text = p.text.replace(old, new)
            replaced += 1

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if old in cell.text:
                    cell.text = cell.text.replace(old, new)
                    replaced += 1

    doc.save(str(outfile))
    print(json.dumps({"out": str(outfile), "replaced_blocks": replaced}, ensure_ascii=False))


def existing_docx(path_str: str) -> Path:
    p = Path(path_str)
    if not p.exists():
        raise argparse.ArgumentTypeError(f"Datei nicht gefunden: {path_str}")
    if p.suffix.lower() != ".docx":
        raise argparse.ArgumentTypeError(f"Keine .docx-Datei: {path_str}")
    return p


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="DOCX Ops via python-docx")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_text = sub.add_parser("text", help="Text aus DOCX extrahieren")
    p_text.add_argument("--in", dest="infile", type=existing_docx, required=True)

    p_stats = sub.add_parser("stats", help="Basis-Statistik als JSON")
    p_stats.add_argument("--in", dest="infile", type=existing_docx, required=True)

    p_replace = sub.add_parser("replace", help="Einfache Text-Ersetzung")
    p_replace.add_argument("--in", dest="infile", type=existing_docx, required=True)
    p_replace.add_argument("--out", dest="outfile", type=Path, required=True)
    p_replace.add_argument("--find", dest="old", required=True)
    p_replace.add_argument("--replace", dest="new", required=True)

    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.cmd == "text":
        cmd_text(args.infile)
    elif args.cmd == "stats":
        cmd_stats(args.infile)
    elif args.cmd == "replace":
        cmd_replace(args.infile, args.outfile, args.old, args.new)


if __name__ == "__main__":
    main()
