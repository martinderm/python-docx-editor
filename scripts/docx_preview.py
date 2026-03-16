#!/usr/bin/env python3
"""Preview helpers for DOCX rewrite workflows.

Examples:
  py -3 scripts/docx_preview.py v2-section --json Tests/file.structure.v2.json --title "3. Governance and Roles"
  py -3 scripts/docx_preview.py docx-around --in Tests/file.docx --contains "3. Governance and Roles" --lines 25
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from docx import Document


def existing_file(path_str: str) -> Path:
    p = Path(path_str)
    if not p.exists():
        raise argparse.ArgumentTypeError(f"Datei nicht gefunden: {path_str}")
    return p


def cmd_v2_section(json_path: Path, title: str) -> int:
    data = json.loads(json_path.read_text(encoding="utf-8"))
    if data.get("schema") != "docx-structure.v2":
        raise ValueError("Kein docx-structure.v2 JSON.")

    def dump(sec: dict, indent: int = 0) -> None:
        pad = "  " * indent
        print(f"{pad}{sec.get('title')} ({sec.get('block_id')})")
        for n in sec.get("content", []):
            text = " ".join((n.get("text", "") or "").split())
            print(f"{pad}- {n.get('type')} {n.get('block_id')}: {text}")
        for c in sec.get("children", []):
            dump(c, indent + 1)

    def find(nodes: list[dict]) -> dict | None:
        for sec in nodes:
            if sec.get("title") == title:
                return sec
            hit = find(sec.get("children", []))
            if hit is not None:
                return hit
        return None

    section = find(data["document"]["sections"])
    if section is None:
        raise ValueError(f"Section nicht gefunden: {title}")

    dump(section)
    return 0


def cmd_docx_around(docx_path: Path, contains: str, lines: int) -> int:
    doc = Document(str(docx_path))
    rows = []
    for p in doc.paragraphs:
        txt = " ".join((p.text or "").split())
        if txt:
            rows.append((p.style.name if p.style else None, txt))

    idx = next((i for i, (_, t) in enumerate(rows) if contains in t), None)
    if idx is None:
        raise ValueError(f"Text nicht gefunden: {contains}")

    start = max(0, idx)
    end = min(len(rows), idx + max(1, lines))
    for i in range(start, end):
        style, text = rows[i]
        print(f"{i+1:03d} | {style} | {text}")
    return 0


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="DOCX preview helpers")
    sub = p.add_subparsers(dest="cmd", required=True)

    p_v2 = sub.add_parser("v2-section", help="Dump a section from docx-structure.v2")
    p_v2.add_argument("--json", dest="json_path", type=existing_file, required=True)
    p_v2.add_argument("--title", required=True)

    p_docx = sub.add_parser("docx-around", help="Print lines around matching text in DOCX")
    p_docx.add_argument("--in", dest="infile", type=existing_file, required=True)
    p_docx.add_argument("--contains", required=True)
    p_docx.add_argument("--lines", type=int, default=25)

    return p.parse_args()


def main() -> int:
    args = parse_args()
    if args.cmd == "v2-section":
        return cmd_v2_section(args.json_path, args.title)
    if args.cmd == "docx-around":
        return cmd_docx_around(args.infile, args.contains, args.lines)
    raise ValueError(f"Unbekannter command: {args.cmd}")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(2)
