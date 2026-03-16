#!/usr/bin/env python3
"""Apply minimal, block-targeted patches to .docx files.

Designed to work with block_ids produced by extract_docx_for_llm.py.

Usage:
  py -3 scripts/apply_docx_patch.py --in in.docx --out out.docx --patch patch.json

Patch format:
{
  "ops": [
    {
      "op": "replace_text",
      "block_id": "p_12",
      "find": "Alt",
      "replace": "Neu",
      "expected_matches": 1
    }
  ]
}
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
import sys

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


@dataclass
class ReplaceResult:
    matches: int = 0
    changes: int = 0
    cross_run_conflicts: int = 0


def existing_docx(path_str: str) -> Path:
    p = Path(path_str)
    if not p.exists():
        raise argparse.ArgumentTypeError(f"Datei nicht gefunden: {path_str}")
    if p.suffix.lower() != ".docx":
        raise argparse.ArgumentTypeError(f"Keine .docx-Datei: {path_str}")
    return p


def existing_json(path_str: str) -> Path:
    p = Path(path_str)
    if not p.exists():
        raise argparse.ArgumentTypeError(f"Patch-Datei nicht gefunden: {path_str}")
    if p.suffix.lower() not in {".json", ".jsonl"}:
        raise argparse.ArgumentTypeError(f"Keine JSON-Datei: {path_str}")
    return p


def iter_block_items(parent: DocumentObject):
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def parse_table_block_id(block_id: str) -> tuple[int, int, int] | None:
    m = re.fullmatch(r"t_(\d+)_r(\d+)_(\d+)", block_id)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2)), int(m.group(3))


def replace_in_runs(paragraph: Paragraph, old: str, new: str) -> ReplaceResult:
    res = ReplaceResult()
    if not old:
        return res

    para_text = paragraph.text
    if old not in para_text:
        return res

    # Count all textual matches in paragraph.
    res.matches = para_text.count(old)

    # Minimal-safe path: replace only matches fully contained in a single run.
    for run in paragraph.runs:
        if old in run.text:
            c = run.text.count(old)
            run.text = run.text.replace(old, new)
            res.changes += c

    if res.changes < res.matches:
        # Remaining matches likely span run boundaries -> don't rewrite whole paragraph.
        res.cross_run_conflicts = res.matches - res.changes

    return res


def build_paragraph_index(doc: DocumentObject) -> dict[str, Paragraph]:
    out: dict[str, Paragraph] = {}
    p_i = 0
    h_i = 0
    f_i = 0

    # Keep block_id mapping aligned with extract_docx_for_llm.py:
    # only non-empty body paragraphs receive p_<n> ids.
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            if not item.text or not item.text.strip():
                continue
            p_i += 1
            out[f"p_{p_i}"] = item

    for section in doc.sections:
        for p in section.header.paragraphs:
            if p.text.strip():
                h_i += 1
                out[f"h_{h_i}"] = p
        for p in section.footer.paragraphs:
            if p.text.strip():
                f_i += 1
                out[f"f_{f_i}"] = p

    return out


def build_table_index(doc: DocumentObject) -> dict[int, Table]:
    out: dict[int, Table] = {}
    t_i = 0
    for item in iter_block_items(doc):
        if isinstance(item, Table):
            t_i += 1
            out[t_i] = item
    return out


def replace_in_table_range(table: Table, row_start: int, row_end: int, old: str, new: str) -> ReplaceResult:
    res = ReplaceResult()
    if not old:
        return res

    # block id rows are 1-based inclusive
    for row_idx in range(max(1, row_start), min(len(table.rows), row_end) + 1):
        row = table.rows[row_idx - 1]
        for cell in row.cells:
            for p in cell.paragraphs:
                p_res = replace_in_runs(p, old, new)
                res.matches += p_res.matches
                res.changes += p_res.changes
                res.cross_run_conflicts += p_res.cross_run_conflicts

    return res


def load_patch(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or "ops" not in data or not isinstance(data["ops"], list):
        raise ValueError("Patch-Datei muss ein JSON-Objekt mit 'ops' (Liste) sein.")
    return data


def apply_replace_op(op: dict, paragraph_index: dict[str, Paragraph], table_index: dict[int, Table]) -> dict:
    block_id = op.get("block_id")
    old = op.get("find")
    new = op.get("replace", "")
    expected = int(op.get("expected_matches", 1))

    if not isinstance(block_id, str) or not block_id:
        raise ValueError("replace_text op braucht 'block_id'.")
    if not isinstance(old, str) or old == "":
        raise ValueError("replace_text op braucht nicht-leeres 'find'.")
    if not isinstance(new, str):
        raise ValueError("replace_text op braucht String in 'replace'.")

    table_spec = parse_table_block_id(block_id)
    if table_spec is not None:
        t_idx, r_start, r_end = table_spec
        table = table_index.get(t_idx)
        if table is None:
            raise ValueError(f"Table block_id nicht gefunden: {block_id}")
        res = replace_in_table_range(table, r_start, r_end, old, new)
    else:
        paragraph = paragraph_index.get(block_id)
        if paragraph is None:
            raise ValueError(f"Paragraph/Header/Footer block_id nicht gefunden: {block_id}")
        res = replace_in_runs(paragraph, old, new)

    if res.matches != expected:
        raise ValueError(
            f"Op auf {block_id}: expected_matches={expected}, gefunden={res.matches}. "
            "Keine stillen Mehrfachtreffer erlaubt."
        )
    if res.cross_run_conflicts > 0:
        raise ValueError(
            f"Op auf {block_id}: {res.cross_run_conflicts} Treffer über Run-Grenzen. "
            "Minimal-Edit verweigert (kein Full-Rewrite)."
        )

    return {
        "op": "replace_text",
        "block_id": block_id,
        "matches": res.matches,
        "changes": res.changes,
        "status": "ok",
    }


def clear_paragraph_numbering(paragraph: Paragraph) -> None:
    ppr = paragraph._p.pPr  # pylint: disable=protected-access
    if ppr is not None and ppr.numPr is not None:
        ppr.remove(ppr.numPr)


def style_looks_like_list(style_name: str | None) -> bool:
    if not style_name:
        return False
    s = style_name.lower()
    return any(k in s for k in ["list", "bullet", "number", "aufz", "aufl"])


def apply_set_paragraph_op(op: dict, paragraph_index: dict[str, Paragraph]) -> dict:
    block_id = op.get("block_id")
    text = op.get("text")
    style = op.get("style")
    expected_contains = op.get("expected_contains")
    clear_list_format = bool(op.get("clear_list_format", True))

    if not isinstance(block_id, str) or not block_id:
        raise ValueError("set_paragraph op braucht 'block_id'.")
    if not isinstance(text, str):
        raise ValueError("set_paragraph op braucht String in 'text'.")
    if style is not None and not isinstance(style, str):
        raise ValueError("set_paragraph op: 'style' muss String sein, wenn gesetzt.")
    if expected_contains is not None and not isinstance(expected_contains, str):
        raise ValueError("set_paragraph op: 'expected_contains' muss String sein, wenn gesetzt.")

    if parse_table_block_id(block_id) is not None:
        raise ValueError("set_paragraph unterstützt keine table block_ids (t_*).")

    paragraph = paragraph_index.get(block_id)
    if paragraph is None:
        raise ValueError(f"Paragraph/Header/Footer block_id nicht gefunden: {block_id}")

    current = paragraph.text or ""
    if expected_contains is not None and expected_contains not in current:
        raise ValueError(
            f"Op auf {block_id}: expected_contains nicht gefunden. "
            "Abbruch, um falsche Zielstellen zu vermeiden."
        )

    paragraph.text = text
    if style:
        paragraph.style = style

    # Avoid leftover bullets/numbering when converting list items to prose.
    if clear_list_format and not style_looks_like_list(style):
        clear_paragraph_numbering(paragraph)

    return {
        "op": "set_paragraph",
        "block_id": block_id,
        "status": "ok",
        "style": style if style else None,
        "new_length": len(text),
    }


def apply_delete_paragraph_op(op: dict, paragraph_index: dict[str, Paragraph]) -> dict:
    block_id = op.get("block_id")
    expected_contains = op.get("expected_contains")

    if not isinstance(block_id, str) or not block_id:
        raise ValueError("delete_paragraph op braucht 'block_id'.")
    if expected_contains is not None and not isinstance(expected_contains, str):
        raise ValueError("delete_paragraph op: 'expected_contains' muss String sein, wenn gesetzt.")
    if parse_table_block_id(block_id) is not None:
        raise ValueError("delete_paragraph unterstützt keine table block_ids (t_*).")

    paragraph = paragraph_index.get(block_id)
    if paragraph is None:
        raise ValueError(f"Paragraph/Header/Footer block_id nicht gefunden: {block_id}")

    current = paragraph.text or ""
    if expected_contains is not None and expected_contains not in current:
        raise ValueError(
            f"Op auf {block_id}: expected_contains nicht gefunden. "
            "Abbruch, um falsche Zielstellen zu vermeiden."
        )

    p = paragraph._element  # pylint: disable=protected-access
    parent = p.getparent()
    if parent is None:
        raise ValueError(f"Op auf {block_id}: Paragraph hat kein Parent-Element.")
    parent.remove(p)

    return {
        "op": "delete_paragraph",
        "block_id": block_id,
        "status": "ok",
    }


def build_body_paragraph_order(doc: DocumentObject) -> list[str]:
    order: list[str] = []
    p_i = 0
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            if not item.text or not item.text.strip():
                continue
            p_i += 1
            order.append(f"p_{p_i}")
    return order


def insert_paragraph_after(paragraph: Paragraph, text: str, style: str | None = None) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)  # pylint: disable=protected-access
    new_para = Paragraph(new_p, paragraph._parent)  # pylint: disable=protected-access
    new_para.text = text
    if style:
        new_para.style = style
    return new_para


def apply_replace_paragraph_range_op(op: dict, doc: DocumentObject, paragraph_index: dict[str, Paragraph]) -> dict:
    start_block_id = op.get("start_block_id")
    end_block_id = op.get("end_block_id")
    new_paragraphs = op.get("new_paragraphs")
    expected_start_contains = op.get("expected_start_contains")
    expected_end_contains = op.get("expected_end_contains")
    allow_headings = bool(op.get("allow_headings", False))

    if not isinstance(start_block_id, str) or not start_block_id.startswith("p_"):
        raise ValueError("replace_paragraph_range braucht 'start_block_id' vom Typ p_<n>.")
    if not isinstance(end_block_id, str) or not end_block_id.startswith("p_"):
        raise ValueError("replace_paragraph_range braucht 'end_block_id' vom Typ p_<n>.")
    if not isinstance(new_paragraphs, list) or len(new_paragraphs) == 0:
        raise ValueError("replace_paragraph_range braucht nicht-leere 'new_paragraphs'-Liste.")

    start_para = paragraph_index.get(start_block_id)
    end_para = paragraph_index.get(end_block_id)
    if start_para is None or end_para is None:
        raise ValueError("replace_paragraph_range: start/end block_id nicht gefunden.")

    if expected_start_contains is not None:
        if not isinstance(expected_start_contains, str):
            raise ValueError("expected_start_contains muss String sein.")
        if expected_start_contains not in (start_para.text or ""):
            raise ValueError("replace_paragraph_range: expected_start_contains nicht gefunden.")

    if expected_end_contains is not None:
        if not isinstance(expected_end_contains, str):
            raise ValueError("expected_end_contains muss String sein.")
        if expected_end_contains not in (end_para.text or ""):
            raise ValueError("replace_paragraph_range: expected_end_contains nicht gefunden.")

    order = build_body_paragraph_order(doc)
    try:
        i_start = order.index(start_block_id)
        i_end = order.index(end_block_id)
    except ValueError as exc:
        raise ValueError("replace_paragraph_range: start/end block_id nicht in Body-Reihenfolge gefunden.") from exc
    if i_start > i_end:
        raise ValueError("replace_paragraph_range: start_block_id muss vor end_block_id liegen.")

    old_ids = order[i_start : i_end + 1]
    old_paragraphs = [paragraph_index[bid] for bid in old_ids]

    if not allow_headings:
        heading_hits = []
        for bid, p in zip(old_ids, old_paragraphs):
            style_name = p.style.name if p.style else ""
            if style_name and (style_name.startswith("Heading") or style_name.startswith("Überschrift")):
                heading_hits.append(f"{bid}({style_name})")
        if heading_hits:
            raise ValueError(
                "replace_paragraph_range enthält Heading-Absätze: "
                + ", ".join(heading_hits)
                + ". Setze allow_headings=true nur wenn das absichtlich ist."
            )

    anchor = end_para
    inserted = 0
    for entry in new_paragraphs:
        if not isinstance(entry, dict):
            raise ValueError("replace_paragraph_range: jedes Element in new_paragraphs muss Objekt sein.")
        text = entry.get("text")
        style = entry.get("style")
        if not isinstance(text, str):
            raise ValueError("replace_paragraph_range: new_paragraphs[].text muss String sein.")
        if style is not None and not isinstance(style, str):
            raise ValueError("replace_paragraph_range: new_paragraphs[].style muss String sein, wenn gesetzt.")
        anchor = insert_paragraph_after(anchor, text=text, style=style)
        # New prose paragraphs should not accidentally inherit numbering.
        if not style_looks_like_list(style):
            clear_paragraph_numbering(anchor)
        inserted += 1

    removed_preview = [" ".join((p.text or "").split())[:120] for p in old_paragraphs]

    for p in old_paragraphs:
        el = p._element  # pylint: disable=protected-access
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)

    return {
        "op": "replace_paragraph_range",
        "start_block_id": start_block_id,
        "end_block_id": end_block_id,
        "removed": len(old_paragraphs),
        "inserted": inserted,
        "removed_preview": removed_preview,
        "status": "ok",
    }


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Apply minimal block-targeted DOCX patches")
    p.add_argument("--in", dest="infile", type=existing_docx, required=True)
    p.add_argument("--out", dest="outfile", type=Path, required=True)
    p.add_argument("--patch", dest="patchfile", type=existing_json, required=True)
    return p.parse_args()


def main() -> int:
    args = parse_args()
    patch = load_patch(args.patchfile)

    doc = Document(str(args.infile))

    results: list[dict] = []
    for i, op in enumerate(patch["ops"], start=1):
        if not isinstance(op, dict):
            raise ValueError(f"Op #{i} ist kein Objekt.")

        # Rebuild indexes before each op to stay stable after insert/delete operations.
        paragraph_index = build_paragraph_index(doc)
        table_index = build_table_index(doc)

        op_kind = op.get("op")
        if op_kind == "replace_text":
            results.append(apply_replace_op(op, paragraph_index, table_index))
        elif op_kind == "set_paragraph":
            results.append(apply_set_paragraph_op(op, paragraph_index))
        elif op_kind == "delete_paragraph":
            results.append(apply_delete_paragraph_op(op, paragraph_index))
        elif op_kind == "replace_paragraph_range":
            results.append(apply_replace_paragraph_range_op(op, doc, paragraph_index))
        else:
            raise ValueError(
                f"Unbekannte op '{op_kind}' bei Op #{i}. "
                "Unterstützt: replace_text, set_paragraph, delete_paragraph, replace_paragraph_range"
            )

    args.outfile.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(args.outfile))

    print(
        json.dumps(
            {
                "in": str(args.infile),
                "out": str(args.outfile),
                "ops": len(results),
                "results": results,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(2)
