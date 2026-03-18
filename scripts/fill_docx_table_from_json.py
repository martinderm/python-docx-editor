#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from docx import Document


HEADER = ["output", "risk", "factors", "mitigation", "warning"]


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
        raise argparse.ArgumentTypeError(f"JSON-Datei nicht gefunden: {path_str}")
    if p.suffix.lower() not in {".json", ".jsonl"}:
        raise argparse.ArgumentTypeError(f"Keine JSON-Datei: {path_str}")
    return p


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Fill DOCX table rows from JSON with merge-aware column handling")
    ap.add_argument("--in", dest="infile", type=existing_docx, required=True)
    ap.add_argument("--out", dest="outfile", type=Path, required=True)
    ap.add_argument("--spec", dest="specfile", type=existing_json, required=True)
    return ap.parse_args()


def load_spec(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Spec muss ein JSON-Objekt sein.")
    if "table_index" not in data or "rows" not in data:
        raise ValueError("Spec braucht mindestens 'table_index' und 'rows'.")
    if not isinstance(data["table_index"], int) or data["table_index"] < 1:
        raise ValueError("table_index muss Integer >= 1 sein.")
    if not isinstance(data["rows"], list):
        raise ValueError("rows muss eine Liste sein.")
    return data


def merged(row, a: int, b: int) -> bool:
    return row.cells[a]._tc is row.cells[b]._tc


def set_cell(cell, text: str) -> None:
    cell.text = text


def clear_row(row) -> None:
    for cell in row.cells:
        if cell.text:
            cell.text = ""


def normalize_row_payload(item: dict[str, Any]) -> dict[str, str]:
    if "values" in item:
        vals = item["values"]
        if not isinstance(vals, dict):
            raise ValueError("rows[].values muss ein Objekt sein.")
        src = vals
    else:
        src = item

    out = {}
    for key in HEADER:
        val = src.get(key, "")
        if val is None:
            val = ""
        if not isinstance(val, str):
            raise ValueError(f"rows[].{key} muss String sein.")
        out[key] = val
    return out


def fill_logical_row_merge_aware(row, values: dict[str, str]) -> dict[str, Any]:
    # Canonical logical mapping:
    # 0 output | 1 risk | 2 factors | 3 mitigation | 4 warning? | 5 warning?
    # Some templates merge 3+4, others merge 4+5.
    clear_row(row)
    set_cell(row.cells[0], values["output"])
    set_cell(row.cells[1], values["risk"])
    set_cell(row.cells[2], values["factors"])

    if len(row.cells) < 6:
        raise ValueError("Erwartet mindestens 6 sichtbare Zellen in der Tabellenzeile.")

    merge_34 = merged(row, 3, 4)
    merge_45 = merged(row, 4, 5)

    if merge_34:
        set_cell(row.cells[3], values["mitigation"])
        set_cell(row.cells[5], values["warning"])
        strategy = "merge_3_4"
    elif merge_45:
        set_cell(row.cells[3], values["mitigation"])
        set_cell(row.cells[4], values["warning"])
        strategy = "merge_4_5"
    else:
        set_cell(row.cells[3], values["mitigation"])
        set_cell(row.cells[5], values["warning"])
        strategy = "no_merge_detected"

    return {
        "merge_3_4": merge_34,
        "merge_4_5": merge_45,
        "strategy": strategy,
    }


def main() -> int:
    args = parse_args()
    spec = load_spec(args.specfile)
    doc = Document(str(args.infile))

    table_index = spec["table_index"]
    if len(doc.tables) < table_index:
        raise ValueError(f"Dokument hat nur {len(doc.tables)} Tabellen; table_index={table_index} ist ungültig.")
    table = doc.tables[table_index - 1]

    results = []
    for item in spec["rows"]:
        if not isinstance(item, dict):
            raise ValueError("Jeder rows[]-Eintrag muss Objekt sein.")
        row_index = item.get("row_index")
        if not isinstance(row_index, int) or row_index < 1:
            raise ValueError("rows[].row_index muss Integer >= 1 sein.")
        if row_index > len(table.rows):
            raise ValueError(f"row_index {row_index} außerhalb der Tabelle (rows={len(table.rows)}).")

        mode = item.get("mode", "fill")
        if mode not in {"fill", "clear"}:
            raise ValueError("rows[].mode muss 'fill' oder 'clear' sein.")

        row = table.rows[row_index - 1]
        if mode == "clear":
            clear_row(row)
            results.append({"row_index": row_index, "mode": "clear", "status": "ok"})
            continue

        values = normalize_row_payload(item)
        diag = fill_logical_row_merge_aware(row, values)
        results.append({"row_index": row_index, "mode": "fill", "status": "ok", **diag})

    args.outfile.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(args.outfile))
    print(json.dumps({
        "in": str(args.infile),
        "out": str(args.outfile),
        "table_index": table_index,
        "rows_processed": len(results),
        "results": results,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
