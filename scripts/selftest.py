#!/usr/bin/env python3
"""Project self-test for python-docx-editor.

Creates a temporary DOCX, runs v2 extraction and patch writeback,
and validates key contracts without committing test artifacts.

Usage:
  py -3 scripts/selftest.py
"""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


ROOT = Path(__file__).resolve().parents[1]
EXTRACT = ROOT / "scripts" / "extract_docx_for_llm.py"
PATCH = ROOT / "scripts" / "apply_docx_patch.py"
FILL_TABLE = ROOT / "scripts" / "fill_docx_table_from_json.py"
MARKDOWN_TO_DOCX = ROOT / "scripts" / "markdown_to_docx.py"
MARKDOWN_TO_TEMPLATE_DOCX = ROOT / "scripts" / "markdown_to_template_docx.py"
GENERIC_TEMPLATE = ROOT / "templates" / "GenericTemplate.docx"


def run(cmd: list[str], cwd: Path) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, cwd=str(cwd), text=True, capture_output=True, check=True)


def body_element_text(element) -> str:
    return "".join(node.text or "" for node in element.iter(qn("w:t")))


def has_section_break(element) -> bool:
    return element.find(f".//{qn('w:sectPr')}") is not None


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="python-docx-editor-") as tmpdir:
        t = Path(tmpdir)

        in_docx = t / "in.docx"
        v2_json = t / "structure.v2.json"
        patch_json = t / "patch.json"
        out_docx = t / "out.docx"

        # Build minimal fixture document.
        d = Document()
        d.add_heading("1. Intro", level=1)
        d.add_paragraph("Hello base text")
        d.add_heading("2. Quality Principles", level=1)
        d.add_paragraph("Quality in MESHE is defined through five core principles:")
        d.add_paragraph("First item", style="List Bullet")
        d.add_paragraph("Second item", style="List Bullet")
        d.save(str(in_docx))

        # Extract v2 (default).
        run([sys.executable, str(EXTRACT), "--in", str(in_docx), "--out", str(v2_json)], ROOT)
        v2 = json.loads(v2_json.read_text(encoding="utf-8"))
        assert v2["schema"] == "docx-structure.v2", "Unexpected schema"

        # Find target paragraph block_id under section title.
        target_block = None
        for sec in v2["document"]["sections"]:
            if "Quality Principles" in sec.get("title", ""):
                for node in sec.get("content", []):
                    if node.get("type") == "paragraph" and "Quality in MESHE" in node.get("text", ""):
                        target_block = node.get("block_id")
                        break
        assert target_block, "Target block_id not found in v2 output"

        patch = {
            "ops": [
                {
                    "op": "replace_text",
                    "block_id": target_block,
                    "find": "Quality in MESHE is defined through five core principles:",
                    "replace": (
                        "Intro sentence for context. "
                        "Quality in MESHE is defined through five core principles:"
                    ),
                    "expected_matches": 1,
                }
            ]
        }
        patch_json.write_text(json.dumps(patch, ensure_ascii=False, indent=2), encoding="utf-8")

        # Apply patch.
        res = run(
            [
                sys.executable,
                str(PATCH),
                "--in",
                str(in_docx),
                "--out",
                str(out_docx),
                "--patch",
                str(patch_json),
            ],
            ROOT,
        )
        out = json.loads(res.stdout)
        assert out["results"][0]["status"] == "ok", "Patch operation did not succeed"

        d2 = Document(str(out_docx))
        texts = [" ".join(p.text.split()) for p in d2.paragraphs if p.text and p.text.strip()]
        assert any("Intro sentence for context." in t for t in texts), "Patched text missing"

        # Generic table fill test (cell-map).
        table_in = t / "table-in.docx"
        table_out = t / "table-out.docx"
        table_spec = t / "table-spec.json"

        td = Document()
        tt = td.add_table(rows=3, cols=3)
        for r in range(3):
            for c in range(3):
                tt.cell(r, c).text = f"R{r + 1}C{c + 1}"
        td.save(str(table_in))

        spec = {
            "table_index": 1,
            "layout": "cell-map",
            "cells": [
                {"row": 1, "col": 1, "text": "Top Left"},
                {"row": 2, "col": 2, "text": "Mid", "mode": "replace"},
                {"row": 3, "col": 3, "text": " + Appended", "mode": "append"},
            ],
        }
        table_spec.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")

        run(
            [
                sys.executable,
                str(FILL_TABLE),
                "--in",
                str(table_in),
                "--out",
                str(table_out),
                "--spec",
                str(table_spec),
            ],
            ROOT,
        )

        td2 = Document(str(table_out))
        tbl = td2.tables[0]
        assert tbl.rows[0].cells[0].text == "Top Left", "cell-map replace (1,1) failed"
        assert tbl.rows[1].cells[1].text == "Mid", "cell-map replace (2,2) failed"
        assert tbl.rows[2].cells[2].text == "R3C3 + Appended", "cell-map append (3,3) failed"

        # Markdown converters must suppress leading YAML front matter.
        markdown_in = t / "report.md"
        markdown_out = t / "report.docx"
        markdown_in.write_text(
            "---\n"
            "title: Internal report metadata\n"
            "project_slug: test-project\n"
            "---\n\n"
            "# Visible title\n\n"
            "Visible body text.\n",
            encoding="utf-8",
        )
        run(
            [sys.executable, str(MARKDOWN_TO_DOCX), "--in", str(markdown_in), "--out", str(markdown_out)],
            ROOT,
        )
        markdown_doc = Document(str(markdown_out))
        markdown_text = "\n".join(p.text for p in markdown_doc.paragraphs)
        assert "Visible title" in markdown_text, "Markdown title missing"
        assert "Visible body text." in markdown_text, "Markdown body missing"
        assert "project_slug:" not in markdown_text, "YAML front matter was rendered"

        # In landscape-table mode, a directly preceding heading must be added
        # after the landscape section break rather than on a portrait-only page.
        template_in = t / "landscape-report.md"
        template_out = t / "landscape-report.docx"
        template_in.write_text(
            "---\n"
            "title: Internal report metadata\n"
            "---\n\n"
            "# Visible title\n\n"
            "## Risks\n\n"
            "| Area | Risk |\n"
            "| --- | --- |\n"
            "| Delivery | Limited funding |\n\n"
            "## Actions\n\n"
            "| Owner | Action |\n"
            "| --- | --- |\n"
            "| Team | Confirm funding |\n\n"
            "## Closing\n\n"
            "Visible closing text.\n",
            encoding="utf-8",
        )
        run(
            [
                sys.executable,
                str(MARKDOWN_TO_TEMPLATE_DOCX),
                "--in",
                str(template_in),
                "--out",
                str(template_out),
                "--template",
                str(GENERIC_TEMPLATE),
                "--table-orientation",
                "landscape",
            ],
            ROOT,
        )
        template_doc = Document(str(template_out))
        template_text = "\n".join(p.text for p in template_doc.paragraphs)
        assert "Visible title" in template_text, "Template title missing"
        assert "Visible closing text." in template_text, "Template body missing"
        assert "title: Internal report metadata" not in template_text, "Template YAML front matter was rendered"
        assert len(template_doc.tables) == 2, "Expected two landscape tables"

        body_children = list(template_doc.element.body)
        for heading_text in ("Risks", "Actions"):
            heading_index = next(
                index
                for index, child in enumerate(body_children)
                if child.tag == qn("w:p") and body_element_text(child).strip() == heading_text
            )
            table_index = next(
                index
                for index in range(heading_index + 1, len(body_children))
                if body_children[index].tag == qn("w:tbl")
            )
            assert not any(
                has_section_break(child)
                for child in body_children[heading_index + 1 : table_index]
            ), f"Section break inserted between '{heading_text}' and its table"

    print("selftest: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
