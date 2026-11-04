#!/usr/bin/env python3
"""Isolated end-to-end test for the config-driven DOCX roundtrip engine."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


ROOT = Path(__file__).resolve().parents[1]
ENGINE = ROOT / "scripts" / "docx_roundtrip.py"
TEMPLATE = ROOT / "templates" / "GenericTemplate.docx"


def tracked_run(text: str, *, deleted: bool) -> OxmlElement:
    run = OxmlElement("w:r")
    node = OxmlElement("w:delText" if deleted else "w:t")
    node.text = text
    node.set(qn("xml:space"), "preserve")
    run.append(node)
    return run


def add_tracked_change(docx_path: Path, *, anchor: str, author: str = "Partner") -> None:
    """Insert one tracked insertion and one tracked deletion into a paragraph."""
    document = Document(str(docx_path))
    paragraph = next(
        item for item in document.paragraphs if item.text.strip().startswith(anchor)
    )
    insertion = OxmlElement("w:ins")
    insertion.set(qn("w:id"), "9001")
    insertion.set(qn("w:author"), author)
    insertion.set(qn("w:date"), "2026-11-04T10:00:00Z")
    insertion.append(tracked_run(" Accepted change.", deleted=False))
    deletion = OxmlElement("w:del")
    deletion.set(qn("w:id"), "9002")
    deletion.set(qn("w:author"), author)
    deletion.set(qn("w:date"), "2026-11-04T10:00:00Z")
    deletion.append(tracked_run(" Removed change.", deleted=True))
    paragraph._p.append(insertion)
    paragraph._p.append(deletion)
    document.save(str(docx_path))


def check_converter() -> None:
    """Direct checks of the revision views in ``docx_to_markdown``."""
    sys.path.insert(0, str(ROOT / "scripts"))
    from docx_to_markdown import docx_to_markdown  # noqa: PLC0415

    document = Document()
    paragraph = document.add_paragraph("Base")
    insertion = OxmlElement("w:ins")
    insertion.set(qn("w:id"), "1")
    insertion.set(qn("w:author"), "Partner")
    insertion.append(tracked_run(" NEW", deleted=False))
    deletion = OxmlElement("w:del")
    deletion.set(qn("w:id"), "2")
    deletion.set(qn("w:author"), "Partner")
    deletion.append(tracked_run(" GONE", deleted=True))
    paragraph._p.append(insertion)
    paragraph._p.append(deletion)

    link = OxmlElement("w:hyperlink")
    link_insertion = OxmlElement("w:ins")
    link_insertion.set(qn("w:id"), "5")
    link_insertion.set(qn("w:author"), "Partner")
    link_insertion.append(tracked_run(" LINK-NEW", deleted=False))
    link_deletion = OxmlElement("w:del")
    link_deletion.set(qn("w:id"), "6")
    link_deletion.set(qn("w:author"), "Partner")
    link_deletion.append(tracked_run(" LINK-GONE", deleted=True))
    link.append(link_insertion)
    link.append(link_deletion)
    paragraph._p.append(link)

    table = document.add_table(rows=1, cols=1)
    cell_paragraph = table.rows[0].cells[0].paragraphs[0]
    cell_paragraph.add_run("Cell")
    cell_insertion = OxmlElement("w:ins")
    cell_insertion.set(qn("w:id"), "3")
    cell_insertion.set(qn("w:author"), "Partner")
    cell_insertion.append(tracked_run(" NEW", deleted=False))
    cell_paragraph._p.append(cell_insertion)

    inserted_paragraph = OxmlElement("w:p")
    inserted_run = OxmlElement("w:r")
    inserted_text = OxmlElement("w:t")
    inserted_text.text = "Whole new paragraph."
    inserted_run.append(inserted_text)
    inserted_paragraph.append(inserted_run)
    block_insertion = OxmlElement("w:ins")
    block_insertion.set(qn("w:id"), "4")
    block_insertion.set(qn("w:author"), "Partner")
    block_insertion.append(inserted_paragraph)
    document.element.body.append(block_insertion)

    field_paragraph = document.add_paragraph()
    field_run = OxmlElement("w:r")
    instruction = OxmlElement("w:instrText")
    instruction.text = "PAGE"
    field_run.append(instruction)
    result = OxmlElement("w:t")
    result.text = "7"
    field_run.append(result)
    field_paragraph._p.append(field_run)

    move_paragraph = document.add_paragraph("Move")
    moved_from = OxmlElement("w:moveFrom")
    moved_from.set(qn("w:id"), "7")
    moved_from.set(qn("w:author"), "Partner")
    moved_from.append(tracked_run(" SRC", deleted=False))
    moved_to = OxmlElement("w:moveTo")
    moved_to.set(qn("w:id"), "8")
    moved_to.set(qn("w:author"), "Partner")
    moved_to.append(tracked_run(" DST", deleted=False))
    move_paragraph._p.append(moved_from)
    move_paragraph._p.append(moved_to)

    deleted_paragraph = OxmlElement("w:p")
    deleted_run = OxmlElement("w:r")
    deleted_text = OxmlElement("w:delText")
    deleted_text.text = "Whole deleted paragraph."
    deleted_run.append(deleted_text)
    deleted_paragraph.append(deleted_run)
    block_deletion = OxmlElement("w:del")
    block_deletion.set(qn("w:id"), "9")
    block_deletion.set(qn("w:author"), "Partner")
    block_deletion.append(deleted_paragraph)
    document.element.body.append(block_deletion)

    accepted = docx_to_markdown(document, "accept")
    original = docx_to_markdown(document, "reject")

    checks = {
        "accept keeps insertion": "Base NEW" in accepted,
        "accept drops deletion": "GONE" not in accepted,
        "reject keeps original": "Base GONE" in original,
        "reject drops insertion": "NEW" not in original,
        "hyperlink accept": "LINK-NEW" in accepted and "LINK-GONE" not in accepted,
        "hyperlink reject": "LINK-GONE" in original and "LINK-NEW" not in original,
        "table accept": "Cell NEW" in accepted,
        "table reject": "Cell" in original and "Cell NEW" not in original,
        "block insertion accept": "Whole new paragraph." in accepted,
        "block insertion reject": "Whole new paragraph." not in original,
        "field result only": "7" in accepted and "PAGE" not in accepted,
        "move accept": "Move DST" in accepted and "Move SRC" not in accepted,
        "move reject": "Move SRC" in original and "Move DST" not in original,
        "block deletion accept": "Whole deleted paragraph." not in accepted,
        "block deletion reject": "Whole deleted paragraph." in original,
    }
    failed = sorted(name for name, ok in checks.items() if not ok)
    if failed:
        raise AssertionError(
            "converter checks failed: " + ", ".join(failed) + f"\naccept:\n{accepted}\nreject:\n{original}"
        )


def run(arguments: list[str], workspace: Path) -> dict:
    result = subprocess.run(
        [sys.executable, str(ENGINE), *arguments],
        cwd=workspace,
        check=False,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )
    if result.returncode:
        raise AssertionError(
            f"Roundtrip command failed ({result.returncode}):\n{result.stdout}\n{result.stderr}"
        )
    return json.loads(result.stdout)


def run_expect_failure(arguments: list[str], workspace: Path) -> subprocess.CompletedProcess:
    result = subprocess.run(
        [sys.executable, str(ENGINE), *arguments],
        cwd=workspace,
        check=False,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )
    if result.returncode == 0:
        raise AssertionError(f"Roundtrip command unexpectedly passed:\n{result.stdout}")
    return result


def main() -> int:
    check_converter()
    with tempfile.TemporaryDirectory(prefix="python-docx-roundtrip-") as temp_dir:
        workspace = Path(temp_dir) / "workspace"
        definition = workspace / "memory" / "operations" / "sample"
        cloud = workspace / "data" / "cloud" / "sample"
        canonical = definition / "current.md"
        config_path = definition / "roundtrip.json"
        working = workspace / "data" / "docx-roundtrips" / "sample"
        definition.mkdir(parents=True)
        cloud.mkdir(parents=True)
        canonical.write_text(
            "---\n"
            "title: Roundtrip test\n"
            "---\n\n"
            "# Roundtrip test\n\n"
            "## Overview\n\n"
            "Initial paragraph.\n\n"
            "1. First item\n"
            "2. Second item\n\n"
            "| Key | Value |\n"
            "| --- | --- |\n"
            "| A | B |\n",
            encoding="utf-8",
        )
        config = {
            "schema": "docx-roundtrip-config.v1",
            "document_id": "sample-roundtrip",
            "canonical_markdown": "memory/operations/sample/current.md",
            "template": str(TEMPLATE),
            "working_directory": "data/docx-roundtrips/sample",
            "cloud": {
                "directory": "data/cloud/sample",
                "review_pattern": "Sample_REVIEW-{revision}.docx",
                "archive_directory": "data/cloud/sample/Archiv",
            },
            "conversion": {
                "table_orientation": "landscape",
                "bookmark_prefix": "sample",
                "preserve_frontmatter": True,
            },
            "import_policy": {
                "automatic_overwrite": False,
                "require_candidate_review": True,
                "refuse_if_canonical_changed_since_export": True,
                "high_risk_requires_explicit_override": True,
            },
        }
        config_path.write_text(
            json.dumps(config, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        base_args = ["--config", str(config_path), "--root", str(workspace)]

        validated = run([*base_args, "validate"], workspace)
        assert validated["status"] == "valid"
        exported = run(
            [*base_args, "export", "--revision", "R0001", "--confirm"],
            workspace,
        )
        assert exported["status"] == "exported"
        current_docx = cloud / "Sample_REVIEW-R0001.docx"
        assert current_docx.exists()
        assert (working / "revisions" / "R0001" / "manifest.json").exists()
        assert not (definition / "revisions").exists()

        status = run([*base_args, "status"], workspace)
        assert status["status"] == "unchanged"
        assert status["warnings"] == []

        document = Document(str(current_docx))
        document.add_paragraph("Editorial change from Word.")
        document.save(str(current_docx))
        status = run([*base_args, "status"], workspace)
        assert status["status"] == "modified"

        add_tracked_change(current_docx, anchor="Initial paragraph")
        refuse_config = dict(config)
        refuse_config["import_policy"] = dict(config["import_policy"], open_revisions="refuse")
        refuse_path = definition / "roundtrip-refuse.json"
        refuse_path.write_text(
            json.dumps(refuse_config, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        failed = run_expect_failure(
            ["--config", str(refuse_path), "--root", str(workspace), "prepare-import"],
            workspace,
        )
        assert "open_revisions=refuse" in failed.stderr, failed.stderr
        assert not (working / "revisions" / "R0001" / "candidate.md").exists()
        prepared = run([*base_args, "prepare-import"], workspace)
        assert prepared["status"] == "candidate_prepared"
        assert (working / "revisions" / "R0001" / "changes.diff").exists()
        candidate = (working / "revisions" / "R0001" / "candidate.md").read_text(
            encoding="utf-8"
        )
        assert "Initial paragraph. Accepted change." in candidate
        assert "Removed change." not in candidate
        report_path = working / "revisions" / "R0001" / "import-report.json"
        report = json.loads(report_path.read_text(encoding="utf-8"))
        assert report["review"]["open_revisions"] == 2, report["review"]
        assert report["review"]["revision_authors"] == ["Partner"], report["review"]
        assert report["risk"] == "medium", report["risk"]
        cached_refusal = run_expect_failure(
            ["--config", str(refuse_path), "--root", str(workspace), "prepare-import"],
            workspace,
        )
        assert "open_revisions=refuse" in cached_refusal.stderr, cached_refusal.stderr
        applied_refusal = run_expect_failure(
            ["--config", str(refuse_path), "--root", str(workspace), "apply-import", "--confirm"],
            workspace,
        )
        assert "open_revisions=refuse" in applied_refusal.stderr, applied_refusal.stderr
        imported = run([*base_args, "apply-import", "--confirm"], workspace)
        assert imported["status"] == "imported"
        assert canonical.read_text(encoding="utf-8").startswith(
            "---\ntitle: Roundtrip test\n---\n"
        )
        status = run([*base_args, "status"], workspace)
        assert status["status"] == "unchanged_since_import"

        advanced = run([*base_args, "advance-revision", "--confirm"], workspace)
        assert advanced["status"] == "revision_advanced"
        assert advanced["revision"] == "R0002"
        assert (cloud / "Archiv" / "Sample_REVIEW-R0001.docx").exists()
        assert (cloud / "Sample_REVIEW-R0002.docx").exists()
        assert (working / "revisions" / "R0002" / "manifest.json").exists()

        unexpected = cloud / "Sample_REV0003.docx"
        unexpected.write_bytes((cloud / "Sample_REVIEW-R0002.docx").read_bytes())
        status = run([*base_args, "status"], workspace)
        assert status["warnings"]
        assert status["inventory"]["unrecognized_docx"]

    print("roundtrip selftest: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
