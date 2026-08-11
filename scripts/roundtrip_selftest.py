#!/usr/bin/env python3
"""Isolated end-to-end test for the config-driven DOCX roundtrip engine."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
ENGINE = ROOT / "scripts" / "docx_roundtrip.py"
TEMPLATE = ROOT / "templates" / "GenericTemplate.docx"


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


def main() -> int:
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

        prepared = run([*base_args, "prepare-import"], workspace)
        assert prepared["status"] == "candidate_prepared"
        assert (working / "revisions" / "R0001" / "changes.diff").exists()
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
