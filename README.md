# python-docx-editor

A DOCX editor for AI Agents: lightweight CLI toolkit for safe, reproducible `.docx` extraction and patching with `python-docx`.

## What it does

- Extract DOCX into structured JSON (`docx-structure.v2`)
- Optionally emit RAG-friendly chunks (`docx-llm-chunks.v1`)
- Apply block-targeted patches with validation guards
- Support Markdown-aware writeback (inline + block-level)

## Requirements

- Python 3.10+
- `python-docx`

Install:

```bash
py -3 -m pip install python-docx
```

Upstream project (`python-docx`):

- <https://github.com/python-openxml/python-docx>

## Quickstart

Extract structure:

```bash
py -3 scripts/extract_docx_for_llm.py --in input.docx --out structure.v2.json
```

Optional RAG output:

```bash
py -3 scripts/extract_docx_for_llm.py --in input.docx --out structure.v2.json --rag-output rag.v1.json
```

Apply patch:

```bash
py -3 scripts/apply_docx_patch.py --in input.docx --out output.docx --patch patch.json
```

Preview helpers:

```bash
py -3 scripts/docx_preview.py --help
```

## Patch operations

Currently supported in `scripts/apply_docx_patch.py`:

- `replace_text`
- `set_paragraph`
- `delete_paragraph`
- `replace_paragraph_range`
- `replace_paragraph_range_markdown`

Full schemas and contracts are documented in `SKILL.md` and `references/python-docx-quickref.md`.

## Safety model

- Fails on unexpected match counts (`expected_matches`)
- Refuses risky cross-run replacements
- Supports expectation guards (`expected_*`) to avoid wrong target writes
- Heading-protection for range replacements by default

## Test

```bash
py -3 scripts/selftest.py
```

## Repository structure

- `scripts/` – executable tools
- `references/` – quick reference notes and spec snippets:
  - `references/ooxml-numbering-notes.md` (OOXML numbering model + restart behavior)
  - `references/python-docx-list-behavior-notes.md` (official python-docx list/style behavior)
- `SKILL.md` – canonical workflow/spec documentation

For numbering/list bugs, prefer Microsoft Learn OpenXML docs as the trusted external source before applying fixes.

## Sync to agent skill installation

This repository is the source of truth. After changes, sync tracked/non-ignored files into the agent skill folder (e.g. `boku-martin/skills/python-docx-editor`) so runtime uses the same version.

Example (PowerShell):

```powershell
git -C C:\Users\dagobert-ai\.openclaw\workspace\projects\python-docx-editor `
  ls-files --cached --others --exclude-standard
```

Copy that file list to the target skill directory.

## License

GNU General Public License v3.0 (GPL-3.0). See `LICENSE`.

## Status

Project is focused on practical CLI workflows and LLM-assisted editing pipelines.
