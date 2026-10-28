# python-docx-editor

A DOCX editor for AI Agents: lightweight CLI toolkit for safe, reproducible `.docx` extraction and patching with `python-docx`.

## What it does

- Extract DOCX into structured JSON (`docx-structure.v2`)
- Emit a read-only contextual review (`docx-review.v1`) for tracked changes, comments, and anchors
- Optionally emit RAG-friendly chunks (`docx-llm-chunks.v1`)
- Apply block-targeted patches with validation guards
- Support Markdown-aware writeback (inline + block-level)
- Convert Markdown to standalone or template-based DOCX files, including optional landscape tables
- Fill arbitrary table cells from JSON (`layout: "cell-map"`) for generic matrix/table use cases

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

Inspect comments, revisions, and highlights:

```bash
py -3 scripts/docx_ops.py inspect --in input.docx [--json]
```

With `--json`, the output keeps the legacy `comments`, `revisions`, and
`highlights` fields and adds a top-level `review` object (`docx-review.v1`);
the same object is added to the structured extraction below.

Fill generic table cells from JSON spec:

```bash
py -3 scripts/fill_docx_table_from_json.py --in input.docx --out output.docx --spec cell-map.json
```

Convert Markdown to DOCX:

```bash
py -3 scripts/markdown_to_docx.py --in input.md --out output.docx
```

Both Markdown converters skip optional leading YAML front matter (`---` … `---`), so document metadata is not rendered as visible text. A standalone `---` without a closing marker remains a normal Markdown horizontal rule.

Convert Markdown to DOCX using a styled Word template (preserves headers, footers, page size, margins, and styles):

```bash
py -3 scripts/markdown_to_template_docx.py --in input.md --out output.docx [--template my_template.docx] [--table-orientation portrait|landscape]
```

With `--table-orientation landscape`, an H2/H3 heading immediately followed by a Markdown table (blank lines allowed) is placed with that table on the landscape page. Consecutive heading-table blocks do not create an empty portrait page between them; subsequent prose returns to the template's portrait layout.

Example spec (SWOT 2x2):

```bash
references/examples/table-cell-map-swot.json
```

Notes:
- `table_index`, `row`, `col` are 1-based.
- Optional per-cell flags: `mode` (`replace|append`) and `clear_first` (`true|false`).

Preview helpers:

```bash
py -3 scripts/docx_preview.py --help
```

## Contextual review (`docx-review.v1`)

`scripts/docx_ops.py inspect --json`, `scripts/extract_docx_for_llm.py --out`,
and the optional `--rag-output` payload all carry the same additive top-level
`review` object. It is read-only: the source DOCX is never opened for writing.

Legacy compatibility:

- `inspect` keeps its `comments`, `revisions`, and `highlights` fields unchanged.
- v2 keeps `docx-structure.v2`, its block IDs, text, and stats unchanged.
- v1 keeps `docx-llm-chunks.v1`, its block IDs, chunks, and text unchanged.
- `review` never replaces or reinterprets `p_*` patch IDs; it uses XML locations.

Shape:

```json
{
  "schema": "docx-review.v1",
  "paragraphs": [
    {
      "location": {"part": "word/document.xml", "path": "/document[1]/body[1]/p[1]"},
      "text_original": "ADEFG",
      "text_proposed": "ABCDG",
      "text_markup": "ABCDEFG"
    }
  ],
  "revisions": [
    {
      "kind": "insert",
      "id": "10",
      "author": "Alice",
      "date": "2026-10-01T10:00:00Z",
      "text": "BC",
      "location": {"part": "word/document.xml", "path": "/document[1]/body[1]/p[1]"},
      "ranges": {"original": [1, 1], "proposed": [1, 3], "markup": [1, 3]},
      "missing_metadata": []
    }
  ],
  "comments": [],
  "limitations": [],
  "coverage": {"supported_parts": ["word/document.xml"], "text_views_complete": true}
}
```

Coordinates and views:

- `location.part` has no leading slash; `location.path` uses 1-based sibling
  indexes, for example `/document[1]/body[1]/tbl[1]/tr[1]/tc[1]/p[1]`.
  Nested tables add another `tbl/tr/tc/p` chain.
- Text views are UTF-8 codepoint offsets, 0-based and end-exclusive. `w:tab`
  becomes a tab, `w:br`/`w:cr` become a newline, and whitespace is preserved.
- `original` omits insertions and keeps deletions; `proposed` keeps insertions
  and omits deletions; `markup` keeps both. Text omitted from a view keeps a
  collapsed `[n, n]` range there, where `n` is the actual zero-width text
  boundary before the omitted content: a deletion of `EF` after proposed
  `ABCD` reports `proposed: [4, 4]` (not `[3, 3]`), a deletion at the paragraph
  start reports `[0, 0]`, and one at the end reports `[len, len]`.
- Only Word-namespace `w:ins`/`w:del`/`w:t` are read as Word content. Foreign
  lookalikes in another namespace are reported as unsupported content and never
  enter a text view.
- All paragraphs of the main document are included, including inner table
  cells and paragraphs inside `w:customXml` wrappers. Text-box content is not
  merged into the surrounding paragraph.

Revisions, comments, and limitations:

- One record per text-carrying `w:ins`/`w:del`; multiple runs inside one
  revision are merged. Missing `author`/`date` are `null` and named in
  `missing_metadata`; IDs stay strings.
- Comments carry `resolved` (`true`/`false`/`null`), `parent_id`, `reply_ids`,
  and an `anchor`. A missing done state is never reported as open. Anchors are
  `resolved`, `unresolved` (with diagnostics such as `orphan_anchor`,
  `duplicate_anchor`, `reversed_anchor`), or `thread_only` for replies without
  their own range. Segments preserve cross-paragraph ranges in all three views.
- Move, property, structural, and unknown revisions, text boxes, and review
  markers in any other WordprocessingML part are reported as structured
  `limitations`: headers, footers, footnotes, endnotes, `word/comments.xml`, the
  glossary, and every further `word/*.xml` story. `unsupported_content` covers
  body wrappers that are not traversed (for example `altChunk`) and
  foreign-namespace elements. `coverage.text_views_complete` is `false` when
  such a marker makes the text views incomplete; it is never a completeness
  guarantee for the whole package.
- Comment and thread ambiguity is diagnosed instead of guessed, and every
  diagnostic keeps the exact part, XML path, and comment/revision ID.
  `ambiguous_comment_id` leaves all affected anchors `unresolved`; duplicate
  `paraId`s yield `ambiguous_thread_metadata` without binding a done state or
  parent; self and cyclic `paraIdParent` relations yield
  `self_parent_reference`/`cyclic_parent_reference` with no trusted
  `parent_id` or `reply_ids`. `w15:done` accepts `true`/`false`/`1`/`0`, a
  missing value is `null`, and an invalid value stays `null` with
  `malformed_thread_metadata` at the offending `commentEx` location.
- Malformed ZIP/XML input fails with a clear error instead of returning an
  empty successful review.
- On Windows, piping CLI output can hit the legacy console code page (`cp1252`)
  and fail on non-ASCII review text. Set `PYTHONIOENCODING=utf-8` for review
  commands, for example
  `$env:PYTHONIOENCODING='utf-8'; py -3 scripts/docx_ops.py inspect --in input.docx --json`.
  This feature does not change the existing CLI console-encoding behaviour.

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
py -3 -m unittest discover -s tests -p test_docx_review.py -v
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

## Windows PDF export

Export DOCX to PDF via Microsoft Word on Windows:

`powershell
powershell -ExecutionPolicy Bypass -File scripts/export_docx_to_pdf.ps1 -InputPath input.docx -OutputPath output.pdf
`

Notes:
- This export path currently works only on Windows.
- It requires a local Microsoft Word installation because it uses the Word COM interface.
