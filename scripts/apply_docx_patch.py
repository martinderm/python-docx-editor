#!/usr/bin/env python3
"""Apply minimal, block-targeted patches to .docx files.

Designed to work with block_ids produced by extract_docx_for_llm.py.

Supports plain text ops and markdown-aware rewrite ops.
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
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.opc.constants import RELATIONSHIP_TYPE as RT
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

    res.matches = para_text.count(old)

    for run in paragraph.runs:
        if old in run.text:
            c = run.text.count(old)
            run.text = run.text.replace(old, new)
            res.changes += c

    if res.changes < res.matches:
        res.cross_run_conflicts = res.matches - res.changes

    return res


def build_paragraph_index(doc: DocumentObject) -> dict[str, Paragraph]:
    out: dict[str, Paragraph] = {}
    p_i = 0
    h_i = 0
    f_i = 0

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


def clear_paragraph_numbering(paragraph: Paragraph) -> None:
    ppr = paragraph._p.pPr  # pylint: disable=protected-access
    if ppr is not None and ppr.numPr is not None:
        ppr.remove(ppr.numPr)


def style_looks_like_list(style_name: str | None) -> bool:
    if not style_name:
        return False
    s = style_name.lower()
    return any(k in s for k in ["list", "bullet", "number", "aufz", "aufl"])


def clear_runs(paragraph: Paragraph) -> None:
    for r in list(paragraph._p.r_lst):  # pylint: disable=protected-access
        paragraph._p.remove(r)  # pylint: disable=protected-access


def add_hyperlink(paragraph: Paragraph, text: str, url: str, *, bold: bool = False, italic: bool = False) -> None:
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")

    if bold:
        b = OxmlElement("w:b")
        r_pr.append(b)
    if italic:
        i = OxmlElement("w:i")
        r_pr.append(i)

    r_style = OxmlElement("w:rStyle")
    r_style.set(qn("w:val"), "Hyperlink")
    r_pr.append(r_style)

    new_run.append(r_pr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)  # pylint: disable=protected-access


def _unescape_md(text: str) -> str:
    return re.sub(r"\\([`*_\[\]()|>#-])", r"\1", text)


def render_inline_markdown(paragraph: Paragraph, text: str) -> None:
    clear_runs(paragraph)
    rest = text

    # order matters: link, bold+italic, bold, italic, code
    patterns = [
        ("link", re.compile(r"\[([^\]]+)\]\((https?://[^)]+)\)")),
        ("bolditalic", re.compile(r"\*\*\*(.+?)\*\*\*")),
        ("bold", re.compile(r"\*\*(.+?)\*\*")),
        ("italic", re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)")),
        ("code", re.compile(r"`([^`]+)`")),
    ]

    while rest:
        earliest = None
        earliest_kind = None
        earliest_match = None
        for kind, rx in patterns:
            m = rx.search(rest)
            if m and (earliest is None or m.start() < earliest):
                earliest = m.start()
                earliest_kind = kind
                earliest_match = m

        if earliest_match is None:
            paragraph.add_run(_unescape_md(rest))
            break

        if earliest_match.start() > 0:
            paragraph.add_run(_unescape_md(rest[: earliest_match.start()]))

        if earliest_kind == "link":
            add_hyperlink(paragraph, earliest_match.group(1), earliest_match.group(2))
        elif earliest_kind == "bolditalic":
            run = paragraph.add_run(_unescape_md(earliest_match.group(1)))
            run.bold = True
            run.italic = True
        elif earliest_kind == "bold":
            run = paragraph.add_run(_unescape_md(earliest_match.group(1)))
            run.bold = True
        elif earliest_kind == "italic":
            run = paragraph.add_run(_unescape_md(earliest_match.group(1)))
            run.italic = True
        elif earliest_kind == "code":
            run = paragraph.add_run(_unescape_md(earliest_match.group(1)))
            run.font.name = "Consolas"

        rest = rest[earliest_match.end() :]


def apply_text_or_runs(
    paragraph: Paragraph,
    *,
    text: str | None = None,
    runs: list[dict] | None = None,
    markdown: bool = False,
) -> None:
    if runs is not None:
        clear_runs(paragraph)
        for r in runs:
            if not isinstance(r, dict):
                raise ValueError("runs[] enthält Nicht-Objekt.")
            r_text = r.get("text", "")
            if not isinstance(r_text, str):
                raise ValueError("runs[].text muss String sein.")
            link = r.get("link")
            bold = bool(r.get("bold", False))
            italic = bool(r.get("italic", False))
            code = bool(r.get("code", False))
            if link:
                if not isinstance(link, str):
                    raise ValueError("runs[].link muss String (URL) sein.")
                add_hyperlink(paragraph, r_text, link, bold=bold, italic=italic)
            else:
                run = paragraph.add_run(r_text)
                run.bold = bold
                run.italic = italic
                if code:
                    run.font.name = "Consolas"
        return

    if text is None:
        raise ValueError("Es muss entweder text oder runs gesetzt sein.")

    if markdown:
        render_inline_markdown(paragraph, text)
    else:
        paragraph.text = text


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


def apply_set_paragraph_op(op: dict, paragraph_index: dict[str, Paragraph]) -> dict:
    block_id = op.get("block_id")
    text = op.get("text")
    runs = op.get("runs")
    style = op.get("style")
    markdown = bool(op.get("markdown", False))
    expected_contains = op.get("expected_contains")
    clear_list_format = bool(op.get("clear_list_format", True))

    if not isinstance(block_id, str) or not block_id:
        raise ValueError("set_paragraph op braucht 'block_id'.")
    if runs is None and not isinstance(text, str):
        raise ValueError("set_paragraph op braucht String in 'text' oder runs[].")
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

    if style:
        paragraph.style = style
    apply_text_or_runs(paragraph, text=text if isinstance(text, str) else None, runs=runs, markdown=markdown)

    if clear_list_format and not style_looks_like_list(style):
        clear_paragraph_numbering(paragraph)

    return {
        "op": "set_paragraph",
        "block_id": block_id,
        "status": "ok",
        "style": style if style else None,
        "new_length": len(paragraph.text or ""),
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

    return {"op": "delete_paragraph", "block_id": block_id, "status": "ok"}


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


def insert_paragraph_after(paragraph: Paragraph, style: str | None = None) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)  # pylint: disable=protected-access
    new_para = Paragraph(new_p, paragraph._parent)  # pylint: disable=protected-access
    if style:
        try:
            new_para.style = style
        except Exception:
            pass
    return new_para


def set_horizontal_rule(paragraph: Paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()  # pylint: disable=protected-access
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "auto")
    p_bdr.append(bottom)


def insert_table_after(paragraph: Paragraph, rows: list[list[str]], table_style: str | None = None) -> Paragraph:
    max_cols = max((len(r) for r in rows), default=1)
    body = paragraph._parent  # pylint: disable=protected-access
    table = body.add_table(rows=len(rows), cols=max_cols)
    if table_style:
        try:
            table.style = table_style
        except Exception:
            pass
    for r_i, row in enumerate(rows):
        for c_i in range(max_cols):
            table.cell(r_i, c_i).text = row[c_i] if c_i < len(row) else ""

    tbl = table._tbl
    tbl.getparent().remove(tbl)
    paragraph._p.addnext(tbl)  # pylint: disable=protected-access

    after = OxmlElement("w:p")
    tbl.addnext(after)
    return Paragraph(after, paragraph._parent)  # pylint: disable=protected-access


def parse_markdown_blocks(markdown: str) -> list[dict]:
    lines = markdown.splitlines()
    i = 0
    blocks: list[dict] = []

    def flush_para(buf: list[str]) -> None:
        if buf:
            blocks.append({"type": "paragraph", "text": " ".join(s.strip() for s in buf if s.strip())})
            buf.clear()

    para_buf: list[str] = []

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            flush_para(para_buf)
            i += 1
            continue

        # table
        if "|" in stripped and i + 1 < len(lines) and re.match(r"^\s*\|?[\s:\-\|]+\|?\s*$", lines[i + 1].strip()):
            flush_para(para_buf)
            table_lines = [stripped]
            i += 2
            while i < len(lines) and "|" in lines[i]:
                table_lines.append(lines[i].strip())
                i += 1
            rows = []
            for tl in table_lines:
                t = tl.strip().strip("|")
                cells = [c.strip() for c in t.split("|")]
                rows.append(cells)
            blocks.append({"type": "table", "rows": rows})
            continue

        # heading
        hm = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if hm:
            flush_para(para_buf)
            blocks.append({"type": "heading", "level": len(hm.group(1)), "text": hm.group(2).strip()})
            i += 1
            continue

        # horizontal rule
        if re.match(r"^(-{3,}|\*{3,}|_{3,})$", stripped):
            flush_para(para_buf)
            blocks.append({"type": "hr"})
            i += 1
            continue

        # blockquote
        if stripped.startswith(">"):
            flush_para(para_buf)
            qbuf = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                qbuf.append(lines[i].strip().lstrip(">").strip())
                i += 1
            blocks.append({"type": "quote", "text": " ".join(qbuf)})
            continue

        # unordered list
        um = re.match(r"^\s*[-*]\s+(.+)$", line)
        if um:
            flush_para(para_buf)
            items = []
            while i < len(lines):
                m = re.match(r"^\s*[-*]\s+(.+)$", lines[i])
                if not m:
                    break
                items.append(m.group(1).strip())
                i += 1
            blocks.append({"type": "ul", "items": items})
            continue

        # ordered list
        om = re.match(r"^\s*\d+[\.)]\s+(.+)$", line)
        if om:
            flush_para(para_buf)
            items = []
            while i < len(lines):
                m = re.match(r"^\s*\d+[\.)]\s+(.+)$", lines[i])
                if not m:
                    break
                items.append(m.group(1).strip())
                i += 1
            blocks.append({"type": "ol", "items": items})
            continue

        para_buf.append(line)
        i += 1

    flush_para(para_buf)
    return blocks


def render_markdown_blocks_after(anchor: Paragraph, markdown: str) -> tuple[Paragraph, int]:
    blocks = parse_markdown_blocks(markdown)
    inserted = 0
    cur = anchor

    for b in blocks:
        btype = b.get("type")
        if btype == "table":
            cur = insert_table_after(cur, b.get("rows", []), table_style="Table Grid")
            inserted += 1
            continue

        if btype == "hr":
            p = insert_paragraph_after(cur)
            set_horizontal_rule(p)
            cur = p
            inserted += 1
            continue

        if btype == "heading":
            level = int(b.get("level", 1))
            style = f"Heading {max(1, min(6, level))}"
            p = insert_paragraph_after(cur, style=style)
            render_inline_markdown(p, b.get("text", ""))
            clear_paragraph_numbering(p)
            cur = p
            inserted += 1
            continue

        if btype == "quote":
            p = insert_paragraph_after(cur, style="Quote")
            render_inline_markdown(p, b.get("text", ""))
            clear_paragraph_numbering(p)
            cur = p
            inserted += 1
            continue

        if btype in {"ul", "ol"}:
            style = "List Bullet" if btype == "ul" else "List Number"
            for item in b.get("items", []):
                p = insert_paragraph_after(cur, style=style)
                render_inline_markdown(p, item)
                cur = p
                inserted += 1
            continue

        # paragraph fallback
        p = insert_paragraph_after(cur, style="Normal")
        render_inline_markdown(p, b.get("text", ""))
        clear_paragraph_numbering(p)
        cur = p
        inserted += 1

    return cur, inserted


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

        style = entry.get("style")
        text = entry.get("text")
        runs = entry.get("runs")
        markdown = bool(entry.get("markdown", False))

        if style is not None and not isinstance(style, str):
            raise ValueError("replace_paragraph_range: new_paragraphs[].style muss String sein, wenn gesetzt.")
        if runs is None and not isinstance(text, str):
            raise ValueError("replace_paragraph_range: new_paragraphs[] braucht text oder runs[].")

        p = insert_paragraph_after(anchor, style=style)
        apply_text_or_runs(p, text=text if isinstance(text, str) else None, runs=runs, markdown=markdown)
        if not style_looks_like_list(style):
            clear_paragraph_numbering(p)
        anchor = p
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


def apply_replace_paragraph_range_markdown_op(op: dict, doc: DocumentObject, paragraph_index: dict[str, Paragraph]) -> dict:
    start_block_id = op.get("start_block_id")
    end_block_id = op.get("end_block_id")
    markdown = op.get("markdown")
    allow_headings = bool(op.get("allow_headings", False))

    if not isinstance(markdown, str) or not markdown.strip():
        raise ValueError("replace_paragraph_range_markdown braucht nicht-leeren markdown-String.")

    # Reuse validation and range selection by calling base op with placeholder paragraph.
    base_result = apply_replace_paragraph_range_op(
        {
            "start_block_id": start_block_id,
            "end_block_id": end_block_id,
            "new_paragraphs": [{"text": "__TMP__", "style": "Normal"}],
            "allow_headings": allow_headings,
            "expected_start_contains": op.get("expected_start_contains"),
            "expected_end_contains": op.get("expected_end_contains"),
        },
        doc,
        paragraph_index,
    )

    # remove tmp paragraph and insert markdown blocks in its place
    p_index = build_paragraph_index(doc)
    # find temp paragraph by marker
    tmp_para = next((p for p in p_index.values() if (p.text or "") == "__TMP__"), None)
    if tmp_para is None:
        raise ValueError("replace_paragraph_range_markdown: temporärer Absatz nicht gefunden.")

    anchor = tmp_para
    parent = tmp_para._element.getparent()  # pylint: disable=protected-access
    if parent is None:
        raise ValueError("replace_paragraph_range_markdown: kein Parent für temp Absatz.")

    # delete tmp first, keep previous sibling as anchor base
    prev = tmp_para._element.getprevious()  # pylint: disable=protected-access
    parent.remove(tmp_para._element)  # pylint: disable=protected-access
    if prev is None:
        # create anchor paragraph at top if needed
        new_p = OxmlElement("w:p")
        body = doc.element.body
        body.insert(0, new_p)
        anchor = Paragraph(new_p, doc)
    else:
        anchor = Paragraph(prev, doc)

    _, inserted_blocks = render_markdown_blocks_after(anchor, markdown)
    base_result["op"] = "replace_paragraph_range_markdown"
    base_result["inserted_markdown_blocks"] = inserted_blocks
    return base_result


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
        elif op_kind == "replace_paragraph_range_markdown":
            results.append(apply_replace_paragraph_range_markdown_op(op, doc, paragraph_index))
        else:
            raise ValueError(
                f"Unbekannte op '{op_kind}' bei Op #{i}. "
                "Unterstützt: replace_text, set_paragraph, delete_paragraph, "
                "replace_paragraph_range, replace_paragraph_range_markdown"
            )

    args.outfile.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(args.outfile))

    print(json.dumps({"in": str(args.infile), "out": str(args.outfile), "ops": len(results), "results": results}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(2)
