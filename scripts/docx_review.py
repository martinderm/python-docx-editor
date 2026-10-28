"""Read-only contextual review extraction for ``.docx`` packages.

This module implements the additive ``docx-review.v1`` payload shared by
``scripts/docx_ops.py inspect`` and ``scripts/extract_docx_for_llm.py``.  It
never mutates the source package: the ZIP is opened for reading only and the
XML is parsed with the standard library.

Coordinate conventions
----------------------
* ``location`` identifies a package part and a stable XPath-like XML path with
  1-based sibling indexes, for example
  ``/document[1]/body[1]/tbl[1]/tr[1]/tc[1]/p[1]``.
* Text views are built in document order: ``original`` omits insertions but
  keeps deletions, ``proposed`` keeps insertions but omits deletions, and
  ``markup`` contains both.  Ranges are 0-based, end-exclusive codepoint
  offsets.  Text that is omitted from a view keeps a collapsed ``[n, n]``
  range there, where ``n`` is the actual text boundary before the omitted
  content (for a deletion directly after proposed ``ABCD`` the boundary is
  ``[4, 4]``, never ``[3, 3]``).
* ``w:tab`` becomes ``\\t`` and ``w:br``/``w:cr`` become ``\\n``.  Whitespace is
  never normalised.

Only the main document story and its tables are supported for text extraction.
Every other WordprocessingML package part (comments, glossary, headers, footers,
footnotes, endnotes, and any further ``word/*.xml`` story) is scanned for
review markers and reported as a limitation instead of silently claiming
complete text views.  Body wrappers that are not traversed and foreign-namespace
lookalikes are diagnosed as unsupported content; foreign ``ins``/``del``/``t``
elements are never interpreted as Word revisions.
"""

from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"

DOCUMENT_PART = "word/document.xml"
COMMENTS_PART = "word/comments.xml"
COMMENTS_EXTENDED_PART = "word/commentsExtended.xml"
REVIEW_SCHEMA = "docx-review.v1"

_VIEWS = ("original", "proposed", "markup")

_TAB = "tab"
_BREAKS = {"br", "cr"}
_TEXT = "t"
_DELETED_TEXT = "delText"

_PROPERTY_CHANGE_TAGS = {
    "pPrChange",
    "rPrChange",
    "sectPrChange",
    "tblPrChange",
    "tblPrExChange",
    "trPrChange",
    "tcPrChange",
    "tblGridChange",
    "numberingChange",
}

_MOVE_TAGS = {
    "moveFrom",
    "moveTo",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
}

_STRUCTURAL_REVISION_TAGS = {
    "cellIns",
    "cellDel",
    "cellMerge",
}

_UNKNOWN_REVISION_TAGS = {
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
}

_TEXT_VIEW_LIMITATIONS = {
    "unsupported_structural_revision",
    "unsupported_move",
    "unsupported_property_change",
    "unsupported_revision",
    "unsupported_story",
    "unsupported_content",
}

# Every WordprocessingML package part is a potential story.  Parts under
# ``word/_rels`` are relationship graphs, not stories.
_WORD_XML_PART_RE = re.compile(r"^word/(?!_rels/).+\.xml$")

_SKIP_TEXT_SUBTREES = {
    "pPr",
    "rPr",
    "txbxContent",
    "pict",
    "drawing",
    "object",
    "instrText",
    "delInstrText",
    "sym",
    "footnoteReference",
    "endnoteReference",
    "commentReference",
}

# Property containers and terminal body metadata that carry no story text.
_NON_STORY_BLOCK_TAGS = {
    "sectPr",
    "tcPr",
    "trPr",
    "tblPr",
    "tblPrEx",
    "tblGrid",
    "customXmlPr",
    "sdtPr",
    "sdtEndPr",
}

_COMMENT_MARKER_TAGS = {
    "commentRangeStart",
    "commentRangeEnd",
    "commentReference",
}


class DocxReviewError(Exception):
    """Raised when a package cannot be interpreted as a reviewable DOCX."""


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _namespace(tag: str) -> str:
    if tag.startswith("{"):
        return tag[1:].split("}", 1)[0]
    return ""


def _is_word(element: ET.Element) -> bool:
    return _namespace(element.tag) == W_NS


def _attr(element: ET.Element, name: str) -> str | None:
    for key, value in element.attrib.items():
        if _local(key) == name:
            return value
    return None


def _has_text_content(element: ET.Element) -> bool:
    for descendant in element.iter():
        if not _is_word(descendant):
            continue
        if _local(descendant.tag) in {_TEXT, _DELETED_TEXT, _TAB, *_BREAKS}:
            return True
    return False


def _contains_revision_marker(element: ET.Element) -> bool:
    for descendant in element.iter():
        if not _is_word(descendant):
            continue
        tag = _local(descendant.tag)
        if tag in {"ins", "del"}:
            return True
        if tag in _COMMENT_MARKER_TAGS:
            return True
        if tag in _PROPERTY_CHANGE_TAGS or tag in _MOVE_TAGS:
            return True
        if tag in _STRUCTURAL_REVISION_TAGS or tag in _UNKNOWN_REVISION_TAGS:
            return True
    return False


def _parse_done(value: str | None) -> tuple[bool | None, bool]:
    """Return ``(resolved, malformed)`` for an ``xsd:boolean`` done value."""

    if value is None:
        return None, False
    normalized = value.strip().lower()
    if normalized in {"1", "true"}:
        return True, False
    if normalized in {"0", "false"}:
        return False, False
    return None, True


@dataclass
class _Marker:
    """A comment range boundary recorded inside one paragraph."""

    comment_id: str
    kind: str
    paragraph_index: int
    location: dict
    offsets: dict
    order: int


@dataclass
class _OpenRevision:
    kind: str
    element: ET.Element
    start: dict
    text_parts: list[str] = field(default_factory=list)


@dataclass
class _Paragraph:
    location: dict
    text_original: str
    text_proposed: str
    text_markup: str
    markers: list[_Marker] = field(default_factory=list)
    revisions: list[dict] = field(default_factory=list)

    def text(self, view: str) -> str:
        if view == "original":
            return self.text_original
        if view == "proposed":
            return self.text_proposed
        return self.text_markup


@dataclass
class _Comment:
    id: str | None
    author: str | None
    date: str | None
    text: str
    missing_metadata: list[str]
    last_para_id: str | None
    path: str
    resolved: bool | None = None
    parent_id: str | None = None
    reply_ids: list[str] = field(default_factory=list)
    thread_path: str | None = None


@dataclass
class _ThreadEntry:
    """One ``w15:commentEx`` record with its exact XML location."""

    element: ET.Element
    path: str


class _ParagraphBuilder:
    """Accumulates the three text views and coordinate events of a paragraph."""

    def __init__(self, location: dict) -> None:
        self.location = location
        self._parts: dict[str, list[str]] = {view: [] for view in _VIEWS}
        self._lengths: dict[str, int] = {view: 0 for view in _VIEWS}
        self._marker_order = 0
        self.markers: list[_Marker] = []
        self.revisions: list[dict] = []

    @property
    def offsets(self) -> dict[str, int]:
        return dict(self._lengths)

    def emit(self, text: str, mode: str) -> None:
        if not text:
            return
        if mode == "normal":
            views = _VIEWS
        elif mode == "insert":
            views = ("proposed", "markup")
        elif mode == "delete":
            views = ("original", "markup")
        else:
            return
        for view in views:
            self._parts[view].append(text)
            self._lengths[view] += len(text)

    def add_marker(self, comment_id: str, kind: str) -> None:
        self.markers.append(
            _Marker(
                comment_id=comment_id,
                kind=kind,
                paragraph_index=-1,
                location=dict(self.location),
                offsets=self.offsets,
                order=self._marker_order,
            )
        )
        self._marker_order += 1

    def text(self, view: str) -> str:
        return "".join(self._parts[view])


class _ReviewParser:
    def __init__(self, infile: Path) -> None:
        try:
            self._archive = zipfile.ZipFile(str(infile))
        except (OSError, zipfile.BadZipFile) as exc:
            raise DocxReviewError(f"cannot open DOCX package {infile}: {exc}") from exc
        self._names = set(self._archive.namelist())
        self.limitations: list[dict] = []
        self._limitation_keys: set[tuple] = set()
        self.paragraphs: list[_Paragraph] = []
        self.revisions: list[dict] = []
        self._markers: dict[str, list[_Marker]] = {}
        self._ambiguous_comment_ids: set[str] = set()

    def close(self) -> None:
        self._archive.close()

    # -- package access ----------------------------------------------------

    def _read_xml(self, part: str) -> ET.Element | None:
        if part not in self._names:
            return None
        try:
            data = self._archive.read(part)
        except (KeyError, OSError, zipfile.BadZipFile) as exc:
            raise DocxReviewError(f"cannot read {part}: {exc}") from exc
        try:
            return ET.fromstring(data)
        except ET.ParseError as exc:
            raise DocxReviewError(f"malformed XML in {part}: {exc}") from exc

    def _add_limitation(
        self, code: str, part: str, path: str, details: dict
    ) -> None:
        key = (code, part, path, tuple(sorted((k, str(v)) for k, v in details.items())))
        if key in self._limitation_keys:
            return
        self._limitation_keys.add(key)
        self.limitations.append(
            {
                "code": code,
                "part": part,
                "location": {"part": part, "path": path},
                "details": details,
            }
        )

    # -- build -------------------------------------------------------------

    def build(self) -> dict:
        if DOCUMENT_PART not in self._names:
            raise DocxReviewError(f"DOCX package has no {DOCUMENT_PART}")
        document = self._read_xml(DOCUMENT_PART)
        assert document is not None
        body = document.find(f"{{{W_NS}}}body")
        if body is None:
            raise DocxReviewError(f"{DOCUMENT_PART} has no w:body element")

        self._walk_block_container(body, "/document[1]/body[1]")
        for index, paragraph in enumerate(self.paragraphs):
            for marker in paragraph.markers:
                marker.paragraph_index = index
                self._markers.setdefault(marker.comment_id, []).append(marker)
            self.revisions.extend(paragraph.revisions)

        self._scan_part(document, DOCUMENT_PART, "/document[1]", story=False)
        for part in sorted(self._names):
            if part == DOCUMENT_PART or not _WORD_XML_PART_RE.match(part):
                continue
            root = self._read_xml(part)
            if root is None:
                continue
            self._scan_part(root, part, f"/{_local(root.tag)}[1]", story=True)

        comments = self._build_comments()
        self._report_orphan_markers(comments)

        text_views_complete = not any(
            limitation["code"] in _TEXT_VIEW_LIMITATIONS
            for limitation in self.limitations
        )
        return {
            "schema": REVIEW_SCHEMA,
            "paragraphs": [
                {
                    "location": paragraph.location,
                    "text_original": paragraph.text_original,
                    "text_proposed": paragraph.text_proposed,
                    "text_markup": paragraph.text_markup,
                }
                for paragraph in self.paragraphs
            ],
            "revisions": self.revisions,
            "comments": comments,
            "limitations": self.limitations,
            "coverage": {
                "supported_parts": [DOCUMENT_PART],
                "text_views_complete": text_views_complete,
            },
        }

    # -- paragraphs --------------------------------------------------------

    def _walk_block_container(self, container: ET.Element, path: str) -> None:
        counters: dict[str, int] = {}
        for child in container:
            tag = _local(child.tag)
            counters[tag] = counters.get(tag, 0) + 1
            child_path = f"{path}/{tag}[{counters[tag]}]"
            if not _is_word(child):
                self._report_unsupported_content(
                    child, DOCUMENT_PART, child_path, reason="foreign_namespace"
                )
                continue
            if tag == "p":
                self._add_paragraph(child, child_path)
            elif tag == "tbl":
                self._walk_table(child, child_path)
            elif tag == "sdt":
                content = child.find(f"{{{W_NS}}}sdtContent")
                if content is not None:
                    self._walk_block_container(
                        content, f"{child_path}/sdtContent[1]"
                    )
            elif tag == "customXml":
                self._walk_block_container(child, child_path)
            elif tag in {"ins", "del"}:
                # Paragraph-mark revision: the paragraphs it wraps are not
                # extracted, so the text views cannot be claimed complete.
                self._add_limitation(
                    "unsupported_structural_revision",
                    DOCUMENT_PART,
                    child_path,
                    {
                        "element": tag,
                        "revision_id": _attr(child, "id"),
                        "context": "paragraph_mark",
                    },
                )
            elif tag in _NON_STORY_BLOCK_TAGS:
                continue
            else:
                self._report_unsupported_content(
                    child, DOCUMENT_PART, child_path, reason="unhandled_block"
                )

    def _walk_table(self, table: ET.Element, path: str) -> None:
        counters: dict[str, int] = {}
        for child in table:
            tag = _local(child.tag)
            counters[tag] = counters.get(tag, 0) + 1
            child_path = f"{path}/{tag}[{counters[tag]}]"
            if not _is_word(child):
                self._report_unsupported_content(
                    child, DOCUMENT_PART, child_path,
                    reason="foreign_namespace",
                )
            elif tag == "tr":
                self._walk_row(child, child_path)
            elif tag == "sdt":
                content = child.find(f"{{{W_NS}}}sdtContent")
                if content is not None:
                    self._walk_table(content, f"{child_path}/sdtContent[1]")
            elif tag in {"tblPr", "tblGrid", "tblPrEx"}:
                continue
            else:
                self._report_unsupported_content(
                    child, DOCUMENT_PART, child_path, reason="unhandled_block"
                )

    def _walk_row(self, row: ET.Element, path: str) -> None:
        counters: dict[str, int] = {}
        for child in row:
            tag = _local(child.tag)
            counters[tag] = counters.get(tag, 0) + 1
            child_path = f"{path}/{tag}[{counters[tag]}]"
            if not _is_word(child):
                self._report_unsupported_content(
                    child, DOCUMENT_PART, child_path,
                    reason="foreign_namespace",
                )
            elif tag == "tc":
                self._walk_block_container(child, child_path)
            elif tag == "sdt":
                content = child.find(f"{{{W_NS}}}sdtContent")
                if content is not None:
                    self._walk_row(content, f"{child_path}/sdtContent[1]")
            elif tag == "trPr":
                continue
            else:
                self._report_unsupported_content(
                    child, DOCUMENT_PART, child_path, reason="unhandled_block"
                )

    def _add_paragraph(self, paragraph: ET.Element, path: str) -> None:
        builder = _ParagraphBuilder({"part": DOCUMENT_PART, "path": path})
        self._walk_paragraph(paragraph, builder, "normal", [], path)
        self.paragraphs.append(
            _Paragraph(
                location=builder.location,
                text_original=builder.text("original"),
                text_proposed=builder.text("proposed"),
                text_markup=builder.text("markup"),
                markers=builder.markers,
                revisions=builder.revisions,
            )
        )

    def _walk_paragraph(
        self,
        element: ET.Element,
        builder: _ParagraphBuilder,
        mode: str,
        open_revisions: list[_OpenRevision],
        path: str,
    ) -> None:
        counters: dict[str, int] = {}
        for child in element:
            tag = _local(child.tag)
            counters[tag] = counters.get(tag, 0) + 1
            child_path = f"{path}/{tag}[{counters[tag]}]"

            if not _is_word(child):
                # Foreign lookalikes (``x:ins``, ``x:del``, ``x:t``) are never
                # Word revisions and their text must not enter a text view.
                self._report_unsupported_content(
                    child, DOCUMENT_PART, child_path, reason="foreign_namespace"
                )
                continue
            if tag in _COMMENT_MARKER_TAGS and _attr(child, "id") is None:
                self._add_limitation(
                    "missing_marker_id",
                    DOCUMENT_PART,
                    child_path,
                    {"element": tag, "reason": "missing w:id"},
                )
                continue
            if tag == "commentRangeStart":
                comment_id = _attr(child, "id")
                if comment_id is not None:
                    builder.add_marker(comment_id, "start")
                continue
            if tag == "commentRangeEnd":
                comment_id = _attr(child, "id")
                if comment_id is not None:
                    builder.add_marker(comment_id, "end")
                continue
            if tag in ("ins", "del"):
                if not _has_text_content(child):
                    continue
                revision = _OpenRevision(
                    kind="insert" if tag == "ins" else "delete",
                    element=child,
                    start=builder.offsets,
                )
                self._walk_paragraph(
                    child,
                    builder,
                    revision.kind,
                    [*open_revisions, revision],
                    child_path,
                )
                self._finish_revision(builder, revision)
                continue
            if tag in _MOVE_TAGS or tag in _PROPERTY_CHANGE_TAGS:
                continue
            if tag in _SKIP_TEXT_SUBTREES:
                continue
            if tag == _TEXT:
                self._emit_text(builder, child.text or "", mode, open_revisions)
                continue
            if tag == _DELETED_TEXT:
                self._emit_text(
                    builder, child.text or "", "delete", open_revisions
                )
                continue
            if tag == _TAB:
                self._emit_text(builder, "\t", mode, open_revisions)
                continue
            if tag in _BREAKS:
                self._emit_text(builder, "\n", mode, open_revisions)
                continue
            if tag in _STRUCTURAL_REVISION_TAGS or tag in _UNKNOWN_REVISION_TAGS:
                continue
            self._walk_paragraph(
                child, builder, mode, open_revisions, child_path
            )

    def _emit_text(
        self,
        builder: _ParagraphBuilder,
        text: str,
        mode: str,
        open_revisions: list[_OpenRevision],
    ) -> None:
        builder.emit(text, mode)
        for revision in open_revisions:
            revision.text_parts.append(text)

    def _finish_revision(
        self, builder: _ParagraphBuilder, revision: _OpenRevision
    ) -> None:
        text = "".join(revision.text_parts)
        if not text:
            return
        end = builder.offsets
        if revision.kind == "insert":
            ranges = {
                "original": [revision.start["original"], revision.start["original"]],
                "proposed": [revision.start["proposed"], end["proposed"]],
                "markup": [revision.start["markup"], end["markup"]],
            }
        else:
            # Deleted text is absent from the proposed view.  The collapsed
            # anchor is the actual text boundary before the omitted content,
            # never the position of the preceding codepoint.
            boundary = revision.start["proposed"]
            ranges = {
                "original": [revision.start["original"], end["original"]],
                "proposed": [boundary, boundary],
                "markup": [revision.start["markup"], end["markup"]],
            }
        revision_id = _attr(revision.element, "id")
        author = _attr(revision.element, "author")
        date = _attr(revision.element, "date")
        missing_metadata: list[str] = []
        if author is None:
            missing_metadata.append("author")
        if date is None:
            missing_metadata.append("date")
        if revision_id is None:
            missing_metadata.append("id")
        builder.revisions.append(
            {
                "kind": revision.kind,
                "id": revision_id,
                "author": author,
                "date": date,
                "text": text,
                "location": dict(builder.location),
                "ranges": ranges,
                "missing_metadata": missing_metadata,
            }
        )

    # -- unsupported markers and other stories -----------------------------

    def _report_unsupported_content(
        self, element: ET.Element, part: str, path: str, *, reason: str
    ) -> None:
        namespace = _namespace(element.tag)
        self._add_limitation(
            "unsupported_content",
            part,
            path,
            {
                "element": _local(element.tag),
                "namespace": namespace or None,
                "reason": reason,
            },
        )

    def _scan_part(
        self, element: ET.Element, part: str, path: str, *, story: bool
    ) -> None:
        self._scan_children(element, part, path, story=story)

    def _scan_children(
        self,
        parent: ET.Element,
        part: str,
        parent_path: str,
        *,
        story: bool,
    ) -> None:
        counters: dict[str, int] = {}
        for child in parent:
            tag = _local(child.tag)
            counters[tag] = counters.get(tag, 0) + 1
            child_path = f"{parent_path}/{tag}[{counters[tag]}]"
            self._scan_node(
                child,
                part,
                child_path,
                story=story,
            )

    def _scan_node(
        self,
        element: ET.Element,
        part: str,
        path: str,
        *,
        story: bool,
    ) -> None:
        tag = _local(element.tag)

        if not _is_word(element):
            # Never read a foreign ``ins``/``del``/``t`` as a Word revision;
            # recurse only to find genuine Word markers inside a wrapper.
            self._scan_children(element, part, path, story=story)
            return

        if tag == "txbxContent":
            if _contains_revision_marker(element):
                self._add_limitation(
                    "unsupported_story",
                    part,
                    path,
                    {"element": "txbxContent", "context": "textbox"},
                )
            return

        if story and (
            tag in _COMMENT_MARKER_TAGS or tag in {"ins", "del"}
        ):
            # Every non-main WordprocessingML part is an unsupported story;
            # recognising the marker keeps it visible instead of dropping it.
            self._add_limitation(
                "unsupported_story",
                part,
                path,
                {
                    "element": tag,
                    "revision_id": _attr(element, "id"),
                    "context": "story",
                },
            )
            # Keep scanning nested markers so no further occurrence is hidden.
            self._scan_children(element, part, path, story=story)
            return

        if tag in _PROPERTY_CHANGE_TAGS:
            self._add_limitation(
                "unsupported_property_change",
                part,
                path,
                {
                    "element": tag,
                    "revision_id": _attr(element, "id"),
                    "context": _local(element.tag),
                },
            )
            return
        if tag in _MOVE_TAGS:
            self._add_limitation(
                "unsupported_move",
                part,
                path,
                {
                    "element": tag,
                    "revision_id": _attr(element, "id"),
                    "context": "move",
                },
            )
            return
        if tag in _STRUCTURAL_REVISION_TAGS:
            self._add_limitation(
                "unsupported_structural_revision",
                part,
                path,
                {
                    "element": tag,
                    "revision_id": _attr(element, "id"),
                    "context": "structural",
                },
            )
            return
        if tag in _UNKNOWN_REVISION_TAGS:
            self._add_limitation(
                "unsupported_revision",
                part,
                path,
                {
                    "element": tag,
                    "revision_id": _attr(element, "id"),
                    "context": "unknown",
                },
            )
            return
        if tag in ("ins", "del"):
            if not _has_text_content(element):
                self._add_limitation(
                    "unsupported_structural_revision",
                    part,
                    path,
                    {
                        "element": tag,
                        "revision_id": _attr(element, "id"),
                        "context": "structural",
                    },
                )
                return
            self._scan_children(element, part, path, story=story)
            return

        self._scan_children(element, part, path, story=story)

    # -- comments ----------------------------------------------------------

    def _build_comments(self) -> list[dict]:
        comments = self._read_comments()
        extended = self._read_comments_extended()

        by_id: dict[str, list[_Comment]] = {}
        for comment in comments:
            if comment.id is not None:
                by_id.setdefault(comment.id, []).append(comment)
        self._ambiguous_comment_ids = {
            comment_id
            for comment_id, matches in by_id.items()
            if len(matches) > 1
        }
        for comment_id in sorted(self._ambiguous_comment_ids):
            matches = by_id[comment_id]
            for comment in matches:
                self._add_limitation(
                    "ambiguous_comment_id",
                    COMMENTS_PART,
                    comment.path,
                    {
                        "comment_id": comment_id,
                        "occurrences": len(matches),
                    },
                )

        self._apply_thread_metadata(comments, extended)

        for comment in comments:
            if comment.parent_id is None:
                continue
            parents = by_id.get(comment.parent_id, [])
            if (
                len(parents) == 1
                and parents[0] is not comment
                and comment.id not in self._ambiguous_comment_ids
            ):
                parents[0].reply_ids.append(comment.id)
            else:
                self._add_limitation(
                    "ambiguous_parent_reference",
                    COMMENTS_PART,
                    comment.path,
                    {
                        "comment_id": comment.id,
                        "parent_id": comment.parent_id,
                    },
                )

        return [
            {
                "id": comment.id,
                "author": comment.author,
                "date": comment.date,
                "text": comment.text,
                "missing_metadata": comment.missing_metadata,
                "resolved": comment.resolved,
                "parent_id": comment.parent_id,
                "reply_ids": comment.reply_ids,
                "anchor": self._build_anchor(comment),
            }
            for comment in comments
        ]

    def _read_comments(self) -> list[_Comment]:
        root = self._read_xml(COMMENTS_PART)
        if root is None:
            return []
        root_path = f"/{_local(root.tag)}[1]"
        counters: dict[str, int] = {}
        comments: list[_Comment] = []
        for element in root:
            tag = _local(element.tag)
            counters[tag] = counters.get(tag, 0) + 1
            if tag != "comment" or not _is_word(element):
                continue
            path = f"{root_path}/{tag}[{counters[tag]}]"
            comment_id = _attr(element, "id")
            author = _attr(element, "author")
            date = _attr(element, "date")
            missing_metadata: list[str] = []
            if author is None:
                missing_metadata.append("author")
            if date is None:
                missing_metadata.append("date")
            if comment_id is None:
                missing_metadata.append("id")
            paragraphs = [
                child
                for child in element
                if _local(child.tag) == "p" and _is_word(child)
            ]
            texts = [self._plain_text(child) for child in paragraphs]
            while texts and texts[-1] == "":
                texts.pop()
            last_para_id = (
                _attr(paragraphs[-1], "paraId") if paragraphs else None
            )
            comments.append(
                _Comment(
                    id=comment_id,
                    author=author,
                    date=date,
                    text="\n".join(texts),
                    missing_metadata=missing_metadata,
                    last_para_id=last_para_id,
                    path=path,
                )
            )
        return comments

    def _plain_text(self, element: ET.Element) -> str:
        parts: list[str] = []

        def walk(node: ET.Element) -> None:
            if not _is_word(node):
                return
            tag = _local(node.tag)
            if tag in {"pPr", "rPr", "txbxContent", "pict", "drawing"}:
                return
            if tag == _TEXT or tag == _DELETED_TEXT:
                parts.append(node.text or "")
            elif tag == _TAB:
                parts.append("\t")
            elif tag in _BREAKS:
                parts.append("\n")
            else:
                for child in node:
                    walk(child)

        walk(element)
        return "".join(parts)

    def _read_comments_extended(self) -> dict[str, list[_ThreadEntry]]:
        root = self._read_xml(COMMENTS_EXTENDED_PART)
        if root is None:
            return {}
        root_path = f"/{_local(root.tag)}[1]"
        counters: dict[str, int] = {}
        entries: dict[str, list[_ThreadEntry]] = {}
        for element in root:
            tag = _local(element.tag)
            counters[tag] = counters.get(tag, 0) + 1
            if tag != "commentEx" or _namespace(element.tag) != W15_NS:
                continue
            path = f"{root_path}/{tag}[{counters[tag]}]"
            para_id = _attr(element, "paraId")
            if para_id is None:
                self._add_limitation(
                    "malformed_thread_metadata",
                    COMMENTS_EXTENDED_PART,
                    path,
                    {"reason": "missing paraId"},
                )
                continue
            entries.setdefault(para_id, []).append(_ThreadEntry(element, path))
        return entries

    def _apply_thread_metadata(
        self,
        comments: list[_Comment],
        extended: dict[str, list[_ThreadEntry]],
    ) -> None:
        by_para: dict[str, list[_Comment]] = {}
        for comment in comments:
            if comment.last_para_id is not None:
                by_para.setdefault(comment.last_para_id, []).append(comment)
        shared = {
            para_id: owners
            for para_id, owners in by_para.items()
            if len(owners) > 1
        }
        for para_id, owners in shared.items():
            entries = extended.get(para_id, [])
            path = entries[0].path if entries else owners[0].path
            self._add_limitation(
                "ambiguous_thread_metadata",
                COMMENTS_EXTENDED_PART,
                path,
                {
                    "para_id": para_id,
                    "comment_ids": [owner.id for owner in owners],
                    "occurrences": len(owners),
                },
            )

        candidates: dict[int, _Comment] = {}
        candidate_owners: dict[int, _Comment] = {}
        for comment in comments:
            if comment.last_para_id is None:
                continue
            if comment.last_para_id in shared:
                # A paraId shared by several comments identifies none of them,
                # so neither the done state nor a parent relation is trusted.
                continue
            entries = extended.get(comment.last_para_id, [])
            if len(entries) > 1:
                self._add_limitation(
                    "ambiguous_thread_metadata",
                    COMMENTS_EXTENDED_PART,
                    entries[0].path,
                    {
                        "comment_id": comment.id,
                        "para_id": comment.last_para_id,
                        "occurrences": len(entries),
                    },
                )
                continue
            if not entries:
                continue
            entry = entries[0]
            comment.thread_path = entry.path
            done_value = _attr(entry.element, "done")
            comment.resolved, malformed = _parse_done(done_value)
            if malformed:
                self._add_limitation(
                    "malformed_thread_metadata",
                    COMMENTS_EXTENDED_PART,
                    entry.path,
                    {
                        "comment_id": comment.id,
                        "reason": f"invalid done={done_value!r}",
                    },
                )

            parent_para = _attr(entry.element, "paraIdParent")
            if parent_para is None:
                continue
            matches = [
                candidate
                for candidate in comments
                if candidate.last_para_id == parent_para
            ]
            if len(matches) == 1 and matches[0] is comment:
                self._add_limitation(
                    "self_parent_reference",
                    COMMENTS_EXTENDED_PART,
                    entry.path,
                    {"comment_id": comment.id, "para_id": parent_para},
                )
            elif len(matches) > 1:
                self._add_limitation(
                    "ambiguous_parent_reference",
                    COMMENTS_EXTENDED_PART,
                    entry.path,
                    {"comment_id": comment.id, "parent_para_id": parent_para},
                )
            elif not matches:
                self._add_limitation(
                    "unknown_parent_reference",
                    COMMENTS_EXTENDED_PART,
                    entry.path,
                    {"comment_id": comment.id, "parent_para_id": parent_para},
                )
            else:
                candidates[id(comment)] = matches[0]
                candidate_owners[id(comment)] = comment

        cyclic = self._cyclic_comments(candidates)
        cyclic_ids = {id(comment) for comment in cyclic}
        for comment in cyclic:
            self._add_limitation(
                "cyclic_parent_reference",
                COMMENTS_EXTENDED_PART,
                comment.thread_path or comment.path,
                {"comment_id": comment.id},
            )
        for owner_id, parent in candidates.items():
            comment = candidate_owners[owner_id]
            if id(comment) in cyclic_ids or id(parent) in cyclic_ids:
                # A cyclic reply relation is never published as a thread.
                continue
            if comment.id is None or comment.id in self._ambiguous_comment_ids:
                continue
            if parent.id is None or parent.id in self._ambiguous_comment_ids:
                self._add_limitation(
                    "ambiguous_parent_reference",
                    COMMENTS_EXTENDED_PART,
                    comment.thread_path or comment.path,
                    {"comment_id": comment.id, "parent_id": parent.id},
                )
                continue
            comment.parent_id = parent.id

    def _cyclic_comments(
        self, candidates: dict[int, _Comment]
    ) -> list[_Comment]:
        cyclic: list[_Comment] = []
        visited: set[int] = set()
        for start in candidates.values():
            chain: list[_Comment] = []
            node: _Comment | None = start
            while node is not None and id(node) not in visited:
                visited.add(id(node))
                chain.append(node)
                node = candidates.get(id(node))
            if node is None:
                continue
            entry_index = next(
                (index for index, item in enumerate(chain) if item is node),
                None,
            )
            if entry_index is None:
                continue
            for member in chain[entry_index:]:
                if not any(member is item for item in cyclic):
                    cyclic.append(member)
        return cyclic

    def _report_orphan_markers(self, comments: list[dict]) -> None:
        known_ids = {
            comment["id"] for comment in comments if comment["id"] is not None
        }
        for comment_id, markers in self._markers.items():
            if comment_id in known_ids:
                continue
            marker = markers[0]
            self._add_limitation(
                "unresolved_comment_reference",
                DOCUMENT_PART,
                marker.location["path"],
                {"comment_id": comment_id, "markers": len(markers)},
            )

    # -- anchors -----------------------------------------------------------

    def _build_anchor(self, comment: _Comment) -> dict:
        if comment.id is None:
            return self._unresolved_anchor("missing_comment_id", {})
        if comment.id in self._ambiguous_comment_ids:
            # Several comments share this id, so no anchor can be attributed
            # to this specific record.
            return self._unresolved_anchor(
                "ambiguous_comment_id", {"comment_id": comment.id}
            )
        markers = self._markers.get(comment.id, [])
        starts = [marker for marker in markers if marker.kind == "start"]
        ends = [marker for marker in markers if marker.kind == "end"]

        if not starts and not ends:
            if comment.parent_id is not None:
                return {
                    "status": "thread_only",
                    "start": None,
                    "end": None,
                    "segments": [],
                    "diagnostics": [],
                }
            return self._unresolved_anchor("missing_anchor", {"comment_id": comment.id})

        if len(starts) > 1 or len(ends) > 1:
            return self._unresolved_anchor(
                "duplicate_anchor",
                {
                    "comment_id": comment.id,
                    "start_markers": len(starts),
                    "end_markers": len(ends),
                },
            )
        if not starts or not ends:
            return self._unresolved_anchor(
                "orphan_anchor",
                {
                    "comment_id": comment.id,
                    "start_markers": len(starts),
                    "end_markers": len(ends),
                },
            )

        start = starts[0]
        end = ends[0]
        if (start.paragraph_index, start.order) >= (
            end.paragraph_index,
            end.order,
        ):
            return self._unresolved_anchor(
                "reversed_anchor", {"comment_id": comment.id}
            )

        return {
            "status": "resolved",
            "start": {
                "location": dict(start.location),
                "offset": start.offsets["markup"],
            },
            "end": {
                "location": dict(end.location),
                "offset": end.offsets["markup"],
            },
            "segments": self._build_segments(start, end),
            "diagnostics": [],
        }

    def _unresolved_anchor(self, code: str, details: dict) -> dict:
        return {
            "status": "unresolved",
            "start": None,
            "end": None,
            "segments": [],
            "diagnostics": [{"code": code, "details": details}],
        }

    def _build_segments(self, start: _Marker, end: _Marker) -> list[dict]:
        segments: list[dict] = []
        for index in range(start.paragraph_index, end.paragraph_index + 1):
            paragraph = self.paragraphs[index]
            if index == start.paragraph_index:
                start_offsets = start.offsets
            else:
                start_offsets = {view: 0 for view in _VIEWS}
            if index == end.paragraph_index:
                end_offsets = end.offsets
            else:
                end_offsets = {
                    view: len(paragraph.text(view)) for view in _VIEWS
                }
            ranges = {
                view: [start_offsets[view], end_offsets[view]]
                for view in _VIEWS
            }
            segment = {"location": dict(paragraph.location), "ranges": ranges}
            for view in _VIEWS:
                segment[f"text_{view}"] = paragraph.text(view)[
                    start_offsets[view] : end_offsets[view]
                ]
            segments.append(segment)
        return segments


def build_review(infile: Path) -> dict:
    """Return the additive ``docx-review.v1`` payload for ``infile``.

    The source package is opened read-only.  Malformed ZIP/XML input raises
    :class:`DocxReviewError` instead of producing an empty successful payload.
    """

    parser = _ReviewParser(Path(infile))
    try:
        return parser.build()
    finally:
        parser.close()
