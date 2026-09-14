"""Derive a test deck from a fixture by splicing extra parts into it.

Two of the features this suite covers -- EMF previews and SmartArt -- cannot be authored
by ``python-pptx`` and are absent from every fixture in the corpus.  Rather than commit
opaque binaries whose provenance nobody can check, the decks are *derived*: an existing
fixture is unzipped, new parts and relationships are added, a shape is spliced into
``slide1``'s shape tree, and the result is re-zipped in memory.

That keeps the added markup visible and reviewable in the test that uses it, which for
hand-written OOXML is the property that matters most -- a wrong namespace or a missing
relationship type is otherwise invisible.
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

SLIDE_PATH = "ppt/slides/slide1.xml"
SLIDE_RELS_PATH = "ppt/slides/_rels/slide1.xml.rels"
CONTENT_TYPES_PATH = "[Content_Types].xml"

RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def derive_deck(
    source: Path,
    *,
    parts: dict[str, bytes] | None = None,
    shapes_xml: str = "",
    slide_relationships: list[tuple[str, str, str]] | None = None,
    defaults: dict[str, str] | None = None,
    overrides: dict[str, str] | None = None,
) -> bytes:
    """Return the bytes of ``source`` with the given additions.

    ``slide_relationships`` are ``(id, type, target)`` triples added to slide 1's rels.
    ``defaults`` maps a file extension to a content type, ``overrides`` a part path to
    one.  ``shapes_xml`` is appended to slide 1's ``p:spTree``, so it lands last in
    z-order, on top of whatever the fixture already draws.
    """
    parts = parts or {}
    original = zipfile.ZipFile(source)
    names = original.namelist()

    written: dict[str, bytes] = {}
    for name in names:
        written[name] = original.read(name)

    if shapes_xml:
        slide = written[SLIDE_PATH].decode("utf-8")
        # Close the tree after the last child rather than before it: `</p:spTree>` is
        # unique in the document, so a plain replace on the end tag is unambiguous.
        written[SLIDE_PATH] = slide.replace("</p:spTree>", f"{shapes_xml}</p:spTree>").encode()

    if slide_relationships:
        written[SLIDE_RELS_PATH] = _with_relationships(
            written.get(SLIDE_RELS_PATH), slide_relationships
        )

    if defaults or overrides:
        written[CONTENT_TYPES_PATH] = _with_content_types(
            written[CONTENT_TYPES_PATH], defaults or {}, overrides or {}
        )

    written.update(parts)

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as out:
        for name, payload in written.items():
            out.writestr(name, payload)
    return buffer.getvalue()


def _with_relationships(existing: bytes | None, additions: list[tuple[str, str, str]]) -> bytes:
    if existing is None:
        existing = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{RELATIONSHIP_NS}"></Relationships>'
        ).encode()
    text = existing.decode("utf-8")
    added = "".join(
        f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"/>'
        for rel_id, rel_type, target in additions
    )
    return text.replace("</Relationships>", f"{added}</Relationships>").encode()


def _with_content_types(
    existing: bytes, defaults: dict[str, str], overrides: dict[str, str]
) -> bytes:
    text = existing.decode("utf-8")
    added = ""
    for extension, content_type in defaults.items():
        # A duplicate Default for an extension is invalid OPC, so only add what is new.
        if not re.search(rf'Extension="{extension}"', text, re.IGNORECASE):
            added += f'<Default Extension="{extension}" ContentType="{content_type}"/>'
    for part_path, content_type in overrides.items():
        added += f'<Override PartName="/{part_path}" ContentType="{content_type}"/>'
    return text.replace("</Types>", f"{added}</Types>").encode()
