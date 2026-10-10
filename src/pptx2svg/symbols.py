"""Symbol faces (Wingdings, Symbol, ...) a conversion cannot draw, drawn as Unicode.

A Wingdings box bullet is ``<a:buChar char="q"/>``; without Wingdings, a renderer draws
"q".  When a deck asks for Symbol, Wingdings, Wingdings 2, Wingdings 3 or Webdings and the
conversion has no copy of the face -- not embedded, not in the application's font folders,
not on the host where the host's faces are in use -- its text and bullets are drawn as the
Unicode characters they stand for (:mod:`ooxml_common.text.symbol_fonts`: ``q`` is U+2751
❑, ``ü`` U+2713 ✓), in faces that have them, and a ``symbol-font-mapped`` warning says
so.  Symbol and Wingdings keep their own advances meanwhile, recorded from Word's copies,
so the rest of the line is laid out where PowerPoint puts it.  Where the face *is*
available it draws itself, which is the faithful answer.
"""

from __future__ import annotations

from typing import Iterable

from ooxml_common.text.symbol_fonts import symbol_face, translate

from . import model as m


def symbol_text(resolved) -> dict[str, tuple[str, str]]:
    """``{face key: (family as the deck names it, every character drawn in it)}`` for the
    symbol faces a resolved deck draws text or a character bullet in."""
    found: dict[str, tuple[str, list[str]]] = {}

    def note(family: str | None, text: str) -> None:
        face = symbol_face(family)
        if face is None or not text:
            return
        entry = found.setdefault(face, (family, []))
        entry[1].append(text)

    def visit_text(body) -> None:
        if body is None:
            return
        for paragraph in body.paragraphs:
            properties = paragraph.properties
            bullet = getattr(properties, "bullet", None)
            if isinstance(bullet, m.CharBullet):
                first = next((run for run in paragraph.runs if run.text), None)
                family = properties.bullet_font or (
                    first.properties.font_family if first is not None else None
                )
                note(family, bullet.char)
            for run in paragraph.runs:
                note(run.properties.font_family, run.text)

    def visit(element) -> None:
        if isinstance(element, m.GroupElement):
            for child in element.children:
                visit(child)
        elif isinstance(element, m.TableElement):
            for row in element.table.rows:
                for cell in row.cells:
                    visit_text(cell.text_body)
        else:
            visit_text(getattr(element, "text_body", None))

    for slide in resolved.slides:
        for element in slide.elements:
            visit(element)
    return {face: (family, "".join(texts)) for face, (family, texts) in found.items()}


def absent_faces(
    used: dict[str, tuple[str, str]],
    *,
    available: Iterable[str],
    host: bool,
) -> dict[str, str]:
    """The symbol faces of ``used`` the conversion has no copy of, as ``{face: family}``.

    ``available`` holds the normalised keys (:func:`ooxml_common.text.fontmap.family_key`)
    of faces the deck embeds or the application supplies; ``host`` says whether the
    host's own faces draw this conversion, in which case an installed face counts too.
    """
    from ooxml_common.text.fontmap import family_key

    keys = set(available)
    absent: dict[str, str] = {}
    for face, (family, _text) in used.items():
        if family_key(family) in keys:
            continue
        if host:
            from .fonts import office

            if office.find(family):
                continue
        absent[face] = family
    return absent


def mapped_warning(family: str, text: str) -> str:
    """The ``symbol-font-mapped`` message for one face."""
    distinct = "".join(dict.fromkeys(char for char in text if not char.isspace()))
    face = symbol_face(family) or ""
    drawn, unmapped = translate(face, distinct)
    pairs = ", ".join(
        f"{char!r}->{shown!r}" for char, shown in zip(distinct, drawn) if char != shown
    )
    message = (
        f"{family} is not available: its characters are drawn as the Unicode characters "
        f"they stand for ({pairs or 'none'}), in a face that has them"
    )
    if unmapped:
        message += f"; {len(unmapped)} with no Unicode equivalent are drawn as they are ({unmapped!r})"
    return message
