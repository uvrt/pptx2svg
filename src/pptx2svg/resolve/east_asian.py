"""Which face PowerPoint draws a run's Japanese in.

Measured in PowerPoint for Mac with ``tools/make_font_resolution_probe.py`` (the
observations are ``tests/fixtures/font-resolution-probe.json``), and two steps:

1. **The run's East Asian name.**  Its own ``<a:ea>``, inherited as any typeface is; a
   ``+mn-ea``/``+mj-ea`` pointer -- or nothing at all, which every master's default text
   style spells ``+mn-ea`` -- names the theme's ``Jpan`` entry for a Japanese run and its
   ``<a:ea>`` otherwise (:func:`ooxml_common.text.fontmap.theme_east_asian`).  An empty
   slot names nothing; the ``Jpan`` entry is not a fallback behind it.
2. **Whether that face draws it.**  Where the name finds an installed face with Japanese
   glyphs, that face does; where nothing is named and the run's *Latin* face has them, it
   does.  Where it finds none -- nothing named, a face that is not
   installed, a Latin face named as the East Asian one -- PowerPoint draws MS Gothic or MS
   Mincho by the PANOSE of the face named (else the Latin face), and MS Gothic where that
   is not installed either (:func:`ooxml_common.text.fontmap.japanese_fallback`).

Shapes, text boxes and table cells all follow it.  Charts do not: their labels fall
through to the theme's ``Jpan`` entry (:func:`ooxml_common.text.fontmap.east_asian_family`).

"Installed" is a question about a machine.  With the host's faces in use
(``ConvertOptions.host_fonts``), a face :func:`pptx2svg.fonts.office.find` finds here is
judged by its own cmap and PANOSE, as PowerPoint would judge it, and on a Mac with Office
a face it does not find is one PowerPoint would not find either.  Elsewhere, and without
the host's faces, a face is judged by what this library knows: the Japanese faces of the
substitution table draw Japanese, and a Latin face's PANOSE is the one recorded for
the bundled face that draws it (Carlito's for Calibri), or else taken from its CSS
generic.
"""

from __future__ import annotations

import functools

from ooxml_common.text.fontmap import (
    JAPANESE_GOTHIC,
    covers_east_asian,
    family_key,
    generic_family,
    japanese_fallback,
    metrics_fallback_font,
)

#: What a face must have to count as drawing Japanese: a kana.
_PROBE = "あ"
#: A PANOSE that :func:`ooxml_common.text.fontmap.panose_is_sans` reads as sans serif,
#: for a face known only by its CSS generic.
_SANS = (2, 11, 0, 0)


class EastAsianFaces:
    """``face(named, latin)``: the face a run's East Asian text is drawn in."""

    def __init__(self, host: bool = False) -> None:
        self.host = host

    def face(self, named: str | None, latin: str | None) -> str | None:
        if named and self.draws_japanese(named):
            return named
        if not named and latin and self.draws_japanese(latin):
            # A Japanese face named only as the Latin one draws the Japanese too.
            return latin
        basis = named or latin
        return japanese_fallback(self.panose(basis) if basis else None)

    # -- what this machine has -------------------------------------------------------

    @functools.lru_cache(maxsize=None)  # noqa: B019 -- one instance per conversion
    def draws_japanese(self, name: str) -> bool:
        face = _regular(name) if self.host else None
        if face is not None:
            return _covers(face, _PROBE)
        if self.host and _office_here():
            return False  # PowerPoint's own places were searched: it is not installed
        # Not looked for, or looked for on a machine without Office: what the library
        # knows, so that the SVG does not hang on whether a bundle happens to be there.
        return covers_east_asian(name)

    @functools.lru_cache(maxsize=None)  # noqa: B019
    def panose(self, name: str) -> tuple | None:
        """The PANOSE of the face ``name`` finds, ``None`` where it finds none."""
        face = _regular(name) if self.host else None
        if face is not None:
            return face.panose
        if self.host and _office_here():
            return None
        bundled = _bundled_panose().get(family_key(metrics_fallback_font(name) or name))
        if bundled is not None:
            return bundled
        # A face known only by name.  An unknown one is "sans-serif" too, which is what
        # PowerPoint does with a face it does not have.
        return _SANS if generic_family(name) == "sans-serif" else (0,)


def _office_here() -> bool:
    """Whether this machine has Office's fonts, so that a face not found is one PowerPoint
    would not find either (:func:`pptx2svg.fonts.office.available`)."""
    from ..fonts import office

    return office.available()


def _regular(name: str):
    from ..fonts import office

    faces = office.find(name)
    for face in faces:
        if not face.bold and not face.italic:
            return face
    return faces[0] if faces else None


@functools.lru_cache(maxsize=None)
def _covers(face, char: str) -> bool:
    from ooxml_common.fonts.office import Face, font_offsets, read_file

    try:
        data = read_file(face.path)
        return Face(data, font_offsets(data)[face.number]).glyph(char) is not None
    except Exception:  # noqa: BLE001 -- an unreadable face draws nothing
        return False


#: The first four PANOSE bytes of the bundled families' regular faces, read from the files
#: and written down so that the SVG does not depend on whether the bundle is installed.
#: Raleway's and Tinos's are all zero, which PowerPoint reads as "not sans": MS Mincho.
_BUNDLED_PANOSE = {
    "arimo": (2, 11, 6, 4), "caladea": (2, 4, 5, 3), "carlito": (2, 15, 5, 2),
    "cousine": (2, 7, 4, 9), "lato": (2, 15, 5, 2), "noto sans jp": (2, 11, 2, 0),
    "raleway": (0, 0, 0, 0), "tinos": (0, 0, 0, 0),
}


def _bundled_panose() -> dict:
    return _BUNDLED_PANOSE


__all__ = ["EastAsianFaces", "JAPANESE_GOTHIC"]
