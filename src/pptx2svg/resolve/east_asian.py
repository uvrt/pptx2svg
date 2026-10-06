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
(``ConvertOptions.host_fonts``) it is answered as PowerPoint would answer it here, by
:func:`pptx2svg.fonts.office.find` and the face's own cmap and PANOSE.  Without them --
the reproducible render -- by what this library knows: the Japanese faces of the
substitution table draw Japanese, and a Latin face's PANOSE is read from the bundled
face that draws it (Carlito's for Calibri), or else taken from its CSS generic.
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
        if self.host:
            face = _regular(name)
            return face is not None and _covers(face, _PROBE)
        return covers_east_asian(name)

    @functools.lru_cache(maxsize=None)  # noqa: B019
    def panose(self, name: str) -> tuple | None:
        """The PANOSE of the face ``name`` finds, ``None`` where it finds none."""
        if self.host:
            face = _regular(name)
            return face.panose if face is not None else None
        bundled = _bundled_panose().get(family_key(metrics_fallback_font(name) or name))
        if bundled is not None:
            return bundled
        # A face known only by name.  An unknown one is "sans-serif" too, which is what
        # PowerPoint does with a face it does not have.
        return _SANS if generic_family(name) == "sans-serif" else (0,)


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


@functools.lru_cache(maxsize=1)
def _bundled_panose() -> dict:
    """``family key -> PANOSE`` of the bundled faces' regular cuts (empty without the bundle)."""
    from ooxml_common.fonts import bundle_dir
    from ooxml_common.fonts.office import index

    directory = bundle_dir()
    if directory is None:
        return {}
    out = {}
    for key, places in index((("bundle", directory),)).items():
        faces = [face for face in places.get("bundle", ()) if not face.bold and not face.italic]
        if faces and faces[0].panose:
            out[key] = faces[0].panose
    return out


__all__ = ["EastAsianFaces", "JAPANESE_GOTHIC"]
