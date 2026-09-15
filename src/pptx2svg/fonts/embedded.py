"""Fonts the deck brought with it.

``ppt/presentation.xml`` may carry a ``<p:embeddedFontLst>``: one entry per family, each
naming up to four ``ppt/fonts/*.fntdata`` parts by relationship id -- regular, bold,
italic, bold-italic.  PowerPoint writes it whenever the author ticks "Embed fonts in the
file", and template vendors do, because their whole product is a typeface choice the
recipient does not have.

Reading it beats any bundle that could be shipped.  Two third-party template decks
measured locally embed Anton, Arimo, Literata, Merriweather Sans, Merriweather Sans Light
and Inclusive Sans; every one of those raised ``font-substituted`` -- "no substitute
known; widths guessed" -- while sitting *inside the file being rendered*.  There are about
1,800 Google Fonts families and a bundle can never catch up, but a deck that carries its
face always can.

What this module does, in order:

1. Decode each referenced part (:mod:`pptx2svg.fonts.eot`), which means an EOT header
   parse, an embedding-rights check and usually MicroType Express decompression.
2. Re-check the rights on the *decoded* ``OS/2`` table, which is the foundry's own
   statement rather than the embedder's copy of it.
3. Relabel the face to the family and slot the deck assigned it
   (:func:`pptx2svg.fonts.sfnt.relabel`), without which it is not addressable.
4. Build a :class:`~pptx2svg.text.metrics.FontMetrics` table from the same bytes.

Step 4 is the point of the whole exercise.  Handing the extracted files to ``svg_to_png``
through ``font_files=`` would be half an implementation and the wrong half: by then
``convert_pptx_to_svg`` has already wrapped, autofitted and centred every line from the
static tables, so the deck would draw in the author's face at some other face's widths --
right glyphs, wrong line breaks.  The metrics built here go into the measurer, so the
extracted file is measured with and drawn with, which is the invariant the rest of this
package exists to protect.

**Failure is never fatal.**  A payload that will not decode, or a face whose ``fsType``
refuses us, drops back to the substitution path with a warning naming the reason -- the
same degradation contract ``metafile-rasterizer-missing`` implements.
"""

from __future__ import annotations

import hashlib
import os
import re
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Iterable

from ..opc import REL_FONT
from ..text.fontmap import family_key
from ..text.metrics import FontMetrics
from .eot import EotError, decode_eot, embedding_refusal
from .sfnt import SfntError, read_sfnt, relabel

__all__ = [
    "EmbeddedFace",
    "EmbeddedFonts",
    "NO_EMBEDDED_FONTS",
    "extract_embedded_fonts",
]

#: ``(bold, italic)`` for each ``<p:embeddedFont>`` child, in schema order.
_SLOTS = (
    ("regular", False, False),
    ("bold", True, False),
    ("italic", False, True),
    ("boldItalic", True, True),
)


@dataclass(frozen=True)
class EmbeddedFace:
    """One decoded and rights-checked face."""

    #: Family name as the deck spells it, which is what the SVG will ask for.
    family: str
    bold: bool
    italic: bool
    #: The TrueType file as it came out of the EOT, still carrying its own name table.
    source: bytes

    @property
    def data(self) -> bytes:
        """The face relabelled to this family and slot -- a complete TrueType file.

        Computed on demand and then cached, because only the PNG path needs it.  An SVG
        names fonts and embeds none, so a caller who only wants SVG should not pay to
        rewrite a name table and re-checksum 300 kB of tables per face.
        """
        cached = self.__dict__.get("_relabelled")
        if cached is None:
            cached = relabel(self.source, self.family, bold=self.bold, italic=self.italic)
            # frozen=True, so the cache goes in through __dict__ directly.
            object.__setattr__(self, "_relabelled", cached)
        return cached

    @property
    def filename(self) -> str:
        stem = re.sub(r"[^A-Za-z0-9]+", "-", self.family).strip("-") or "EmbeddedFont"
        suffix = ("-Bold" if self.bold else "") + ("-Italic" if self.italic else "")
        return f"{stem}{suffix or '-Regular'}.ttf"


@dataclass(frozen=True)
class EmbeddedFonts:
    """Everything one deck's ``<p:embeddedFontLst>`` yielded."""

    faces: tuple[EmbeddedFace, ...] = ()
    #: Normalised family key -> metrics measured from the extracted files.
    metrics: dict[str, FontMetrics] = field(default_factory=dict)
    #: ``(code, message)`` pairs for families that could not be used.
    problems: tuple[tuple[str, str], ...] = ()

    @property
    def families(self) -> frozenset[str]:
        """Normalised keys of the families this supplies, for the substitution report."""
        return frozenset(self.metrics)

    def __bool__(self) -> bool:
        return bool(self.faces)

    def write(self, directory: str) -> list[str]:
        """Write every face into ``directory`` and return the paths.

        Temporary files on disk are how the rasteriser is reached -- resvg's font
        database indexes files, not buffers -- and they are what LibreOffice does for the
        same reason (``EmbeddedFontsHelper::addEmbeddedFont`` writes a temp file and calls
        ``AddTempDevFont``).  The caller owns the directory's lifetime.
        """
        os.makedirs(directory, exist_ok=True)
        paths = []
        for face in self.faces:
            path = os.path.join(directory, face.filename)
            with open(path, "wb") as handle:
                handle.write(face.data)
            paths.append(path)
        return paths


NO_EMBEDDED_FONTS = EmbeddedFonts()


def extract_embedded_fonts(
    package,
    embedded_fonts,
    *,
    wanted_families: "Iterable[str] | None" = None,
) -> EmbeddedFonts:
    """Decode a deck's embedded faces.

    ``wanted_families`` limits the work to families the deck actually draws with; names
    are matched through :func:`~pptx2svg.text.fontmap.family_key`, so the caller can pass
    them as the deck spells them.  Decoding costs real time -- 0.15 s to 1.1 s per face --
    and a template deck routinely embeds a family that survives in the theme but appears
    on no slide.  Pass ``None`` to decode everything.
    """
    if not embedded_fonts:
        return NO_EMBEDDED_FONTS
    wanted = None if wanted_families is None else {family_key(n) for n in wanted_families}

    faces: list[EmbeddedFace] = []
    metrics: dict[str, FontMetrics] = {}
    problems: list[tuple[str, str]] = []

    for entry in embedded_fonts:
        family = (entry.typeface or "").strip()
        if not family:
            continue
        key = family_key(family)
        if wanted is not None and key not in wanted:
            continue

        decoded: dict[tuple[bool, bool], bytes] = {}
        for slot, bold, italic in _SLOTS:
            rel_id = getattr(entry, _SLOT_FIELDS[slot])
            if not rel_id:
                continue
            payload = _read_font_part(package, entry.part_path, rel_id)
            if payload is None:
                problems.append(
                    (
                        "font-embedded-unreadable",
                        f"{family} ({slot}): relationship {rel_id} has no font part",
                    )
                )
                continue
            try:
                font = _decode(payload)
                refusal = embedding_refusal(read_sfnt(font).fs_type)
            except (EotError, SfntError) as error:
                problems.append(
                    (
                        _problem_code(error),
                        f"{family} ({slot}): {error}; falling back to substitution",
                    )
                )
                continue
            if refusal is not None:
                problems.append(
                    (
                        "font-embedded-restricted",
                        f"{family} ({slot}) will not be used because {refusal}; "
                        "falling back to substitution",
                    )
                )
                continue
            decoded[(bold, italic)] = font

        if not decoded:
            continue

        try:
            table = _metrics_for(decoded)
        except SfntError as error:  # pragma: no cover - read_sfnt already succeeded
            problems.append(
                ("font-embedded-unreadable", f"{family}: {error}; falling back to substitution")
            )
            continue

        metrics[key] = table
        for (bold, italic), font in decoded.items():
            faces.append(
                EmbeddedFace(family=family, bold=bold, italic=italic, source=font)
            )

    return EmbeddedFonts(tuple(faces), metrics, tuple(problems))


#: Decoded fonts, keyed by a digest of the payload they came from.
#:
#: Decoding an embedded face costs 0.15 s to 1.1 s, and the same deck gets rendered more
#: than once often enough to matter: ``tests/fixtures/real-basic-theme.pptx`` embeds eight
#: faces and the test suite renders decks derived from it about forty times, which without
#: this took the suite from 21 s to 126 s.  Keyed by content digest rather than by part
#: path so that a deck derived from another still hits, which is exactly what those tests
#: do.
#:
#: Bounded, because the values are whole font files: 32 entries is at most ~11 MB, the
#: same order as the font bundle.  Failures are cached too -- a deck whose fonts cannot be
#: decoded should not pay for the attempt on every render.
_DECODE_CACHE_SIZE = 32
_decode_cache: "OrderedDict[bytes, bytes | EotError]" = OrderedDict()


def _decode(payload: bytes) -> bytes:
    digest = hashlib.sha256(payload).digest()
    result = _decode_cache.get(digest)
    if result is None:
        try:
            result = decode_eot(payload)
        except EotError as error:
            result = error
        _decode_cache[digest] = result
        if len(_decode_cache) > _DECODE_CACHE_SIZE:
            _decode_cache.popitem(last=False)
    else:
        _decode_cache.move_to_end(digest)
    if isinstance(result, EotError):
        raise result
    return result


_SLOT_FIELDS = {
    "regular": "regular",
    "bold": "bold",
    "italic": "italic",
    "boldItalic": "bold_italic",
}


def _problem_code(error: Exception) -> str:
    message = str(error)
    if message.startswith("embedding is not permitted"):
        return "font-embedded-restricted"
    return "font-embedded-undecodable"


def _read_font_part(package, part_path: str, rel_id: str) -> bytes | None:
    relationship = package.relationships(part_path).get(rel_id)
    if relationship is None or relationship.target_part is None:
        return None
    if relationship.type != REL_FONT:
        # A slot pointing at something that is not a font relationship is a malformed
        # deck, not a font we should try to parse.
        return None
    return package.read(relationship.target_part)


#: The ideograph ``cjk_width`` stands for, matching ``tools/extract_font_metrics.py``.
#: A kanji rather than a kana on purpose: kana are enumerated individually there and
#: ideographs are not, and in a proportional face the two disagree.
_CJK_PROBE = 0x7DE8  # CJK UNIFIED IDEOGRAPH-7DE8


def _metrics_for(decoded: dict[tuple[bool, bool], bytes]) -> FontMetrics:
    """Build one metrics table from the slots a family supplied.

    The upright table comes from the regular cut where there is one.  Failing that it
    comes from the italic, then the bold -- which matches the standing decision that
    italic is measured from the upright table anyway (divergence is at most 2.5% across
    the Office substitutes), and beats the 0.6 em per-character guess by a wide margin.
    """
    upright = _first(decoded, [(False, False), (False, True), (True, False), (True, True)])
    bold = _first(decoded, [(True, False), (True, True)])

    face = read_sfnt(upright)
    widths, default_width, cjk_width = _widths(face)

    bold_widths: dict[str, int] = {}
    bold_default = bold_cjk = 0
    if bold is not None and bold is not upright:
        bold_face = read_sfnt(bold)
        bold_widths, bold_default, bold_cjk = _widths(bold_face)
        if bold_face.units_per_em != face.units_per_em:
            # Different designs for the same family can disagree on unitsPerEm; the
            # measurer divides both tables by one value, so rescale rather than discard.
            scale = face.units_per_em / bold_face.units_per_em
            bold_widths = {char: round(width * scale) for char, width in bold_widths.items()}
            bold_default = round(bold_default * scale)
            bold_cjk = round(bold_cjk * scale)

    return FontMetrics(
        units_per_em=face.units_per_em,
        ascender=face.ascender,
        descender=face.descender,
        default_width=default_width,
        cjk_width=cjk_width,
        widths=widths,
        bold_default_width=bold_default,
        bold_cjk_width=bold_cjk,
        bold_widths=bold_widths,
    )


def _first(decoded, order):
    for slot in order:
        if slot in decoded:
            return decoded[slot]
    return None


def _widths(face) -> tuple[dict[str, int], int, int]:
    """Every character the face maps, plus the two fallbacks.

    The generated tables in :mod:`pptx2svg.text.metrics` store a fixed sample of about
    450 characters because they are source code a human reads.  This one is built at run
    time and thrown away, so it stores the whole cmap -- typically 500 to 3,000 entries,
    a fraction of a megabyte -- and nothing a deck can contain falls back to a guess.
    """
    widths = {
        chr(code_point): advance
        for code_point, advance in face.advances.items()
        if code_point <= 0x10FFFF
    }
    alphabet = [widths[c] for c in "abcdefghijklmnopqrstuvwxyz" if c in widths]
    # The mean lower-case advance, as in tools/extract_font_metrics.py: it beats the
    # font's advanceWidthMax and the width of a space, both of which earlier tables used
    # and both of which are far from typical.
    default_width = round(sum(alphabet) / len(alphabet)) if alphabet else face.units_per_em // 2
    cjk_width = face.advances.get(_CJK_PROBE, face.units_per_em)
    return widths, default_width, cjk_width
