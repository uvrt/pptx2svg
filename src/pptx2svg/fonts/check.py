"""Answer "will this deck draw at the widths we measured?" before anyone renders it.

The report itself -- :class:`FaceStatus`, :class:`FontReport`, the four verdicts and
:func:`check_families` -- moved to :mod:`ooxml_common.fonts.check` with its history, and
takes plain family names.  What stays here is the half that reads them out of a deck:
:func:`resolved_families` walks pptx2svg's resolved slide model, and :func:`check_deck`
and :func:`deck_families` parse a deck to get one.
"""

from __future__ import annotations

from ooxml_common.fonts import check as _shared
from ooxml_common.fonts.check import (  # noqa: F401  -- the old import paths
    VERDICTS,
    FaceStatus,
    FontReport,
    check_families,
)

__all__ = [
    "FaceStatus",
    "FontReport",
    "check_deck",
    "check_families",
    "deck_families",
    "resolved_families",
]


def __getattr__(name: str):
    """Anything else the module had before it moved, read from where it lives now."""
    try:
        return getattr(_shared, name)
    except AttributeError:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}") from None


def resolved_families(resolved) -> list[str]:
    """Every typeface a resolved deck actually draws with.

    Deliberately taken from the *resolved* model rather than by grepping the XML for
    ``typeface="..."``: a theme's font scheme lists forty-odd script fallbacks a deck
    never uses, and ``+mj-lt`` is a pointer, not a face.  What matters is what a run ends
    up asking for after the inheritance cascade has run, which is exactly what the
    renderer will put in the SVG.
    """
    from .. import model as m
    from ..text.measure import is_cjk

    names: set[str] = set()

    def visit_text(body) -> None:
        if body is None:
            return
        for paragraph in body.paragraphs:
            bullet_font = getattr(paragraph.properties, "bullet_font", None)
            if bullet_font:
                names.add(bullet_font)
            for run in paragraph.runs:
                for field_name in ("font_family", "font_family_cs"):
                    value = getattr(run.properties, field_name, None)
                    if value:
                        names.add(value)
                # The East Asian face only counts when the run has East Asian text in
                # it.  `font_family_ea` is the *resolved* face -- the run's `<a:ea>`, or
                # the theme's, or its `<a:font script="Jpan"/>` -- so on a theme that
                # offers a Jpan face, every run in the deck carries one whether or not a
                # single CJK character is drawn.  Reporting those would put this
                # function back to listing the script fallbacks its own docstring says
                # it exists to exclude, and it showed up immediately: two Google Slides
                # templates with no Japanese anywhere began warning that ＭＳ Ｐゴシック
                # would be substituted.  The test is the one `render/text.py` splits on,
                # so the report agrees with what is actually drawn.
                east_asian = getattr(run.properties, "font_family_ea", None)
                if east_asian and any(is_cjk(ord(ch)) for ch in run.text or ""):
                    names.add(east_asian)

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

    # The theme's own major/minor faces matter even when no run names them: an empty
    # placeholder still has to be measured, and a deck's identity is in its theme.
    scheme = resolved.font_scheme
    for field_name in ("major_font", "minor_font", "major_font_ea", "minor_font_ea"):
        value = getattr(scheme, field_name, None)
        if value:
            names.add(value)

    return sorted(names)


def check_deck(source, *, system_fonts: bool = False) -> FontReport:
    """Report on every face a deck asks for."""
    from .. import ConvertOptions, convert_pptx_to_model  # circular at module scope

    resolved = convert_pptx_to_model(source, ConvertOptions())
    return check_families(
        resolved_families(resolved),
        system_fonts=system_fonts,
        embedded=resolved.embedded_fonts.families,
    )


def deck_families(source) -> list[str]:
    """Every typeface a deck asks for, parsed and resolved from scratch."""
    from .. import ConvertOptions, convert_pptx_to_model  # circular at module scope

    return resolved_families(convert_pptx_to_model(source, ConvertOptions()))
