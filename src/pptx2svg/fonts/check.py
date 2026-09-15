"""Answer "will this deck draw at the widths we measured?" before anyone renders it.

Silent substitution is the failure this whole subsystem exists to stop, and the only way
to keep it from being silent is to have something that says it out loud.  That is this
module: it takes the faces a deck asks for, runs each one through the same
:data:`pptx2svg.text.fontmap.SUBSTITUTIONS` table the renderer uses, and reports where
each one lands.

Four outcomes, in decreasing order of trust:

``exact``
    The deck asked for a face we ship and we will draw with it.  Lato and Raleway, and
    equally Carlito, Arimo, Tinos, Cousine, Caladea and Noto Sans JP when a deck names
    one of those directly -- which decks authored on Linux or round-tripped through
    LibreOffice do, since Carlito and Caladea are LibreOffice's own Calibri and Cambria
    substitutes.  Every one of those used to report ``missing`` with "no substitute
    known", which was exactly backwards: we ship the file and measured the table.
``compatible``
    The deck asked for an Office face and we will draw a clone built to the same advance
    widths.  Glyph outlines differ from PowerPoint's; nothing else does.
``approximate``
    We will draw *something*, but its widths are not the ones we measured with -- Aptos
    drawn as Carlito, Japanese Mincho drawn as a Gothic.  Layout is close, not right.
``missing``
    The face is one we know nothing about, or the ``pptx2svg-fonts`` distribution is not
    installed so we can draw nothing at all.  In the first case the rasteriser falls back
    to the generic family -- a face of roughly the right size and definitely the wrong
    widths.  In the second, rendering falls back to the host's fonts entirely and the
    report says so once, loudly, instead of repeating it for every face.

Only ``exact`` and ``compatible`` count as faithful; :attr:`FontReport.faithful` is what
``pptx2svg fonts --check`` turns into an exit status, so a deck that cannot be rendered
faithfully fails a build instead of shipping wrong pixels.
"""

from __future__ import annotations

from dataclasses import dataclass

from ..text.fontmap import substitution_for
from ..text.metrics import METRICS
from . import INSTALL_HINT, available_families, bundle_mode

__all__ = [
    "FaceStatus",
    "FontReport",
    "check_deck",
    "check_families",
    "deck_families",
    "resolved_families",
]

#: Verdicts, best first.  Used for sorting the report and for the exit status.
VERDICTS = ("exact", "compatible", "approximate", "missing")


@dataclass(frozen=True)
class FaceStatus:
    """Where one requested face ends up."""

    #: Family name as the deck spells it.
    requested: str
    #: Family we will hand the rasteriser, or ``None`` when we have no answer at all.
    substitute: str | None
    #: Metrics table the layout was computed from, or ``None`` when it was guessed.
    metrics: str | None
    #: One of :data:`VERDICTS`.
    verdict: str
    #: Why, in a line, for the ``fonts --check`` table.
    reason: str

    @property
    def faithful(self) -> bool:
        return self.verdict in ("exact", "compatible")


@dataclass(frozen=True)
class FontReport:
    faces: tuple[FaceStatus, ...]
    #: Families the bundle can supply on this installation.
    available: frozenset[str]
    #: True when the caller asked to use the host's fonts as well as (or instead of) the
    #: bundle, which makes every verdict here a lower bound rather than the whole story.
    system_fonts: bool = False
    #: "bundled" or "system" -- see :func:`pptx2svg.fonts.bundle_mode`.
    mode: str = "bundled"

    @property
    def reproducible(self) -> bool:
        """Will two machines rendering this deck produce the same pixels?

        Separate from :attr:`faithful` on purpose.  A deck can be drawn faithfully and
        still not reproducibly (system fonts that happen to be right today), and it can
        be perfectly reproducible while drawing Aptos as Carlito.  Conflating the two is
        how "it looks fine on my machine" becomes a production incident.
        """
        return self.mode == "bundled" and not self.system_fonts

    @property
    def faithful(self) -> bool:
        """Every face draws at the widths the layout was computed from."""
        return self.mode == "bundled" and all(face.faithful for face in self.faces)

    def by_verdict(self, verdict: str) -> tuple[FaceStatus, ...]:
        return tuple(face for face in self.faces if face.verdict == verdict)


def _status(requested: str, available: frozenset[str]) -> FaceStatus:
    substitution = substitution_for(requested)

    if substitution is None:
        return FaceStatus(
            requested=requested,
            substitute=None,
            metrics=None,
            verdict="missing",
            reason="no substitute known; widths guessed, drawn with the generic family",
        )

    if substitution.substitute not in available:
        return FaceStatus(
            requested=requested,
            substitute=substitution.substitute,
            metrics=substitution.metrics if substitution.metrics in METRICS else None,
            verdict="missing",
            reason=f"{substitution.substitute} is not installed; run `{INSTALL_HINT}`",
        )

    if substitution.exact:
        return FaceStatus(
            requested=requested,
            substitute=substitution.substitute,
            metrics=substitution.metrics,
            verdict="exact",
            reason="drawn with the face the deck asked for",
        )

    if substitution.metric_compatible:
        return FaceStatus(
            requested=requested,
            substitute=substitution.substitute,
            metrics=substitution.metrics,
            verdict="compatible",
            reason=(
                substitution.caveat
                or f"{substitution.substitute} has {requested}'s advance widths"
            ),
        )

    measured = substitution.metrics
    drawn = substitution.substitute
    return FaceStatus(
        requested=requested,
        substitute=drawn,
        metrics=measured if measured in METRICS else None,
        verdict="approximate",
        reason=(
            substitution.caveat
            or (
                f"measured as {measured}, drawn as {drawn}; no metric-compatible clone "
                "exists"
                if measured != drawn
                else f"{drawn} is not built to {requested}'s widths"
            )
        ),
    )


def check_families(
    families: list[str] | tuple[str, ...], *, system_fonts: bool = False
) -> FontReport:
    """Report on an explicit list of face names."""
    available = available_families()
    mode = bundle_mode()
    seen: set[str] = set()
    statuses: list[FaceStatus] = []
    for name in families:
        # "+mj-lt" and friends are theme pointers, not faces.  One survives resolution
        # whenever a font scheme has no `cs` entry; reporting it as a missing font would
        # be noise that hides the real ones.
        if not name or name.startswith("+") or name in seen:
            continue
        seen.add(name)
        statuses.append(_status(name, available))
    statuses.sort(key=lambda face: (VERDICTS.index(face.verdict), face.requested))
    return FontReport(tuple(statuses), available, system_fonts, mode)


def resolved_families(resolved) -> list[str]:
    """Every typeface a resolved deck actually draws with.

    Deliberately taken from the *resolved* model rather than by grepping the XML for
    ``typeface="..."``: a theme's font scheme lists forty-odd script fallbacks a deck
    never uses, and ``+mj-lt`` is a pointer, not a face.  What matters is what a run ends
    up asking for after the inheritance cascade has run, which is exactly what the
    renderer will put in the SVG.
    """
    from .. import model as m

    names: set[str] = set()

    def visit_text(body) -> None:
        if body is None:
            return
        for paragraph in body.paragraphs:
            bullet_font = getattr(paragraph.properties, "bullet_font", None)
            if bullet_font:
                names.add(bullet_font)
            for run in paragraph.runs:
                for field_name in ("font_family", "font_family_ea", "font_family_cs"):
                    value = getattr(run.properties, field_name, None)
                    if value:
                        names.add(value)

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
    return check_families(deck_families(source), system_fonts=system_fonts)


def deck_families(source) -> list[str]:
    """Every typeface a deck asks for, parsed and resolved from scratch."""
    from .. import ConvertOptions, convert_pptx_to_model  # circular at module scope

    return resolved_families(convert_pptx_to_model(source, ConvertOptions()))
