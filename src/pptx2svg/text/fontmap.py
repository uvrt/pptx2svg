"""What we measure a face with, and what we draw it with -- kept to one answer.

The bug this module is shaped around: layout was computed from Calibri's metrics while
the SVG named a ``font-family`` that resolved, on most machines, to something else
entirely.  Nothing complained.  resvg does not warn when it substitutes a face; rendering
one string in Calibri, Carlito, Aptos, Noto Sans JP and Lato on a Mac without those fonts
produced five byte-identical PNGs.  Text appears, at widths nothing computed.

So every Office face gets one :class:`Substitution` record, and that record answers both
questions at once:

* *What do we measure with?*  :func:`metrics_for` -> ``METRICS[sub.metrics]``.
* *What do we draw with?*  :func:`substitute_for` -> ``sub.substitute``, a family
  :mod:`pptx2svg.fonts` actually ships.

For every face with a metric-compatible clone those are the same family, which is the
point: the two cannot drift apart because they are one field apart in one table, and
``tests/test_fonts.py`` asserts ``sub.metrics == sub.substitute`` wherever
``metric_compatible`` is set.

Aptos is the exception, and it is deliberate.  Microsoft's Office default since 2023 is
proprietary and has no open metric-compatible clone -- not Carlito, not Arimo, nothing.
We cannot both draw it correctly and measure it correctly, so we choose: measure with
*Aptos's own* advance widths (see :mod:`pptx2svg.text.metrics` for how those were
obtained without redistributing anything), draw with Carlito, whose widths sit closest of
anything we ship (-3.6% on a representative sentence, against Arimo's +5.2%).  Line
breaks then match PowerPoint's and each drawn line is a few percent short, which scores
better than re-wrapping the paragraph somewhere else.  :func:`substitution_report` calls
this out so it is visible rather than inferred.
"""

from __future__ import annotations

from dataclasses import dataclass

from .metrics import METRICS, FontMetrics

__all__ = [
    "DEFAULT_FONT_MAPPING",
    "SUBSTITUTIONS",
    "Substitution",
    "create_font_mapping",
    "font_family_value",
    "generic_family",
    "mapped_font",
    "metrics_fallback_font",
    "metrics_for",
    "substitution_for",
]


@dataclass(frozen=True)
class Substitution:
    """One Office face, and the single family we both measure and draw it with."""

    #: Family the deck asks for, spelled as PowerPoint spells it.
    office: str
    #: Family we hand the rasteriser.  Always one :mod:`pptx2svg.fonts` can supply.
    substitute: str
    #: Key into :data:`pptx2svg.text.metrics.METRICS`.  Equal to ``substitute`` unless
    #: no clone exists and we knowingly measure one face while drawing another.
    metrics: str
    #: True when ``substitute`` was built to ``office``'s advance widths, so the
    #: substitution moves glyph shapes but not a single line break.
    metric_compatible: bool = True
    #: Anything a caller should know that the flags above cannot say.  Surfaced verbatim
    #: by ``pptx2svg fonts --check``, so it is written for a human reading a table.
    caveat: str = ""

    @property
    def exact(self) -> bool:
        """True when we draw the very family the deck asked for (Lato, Raleway)."""
        return self.office == self.substitute


def _entries() -> list[Substitution]:
    """The substitution table, written out so each group can carry its reason."""
    rows: list[Substitution] = []

    # -- Metric-compatible clones.  Same widths, same line breaks, different outlines.
    #    Verified rather than assumed: Carlito/Calibri, Arimo/Arial, Tinos/Times New
    #    Roman and Cousine/Courier New each came back at a ratio of exactly 1.0000 over
    #    four representative strings, measured against the copies Office installs here.
    rows.append(Substitution("Calibri", "Carlito", "Carlito"))
    # Calibri Light is a distinct face and Carlito has no light cut.  It is 1.3% narrower
    # than Calibri, so Carlito draws it 1.3% wide -- small, but real, and there is no
    # light-weight Carlito to do better with.  It stays "compatible" because the verdict
    # is about *internal* consistency: we measure with Carlito and we draw with Carlito.
    rows.append(
        Substitution(
            "Calibri Light", "Carlito", "Carlito",
            caveat="Carlito clones Calibri, not Calibri Light; measured 1.3% wider",
        )
    )
    for office in ("Arial", "Helvetica"):
        rows.append(Substitution(office, "Arimo", "Arimo"))
    for office in ("Times New Roman", "Times"):
        rows.append(Substitution(office, "Tinos", "Tinos"))
    for office in ("Courier New", "Courier"):
        rows.append(Substitution(office, "Cousine", "Cousine"))

    # -- Faces that are themselves open.  Nothing is being substituted; the bundle just
    #    supplies them so a host without them still renders the deck as authored.
    rows.append(Substitution("Lato", "Lato", "Lato"))
    rows.append(Substitution("Raleway", "Raleway", "Raleway"))

    # -- Cambria.  Caladea is everywhere described as metric-compatible with Cambria, and
    #    for the line box it is, but its advance widths are not Cambria's: measured
    #    against the installed Cambria it runs 4.5% narrow.  So this gets the same
    #    treatment as Aptos -- measured with Cambria's own widths, drawn with Caladea,
    #    which is the closest serif we can ship.
    rows.append(Substitution("Cambria", "Caladea", "Cambria", metric_compatible=False))

    # -- Aptos.  No clone exists; see the module docstring.  Measured with its own
    #    advance widths, drawn as Carlito.
    rows.append(Substitution("Aptos", "Carlito", "Aptos", metric_compatible=False))
    # Aptos Display is a separate design, not a size of Aptos, and it is measurably wider.
    # Office does not install it: it downloads it on demand into its cloud-font cache,
    # which is why it looked absent until PowerPoint's own PDF export was found to embed
    # it.  It gets its own table for the same reason Aptos does.
    rows.append(
        Substitution("Aptos Display", "Carlito", "Aptos Display", metric_compatible=False)
    )
    # Aptos Narrow is a condensed cut with no table of its own; Aptos's is closer than
    # anything else we have, and saying so beats falling through to the 0.6 em guess.
    rows.append(Substitution("Aptos Narrow", "Carlito", "Aptos", metric_compatible=False))

    # -- Japanese.  Noto Sans JP is not metric-compatible with any of these (nothing is;
    #    the MS faces are proprietary and were never cloned), but every one of them is
    #    full-width for CJK, which is what dominates the measurement, and a Gothic stands
    #    in for a Gothic far better than a Latin fallback does.
    japanese_gothic = (
        "MS Gothic", "MS ゴシック", "MS PGothic", "MS Pゴシック",
        "Meiryo", "メイリオ", "Meiryo UI",
        "Yu Gothic", "游ゴシック", "Yu Gothic UI",
        "Hiragino Sans", "Hiragino Kaku Gothic ProN",
        "Noto Sans JP", "Noto Sans CJK JP",
    )
    for office in japanese_gothic:
        rows.append(
            Substitution(
                office, "Noto Sans JP", "Noto Sans JP",
                metric_compatible=(office in ("Noto Sans JP", "Noto Sans CJK JP")),
            )
        )
    # Mincho is a serif; Noto Sans JP is not, and we ship no Japanese serif.  Mapping it
    # anyway is still right: the alternative is no Japanese glyphs at all.
    japanese_mincho = (
        "MS Mincho", "MS 明朝", "MS PMincho", "MS P明朝",
        "Yu Mincho", "游明朝", "Hiragino Mincho ProN",
        "Noto Serif CJK JP", "Noto Serif JP",
    )
    for office in japanese_mincho:
        rows.append(
            Substitution(
                office, "Noto Sans JP", "Noto Sans JP", metric_compatible=False,
            )
        )
    return rows


#: Office family (as the deck spells it) -> what we measure and draw it with.
SUBSTITUTIONS: dict[str, Substitution] = {row.office: row for row in _entries()}

#: Lower-cased, width-normalised index, so lookups need not normalise at every call site.
_INDEX: dict[str, Substitution] = {}

#: PPTX font family -> replacement family name.  Kept as a plain ``dict[str, str]``
#: because it is public API and ``ConvertOptions.font_mapping`` merges over it; the
#: richer information lives in :data:`SUBSTITUTIONS`.
DEFAULT_FONT_MAPPING: dict[str, str] = {
    row.office: row.substitute for row in SUBSTITUTIONS.values() if not row.exact
}

#: Names that read as serif faces, for the generic family at the end of a font stack.
_SERIF_HINTS = (
    "mincho", "明朝", "times new roman", "georgia", "cambria", "garamond",
    "book antiqua", "palatino", "caslon", "baskerville", "constantia",
)


def _normalize_full_width(value: str) -> str:
    """Some PPTX themes spell font names full-width (``ＭＳ Ｐゴシック``)."""
    return "".join(
        " " if ch == "　" else (chr(ord(ch) - 0xFEE0) if "！" <= ch <= "～" else ch)
        for ch in value
    )


def _key(value: str) -> str:
    return _normalize_full_width(value).strip().lower()


for _row in SUBSTITUTIONS.values():
    _INDEX.setdefault(_key(_row.office), _row)


#: Weight and optical-size qualifiers a family name may carry.  Real decks name faces
#: like "游ゴシック Light", "Segoe UI Semibold" or "Aptos Narrow Bold": the same design at
#: a different weight, which our substitute covers with its own weight axis.  Stripping
#: them and retrying beats enumerating every combination, and beats the alternative of
#: treating "Yu Gothic Light" as a face we have never heard of.
_WEIGHT_SUFFIXES = (
    " light", " semilight", " semibold", " demibold", " medium", " black", " heavy",
    " thin", " extralight", " extrabold", " ultralight", " bold", " italic", " ui",
)


def substitution_for(font_family: str | None) -> Substitution | None:
    """The substitution record for a PPTX font name, or ``None`` if we know nothing."""
    if not font_family:
        return None
    # ``+mj-lt`` and friends are theme pointers that survived an unresolvable lookup --
    # a font scheme with no `cs` entry, usually.  They are not faces and never resolve.
    if font_family.startswith("+"):
        return None

    key = _key(font_family)
    found = _INDEX.get(key)
    if found is not None:
        return found

    # Peel qualifiers off the end, longest name first, so "Yu Gothic UI Light" tries
    # "Yu Gothic UI" before "Yu Gothic".
    while True:
        for suffix in _WEIGHT_SUFFIXES:
            if key.endswith(suffix) and len(key) > len(suffix):
                key = key[: -len(suffix)].strip()
                break
        else:
            return None
        found = _INDEX.get(key)
        if found is not None:
            return found


def create_font_mapping(user_mapping: dict[str, str] | None = None) -> dict[str, str]:
    mapping = dict(DEFAULT_FONT_MAPPING)
    if user_mapping:
        mapping.update(user_mapping)
    return mapping


def mapped_font(font_family: str | None, mapping: dict[str, str]) -> str | None:
    """Replacement family for a PPTX font name under ``mapping``.

    Honours a caller's overrides first, so ``ConvertOptions(font_mapping=...)`` still
    wins over the built-in table, then falls back to the normalised index.
    """
    if not font_family:
        return None
    direct = mapping.get(font_family)
    if direct is not None:
        return direct

    normalized = _normalize_full_width(font_family)
    if normalized != font_family:
        direct = mapping.get(normalized)
        if direct is not None:
            return direct

    lowered = normalized.strip().lower()
    for key, value in mapping.items():
        if _key(key) == lowered:
            return value

    substitution = substitution_for(font_family)
    return substitution.substitute if substitution else None


def metrics_for(font_family: str | None) -> FontMetrics | None:
    """Metrics table for a PPTX font name, or ``None`` when we have no data for it."""
    substitution = substitution_for(font_family)
    return METRICS.get(substitution.metrics) if substitution else None


def metrics_fallback_font(font_family: str | None) -> str | None:
    """The family we will actually draw with, so the SVG can ask for it by name."""
    substitution = substitution_for(font_family)
    return substitution.substitute if substitution else None


def generic_family(font_family: str) -> str:
    lowered = font_family.lower()
    if any(hint in lowered for hint in _SERIF_HINTS):
        return "serif"
    if "serif" in lowered and "sans" not in lowered:
        return "serif"
    if "mono" in lowered or "courier" in lowered or "consolas" in lowered:
        return "monospace"
    return "sans-serif"


def _escape(name: str) -> str:
    return name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def font_family_value(
    fonts: list[str | None], mapping: dict[str, str] | None = None
) -> str | None:
    """Build an SVG ``font-family`` stack: originals, then substitutes, then generic.

    Ordering matters -- the original name goes first so a host that really has Calibri
    uses it and matches PowerPoint exactly; the substitute only kicks in when it is
    missing, and because it is metric-compatible the layout computed from our tables
    still holds.  The trailing generic keyword is not decoration: with
    ``skip_system_fonts`` set, resvg draws *nothing at all* for a family it cannot
    resolve, and :data:`pptx2svg.fonts.GENERIC_FAMILY_DEFAULTS` points the generics at
    the bundle so an unknown face degrades to visible text at roughly the right size
    instead of a blank slide.
    """
    mapping = mapping if mapping is not None else DEFAULT_FONT_MAPPING
    unique: list[str] = []
    seen: set[str] = set()

    for font in fonts:
        if not font or font in seen:
            continue
        seen.add(font)
        unique.append(font)

        substitute = mapped_font(font, mapping)
        if substitute and substitute not in seen:
            seen.add(substitute)
            unique.append(substitute)

        fallback = metrics_fallback_font(font)
        if fallback and fallback not in seen:
            seen.add(fallback)
            unique.append(fallback)

    if not unique:
        return None

    parts = [f"'{_escape(name)}'" if " " in name else _escape(name) for name in unique]
    parts.append(generic_family(unique[0]))
    return ", ".join(parts)
