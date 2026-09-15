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

The table is indexed by *any name a deck may spell*, not only by Office's names, and that
distinction was once a bug rather than a design.  Because every row was written from the
Office side, a family was recognisable only if some Office face pointed at it: Carlito,
Arimo, Tinos, Cousine and Caladea -- the five faces we ship, measure and draw with --
were not names we accepted.  ``metrics_for("Carlito")`` returned ``None`` and the string
was laid out from the 0.6 em per-character guess, measuring identically to a face that
does not exist.  Decks name them: Carlito and Caladea are LibreOffice's Calibri and
Cambria substitutes, and Arimo/Tinos/Cousine are the Chrome OS core fonts.  The identity
rows in :func:`_entries` now come from :data:`pptx2svg.fonts.BUNDLED_FAMILIES` itself, so
the invariant is structural: *every family we can draw is a family we can be asked for,
and it resolves to itself.*
"""

from __future__ import annotations

from dataclasses import dataclass

from ..fonts import BUNDLED_FAMILIES
from .metrics import METRICS, FontMetrics

__all__ = [
    "DEFAULT_FONT_MAPPING",
    "SUBSTITUTIONS",
    "Substitution",
    "covers_east_asian",
    "create_font_mapping",
    "east_asian_family",
    "family_key",
    "font_family_value",
    "generic_family",
    "mapped_font",
    "metrics_fallback_font",
    "metrics_for",
    "substitution_for",
    "typographic_family",
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
    #: False when neither the Office face nor ``substitute`` ships an italic design, so
    #: a rasteriser asked for one draws upright and PowerPoint's slant has to be
    #: synthesised.  See :data:`pptx2svg.render.text.SYNTHETIC_OBLIQUE_SHEAR`.
    has_italic_cut: bool = True
    #: True when the face draws kana and ideographs, so a CJK run may resolve to it.
    #: See :func:`covers_east_asian` for why this is a field rather than something read
    #: off the metric table.
    east_asian: bool = False
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
    #    the MS faces are proprietary and were never cloned), and it is the only Japanese
    #    face we ship, so it is what every one of them is *drawn* with.
    #
    #    What they are *measured* with is a separate question, and for the four MS faces
    #    the answer is their own table.  The "P" in MS PGothic means proportional: its
    #    katakana run from 0.648 em to 1.0 while Noto Sans JP's are uniformly 1.0, so
    #    measuring a line of katakana with Noto Sans JP's widths overstates it by up to a
    #    third and wraps it early.  This is the Aptos arrangement -- measured as one face,
    #    drawn as another -- and it is right for the same reason: PowerPoint laid the deck
    #    out with MS PGothic's advances, so matching those is what matches PowerPoint.
    japanese_gothic = (
        ("MS Gothic", "ＭＳ ゴシック"), ("MS ゴシック", "ＭＳ ゴシック"),
        ("MS PGothic", "ＭＳ Ｐゴシック"), ("MS Pゴシック", "ＭＳ Ｐゴシック"),
        ("Meiryo", "Noto Sans JP"), ("メイリオ", "Noto Sans JP"),
        ("Meiryo UI", "Noto Sans JP"),
        ("Yu Gothic", "Noto Sans JP"), ("游ゴシック", "Noto Sans JP"),
        ("Yu Gothic UI", "Noto Sans JP"),
        ("Hiragino Sans", "Noto Sans JP"),
        ("Hiragino Kaku Gothic ProN", "Noto Sans JP"),
        ("Noto Sans JP", "Noto Sans JP"), ("Noto Sans CJK JP", "Noto Sans JP"),
    )
    for office, table in japanese_gothic:
        rows.append(
            Substitution(
                office, "Noto Sans JP", table,
                metric_compatible=(office in ("Noto Sans JP", "Noto Sans CJK JP")),
                east_asian=True,
                # No Japanese face here has an italic cut -- not MS Gothic or MS Mincho
                # inside Office's .ttc files, and not the Noto Sans JP we ship, which is
                # a weight-axis variable font with no slant axis and no oblique sibling.
                # So whichever of them the rasteriser picks, italic text draws upright.
                has_italic_cut=False,
            )
        )
    # Mincho is a serif; Noto Sans JP is not, and we ship no Japanese serif.  Mapping it
    # anyway is still right: the alternative is no Japanese glyphs at all.  The MS cuts
    # measure from their own tables for the reason given above.
    japanese_mincho = (
        ("MS Mincho", "ＭＳ 明朝"), ("MS 明朝", "ＭＳ 明朝"),
        ("MS PMincho", "ＭＳ Ｐ明朝"), ("MS P明朝", "ＭＳ Ｐ明朝"),
        ("Yu Mincho", "Noto Sans JP"), ("游明朝", "Noto Sans JP"),
        ("Hiragino Mincho ProN", "Noto Sans JP"),
        ("Noto Serif CJK JP", "Noto Sans JP"), ("Noto Serif JP", "Noto Sans JP"),
    )
    for office, table in japanese_mincho:
        rows.append(
            Substitution(
                office, "Noto Sans JP", table,
                metric_compatible=False, has_italic_cut=False, east_asian=True,
            )
        )

    # -- The substitutes, under their own names.  Every row above answers "a deck asked
    #    for an Office face, what do we draw?", and for a long time that was the only
    #    question the table could answer: a family was a legal *input* only if some deck
    #    spelled it that way.  Carlito, Arimo, Tinos, Cousine and Caladea are therefore
    #    families we ship, measure and draw with, which we did not recognise when a deck
    #    named one of them -- ``substitution_for("Carlito")`` returned ``None`` and the
    #    string fell through to the 0.6 em per-character guess in
    #    :mod:`pptx2svg.text.measure`.  Measured at 18 pt on "Hamburgefonstiv 12345":
    #    Carlito, Arimo, Tinos, Cousine and Caladea each came back 280.800 px -- the same
    #    280.800 px as "Nonexistent Face" -- against the 235.055, 254.824, 233.965,
    #    302.449 and 231.312 px their own tables hold.  Cousine is the one that shows how
    #    little the guess is worth in either direction: monospaced, it is 7.7% *wider*
    #    than the fallback assumed, while Caladea is 17.6% narrower.  Lato, Raleway and
    #    Noto Sans JP escaped only because a deck spells them the way we bundle them, so
    #    they already had rows of their own.
    #
    #    Decks really do name these.  Carlito and Caladea exist because LibreOffice
    #    ships them as its Calibri and Cambria substitutes, so anything round-tripped
    #    through LibreOffice names them directly; Arimo, Tinos and Cousine are the Chrome
    #    OS core fonts, packaged on Debian as ``fonts-croscore``.
    #
    #    Derived from :data:`pptx2svg.fonts.BUNDLED_FAMILIES` rather than listed by hand,
    #    so a ninth bundled family cannot ship without being reachable by its own name.
    #    The guard skips a family that already has a row, which is how Noto Sans JP keeps
    #    its ``has_italic_cut=False`` -- it is a weight-axis variable font with no slant
    #    axis.  The five added here each ship an italic cut in the bundle
    #    (``Carlito-Italic.ttf`` and siblings), so the default is right for them.
    spoken_for = {row.office for row in rows}
    for family in sorted(BUNDLED_FAMILIES - spoken_for):
        rows.append(Substitution(family, family, family))

    # -- Liberation.  The same three designs again under another name: Liberation 2.x is
    #    built *from* the Chrome OS core fonts -- Sans from Arimo, Serif from Tinos, Mono
    #    from Cousine -- which is why ``tools/install-fonts-debian.sh`` installs
    #    ``fonts-liberation2`` deliberately, "because plenty of decks and themes name
    #    Liberation Sans directly".  They did not resolve either, for the same reason the
    #    five above did not: nothing made a name we can draw a name we can be asked for.
    #
    #    The width equality is inherited, not re-measured here: no Liberation file is
    #    installed on this machine, and none may enter the repository.  It rests on two
    #    measurements already recorded -- the character-by-character comparison in
    #    ``README.md`` (0 of 191 advances differ) and resvg drawing Arimo at ``wght=700``
    #    pixel-identically to static Liberation Sans Bold (``ROADMAP.md``).  Being built
    #    from the same outlines, they are the strongest metric-compatibility claim in
    #    this table rather than the weakest.
    #
    #    Only the three base families.  "Liberation Sans Narrow" is a genuinely different
    #    condensed design with no counterpart in the bundle, and "narrow" is not in
    #    :data:`_WEIGHT_SUFFIXES`, so it keeps falling through rather than being measured
    #    as its un-condensed parent.
    liberation = (
        ("Liberation Sans", "Arimo"),
        ("Liberation Serif", "Tinos"),
        ("Liberation Mono", "Cousine"),
    )
    for office, substitute in liberation:
        rows.append(
            Substitution(
                office, substitute, substitute,
                caveat=f"{substitute} and {office} are the same design under two names",
            )
        )
    return rows


#: Family name a deck may ask for -> what we measure and draw it with.  Keyed by the
#: *requested* spelling, which includes Office's names, the open faces those are
#: substituted by, and the aliases those in turn go by -- see :func:`_entries`.
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


def family_key(value: str) -> str:
    """Normalise a requested family name to its index key.

    Case-folded and outer-whitespace-stripped on purpose, so ``"carlito"``, ``"CARLITO"``
    and ``"Carlito "`` all reach the same row: OOXML puts no constraint on how a
    ``typeface`` attribute is capitalised, and a deck hand-edited or written by a
    generator carries whatever its author typed.  Public because it has become the whole
    font subsystem's notion of family identity: the substitution index, the metrics
    overlay a deck's embedded fonts install, and the `fonts --check` report all have to
    agree on when two spellings name the same face.  A name we fail to match does not fail
    loudly -- it silently becomes the 0.6 em guess -- so the lookup is deliberately the
    forgiving end of this module.

    *Inner* whitespace is left alone: ``"Times  New Roman"`` with two spaces does not
    resolve, and that is the decision rather than an oversight.  Collapsing runs of
    spaces would also have to collapse them in the :func:`font_family_value` output to
    stay honest, and a rasteriser matches ``font-family`` on the exact string; no deck in
    the corpus spells a family that way, so this buys a mismatch risk for nothing.
    """
    return _normalize_full_width(value).strip().lower()


for _row in SUBSTITUTIONS.values():
    _INDEX.setdefault(family_key(_row.office), _row)


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

    key = family_key(font_family)
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
        if family_key(key) == lowered:
            return value

    substitution = substitution_for(font_family)
    return substitution.substitute if substitution else None


def metrics_for(font_family: str | None) -> FontMetrics | None:
    """Metrics table for a PPTX font name, or ``None`` when we have no data for it."""
    substitution = substitution_for(font_family)
    return METRICS.get(substitution.metrics) if substitution else None


def covers_east_asian(font_family: str | None) -> bool:
    """Whether this family can actually draw Japanese, Chinese or Korean text.

    ``False`` for a face we know nothing about as well as for a Latin one, and the two
    cases want the same treatment: the East Asian answer would be a guess rather than a
    measurement, so the caller should try the next name in its cascade.

    **It cannot be derived from the metric tables, and finding that out is the useful
    part.**  ``FontMetrics.cjk_width`` is 1.0 em in every table, Latin ones included,
    because ``tools/extract_font_metrics.py`` writes ``units_per_em`` when the face has
    no glyph for its probe kanji -- "one em" is what a face says whether it draws the
    character beautifully or not at all.  Nor do the per-character rows help: only the
    two *proportional* MS faces earn any, because a row is written only where the advance
    disagrees with ``cjk_width``, so ＭＳ ゴシック -- a genuine Japanese face -- carries
    none while Cambria carries four (its four bracket forms).  Counting rows would have
    called Cambria Japanese and MS Gothic not.

    So it is a property of the *face*, recorded beside the other two in
    :class:`Substitution` where the Japanese rows already sit together.
    """
    substitution = substitution_for(font_family)
    return substitution is not None and substitution.east_asian


def east_asian_family(*candidates: str | None) -> str | None:
    """The face East Asian characters are actually drawn in, given a cascade of names.

    Call it with the names in OOXML's own order of precedence -- the run's ``<a:ea>``,
    then the theme font collection's ``<a:ea>``, then its ``<a:font script="Jpan"/>``,
    then the Latin face -- and it answers with the one that will draw the glyphs.

    **The precedence was read out of PowerPoint's own PDF export of
    ``real-financial-report.pptx``.**  Its charts name ``<a:latin typeface="Arial"/>``
    and nothing else, and its theme writes ``<a:ea typeface=""/>`` in both collections
    with ``<a:font script="Jpan" typeface="游ゴシック"/>`` beside it.  PowerPoint drew
    every Japanese category label in **YuGothic-Regular** and the Latin runs of the same
    labels (``DX``, ``CO2``) in **ArialMT** -- so an empty ``<a:ea>`` falls through to the
    script list, the script list beats the Latin face, and the two faces split *within*
    one label rather than one of them winning the whole of it.

    Two rules that are not obvious from the spec, and both are measured:

    * **An empty ``typeface=""`` is not a name.**  Every theme in this corpus writes one,
      and it means "this collection names no East Asian face", not "the empty face".
    * **A name with no East Asian glyphs loses to one that has them.**
      ``real-basic-theme.pptx`` writes ``<a:ea typeface="Raleway"/>`` -- a Latin face --
      on 9 of its runs, and PowerPoint drew their Japanese in MS Gothic and MS Mincho
      rather than in Raleway.  Taking the name at face value measures kana with Raleway's
      1.0 em ``cjk_width``, which is a guess wearing a measurement's clothes.

    The last resort is the first real name, whatever it is: a cascade that resolved to
    nothing would leave the emitted ``font-family`` empty, and a named face we have no
    table for still tells the rasteriser something.
    """
    named = [
        name for name in candidates if name and name.strip() and not name.startswith("+")
    ]
    for name in named:
        if covers_east_asian(name):
            return name
    return named[0] if named else None


def synthesises_italic(font_family: str | None) -> bool:
    """Whether italic for this family has to be faked, because no cut of it exists.

    PowerPoint slants such a face itself; resvg does not synthesise obliques at all, so
    `font-style="italic"` on a family with no italic design is silently a no-op and the
    run draws upright.  A family we know nothing about is assumed to have one, because
    skewing a face that does ship an italic would be a double slant.
    """
    substitution = substitution_for(font_family)
    return substitution is not None and not substitution.has_italic_cut


def metrics_fallback_font(font_family: str | None) -> str | None:
    """The family we will actually draw with, so the SVG can ask for it by name."""
    substitution = substitution_for(font_family)
    return substitution.substitute if substitution else None


#: The CSS generic each bundled family really belongs to.
#:
#: The heuristics below read a *name*, and these names say nothing: there is no "serif"
#: in "Tinos" or "Caladea" and no "mono" in "Cousine", so all three used to end their
#: stack in ``sans-serif``.  That is only the last resort in the stack, but the last
#: resort is exactly where it bites -- in ``system`` mode, or with the bundle
#: half-installed, a deck set in Tinos degraded to resvg's sans default.  Exact knowledge
#: about the families we ship, so it is consulted ahead of the guesses rather than added
#: to them.
_BUNDLED_GENERICS = {
    "arimo": "sans-serif",
    "caladea": "serif",
    "carlito": "sans-serif",
    "cousine": "monospace",
    "lato": "sans-serif",
    "noto sans jp": "sans-serif",
    "raleway": "sans-serif",
    "tinos": "serif",
}


def generic_family(font_family: str) -> str:
    known = _BUNDLED_GENERICS.get(family_key(font_family))
    if known is not None:
        return known
    lowered = font_family.lower()
    if any(hint in lowered for hint in _SERIF_HINTS):
        return "serif"
    if "serif" in lowered and "sans" not in lowered:
        return "serif"
    if "mono" in lowered or "courier" in lowered or "consolas" in lowered:
        return "monospace"
    return "sans-serif"


#: Style words that an Office family name appends to its *typographic* family, so that
#: stripping one recovers the name a rasteriser is likely to index the face under.  See
#: :func:`typographic_family` for why that matters and where this list came from.
#:
#: ``Ornaments`` is measured-and-excluded on purpose: "Hoefler Text Ornaments" really is
#: built that way, but it is a different glyph set rather than a weight of "Hoefler
#: Text", so falling back to the text family would draw letters where ornaments belong.
_TYPOGRAPHIC_STYLE_WORDS = frozenset(
    {
        "black",
        "bold",
        "book",
        "demibold",
        "display",
        "extrabold",
        "extralight",
        "hairline",
        "heavy",
        "light",
        "medium",
        "regular",
        "semibold",
        "semilight",
        "thin",
        "ultrabold",
        "ultralight",
    }
)


def typographic_family(font_family: str | None) -> str | None:
    """The family name a rasteriser indexes ``font_family`` under, if it differs.

    OpenType carries two family names: name ID 1, the four-style "compatible" family a
    PPTX spells (``Aptos Display``, ``Calibri Light``), and name ID 16, the *typographic*
    family that groups a whole superfamily under one name (``Aptos``, ``Calibri``) with
    the style carried in ID 17.  fontdb -- which is what resvg matches ``font-family``
    against -- indexes a face under ID 16 whenever the face has one, and never under ID 1.

    Measured on this machine rather than assumed: with only ``Aptos-Light.ttf`` loaded,
    resvg draws nothing for ``font-family="Aptos Light"`` and draws the Light face for
    ``font-family="Aptos"``.  64 of the 580 families installed here are unreachable by
    the name a deck would use, including ``Calibri Light`` and ``Yu Gothic Light``.

    So the emitted stack names the typographic family behind the exact one.  It is only
    ever consulted when the exact name fails to resolve, which is precisely the case this
    exists for; where the exact name works, this token is dead weight and costs nothing.

    The rule -- strip one trailing style word -- is derived from the fonts themselves, not
    guessed: of the installed faces whose ID 1 and ID 16 disagree, every one that a deck
    might name is exactly ``ID16 + " " + <style word>``.  (The rest are Noto's abbreviated
    script names, where ID 1 is a *contraction* of ID 16 rather than an extension of it;
    no suffix rule recovers those and they are left alone.)
    """
    if not font_family:
        return None
    head, _, tail = font_family.rpartition(" ")
    if not head or tail.lower() not in _TYPOGRAPHIC_STYLE_WORDS:
        return None
    return head


def _escape(name: str) -> str:
    return name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def font_family_value(
    fonts: list[str | None], mapping: dict[str, str] | None = None
) -> str | None:
    """Build an SVG ``font-family`` stack: originals, then substitutes, then generic.

    Ordering matters -- the original name goes first so a host that really has Calibri
    uses it and matches PowerPoint exactly; the substitute only kicks in when it is
    missing, and because it is metric-compatible the layout computed from our tables
    still holds.  Between the two sits :func:`typographic_family`, for the faces a
    rasteriser files under a name the deck never spells: without it ``Aptos Display``
    resolved to the *generic* on a machine that has Aptos installed, skipping every
    named face in the stack, and the title of ``table-test`` drew 10% wide.  The trailing generic keyword is not decoration: with
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
        if font.startswith("+"):
            # An unexpanded theme pointer (``+mn-cs``).  :func:`substitution_for` already
            # declines to map these; letting one into the emitted list is worse than
            # useless, because ``+`` is not a legal start for a CSS identifier and resvg
            # discards the entire ``font-family`` declaration rather than just the bad
            # token.  One stray pointer therefore costs every other face in the stack.
            continue
        seen.add(font)
        unique.append(font)

        typographic = typographic_family(font)
        if typographic and typographic not in seen:
            seen.add(typographic)
            unique.append(typographic)

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
