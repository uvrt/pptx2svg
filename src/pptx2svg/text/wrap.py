"""Line breaking for a paragraph of styled runs.

Wrapping happens across runs, not within them: a paragraph whose first half is bold and
second half italic is one continuous line of text, so tokens carry their own run
properties and the line's segments are re-merged by property identity afterwards.

Break opportunities follow the two scripts PowerPoint decks actually mix:

* **Latin** -- break at spaces; a word is atomic unless it alone exceeds the line, in
  which case it is split by character.
* **CJK** -- every character is its own break opportunity, because CJK text has no
  spaces.

``WRAP_TOLERANCE_RATIO`` used to be an allowance for the ``kern`` feature we did not
apply.  We apply it now -- see :mod:`pptx2svg.text.kerning` -- and the constant is a
floating-point epsilon, which is what **PowerPoint's own budget has no slack in it at
all** means once the measurement it was covering for is right.  See the constant.

Kerning reaches the wrap through :func:`_join_kern` rather than through the token widths
alone, and that is not an optimisation: a CJK paragraph gives every character its own
token, so a kern pair almost always straddles a token boundary and measuring tokens
separately would charge none of them.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .. import model as m
from .measure import DefaultTextMeasurer, TextMeasurer, is_cjk

DEFAULT_FONT_SIZE = 18.0

#: Slack allowed when deciding whether a token fits, as a fraction of the line width.
#:
#: **PowerPoint itself allows none.**  ``tools/make_cjk_wrap_probe.py`` sweeps a text box
#: whose width is ``k * size + delta`` over 53 slides, three box widths and five font
#: sizes, with one character repeated whose advance is exactly 1 em; every slide fits
#: ``k`` characters at ``delta = 0`` and ``k - 1`` at ``delta = -0.25 pt``.  The rule is
#: ``sum of advances <= width - lIns - rIns - marL``, inclusive, to a quarter of a point.
#:
#: **So this is now an epsilon and nothing else**, which is the whole of what modelling
#: ``kern`` bought here.  It was 0.02, then 0.005, and both were allowances for *our*
#: measurement error rather than models of PowerPoint's rule: PowerPoint applies the
#: face's OpenType ``kern`` feature and we did not, so every line we measured was wider
#: than the one it laid out -- by 0% to 0.684% over ``sample-cjk``'s fifteen lines, and by
#: 1.5% for a string of nothing but kerned pairs (キスキスキス), which is outside any
#: window a single constant could sit in.  :mod:`pptx2svg.text.kerning` removed the error
#: instead of budgeting for it.
#:
#: What is left is arithmetic, not typography.  A line of ``k`` full-width glyphs is a sum
#: of ``k`` floating-point advances and the box is one multiplication, and the two agree
#: only to a few parts in 10^15; at exactly zero the probe's own case -- a box of
#: ``k * size`` fitting ``k`` -- fails on the last bit of the mantissa.  ``1e-6`` of a
#: line is 0.0004 px on a 400 px box, nine orders of magnitude inside the narrowest thing
#: it could wrongly admit (a glyph is ~2.5% of a forty-character line), and five thousand
#: times smaller than the constant it replaces.
#:
#: The one caller still measuring wide is an **embedded** face:
#: :mod:`pptx2svg.fonts.embedded` builds its table from a cut-down sfnt reader that does
#: not parse GPOS, so its ``FontMetrics.kerning`` is ``None``.  Widening this constant is
#: not the fix for that -- the fix is a GPOS reader -- and no deck in the corpus embeds a
#: face that kerns.
WRAP_TOLERANCE_RATIO = 1e-6

#: Characters a line may not *begin* with (行頭禁則), and which therefore pull their
#: left-hand neighbour down with them when a break would land before one.
#:
#: Every member is inside :func:`~pptx2svg.text.measure.is_cjk`'s ranges, which is
#: deliberate and is tested: the ASCII members of the same class -- ``)``, ``.``, ``,`` --
#: would change where *Latin* wraps, and nothing here measured that.
NOT_LINE_START = frozenset(
    "、。，．・：；？！゛゜ゝゞヽヾ々ー゠"       # punctuation, iteration and length marks
    "ぁぃぅぇぉっゃゅょゎゕゖ"                     # small hiragana
    "ァィゥェォッャュョヮヵヶ"                     # small katakana
    "）］｝〕〉》」』】〗〙〟｠"               # closing brackets and quotes
)

#: Characters a line may not *end* with (行末禁則), which go down to the next line
#: instead.  Same containment rule as :data:`NOT_LINE_START`.
NOT_LINE_END = frozenset("（［｛〔〈《「『【〖〘〝｟")


@dataclass
class LineSegment:
    text: str
    properties: m.RunProperties


@dataclass
class WrappedLine:
    segments: list[LineSegment] = field(default_factory=list)


@dataclass
class _Token:
    text: str
    properties: m.RunProperties
    width: float
    #: May a line break happen *before* this token?
    breakable: bool
    force_break: bool = False


def _is_whitespace(code_point: int) -> bool:
    return code_point in (0x20, 0x09, 0x0A, 0x0D)


def _split_into_fragments(text: str) -> list[tuple[str, bool]]:
    """Split a run into break units: ``(fragment, may_break_before)``."""
    fragments: list[tuple[str, bool]] = []
    current = ""
    current_type: str | None = None

    for char in text:
        code_point = ord(char)
        if _is_whitespace(code_point):
            if current and current_type != "space":
                fragments.append((current, current_type == "cjk"))
                current = ""
            current_type = "space"
            current += char
        elif is_cjk(code_point):
            if current:
                fragments.append((current, current_type in ("cjk", "space")))
                current = ""
            # Every CJK character is its own break opportunity.
            fragments.append((char, True))
            current_type = "cjk"
        else:
            if current and current_type != "latin":
                fragments.append((current, current_type == "space"))
                current = ""
            current_type = "latin"
            current += char

    if current:
        fragments.append((current, current_type in ("space", "cjk")))
    return fragments


def _tokenize(
    runs: list[m.TextRun],
    default_font_size: float,
    font_scale: float,
    measurer: TextMeasurer,
) -> list[_Token]:
    tokens: list[_Token] = []
    is_first = True

    for run in runs:
        if not run.text:
            continue
        properties = run.properties
        font_size = (
            properties.font_size * font_scale if properties.font_size else default_font_size
        )

        # A run may contain literal newlines (from `a:br`); each forces a line break.
        parts = run.text.split("\n")
        for index, part in enumerate(parts):
            if index > 0:
                tokens.append(
                    _Token(
                        text="",
                        properties=properties,
                        width=0.0,
                        breakable=True,
                        force_break=True,
                    )
                )
                is_first = False
            if not part:
                continue
            for fragment, breakable in _split_into_fragments(part):
                width = measurer.measure_text_width(
                    fragment,
                    font_size,
                    properties.bold,
                    properties.font_family,
                    properties.font_family_ea,
                )
                tokens.append(
                    _Token(
                        text=fragment,
                        properties=properties,
                        width=width,
                        # Never break before the very first token of a paragraph.
                        breakable=False if is_first else breakable,
                    )
                )
                is_first = False

    return tokens


def _is_space_only(text: str) -> bool:
    return all(_is_whitespace(ord(char)) for char in text)


def _split_token_by_chars(
    token: _Token,
    available_width: float,
    default_font_size: float,
    font_scale: float,
    measurer: TextMeasurer,
) -> list[list[_Token]]:
    """Break a single over-long token (a URL, a long CJK-free word) by character."""
    lines: list[list[_Token]] = []
    current: list[_Token] = []
    current_width = 0.0
    properties = token.properties
    font_size = properties.font_size * font_scale if properties.font_size else default_font_size

    for char in token.text:
        char_width = measurer.measure_text_width(
            char, font_size, properties.bold, properties.font_family, properties.font_family_ea
        )
        join = (
            _kern_between(
                measurer,
                current[-1].text,
                char,
                font_size,
                properties.bold,
                properties.font_family,
                properties.font_family_ea,
            )
            if current
            else 0.0
        )
        if current_width + join + char_width > available_width and current:
            lines.append(current)
            current = []
            current_width = 0.0
            join = 0.0
        current.append(
            _Token(text=char, properties=properties, width=char_width, breakable=False)
        )
        current_width += join + char_width

    if current:
        lines.append(current)
    return lines


def _join_kern(
    previous: _Token | None,
    token: _Token,
    default_font_size: float,
    font_scale: float,
    measurer: TextMeasurer,
) -> float:
    """What the ``kern`` feature takes off the join between two adjacent tokens.

    Tokens are measured one at a time, and for CJK *every character is its own token*, so
    without this the Japanese kern pairs -- the ones that made this worth modelling --
    would all fall between tokens and none would be charged.

    The two tokens must come from the same run, tested by identity of their properties
    exactly as :func:`_merge_segments` tests it.  That is not a shortcut: the width this
    paragraph is finally *drawn* at is the sum of the merged segments' widths, so a join
    charged here that `_merge_segments` will not merge would make the line the layout
    fits and the line the renderer centres disagree.
    """
    if previous is None or previous.properties is not token.properties:
        return 0.0
    if not previous.text or not token.text:
        return 0.0
    properties = token.properties
    font_size = properties.font_size * font_scale if properties.font_size else default_font_size
    return _kern_between(
        measurer,
        previous.text,
        token.text,
        font_size,
        properties.bold,
        properties.font_family,
        properties.font_family_ea,
    )


def _kern_between(
    measurer: TextMeasurer,
    left: str,
    right: str,
    font_size: float,
    bold: bool,
    font_family: str | None,
    font_family_ea: str | None,
) -> float:
    """:meth:`TextMeasurer.kern_between`, or zero for a measurer that has no such method.

    ``TextMeasurer`` is a structural protocol and a public one: a caller may pass its own
    object, and one written before kerning was modelled answers widths perfectly well.
    Such a measurer simply does not kern, which is what it did before and what its own
    ``measure_text_width`` will keep doing -- the alternative is an ``AttributeError`` in
    the middle of a wrap.
    """
    method = getattr(measurer, "kern_between", None)
    if method is None:
        return 0.0
    return method(left, right, font_size, bold, font_family, font_family_ea)


def _line_width(
    tokens: list[_Token],
    default_font_size: float,
    font_scale: float,
    measurer: TextMeasurer,
) -> float:
    """A token run's width, kerning across the joins included."""
    total = 0.0
    previous: _Token | None = None
    for token in tokens:
        total += _join_kern(previous, token, default_font_size, font_scale, measurer)
        total += token.width
        previous = token
    return total


def _merge_segments(tokens: list[_Token]) -> list[LineSegment]:
    """Re-join adjacent tokens that share run properties into one tspan's worth of text."""
    segments: list[LineSegment] = []
    for token in tokens:
        if segments and segments[-1].properties is token.properties:
            segments[-1].text += token.text
            continue
        segments.append(LineSegment(text=token.text, properties=token.properties))
    return segments


def _trim_trailing_spaces(segments: list[LineSegment]) -> list[LineSegment]:
    while segments:
        last = segments[-1]
        trimmed = last.text.rstrip()
        if trimmed:
            if trimmed != last.text:
                segments = segments[:-1] + [LineSegment(text=trimmed, properties=last.properties)]
            return segments
        segments = segments[:-1]
    return segments


def _kinsoku_pushback(current: list[_Token], token: _Token) -> int:
    """How many of ``current``'s trailing tokens must go down with ``token``.

    Japanese forbids a line that *begins* with a closing bracket, a small kana, a
    sound mark or a full stop, and one that *ends* with an opening bracket.  PowerPoint
    honours both by moving the offending character's neighbour down with it -- 追い出し,
    push-out -- and the loop here is why one pass is not enough: two forbidden characters
    in a row, or an opening bracket that lands at the line end only because something was
    already pushed down, each need another turn.

    Measured rather than assumed, in three respects.  ``tools/make_cjk_wrap_probe.py``'s
    ``k`` family puts each class of character at the break in a box whose width is known
    to the quarter point: PowerPoint moved the neighbour down in all seven, including for
    two punctuation marks in a row.  Its ``q`` family then tried the two attributes that
    might have governed it: **``hangingPunct`` makes no difference at all** -- the
    alternative layout, where the character hangs past the right edge, is not what this
    PowerPoint does with either setting -- while **``eaLnBrk="0"`` turns the rule off**
    and lets 、 open a line.  ``eaLnBrk`` defaults to on and every master in this corpus
    writes it on, so the rule is applied unconditionally here; reading the attribute is
    what a deck that turns it off would need and no deck in the corpus does.

    A line is never emptied: pushing its last token down would move the problem rather
    than solve it, and PowerPoint does not do that either.
    """
    moved = 0
    while True:
        kept = len(current) - moved
        if kept <= 1:
            return 0
        head = current[kept].text if moved else token.text
        tail = current[kept - 1].text
        if head[:1] in NOT_LINE_START or tail[-1:] in NOT_LINE_END:
            moved += 1
            continue
        return moved


def _layout_tokens(
    tokens: list[_Token],
    available_width: float,
    default_font_size: float,
    font_scale: float,
    measurer: TextMeasurer,
) -> list[WrappedLine]:
    if not tokens:
        return [WrappedLine()]

    lines: list[WrappedLine] = []
    current: list[_Token] = []
    current_width = 0.0
    tolerance = available_width * WRAP_TOLERANCE_RATIO

    for token in tokens:
        if token.force_break:
            lines.append(WrappedLine(segments=_trim_trailing_spaces(_merge_segments(current))))
            current = []
            current_width = 0.0
            continue

        # The join is charged only while the token stays on this line; a token that goes
        # down starts a line, and a line's first token kerns against nothing.
        join = _join_kern(
            current[-1] if current else None, token, default_font_size, font_scale, measurer
        )
        if current_width + join + token.width <= available_width + tolerance:
            current.append(token)
            current_width += join + token.width
            continue

        if not current:
            # Nothing on the line and the token still does not fit: split it.
            if _is_space_only(token.text):
                continue
            chunks = _split_token_by_chars(
                token, available_width, default_font_size, font_scale, measurer
            )
            for index, chunk in enumerate(chunks):
                if index < len(chunks) - 1:
                    segments = _trim_trailing_spaces(_merge_segments(chunk))
                    if segments:
                        lines.append(WrappedLine(segments=segments))
                else:
                    current = chunk
                    current_width = _line_width(
                        chunk, default_font_size, font_scale, measurer
                    )
            continue

        # The line is full. Whether we may break here or not, the token moves down --
        # but a leading space on the new line is dropped.  Kinsoku may take one or more
        # of the line's own trailing tokens down with it; see :func:`_kinsoku_pushback`.
        moved = _kinsoku_pushback(current, token)
        carried = current[len(current) - moved:] if moved else []
        current = current[: len(current) - moved] if moved else current
        segments = _trim_trailing_spaces(_merge_segments(current))
        if segments:
            lines.append(WrappedLine(segments=segments))
        if token.breakable and _is_space_only(token.text):
            current = carried
        else:
            current = carried + [token]
        current_width = _line_width(current, default_font_size, font_scale, measurer)

    if current:
        segments = _trim_trailing_spaces(_merge_segments(current))
        if segments:
            lines.append(WrappedLine(segments=segments))

    return lines or [WrappedLine()]


def wrap_paragraph(
    paragraph: m.Paragraph,
    available_width: float,
    default_font_size: float = DEFAULT_FONT_SIZE,
    font_scale: float = 1.0,
    measurer: TextMeasurer | None = None,
) -> list[WrappedLine]:
    """Break one paragraph into lines that fit ``available_width`` pixels."""
    measurer = measurer or DefaultTextMeasurer()
    if not paragraph.runs or not any(run.text for run in paragraph.runs):
        return [WrappedLine()]

    safe_width = max(available_width, 1.0)
    tokens = _tokenize(paragraph.runs, default_font_size, font_scale, measurer)
    if not tokens:
        return [WrappedLine()]

    return _layout_tokens(tokens, safe_width, default_font_size, font_scale, measurer)
