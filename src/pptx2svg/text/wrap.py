"""Line breaking for a paragraph of styled runs.

Wrapping happens across runs, not within them: a paragraph whose first half is bold and
second half italic is one continuous line of text, so tokens carry their own run
properties and the line's segments are re-merged by property identity afterwards.

Break opportunities follow the two scripts PowerPoint decks actually mix:

* **Latin** -- break at spaces; a word is atomic unless it alone exceeds the line, in
  which case it is split by character.
* **CJK** -- every character is its own break opportunity, because CJK text has no
  spaces.

``WRAP_TOLERANCE_RATIO`` exists because the measurements are estimates.  A substitute
font that measures ~1% wide would otherwise push the last word of a line that fits in
PowerPoint onto a line of its own, and that error compounds down a text box.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .. import model as m
from .measure import DefaultTextMeasurer, TextMeasurer, is_cjk

DEFAULT_FONT_SIZE = 18.0

#: Slack allowed when deciding whether a token fits, as a fraction of the line width.
WRAP_TOLERANCE_RATIO = 0.02


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
        if current_width + char_width > available_width and current:
            lines.append(current)
            current = []
            current_width = 0.0
        current.append(
            _Token(text=char, properties=properties, width=char_width, breakable=False)
        )
        current_width += char_width

    if current:
        lines.append(current)
    return lines


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

        if current_width + token.width <= available_width + tolerance:
            current.append(token)
            current_width += token.width
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
                    current_width = sum(item.width for item in chunk)
            continue

        # The line is full. Whether we may break here or not, the token moves down --
        # but a leading space on the new line is dropped.
        segments = _trim_trailing_spaces(_merge_segments(current))
        if segments:
            lines.append(WrappedLine(segments=segments))
        if token.breakable and _is_space_only(token.text):
            current = []
            current_width = 0.0
        else:
            current = [token]
            current_width = token.width

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
