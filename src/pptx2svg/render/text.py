"""Text body layout and SVG generation.

Output is one ``<text>`` element per text body, with a ``<tspan>`` per run fragment.
Line advance is expressed as ``dy`` on the first tspan of each line, which is why the
whole body has to be laid out before any of it is emitted: the ``y`` of the ``<text>``
element depends on the *total* height (for middle/bottom anchoring), and the height
depends on how the text wraps.

The order of operations, therefore:

1. swap width/height and rotate margins if the body is vertical;
2. pick the default font size and, for ``normAutofit``, shrink until the text fits;
3. wrap each paragraph and emit its tspans, accumulating line advances;
4. measure the total height and offset the ``<text>`` element's ``y`` for the anchor;
5. add the first line's ascender so ``y`` lands on the baseline, not the line top.

PowerPoint's own line metrics are font-dependent, so a line's height comes from the
tallest run on it, using the font's ``(ascender + |descender|) / unitsPerEm``.
"""

from __future__ import annotations

from dataclasses import dataclass, replace

from .. import model as m
from ..text.fontmap import font_family_value, synthesises_italic
from ..text.measure import is_cjk
from ..text.wrap import LineSegment, wrap_paragraph
from ..units import PX_PER_PT, emu_to_px, px_to_emu
from .context import RenderContext, escape_xml_attr, escape_xml_text, num

DEFAULT_LINE_SPACING = 1.0
DEFAULT_FONT_SIZE_PT = 18.0

#: The shear PowerPoint applies when it has to fake an italic, as ``dx/dy``.
#:
#: No Japanese face involved here has an italic cut -- not MS Gothic or MS Mincho inside
#: Office's .ttc files, not the Noto Sans JP we ship -- and resvg does not synthesise
#: obliques, so `font-style="italic"` on one of them is silently a no-op and the run
#: draws bolt upright.  PowerPoint slants it.
#:
#: Read straight out of the text matrix in PowerPoint's PDF export of `sample.pptx`,
#: whose slide 2 sets one Japanese run italic::
#:
#:     45.3125 / 133.3333 = 0.33984375
#:
#: and confirmed independently against the raster: rendering that run through this skew
#: and scoring it against PowerPoint's own pixels peaks at 0.34 (SSIM 0.87, against 0.23
#: upright), with 0.30 and 0.36 both clearly worse.  It is a steep slant -- 18.8 degrees,
#: where a designed italic is usually 10-15 -- which is why it was worth confirming twice.
#:
#: A shear leaves the advance alone, and so does PowerPoint: the pen origins either side
#: of that run are exactly 32.000 pt apart at 32 pt, the same as the upright runs around
#: it.  So this changes no line break.
SYNTHETIC_OBLIQUE_SHEAR = 0.33984

_VERTICAL_TYPES = frozenset({"vert", "eaVert", "wordArtVert", "mongolianVert"})

#: ``a:rPr@u`` -> the nearest ``text-decoration-style``.  OOXML distinguishes weights
#: that CSS does not ("heavy", "dashLongHeavy"), so those collapse onto the plain form.
UNDERLINE_STYLES = {
    "dbl": "double",
    "wavyDbl": "double",
    "wavy": "wavy",
    "wavyHeavy": "wavy",
    "dotted": "dotted",
    "dottedHeavy": "dotted",
    "dotDash": "dotted",
    "dotDashHeavy": "dotted",
    "dotDotDash": "dotted",
    "dotDotDashHeavy": "dotted",
    "dash": "dashed",
    "dashHeavy": "dashed",
    "dashLong": "dashed",
    "dashLongHeavy": "dashed",
}


@dataclass
class _Dimensions:
    width: float
    height: float
    margin_left: float
    margin_right: float
    margin_top: float
    margin_bottom: float


def _resolve_dimensions(
    body: m.BodyProperties, original_width: float, original_height: float
) -> _Dimensions:
    """Vertical text lays out in a rotated box, so swap the axes and the margins with it."""
    if body.vert in _VERTICAL_TYPES:
        # 90 degrees clockwise.
        return _Dimensions(
            width=original_height,
            height=original_width,
            margin_left=emu_to_px(body.margin_top),
            margin_right=emu_to_px(body.margin_bottom),
            margin_top=emu_to_px(body.margin_right),
            margin_bottom=emu_to_px(body.margin_left),
        )
    if body.vert == "vert270":
        # 90 degrees counter-clockwise.
        return _Dimensions(
            width=original_height,
            height=original_width,
            margin_left=emu_to_px(body.margin_bottom),
            margin_right=emu_to_px(body.margin_top),
            margin_top=emu_to_px(body.margin_left),
            margin_bottom=emu_to_px(body.margin_right),
        )
    return _Dimensions(
        width=original_width,
        height=original_height,
        margin_left=emu_to_px(body.margin_left),
        margin_right=emu_to_px(body.margin_right),
        margin_top=emu_to_px(body.margin_top),
        margin_bottom=emu_to_px(body.margin_bottom),
    )


def render_text_body(
    text_body: m.TextBody, transform: m.Transform, context: RenderContext
) -> str:
    """Render a text body, flowing it into columns when ``a:bodyPr@numCol`` asks for them."""
    body = text_body.body_properties
    if body.num_col > 1 and body.vert == "horz":
        return _render_columns(text_body, transform, context)
    return _render_column(text_body, transform, context)


def _render_columns(
    text_body: m.TextBody, transform: m.Transform, context: RenderContext
) -> str:
    """Fill each column to the body's height before starting the next.

    Columns are laid out by rendering the same body once per column with the side
    margins widened to leave only that column's slice exposed, which keeps every bit of
    wrapping, anchoring and bullet logic in one place instead of two.  Vertical text is
    excluded: its columns run along the other axis, and the margin trick would move them
    in the wrong direction.
    """
    body = text_body.body_properties
    paragraphs = text_body.paragraphs
    if not any(run.text for para in paragraphs for run in para.runs):
        return ""

    columns = max(1, body.num_col)
    dims = _resolve_dimensions(
        body, emu_to_px(transform.extent_width), emu_to_px(transform.extent_height)
    )
    column_width = (dims.width - dims.margin_left - dims.margin_right) / columns
    available = dims.height - dims.margin_top - dims.margin_bottom

    groups = _split_into_columns(
        paragraphs, columns, column_width, available, body, context
    )

    # One shared counter dict, so a numbered list carries on across a column break
    # rather than restarting.
    counters: dict[str, int] = {}
    parts: list[str] = []
    for index, group in enumerate(groups):
        if not group:
            continue
        column_body = m.TextBody(
            paragraphs=group,
            body_properties=replace(
                body,
                num_col=1,
                margin_left=body.margin_left + px_to_emu(index * column_width),
                margin_right=body.margin_right
                + px_to_emu((columns - 1 - index) * column_width),
            ),
        )
        parts.append(_render_column(column_body, transform, context, counters))
    return "".join(parts)


def _split_into_columns(
    paragraphs: list[m.Paragraph],
    columns: int,
    column_width: float,
    available_height: float,
    body: m.BodyProperties,
    context: RenderContext,
) -> list[list[m.Paragraph]]:
    """Assign whole paragraphs to columns, breaking when one would overflow.

    PowerPoint breaks mid-paragraph; we do not, because a paragraph is the unit the
    wrapper and the height estimator both work in.  For the body text that `numCol` is
    normally used on -- several short paragraphs -- the two agree.
    """
    default_font_size = _default_font_size(paragraphs)
    should_wrap = body.wrap != "none"
    heights = [
        _estimate_text_height(
            [paragraph], default_font_size, should_wrap, column_width,
            body.ln_spc_reduction, body.font_scale, context,
        )
        for paragraph in paragraphs
    ]

    groups: list[list[m.Paragraph]] = [[]]
    used = 0.0
    for paragraph, height in zip(paragraphs, heights):
        if groups[-1] and used + height > available_height and len(groups) < columns:
            groups.append([])
            used = 0.0
        groups[-1].append(paragraph)
        used += height
    while len(groups) < columns:
        groups.append([])
    return groups


def _render_column(
    text_body: m.TextBody,
    transform: m.Transform,
    context: RenderContext,
    counters: dict[str, int] | None = None,
) -> str:
    body = text_body.body_properties
    paragraphs = text_body.paragraphs

    if not any(run.text for para in paragraphs for run in para.runs):
        return ""

    original_width = emu_to_px(transform.extent_width)
    original_height = emu_to_px(transform.extent_height)
    dims = _resolve_dimensions(body, original_width, original_height)

    full_text_width = dims.width - dims.margin_left - dims.margin_right
    num_col = max(1, body.num_col)
    text_width = full_text_width / num_col if num_col > 1 else full_text_width

    default_font_size = _default_font_size(paragraphs)
    should_wrap = body.wrap != "none"

    font_scale = body.font_scale
    ln_spc_reduction = body.ln_spc_reduction

    if body.auto_fit == "normAutofit" and should_wrap:
        available_height = dims.height - dims.margin_top - dims.margin_bottom
        font_scale = _shrink_to_fit_scale(
            paragraphs,
            default_font_size,
            font_scale,
            ln_spc_reduction,
            text_width,
            available_height,
            context,
        )

    scaled_default_size = default_font_size * font_scale
    default_line_ratio = _default_line_height_ratio(paragraphs, context)
    default_ascender_ratio = _default_ascender_ratio(paragraphs, context)
    default_natural_height = scaled_default_size * default_line_ratio

    tspans: list[str] = []
    is_first_line = True
    auto_num_counters: dict[str, int] = {} if counters is None else counters
    previous_space_after = 0.0
    # SVG has no text background, so `a:highlight` is drawn as rectangles behind the
    # <text> element.  Their vertical position is the running sum of the `dy` advances,
    # which is why it is accumulated here rather than recovered afterwards.
    baseline = 0.0
    highlights: list[_Highlight] = []
    obliques: list[_Oblique] = []

    for paragraph in paragraphs:
        properties = paragraph.properties
        para_margin_left = emu_to_px(properties.margin_left or 0)
        para_indent = emu_to_px(properties.indent or 0)

        text_start_x = dims.margin_left + para_margin_left
        # `indent` is normally negative: it hangs the bullet left of the text.
        bullet_x = text_start_x + para_indent
        effective_text_width = text_width - para_margin_left

        bullet_text = _bullet_text(properties, auto_num_counters)
        x_pos, anchor = _alignment(
            properties.alignment, text_start_x, effective_text_width, dims.width, dims.margin_right
        )

        para_font_size = _paragraph_font_size(paragraph, default_font_size) * font_scale
        para_natural = _paragraph_natural_height(
            paragraph, default_font_size, font_scale, context, default_line_ratio
        )
        # PowerPoint *adds* the space after one paragraph to the space before the next;
        # it does not collapse them the way CSS margins or Word do.  Measured on
        # `real-college-template` slide 6: `spcAft` 6 pt and an inherited `spcBef` of 20%
        # of a 24.09 pt line (4.78 pt) come out as a 10.78 pt gap, not as 6 pt.
        paragraph_gap = previous_space_after + _spacing_px(
            properties.space_before, para_natural
        )

        if not any(run.text for run in paragraph.runs):
            # An empty paragraph is a line's worth of nothing, so it is as tall as a line
            # -- the font size alone, which this used to use, is short by the ascent the
            # face puts above its em.  Measured on `real-college-template` slide 7, whose
            # two text blocks are separated by one empty 20 pt Arial paragraph:
            # PowerPoint leaves 106 px between the blocks' line tops and the font size
            # gave 100.  The natural height gives 107.2, so the residual is 1.2 px rather
            # than 6.  `_height` used the natural height here already, which is also why
            # the drawn and measured heights used to disagree for a body with a blank line
            # in it.
            dy = _compute_dy(
                is_first_line,
                _line_height_px(paragraph, para_natural, ln_spc_reduction),
                paragraph_gap,
            )
            tspans.append(f'<tspan x="{num(x_pos)}" dy="{dy}" text-anchor="{anchor}"> </tspan>')
            baseline += float(dy)
            is_first_line = False
            previous_space_after = _spacing_px(properties.space_after, para_natural)
            continue

        if should_wrap:
            lines = wrap_paragraph(
                paragraph,
                effective_text_width,
                scaled_default_size,
                font_scale,
                context.measurer,
            )
            for line_index, line in enumerate(lines):
                line_gap = paragraph_gap if line_index == 0 else 0.0

                if not line.segments:
                    dy = _compute_dy(
                        is_first_line,
                        _line_height_px(paragraph, default_natural_height, ln_spc_reduction),
                        line_gap,
                    )
                    tspans.append(
                        f'<tspan x="{num(x_pos)}" dy="{dy}" text-anchor="{anchor}"> </tspan>'
                    )
                    baseline += float(dy)
                    is_first_line = False
                    continue

                natural_height = _line_natural_height(
                    line.segments, default_font_size, font_scale, context
                )
                dy = _compute_dy(
                    is_first_line,
                    _line_height_px(paragraph, natural_height, ln_spc_reduction),
                    line_gap,
                )

                if line_index == 0 and bullet_text:
                    line_font_size = (
                        _line_font_size(line.segments, default_font_size) * font_scale
                    )
                    first_segment = line.segments[0]
                    tspans.append(
                        f'<tspan x="{num(bullet_x)}" dy="{dy}" text-anchor="start" '
                        f'{_bullet_style_attrs(properties, line_font_size, first_segment, context)}>'
                        f"{escape_xml_text(bullet_text)}</tspan>"
                    )
                    # The bullet already advanced the line, so the text only sets x.
                    line_dy = ""
                else:
                    line_dy = dy
                tspans.extend(
                    _render_line(
                        line.segments, x_pos, anchor, line_dy, dims.margin_left,
                        properties, body.default_tab_size, default_font_size,
                        font_scale, context,
                        # The same baseline the highlights get: one line advance on from
                        # where the accumulator currently stands.
                        baseline + float(dy), obliques,
                    )
                )
                baseline += float(dy)
                highlights.extend(
                    _line_highlights(
                        line.segments, x_pos, anchor, baseline,
                        default_font_size, font_scale, context,
                    )
                )
                is_first_line = False
        else:
            natural_height = _line_natural_height(
                [LineSegment(run.text, run.properties) for run in paragraph.runs],
                default_font_size,
                font_scale,
                context,
            )
            dy = _compute_dy(
                is_first_line,
                _line_height_px(paragraph, natural_height, ln_spc_reduction),
                paragraph_gap,
            )
            if bullet_text:
                first_run = next((run for run in paragraph.runs if run.text), None)
                size = (
                    (first_run.properties.font_size if first_run and first_run.properties.font_size
                     else default_font_size)
                    * font_scale
                )
                segment = LineSegment(
                    text="",
                    properties=first_run.properties if first_run else m.RunProperties(),
                )
                tspans.append(
                    f'<tspan x="{num(bullet_x)}" dy="{dy}" text-anchor="start" '
                    f"{_bullet_style_attrs(properties, size, segment, context)}>"
                    f"{escape_xml_text(bullet_text)}</tspan>"
                )

            line_dy = "" if bullet_text else dy
            tspans.extend(
                _render_line(
                    [LineSegment(run.text, run.properties) for run in paragraph.runs if run.text],
                    x_pos, anchor, line_dy, dims.margin_left, properties,
                    body.default_tab_size, default_font_size, font_scale, context,
                    baseline + float(dy), obliques,
                )
            )
            baseline += float(dy)
            highlights.extend(
                _line_highlights(
                    [LineSegment(run.text, run.properties) for run in paragraph.runs if run.text],
                    x_pos, anchor, baseline, default_font_size, font_scale, context,
                )
            )
            is_first_line = False

        previous_space_after = _spacing_px(properties.space_after, para_natural)

    if not tspans and not obliques:
        # A body whose every run was detached for shearing still has text to draw.
        return ""

    # Vertical anchoring needs the total height, which is only known now.
    y_start = dims.margin_top
    total_height = _estimate_text_height(
        paragraphs,
        default_font_size,
        should_wrap,
        text_width,
        ln_spc_reduction,
        font_scale,
        context,
    )
    if body.anchor == "ctr":
        y_start = max(dims.margin_top, (dims.height - total_height) / 2)
    elif body.anchor == "b":
        y_start = max(dims.margin_top, dims.height - total_height - dims.margin_bottom)

    # `y` on <text> is the baseline, not the top of the line box.
    first_font_size = _paragraph_font_size(paragraphs[0], default_font_size) * font_scale
    y_start += _first_baseline_px(
        paragraphs[0], first_font_size, default_ascender_ratio, ln_spc_reduction, context
    )

    element = f'<text x="0" y="{num(y_start)}" xml:space="preserve">{"".join(tspans)}</text>'
    if obliques:
        # After the main <text>, not before: these are glyphs, not backgrounds, and a
        # sheared run leans into its neighbours' space by design.
        element += "".join(run.svg(y_start) for run in obliques)
    if highlights:
        # Behind the text, and in one go: a highlight run is a background, so it must not
        # paint over a neighbouring run's glyphs.
        element = "".join(rect.svg(y_start) for rect in highlights) + element

    if body.vert in _VERTICAL_TYPES:
        element = (
            f'<g transform="translate({num(original_width)}, 0) rotate(90)">{element}</g>'
        )
    elif body.vert == "vert270":
        element = (
            f'<g transform="translate(0, {num(original_height)}) rotate(-90)">{element}</g>'
        )

    if body.rotation:
        # `a:bodyPr@rot` turns the text within the shape without turning the shape, so
        # the pivot is the text box's centre and the rotation composes on top of any
        # vertical-text transform rather than replacing it.
        element = (
            f'<g transform="rotate({num(body.rotation)}, '
            f'{num(original_width / 2)}, {num(original_height / 2)})">{element}</g>'
        )
    return element


# --------------------------------------------------------------------------------------
# Tab stops
# --------------------------------------------------------------------------------------

#: ``a:tab@algn`` -> the SVG ``text-anchor`` that puts the text on the right side of the
#: stop.  A decimal tab lines the decimal point up with the stop; with no way to find
#: that point in SVG, right-aligning is much closer than left-aligning for the numbers
#: decimal tabs are used on.
TAB_ANCHORS = {"l": "start", "ctr": "middle", "r": "end", "dec": "end"}


def _next_tab_stop(
    x: float, origin: float, stops: list[m.TabStop], default_size: float
) -> tuple[float, str]:
    """Where a tab at ``x`` lands, and how the text after it is anchored.

    Explicit ``a:tabLst`` stops come first; past the last of them PowerPoint falls back
    to the implicit grid of ``a:bodyPr@defTabSz``.  Both are measured from the text
    body's left inset, not from the paragraph indent.
    """
    for stop in stops:
        position = origin + emu_to_px(stop.position)
        if position > x + 0.01:
            return position, TAB_ANCHORS.get(stop.alignment, "start")
    step = emu_to_px(default_size) or 1.0
    steps = int((x - origin) / step) + 1
    return origin + steps * step, "start"


def _split_on_tabs(segments: list[LineSegment]) -> list[LineSegment | None]:
    """Flatten a line into pieces, with ``None`` standing in for each tab character."""
    pieces: list[LineSegment | None] = []
    for segment in segments:
        if "\t" not in segment.text:
            pieces.append(segment)
            continue
        parts = segment.text.split("\t")
        for index, part in enumerate(parts):
            if index:
                pieces.append(None)
            if part:
                pieces.append(LineSegment(part, segment.properties))
    return pieces


def _leading(x: float, dy: str, anchor: str) -> str:
    gap = f'dy="{dy}" ' if dy else ""
    return f'x="{num(x)}" {gap}text-anchor="{anchor}" '


def _render_line(
    segments: list[LineSegment],
    x_pos: float,
    anchor: str,
    dy: str,
    origin: float,
    properties: m.ParagraphProperties,
    default_tab_size: float,
    default_font_size: float,
    font_scale: float,
    context: RenderContext,
    baseline: float = 0.0,
    obliques: list[_Oblique] | None = None,
) -> list[str]:
    """Emit one line's tspans, starting a new chunk at every tab stop and font change.

    A tab is not a character with a width -- it is a jump to the next stop -- so each
    piece after one gets its own absolute ``x``, which in SVG starts a fresh text chunk
    and lets a centre or right stop be expressed as that chunk's ``text-anchor``.

    Tabs are only honoured in left-aligned paragraphs.  In a centred or right-aligned
    one the line's own anchor already decides where the text sits, and a stop measured
    from the left inset would fight with it; PowerPoint effectively ignores them there
    too.

    **A font change ends a chunk too, and that is not a nicety.**  ``font-family`` is a
    per-character property in SVG, so a conforming renderer falls back per glyph; resvg
    does not.  It picks one face for a whole chunk, and if the requested family cannot
    cover every character in it, the chunk is drawn in resvg's *default* face -- not just
    the characters the family was missing.  Measured: ``Markdown`` at 42.667 px renders
    182 px wide under ``font-family="Calibri"``, matching PowerPoint's 181 px exactly;
    put ``から`` in the same chunk and the Latin is redrawn 202 px wide, byte-identical to
    the same string under ``sans-serif`` and under ``Noto Sans JP``.  One kana silently
    cost Calibri for the entire line.

    Every mixed-script line in ``sample.pptx`` was drawing that way, which is most of why
    it scored 0.03 SSIM while the Latin-only decks scored 0.95+: the glyphs were the
    wrong outlines at accumulating wrong offsets.  Giving the following tspan an explicit
    ``x`` restores it -- the Latin chunk then rasterises byte-identically to drawing it
    alone -- so the positions we already computed to wrap the line are also what places
    each run.

    A line that never changes face is left flowing, both because it costs nothing and
    because letting the rasteriser accumulate advances with the real font is *better*
    than trusting our tables when there is no reason not to.
    """
    pieces = _split_on_tabs(segments)
    honour_tabs = anchor == "start" and any(piece is None for piece in pieces)
    if obliques is None:
        obliques = []

    # Plan the whole line before emitting any of it: a chunk boundary needs to know how
    # far along the line it falls, and a centred line needs its total width first.
    planned: list[tuple[LineSegment, list[tuple[str | None, str, str, bool]]] | None] = [
        None if piece is None
        else (piece, _segment_tspans(piece, font_scale, context, default_font_size))
        for piece in pieces
    ]
    families = [
        family
        for entry in planned
        if entry is not None
        for family, _styles, _text, _oblique in entry[1]
    ]
    any_oblique = any(
        oblique
        for entry in planned
        if entry is not None
        for _family, _styles, _text, oblique in entry[1]
    )

    if len(set(families)) <= 1 and not honour_tabs and not any_oblique:
        out = []
        first = True
        for entry in planned:
            if entry is None:
                continue
            out.append(
                _render_segment(
                    entry[0], font_scale,
                    _leading(x_pos, dy, anchor) if first else "",
                    context, default_font_size,
                )
            )
            first = False
        return out

    if honour_tabs:
        left = x_pos
    else:
        # Anchoring chunks absolutely means resolving the line's own anchor ourselves,
        # the same way :func:`_line_highlights` has to.
        total = sum(
            _tspan_width(text, entry[0].properties, default_font_size, font_scale, context)
            for entry in planned
            if entry is not None
            for _family, _styles, text, _oblique in entry[1]
        )
        left = (
            x_pos - total / 2 if anchor == "middle"
            else x_pos - total if anchor == "end"
            else x_pos
        )

    stops = properties.tab_stops
    out = []
    # `cursor` tracks where the next glyph would land; `chunk_*` remember where the
    # current chunk began and how wide it has grown, because a right- or centre-anchored
    # chunk does not end where it started plus its width.
    cursor = chunk_start = left
    chunk_width = 0.0
    chunk_anchor = "start"
    previous_family: object = _NO_FAMILY
    # A tab stop owns the chunk it opens -- it is the thing that carries the stop's
    # alignment -- so its prefix is held here until a tspan consumes it, rather than
    # being rebuilt from the font-change rule below.
    pending: str | None = _leading(left, dy, anchor)
    for entry in planned:
        if entry is None:
            cursor = _chunk_end(chunk_start, chunk_width, chunk_anchor)
            stop, chunk_anchor = _next_tab_stop(cursor, origin, stops, default_tab_size)
            pending = f'x="{num(stop)}" text-anchor="{chunk_anchor}" '
            cursor = chunk_start = stop
            chunk_width = 0.0
            previous_family = _NO_FAMILY
            continue

        piece, tspans = entry
        rendered: list[str] = []
        for family, styles, text, oblique in tspans:
            width = _tspan_width(
                text, piece.properties, default_font_size, font_scale, context
            )
            if oblique and chunk_anchor == "start":
                # Detached into its own <text> sibling, because the shear has to live on
                # an element and a <tspan> is not one resvg will transform.  The run
                # leaves the flow, so whatever follows has to re-anchor: forcing
                # `previous_family` to the sentinel makes the next tspan open a chunk.
                obliques.append(
                    _Oblique(x=cursor, baseline=baseline, styles=styles, text=text)
                )
                pending, previous_family = None, _NO_FAMILY
                cursor += width
                chunk_width += width
                continue

            if pending is not None:
                prefix, pending = pending, None
                previous_family = family
            elif family != previous_family:
                previous_family = family
                # A centre- or right-anchored tab chunk is positioned by its own total
                # width, so a sub-chunk inside it has no absolute x to give.  Leave that
                # one flowing and accept resvg's fallback rather than move the text.
                prefix = (
                    f'x="{num(cursor)}" text-anchor="start" '
                    if chunk_anchor == "start" else ""
                )
            else:
                prefix = ""
            rendered.append(f"<tspan {prefix}{styles}>{escape_xml_text(text)}</tspan>")
            cursor += width
            chunk_width += width

        content = "".join(rendered)
        if content and piece.properties.hyperlink is not None:
            content = (
                f'<a href="{escape_xml_attr(piece.properties.hyperlink.url)}">{content}</a>'
            )
        if content:
            out.append(content)
    return out


#: Sentinel for "no chunk open yet", distinct from a real ``font-family`` of ``None``.
_NO_FAMILY = object()


def _tspan_width(
    text: str,
    properties: m.RunProperties,
    default_font_size: float,
    font_scale: float,
    context: RenderContext,
) -> float:
    return context.measurer.measure_text_width(
        text,
        (properties.font_size or default_font_size) * font_scale,
        properties.bold,
        properties.font_family,
        properties.font_family_ea,
    )


def _chunk_end(start: float, width: float, anchor: str) -> float:
    if anchor == "end":
        return start
    if anchor == "middle":
        return start + width / 2
    return start + width


# --------------------------------------------------------------------------------------
# Highlight
# --------------------------------------------------------------------------------------


@dataclass
class _Oblique:
    """An italic run in a face with no italic, drawn as a sheared ``<text>`` sibling.

    It cannot stay a ``<tspan>``: SVG 1.1 puts ``transform`` on container and graphics
    elements, not on ``tspan``, and resvg follows that to the letter -- a ``transform``
    on a ``tspan`` rasterises byte-identically to no transform at all, which was checked
    rather than assumed.  So the run leaves the parent ``<text>`` flow and is positioned
    absolutely instead, which is affordable only because we already know where every run
    on the line starts.
    """

    x: float
    #: distance from the <text> element's y down to this line's baseline
    baseline: float
    styles: str
    text: str

    def svg(self, y_start: float) -> str:
        y = y_start + self.baseline
        # Shear about this run's own origin, so the baseline stays put and only the
        # verticals lean; skewX alone would slide the whole run sideways by x * shear.
        transform = (
            f"translate({num(self.x)},{num(y)}) "
            f"skewX({num(-_shear_degrees())}) "
            f"translate({num(-self.x)},{num(-y)})"
        )
        return (
            f'<text x="{num(self.x)}" y="{num(y)}" transform="{transform}" '
            f'xml:space="preserve"><tspan {self.styles}>'
            f"{escape_xml_text(self.text)}</tspan></text>"
        )


def _shear_degrees() -> float:
    import math

    return math.degrees(math.atan(SYNTHETIC_OBLIQUE_SHEAR))


@dataclass
class _Highlight:
    """One run's highlight, positioned relative to the ``<text>`` element's own ``y``."""

    x: float
    #: distance from the <text> element's y down to this line's baseline
    baseline: float
    width: float
    ascent: float
    descent: float
    color: m.ResolvedColor

    def svg(self, y_start: float) -> str:
        y = y_start + self.baseline - self.ascent
        attrs = f'fill="{self.color.hex}"'
        if self.color.alpha < 1:
            attrs += f' fill-opacity="{num(self.color.alpha)}"'
        return (
            f'<rect x="{num(self.x)}" y="{num(y)}" width="{num(self.width)}" '
            f'height="{num(self.ascent + self.descent)}" {attrs}/>'
        )


def _line_highlights(
    segments: list[LineSegment],
    x_pos: float,
    anchor: str,
    baseline: float,
    default_font_size: float,
    font_scale: float,
    context: RenderContext,
) -> list[_Highlight]:
    """Rectangles for whichever runs on this line are highlighted.

    ``x_pos`` is the anchor point rather than the left edge, so a centred or
    right-aligned line has to be measured in full before any run's position is known.
    """
    if not any(segment.properties.highlight for segment in segments):
        return []

    widths = [_segment_width(segment, default_font_size, font_scale, context) for segment in segments]
    total = sum(widths)
    if anchor == "middle":
        left = x_pos - total / 2
    elif anchor == "end":
        left = x_pos - total
    else:
        left = x_pos

    rectangles: list[_Highlight] = []
    for segment, width in zip(segments, widths):
        highlight = segment.properties.highlight
        if highlight is not None and width > 0:
            size = (segment.properties.font_size or default_font_size) * font_scale * PX_PER_PT
            ascent = context.measurer.ascender_ratio(
                segment.properties.font_family, segment.properties.font_family_ea
            )
            line = context.measurer.line_height_ratio(
                segment.properties.font_family, segment.properties.font_family_ea
            )
            rectangles.append(
                _Highlight(
                    x=left,
                    baseline=baseline,
                    width=width,
                    ascent=size * ascent,
                    descent=size * max(0.0, line - ascent),
                    color=highlight,
                )
            )
        left += width
    return rectangles


def _segment_width(
    segment: LineSegment, default_font_size: float, font_scale: float, context: RenderContext
) -> float:
    properties = segment.properties
    return context.measurer.measure_text_width(
        segment.text,
        (properties.font_size or default_font_size) * font_scale,
        properties.bold,
        properties.font_family,
        properties.font_family_ea,
    )


# --------------------------------------------------------------------------------------
# Bullets
# --------------------------------------------------------------------------------------


def _bullet_text(
    properties: m.ParagraphProperties, counters: dict[str, int]
) -> str | None:
    bullet = properties.bullet
    if bullet is None or isinstance(bullet, m.NoBullet):
        return None
    if isinstance(bullet, m.CharBullet):
        return bullet.char
    if isinstance(bullet, m.AutoNumBullet):
        # Numbering restarts per scheme+level, matching PowerPoint's list behaviour.
        key = f"{bullet.scheme}-{properties.level}"
        counters[key] = counters.get(key, 0) + 1
        return format_auto_num(bullet.scheme, bullet.start_at + counters[key] - 1)
    return None


def format_auto_num(scheme: str, index: int) -> str:
    if scheme == "arabicPeriod":
        return f"{index}."
    if scheme == "arabicParenR":
        return f"{index})"
    if scheme == "arabicPlain":
        return str(index)
    if scheme == "romanUcPeriod":
        return f"{_to_roman(index)}."
    if scheme == "romanLcPeriod":
        return f"{_to_roman(index).lower()}."
    if scheme == "alphaUcPeriod":
        return f"{_to_alpha(index)}."
    if scheme == "alphaLcPeriod":
        return f"{_to_alpha(index).lower()}."
    if scheme == "alphaUcParenR":
        return f"{_to_alpha(index)})"
    if scheme == "alphaLcParenR":
        return f"{_to_alpha(index).lower()})"
    return f"{index}."


_ROMAN_NUMERALS = (
    (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
    (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
)


def _to_roman(number: int) -> str:
    result = ""
    remaining = number
    for value, symbol in _ROMAN_NUMERALS:
        while remaining >= value:
            result += symbol
            remaining -= value
    return result


def _to_alpha(number: int) -> str:
    """1 -> A, 26 -> Z, 27 -> AA."""
    result = ""
    remaining = number
    while remaining > 0:
        remaining -= 1
        result = chr(65 + remaining % 26) + result
        remaining //= 26
    return result


def _bullet_style_attrs(
    properties: m.ParagraphProperties,
    text_font_size_pt: float,
    first_segment: LineSegment,
    context: RenderContext,
) -> str:
    styles: list[str] = []

    if properties.bullet_size_pct is not None:
        size = text_font_size_pt * (properties.bullet_size_pct / 100000)
        styles.append(f'font-size="{num(size * PX_PER_PT)}"')
    elif text_font_size_pt:
        styles.append(f'font-size="{num(text_font_size_pt * PX_PER_PT)}"')

    # A bullet with no `buFont` inherits the first run's typeface.
    chain = [
        properties.bullet_font,
        first_segment.properties.font_family,
        first_segment.properties.font_family_ea,
    ]
    family = font_family_value(chain, context.font_mapping)
    if family:
        styles.append(f'font-family="{escape_xml_attr(family)}"')

    if properties.bullet_color is not None:
        styles.append(f'fill="{properties.bullet_color.hex}"')
        if properties.bullet_color.alpha < 1:
            styles.append(f'fill-opacity="{num(properties.bullet_color.alpha)}"')
    elif first_segment.properties.color is not None:
        styles.append(f'fill="{first_segment.properties.color.hex}"')

    return " ".join(styles)


# --------------------------------------------------------------------------------------
# Segments
# --------------------------------------------------------------------------------------


def _needs_script_split(properties: m.RunProperties) -> bool:
    """A run with distinct Latin and East Asian typefaces must be split per script."""
    return (
        properties.font_family is not None
        and properties.font_family_ea is not None
        and properties.font_family != properties.font_family_ea
    )


def _split_by_script(text: str) -> list[tuple[str, bool]]:
    parts: list[tuple[str, bool]] = []
    current = ""
    current_is_ea: bool | None = None

    for char in text:
        east_asian = is_cjk(ord(char))
        if current_is_ea is not None and east_asian != current_is_ea:
            parts.append((current, current_is_ea))
            current = ""
        current_is_ea = east_asian
        current += char

    if current and current_is_ea is not None:
        parts.append((current, current_is_ea))
    return parts


def _segment_tspans(
    segment: LineSegment,
    font_scale: float,
    context: RenderContext,
    default_font_size: float = DEFAULT_FONT_SIZE_PT,
) -> list[tuple[str | None, str, str, bool]]:
    """``(font-family, style attributes, text, needs oblique)`` per tspan.

    The family is returned alongside the attribute string it is already inside because
    :func:`_render_line` has to compare it against the next tspan's: an SVG text chunk
    that changes face part-way through is the one thing resvg cannot draw (see
    :func:`_render_line`), so the family is what decides where a chunk ends.

    The last field says this run is italic in a face that has no italic to draw, so the
    slant has to be sheared on rather than asked for.  It also ends a chunk, because the
    shear can only be carried by an element a ``<tspan>`` is not allowed to be.
    """
    properties = segment.properties
    if not _needs_script_split(properties):
        chains: list[tuple[list[str | None] | None, str]] = [(None, segment.text)]
    else:
        chains = []
        for part_text, east_asian in _split_by_script(segment.text):
            fonts = (
                [properties.font_family_ea, context.jpan_fallback_font, properties.font_family]
                if east_asian
                else [properties.font_family, properties.font_family_ea]
            ) + [properties.font_family_cs]
            chains.append((fonts, part_text))

    out: list[tuple[str | None, str, str, bool]] = []
    for fonts, part_text in chains:
        styles = _style_attrs(properties, font_scale, fonts, context, default_font_size)
        chain = fonts if fonts is not None else [
            properties.font_family, properties.font_family_ea, properties.font_family_cs
        ]
        # The face that will actually draw this run is the first name in the stack, which
        # is the first name in the chain that resolves.
        oblique = bool(properties.italic) and any(
            synthesises_italic(name) for name in chain if name
        )
        if oblique:
            # We are about to shear the upright face ourselves, so asking for an italic
            # as well would be a second slant on any host that turns out to have one.
            styles = styles.replace(' font-style="italic"', "")
        out.append(
            (font_family_value(chain, context.font_mapping), styles, part_text, oblique)
        )
    return out


def _render_segment(
    segment: LineSegment,
    font_scale: float,
    prefix: str,
    context: RenderContext,
    default_font_size: float = DEFAULT_FONT_SIZE_PT,
) -> str:
    content = "".join(
        f"<tspan {prefix if index == 0 else ''}{styles}>{escape_xml_text(text)}</tspan>"
        for index, (_family, styles, text, _oblique) in enumerate(
            _segment_tspans(segment, font_scale, context, default_font_size)
        )
    )
    if segment.properties.hyperlink is not None:
        return f'<a href="{escape_xml_attr(segment.properties.hyperlink.url)}">{content}</a>'
    return content


def _style_attrs(
    properties: m.RunProperties,
    font_scale: float,
    fonts: list[str | None] | None,
    context: RenderContext,
    default_font_size: float = DEFAULT_FONT_SIZE_PT,
) -> str:
    styles: list[str] = []

    # Always emitted, even when the run itself states no size.  Nothing in OOXML obliges
    # a run to state one and plenty of real decks state one nowhere -- every run in
    # `authoring-integration.pptx` reaches the renderer with `font_size=None`.  The size
    # the *layout* used in that case is `default_font_size`, and leaving the attribute
    # off did not mean "same as the layout": it meant the rasteriser drew at the CSS
    # initial value of 16 px while the line had been wrapped, centred and spaced for
    # 18 pt.  Measured on that deck, every glyph came out about 1.65x too small.
    #
    # Written as user units (px), not `pt`: resvg's presentation-attribute parser
    # rejects unit suffixes on font-size, and px is understood by every backend.
    size = properties.font_size or default_font_size
    if size:
        styles.append(f'font-size="{num(size * font_scale * PX_PER_PT)}"')

    # `a:cs` names the typeface for complex scripts -- Arabic, Hebrew, Thai, Devanagari.
    # There is no per-script selection to make here the way `_split_by_script` makes one
    # for East Asian text, because the run is not split on script ranges; putting it last
    # in the stack lets the renderer fall through to it for glyphs the Latin face lacks,
    # which is what a `font-family` list is for.
    chain = fonts if fonts is not None else [
        properties.font_family, properties.font_family_ea, properties.font_family_cs
    ]
    family = font_family_value(chain, context.font_mapping)
    if family:
        styles.append(f'font-family="{escape_xml_attr(family)}"')

    if properties.bold:
        styles.append('font-weight="bold"')
    if properties.italic:
        styles.append('font-style="italic"')

    if properties.color is not None:
        styles.append(f'fill="{properties.color.hex}"')
        if properties.color.alpha < 1:
            styles.append(f'fill-opacity="{num(properties.color.alpha)}"')

    decorations = []
    if properties.underline:
        decorations.append("underline")
    if properties.strikethrough:
        decorations.append("line-through")
    if decorations:
        styles.append(f'text-decoration="{" ".join(decorations)}"')
    line_style = UNDERLINE_STYLES.get(properties.underline_style or "")
    if properties.underline and line_style:
        # `text-decoration-style` is a presentation attribute, not a CSS rule, so this
        # stays within the inline-attributes-only constraint.  Renderers that do not
        # implement it ignore it and draw the plain rule, which is the right fallback.
        styles.append(f'text-decoration-style="{line_style}"')

    if properties.baseline > 0:
        styles.append('baseline-shift="super"')
    elif properties.baseline < 0:
        styles.append('baseline-shift="sub"')

    if properties.outline is not None:
        styles.append(f'stroke="{properties.outline.color.hex}"')
        styles.append(f'stroke-width="{num(emu_to_px(properties.outline.width))}"')
        if properties.outline.color.alpha < 1:
            styles.append(f'stroke-opacity="{num(properties.outline.color.alpha)}"')
        # Stroke first so the outline sits behind the glyph fill, as PowerPoint does.
        styles.append('paint-order="stroke"')

    return " ".join(styles)


# --------------------------------------------------------------------------------------
# Metrics
# --------------------------------------------------------------------------------------


def _alignment(
    alignment: str | None,
    margin_left: float,
    text_width: float,
    width: float,
    margin_right: float,
) -> tuple[float, str]:
    if alignment == "ctr":
        return margin_left + text_width / 2, "middle"
    if alignment == "r":
        return width - margin_right, "end"
    return margin_left, "start"


def _first_baseline_px(
    paragraph: m.Paragraph,
    font_size_pt: float,
    ascender_ratio: float,
    ln_spc_reduction: float,
    context: RenderContext,
) -> float:
    """How far below the top of the text the first baseline sits.

    PowerPoint has two rules here and switches between them at exactly 100% line
    spacing, which is odd enough to be worth spelling out.  Both were measured by
    exporting a top-anchored, zero-inset box holding a single "H" -- whose ink bottom is
    the baseline -- and reading the offset off the raster.

    * **At or below 100%** the baseline hangs off the *bottom* of the line box: it sits
      one font descent up from it, so the leftover space becomes leading above the text.
      Measured 14 pt for 14 pt Arial (descent 0.212 em), 14 pt for Times New Roman
      (0.216) and 13 pt for Calibri (0.269) -- and 12 pt for Arial at 90%, i.e. the same
      figure scaled by the spacing.  That descent-derived offset is ``ascender_ratio``.

    * **Above 100%** the font drops out entirely and the baseline lands at three
      quarters of the line box, whatever the face.  Arial and Calibri both measured
      19 pt at 150% despite differing descents, and 105/110/120/130/150/200% all fit
      ``0.75 * line box`` to within the whole point PowerPoint rounds the offset to.

    The two rules do not meet: going from 100% to 105% moves the baseline *up*, which
    looked like a bad measurement until the second rule explained it.

    ``a:spcPts`` needs no special case -- an absolute line height is just a ratio
    against the natural one, and the same branch picks it up.
    """
    natural_pt = font_size_pt * context.measurer.line_height_ratio(
        *_first_run_fonts(paragraph)
    )
    line_px = _line_height_px(paragraph, natural_pt, ln_spc_reduction)
    natural_px = natural_pt * PX_PER_PT
    if natural_px <= 0:
        return font_size_pt * ascender_ratio * PX_PER_PT
    factor = line_px / natural_px
    if factor > 1.0:
        return 0.75 * line_px
    return font_size_pt * ascender_ratio * factor * PX_PER_PT


def _first_run_fonts(paragraph: m.Paragraph) -> tuple[str | None, str | None]:
    for run in paragraph.runs:
        if run.properties.font_family or run.properties.font_family_ea:
            return run.properties.font_family, run.properties.font_family_ea
    return None, None


def _line_height_px(
    paragraph: m.Paragraph, natural_height_pt: float, ln_spc_reduction: float = 0.0
) -> float:
    """``a:lnSpc`` is either a fixed point size or a multiplier on the natural height."""
    line_spacing = paragraph.properties.line_spacing
    if isinstance(line_spacing, m.PointsSpacing):
        return (line_spacing.value / 100) * PX_PER_PT * (1 - ln_spc_reduction)
    factor = (
        max(0.5, line_spacing.value / 100000)
        if isinstance(line_spacing, m.PercentSpacing)
        else DEFAULT_LINE_SPACING
    )
    return natural_height_pt * PX_PER_PT * factor * (1 - ln_spc_reduction)


def _spacing_px(spacing: m.SpacingValue, natural_height_pt: float) -> float:
    """``a:spcBef`` / ``a:spcAft`` in pixels.

    ``a:spcPct`` is a percentage of the paragraph's **natural line height**, not of its
    font size.  Measured on ``real-college-template`` slide 6, whose bullets are 20 pt
    Arial with ``spcAft`` 6 pt and an inherited ``spcBef`` of 20%: PowerPoint steps
    62.0 px between line tops at 1280 px wide (128 px/in), our line height is 42.83 px
    (24.09 pt, Arial's 1.2045 ratio at 20 pt), so the two gaps are 10.78 pt together and
    the ``spcBef`` half is 4.78 pt.  That is 19.84% of the line height and 23.9% of the
    font size, so the base is the line.

    ``a:spcPts`` stays absolute -- it is hundredths of a point and says so.
    """
    if isinstance(spacing, m.PointsSpacing):
        return (spacing.value / 100) * PX_PER_PT
    return natural_height_pt * (spacing.value / 100000) * PX_PER_PT


def _compute_dy(is_first_line: bool, line_height_px: float, paragraph_gap_px: float) -> str:
    if is_first_line:
        return "0"
    return f"{line_height_px + paragraph_gap_px:.2f}"


def _default_font_size(paragraphs: list[m.Paragraph]) -> float:
    for paragraph in paragraphs:
        for run in paragraph.runs:
            if run.properties.font_size:
                return run.properties.font_size
    return DEFAULT_FONT_SIZE_PT


def _paragraph_font_size(paragraph: m.Paragraph, default_font_size: float) -> float:
    for run in paragraph.runs:
        if run.text and run.properties.font_size:
            return run.properties.font_size
    if paragraph.end_para_run_properties and paragraph.end_para_run_properties.font_size:
        return paragraph.end_para_run_properties.font_size
    return default_font_size


def _line_font_size(segments: list[LineSegment], default_font_size: float) -> float:
    for segment in segments:
        if segment.properties.font_size:
            return segment.properties.font_size
    return default_font_size


def _paragraph_natural_height(
    paragraph: m.Paragraph,
    default_font_size: float,
    font_scale: float,
    context: RenderContext,
    default_ratio: float,
) -> float:
    """A paragraph's single-line height in points, before ``a:lnSpc``.

    This is what ``a:spcPct`` is a percentage of -- see :func:`_spacing_px` -- so the
    drawing pass and the height measurement have to agree on it, which is why it is one
    function rather than the same four lines twice.
    """
    has_text = any(run.text for run in paragraph.runs)
    end = paragraph.end_para_run_properties
    if not has_text and end is not None and end.font_size:
        return end.font_size * font_scale * default_ratio
    height = _line_natural_height(
        [LineSegment(run.text, run.properties) for run in paragraph.runs],
        default_font_size,
        font_scale,
        context,
    )
    return height if height > 0 else default_font_size * font_scale * default_ratio


def _line_natural_height(
    segments: list[LineSegment],
    default_font_size: float,
    font_scale: float,
    context: RenderContext,
) -> float:
    """A line is as tall as its tallest run."""
    tallest = 0.0
    for segment in segments:
        font_size = (segment.properties.font_size or default_font_size) * font_scale
        ratio = context.measurer.line_height_ratio(
            segment.properties.font_family, segment.properties.font_family_ea
        )
        tallest = max(tallest, font_size * ratio)
    return tallest if tallest > 0 else default_font_size * font_scale * 1.2


def _default_line_height_ratio(paragraphs: list[m.Paragraph], context: RenderContext) -> float:
    for paragraph in paragraphs:
        for run in paragraph.runs:
            if run.properties.font_family or run.properties.font_family_ea:
                return context.measurer.line_height_ratio(
                    run.properties.font_family, run.properties.font_family_ea
                )
    return 1.2


def _default_ascender_ratio(paragraphs: list[m.Paragraph], context: RenderContext) -> float:
    for paragraph in paragraphs:
        for run in paragraph.runs:
            if run.properties.font_family or run.properties.font_family_ea:
                return context.measurer.ascender_ratio(
                    run.properties.font_family, run.properties.font_family_ea
                )
    return 1.0


def _estimate_text_height(
    paragraphs: list[m.Paragraph],
    default_font_size: float,
    should_wrap: bool,
    text_width: float,
    ln_spc_reduction: float,
    font_scale: float,
    context: RenderContext,
) -> float:
    total = 0.0
    default_ratio = _default_line_height_ratio(paragraphs, context)
    previous_space_after = 0.0
    scaled_default = default_font_size * font_scale

    for index, paragraph in enumerate(paragraphs):
        has_text = any(run.text for run in paragraph.runs)
        natural_height = _paragraph_natural_height(
            paragraph, default_font_size, font_scale, context, default_ratio
        )
        line_height = _line_height_px(paragraph, natural_height, ln_spc_reduction)

        if should_wrap and has_text:
            line_count = len(
                wrap_paragraph(
                    paragraph, text_width, scaled_default, font_scale, context.measurer
                )
            )
        else:
            line_count = 1

        total += line_count * line_height

        if index > 0:
            # Added, not collapsed -- see the same sum in :func:`_render_column`.
            total += previous_space_after + _spacing_px(
                paragraph.properties.space_before, natural_height
            )

        previous_space_after = _spacing_px(
            paragraph.properties.space_after, natural_height
        )

    return total


def _shrink_to_fit_scale(
    paragraphs: list[m.Paragraph],
    default_font_size: float,
    font_scale: float,
    ln_spc_reduction: float,
    text_width: float,
    available_height: float,
    context: RenderContext,
) -> float:
    """``normAutofit``: shrink text until it fits, converging in a few passes.

    The stored ``fontScale`` is PowerPoint's own answer, but it was computed against
    PowerPoint's font metrics; re-deriving it keeps text inside the box when our
    measurements differ.
    """
    if available_height <= 0:
        return font_scale

    min_scale = font_scale * 0.1
    scale = font_scale

    for _ in range(5):
        height = _estimate_text_height(
            paragraphs, default_font_size, True, text_width, ln_spc_reduction, scale, context
        )
        if height <= available_height:
            break
        scale = max(scale * (available_height / height), min_scale)
        if scale <= min_scale:
            break

    return scale


def compute_sp_autofit_height(
    text_body: m.TextBody, transform: m.Transform, context: RenderContext
) -> float | None:
    """``spAutofit``: grow the shape to fit its text.  ``None`` when it already fits."""
    body = text_body.body_properties
    paragraphs = text_body.paragraphs

    if not any(run.text for para in paragraphs for run in para.runs):
        return None

    dims = _resolve_dimensions(
        body, emu_to_px(transform.extent_width), emu_to_px(transform.extent_height)
    )
    full_text_width = dims.width - dims.margin_left - dims.margin_right
    num_col = max(1, body.num_col)
    text_width = full_text_width / num_col if num_col > 1 else full_text_width

    height = _estimate_text_height(
        paragraphs,
        _default_font_size(paragraphs),
        body.wrap != "none",
        text_width,
        0.0,
        1.0,
        context,
    )
    if num_col > 1 and body.vert == "horz":
        # The text now flows into `num_col` columns, so it needs roughly that fraction
        # of the height it would take in one.
        height /= num_col
    required = height + dims.margin_top + dims.margin_bottom
    if required <= dims.height:
        return None
    return px_to_emu(required)
