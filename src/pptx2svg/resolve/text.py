"""Text resolution: the run- and paragraph-property inheritance cascade.

A run's font size can be specified in six different places.  PowerPoint looks for it, in
order, in:

1. the run's own ``a:rPr``;
2. the paragraph's ``a:pPr/a:defRPr``;
3. the shape's ``a:lstStyle`` entry for the paragraph's outline level;
4. the *layout* placeholder's ``a:lstStyle`` for that level;
5. the *master* placeholder's ``a:lstStyle`` for that level;
6. the master's ``p:txStyles`` (title / body / other, chosen by placeholder type);
7. the presentation's ``a:defaultTextStyle``.

Levels 4-7 are "inherited defaults" rather than the shape's own styling, and pptx-glimpse
draws a distinction there: underline/strike/baseline/highlight are *not* inherited from
them, only size, typeface and colour are.

**Bold and italic used to be on the wrong side of that line**, on pptx-glimpse's reading
that "making a master's body text bold does not bold every slide's body text".  Measured
against PowerPoint, it does.  ``real-college-template.pptx`` settles it twice over, in
the two places that can be told apart:

* its master's ``p:titleStyle`` says ``b="1"`` and nothing else in the deck does, and
  PowerPoint drew "List Title", "GraphTitle" and "Slide Title" -- three different layouts
  -- bold;
* its ``slideLayout10`` repeats ``b="1"`` in the title placeholder's own ``a:lstStyle``,
  and PowerPoint drew slide 2's "Presentation Title" bold although that slide says
  nothing.  Slide 3 uses the same layout and *does* say ``b="1"`` on the run, which is
  why only slide 2 could show the difference.

The export's ``/BaseFont`` list is the corroborating fact: it holds ``Arial-BoldMT`` and
the deck has not one explicit bold run outside slides 3, 5 and 9.

Underline, strike, baseline and highlight stay excluded: no deck here sets one at levels
4-7, so there is nothing to check a change against.
"""

from __future__ import annotations

import base64
from dataclasses import dataclass, replace
from typing import Sequence

from .. import model as m
from ..parse import source as s
from ..text.fontmap import east_asian_family
from ..units import ROTATION_UNIT
from .color import resolve_color

#: Properties inherited from every level of the chain, including the layout's and
#: master's.  ``bold`` and ``italic`` belong here -- see the module docstring for the
#: measurement that moved them.
_ALWAYS_INHERITED = (
    "font_size", "typeface", "typeface_ea", "typeface_cs", "color", "bold", "italic",
)
#: Properties inherited only from the shape's own paragraph/list style.
_DECORATIONS = (
    "underline", "underline_style", "strikethrough", "baseline",
    "highlight",
)

#: ``p:txStyles`` child chosen by placeholder type.
_TITLE_PLACEHOLDERS = frozenset({"title", "ctrTitle"})
_BODY_PLACEHOLDERS = frozenset({"body", "subTitle", "obj"})


@dataclass(frozen=True)
class _StyleEntry:
    style: s.SourceTextStyle
    #: False for inherited defaults, whose decorations must not leak into the run.
    include_decorations: bool


def resolve_text_body(
    context,
    text_body: s.SourceTextBody,
    inherited: Sequence[s.SourceTextBody | None],
    placeholder_type: str | None,
    extra_defaults: s.SourceRunProperties | None = None,
) -> m.TextBody:
    chain = _build_style_chain(
        context, text_body, inherited, placeholder_type, extra_defaults
    )

    # Whether this body is a plain text box: no placeholder to inherit a face from, and
    # no table style underneath it.  :func:`_theme_body_latin` needs it and is several
    # calls down with only the context to go on.  Saved and restored rather than just
    # set, because a table's cells resolve their bodies inside a shape that may itself
    # be one.
    outer_unstyled = getattr(context, "unstyled_text_box", False)
    context.unstyled_text_box = placeholder_type is None and extra_defaults is None

    # Body properties layer outward-in: master placeholder, then layout, then the shape.
    properties: s.SourceTextBodyProperties | None = None
    for body in reversed(list(inherited)):
        properties = _merge_body_properties(properties, body.properties if body else None)
    properties = _merge_body_properties(properties, text_body.properties)

    try:
        return m.TextBody(
            paragraphs=[
                _resolve_paragraph(context, paragraph, chain)
                for paragraph in text_body.paragraphs
            ],
            body_properties=_body_properties(properties),
        )
    finally:
        context.unstyled_text_box = outer_unstyled


def _merge_body_properties(
    inherited: s.SourceTextBodyProperties | None,
    local: s.SourceTextBodyProperties | None,
) -> s.SourceTextBodyProperties | None:
    if local is None:
        return inherited
    if inherited is None:
        merged = replace(local)
    else:
        merged = replace(inherited)
        for name, value in vars(local).items():
            if value is not None:
                setattr(merged, name, value)

    # Only normAutofit carries a scale; the other two modes reset it.
    if merged.auto_fit == "normAutofit":
        merged.font_scale = merged.font_scale if merged.font_scale is not None else 1.0
        merged.ln_spc_reduction = (
            merged.ln_spc_reduction if merged.ln_spc_reduction is not None else 0.0
        )
    elif merged.auto_fit in ("spAutofit", "noAutofit"):
        merged.font_scale = 1.0
        merged.ln_spc_reduction = 0.0
    return merged


def _body_properties(properties: s.SourceTextBodyProperties | None) -> m.BodyProperties:
    if properties is None:
        return m.BodyProperties()
    default = m.BodyProperties()
    return m.BodyProperties(
        anchor=properties.anchor or default.anchor,
        margin_left=_or(properties.margin_left, default.margin_left),
        margin_right=_or(properties.margin_right, default.margin_right),
        margin_top=_or(properties.margin_top, default.margin_top),
        margin_bottom=_or(properties.margin_bottom, default.margin_bottom),
        wrap=properties.wrap or default.wrap,
        auto_fit=properties.auto_fit or default.auto_fit,
        font_scale=_or(properties.font_scale, 1.0),
        ln_spc_reduction=_or(properties.ln_spc_reduction, 0.0),
        num_col=_or(properties.num_col, 1),
        vert=properties.vert or default.vert,
        rotation=(properties.rotation or 0.0) / ROTATION_UNIT,
        default_tab_size=_or(properties.default_tab_size, default.default_tab_size),
    )


def _build_style_chain(
    context,
    text_body: s.SourceTextBody,
    inherited: Sequence[s.SourceTextBody | None],
    placeholder_type: str | None,
    extra_defaults: s.SourceRunProperties | None = None,
) -> list[_StyleEntry]:
    chain: list[_StyleEntry] = []

    if text_body.list_style is not None:
        chain.append(_StyleEntry(text_body.list_style, include_decorations=True))

    for body in inherited:
        if body is not None and body.list_style is not None:
            chain.append(_StyleEntry(body.list_style, include_decorations=False))

    if extra_defaults is not None:
        # A table style's `a:tcTxStyle`.  It sits *above* the master's `p:otherStyle` and
        # the presentation default, not below them: a table cell is not a placeholder, and
        # the region that styles it is the more specific statement.  It used to be last,
        # which cost `real-college-template` slide 5 its header -- that deck's
        # "Medium Style 2 - Accent 1" puts `<a:schemeClr val="lt1"/>` on `a:firstRow` and
        # PowerPoint inks "Column A" white on the #C00000 band, while the master's
        # `otherStyle` reached the run first and made it #151515.  The `b="on"` from the
        # same element did land, because nothing above it says anything about weight,
        # which is why the region looked like it was working.
        #
        # Its decorations apply, unlike the other inherited layers': a header row styled
        # bold really does bold the cell's text, which is the whole point of the region.
        chain.append(
            _StyleEntry(
                s.SourceTextStyle(
                    default_paragraph=s.SourceParagraphProperties(
                        default_run_properties=extra_defaults
                    )
                ),
                include_decorations=True,
            )
        )

    master_style = _tx_style_for_placeholder(context.master, placeholder_type)
    if master_style is not None:
        chain.append(_StyleEntry(master_style, include_decorations=False))

    default_style = context.presentation.default_text_style
    if default_style is not None:
        chain.append(_StyleEntry(default_style, include_decorations=False))

    return chain


def _tx_style_for_placeholder(
    master: s.SourceSlideMaster | None, placeholder_type: str | None
) -> s.SourceTextStyle | None:
    if master is None:
        return None
    if placeholder_type is None:
        return master.other_style
    if placeholder_type in _TITLE_PLACEHOLDERS:
        return master.title_style
    if placeholder_type in _BODY_PLACEHOLDERS:
        return master.body_style
    return master.other_style


def _style_level(style: s.SourceTextStyle, level: int) -> s.SourceParagraphProperties | None:
    if 0 <= level < len(style.levels) and style.levels[level] is not None:
        return style.levels[level]
    return style.default_paragraph


def _resolve_paragraph(
    context, paragraph: s.SourceParagraph, chain: Sequence[_StyleEntry]
) -> m.Paragraph:
    level = paragraph.properties.level if paragraph.properties and paragraph.properties.level else 0
    level_entries = [
        (_style_level(entry.style, level), entry.include_decorations) for entry in chain
    ]
    level_properties = [properties for properties, _ in level_entries]

    properties = _resolve_paragraph_properties(context, paragraph.properties, level_properties)

    # The run-property fallback order: the paragraph's own defRPr (decorations included),
    # then each style level's defRPr with that level's decoration policy.
    run_defaults: list[tuple[s.SourceRunProperties | None, bool]] = [
        (paragraph.properties.default_run_properties if paragraph.properties else None, True)
    ]
    run_defaults.extend(
        (props.default_run_properties if props else None, include)
        for props, include in level_entries
    )

    runs = [
        m.TextRun(
            text=run.text,
            properties=_resolve_run_properties(context, run.properties, run_defaults),
        )
        for run in paragraph.runs
    ]

    end_properties = None
    if paragraph.end_para_run_properties is not None:
        end_properties = _resolve_run_properties(
            context, paragraph.end_para_run_properties, run_defaults
        )

    return m.Paragraph(runs=runs, properties=properties, end_para_run_properties=end_properties)


def _resolve_paragraph_properties(
    context,
    local: s.SourceParagraphProperties | None,
    inherited: Sequence[s.SourceParagraphProperties | None],
) -> m.ParagraphProperties:
    def pick(name: str):
        if local is not None and getattr(local, name) is not None:
            return getattr(local, name)
        for properties in inherited:
            if properties is not None and getattr(properties, name) is not None:
                return getattr(properties, name)
        return None

    bullet_color_source = pick("bullet_color")

    return m.ParagraphProperties(
        alignment=pick("align") or "l",
        line_spacing=pick("line_spacing"),
        space_before=pick("space_before") or m.PercentSpacing(0),
        space_after=pick("space_after") or m.PercentSpacing(0),
        level=(local.level if local and local.level is not None else 0),
        bullet=_resolve_bullet(context, pick("bullet")),
        bullet_font=pick("bullet_font"),
        bullet_color=resolve_color(context.colors, bullet_color_source),
        bullet_size_pct=pick("bullet_size_pct"),
        bullet_size_points=pick("bullet_size_points"),
        margin_left=pick("margin_left"),
        indent=pick("indent"),
        tab_stops=pick("tab_stops") or [],
    )


def _resolve_bullet(context, bullet: s.SourceBulletType | None) -> m.BulletType | None:
    """Turn a parsed bullet into a model one.  Only ``a:buBlip`` needs anything done.

    A picture bullet is a relationship id until here; every other spelling is already
    what the renderer wants.  The blip is loaded eagerly, for the same reason
    :class:`~pptx2svg.model.ImageElement` carries its bytes: the package is gone by the
    time anything draws.

    An unreadable blip resolves to *no bullet* rather than to a substituted character.
    A deck that asks for a picture and gets a black disc has been quietly told a lie
    about its own content; an absent bullet plus the warning is the honest answer.
    """
    if not isinstance(bullet, s.SourceBlipBullet):
        return bullet

    # Circular at module scope: view.py imports this file, at its own foot.
    from .view import SUPPORTED_IMAGE_MIME_TYPES, _load_media_bytes

    media = _load_media_bytes(context, bullet.relationship_id)
    if media is None:
        context.warn(
            "unresolved-bullet-image",
            "a paragraph has an a:buBlip picture bullet whose image part is missing or "
            "unreadable; the paragraph is drawn with no bullet at all",
        )
        return None
    payload, mime_type = media
    if mime_type not in SUPPORTED_IMAGE_MIME_TYPES:
        context.warn(
            "unresolved-bullet-image",
            f"a paragraph has an a:buBlip picture bullet of type {mime_type}, which is "
            "not embeddable; the paragraph is drawn with no bullet at all",
        )
        return None
    return m.BlipBullet(
        image_data=base64.b64encode(payload).decode("ascii"), mime_type=mime_type
    )


def _resolve_run_properties(
    context,
    local: s.SourceRunProperties | None,
    defaults: Sequence[tuple[s.SourceRunProperties | None, bool]],
) -> m.RunProperties:
    merged = replace(local) if local is not None else s.SourceRunProperties()

    for properties, include_decorations in defaults:
        if properties is None:
            continue
        for name in _ALWAYS_INHERITED:
            if getattr(merged, name) is None:
                setattr(merged, name, getattr(properties, name))
        if include_decorations:
            for name in _DECORATIONS:
                if getattr(merged, name) is None:
                    setattr(merged, name, getattr(properties, name))
        if merged.outline_width is None and properties.outline_width is not None:
            merged.outline_width = properties.outline_width
            merged.outline_color = properties.outline_color

    outline = None
    if merged.outline_width:
        outline_color = resolve_color(context.colors, merged.outline_color)
        if outline_color is not None:
            outline = m.TextOutline(width=merged.outline_width, color=outline_color)

    return m.RunProperties(
        font_size=merged.font_size,
        font_family=(
            _resolve_typeface(context, merged.typeface) or _theme_body_latin(context)
        ),
        font_family_ea=east_asian_family(
            _resolve_typeface(context, merged.typeface_ea), _theme_east_asian(context)
        ),
        font_family_cs=_resolve_typeface(context, merged.typeface_cs),
        bold=bool(merged.bold),
        italic=bool(merged.italic),
        underline=bool(merged.underline),
        underline_style=merged.underline_style if merged.underline else None,
        strikethrough=bool(merged.strikethrough),
        color=resolve_color(context.colors, merged.color),
        baseline=merged.baseline or 0.0,
        hyperlink=_hyperlink(context, merged),
        outline=outline,
        highlight=resolve_color(context.colors, merged.highlight),
    )


def _theme_east_asian(context) -> str | None:
    """The face this theme draws East Asian text in, for a run that names none.

    The counterpart of :func:`_theme_body_latin`, and the piece that was missing: a run
    stating no ``<a:ea>`` used to reach the measurer with ``font_family_ea=None``, so
    every kana and every ideograph in it was measured through the *Latin* table.  That is
    not a small approximation -- a Latin face has no East Asian glyph at all, so the
    width came from :attr:`FontMetrics.cjk_width`, which is one em in every table because
    the extractor writes ``units_per_em`` where the probe kanji is missing.  Layout was
    computed from a constant while the rasteriser drew with a real face.

    **Body before heading, measured.**  ``real-financial-report.pptx``'s theme offers
    ``游ゴシック Light`` as its major Jpan face and ``游ゴシック`` as its minor, and
    PowerPoint's own export drew the chart's Japanese in **YuGothic-Regular** -- the
    minor.  The Latin side already makes the same choice for the same reason (see
    :func:`_theme_body_latin`): a master that wants the heading face says so with
    ``+mj-ea``, and guessing it for everything else gets a light weight on body copy.

    Within a collection, ``<a:ea>`` comes before the ``<a:font script="..."/>`` list;
    :func:`pptx2svg.text.fontmap.east_asian_family` is what enforces that an empty
    ``typeface=""`` -- which is what every theme in this corpus writes -- is not a name.
    """
    scheme = context.theme.font_scheme if context.theme else None
    if scheme is None:
        return None
    return east_asian_family(
        scheme.minor_east_asian,
        scheme.minor_japanese,
        scheme.major_east_asian,
        scheme.major_japanese,
    )


def _theme_body_latin(context) -> str | None:
    """The theme's body face, for a run that named none anywhere in the cascade.

    Nothing in OOXML obliges a run, a placeholder, a layout or a master to state a
    typeface, and plenty of real decks state one nowhere: five of the seven corpus
    fixtures reached the renderer with ``font_family=None`` on most of their runs.  That
    is not "no font" -- PowerPoint draws those in the theme's minor (body) face, which is
    exactly what ``+mn-lt`` points at.  Leaving it None had two costs, and the second is
    the expensive one:

    * the SVG carried no ``font-family``, so the rasteriser drew its own default;
    * :func:`pptx2svg.text.fontmap.metrics_for` had nothing to look up, so every string
      was measured with the crude per-category ratios in :mod:`pptx2svg.text.measure`
      (0.6 em for a typical glyph) instead of real advance widths.

    So wrapping, autofit and centring were computed for a font nobody named, and then
    drawn in a font nobody chose.  The major (heading) face is deliberately not used
    here: a master that wants it says ``+mj-lt`` in its ``titleStyle``, and guessing
    "this looks like a title" would be a second, worse heuristic on top of this one.

    The theme is not the answer everywhere, though -- see the comment in the body for the
    one place PowerPoint measurably does something else.
    """
    if getattr(context, "unstyled_text_box", False):
        # ...except in a plain text box, where PowerPoint does not consult the theme at
        # all.  Measured on `authoring-integration`: every text box on its slide, its
        # master and its layout is `txBox="1"` with no `rPr`, no `lstStyle`, an empty
        # `p:txStyles` and an empty `defaultTextStyle`, and PowerPoint's export draws all
        # of them in Arial -- /BaseFont says ArialMT, and the drawn advance for "MASTER
        # CONTRACT" is 178.28 pt against Arial's 180.00 pt at 18 pt, ink against advance.
        # In the same export that deck's table cells are Aptos-Bold and its chart labels
        # Aptos, so this is not PowerPoint lacking the theme face: it is a text box with
        # nothing specified anywhere falling back to the application default instead of
        # to the theme.
        #
        # It costs more than two headings looking slightly wrong.  Arial is wider than
        # Aptos, so PowerPoint wraps "LAYOUT CONTRACT" onto two lines in a box we fitted
        # it into on one, and a wrap that disagrees moves every line under it.
        #
        # Placeholders and table cells are deliberately excluded: a placeholder takes its
        # face from the master's `titleStyle`/`bodyStyle` and a cell from the table
        # style's `a:tcTxStyle`, both of which normally do point at the theme.
        return "Arial"
    scheme = context.theme.font_scheme if context.theme else None
    if scheme is None:
        return None
    return _non_empty(scheme.minor_latin) or _non_empty(scheme.major_latin)


def _resolve_typeface(context, typeface: str | None) -> str | None:
    """Expand the theme font placeholders ``+mj-lt`` / ``+mn-ea`` / ``+mn-cs`` etc.

    A placeholder that cannot be expanded resolves to ``None``, not to itself.  A theme
    that writes ``<a:cs typeface=""/>`` -- which is most of them, and every deck in this
    corpus -- is saying it names no complex-script face, so the correct answer is "this
    run states no typeface for that script" and the cascade should carry on as if the
    attribute were absent.  Returning the pointer instead put the literal string
    ``+mn-cs`` into the model, and from there into the SVG's ``font-family`` list, where
    it did far more damage than an unused name: ``+`` cannot start a CSS identifier, so
    resvg rejected the *whole* declaration and fell back to its default family.  Every
    Japanese glyph in ``sample.pptx`` was drawn by resvg's fallback instead of the
    ＭＳ Ｐゴシック the stack asked for, which is measurable -- "テンプレート" at 100 px
    inks 563 px wide through the poisoned stack and 503 px with the same stack minus the
    pointer.
    """
    if typeface is None:
        return None
    scheme = context.theme.font_scheme if context.theme else None
    if scheme is None:
        # No theme to expand against.  A pointer is still not a face name.
        return None if typeface.startswith("+") else typeface

    if typeface == "+mj-lt":
        return _non_empty(scheme.major_latin)
    if typeface == "+mn-lt":
        return _non_empty(scheme.minor_latin)
    if typeface == "+mj-ea":
        return _non_empty(scheme.major_east_asian) or _non_empty(scheme.major_japanese)
    if typeface == "+mn-ea":
        return _non_empty(scheme.minor_east_asian) or _non_empty(scheme.minor_japanese)
    if typeface == "+mj-cs":
        return _non_empty(scheme.major_complex_script)
    if typeface == "+mn-cs":
        return _non_empty(scheme.minor_complex_script)
    return typeface


def _hyperlink(context, properties: s.SourceRunProperties) -> m.Hyperlink | None:
    if properties.hyperlink_rel_id is None:
        return None
    relationship = context.package.relationships(context.part_path).get(properties.hyperlink_rel_id)
    if relationship is None:
        return None
    return m.Hyperlink(url=relationship.target, tooltip=properties.hyperlink_tooltip)


def _non_empty(value: str | None) -> str | None:
    return value or None


def _or(value, fallback):
    return fallback if value is None else value
