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
draws a distinction there that matters: bold/italic/underline/strike/baseline are
*not* inherited from them, only size, typeface and colour are.  That mirrors PowerPoint,
where making a master's body text bold does not bold every slide's body text.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Sequence

from .. import model as m
from ..parse import source as s
from .color import resolve_color

#: Properties inherited from every level of the chain.
_ALWAYS_INHERITED = ("font_size", "typeface", "typeface_ea", "typeface_cs", "color")
#: Properties inherited only from the shape's own paragraph/list style.
_DECORATIONS = ("bold", "italic", "underline", "strikethrough", "baseline", "highlight")

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
) -> m.TextBody:
    chain = _build_style_chain(context, text_body, inherited, placeholder_type)

    # Body properties layer outward-in: master placeholder, then layout, then the shape.
    properties: s.SourceTextBodyProperties | None = None
    for body in reversed(list(inherited)):
        properties = _merge_body_properties(properties, body.properties if body else None)
    properties = _merge_body_properties(properties, text_body.properties)

    return m.TextBody(
        paragraphs=[
            _resolve_paragraph(context, paragraph, chain) for paragraph in text_body.paragraphs
        ],
        body_properties=_body_properties(properties),
    )


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
    )


def _build_style_chain(
    context,
    text_body: s.SourceTextBody,
    inherited: Sequence[s.SourceTextBody | None],
    placeholder_type: str | None,
) -> list[_StyleEntry]:
    chain: list[_StyleEntry] = []

    if text_body.list_style is not None:
        chain.append(_StyleEntry(text_body.list_style, include_decorations=True))

    for body in inherited:
        if body is not None and body.list_style is not None:
            chain.append(_StyleEntry(body.list_style, include_decorations=False))

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
        bullet=pick("bullet"),
        bullet_font=pick("bullet_font"),
        bullet_color=resolve_color(context.colors, bullet_color_source),
        bullet_size_pct=pick("bullet_size_pct"),
        margin_left=pick("margin_left"),
        indent=pick("indent"),
        tab_stops=pick("tab_stops") or [],
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
        font_family=_resolve_typeface(context, merged.typeface),
        font_family_ea=_resolve_typeface(context, merged.typeface_ea),
        font_family_cs=_resolve_typeface(context, merged.typeface_cs),
        bold=bool(merged.bold),
        italic=bool(merged.italic),
        underline=bool(merged.underline),
        strikethrough=bool(merged.strikethrough),
        color=resolve_color(context.colors, merged.color),
        baseline=merged.baseline or 0.0,
        hyperlink=_hyperlink(context, merged),
        outline=outline,
        highlight=resolve_color(context.colors, merged.highlight),
    )


def _resolve_typeface(context, typeface: str | None) -> str | None:
    """Expand the theme font placeholders ``+mj-lt`` / ``+mn-ea`` / ``+mn-cs`` etc."""
    if typeface is None:
        return None
    scheme = context.theme.font_scheme if context.theme else None
    if scheme is None:
        return typeface

    if typeface == "+mj-lt":
        return scheme.major_latin or typeface
    if typeface == "+mn-lt":
        return scheme.minor_latin or typeface
    if typeface == "+mj-ea":
        return _non_empty(scheme.major_east_asian) or scheme.major_japanese or typeface
    if typeface == "+mn-ea":
        return _non_empty(scheme.minor_east_asian) or scheme.minor_japanese or typeface
    if typeface == "+mj-cs":
        return _non_empty(scheme.major_complex_script) or typeface
    if typeface == "+mn-cs":
        return _non_empty(scheme.minor_complex_script) or typeface
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
