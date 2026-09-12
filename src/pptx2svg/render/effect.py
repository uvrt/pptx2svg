"""Shape effects as SVG filters.

DrawingML effects compose in a fixed order -- soft edge, glow, outer shadow, inner shadow
-- and each one consumes the previous stage's result, so the primitives are chained
through named ``result`` values rather than each reading ``SourceGraphic``.

The filter region is enlarged to 200% so blurs and offsets are not clipped, and
``color-interpolation-filters="sRGB"`` is set because the SVG default (linearRGB) makes
shadows visibly lighter than PowerPoint's.
"""

from __future__ import annotations

import math

from .. import model as m
from ..units import emu_to_px
from .context import RenderContext, num


def render_effects(effects: m.EffectList | None, context: RenderContext) -> str:
    """Return a ``filter="url(#...)"`` attribute, registering the filter in ``<defs>``."""
    if effects is None:
        return ""

    primitives: list[str] = []
    last_result = "SourceGraphic"

    if effects.soft_edge is not None:
        radius = emu_to_px(effects.soft_edge.radius)
        primitives.append(
            f'<feGaussianBlur in="SourceAlpha" stdDeviation="{num(radius)}" result="softEdgeMask"/>'
        )
        primitives.append(
            '<feComposite in="SourceGraphic" in2="softEdgeMask" operator="in" '
            'result="softEdgeResult"/>'
        )
        last_result = "softEdgeResult"

    if effects.glow is not None:
        radius = emu_to_px(effects.glow.radius)
        color = effects.glow.color
        if last_result != "SourceGraphic":
            # Take the alpha of the current stage to blur, not of the original graphic.
            primitives.append(
                f'<feColorMatrix in="{last_result}" type="matrix" '
                'values="0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 1 0" result="glowAlpha"/>'
            )
        blur_in = "SourceAlpha" if last_result == "SourceGraphic" else "glowAlpha"
        primitives.extend(
            [
                f'<feGaussianBlur in="{blur_in}" stdDeviation="{num(radius)}" result="glowBlur"/>',
                f'<feFlood flood-color="{color.hex}" flood-opacity="{num(color.alpha)}" '
                'result="glowColor"/>',
                '<feComposite in="glowColor" in2="glowBlur" operator="in" result="glowFinal"/>',
                '<feMerge result="glowMerge">',
                '<feMergeNode in="glowFinal"/>',
                f'<feMergeNode in="{last_result}"/>',
                "</feMerge>",
            ]
        )
        last_result = "glowMerge"

    if effects.outer_shadow is not None:
        shadow = effects.outer_shadow
        # DrawingML blurRad is a diameter-like radius; halve it for a Gaussian sigma.
        std_dev = emu_to_px(shadow.blur_radius) / 2
        distance = emu_to_px(shadow.distance)
        radians = math.radians(shadow.direction)
        dx = round(distance * math.cos(radians), 2)
        dy = round(distance * math.sin(radians), 2)
        primitives.extend(
            [
                f'<feGaussianBlur in="SourceAlpha" stdDeviation="{num(std_dev)}" '
                'result="shadowBlur"/>',
                f'<feOffset in="shadowBlur" dx="{num(dx)}" dy="{num(dy)}" result="shadowOffset"/>',
                f'<feFlood flood-color="{shadow.color.hex}" '
                f'flood-opacity="{num(shadow.color.alpha)}" result="shadowColor"/>',
                '<feComposite in="shadowColor" in2="shadowOffset" operator="in" '
                'result="shadowFinal"/>',
                '<feMerge result="outerShadowMerge">',
                '<feMergeNode in="shadowFinal"/>',
                f'<feMergeNode in="{last_result}"/>',
                "</feMerge>",
            ]
        )
        last_result = "outerShadowMerge"

    if effects.inner_shadow is not None:
        shadow = effects.inner_shadow
        std_dev = emu_to_px(shadow.blur_radius) / 2
        distance = emu_to_px(shadow.distance)
        radians = math.radians(shadow.direction)
        dx = round(distance * math.cos(radians), 2)
        dy = round(distance * math.sin(radians), 2)
        primitives.extend(
            [
                # Invert the alpha so the blur spreads inward from the shape's edge.
                '<feComponentTransfer in="SourceAlpha" result="innerShdwInverse">',
                '<feFuncA type="table" tableValues="1 0"/>',
                "</feComponentTransfer>",
                f'<feGaussianBlur in="innerShdwInverse" stdDeviation="{num(std_dev)}" '
                'result="innerShdwBlur"/>',
                f'<feOffset in="innerShdwBlur" dx="{num(dx)}" dy="{num(dy)}" '
                'result="innerShdwOffset"/>',
                f'<feFlood flood-color="{shadow.color.hex}" '
                f'flood-opacity="{num(shadow.color.alpha)}" result="innerShdwFill"/>',
                '<feComposite in="innerShdwFill" in2="innerShdwOffset" operator="in" '
                'result="innerShdwColored"/>',
                '<feComposite in="innerShdwColored" in2="SourceAlpha" operator="in" '
                'result="innerShdwClipped"/>',
                f'<feComposite in="innerShdwClipped" in2="{last_result}" operator="over"/>',
            ]
        )

    if not primitives:
        return ""

    filter_id = context.new_id("effect")
    context.add_def(
        f'<filter id="{filter_id}" x="-50%" y="-50%" width="200%" height="200%" '
        f'color-interpolation-filters="sRGB">{"".join(primitives)}</filter>'
    )
    return f'filter="url(#{filter_id})"'


def render_blip_effects(effects: m.BlipEffects | None, context: RenderContext) -> str:
    """Colour adjustments applied to a picture (``a:grayscl``, ``a:duotone``, ...)."""
    if effects is None:
        return ""

    primitives: list[str] = []

    if effects.grayscale:
        primitives.append(
            '<feColorMatrix type="matrix" values="'
            "0.2126 0.7152 0.0722 0 0  "
            "0.2126 0.7152 0.0722 0 0  "
            "0.2126 0.7152 0.0722 0 0  "
            '0 0 0 1 0"/>'
        )

    if effects.lum is not None:
        # brightness shifts the intercept, contrast scales the slope
        slope = 1 + effects.lum.contrast
        intercept = effects.lum.brightness
        primitives.append(
            '<feComponentTransfer>'
            f'<feFuncR type="linear" slope="{num(slope)}" intercept="{num(intercept)}"/>'
            f'<feFuncG type="linear" slope="{num(slope)}" intercept="{num(intercept)}"/>'
            f'<feFuncB type="linear" slope="{num(slope)}" intercept="{num(intercept)}"/>'
            "</feComponentTransfer>"
        )

    if effects.bi_level is not None:
        # Threshold to pure black/white.  A discrete transfer splits its input into
        # equal buckets, so the split point is set by how many of them map to black:
        # 20 buckets puts the threshold within 2.5% of wherever `a:biLevel@thresh` asks
        # for, which is finer than the effect itself is ever authored to.
        buckets = 20
        black = max(1, min(buckets - 1, round(effects.bi_level.threshold * buckets)))
        table = " ".join(["0"] * black + ["1"] * (buckets - black))
        primitives.append(
            '<feColorMatrix type="matrix" values="'
            "0.2126 0.7152 0.0722 0 0  "
            "0.2126 0.7152 0.0722 0 0  "
            "0.2126 0.7152 0.0722 0 0  "
            '0 0 0 1 0"/>'
            "<feComponentTransfer>"
            f'<feFuncR type="discrete" tableValues="{table}"/>'
            f'<feFuncG type="discrete" tableValues="{table}"/>'
            f'<feFuncB type="discrete" tableValues="{table}"/>'
            "</feComponentTransfer>"
        )

    if effects.duotone is not None:
        first, second = effects.duotone.color1, effects.duotone.color2
        # Map luminance onto the ramp between the two colours.
        primitives.append(
            '<feColorMatrix type="matrix" values="'
            "0.2126 0.7152 0.0722 0 0  "
            "0.2126 0.7152 0.0722 0 0  "
            "0.2126 0.7152 0.0722 0 0  "
            '0 0 0 1 0"/>'
            "<feComponentTransfer>"
            f'<feFuncR type="table" tableValues="{_channel(first.hex, 0)} {_channel(second.hex, 0)}"/>'
            f'<feFuncG type="table" tableValues="{_channel(first.hex, 1)} {_channel(second.hex, 1)}"/>'
            f'<feFuncB type="table" tableValues="{_channel(first.hex, 2)} {_channel(second.hex, 2)}"/>'
            "</feComponentTransfer>"
        )

    if effects.blur is not None:
        primitives.append(
            f'<feGaussianBlur stdDeviation="{num(emu_to_px(effects.blur.radius) / 2)}"/>'
        )

    if not primitives:
        return ""

    filter_id = context.new_id("blip")
    context.add_def(
        f'<filter id="{filter_id}" x="0%" y="0%" width="100%" height="100%" '
        f'color-interpolation-filters="sRGB">{"".join(primitives)}</filter>'
    )
    return f'filter="url(#{filter_id})"'


def _channel(hex_color: str, index: int) -> str:
    value = hex_color.lstrip("#")
    return num(int(value[index * 2 : index * 2 + 2], 16) / 255)
