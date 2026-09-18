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

#: How sharply `a:clrChange`'s mask falls away from the keyed colour.  Alpha is
#: ``1 - K * (|dR| + |dG| + |dB|)``, so K = 64 means the three channels together may
#: differ by 1/64 of full scale -- about four 8-bit steps shared between them -- before
#: the pixel stops being replaced.  `a:clrChange` matches exactly, so this is a guard
#: against a rasteriser's rounding rather than a tolerance: a PNG that stores the keyed
#: colour exactly still reads back exactly, and anything a designer would call a
#: different colour is already far outside it.
CLR_CHANGE_SHARPNESS = 64


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
    """Colour adjustments applied to a picture (``a:grayscl``, ``a:duotone``, ...).

    The primitives chain implicitly: a primitive with no ``in`` reads whatever the one
    before it produced.  ``a:clrChange`` goes first because it is a substitution on the
    original artwork -- every other effect here is an adjustment *of* the colours, and
    adjusting before substituting would key on colours the file never contained.
    """
    if effects is None:
        return ""

    primitives: list[str] = []

    if effects.clr_change is not None:
        # `a:clrChange` replaces one exact colour throughout the picture with another --
        # the recolouring behind PowerPoint's "Set Transparent Color" and its recolour
        # presets.  A per-channel `feComponentTransfer` cannot do it: whether a pixel is
        # the source colour is a fact about all three channels together, and a transfer
        # function sees one channel at a time.
        #
        # So the mask is built with arithmetic instead:
        #
        # 1. Two `feColorMatrix` passes subtract the source colour each way round.  A
        #    filter clamps to 0, so one pass keeps where the channel is above the key and
        #    the other keeps where it is below, and adding them gives |delta| per channel
        #    -- the absolute value a single linear matrix cannot express.
        # 2. `feComposite operator="arithmetic"` with k2=k3=1 performs that addition.
        # 3. A third matrix sets RGB to the replacement colour from its constant column
        #    and puts `1 - K*(dR+dG+dB)` in alpha, so alpha is 1 only where every channel
        #    matched and 0 a hair's breadth away.  `a:clrChange` is an exact-match effect,
        #    which is why K is large rather than a tolerance anyone tunes.
        # 4. Two composites clip that flat colour to the picture's own alpha and lay it
        #    over the original.
        #
        # **Checked against resvg rather than assumed**, the way `feDiffuseLighting` and
        # `feSpecularLighting` were: a strip of #ff0000, #00ff00 and #fa0505 through this
        # chain keyed on red comes back (0,0,255), (0,255,0) and (250,5,5) -- the exact
        # match replaced, the near miss and the unrelated colour untouched.
        source, target = effects.clr_change.clr_from, effects.clr_change.clr_to
        r0, g0, b0 = (_unit(source.hex, index) for index in range(3))
        primitives.append(
            f'<feColorMatrix type="matrix" in="SourceGraphic" result="clrChangeAbove" '
            f'values="1 0 0 0 {num(-r0)}  0 1 0 0 {num(-g0)}  0 0 1 0 {num(-b0)}  0 0 0 0 1"/>'
            f'<feColorMatrix type="matrix" in="SourceGraphic" result="clrChangeBelow" '
            f'values="-1 0 0 0 {num(r0)}  0 -1 0 0 {num(g0)}  0 0 -1 0 {num(b0)}  0 0 0 0 1"/>'
            '<feComposite in="clrChangeAbove" in2="clrChangeBelow" operator="arithmetic" '
            'k1="0" k2="1" k3="1" k4="0" result="clrChangeDistance"/>'
            f'<feColorMatrix type="matrix" in="clrChangeDistance" result="clrChangeMask" '
            f'values="0 0 0 0 {_channel(target.hex, 0)}  '
            f'0 0 0 0 {_channel(target.hex, 1)}  '
            f'0 0 0 0 {_channel(target.hex, 2)}  '
            f'{-CLR_CHANGE_SHARPNESS} {-CLR_CHANGE_SHARPNESS} {-CLR_CHANGE_SHARPNESS} 0 1"/>'
            '<feComposite in="clrChangeMask" in2="SourceGraphic" operator="in" '
            'result="clrChangeKeyed"/>'
            '<feComposite in="clrChangeKeyed" in2="SourceGraphic" operator="over"/>'
        )

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

    if effects.alpha is not None and effects.alpha < 1:
        # `a:alphaModFix` scales the picture's existing alpha rather than replacing it,
        # so a PNG's own transparency survives -- which is why this is a `linear` transfer
        # on the alpha channel and not an `opacity` attribute on the <image>.
        primitives.append(
            "<feComponentTransfer>"
            f'<feFuncA type="linear" slope="{num(effects.alpha)}" intercept="0"/>'
            "</feComponentTransfer>"
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
    """One channel of ``#rrggbb`` as a filter-ready 0..1 string."""
    return num(_unit(hex_color, index))


def _unit(hex_color: str, index: int) -> float:
    """One channel of ``#rrggbb`` as a 0..1 float."""
    value = hex_color.lstrip("#")
    return int(value[index * 2 : index * 2 + 2], 16) / 255
