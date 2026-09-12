"""Colour resolution: scheme lookup through the colour map, then DrawingML transforms.

Two indirections stack here.  ``a:schemeClr val="tx1"`` names a *colour map slot*, not a
theme colour; the slide master's ``p:clrMap`` says which theme entry (``dk1``, ``lt1``,
...) that slot points at, and a layout or slide may override the map.  Only then do the
transforms -- ``lumMod``, ``lumOff``, ``tint``, ``shade``, ``alpha`` -- apply, in the
order PowerPoint applies them.
"""

from __future__ import annotations

import colorsys

from ..model import ResolvedColor
from ..parse.source import SourceColor, SourceColorMap, SourceTheme

#: ``p:clrMap`` when a master does not declare one.
DEFAULT_COLOR_MAP: dict[str, str] = {
    "bg1": "lt1",
    "tx1": "dk1",
    "bg2": "lt2",
    "tx2": "dk2",
    "accent1": "accent1",
    "accent2": "accent2",
    "accent3": "accent3",
    "accent4": "accent4",
    "accent5": "accent5",
    "accent6": "accent6",
    "hlink": "hlink",
    "folHlink": "folHlink",
}

#: Office default theme, used when a deck's theme is missing or incomplete.
FALLBACK_SCHEME_COLORS: dict[str, str] = {
    "dk1": "#000000",
    "lt1": "#ffffff",
    "dk2": "#44546a",
    "lt2": "#e7e6e6",
    "accent1": "#4472c4",
    "accent2": "#ed7d31",
    "accent3": "#a5a5a5",
    "accent4": "#ffc000",
    "accent5": "#5b9bd5",
    "accent6": "#70ad47",
    "hlink": "#0563c1",
    "folHlink": "#954f72",
}

BLACK = ResolvedColor(hex="#000000", alpha=1.0)


class ColorContext:
    """Everything needed to turn a :class:`SourceColor` into a concrete colour."""

    __slots__ = ("theme", "color_map", "scheme")

    def __init__(self, theme: SourceTheme | None, color_map: dict[str, str]):
        self.theme = theme
        self.color_map = color_map
        self.scheme = build_color_scheme(theme, color_map)


def build_effective_color_map(
    master: SourceColorMap | None = None,
    layout_override: SourceColorMap | None = None,
    slide_override: SourceColorMap | None = None,
) -> dict[str, str]:
    """Layer the colour maps: schema default < master < layout override < slide override."""
    mapping = dict(DEFAULT_COLOR_MAP)
    for source in (master, layout_override, slide_override):
        if source is not None:
            mapping.update(source.mapping)
    return mapping


def build_color_scheme(theme: SourceTheme | None, color_map: dict[str, str]) -> dict[str, str]:
    """Flatten the theme's ``a:clrScheme`` into ``{name: "#rrggbb"}``."""
    colors = dict(FALLBACK_SCHEME_COLORS)
    if theme is None:
        return colors
    # A partial context is enough here: scheme entries are srgb/sysClr in practice, and a
    # scheme entry that references another slot is resolved through `colors` as it fills in.
    context = _BootstrapContext(theme, color_map, colors)
    for name, color in theme.color_scheme.items():
        resolved = resolve_color(context, color)
        if resolved is not None:
            colors[name] = resolved.hex
    return colors


class _BootstrapContext:
    """Stand-in context used while the scheme table is still being built."""

    __slots__ = ("theme", "color_map", "scheme")

    def __init__(self, theme, color_map, scheme):
        self.theme = theme
        self.color_map = color_map
        self.scheme = scheme


def resolve_color(
    context: ColorContext | _BootstrapContext,
    color: SourceColor | None,
    visited: frozenset[str] = frozenset(),
) -> ResolvedColor | None:
    if color is None:
        return None
    base = _resolve_base_hex(context, color, visited)
    if base is None:
        return None
    return _apply_transforms(base, color.transforms)


def resolve_color_or(
    context: ColorContext | _BootstrapContext,
    color: SourceColor | None,
    default: ResolvedColor = BLACK,
) -> ResolvedColor:
    return resolve_color(context, color) or default


def _resolve_base_hex(context, color: SourceColor, visited: frozenset[str]) -> str | None:
    if color.kind == "srgb":
        return _normalize_hex(color.hex)
    if color.kind == "system":
        return _normalize_hex(color.last_color or "000000")
    if color.kind == "scheme":
        mapped = context.color_map.get(color.scheme, color.scheme)
        if mapped in visited:
            # A colour map cycle (tx1 -> dk1 -> tx1); stop rather than recurse.
            return "#000000"
        scheme_color = context.theme.color_scheme.get(mapped) if context.theme else None
        if scheme_color is not None:
            resolved = resolve_color(context, scheme_color, visited | {mapped})
            return resolved.hex if resolved is not None else None
        return context.scheme.get(mapped) or FALLBACK_SCHEME_COLORS.get(mapped)
    return None


def _apply_transforms(initial_hex: str, transforms) -> ResolvedColor:
    hex_value = initial_hex
    alpha = 1.0
    kinds = {transform.kind for transform in transforms}

    for transform in transforms:
        kind = transform.kind
        if kind == "lumMod":
            # lumOff is a companion of lumMod; apply both in one pass.
            lum_off = next((t.value for t in transforms if t.kind == "lumOff"), 0)
            hex_value = _apply_luminance(hex_value, transform.value / 100000, lum_off / 100000)
        elif kind == "lumOff":
            if "lumMod" not in kinds:
                hex_value = _apply_luminance(hex_value, 1.0, transform.value / 100000)
        elif kind == "tint":
            hex_value = _apply_tint(hex_value, transform.value / 100000)
        elif kind == "shade":
            hex_value = _apply_shade(hex_value, transform.value / 100000)
        elif kind == "alpha":
            alpha = transform.value / 100000
        elif kind == "satMod":
            hex_value = _apply_saturation(hex_value, transform.value / 100000)

    return ResolvedColor(hex=hex_value, alpha=alpha)


def _normalize_hex(value: str) -> str:
    normalized = value.lstrip("#").lower()
    return f"#{normalized.rjust(6, '0')[:6]}"


def _hex_to_rgb(value: str) -> tuple[int, int, int]:
    normalized = value.lstrip("#")
    return (
        int(normalized[0:2], 16),
        int(normalized[2:4], 16),
        int(normalized[4:6], 16),
    )


def _rgb_to_hex(r: float, g: float, b: float) -> str:
    return "#" + "".join(f"{max(0, min(255, int(round(v)))):02x}" for v in (r, g, b))


def _apply_luminance(value: str, lum_mod: float, lum_off: float) -> str:
    r, g, b = _hex_to_rgb(value)
    hue, lum, sat = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    lum = max(0.0, min(1.0, lum * lum_mod + lum_off))
    red, green, blue = colorsys.hls_to_rgb(hue, lum, sat)
    return _rgb_to_hex(red * 255, green * 255, blue * 255)


def _apply_saturation(value: str, sat_mod: float) -> str:
    r, g, b = _hex_to_rgb(value)
    hue, lum, sat = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    sat = max(0.0, min(1.0, sat * sat_mod))
    red, green, blue = colorsys.hls_to_rgb(hue, lum, sat)
    return _rgb_to_hex(red * 255, green * 255, blue * 255)


def _srgb_to_linear(channel: int) -> float:
    value = channel / 255
    if value <= 0.04045:
        return value / 12.92
    return ((value + 0.055) / 1.055) ** 2.4


def _linear_to_srgb(value: float) -> float:
    value = max(0.0, min(1.0, value))
    if value <= 0.0031308:
        return value * 12.92 * 255
    return (1.055 * value ** (1 / 2.4) - 0.055) * 255


def _apply_tint(value: str, amount: float) -> str:
    """Keep ``amount`` of the colour and make up the rest with white.

    ECMA-376 defines tint as "a 10% tint is 10% of the input colour combined with 90%
    white" -- so the value is how much of the *original* survives, not how far it moves.
    PowerPoint does the blend in linear-light space, which is why a 40% tint of a mid
    blue comes out visibly paler than a naive sRGB interpolation predicts; verified
    swatch-by-swatch against PowerPoint's own PDF export.
    """
    return _rgb_to_hex(
        *(
            _linear_to_srgb(_srgb_to_linear(channel) * amount + (1 - amount))
            for channel in _hex_to_rgb(value)
        )
    )


def _apply_shade(value: str, amount: float) -> str:
    """Keep ``amount`` of the colour and make up the rest with black.

    The linear-light note on :func:`_apply_tint` applies here too.
    """
    return _rgb_to_hex(
        *(
            _linear_to_srgb(_srgb_to_linear(channel) * amount)
            for channel in _hex_to_rgb(value)
        )
    )
