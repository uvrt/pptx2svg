"""Mapping proprietary Office fonts to open replacements.

Two different questions get answered here, and they have different answers:

* *What should I measure with?*  -> :func:`metrics_for`, which returns the metric-
  compatible table (Calibri measures as Carlito).
* *What should the SVG ask the rasteriser for?* -> :func:`font_family_value`, which
  emits the original name first so a machine that really has Calibri uses it, then the
  substitutes, then a generic family.
"""

from __future__ import annotations

from .metrics import FONT_NAME_TO_METRICS, METRICS, METRICS_KEY_TO_FONT_NAME, FontMetrics

#: PPTX font family -> a font likely installed on a Linux/CI rendering host.
DEFAULT_FONT_MAPPING: dict[str, str] = {
    # Latin
    "Calibri": "Carlito",
    "Calibri Light": "Carlito",
    "Arial": "Arimo",
    "Times New Roman": "Tinos",
    "Courier New": "Cousine",
    "Cambria": "Caladea",
    # Japanese Gothic -> Noto Sans JP
    "メイリオ": "Noto Sans JP",
    "Meiryo": "Noto Sans JP",
    "游ゴシック": "Noto Sans JP",
    "Yu Gothic": "Noto Sans JP",
    "MS ゴシック": "Noto Sans JP",
    "MS Gothic": "Noto Sans JP",
    "MS Pゴシック": "Noto Sans JP",
    "MS PGothic": "Noto Sans JP",
    # Japanese Mincho -> Noto Serif CJK JP
    "MS 明朝": "Noto Serif CJK JP",
    "MS Mincho": "Noto Serif CJK JP",
    "MS P明朝": "Noto Serif CJK JP",
    "MS PMincho": "Noto Serif CJK JP",
    "游明朝": "Noto Serif CJK JP",
    "Yu Mincho": "Noto Serif CJK JP",
}

#: Names that read as serif faces, for the generic family at the end of a font stack.
_SERIF_HINTS = ("mincho", "明朝", "times new roman", "georgia", "cambria", "garamond", "book antiqua")


def create_font_mapping(user_mapping: dict[str, str] | None = None) -> dict[str, str]:
    mapping = dict(DEFAULT_FONT_MAPPING)
    if user_mapping:
        mapping.update(user_mapping)
    return mapping


def _normalize_full_width(value: str) -> str:
    """Some PPTX themes spell font names full-width (``ＭＳ Ｐゴシック``)."""
    return "".join(
        " " if ch == "　" else (chr(ord(ch) - 0xFEE0) if "！" <= ch <= "～" else ch)
        for ch in value
    )


def mapped_font(font_family: str | None, mapping: dict[str, str]) -> str | None:
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

    lowered = normalized.lower()
    for key, value in mapping.items():
        if _normalize_full_width(key).lower() == lowered:
            return value
    return None


def metrics_for(font_family: str | None) -> FontMetrics | None:
    """Metrics table for a PPTX font name, or ``None`` when we have no data for it."""
    if not font_family:
        return None
    key = FONT_NAME_TO_METRICS.get(font_family)
    if key is None:
        key = FONT_NAME_TO_METRICS.get(font_family.lower())
    if key is None:
        key = FONT_NAME_TO_METRICS.get(_normalize_full_width(font_family))
    return METRICS.get(key) if key else None


def metrics_fallback_font(font_family: str | None) -> str | None:
    """The real OSS font whose metrics we used, so the SVG can ask for it by name."""
    if not font_family:
        return None
    key = FONT_NAME_TO_METRICS.get(font_family) or FONT_NAME_TO_METRICS.get(font_family.lower())
    return METRICS_KEY_TO_FONT_NAME.get(key) if key else None


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

    Ordering matters -- the original name goes first so a host that actually has the
    font uses it and matches PowerPoint exactly; the substitutes only kick in when it is
    missing, and they are metric-compatible so the layout computed from our metrics
    still holds.
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
