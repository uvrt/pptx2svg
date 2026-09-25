"""Font metrics, text measurement and line breaking."""

# kerning and metrics moved to ooxml_common.text with fontmap and measure.  Importing
# them here keeps `pptx2svg.text.metrics` an attribute of this package without an
# import, as it was when fontmap imported it relatively.
from . import kerning, metrics  # noqa: F401
from .fontmap import DEFAULT_FONT_MAPPING, create_font_mapping, font_family_value
from .measure import DefaultTextMeasurer, FontToolsTextMeasurer, TextMeasurer, is_cjk
from .wrap import WrappedLine, wrap_paragraph

__all__ = [
    "DEFAULT_FONT_MAPPING",
    "DefaultTextMeasurer",
    "FontToolsTextMeasurer",
    "TextMeasurer",
    "WrappedLine",
    "create_font_mapping",
    "font_family_value",
    "is_cjk",
    "wrap_paragraph",
]
