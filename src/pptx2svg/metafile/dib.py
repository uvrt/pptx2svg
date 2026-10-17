"""Device-independent bitmap -> PNG: moved to :mod:`ooxml_common.metafile.dib` (round 7).

Kept so that ``pptx2svg.metafile.dib`` still imports; the code lives in ooxml-common,
where the GDI player that both renderers use decodes bitmaps with it.
"""

from ooxml_common.metafile.dib import *  # noqa: F401,F403
from ooxml_common.metafile.dib import (  # noqa: F401
    BI_ALPHABITFIELDS,
    BI_BITFIELDS,
    BI_JPEG,
    BI_PNG,
    BI_RGB,
    BI_RLE4,
    BI_RLE8,
    MAX_PIXELS,
    PNG_COMPRESS_LEVEL,
    DecodedImage,
    DecodedPixels,
    decode_dib,
    dib_to_image,
    encode_png,
    split_packed_dib,
)
