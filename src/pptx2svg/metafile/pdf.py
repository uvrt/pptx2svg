"""Rasterise an EMF's embedded PDF preview.

This is the one part of metafile support that cannot be done with the standard library:
a PDF page has to be *rendered*, and nothing in the stdlib does that.  The core of this
library is deliberately dependency-free, so the import is local to the call, the extra is
opt-in (``pip install pptx2svg[metafile]``), and the absence of it is a
:class:`PdfRasterizerNotAvailable` that the resolver turns back into the same grey
placeholder and warning we had before.  Nothing regresses when it is missing.

``pypdfium2`` is also what ``tools/fidelity.py`` already uses to rasterise PowerPoint's
own PDF export, so on a development machine it is present regardless.
"""

from __future__ import annotations

import importlib.util

from .dib import encode_png

#: Pixels per inch to render at, given a target size in EMU.  PowerPoint artwork is
#: usually viewed at well above 1:1 (projected, or on a retina display), and a metafile
#: preview blown up from 96 DPI looks obviously soft next to the vector shapes around it.
#: 192 DPI is the smallest multiple that holds up and keeps the base64 payload sane.
PREVIEW_DPI = 192

#: EMU per inch (914400) -- the OOXML unit.
EMU_PER_INCH = 914400

#: Clamp on the rendered raster.  A slide-sized preview at 192 DPI is ~2600 px wide; the
#: ceiling only exists so a metafile claiming a metre-wide extent cannot make us render
#: a gigapixel page, and the floor so a degenerate zero extent still produces something.
MIN_RENDER_PX = 16
MAX_RENDER_PX = 4096


class PdfRasterizerNotAvailable(RuntimeError):
    """``pypdfium2`` is not installed, so embedded-PDF previews cannot be rendered."""


def pdf_rasterizer_available() -> bool:
    """Whether embedded-PDF previews can be rendered right now.

    Checked with ``find_spec`` rather than a trial import: pdfium loads a sizeable
    shared library, and a mere availability probe should not pay for that.
    """
    return importlib.util.find_spec("pypdfium2") is not None


def rasterise_pdf(data: bytes, *, width_emu: float | None = None) -> bytes | None:
    """Render the first page of ``data`` to PNG bytes.

    ``width_emu`` is the on-slide width the preview will be drawn at, used only to choose
    a resolution; the PDF's own aspect ratio is what decides the pixel height.  Returns
    ``None`` when the PDF cannot be opened or has no pages -- an EMF's embedded preview
    is as untrusted as the EMF around it.

    :raises PdfRasterizerNotAvailable: if ``pypdfium2`` is not installed.
    """
    try:
        import pypdfium2 as pdfium
    except ImportError as error:  # pragma: no cover - exercised only without the extra
        raise PdfRasterizerNotAvailable(
            "rendering an EMF's embedded PDF preview needs pypdfium2: "
            "pip install pptx2svg[metafile]"
        ) from error

    document = None
    try:
        document = pdfium.PdfDocument(data)
        if len(document) == 0:
            return None
        page = document[0]
        page_width, page_height = page.get_size()
        if page_width <= 0 or page_height <= 0:
            return None

        scale = _render_scale(page_width, width_emu)
        bitmap = page.render(scale=scale)
        return _bitmap_to_png(bitmap)
    except Exception:
        # pdfium reports malformed documents as PdfiumError, but the helper layer can
        # also raise ValueError/TypeError on odd input.  A bad preview must degrade to
        # the placeholder, never abort the conversion of the deck.
        return None
    finally:
        if document is not None:
            try:
                document.close()
            except Exception:
                pass


def _render_scale(page_width_pt: float, width_emu: float | None) -> float:
    """Multiplier on the PDF's own 72-DPI point size.

    With a known on-slide extent the preview is rendered at exactly the pixel width it
    will occupy at :data:`PREVIEW_DPI`; without one, fall back to rendering the page at
    that DPI, which is the same thing for a preview authored at its natural size.
    """
    if width_emu and width_emu > 0:
        target_px = width_emu / EMU_PER_INCH * PREVIEW_DPI
    else:
        target_px = page_width_pt / 72.0 * PREVIEW_DPI
    target_px = max(MIN_RENDER_PX, min(MAX_RENDER_PX, target_px))
    return target_px / page_width_pt


def _bitmap_to_png(bitmap) -> bytes | None:
    """Repack a pdfium bitmap as PNG without going through Pillow.

    pdfium hands back a padded buffer in one of four channel orders; the rows are
    stride-padded exactly like a DIB, so the unpacking is the same shape as
    :mod:`.dib`'s.  Doing it here rather than via ``to_pil()`` keeps Pillow out of the
    dependency set -- one optional dependency for this feature, not two.
    """
    width, height, stride, mode = bitmap.width, bitmap.height, bitmap.stride, bitmap.mode
    if width <= 0 or height <= 0:
        return None

    order = {"BGRA": (2, 1, 0, 3), "BGRX": (2, 1, 0, None), "BGR": (2, 1, 0, None),
             "RGBA": (0, 1, 2, 3), "RGBX": (0, 1, 2, None), "RGB": (0, 1, 2, None)}.get(mode)
    if order is None:
        return None
    channels = 4 if len(mode) == 4 else 3
    buffer = bytes(bitmap.buffer)
    if len(buffer) < stride * height:
        return None

    red_i, green_i, blue_i, alpha_i = order
    has_alpha = alpha_i is not None
    rgba = bytearray(width * height * 4)
    for y in range(height):
        row = y * stride
        out = y * width * 4
        for x in range(width):
            source = row + x * channels
            target = out + x * 4
            rgba[target] = buffer[source + red_i]
            rgba[target + 1] = buffer[source + green_i]
            rgba[target + 2] = buffer[source + blue_i]
            rgba[target + 3] = buffer[source + alpha_i] if has_alpha else 255

    return encode_png(width, height, bytes(rgba), has_alpha=has_alpha)
