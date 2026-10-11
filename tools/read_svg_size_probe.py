#!/usr/bin/env python3
"""Read the SVG size probe: where each frame's blue square landed, against the prediction.

``tools/make_svg_size_probe.py`` builds the deck.  For each probe this finds the blue
square's box as fractions of its frame -- in PowerPoint's PDF export, or in a PNG of the
slide (pptx2svg's, at any width) -- and prints it beside what
:func:`ooxml_common.svgimage.natural_size` predicts: the root at that size, its 20x20
``viewBox`` letterboxed into it (``preserveAspectRatio`` ``xMidYMid meet``), stretched to
the frame.  A probe drawn from the green fallback PNG says so.

Usage::

    python3 tools/read_svg_size_probe.py svg-size.pdf            # PowerPoint's export
    python3 tools/read_svg_size_probe.py slide.png [more.png]    # rasterised slides

Needs PyMuPDF for a PDF and Pillow for both (the ``fidelity`` tooling has them).
"""

from __future__ import annotations

import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from make_svg_size_probe import (  # noqa: E402
    CASES, EMU_PER_INCH, attributes, blue_side, frame,
)

from ooxml_common.svgimage import natural_size  # noqa: E402

SLIDE_WIDTH_IN = 10.0
PDF_DPI = 400
TOLERANCE = 0.01  # of the frame: 1.4 x 0.7 inches at 400 dpi is 560 x 280 px


def page_image(path: Path):
    from PIL import Image

    if path.suffix.lower() == ".pdf":
        import pymupdf

        pixmap = pymupdf.open(path)[0].get_pixmap(dpi=PDF_DPI)
        return Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")
    return Image.open(path).convert("RGB")


def predicted(name: str) -> tuple[float, float, float, float]:
    """The blue box as fractions of the frame: left, top, right, bottom."""
    size, has_box = next((size, box) for case, size, box in CASES if case == name)
    width, height = natural_size(attributes(name)) or (0.0, 0.0)
    if not has_box:
        side = blue_side(name)
        return 0.0, 0.0, side / width, side / height
    scale = min(width, height) / 20
    left, top = (width - 20 * scale) / 2, (height - 20 * scale) / 2
    return (left / width, top / height,
            (left + 10 * scale) / width, (top + 10 * scale) / height)


def observed(image, index: int) -> tuple[tuple[float, float, float, float] | None, bool]:
    """The blue box as fractions of probe ``index``'s frame, and whether the frame is the
    green fallback."""
    dpi = image.width / SLIDE_WIDTH_IN
    x, y, cx, cy = (value / EMU_PER_INCH * dpi for value in frame(index))
    left, top, width, height = round(x), round(y), round(cx), round(cy)
    pixels = image.load()

    def blue(px: int, py: int) -> bool:
        r, g, b = pixels[px, py]
        return b > 160 and r < 100 and g < 100

    centre = pixels[left + width // 2, top + height // 2]
    green = centre[1] > 150 and centre[0] < 80 and centre[2] < 80
    columns = [i for i in range(width) if any(blue(left + i, top + j) for j in range(0, height, 2))]
    rows = [j for j in range(height) if any(blue(left + i, top + j) for i in range(0, width, 2))]
    if not columns or not rows:
        return None, green
    return ((columns[0]) / width, rows[0] / height,
            (columns[-1] + 1) / width, (rows[-1] + 1) / height), green


def report(path: Path) -> int:
    image = page_image(path)
    wrong = 0
    print(f"{path.name}: blue box as fractions of the frame (left, top, right, bottom)")
    for index, (name, size, _box) in enumerate(CASES):
        want = predicted(name)
        got, green = observed(image, index)
        if got is None:
            verdict, wrong = ("GREEN FALLBACK" if green else "NO BLUE"), wrong + 1
            shown = "-"
        else:
            error = max(abs(a - b) for a, b in zip(got, want))
            ok = error <= TOLERANCE
            wrong += not ok
            verdict = "ok" if ok else f"MISMATCH ({error:.3f})"
            shown = " ".join(f"{v:.3f}" for v in got)
        print(f"  {name:22} {size or '(none)':40} got {shown:24} "
              f"want {' '.join(f'{v:.3f}' for v in want)}  {verdict}")
    print(f"  {len(CASES) - wrong} of {len(CASES)} as predicted")
    return 1 if wrong else 0


def main(argv: list[str]) -> int:
    if not argv:
        print(__doc__.split("Usage::")[1].strip().splitlines()[0], file=sys.stderr)
        return 2
    return max(report(Path(arg)) for arg in argv)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
