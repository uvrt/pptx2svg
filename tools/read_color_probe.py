#!/usr/bin/env python3
"""Read a ``make_color_probe.py`` deck's PDF export and say where pptx2svg disagrees.

PowerPoint writes each swatch as a filled path whose colour is three numbers on 0-1,
every one a whole level over 255, and its opacity in the path's graphics state; PyMuPDF's
``get_drawings`` returns both.  A swatch is found by its top left corner on the grid
``make_color_probe.place`` lays out, so the reader needs neither captions nor names.

Each swatch's colour element is read by ``ooxml_common.drawingml.read`` and resolved by
``ooxml_common.drawingml.color`` against the deck's theme colours, under
``ColorRules.POWERPOINT`` (and, for reference, ``ColorRules.WORD``): a row agrees when
every channel is the same level and the opacity is the same to 1/255.

Usage::

    python3 tools/read_color_probe.py deck.pdf [--all]

``--all`` prints every swatch, not only those that disagree.  Needs PyMuPDF.
"""

from __future__ import annotations

import sys
from pathlib import Path
from xml.etree.ElementTree import fromstring

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "src"))
sys.path.insert(0, str(HERE))

import pymupdf  # noqa: E402

from make_color_probe import PROBES, SCHEME, color_xml, key, place  # noqa: E402
from ooxml_common.drawingml import color, read  # noqa: E402
from ooxml_common.drawingml import model as m  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


class _Theme:
    color_scheme = {name: m.SrgbColor(hex_value) for name, hex_value in SCHEME.items()}


def read_export(pdf: Path) -> dict[int, tuple[str, float]]:
    """Swatch index -> (``#rrggbb``, opacity) as PowerPoint drew it."""
    corners = {}
    for index in range(len(PROBES)):
        slide, x, y = place(index)
        corners[(slide, round(x), round(y))] = index
    found: dict[int, tuple[str, float]] = {}
    with pymupdf.open(pdf) as document:
        for number, page in enumerate(document):
            for drawing in page.get_drawings():
                fill = drawing.get("fill")
                rect = drawing["rect"]
                index = corners.get((number, round(rect.x0), round(rect.y0)))
                if fill is None or index is None:
                    continue
                hex_value = "#" + "".join(f"{round(c * 255):02x}" for c in fill[:3])
                found[index] = (hex_value, drawing.get("fill_opacity") or 1.0)
    return found


def resolve(index: int, rules: color.ColorRules) -> tuple[str, float]:
    """What ooxml-common reads and resolves for swatch ``index`` under ``rules``."""
    node = fromstring(f"<a:solidFill xmlns:a='{A}'>{color_xml(*PROBES[index])}</a:solidFill>")
    context = color.ColorContext(_Theme(), color.build_effective_color_map(), rules)
    resolved = color.resolve_color(context, read.parse_color(node))
    return resolved.hex, resolved.alpha


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 1:
        print(__doc__)
        return 2
    show_all = "--all" in sys.argv
    drawn = read_export(Path(args[0]).expanduser())
    agree = {"powerpoint": 0, "word": 0}
    for index in range(len(PROBES)):
        if index not in drawn:
            print(f"{index:4d} {key(*PROBES[index]):45s} not drawn")
            continue
        power = drawn[index]
        ours = resolve(index, color.POWERPOINT)
        word = resolve(index, color.WORD)
        same = ours[0] == power[0] and abs(ours[1] - power[1]) <= 1 / 255
        agree["powerpoint"] += same
        agree["word"] += word[0] == power[0] and abs(word[1] - power[1]) <= 1 / 255
        if show_all or not same:
            print(
                f"{index:4d} {key(*PROBES[index]):45s} PowerPoint {power[0]} {power[1]:.3f}"
                f"  ours {ours[0]} {ours[1]:.3f}  Word's rules {word[0]}"
                + ("" if same else "  DIFFERS")
            )
    print(
        f"{len(drawn)} of {len(PROBES)} swatches read; agree under POWERPOINT "
        f"{agree['powerpoint']}, under WORD {agree['word']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
