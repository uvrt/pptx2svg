#!/usr/bin/env python3
"""Build a deck of colour swatches that measures how PowerPoint composes colour transforms.

Every swatch is a rectangle with no outline, filled with one colour expression: a scheme
or an sRGB colour under one transform or a chain of them -- ``tint``, ``shade``,
``satMod``, ``satOff``, ``lumMod``, ``lumOff``, ``hueMod``, ``hueOff``, ``alpha``,
``alphaMod``, ``alphaOff``, ``comp``, ``inv``, ``gray``, ``gamma`` and ``invGamma`` --
singly at ordinary and extreme values, and chained in both orders where the order could
matter (``tint`` then ``satMod`` and ``satMod`` then ``tint``), including the chains
Office's own themes put on their gradient stops.

**PowerPoint's PDF export writes a solid fill as a vector path with its colour as three
numbers**, so ``tools/read_color_probe.py`` reads the level PowerPoint drew exactly, with
no rasteriser and no antialiasing in between, and the fill's opacity from the path's
graphics state.  Swatches carry no caption: the reader finds each one by its place on the
grid, which ``place`` defines for both tools.

The deck takes the master, layouts and theme of ``real-financial-report.pptx`` (the
2013 Office theme: ``accent1`` is ``4472C4``, ``accent2`` ``ED7D31``) with its master and
layouts emptied.

Usage::

    python3 tools/make_color_probe.py ~/Documents/color-probe.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/Documents/color-probe.pptx ~/Documents/color-probe.pdf
    python3 tools/read_color_probe.py ~/Documents/color-probe.pdf

A probe deck and its export are throwaway and must be deleted again.
"""

from __future__ import annotations

import random
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from make_feature_sweep import NAMESPACES, PT  # noqa: E402
from make_style_probe import write_deck  # noqa: E402

SOURCE = ROOT / "tests/fixtures/real-financial-report.pptx"

#: The scheme colours of ``SOURCE``'s theme that the swatches name.
SCHEME = {"accent1": "4472C4", "accent2": "ED7D31", "accent6": "70AD47", "dk1": "000000",
          "lt1": "FFFFFF"}

#: The slide is 960 x 540 pt.  A swatch is ``SWATCH_W`` x ``SWATCH_H`` on a ``CELL_W`` x
#: ``CELL_H`` grid starting ``MARGIN`` in.
SLIDE_WIDTH, SLIDE_HEIGHT = 960, 540
MARGIN = 20
CELL_W, CELL_H = 46, 32
SWATCH_W, SWATCH_H = 38, 24
COLUMNS = (SLIDE_WIDTH - 2 * MARGIN) // CELL_W
ROWS = (SLIDE_HEIGHT - 2 * MARGIN) // CELL_H

#: Transforms that carry no value.
VALUELESS = {"comp", "inv", "gray", "gamma", "invGamma"}

#: Every swatch: its base (a scheme name or six hex digits) and its transforms in order,
#: each ``(name, value)`` with ``None`` for a transform that takes no value.
PROBES: list[tuple[str, tuple[tuple[str, int | None], ...]]] = []


def probe(base: str, *transforms: tuple[str, int | None]) -> None:
    PROBES.append((base, tuple(transforms)))


def T(name: str, value: int | None = None) -> tuple[str, int | None]:
    return (name, value)


# The untransformed bases: the control that says the export writes the colour it was given.
BASES = ["accent1", "accent2", "accent6", "808080", "C00000", "FFE699", "1F3864", "00B0F0",
         "dk1", "lt1"]
for base in BASES:
    probe(base)

# One transform at a time, across a spread of hues, saturations and lightnesses.
SINGLE_BASES = ["accent1", "accent2", "accent6", "808080", "C00000", "FFE699"]
SINGLES = (
    [T("tint", v) for v in (0, 20000, 50000, 80000, 95000)]
    + [T("shade", v) for v in (0, 20000, 50000, 80000, 95000)]
    + [T("satMod", v) for v in (0, 50000, 150000, 200000, 400000)]
    + [T("satOff", v) for v in (-50000, 25000, 100000)]
    + [T("lumMod", v) for v in (25000, 75000, 150000, 300000)]
    + [T("lumOff", v) for v in (-40000, 20000, 60000)]
    + [T("hueMod", v) for v in (50000, 150000)]
    + [T("hueOff", v) for v in (1800000, 10800000, -5400000)]
    + [T("alpha", v) for v in (50000, 25000)]
    + [T("comp"), T("inv"), T("gray"), T("gamma"), T("invGamma")]
)
for base in SINGLE_BASES:
    for transform in SINGLES:
        probe(base, transform)

# Clamps at the ends of the range.
for base in ("dk1", "lt1", "1F3864"):
    for transform in (T("tint", 50000), T("shade", 50000), T("satMod", 200000),
                      T("lumMod", 150000), T("lumOff", 50000), T("hueOff", 5400000),
                      T("comp"), T("inv"), T("gamma"), T("invGamma")):
        probe(base, transform)

# Chains, in both orders where order could matter, and the Office themes' own stops.
CHAINS = [
    (T("tint", 95000), T("satMod", 170000)),
    (T("satMod", 170000), T("tint", 95000)),
    (T("tint", 50000), T("satMod", 300000)),
    (T("satMod", 300000), T("tint", 50000)),
    (T("shade", 30000), T("satMod", 115000)),
    (T("satMod", 115000), T("shade", 30000)),
    (T("shade", 63000), T("satMod", 120000)),
    (T("tint", 93000), T("satMod", 150000), T("shade", 98000)),
    (T("tint", 98000), T("satMod", 130000), T("shade", 90000)),
    (T("tint", 40000), T("satMod", 350000)),
    (T("tint", 45000), T("satMod", 355000)),
    (T("shade", 51000), T("satMod", 130000)),
    (T("shade", 80000), T("satMod", 130000)),
    (T("shade", 94000), T("satMod", 135000)),
    (T("lumMod", 110000), T("satMod", 105000), T("tint", 67000)),
    (T("lumMod", 105000), T("satMod", 103000), T("tint", 73000)),
    (T("lumMod", 105000), T("satMod", 109000), T("tint", 81000)),
    (T("satMod", 103000), T("lumMod", 102000), T("tint", 94000)),
    (T("satMod", 110000), T("lumMod", 100000), T("shade", 100000)),
    (T("lumMod", 99000), T("satMod", 120000), T("shade", 78000)),
    (T("lumMod", 60000), T("lumOff", 40000)),
    (T("lumOff", 40000), T("lumMod", 60000)),
    (T("lumMod", 75000), T("lumOff", 10000)),
    (T("lumOff", 20000), T("lumMod", 50000)),
    (T("lumMod", 200000), T("lumOff", 50000)),
    (T("lumMod", 50000), T("tint", 50000)),
    (T("tint", 50000), T("lumMod", 50000)),
    (T("tint", 50000), T("shade", 50000)),
    (T("shade", 50000), T("tint", 50000)),
    (T("satMod", 200000), T("satMod", 200000)),
    (T("satMod", 200000), T("satMod", 50000)),
    (T("satMod", 300000), T("lumMod", 50000)),
    (T("lumMod", 50000), T("satMod", 300000)),
    (T("satMod", 300000), T("hueOff", 5400000)),
    (T("satOff", 50000), T("satMod", 50000)),
    (T("hueOff", 5400000), T("satMod", 150000)),
    (T("comp"), T("lumMod", 50000)),
    (T("inv"), T("tint", 50000)),
    (T("gray"), T("tint", 50000)),
    (T("tint", 50000), T("gray")),
    (T("gamma"), T("tint", 50000)),
    (T("tint", 50000), T("gamma")),
    (T("invGamma"), T("shade", 50000)),
    (T("gamma"), T("invGamma")),
    (T("inv"), T("inv")),
    (T("alpha", 50000), T("tint", 50000)),
    (T("alpha", 50000), T("alphaMod", 50000)),
    (T("alpha", 50000), T("alphaOff", 25000)),
    (T("satMod", 1000000),),
    (T("tint", 0), T("satMod", 200000)),
    (T("lumOff", 100000), T("satMod", 200000)),
    (T("lumMod", 0), T("tint", 50000)),
]
CHAIN_BASES = ["accent1", "accent2", "accent6", "808080", "C00000", "FFE699"]
for chain in CHAINS:
    for base in CHAIN_BASES:
        probe(base, *chain)

# Achromatic colours given a saturation, and lumMod past the top of the range then lumOff.
for base in ("404040", "808080", "C0C0C0", "202020", "E0E0E0", "7F7F7F"):
    for chain in ((T("satOff", 10000),), (T("satOff", 50000),), (T("satOff", -25000),),
                  (T("satOff", 50000), T("hueOff", 5400000)),
                  (T("satOff", 50000), T("lumMod", 80000))):
        probe(base, *chain)
for base in ("FFE699", "FFCCCC", "CCFFCC", "FFF2CC", "4472C4", "C0C0C0", "ED7D31", "DEEBF7"):
    for chain in ((T("lumMod", 200000), T("lumOff", 50000)),
                  (T("lumMod", 200000), T("lumOff", 0)),
                  (T("lumMod", 150000), T("lumOff", 10000)),
                  (T("lumMod", 300000), T("lumOff", 50000)),
                  (T("lumMod", 200000), T("lumOff", -50000)),
                  (T("lumMod", 120000), T("lumOff", 20000)),
                  (T("lumMod", 200000), T("tint", 99000), T("lumOff", 50000))):
        probe(base, *chain)

# The other colour elements: ``a:scrgbClr`` (linear light, in percentages), ``a:hslClr``,
# ``a:prstClr`` and ``a:sysClr``, bare and under a transform.
OTHER_BASES = ["scrgb:50000,20000,0", "scrgb:100000,50000,0", "scrgb:21404,21404,21404",
               "scrgb:0,0,100000", "scrgb:5000,40000,90000", "hsl:0,100000,50000",
               "hsl:14400000,50000,25000", "hsl:2700000,80000,60000", "prst:orange",
               "prst:teal", "sys:windowText:000000", "sys:window:FFFFFF"]
for base in OTHER_BASES:
    for chain in ((), (T("lumMod", 50000),), (T("tint", 50000),), (T("satMod", 150000),)):
        probe(base, *chain)

#: Then chains drawn at random -- a seeded generator, so the deck is the same every time --
#: from the values decks and themes use.  One swatch in fifteen or so lands within a
#: hair of a half level, where only the exact arithmetic says which level is drawn: these
#: are what separate one account of PowerPoint's precision from another.
RANDOM_SWATCHES = 2400
_THEME_BASES = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "44546A",
                "E7E6E6", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646",
                "1F497D", "EEECE1"]
_VALUES = {
    "lumMod": [20000, 25000, 40000, 50000, 60000, 75000, 80000, 90000, 95000, 105000,
               110000, 120000, 150000],
    "lumOff": [10000, 15000, 20000, 25000, 40000, 50000, 60000, 80000, -10000, -25000],
    "satMod": [50000, 80000, 103000, 105000, 110000, 115000, 120000, 130000, 150000,
               170000, 200000, 300000, 350000],
    "tint": [20000, 40000, 50000, 60000, 67000, 73000, 80000, 90000, 94000, 95000, 98000],
    "shade": [30000, 50000, 63000, 75000, 78000, 90000, 94000, 98000],
    "hueOff": [600000, 1800000, 5400000, -1800000, 10800000],
    "satOff": [-20000, 10000, 25000, 50000],
    "hueMod": [50000, 90000, 110000, 150000],
}


def _random_swatches() -> None:
    rng = random.Random(5)
    names = sorted(_VALUES)
    for _ in range(RANDOM_SWATCHES):
        if rng.random() < 0.7:
            base = "".join(f"{rng.randrange(256):02X}" for _ in range(3))
        else:
            base = rng.choice(_THEME_BASES)
        if rng.random() < 0.3:
            chain = [T("lumMod", rng.choice(_VALUES["lumMod"])),
                     T("lumOff", rng.choice(_VALUES["lumOff"]))]
            if rng.random() < 0.5:
                chain.append(T(rng.choice(["tint", "shade", "satMod"]), 95000))
        else:
            chain = []
            for _ in range(rng.choice([1, 1, 2, 2, 2, 3])):
                name = rng.choice(names)
                chain.append(T(name, rng.choice(_VALUES[name])))
        probe(base, *chain)


_random_swatches()


def color_xml(base: str, transforms) -> str:
    """The colour element: six hex digits are ``a:srgbClr``, ``scrgb:r,g,b`` ``a:scrgbClr``,
    ``hsl:hue,sat,lum`` ``a:hslClr``, ``prst:name`` ``a:prstClr``, ``sys:name:last``
    ``a:sysClr``, and anything else a scheme colour."""
    children = "".join(
        f"<a:{name}/>" if value is None else f"<a:{name} val='{value}'/>"
        for name, value in transforms
    )
    kind, _, rest = base.partition(":")
    if kind == "scrgb":
        r, g, b = rest.split(",")
        return f"<a:scrgbClr r='{r}' g='{g}' b='{b}'>{children}</a:scrgbClr>"
    if kind == "hsl":
        hue, sat, lum = rest.split(",")
        return f"<a:hslClr hue='{hue}' sat='{sat}' lum='{lum}'>{children}</a:hslClr>"
    if kind == "prst":
        return f"<a:prstClr val='{rest}'>{children}</a:prstClr>"
    if kind == "sys":
        name, last = rest.split(":")
        return f"<a:sysClr val='{name}' lastClr='{last}'>{children}</a:sysClr>"
    tag = "srgbClr" if re.fullmatch(r"[0-9A-F]{6}", base) else "schemeClr"
    return f"<a:{tag} val='{base}'>{children}</a:{tag}>"


def key(base: str, transforms) -> str:
    return " ".join([base] + [n if v is None else f"{n}={v}" for n, v in transforms])


def place(index: int) -> tuple[int, float, float]:
    """Slide, x and y (pt) of swatch ``index``'s top left corner."""
    slide, within = divmod(index, COLUMNS * ROWS)
    row, column = divmod(within, COLUMNS)
    return slide, MARGIN + column * CELL_W, MARGIN + row * CELL_H


def swatch(shape_id: int, base: str, transforms, x: float, y: float) -> str:
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{shape_id}' name='Swatch {shape_id}'/>"
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>"
        f"<a:xfrm><a:off x='{int(x * PT)}' y='{int(y * PT)}'/>"
        f"<a:ext cx='{SWATCH_W * PT}' cy='{SWATCH_H * PT}'/></a:xfrm>"
        "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom>"
        f"<a:solidFill>{color_xml(base, transforms)}</a:solidFill>"
        "<a:ln><a:noFill/></a:ln></p:spPr></p:sp>"
    )


def slides() -> list[str]:
    trees: dict[int, list[str]] = {}
    for index, (base, transforms) in enumerate(PROBES):
        slide, x, y = place(index)
        trees.setdefault(slide, []).append(swatch(100 + index, base, transforms, x, y))
    return [
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + "".join(trees[slide])
        + "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        for slide in sorted(trees)
    ]


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    target = Path(sys.argv[1]).expanduser()
    count = write_deck(target, SOURCE, slides())
    print(f"{len(PROBES)} swatches on {count} slides -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
