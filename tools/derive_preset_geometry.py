#!/usr/bin/env python3
"""Compile preset shape geometry out of ECMA-376's own ``presetShapeDefinitions.xml``.

DrawingML's 186 preset shapes are not described in prose; the standard ships their
geometry as a data file, in an electronic addendum to Part 1.  Each entry is an
``avLst`` of adjustment defaults, a ``gdLst`` of guide formulas and a ``pathLst`` of
drawing commands written in terms of those guides -- exactly the vocabulary
:mod:`pptx2svg.guides` already evaluates for ``a:custGeom``.

So the geometry does not have to be re-derived by hand.  This script transcribes the
definitions we use into :mod:`pptx2svg.render.preset_specs`, a pure-data module the
renderer evaluates at draw time.  Hand-written generators still exist for the shapes
where an approximation is the better trade -- a ``rect`` should emit ``<rect>``, not a
four-command path -- and :data:`SPEC_DRIVEN` below is the list of the ones that should
come from the specification instead.

    python3 tools/derive_preset_geometry.py --source presetShapeDefinitions.xml
    python3 tools/derive_preset_geometry.py --check --source ...   # exit 1 if stale

**The source is not redistributed here.**  ECMA's text copyright policy governs the
standard's own files, and vendoring a copy is a decision for the project rather than for
this script, so ``--source`` must point at a copy you obtained yourself and its SHA-256
is verified against :data:`SOURCE_SHA256` before anything is read.  A mismatch is fatal:
silently compiling a different edition would change shapes with no diff to show for it.

Where to get it:

* Landing page: https://ecma-international.org/publications-and-standards/standards/ecma-376/
* Package: ``ECMA-376-1_5th_edition_december_2016.zip`` (5th edition, December 2016)
* Inside it: ``OfficeOpenXML-DrawingMLGeometries.zip``
* Inside that: ``presetShapeDefinitions.xml``

``--source`` accepts either the inner ``.zip`` or the extracted ``.xml``.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
TARGET = HERE.parent / "src/pptx2svg/render/preset_specs.py"

DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

#: SHA-256 of the two files we accept, so a wrong edition fails loudly rather than
#: quietly redrawing every shape.  Both are from ECMA-376 Part 1, 5th edition (2016).
SOURCE_SHA256 = {
    "presetShapeDefinitions.xml": (
        "2f7c868d857c1e3c4b5a6068759fe0e07d77ad58377a6618d1b02ba3507b6939"
    ),
    "OfficeOpenXML-DrawingMLGeometries.zip": (
        "a8517838a3a60492d267d1ff92e1ebd831adc04584de79a38a5532a68902abf6"
    ),
}

#: Presets the renderer takes from the specification rather than from a hand-written
#: generator.  The rule of thumb: a shape whose outline a person can write down exactly
#: in three lines of Python stays hand-written, because a native ``<rect>`` or
#: ``<ellipse>`` is smaller and scales its stroke properly; a shape with a dozen guides
#: and a trigonometric arrowhead comes from here, because approximating it by eye
#: produces something recognisably not what PowerPoint draws.
#:
#: Adding a name to this list and re-running is the whole process for a new preset.
SPEC_DRIVEN = [
    # Callouts.
    "callout1", "callout2", "callout3",
    "accentCallout1", "accentCallout2", "accentCallout3",
    "accentBorderCallout1", "accentBorderCallout2", "accentBorderCallout3",
    "leftArrowCallout", "rightArrowCallout", "upArrowCallout", "downArrowCallout",
    "leftRightArrowCallout", "upDownArrowCallout", "quadArrowCallout",
    # Action buttons.
    "actionButtonBackPrevious", "actionButtonBeginning", "actionButtonBlank",
    "actionButtonDocument", "actionButtonEnd", "actionButtonForwardNext",
    "actionButtonHelp", "actionButtonHome", "actionButtonInformation",
    "actionButtonMovie", "actionButtonReturn", "actionButtonSound",
    # Curved and circular arrows.
    "curvedUpArrow", "curvedDownArrow", "curvedLeftArrow", "curvedRightArrow",
    "circularArrow", "leftCircularArrow", "leftRightCircularArrow", "swooshArrow",
    # Banners and scrolls.
    "verticalScroll", "horizontalScroll",
    "ellipseRibbon", "ellipseRibbon2", "leftRightRibbon",
    # Gears, tabs and the remaining tail.
    "gear6", "gear9", "funnel", "pieWedge",
    "cornerTabs", "squareTabs", "plaqueTabs", "nonIsoscelesTrapezoid",
    "chartPlus", "chartStar", "chartX", "flowChartOfflineStorage",
    "lineInv",
    # Promoted from hand-written approximations after measuring how far they had
    # drifted from the specification; see tests/test_geometry_presets.py.
    "arc",
    "chord",
    "pie",
    "blockArc",
    "donut",
    "noSmoking",
    "moon",
    "sun",
    "smileyFace",
    "lightningBolt",
    "flowChartMagneticTape",
    "bentConnector2",
    "bentConnector3",
    "bentConnector4",
    "bentConnector5",
    "curvedConnector2",
    "curvedConnector3",
    "curvedConnector4",
    "curvedConnector5",
    "bracePair",
    "bracketPair",
    "mathPlus",
    "mathMinus",
    "mathMultiply",
    "mathDivide",
    "mathEqual",
    "mathNotEqual",
    # Stars.  Held back until PowerPoint could be asked directly, because the spec's
    # defaults looked too shallow to be right.  Measured against its own PDF export they
    # match at 0.975-0.999 silhouette overlap; the hand-written ones managed 0.42-0.75.
    "star4",
    "star5",
    "star6",
    "star7",
    "star8",
    "star10",
    "star12",
    "star16",
    "star24",
    "star32",
    # Arrows and ribbons: measured against PowerPoint at two aspect ratios, ours 0.32-0.78, the specification 0.99-1.00.
    "bentArrow",
    "bentUpArrow",
    "chevron",
    "leftArrow",
    "leftRightArrow",
    "leftRightUpArrow",
    "leftUpArrow",
    "notchedRightArrow",
    "quadArrow",
    "ribbon",
    "ribbon2",
    "rightArrow",
    "stripedRightArrow",
    "uturnArrow",
    # Flowchart shapes: ours 0.54-0.90 against PowerPoint, the specification 0.998-1.000.
    "flowChartDocument",
    "flowChartMagneticDisk",
    "flowChartMagneticDrum",
    "flowChartMultidocument",
    "flowChartOnlineStorage",
    "flowChartOr",
    "flowChartPunchedTape",
    "flowChartSummingJunction",
    "flowChartTerminator",
]

HEADER = '''"""Preset shape geometry, compiled from ECMA-376's ``presetShapeDefinitions.xml``.

**Generated file -- do not edit.**  Regenerate with::

    python3 tools/derive_preset_geometry.py --source presetShapeDefinitions.xml

and see that script for where to obtain the source and why it is not vendored here.
``--check`` fails when this file has drifted from what the source compiles to.

Source: ECMA-376 Part 1, 5th edition (December 2016), electronic addendum
``OfficeOpenXML-DrawingMLGeometries.zip`` -> ``presetShapeDefinitions.xml``
SHA-256 ``{sha}``.

Each entry is ``name -> (adjustments, guides, paths)``, holding the shape's ``a:avLst``,
``a:gdLst`` and ``a:pathLst`` verbatim:

* **adjustments** -- ``(name, default)`` pairs.  Defaults are the raw OOXML integers, in
  the 1/1000-percent or 1/60000-degree units the guides expect; a shape's own ``a:avLst``
  overrides them by name.
* **guides** -- ``(name, formula)`` pairs in evaluation order, which is *not* alphabetical:
  a guide may refer to any guide declared before it.
* **paths** -- ``(fill, stroke, space, commands)`` per ``a:path``, where ``fill`` is the
  path's ``@fill`` mode (``norm``, ``none``, ``darken``, ``darkenLess``, ``lighten``,
  ``lightenLess``), ``stroke`` is its ``@stroke``, ``space`` is ``(@w, @h)`` when the path
  is authored in its own coordinate system and ``None`` otherwise, and each command is a
  tuple of an SVG-style letter and its operands as guide-name or literal strings:
  ``("M", x, y)``, ``("L", x, y)``, ``("A", wR, hR, stAng, swAng)``, ``("Q", ...)``,
  ``("C", ...)``, ``("Z",)``.

Nothing here is evaluated until :mod:`pptx2svg.render.geometry` draws it.
"""

# fmt: off
PRESET_SPECS: dict = {{
'''

FOOTER = "}\n"

LINE_LIMIT = 96


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def load_source(path: Path) -> tuple[bytes, str]:
    """Read ``presetShapeDefinitions.xml`` out of ``path``, verifying its hash."""
    raw = path.read_bytes()
    digest = sha256(raw)

    if path.suffix.lower() == ".zip":
        expected = SOURCE_SHA256["OfficeOpenXML-DrawingMLGeometries.zip"]
        if digest != expected:
            raise SystemExit(
                f"{path} has SHA-256 {digest}, expected {expected}.\n"
                "That is not the 5th-edition DrawingML geometry archive; see the module "
                "docstring for where to obtain it."
            )
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            xml = archive.read("presetShapeDefinitions.xml")
    else:
        xml = raw

    digest = sha256(xml)
    expected = SOURCE_SHA256["presetShapeDefinitions.xml"]
    if digest != expected:
        raise SystemExit(
            f"presetShapeDefinitions.xml has SHA-256 {digest}, expected {expected}.\n"
            "Compiling a different edition would change shapes with no diff to show for "
            "it, so this is fatal rather than a warning."
        )
    return xml, digest


def parse(xml: bytes) -> dict[str, ET.Element]:
    root = ET.fromstring(xml)
    return {child.tag.rsplit("}", 1)[-1]: child for child in root}


def tag(element: ET.Element) -> str:
    return element.tag.rsplit("}", 1)[-1]


def find(parent: ET.Element, name: str) -> ET.Element | None:
    for child in parent:
        if tag(child) == name:
            return child
    return None


def guides_of(shape: ET.Element, list_name: str) -> list[tuple[str, str]]:
    node = find(shape, list_name)
    if node is None:
        return []
    return [(gd.get("name"), gd.get("fmla")) for gd in node if tag(gd) == "gd"]


def commands_of(path: ET.Element) -> list[tuple]:
    """``a:path`` children -> SVG-style command tuples of unevaluated operands."""
    out: list[tuple] = []
    for node in path:
        kind = tag(node)
        points = [p for p in node if tag(p) == "pt"]
        if kind == "moveTo":
            out.append(("M", points[0].get("x"), points[0].get("y")))
        elif kind == "lnTo":
            out.append(("L", points[0].get("x"), points[0].get("y")))
        elif kind == "close":
            out.append(("Z",))
        elif kind == "arcTo":
            out.append(
                ("A", node.get("wR"), node.get("hR"), node.get("stAng"), node.get("swAng"))
            )
        elif kind in ("quadBezTo", "cubicBezTo"):
            letter = "Q" if kind == "quadBezTo" else "C"
            operands: list[str] = []
            for point in points:
                operands.extend((point.get("x"), point.get("y")))
            out.append((letter, *operands))
    return out


def paths_of(shape: ET.Element) -> list[tuple]:
    path_list = find(shape, "pathLst")
    if path_list is None:
        return []
    out = []
    for path in path_list:
        if tag(path) != "path":
            continue
        width, height = path.get("w"), path.get("h")
        space = (int(width), int(height)) if width and height else None
        out.append(
            (
                path.get("fill", "norm"),
                path.get("stroke", "true") != "false",
                space,
                commands_of(path),
            )
        )
    return out


def pack(items: list[str], indent: str) -> list[str]:
    """Pack ``items`` onto as few lines as fit, so the output stays reviewable.

    Generated data is still read by people -- when a shape looks wrong the first thing
    anyone does is diff this file against the specification -- so it wraps to the same
    width as the hand-written source rather than emitting one enormous line.
    """
    lines: list[str] = []
    current = indent
    for item in items:
        piece = item + ", "
        if len(current) + len(piece) > LINE_LIMIT and current != indent:
            lines.append(current.rstrip())
            current = indent
        current += piece
    if current != indent:
        lines.append(current.rstrip())
    return lines


def tuple_lines(items: list[str], open_indent: str, indent: str, suffix: str) -> list[str]:
    """``items`` as a parenthesised tuple, kept on one line when it fits."""
    if not items:
        return [f"{open_indent}(){suffix}"]
    packed = pack(items, indent)
    if len(packed) == 1 and len(open_indent) + len(packed[0].strip()) + 2 <= LINE_LIMIT:
        return [f"{open_indent}({packed[0].strip()}){suffix}"]
    return [f"{open_indent}("] + packed + [f"{indent[:-4]}){suffix}"]


def render_shape(name: str, shape: ET.Element) -> str:
    adjustments = [f'("{n}", {f.split()[1]})' for n, f in guides_of(shape, "avLst")]
    guides = [f'("{n}", "{f}")' for n, f in guides_of(shape, "gdLst")]

    lines = [f'    "{name}": (']
    lines += tuple_lines(adjustments, "        ", "            ", ",")
    lines += tuple_lines(guides, "        ", "            ", ",")

    lines.append("        (")
    for fill, stroke, space, commands in paths_of(shape):
        rendered = [repr(tuple(c)).replace("'", '"') for c in commands]
        lines.append(f'            ("{fill}", {stroke}, {space}, (')
        lines += pack(rendered, "                ")
        lines.append("            )),")
    lines.append("        ),")
    lines.append("    ),")
    return "\n".join(lines)


def build(xml: bytes, digest: str) -> str:
    shapes = parse(xml)
    missing = [name for name in SPEC_DRIVEN if name not in shapes]
    if missing:
        raise SystemExit(f"not in the specification: {', '.join(missing)}")

    body = "\n".join(render_shape(name, shapes[name]) for name in SPEC_DRIVEN)
    return HEADER.format(sha=digest) + body + "\n" + FOOTER


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--source",
        type=Path,
        required=True,
        help="presetShapeDefinitions.xml, or the OfficeOpenXML-DrawingMLGeometries.zip "
        "containing it",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help=f"exit 1 if {TARGET.name} differs from what the source compiles to",
    )
    args = parser.parse_args()

    xml, digest = load_source(args.source)
    generated = build(xml, digest)

    if args.check:
        if not TARGET.exists():
            print(f"{TARGET} does not exist", file=sys.stderr)
            return 1
        if TARGET.read_text() != generated:
            print(f"{TARGET} is stale; re-run without --check", file=sys.stderr)
            return 1
        print(f"{TARGET.name} is up to date ({len(SPEC_DRIVEN)} presets)")
        return 0

    TARGET.write_text(generated)
    print(f"wrote {TARGET} ({len(SPEC_DRIVEN)} presets)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
