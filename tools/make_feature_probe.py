#!/usr/bin/env python3
"""Build a deck exercising the features no fixture in the corpus covers.

The six real-world fixtures were checked and none of them uses a tab stop, a text
highlight, ``a:bodyPr@rot``, a hidden shape, a non-single underline or multi-column
text -- so none of those could be validated against PowerPoint from the corpus alone.
This writes a small deck that uses each of them, for exporting through
``powerpoint_export_pdf.applescript`` and diffing against our own render.

Text is pinned to Arial deliberately.  The theme font is Aptos, which is not installed
here and has no metric-compatible substitute in ``text/fontmap.py``, so every measured
position -- where a line wraps, where a highlight rectangle starts -- would be off by
the difference between our metrics and whatever face the rasteriser picked.  Pinning a
font we have real metrics for keeps the comparison about the feature under test rather
than about font substitution.

Usage::

    python3 tools/make_feature_probe.py ~/features.pptx
    osascript tools/powerpoint_export_pdf.applescript ~/features.pptx ~/features.pdf
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from derive_table_styles import slide_document, write_deck  # noqa: E402

HERE = Path(__file__).resolve().parent
SOURCE = HERE.parent / "tests/fixtures/authoring-integration.pptx"
BACKDROP = "FFFFFF"


def textbox(shape_id, name, x, y, cx, cy, body_pr, paragraphs, hidden=False):
    hide = ' hidden="1"' if hidden else ""
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"{hide}/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:ln><a:solidFill><a:srgbClr val="C0C0C0"/></a:solidFill></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr {body_pr}/><a:lstStyle/>{paragraphs}</p:txBody></p:sp>'
    )


def run(text, attrs="", body=""):
    return (
        f'<a:r><a:rPr lang="en-US" sz="1400" {attrs}>{body}'
        '<a:latin typeface="Arial"/></a:rPr>'
        f"<a:t>{text}</a:t></a:r>"
    )


def para(runs, properties=""):
    return f"<a:p><a:pPr {properties}><a:buNone/></a:pPr>{runs}</a:p>"


def slides() -> list[str]:
    # 1 -- tab stops: the implicit one-inch grid, then explicit centre and right stops.
    tabs = para(run("a\tb\tc")) + para(
        run("name\tmiddle\tright"),
        'marL="0" indent="0"><a:tabLst>'
        '<a:tab pos="1828800" algn="ctr"/><a:tab pos="3657600" algn="r"/></a:tabLst',
    )

    # 2 -- underline styles and a highlight.
    decoration = (
        para(run("single", 'u="sng"') + run("  ") + run("double", 'u="dbl"'))
        + para(run("dotted", 'u="dotted"') + run("  ") + run("wavy", 'u="wavy"'))
        + para(
            run("plain then ")
            + run("highlighted", "", '<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>')
            + run(" then plain")
        )
    )

    # 3 -- a rotated text body, and a hidden shape that must not draw at all.
    rotation = textbox(
        30, "Rotated", 400000, 300000, 3000000, 900000,
        'rot="1200000" wrap="square"', para(run("rotated 20 degrees")),
    ) + textbox(
        31, "Hidden", 4000000, 300000, 3000000, 900000, 'wrap="square"',
        para(run("THIS MUST NOT APPEAR")), hidden=True,
    )

    # 4 -- two columns, in a box short enough to force the break.
    columns = "".join(
        para(run(f"Paragraph {i} with enough words in it to wrap onto a second line."))
        for i in range(1, 7)
    )

    bodies = [
        textbox(10, "Tab stops", 400000, 300000, 8000000, 1200000, 'wrap="square"', tabs),
        textbox(20, "Decoration", 400000, 300000, 8000000, 1600000, 'wrap="square"', decoration),
        rotation,
        textbox(40, "Columns", 400000, 300000, 8000000, 1600000,
                'numCol="2" wrap="square"', columns),
    ]
    return [slide_document(BACKDROP, body) for body in bodies]


def main() -> int:
    target = Path(
        sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser("~/features.pptx")
    )
    write_deck(SOURCE, target, slides(), BACKDROP)
    print(f"wrote {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
