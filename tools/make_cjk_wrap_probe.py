#!/usr/bin/env python3
"""Build decks that measure **where PowerPoint breaks a CJK line**.

`sample-cjk.pptx` slide 2 said our break lands one character later than PowerPoint's,
and three explanations fit that observation equally well until something separates them:

* our **glyph widths** are too narrow, so more characters fit;
* our **budget** is too wide -- an inset we do not subtract, or a margin;
* our **comparison at the boundary** admits a character PowerPoint rejects.

Only the third survives these decks, and the reason it is worth a probe rather than a
reading of the fixture is that the fixture confounds all three: its box width, its insets
and its paragraph indent are all inherited through a layout and a master.

Every box here states its own geometry instead.  Insets are zero unless the probe is
about insets, autofit is off, the face is named on the run so neither renderer has to
resolve it, and the text is **one character repeated** -- 東, whose advance is exactly
1 em in Noto Sans JP -- so the line's width in points is the character count times the
font size and nothing else.  A break is then a count, and a count is exact.

    box width = k * size + delta

is the whole experiment.  If the rule is ``n * size <= width`` then ``delta = 0`` fits
``k`` characters and ``delta = -0.25`` fits ``k - 1``; if it is ``<`` then ``delta = 0``
fits ``k - 1`` too.  Whatever tolerance PowerPoint allows shows up as the delta at which
the count changes, in points, at three box widths and four font sizes -- a proportional
slack and a constant one are then different shapes in the table.

| family | holds | varies |
| --- | --- | --- |
| ``w`` | 32 pt text, zero insets | ``k`` and ``delta``: the boundary itself |
| ``z`` | ``delta = 0``, ``k = 10`` | font size 12 through 40 pt -- is slack proportional? |
| ``i`` | 32 pt text, box width | ``lIns``/``rIns``: is the budget the inset box? |
| ``b`` | ``delta = 0``, ``k = 10`` | bold, and bold mixed with upright mid-paragraph |
| ``m`` | ``delta = 0``, ``k = 10`` | a font-size change and a face change mid-paragraph |
| ``p`` | 32 pt text, box width | ``marL``/``indent``: does the indent leave the budget? |
| ``k`` | 32 pt text, ``k = 10`` | 、 and 。 at the break -- kinsoku, or break anywhere? |
| ``q`` | the ``k`` string and box | ``hangingPunct`` and ``eaLnBrk``: which one owns kinsoku |
| ``g`` | one line per paragraph | the two font sizes and ``spcBef``: the vertical advance |

Usage -- the deck's file name picks its probe table::

    python3 tools/make_cjk_wrap_probe.py ~/pptx2svg-oracle/cjk-wrap-box.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/cjk-wrap-box.pptx ~/pptx2svg-oracle/cjk-wrap-box.pdf
    python3 tools/read_cjk_wrap_probe.py ~/pptx2svg-oracle/cjk-wrap-box.pdf --check

The ``g`` family is vertical, so its reader prints baseline advances instead of break
points; it does so on its own, without ``--baselines``.

The decks and their exports are throwaway and are **not** committed; ``~/pptx2svg-oracle``
is the directory PowerPoint is allowed to write to and it is left holding its twenty
corpus files.  See ROADMAP.md section 0.1 before blaming a failed export on the path.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SOURCE = ROOT / "tests/fixtures/authoring-integration.pptx"

NAMESPACES = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
)
SLIDE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"

EMU = 12700
#: Top-left of every probe box, in points.  Away from the edges so nothing is clipped.
BOX_OFF = (36.0, 36.0)
BOX_HEIGHT = 400.0

#: The face both renderers ink, named on the run so no cascade is involved.
FACE = "Noto Sans JP"

#: Advance exactly 1 em in Noto Sans JP, so a line's width is a character count.
FILLER = "東"

EMPTY_TREE = (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>'
)


def emu(points: float) -> int:
    return int(round(points * EMU))


# --------------------------------------------------------------------------------------
# Slide XML
# --------------------------------------------------------------------------------------


def run_xml(text: str, size_pt: float, *, bold: bool = False, face: str = FACE) -> str:
    return (
        f"<a:r><a:rPr lang='ja-JP' sz='{int(round(size_pt * 100))}'"
        f"{' b=\"1\"' if bold else ''} dirty='0'>"
        f"<a:latin typeface='{face}'/><a:ea typeface='{face}'/>"
        f"<a:cs typeface='{face}'/></a:rPr><a:t>{text}</a:t></a:r>"
    )


def slide_xml(index: int, probe: dict) -> str:
    width = probe["width_pt"]
    left_inset = probe.get("l_ins", 0.0)
    right_inset = probe.get("r_ins", 0.0)
    indent = probe.get("indent_pt", 0.0)
    margin_left = probe.get("mar_l_pt", 0.0)
    flags = ""
    if "hanging_punct" in probe:
        flags += f" hangingPunct='{1 if probe['hanging_punct'] else 0}'"
    if "ea_line_break" in probe:
        flags += f" eaLnBrk='{1 if probe['ea_line_break'] else 0}'"
    paragraph_properties = (
        f"<a:pPr marL='{emu(margin_left)}' indent='{emu(indent)}'{flags}>"
        "<a:buNone/></a:pPr>"
    )
    if "paragraphs" in probe:
        # A spacing probe: one paragraph per (size, spcBef %, text), every other
        # property held.  `<a:spcBef>` is stated on each so nothing is inherited.
        body = "".join(
            f"<a:p><a:pPr marL='0' indent='0'{flags}>"
            f"<a:spcBef><a:spcPct val='{int(round(percent * 1000))}'/></a:spcBef>"
            f"<a:buNone/></a:pPr>{run_xml(text, size)}</a:p>"
            for size, percent, text in probe["paragraphs"]
        )
    else:
        runs = "".join(
            run_xml(text, size, bold=bold, face=face)
            for text, size, bold, face in probe["runs"]
        )
        body = f"<a:p>{paragraph_properties}{runs}</a:p>"
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f"<p:sp><p:nvSpPr><p:cNvPr id='{100 + index}' name='{probe['key']}'/>"
        "<p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr>"
        f"<p:spPr><a:xfrm><a:off x='{emu(BOX_OFF[0])}' y='{emu(BOX_OFF[1])}'/>"
        f"<a:ext cx='{emu(width)}' cy='{emu(BOX_HEIGHT)}'/></a:xfrm>"
        "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
        f"<p:txBody><a:bodyPr wrap='square' lIns='{emu(left_inset)}' tIns='0' "
        f"rIns='{emu(right_inset)}' bIns='0' anchor='t'><a:noAutofit/></a:bodyPr>"
        f"<a:lstStyle/>{body}</p:txBody></p:sp>"
        "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )


# --------------------------------------------------------------------------------------
# The probe tables
# --------------------------------------------------------------------------------------


def filler(count: int, size_pt: float = 32.0, *, bold: bool = False, face: str = FACE):
    return (FILLER * count, size_pt, bold, face)


def probe(key: str, width_pt: float, runs: list, **extra) -> dict:
    return {"key": key, "width_pt": width_pt, "runs": runs, **extra}


#: **The boundary.**  Box width ``k * 32 + delta``, and 40 characters to break.  A
#: proportional tolerance shows as a delta threshold that grows with ``k``; a constant
#: one as a threshold that does not.  ``delta = 0`` alone says whether the comparison is
#: ``<=`` or ``<``.
DELTAS = (-1.0, -0.25, 0.0, 0.25, 1.0, 3.0, 6.0, 12.0)

BOX_PROBES: list[dict] = [
    probe(f"w-{k}-{delta:+.2f}", k * 32.0 + delta, [filler(40)])
    for k in (5, 10, 15)
    for delta in DELTAS
]

#: **Font size** at an exact box: ten characters' worth, to the point.  If the answer is
#: ten at every size the rule is scale-free; if a size slips to nine, whatever admits the
#: tenth is a fraction of something.
SIZE_PROBES: list[dict] = [
    probe(f"z-{size:.0f}", 10 * size, [filler(40, size)])
    for size in (12.0, 18.0, 24.0, 32.0, 40.0)
] + [
    # The same sizes one quarter point short, which no rounding rule should let through.
    probe(f"zs-{size:.0f}", 10 * size - 0.25, [filler(40, size)])
    for size in (12.0, 18.0, 24.0, 32.0, 40.0)
]

#: **Insets.**  The box is ten characters wide plus both insets, so a budget that is the
#: inset box fits ten and one that is the frame fits more.
INSET_PROBES: list[dict] = [
    probe("i-0", 320.0, [filler(40)]),
    probe("i-lr7", 320.0 + 14.4, [filler(40)], l_ins=7.2, r_ins=7.2),
    probe("i-l7", 320.0 + 7.2, [filler(40)], l_ins=7.2),
    probe("i-r7", 320.0 + 7.2, [filler(40)], r_ins=7.2),
    probe("i-lr7short", 320.0, [filler(40)], l_ins=7.2, r_ins=7.2),
    probe("i-lr20", 320.0 + 40.0, [filler(40)], l_ins=20.0, r_ins=20.0),
]

#: **Bold**, whose East Asian advance PowerPoint widens by a fixed 0.125 pt when the face
#: has no bold cut.  Noto Sans JP has one, so ten should still fit; a nine says the
#: emboldening applies anyway and enters the wrap.
BOLD_PROBES: list[dict] = [
    probe("b-upright", 320.0, [filler(40)]),
    probe("b-bold", 320.0, [filler(40, bold=True)]),
    probe("b-half", 320.0, [filler(6, bold=True), filler(34)]),
    probe("b-halfafter", 320.0, [filler(6), filler(34, bold=True)]),
    probe("b-bold-1", 320.0 - 0.25, [filler(40, bold=True)]),
]

#: **A format change mid-paragraph**, which is the condition slide 3's overlapping line
#: was about and the condition slide 2's break lands in.  A budget that is a property of
#: the line rather than of the run should be untouched by either.
MIXED_PROBES: list[dict] = [
    probe("m-onesize", 320.0, [filler(40)]),
    # A smaller run first: the break must still be measured in each run's own advances.
    probe("m-small", 320.0, [filler(4, 16.0), filler(36)]),
    probe("m-large", 320.0, [filler(4, 40.0), filler(36)]),
    # A face change with the same advance, so only the run boundary differs.
    probe("m-face", 320.0, [filler(6, face="Meiryo"), filler(34)]),
]

#: **The paragraph indent.**  ``sample-cjk``'s bullets carry ``marL``, and whether it
#: leaves the budget for every line or only the first is what our
#: ``effective_text_width`` assumes.
INDENT_PROBES: list[dict] = [
    probe("p-0", 320.0, [filler(40)]),
    probe("p-marl32", 320.0 + 32.0, [filler(40)], mar_l_pt=32.0),
    probe("p-marl32short", 320.0, [filler(40)], mar_l_pt=32.0),
    probe("p-hang", 320.0 + 32.0, [filler(40)], mar_l_pt=32.0, indent_pt=-32.0),
]

#: **Kinsoku.**  Japanese forbids a line *beginning* with 、 。 ) ｝ ー and friends, and
#: forbids one *ending* with an opening bracket.  Our wrapper breaks between any two CJK
#: characters, so a 、 pushed to the head of a line is a defect we would not otherwise
#: see.  Each string here puts the punctuation at the character the box's width makes the
#: first of line two.
KINSOKU_PROBES: list[dict] = [
    probe("k-plain", 320.0, [filler(40)]),
    # 、 as character 11: the head of line two if the break is blind.
    probe("k-comma11", 320.0, [(FILLER * 10 + "、" + FILLER * 29, 32.0, False, FACE)]),
    probe("k-stop11", 320.0, [(FILLER * 10 + "。" + FILLER * 29, 32.0, False, FACE)]),
    probe("k-close11", 320.0, [(FILLER * 10 + "」" + FILLER * 29, 32.0, False, FACE)]),
    probe("k-dash11", 320.0, [(FILLER * 10 + "ー" + FILLER * 29, 32.0, False, FACE)]),
    probe("k-small11", 320.0, [(FILLER * 10 + "っ" + FILLER * 29, 32.0, False, FACE)]),
    # An opening bracket as character 10, which may not end a line.
    probe("k-open10", 320.0, [(FILLER * 9 + "「" + FILLER * 30, 32.0, False, FACE)]),
    # Two in a row, which is where a naive "push one down" rule fails.
    probe("k-two11", 320.0, [(FILLER * 10 + "、。" + FILLER * 28, 32.0, False, FACE)]),
]

#: **The two switches that own the kinsoku rule.**  ``hangingPunct`` lets a forbidden
#: character hang past the right edge instead of pushing its neighbour down, and
#: ``eaLnBrk`` turns East Asian line breaking off altogether.  ``sample-cjk``'s master
#: writes ``eaLnBrk="1" hangingPunct="1"`` and the ``k`` family above states neither, so
#: without this family we would be reading one setting and acting on the other.
PUNCT_PROBES: list[dict] = [
    probe("q-none", 320.0, [(FILLER * 10 + "、" + FILLER * 29, 32.0, False, FACE)]),
    probe("q-hang1", 320.0, [(FILLER * 10 + "、" + FILLER * 29, 32.0, False, FACE)],
          hanging_punct=True),
    probe("q-hang0", 320.0, [(FILLER * 10 + "、" + FILLER * 29, 32.0, False, FACE)],
          hanging_punct=False),
    probe("q-ea0", 320.0, [(FILLER * 10 + "、" + FILLER * 29, 32.0, False, FACE)],
          ea_line_break=False),
    probe("q-ea1", 320.0, [(FILLER * 10 + "、" + FILLER * 29, 32.0, False, FACE)],
          ea_line_break=True),
    probe("q-hang1ea1", 320.0, [(FILLER * 10 + "、" + FILLER * 29, 32.0, False, FACE)],
          hanging_punct=True, ea_line_break=True),
    # The opening bracket, which is forbidden at a line *end* rather than at its head and
    # which hanging punctuation cannot rescue.
    probe("q-open-hang1", 320.0, [(FILLER * 9 + "「" + FILLER * 30, 32.0, False, FACE)],
          hanging_punct=True),
    # And the plain control, so a deck that reproduces the `k` family is visible as such.
    probe("q-plain-hang1", 320.0, [filler(40)], hanging_punct=True),
]

#: **The paragraph gap when the font size changes**, which is what ``sample-cjk`` slide 4
#: is made of.  Its five paragraphs are 32, 32, 28, 24 and 32 pt with ``spcBef`` 20% on
#: every one, and PowerPoint's baseline-to-baseline advances are 46.08, 42.00, 34.80 and
#: 44.16 pt where ours are 46.08, 40.32, 34.56 and 46.08.  Three readings survive that:
#: the percentage is a share of the *previous* paragraph's line box rather than the
#: following one's; the advance carries the difference of the two faces' descents; or the
#: gap is not a function of the two sizes at all.  **No linear function of the two sizes
#: fits all four**, which is why this family holds one size and sweeps the other rather
#: than sampling pairs out of a real deck.
#:
#: Each slide is two paragraphs of one line each, ``buNone`` and no autofit, so the only
#: number to read is the distance between two baselines.
SPACING_PROBES: list[dict] = [
    # First size held at 32, second swept: if the percentage is a share of the *first*
    # paragraph's box then the gap is constant down this column.
    *[
        probe(f"g-32-{second:.0f}", 400.0, [],
              paragraphs=[(32.0, 20.0, FILLER * 3), (second, 20.0, FILLER * 3)])
        for second in (12.0, 18.0, 24.0, 28.0, 32.0, 40.0)
    ],
    # Second size held at 32, first swept: the mirror image.
    *[
        probe(f"g-{first:.0f}-32", 400.0, [],
              paragraphs=[(first, 20.0, FILLER * 3), (32.0, 20.0, FILLER * 3)])
        for first in (12.0, 18.0, 24.0, 28.0, 40.0)
    ],
    # The percentage itself at one size pair, which says whether the share is linear.
    *[
        probe(f"g-pct{percent:.0f}", 400.0, [],
              paragraphs=[(32.0, percent, FILLER * 3), (24.0, percent, FILLER * 3)])
        for percent in (0.0, 10.0, 20.0, 50.0, 100.0)
    ],
    # And zero spacing at four size pairs, which isolates the line box from the gap.
    *[
        probe(f"g-zero-{first:.0f}-{second:.0f}", 400.0, [],
              paragraphs=[(first, 0.0, FILLER * 3), (second, 0.0, FILLER * 3)])
        for first, second in ((32.0, 32.0), (32.0, 24.0), (24.0, 32.0), (32.0, 12.0))
    ],
]

DECKS = {
    "cjk-wrap-spacing": SPACING_PROBES,
    "cjk-wrap-punct": PUNCT_PROBES,
    "cjk-wrap-box": BOX_PROBES,
    "cjk-wrap-size": SIZE_PROBES,
    "cjk-wrap-inset": INSET_PROBES + BOLD_PROBES + MIXED_PROBES + INDENT_PROBES,
    "cjk-wrap-kinsoku": KINSOKU_PROBES,
}


# --------------------------------------------------------------------------------------
# Deck assembly -- the same shape as tools/make_legend_probe.py
# --------------------------------------------------------------------------------------


def slide_rels() -> str:
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        f"<Relationship Id='rId1' Type='{LAYOUT_REL}' Target='../slideLayouts/slideLayout1.xml'/>"
        "</Relationships>"
    )


def blank_template(xml: str) -> str:
    xml = re.sub(r"<p:spTree>.*?</p:spTree>", EMPTY_TREE, xml, flags=re.S)
    return re.sub(r"<p:bg>.*?</p:bg>", "", xml, flags=re.S)


def probes_for(target: Path) -> list[dict]:
    for name, table in DECKS.items():
        if name in target.stem:
            return table
    raise SystemExit(f"name the deck after one of {sorted(DECKS)}, not {target.stem!r}")


def write_deck(target: Path, probes: list[dict]) -> None:
    source = zipfile.ZipFile(SOURCE)
    presentation = source.read("ppt/presentation.xml").decode()
    pres_rels = source.read("ppt/_rels/presentation.xml.rels").decode()
    content_types = source.read("[Content_Types].xml").decode()
    numbers = range(1, len(probes) + 1)

    presentation = re.sub(
        r"<p:sldIdLst>.*?</p:sldIdLst>",
        "<p:sldIdLst>"
        + "".join(f'<p:sldId id="{255 + n}" r:id="rIdS{n}"/>' for n in numbers)
        + "</p:sldIdLst>",
        presentation,
        flags=re.S,
    )
    pres_rels = re.sub(r'<Relationship[^>]*Type="[^"]*/slide"[^>]*/>', "", pres_rels)
    pres_rels = pres_rels.replace(
        "</Relationships>",
        "".join(
            f'<Relationship Id="rIdS{n}" Type="{SLIDE_REL}" Target="slides/slide{n}.xml"/>'
            for n in numbers
        )
        + "</Relationships>",
    )
    content_types = re.sub(
        r'<Override PartName="/ppt/(slides/slide|charts/chart)\d+\.xml"[^>]*/>', "", content_types
    )
    content_types = content_types.replace(
        "</Types>",
        "".join(
            f'<Override PartName="/ppt/slides/slide{n}.xml" ContentType="{SLIDE_TYPE}"/>'
            for n in numbers
        )
        + "</Types>",
    )

    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.infolist():
            name = item.filename
            if name.startswith("ppt/slides/") or name.startswith("ppt/charts/"):
                continue
            data = source.read(name)
            if name == "ppt/presentation.xml":
                data = presentation.encode()
            elif name == "ppt/_rels/presentation.xml.rels":
                data = pres_rels.encode()
            elif name == "[Content_Types].xml":
                data = content_types.encode()
            elif re.search(r"slide(Layouts|Masters)/slide\w+\d+\.xml$", name):
                data = blank_template(data.decode()).encode()
            out.writestr(name, data)
        for index, item in enumerate(probes):
            out.writestr(f"ppt/slides/slide{index + 1}.xml", slide_xml(index, item))
            out.writestr(f"ppt/slides/_rels/slide{index + 1}.xml.rels", slide_rels())


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    target = Path(sys.argv[1]).expanduser()
    probes = probes_for(target)
    write_deck(target, probes)
    print(f"wrote {target} -- {len(probes)} slides")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
