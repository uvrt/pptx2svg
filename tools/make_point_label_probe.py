#!/usr/bin/env python3
"""Build a deck that shows what PowerPoint draws for a data label with text of its own.

A ``c:dLbl`` may carry ``c:tx``: rich DrawingML where a label was typed over, or Office's
"Value From Cells" -- a ``CELLRANGE`` field over the series' ``c15:datalabelsRange`` -- and
the chart then prints that text instead of composing one from ``c:showVal`` and its
siblings.  One chart per slide:

``rich``
    a scatter whose four points carry rich labels: plain, two runs (bold red, plain),
    18 pt, and one whose ``c:show*`` flags are all off;
``bubble``
    the same labels on a bubble chart;
``range``
    a column chart with a ``c15:datalabelsRange`` cache and ``c15:showDataLabelsRange``
    beside ``c:showCatName`` and ``c:showVal``, no ``c:tx`` -- which part comes first;
``fields``
    per-point rich text made of ``CELLRANGE``, ``CATEGORYNAME``, ``SERIESNAME`` and
    ``VALUE`` fields with stale cached text;
``strref``
    a label whose ``c:tx`` is a ``c:strRef`` with a cached string.

    python3 tools/make_point_label_probe.py ~/point-labels.pptx
    osascript tools/powerpoint_export_pdf.applescript ~/point-labels.pptx ~/point-labels.pdf

The deck and its export are throwaway and are not committed.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import make_bubble_probe as bubble  # noqa: E402

C, A, R = bubble.C, bubble.A, bubble.R
C15 = 'xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"'
FLAGS = ("LegendKey", "Val", "CatName", "SerName", "Percent", "BubbleSize")


def flags(**on: bool) -> str:
    return "".join(f"<c:show{name} val='{int(on.get(name, False))}'/>" for name in FLAGS)


def rich(*runs: str) -> str:
    return f"<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>{''.join(runs)}</a:p></c:rich></c:tx>"


def run(text: str, props: str = "", inner: str = "") -> str:
    return f"<a:r><a:rPr lang='en-US' {props}>{inner}</a:rPr><a:t>{text}</a:t></a:r>"


def fld(kind: str, cached: str) -> str:
    return (f"<a:fld id='{{0F0E0D0C-0000-4000-8000-0000000000{len(kind):02d}}}' type='{kind}'>"
            f"<a:rPr lang='en-US'/><a:t>{cached}</a:t></a:fld>")


def d_lbl(index: int, tx: str, shown: dict | None = None, extra: str = "") -> str:
    shown = {"Val": True} if shown is None else shown
    return (f"<c:dLbl><c:idx val='{index}'/>{tx}<c:dLblPos val='r'/>{flags(**shown)}"
            f"{extra}</c:dLbl>")


RICH_LABELS = (
    d_lbl(0, rich(run("R1")))
    + d_lbl(1, rich(run("Bold", "b='1'", "<a:solidFill><a:srgbClr val='C00000'/></a:solidFill>"),
                    run(" plain")))
    + d_lbl(2, rich(run("Big", "sz='1800'")))
    + d_lbl(3, rich(run("Hidden")), shown={})
)

RANGE_EXT = (
    f"<c:extLst><c:ext uri='{{02D57815-91ED-43cb-92C2-25804820EDAC}}' {C15}>"
    "<c15:datalabelsRange><c15:f>Sheet1!$D$2:$D$4</c15:f><c15:dlblRangeCache>"
    "<c:ptCount val='3'/><c:pt idx='0'><c:v>alpha</c:v></c:pt><c:pt idx='1'><c:v>beta</c:v></c:pt>"
    "<c:pt idx='2'><c:v>gamma</c:v></c:pt></c15:dlblRangeCache></c15:datalabelsRange>"
    "</c:ext></c:extLst>"
)
SHOW_RANGE = (
    f"<c:extLst><c:ext uri='{{CE6537A1-D6FC-4f65-9D91-7224C49458BB}}' {C15}>"
    "<c15:showDataLabelsRange val='1'/></c:ext></c:extLst>"
)


def _scatter(labels: str, bubble_chart: bool) -> str:
    x = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate((1, 2, 3, 4)))
    y = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate((2, 4, 3, 1)))
    size = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate((3, 2, 3, 1)))
    num = "<c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val='4'/>{}</c:numCache></c:numRef>"
    series = (
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/><c:pt idx='0'><c:v>Ideas</c:v></c:pt>"
        "</c:strCache></c:strRef></c:tx>"
        + ("<c:spPr><a:noFill/><a:ln><a:solidFill><a:srgbClr val='4472C4'/></a:solidFill></a:ln></c:spPr>"
           "<c:invertIfNegative val='0'/>" if bubble_chart else
           "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val='circle'/><c:size val='7'/></c:marker>")
        + f"<c:dLbls>{labels}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>"
        f"<c:dLblPos val='r'/>{flags(Val=True)}</c:dLbls>"
        + f"<c:xVal>{num.format(x)}</c:xVal><c:yVal>{num.format(y)}</c:yVal>"
        + (f"<c:bubbleSize>{num.format(size)}</c:bubbleSize><c:bubble3D val='0'/>" if bubble_chart
           else "<c:smooth val='0'/>")
        + "</c:ser>"
    )
    if bubble_chart:
        group = (f"<c:bubbleChart><c:varyColors val='0'/>{series}<c:bubbleScale val='50'/>"
                 "<c:showNegBubbles val='0'/>")
        end = "</c:bubbleChart>"
    else:
        group = f"<c:scatterChart><c:scatterStyle val='lineMarker'/><c:varyColors val='0'/>{series}"
        end = "</c:scatterChart>"
    return (group + f"<c:axId val='{bubble.X_AXIS}'/><c:axId val='{bubble.Y_AXIS}'/>" + end
            + bubble._value_axis(bubble.X_AXIS, bubble.Y_AXIS, "b", False)
            + bubble._value_axis(bubble.Y_AXIS, bubble.X_AXIS, "l", True))


def _column(labels: str, series_ext: str) -> str:
    cats = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(("North", "South", "East")))
    vals = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate((3, 5, 4)))
    return (
        "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/><c:varyColors val='0'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/><c:pt idx='0'><c:v>Sales</c:v></c:pt>"
        "</c:strCache></c:strRef></c:tx>"
        "<c:spPr><a:solidFill><a:srgbClr val='A5A5A5'/></a:solidFill></c:spPr>"
        "<c:invertIfNegative val='0'/>"
        f"<c:dLbls>{labels}</c:dLbls>"
        f"<c:cat><c:strRef><c:strCache><c:ptCount val='3'/>{cats}</c:strCache></c:strRef></c:cat>"
        f"<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val='3'/>"
        f"{vals}</c:numCache></c:numRef></c:val>{series_ext}</c:ser>"
        "<c:gapWidth val='80'/>"
        f"<c:axId val='{bubble.X_AXIS}'/><c:axId val='{bubble.Y_AXIS}'/></c:barChart>"
        f"<c:catAx><c:axId val='{bubble.X_AXIS}'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='b'/><c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='none'/><c:minorTickMark val='none'/><c:tickLblPos val='nextTo'/>"
        f"<c:crossAx val='{bubble.Y_AXIS}'/><c:crosses val='autoZero'/><c:auto val='1'/>"
        "<c:lblAlgn val='ctr'/><c:lblOffset val='100'/><c:noMultiLvlLbl val='0'/></c:catAx>"
        + bubble._value_axis(bubble.Y_AXIS, bubble.X_AXIS, "l", True).replace(
            "<c:crossBetween val='midCat'/>", "<c:crossBetween val='between'/>")
    )


PROBES = {
    "rich": lambda: _scatter(RICH_LABELS, False),
    "bubble": lambda: _scatter(RICH_LABELS, True),
    "range": lambda: _column(
        f"<c:dLblPos val='outEnd'/>{flags(Val=True, CatName=True)}{SHOW_RANGE}", RANGE_EXT),
    "fields": lambda: _column(
        "".join(
            d_lbl(i, rich(fld("CELLRANGE", "[CELLRANGE]"), run(" | "), fld("CATEGORYNAME", "[CATEGORY NAME]"),
                          run(" | "), fld("SERIESNAME", "[SERIES NAME]"), run(" | "), fld("VALUE", "[VALUE]")),
                  shown={"Val": True})
            .replace("<c:dLblPos val='r'/>", "<c:dLblPos val='outEnd'/>")
            for i in range(3)
        ) + f"<c:dLblPos val='outEnd'/>{flags(Val=True)}{SHOW_RANGE}", RANGE_EXT),
    "strref": lambda: _column(
        d_lbl(1, "<c:tx><c:strRef><c:f>Sheet1!$E$3</c:f><c:strCache><c:ptCount val='1'/>"
                 "<c:pt idx='0'><c:v>From a cell</c:v></c:pt></c:strCache></c:strRef></c:tx>")
        .replace("<c:dLblPos val='r'/>", "<c:dLblPos val='outEnd'/>")
        + f"<c:dLblPos val='outEnd'/>{flags(Val=True)}", ""),
}


def chart_xml(probe: dict) -> bytes:
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:date1904 val='0'/><c:lang val='en-US'/><c:roundedCorners val='0'/>"
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        + PROBES[probe["key"]]()
        + "</c:plotArea><c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz='1000'/></a:pPr>"
        "<a:endParaRPr lang='en-US'/></a:p></c:txPr></c:chartSpace>"
    ).encode()


def probes() -> list[dict]:
    return [{"key": key, "frame": bubble._frame(480.0, 300.0)} for key in PROBES]


def build(target: Path) -> None:
    bubble.write_deck(target, probes(), chart=chart_xml)


if __name__ == "__main__":
    build(Path(sys.argv[1]).expanduser())
    print(f"wrote {sys.argv[1]}")
