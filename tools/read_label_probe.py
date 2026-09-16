#!/usr/bin/env python3
"""Read the rotated category band back out of a label probe deck's PDF export.

For each probe slide this reports what PowerPoint actually reserved under the plot and
what it actually **drew** there: the plot's own edges, taken from the value gridlines, and
every text run below the plot with its string, its matrix and its drawn extent.

The string is the reading that matters.  PowerPoint truncates a category label it cannot
fit and the reserve follows the *drawn* label, so a run that comes back short of what the
deck authored is the cap being read directly rather than inverted through a formula.

Two traps:

* **Rotated runs overlap.**  ``FPDFText_GetRect`` groups by rectangle and the 45 degree
  labels' rectangles overlap their neighbours', so bounded-text reads pick up two labels
  at once.  Text *objects* do not overlap: ``FPDFTextObj_GetText`` gives exactly the run.
* **The gridlines are one path, not one path each**, and a column chart's bars are filled
  paths whose edges look like gridlines -- both already handled in
  :mod:`read_axis_probe`, whose stroke readers this reuses.

Usage::

    python3 tools/read_label_probe.py ~/pptx2svg-oracle/label-cap.pdf [substring]
    python3 tools/read_label_probe.py ~/pptx2svg-oracle/real-financial-report.pdf --page 3
"""

from __future__ import annotations

import ctypes
import math
import sys
from pathlib import Path

import pypdfium2 as pdfium
import pypdfium2.raw as raw

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from make_label_probe import PT, probes_for  # noqa: E402
from read_axis_probe import _matrix, _mul, horizontal_strokes, walk  # noqa: E402

from pptx2svg.resolve import chart as chartmod  # noqa: E402

SIN_45 = math.sin(math.radians(45))
#: Where `make_axis_probe.slide_xml` puts every frame, in points.
FRAME_OFF = (18.0, 12.0)


def text_runs(page):
    """Every text object as ``(text, size, angle_deg, x, y, width, height)``.

    ``x``/``y`` is the run's origin in page points, ``angle`` its baseline direction and
    ``width``/``height`` the axis-aligned bounds pdfium reports for it.
    """
    textpage = raw.FPDFText_LoadPage(page)
    out = []
    for obj, kind, parent in walk(page):
        if kind != raw.FPDF_PAGEOBJ_TEXT:
            continue
        length = raw.FPDFTextObj_GetText(obj, textpage, None, 0)
        buf = ctypes.create_string_buffer(length)
        raw.FPDFTextObj_GetText(
            obj, textpage, ctypes.cast(buf, ctypes.POINTER(ctypes.c_ushort)), length
        )
        # pdfium ends a run with a space of its own making on some exports and not on
        # others, which is two characters and 2.78 pt of phantom width in a chain that is
        # walked by advance width.  The probe labels never contain one.
        text = bytes(buf)[: length - 2].decode("utf-16-le", "replace").rstrip()
        size = ctypes.c_float()
        raw.FPDFTextObj_GetFontSize(obj, ctypes.byref(size))
        a, b, c, d, e, f = _mul(parent, _matrix(obj))
        left, bottom, right, top = (ctypes.c_float() for _ in range(4))
        raw.FPDFPageObj_GetBounds(
            obj, *[ctypes.byref(v) for v in (left, bottom, right, top)]
        )
        out.append(
            {
                "text": text,
                "size": round(size.value * math.hypot(a, b), 4),
                "angle": round(math.degrees(math.atan2(b, a)), 3),
                "x": round(e, 3),
                "y": round(f, 3),
                "left": round(left.value, 3),
                "bottom": round(bottom.value, 3),
                "right": round(right.value, 3),
                "top": round(top.value, 3),
            }
        )
    raw.FPDFText_ClosePage(textpage)
    return out


def plot_edges(page):
    """The value gridlines' ``(bottom, top, left, right)`` in page points, or ``None``."""
    lines = horizontal_strokes(page)
    if not lines:
        return None
    ys = [line[0] for line in lines]
    return (min(ys), max(ys), min(l[1] for l in lines), max(l[2] for l in lines))


def read_page(page, page_height: float, frame: tuple[float, float]) -> dict:
    """One slide: the reserve under the plot and the labels drawn in it."""
    width_pt, height_pt = frame
    frame_bottom = page_height - (FRAME_OFF[1] + height_pt)
    frame_left = FRAME_OFF[0]
    edges = plot_edges(page)
    runs = text_runs(page)
    if edges is None:
        return {"reserve": None, "runs": runs}
    bottom, top, left, right = edges
    below = [run for run in runs if run["top"] < bottom + 1.0]
    turned = [run for run in below if abs(run["angle"]) > 1.0]
    # The deepest pen position is the start of the longest drawn label: a run reading up
    # to the right starts at its own bottom-left corner, so ``drop`` is the whole of what
    # the label's own length contributes to the band, plus the fixed offset that puts the
    # label's far end under the axis.
    deepest = min((run["y"] for run in turned), default=None)
    return {
        "reserve": round(bottom - frame_bottom, 3),
        "axis": round(bottom, 3),
        "drop": round(bottom - deepest, 3) if deepest is not None else None,
        "pad": round(deepest - frame_bottom, 3) if deepest is not None else None,
        "turned": len(turned),
        # The band's own labels are the turned ones: a bottom legend also sits under the
        # plot, and joining its entries to the chain reads a label that was never drawn.
        "runs": turned,
        "left_inset": round(left - frame_left, 3),
        "right_inset": round(frame_left + width_pt - right, 3),
        "top_inset": round(page_height - FRAME_OFF[1] - top, 3),
        "plot": (round(left, 3), round(right, 3), round(bottom, 3), round(top, 3)),
        "below": below,
        "all_runs": runs,
    }


def group_runs(runs: list[dict], face: str, size: float) -> list[dict]:
    """Join the text objects that make up one label.

    PowerPoint draws the ellipsis of a truncated label as its **own** text object, whose
    pen starts exactly where the kept text ends, so a label is a chain of runs each
    beginning where the last one stopped.  The chain is walked with our own advance
    widths, which for the Latin probe alphabets are the drawn face's exactly.

    **The y step is what keeps the chain inside one label.**  Matching on x alone joins a
    label's ellipsis to its neighbour's first run whenever the neighbour happens to start
    a label width away, which silently reports a label two characters longer than the one
    PowerPoint drew.  A continuation is up *and* to the right by the same amount; the next
    label starts lower again.
    """
    labels: list[dict] = []
    for run in sorted(runs, key=lambda item: item["x"]):
        # Every open chain is a candidate, not just the last one: a label long enough to
        # reach past its neighbour's start has its ellipsis sorted after that neighbour,
        # and matching only the last chain leaves the ellipsis stranded as a label of its
        # own -- which reads as a one-character allowance.
        for chain in labels:
            reach = chartmod.text_width(chain["text"], face, size) * SIN_45
            if (
                abs(run["x"] - (chain["x"] + reach)) < 2.0
                and abs(run["y"] - (chain["y"] + reach)) < 2.0
            ):
                chain["text"] += run["text"]
                break
        else:
            labels.append(dict(run))
    return labels


def budget(
    drawn: list[str], authored: tuple[str, ...], face: str, size: float
) -> tuple[float, float]:
    """The bracket this chart puts on the width a category label is allowed.

    A label whose own width fits is drawn whole, which says only that the allowance is at
    least that wide.  A label that is cut says both halves: the kept prefix **and its
    ellipsis** fit, and one more character of the same prefix does not.  The ellipsis is
    inside the allowance -- ``f36``/``f37`` separate that from a rule on the prefix alone
    by three characters.

    Every label on the chart is one reading and they are intersected, which matters on a
    chart whose labels differ: ``real-financial-report``'s chart3 is bracketed from below
    by the five-character label it kept and from above by the eight-character one it cut.
    """
    ellipsis = chartmod.text_width("…", face, size)
    low, high = 0.0, math.inf
    for text in drawn:
        if not text.endswith("…"):
            # A run that is not cut must be a whole label; anything else is a chain that
            # failed to join, and reading it as a cut label invents an allowance far
            # below the real one.
            if text in authored:
                low = max(low, chartmod.text_width(text, face, size))
            continue
        kept = text.rstrip("…")
        sources = [item for item in authored if item.startswith(kept) and item != kept]
        if not sources:
            continue
        source = max(sources, key=len)
        low = max(low, chartmod.text_width(kept, face, size) + ellipsis)
        high = min(
            high,
            chartmod.text_width(source[: len(kept) + 1], face, size) + ellipsis,
        )
    return (low, high)


def predict(item: dict, read: dict) -> tuple[float, str, str]:
    """What the shipped rule says this chart reserves and draws.

    Built out of :mod:`pptx2svg.resolve.chart`'s own constants and functions rather than a
    copy of them, so a change to either shows up here as a disagreement.  Only the Latin
    probes get a verdict: the probe deck names no East Asian face, so which one PowerPoint
    fell back to -- and what its ascent is -- is not something this side can know.
    """
    face = item["face"] or "Aptos"
    size = item["size"] / 100
    font = chartmod.ChartFont(family=face, box=chartmod.font_box(face, size))
    height = item["frame"][1] / PT
    if item.get("legend") in ("b", "t", "tr"):
        height -= chartmod.LEGEND_BAND_LINES * font.box.line_height
    if item.get("title"):
        height -= chartmod.TITLE_BAND_LINES * chartmod.font_box("Arial", 18.0).line_height
    allowance = max(
        (height / 2 - chartmod.ROTATED_LABEL_HEADROOM_PT - chartmod.rotated_label_anchor(font.box))
        / SIN_45,
        0.0,
    )
    drawn = [chartmod.truncate_label(text, font, allowance) for text in item["labels"]]
    widest = max(font.width(text.rstrip("…")) for text in drawn)
    band = (
        chartmod.FRAME_PADDING_PT
        + (font.box.line_height + widest) * SIN_45
        + chartmod.CATEGORY_LABEL_GAP_EM * font.box.size
    )
    return band, max(drawn, key=len), f"allowance={allowance:.2f}"


def describe(read: dict, face: str, size: float, authored: str) -> str:
    """What the drawn labels say about the cap, in one line."""
    if read["reserve"] is None:
        return "no gridlines"
    if not read["runs"]:
        return "no labels below the plot"
    labels = group_runs(read["runs"], face, size)
    drawn = max((label["text"] for label in labels), key=len)
    kept = drawn.rstrip("….")
    width = chartmod.text_width(kept, face, size)
    full = chartmod.text_width(authored, face, size)
    return (
        f"drawn={drawn!r} kept={len(kept)}/{len(authored)} chars "
        f"{width:.3f} pt of {full:.3f}  drop={read['drop']} pad={read['pad']}"
    )


def main() -> int:
    path = Path(sys.argv[1]).expanduser()
    doc = pdfium.PdfDocument(path)
    if len(sys.argv) > 2 and sys.argv[2] == "--page":
        index = int(sys.argv[3]) - 1
        page = doc[index]
        for run in sorted(text_runs(page.raw), key=lambda r: (-r["top"], r["left"])):
            print(run)
        print("gridlines", horizontal_strokes(page.raw))
        print("page", page.get_size())
        return 0
    only = sys.argv[2] if len(sys.argv) > 2 else None
    rows = []
    for index, item in enumerate(probes_for(path)):
        if only and only not in item["key"]:
            continue
        page = doc[index]
        page_height = page.get_size()[1]
        frame = tuple(dimension / PT for dimension in item["frame"])
        read = read_page(page.raw, page_height, frame)
        face = item["face"] or "Aptos"
        size = item["size"] / 100
        authored = max(item["labels"], key=len)
        print(
            f"=== {index + 1:3d} {item['key']:14s} frame={frame[0]:.1f}x{frame[1]:.1f} "
            f"size={size:g} face={face} cats={len(item['labels'])} "
            f"reserve={read['reserve']} left={read.get('left_inset')}"
        )
        print(f"    {describe(read, face, size, authored)}")
        angles = sorted({run["angle"] for run in read["runs"]})
        print(f"    angles={angles} labels={[run['text'] for run in read['runs']]}")
        if read["reserve"] is not None and read["runs"]:
            band, drawn, note = predict(item, read)
            labels = group_runs(read["runs"], face, size)
            theirs = max((label["text"] for label in labels), key=len, default="")
            # The legend's own band is inside the reserve but outside the label band.
            reserve = read["reserve"]
            if item.get("legend") == "b":
                reserve -= chartmod.LEGEND_BAND_LINES * chartmod.font_box(face, size).line_height
            cjk = any(ord(ch) > 0x2E80 for ch in theirs)
            verdict = (
                "cjk, no verdict"
                if cjk
                else "MATCH"
                if abs(band - reserve) < 0.4 and drawn == theirs
                else "*** DIFFERS ***"
            )
            print(
                f"    ours     band={band:.3f} vs {reserve:.3f}  {drawn!r} vs {theirs!r}"
                f"  {note}  {verdict}"
            )
        rows.append((item, read))
    print(
        "\n# key\tframe_w\tframe_h\tsize\tface\tcats\tband\tplot_h\treserve\tdrop\tpad"
        "\tkept\tkept_pt\tfull_pt\tbudget_lo\tbudget_hi"
    )
    for item, read in rows:
        frame = tuple(dimension / PT for dimension in item["frame"])
        face = item["face"] or "Aptos"
        size = item["size"] / 100
        labels = group_runs(read["runs"], face, size) if read["runs"] else []
        drawn = max((label["text"] for label in labels), key=len, default="")
        kept = drawn.rstrip("….")
        plot = read.get("plot")
        band = (plot[1] - plot[0]) / len(item["labels"]) if plot else 0.0
        # The label the band is sized from is the widest, which for the probes whose five
        # labels differ only in a trailing letter is not the first.
        authored = max(item["labels"], key=lambda text: chartmod.text_width(text, face, size))
        low, high = budget([label["text"] for label in labels], item["labels"], face, size)
        print(
            f"{item['key']}\t{frame[0]:.4f}\t{frame[1]:.4f}\t{size:g}\t{face}\t"
            f"{len(item['labels'])}\t{band:.3f}\t{plot[3] - plot[2] if plot else 0:.3f}\t"
            f"{read['reserve']}\t{read['drop']}\t{read['pad']}\t"
            f"{len(kept)}\t{chartmod.text_width(kept, face, size):.3f}\t"
            f"{chartmod.text_width(authored, face, size):.3f}\t{low:.3f}\t{high:.3f}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
