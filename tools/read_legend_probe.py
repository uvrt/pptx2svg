#!/usr/bin/env python3
"""Read a legend probe deck's PDF export back as legend geometry.

The quantity every slide is here to yield is the pitch between consecutive legend **keys**
-- the swatch or rule that precedes each name.  Keys are paths, so the pitch comes off the
export with no font metric in it, which matters because the thing under test is a gap
measured in ems.

For entry *i* with name advance ``a(i)``::

    pitch(i) = key_width + key_gap + a(i) + G

so ``K = pitch(i) - a(i)`` is the one number a slide determines, and ``K`` is constant
within a slide exactly when the gap is.  ``G`` is then ``K`` less the key advance, and the
key advance is measured here too: the swatch's own drawn width, and the distance from the
key to its label.

ROADMAP.md 3.5 records what these readings settled.  Two warnings for anyone reading the
raw table rather than ``--check``:

* the **key advance is not the drawn key**.  The layout cell is twice the swatch and
  1.25x the line rule, with the key centred in it, so the ``G`` column below -- which
  subtracts only what is drawn -- is about 0.31 em larger than the gap in the rule.
* ``key_gap`` is the distance to the label's **ink**, which carries a left side bearing
  and drifts by a hundredth of an em across font sizes.  ``K`` is the number to trust.

What each column says:

``K``          ``pitch - advance``, per consecutive pair.  Its **spread** within a slide is
               the check that the layout is a constant gap at all.
``G``          ``K - key_width - key_gap``, in ems -- see the warning above.
``lead``       the first key's x from the frame's left edge.
``trail``      the frame's right edge less the last label's ink right.
               ``lead`` and ``trail`` agreeing is the run being **centred**; a fixed
               ``lead`` that does not move with the content is the run being **pinned**,
               which is what a distributed layout would look like.

Usage::

    python3 tools/read_legend_probe.py ~/pptx2svg-oracle/legend-pack.pdf [substring]
    python3 tools/read_legend_probe.py ~/pptx2svg-oracle/legend-pack.pdf --check [substring]

``--check`` renders the same deck through this library and prints the residual between our
key positions and PowerPoint's, which is the assertion SSIM cannot make.
"""

from __future__ import annotations

import ctypes
import statistics
import sys
from pathlib import Path

import pypdfium2 as pdfium
import pypdfium2.raw as raw

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from make_legend_probe import FRAME_OFF, probes_for  # noqa: E402
from read_axis_probe import _mul, chartmod, labels, path_points  # noqa: E402


def _matrix(obj):
    m = raw.FS_MATRIX()
    raw.FPDFPageObj_GetMatrix(obj, ctypes.byref(m))
    return (m.a, m.b, m.c, m.d, m.e, m.f)


def walk(page_or_form, parent=(1, 0, 0, 1, 0, 0), form=False):
    if form:
        count = raw.FPDFFormObj_CountObjects(page_or_form)

        def get(i):
            return raw.FPDFFormObj_GetObject(page_or_form, i)
    else:
        count = raw.FPDFPage_CountObjects(page_or_form)

        def get(i):
            return raw.FPDFPage_GetObject(page_or_form, i)

    for i in range(count):
        obj = get(i)
        kind = raw.FPDFPageObj_GetType(obj)
        if kind == raw.FPDF_PAGEOBJ_FORM:
            yield from walk(obj, _mul(parent, _matrix(obj)), form=True)
        else:
            yield obj, kind, parent


def page_paths(page) -> list[tuple[float, float, float, float]]:
    """Every path object's bounding box, in page points."""
    out = []
    for obj, kind, parent in walk(page):
        if kind != raw.FPDF_PAGEOBJ_PATH:
            continue
        points = path_points(obj, parent)
        if not points:
            continue
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        out.append((min(xs), min(ys), max(xs), max(ys)))
    return out


def entry_names(probe: dict) -> list[str]:
    """The legend entries this probe should produce, in paint order.

    Paint order is ``areaChart -> barChart -> lineChart`` (ROADMAP.md 3.3), which for a
    combo is not the document order.  A pie legends its **categories**.
    """
    order = {"area": 0, "col": 1, "scatter": 1, "pie": 1, "line": 2, "stock": 2}
    names: list[str] = []
    for group in sorted(probe["groups"], key=lambda g: order[g["kind"]]):
        names.extend(group["names"])
    return names


def read_page(probe: dict, page) -> dict | None:
    """The legend's drawn geometry: one key x and one label box per entry."""
    wanted = entry_names(probe)
    runs = [row for row in labels(page) if row[0] in wanted]
    if len(runs) < 2:
        return None
    # The legend band is the row those labels share.  A category label that happens to
    # repeat a legend name would sit on a different line and is dropped here.
    band = statistics.median(row[2] for row in runs)
    runs = [row for row in runs if abs(row[2] - band) < 4.0]
    runs.sort(key=lambda row: row[1])
    if len(runs) != len(wanted):
        return None

    paths = page_paths(page)
    # A key is a small path sharing the label's line.  The plot's bars and gridlines are
    # far taller or far wider; the bound is generous because a line key is 19.2 pt.
    keys = [
        box
        for box in paths
        if abs((box[1] + box[3]) / 2 - band) < 7.0
        and box[2] - box[0] < 40.0
        and box[3] - box[1] < 12.0
    ]
    out_keys: list[tuple[float, float]] = []
    for index, run in enumerate(runs):
        left_limit = runs[index - 1][1] + runs[index - 1][4] if index else -1e9
        mine = [
            box for box in keys if box[2] <= run[1] + 0.5 and box[0] > left_limit
        ]
        if not mine:
            return None
        out_keys.append((min(b[0] for b in mine), max(b[2] for b in mine)))

    return {
        "names": [row[0] for row in runs],
        "key_x": [round(k[0], 3) for k in out_keys],
        "key_right": [round(k[1], 3) for k in out_keys],
        "text_left": [round(row[1], 3) for row in runs],
        "text_right": [round(row[1] + row[4], 3) for row in runs],
        "band_y": round(band, 3),
    }


# --------------------------------------------------------------------------------------
# What this library computes for the same deck
# --------------------------------------------------------------------------------------


def ours(deck: Path) -> list[dict]:
    """Per chart: the frame, the legend font, and each entry's name and advance width.

    Collected by instrumenting the resolver rather than by parsing the SVG, because the
    advance width is the number wanted and the SVG only carries a position.
    """
    from pptx2svg.resolve import chart as chartmod

    rows: list[dict] = []
    original = chartmod.ChartBuilder._draw_legend

    def record(self, rect, series, *, per_point=False, categories=None, second_band=0.0):
        position = self._legend_position()
        if position in ("b", "t", "tr") and series:
            font = self._legend_font()
            swatch, gap = self._legend_key_size(font)
            if per_point:
                names = [name for name in (categories or []) if name]
            else:
                names = [item.name for item in series if item.name]
            rows.append(
                {
                    # The resolver works in the frame's own space; these decks put every
                    # frame at the same slide offset, which is what the export measures
                    # against.
                    "frame_left": self.frame.left + FRAME_OFF[0] / 12700,
                    "frame_width": self.frame.width,
                    "size": font.box.size,
                    "swatch": swatch,
                    "key_gap": gap,
                    "cell": self._legend_key_cell(font),
                    "names": names,
                    "advance": [font.width(name) for name in names],
                    "position": position,
                }
            )
        return original(
            self, rect, series, per_point=per_point, categories=categories,
            second_band=second_band,
        )

    chartmod.ChartBuilder._draw_legend = record
    try:
        import pptx2svg

        pptx2svg.convert_pptx_to_svg(deck.read_bytes())
    finally:
        chartmod.ChartBuilder._draw_legend = original
    return rows


def solve(read: dict, mine: dict) -> dict:
    """``K``, ``G`` and the run's margins, from one slide's measured geometry."""
    key_x = read["key_x"]
    advance = mine["advance"]
    size = mine["size"]
    pitches = [b - a for a, b in zip(key_x, key_x[1:])]
    ks = [pitch - adv for pitch, adv in zip(pitches, advance[:-1])]
    key_width = statistics.median(r - x for x, r in zip(key_x, read["key_right"]))
    key_gap = statistics.median(
        t - r for t, r in zip(read["text_left"], read["key_right"])
    )
    k_mean = statistics.mean(ks) if ks else float("nan")
    frame_right = mine["frame_left"] + mine["frame_width"]
    return {
        "pitches": [round(p, 3) for p in pitches],
        "K": [round(k, 3) for k in ks],
        "K_mean": k_mean,
        "K_spread": (max(ks) - min(ks)) if ks else 0.0,
        "key_width": key_width,
        "key_gap": key_gap,
        "G_em": (k_mean - key_width - key_gap) / size,
        "lead": key_x[0] - mine["frame_left"],
        "trail": frame_right - read["text_right"][-1],
        "span": read["text_right"][-1] - key_x[0],
    }


def report(path: Path, only: str | None) -> int:
    probes = probes_for(path)
    doc = pdfium.PdfDocument(path)
    mine_all = ours(path.with_suffix(".pptx"))
    print(
        f"{'key':12s} {'frame':>6s} {'n':>2s} {'size':>4s} {'keyw':>6s} {'kgap':>5s} "
        f"{'K':>7s} {'spread':>6s} {'G_em':>6s} {'lead':>6s} {'trail':>6s} {'span':>7s}"
    )
    rows = []
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page = doc[index]
        read = read_page(probe, page.raw)
        mine = mine_all[index] if index < len(mine_all) else None
        if read is None or mine is None:
            print(f"{probe['key']:12s}  -- not read --")
            continue
        if read["names"] != mine["names"]:
            print(f"{probe['key']:12s}  names differ: {read['names']} vs {mine['names']}")
            continue
        answer = solve(read, mine)
        rows.append((probe, read, mine, answer))
        print(
            f"{probe['key']:12s} {mine['frame_width']:6.1f} {len(read['names']):2d} "
            f"{mine['size']:4.0f} {answer['key_width']:6.3f} {answer['key_gap']:5.3f} "
            f"{answer['K_mean']:7.3f} {answer['K_spread']:6.3f} {answer['G_em']:6.3f} "
            f"{answer['lead']:6.2f} {answer['trail']:6.2f} {answer['span']:7.2f}"
        )
    return 0


def check(path: Path, only: str | None) -> int:
    """Our drawn key positions against PowerPoint's, entry by entry."""
    probes = probes_for(path)
    doc = pdfium.PdfDocument(path)
    mine_all = ours(path.with_suffix(".pptx"))
    from pptx2svg.resolve import chart as chartmod

    worst = 0.0
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page = doc[index]
        read = read_page(probe, page.raw)
        mine = mine_all[index] if index < len(mine_all) else None
        if read is None or mine is None:
            print(f"{probe['key']:12s}  -- not read --")
            continue
        cell = mine["cell"]
        widths = [cell + adv for adv in mine["advance"]]
        total = sum(widths)
        slack = min(
            chartmod.LEGEND_ENTRY_SLACK * total,
            chartmod.LEGEND_BAND_MAX_FRACTION * mine["frame_width"] - total,
        )
        gap = max(slack, 0.0) / (len(widths) + 1)
        run = total + gap * (len(widths) - 1)
        x = (
            mine["frame_left"]
            + (mine["frame_width"] - run) / 2
            + chartmod.LEGEND_HORIZONTAL_OFFSET_PT
            + (cell - mine["swatch"]) / 2
        )
        deltas = []
        for measured, width in zip(read["key_x"], widths):
            deltas.append(x - measured)
            x += width + gap
        worst = max(worst, max(abs(d) for d in deltas))
        flag = "" if max(abs(d) for d in deltas) < 0.5 else "   <<<"
        pretty = " ".join(f"{d:+7.2f}" for d in deltas)
        print(f"{probe['key']:12s} key dx {pretty}{flag}")
    print(f"\nworst legend key residual {worst:.2f} pt")
    return 0


def read_text_row(probe: dict, page) -> dict | None:
    """The legend's **labels** only, for a legend whose keys may not be drawn at all.

    :func:`read_page` needs a key path per entry and so cannot read a chart whose keys are
    invisible, which is the whole subject of the ``legend-nokey`` deck.  Every entry there
    is named ``W``, ``Wm``, ``Wmm``... so each label's ink starts at the same left side
    bearing and the pitch between two labels is the pitch between two layout cells with no
    font residue in it.
    """
    wanted = entry_names(probe)
    runs = [row for row in labels(page) if row[0] in wanted]
    if len(runs) < len(wanted):
        return None
    band = statistics.median(row[2] for row in runs)
    runs = [row for row in runs if abs(row[2] - band) < 4.0]
    runs.sort(key=lambda row: row[1])
    if len(runs) != len(wanted):
        return None
    paths = page_paths(page)
    keys = [
        box
        for box in paths
        if abs((box[1] + box[3]) / 2 - band) < 7.0
        and box[2] - box[0] < 40.0
        and box[3] - box[1] < 12.0
        and box[0] < runs[-1][1]
    ]
    return {
        "names": [row[0] for row in runs],
        "text_left": [round(row[1], 3) for row in runs],
        "text_right": [round(row[1] + row[4], 3) for row in runs],
        "keys": [(round(b[0], 3), round(b[2], 3)) for b in sorted(keys)],
        "band_y": round(band, 3),
    }


def solve_cell(read: dict, mine: dict) -> dict:
    """The one unknown a ``legend-nokey`` slide leaves: the width of the entry's key cell.

    3.5's layout is ``W(i) = cell + advance(i)``, ``gap = 0.2 * sum(W) / (n + 1)`` and the
    run centred on the frame plus 0.75 pt.  Hold that rule and let ``cell`` float, and one
    slide determines it twice over:

    * from the **pitch** between two labels, which is ``advance(i) + gap + cell``;
    * from the **lead**, the first label's ink less the run's computed start, which must
      come back as one left side bearing shared by every slide in the deck.

    The two agreeing is what says the layout is the measured one with a different cell,
    rather than a different layout.
    """
    advance = mine["advance"]
    text = read["text_left"]
    n = len(advance)
    frame_width = mine["frame_width"]
    pitches = [b - a for a, b in zip(text, text[1:])]
    # `pitch(i) - advance(i)` is `cell + gap`, one reading per consecutive pair.
    residues = [p - a for p, a in zip(pitches, advance[:-1])]
    total_advance = sum(advance)
    mean = statistics.mean(residues) if residues else float("nan")
    # cell + 0.2 * (n * cell + sum advance) / (n + 1) = mean
    cell = (mean - chartmod.LEGEND_ENTRY_SLACK * total_advance / (n + 1)) / (
        1.0 + chartmod.LEGEND_ENTRY_SLACK * n / (n + 1)
    )

    def lead_for(candidate: float) -> float:
        total = n * candidate + total_advance
        slack = min(
            chartmod.LEGEND_ENTRY_SLACK * total,
            chartmod.LEGEND_BAND_MAX_FRACTION * frame_width - total,
        )
        gap = max(slack, 0.0) / (n + 1)
        run = total + gap * (n - 1)
        start = (
            mine["frame_left"]
            + (frame_width - run) / 2
            + chartmod.LEGEND_HORIZONTAL_OFFSET_PT
        )
        return text[0] - (start + candidate)

    return {
        "residues": [round(r, 3) for r in residues],
        "spread": (max(residues) - min(residues)) if residues else 0.0,
        "cell": cell,
        "cell_em": cell / mine["size"],
        "lsb_fitted": lead_for(cell),
        "lsb_swatch": lead_for(chartmod.LEGEND_ENTRY_KEY_EM * mine["size"]),
        "lsb_line": lead_for(chartmod.LINE_LEGEND_ENTRY_KEY_PT),
        "lsb_none": lead_for(0.0),
        "keys": len(read["keys"]),
    }


def cells(path: Path, only: str | None) -> int:
    """Solve each slide for its legend key **cell**, the ``legend-nokey`` deck's subject."""
    probes = probes_for(path)
    doc = pdfium.PdfDocument(path)
    mine_all = ours(path.with_suffix(".pptx"))
    print(
        f"{'key':12s} {'frame':>6s} {'n':>2s} {'keys':>4s} {'cell':>7s} {'em':>6s} "
        f"{'spread':>6s} {'lsb@fit':>7s} {'lsb@sw':>7s} {'lsb@ln':>7s} {'lsb@0':>7s}  residues"
    )
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page = doc[index]
        read = read_text_row(probe, page.raw)
        mine = mine_all[index] if index < len(mine_all) else None
        if read is None or mine is None:
            print(f"{probe['key']:12s}  -- not read --")
            continue
        if mine["position"] not in ("b", "t"):
            print(f"{probe['key']:12s}  {mine['position']} legend: text {read['text_left']} "
                  f"keys {read['keys']}")
            continue
        answer = solve_cell(read, mine)
        print(
            f"{probe['key']:12s} {mine['frame_width']:6.1f} {len(read['names']):2d} "
            f"{answer['keys']:4d} {answer['cell']:7.3f} {answer['cell_em']:6.4f} "
            f"{answer['spread']:6.3f} {answer['lsb_fitted']:7.3f} {answer['lsb_swatch']:7.3f} "
            f"{answer['lsb_line']:7.3f} {answer['lsb_none']:7.3f}  {answer['residues']}"
        )
    return 0


def main() -> int:
    path = Path(sys.argv[1]).expanduser()
    arguments = sys.argv[2:]
    if "--cells" in arguments:
        arguments.remove("--cells")
        return cells(path, arguments[0] if arguments else None)
    if "--check" in arguments:
        arguments.remove("--check")
        return check(path, arguments[0] if arguments else None)
    return report(path, arguments[0] if arguments else None)


if __name__ == "__main__":
    raise SystemExit(main())
