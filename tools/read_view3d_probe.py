#!/usr/bin/env python3
"""Read a 3-D probe deck's PDF export: the axis PowerPoint chose and where it drew it.

PowerPoint rasterises 3-D chart geometry (ROADMAP.md 3.4) but leaves every string vector,
so the default reading here is nothing but text -- ``--colours`` and ``--shapes`` read the
raster instead, and are documented below.  The value axis' tick labels give

* the **axis** -- its ends and its unit, hence the interval count that produced it, and
* the **plot rectangle** -- the extreme ticks sit on the plot's own top and bottom edges,
  and the category labels' centres bracket its left and right ones.

A label's centre is not the tick, though, and about a point of PDF text-rect convention
rides on every inset read this way: ``tools/read_axis_probe.py --insets`` reads the same
rectangle off the **gridlines** and puts a flat 195 pt frame's insets at 11.103 / 24.965
where the flat controls here read 10.010 / 25.870.  That point is a constant of the
instrument and cancels in a *difference* of two readings, which is what the depth
reservation is -- but it does not cancel in an inset quoted on its own, and quoting one
on its own is how ROADMAP.md 3.4 came to record a plot rectangle 3.7 pt low that was
never there.  See :func:`ours`.

That is the whole of the depth reservation: the difference between those insets and the
ones a flat chart of the same data on the same frame gets.  ``--check`` renders the deck
through this library and prints that difference per probe.

Usage::

    python3 tools/read_view3d_probe.py ~/pptx2svg-oracle/view3d-meter.pdf [substring]
    python3 tools/read_view3d_probe.py ~/pptx2svg-oracle/view3d-meter.pdf --check

``--colours`` and ``--shapes`` read the **raster** instead, which is the other half of
what a 3-D slide holds: the first reports every drawn colour that is one of the stated
fills scaled, which is a prism's three faces and their shading factors, and the second
reports the scene image's own box and the ink inside it, which is the only reading a
``pie3DChart`` offers at all::

    python3 tools/read_view3d_probe.py ~/pptx2svg-oracle/view3d-colour.pdf --colours
    python3 tools/read_view3d_probe.py ~/pptx2svg-oracle/view3d-shape.pdf --shapes

``--mesh`` reads the **solid inside** the scene rather than the scene's own box: each
prism's front face against the front plane its value axis defines, and its depth off its
own right face.  That is what a renderer drawing the scene needs and what the box cannot
say -- where in its row of depth a bar stands::

    python3 tools/read_view3d_probe.py ~/pptx2svg-oracle/view3d-mesh.pdf --mesh
"""

from __future__ import annotations

import collections
import sys
from pathlib import Path

import pypdfium2 as pdfium

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))
sys.path.insert(0, str(HERE.parent / "packages/pptx2svg-fonts/src"))

from make_axis_probe import FRAME_OFF  # noqa: E402
from make_view3d_probe import probes_for  # noqa: E402
from read_axis_probe import _number, labels, unit_from  # noqa: E402

EMU = 12700.0


def read_page(page, probe: dict) -> dict:
    """One probe slide's axis and plot rectangle, in page points."""
    text = labels(page)
    ticks, categories = [], []
    for label, left, y, _height, width, x in text:
        try:
            ticks.append((_number(label), y, width))
        except ValueError:
            categories.append((label, x, y))
    ticks.sort(key=lambda row: row[1])
    categories.sort(key=lambda row: row[1])
    values = [row[0] for row in ticks]
    positions = [row[1] for row in ticks]
    xs = [row[1] for row in categories]
    band = (xs[1] - xs[0]) if len(xs) > 1 else 0.0
    return {
        "values": values,
        "unit": unit_from(values),
        "bottom": positions[0] if positions else 0.0,
        "top": positions[-1] if positions else 0.0,
        "height": round(positions[-1] - positions[0], 3) if len(positions) > 1 else 0.0,
        "left": round(xs[0] - band / 2, 3) if xs else 0.0,
        "right": round(xs[-1] + band / 2, 3) if xs else 0.0,
        "cat_y": round(categories[0][2], 3) if categories else 0.0,
    }


def insets(read: dict, probe: dict, page_height: float) -> dict:
    """The plot rectangle as four insets from the chart frame's own edges, in points."""
    frame_top = page_height - FRAME_OFF[1] / EMU
    frame_bottom = frame_top - probe["frame"][1] / EMU
    frame_left = FRAME_OFF[0] / EMU
    frame_right = frame_left + probe["frame"][0] / EMU
    return {
        "top": round(frame_top - read["top"], 3),
        "bottom": round(read["bottom"] - frame_bottom, 3),
        "left": round(read["left"] - frame_left, 3),
        "right": round(frame_right - read["right"], 3),
        "height": read["height"],
        "width": round(read["right"] - read["left"], 3),
    }


def ours(deck: Path, probes: list[dict]) -> list[dict]:
    """What this library draws for the same deck: the same readings, same units.

    The library's own model is read rather than its SVG, so the numbers are the layout's
    and not a rasteriser's: a chart's tick labels are ordinary text elements inside the
    chart frame.

    **What is compared is the label's ink, not its box**, and that distinction cost this
    file a phantom.  Our text element is a box ``1.5`` line heights tall whose baseline
    sits near its *top* (:meth:`ChartBuilder._text`), so its centre is some 2.7 pt below
    the tick it marks at 10 pt, where the PDF text rect pdfium reports for the same label
    is the digits' own ink and sits on the tick.  Differencing the two read a 3.7 pt
    displacement of the whole plot rectangle -- +3.8 at the top, -3.7 at the bottom, the
    height right -- into a layout that is in fact correct: ``axis-inset``'s gridlines put
    our plot's top edge within 0.002 pt of PowerPoint's at eight sizes in four faces.
    So the box centre is converted to the ink centre here, which is the quantity the
    export's own reading is.
    """
    from pptx2svg import ConvertOptions, convert_pptx_to_model
    from pptx2svg import model as m
    from pptx2svg.resolve.chart import font_box

    presentation = convert_pptx_to_model(deck, ConvertOptions())
    page_height = presentation.slide_size.height / EMU
    rows = []
    for probe, slide in zip(probes, presentation.slides):
        ticks, categories = [], []

        def walk(elements, ox=0.0, oy=0.0):
            for element in elements:
                if isinstance(element, m.ChartElement):
                    walk(
                        element.children,
                        ox + element.transform.offset_x / EMU,
                        oy + element.transform.offset_y / EMU,
                    )
                elif isinstance(element, m.GroupElement):
                    walk(element.children, ox, oy)
                elif isinstance(element, m.ShapeElement) and element.text_body is not None:
                    text = "".join(
                        run.text
                        for paragraph in element.text_body.paragraphs
                        for run in paragraph.runs
                    ).strip()
                    if not text:
                        continue
                    transform = element.transform
                    run = next(
                        run.properties
                        for paragraph in element.text_body.paragraphs
                        for run in paragraph.runs
                        if run.text.strip()
                    )
                    box = font_box(run.font_family, run.font_size or 10.0)
                    x = ox + transform.offset_x / EMU + transform.extent_width / EMU / 2
                    # The box's top plus our own first-baseline rule is the baseline; the
                    # ink sits `ink_centre` above it.  See the docstring.
                    y = page_height - (
                        oy
                        + transform.offset_y / EMU
                        + box.first_baseline
                        - box.ink_centre
                    )
                    try:
                        ticks.append((_number(text), y))
                    except ValueError:
                        categories.append((text, x, y))

        walk(slide.elements)
        ticks.sort(key=lambda row: row[1])
        categories.sort(key=lambda row: row[1])
        xs = [row[1] for row in categories]
        band = (xs[1] - xs[0]) if len(xs) > 1 else 0.0
        positions = [row[1] for row in ticks]
        rows.append(
            {
                "values": [row[0] for row in ticks],
                "unit": unit_from([row[0] for row in ticks]),
                "bottom": positions[0] if positions else 0.0,
                "top": positions[-1] if positions else 0.0,
                "height": round(positions[-1] - positions[0], 3) if len(positions) > 1 else 0.0,
                "left": round(xs[0] - band / 2, 3) if xs else 0.0,
                "right": round(xs[-1] + band / 2, 3) if xs else 0.0,
            }
        )
    return rows


def solved_intervals(read: dict, high: float, strict: bool) -> set[int]:
    """Every interval count that would draw the axis this probe drew.

    The axis is a product of two rules -- how many intervals there is room for, and what
    range they divide -- and only the drawn unit and ends are observable.  So *strict*
    picks the range rule and this reports the counts consistent with it: an empty set says
    the drawn axis is not reachable at any count, which is the rule being refuted.
    """
    from pptx2svg.resolve import chart as chartmod

    values, unit = read["values"], read["unit"]
    if not values or not isinstance(unit, float):
        return set()
    out = set()
    for count in range(1, chartmod.AXIS_MAX_INTERVALS + 1):
        minimum, maximum, step = chartmod.nice_axis_scale(
            0.0, high, intervals=count, strict=strict
        )
        if (
            abs(step - unit) < unit * 1e-6
            and abs(minimum - values[0]) < unit * 1e-6
            and abs(maximum - values[-1]) < unit * 1e-6
        ):
            out.add(count)
    return out


def solve(path: Path, probes: list[dict]) -> None:
    """Per probe, the counts each range rule allows; per cell, what they agree on.

    A "cell" is the probes sharing a frame and a view -- the N-meter, whose datasets'
    drawn units name one count between them.  The rule that survives is the one whose
    cells intersect non-empty.
    """
    doc = pdfium.PdfDocument(path)
    cells: dict[tuple, dict[str, set[int]]] = {}
    for index, probe in enumerate(probes):
        read = read_page(doc[index].raw, probe)
        view = probe["view"]
        cell = (
            probe["kind"],
            probe["frame"][1],
            "absent" if view is None else tuple(sorted(view.items())),
        )
        strictly = solved_intervals(read, probe["high"], strict=True)
        barely = solved_intervals(read, probe["high"], strict=False)
        row = cells.setdefault(cell, {"padded": set(range(1, 11)), "bare": set(range(1, 11)),
                                      "height": read["height"], "keys": []})
        row["padded"] &= strictly
        row["bare"] &= barely
        row["keys"].append(probe["key"])
        print(
            f"{probe['key']:20s} {_axis(read):22s} h={read['height']:8.3f}  "
            f"padded={sorted(strictly)}  bare={sorted(barely)}"
        )
    print("\n# cell\tplot_h\tpadded_N\tbare_N\tprobes")
    for cell, row in cells.items():
        kind, frame, view = cell
        shown = view if isinstance(view, str) else ",".join(f"{k}={v}" for k, v in view)
        print(
            f"{kind}\t{frame / EMU:.0f}\t{shown}\t{row['height']:.3f}\t"
            f"{sorted(row['padded'])}\t{sorted(row['bare'])}\t{len(row['keys'])}"
        )


def _axis(read: dict) -> str:
    values, unit = read["values"], read["unit"]
    if not values:
        return "(no ticks)"
    return f"{values[0]:g}..{values[-1]:g} by {unit}"


# --------------------------------------------------------------------------- #
# The scene itself, read off the raster PowerPoint draws it as.
# --------------------------------------------------------------------------- #

#: A colour has to cover this many of the scene raster's pixels to be a face rather than
#: an antialiasing artefact.  The scenes here run about 2600 x 1200 px, and the smallest
#: real face -- a top face at five degrees of pitch -- covers some thousands.
MIN_FACE_PIXELS = 300

#: How far a drawn colour may sit off ``factor * fill`` and still be called that face, in
#: 8-bit levels.  The export shifts a channel by one either way (``FF3300`` comes back
#: ``FF3200``), so this has to clear one level and nothing like the 5.8 that separate a
#: linear-light scaling from an sRGB one on a dark channel.
FACE_TOLERANCE = 2.5

#: How far the front face's additive lift may run before a colour is something else's.
#: The largest lift measured is eleven levels, at forty-five degrees of negative pitch.
MAX_FRONT_LIFT = 20.0


def scene_images(page) -> list:
    """The image objects on a page.  A 3-D chart's scene is exactly one of them."""
    return [obj for obj in page.get_objects() if obj.type == 3]


def census(obj) -> list[tuple[tuple[int, int, int], int]]:
    """Every colour in one image object's own bitmap, most common first.

    The image's *own* pixels are read rather than the rendered page's, so nothing here is
    a resampling of PowerPoint's raster: these are the bytes it wrote.  pdfium hands them
    over as BGR, which is turned back the right way round here and nowhere else.
    """
    array = obj.get_bitmap().to_numpy()
    flat = array.reshape(-1, array.shape[-1])
    counts = collections.Counter(map(bytes, flat))
    return [((c[2], c[1], c[0]), n) for c, n in counts.most_common()]


def faces(colours: list[tuple[tuple[int, int, int], int]], fill: str) -> list[tuple]:
    """Which drawn colours are *this* fill lit -- and how.

    Two models, because the measurement found two.  A **shaded** face is painted
    ``factor * fill`` per channel, so it lies on the ray through the fill and the factor is
    that ray's least-squares projection.  The **front** face is not on the ray at all: it
    is the fill plus the same small number on every channel, clamped at 255, which is an
    additive white term and not a scaling of anything.  Each colour is fitted both ways and
    kept if either lands within :data:`FACE_TOLERANCE`; the residuals are what say which
    model it is, and they are printed rather than resolved here.

    Fitting both is what keeps the front face from being lost: a lift of six levels reads
    as 1.04 of a dark channel and 1.00 of a bright one, which no single factor covers.
    """
    base = tuple(int(fill[i : i + 2], 16) for i in (0, 2, 4))
    norm = sum(v * v for v in base)
    if not norm:
        return []
    out = []
    for colour, count in colours:
        if count < MIN_FACE_PIXELS:
            continue
        factor = sum(a * b for a, b in zip(colour, base)) / norm
        scaled = max(abs(a - factor * b) for a, b in zip(colour, base))
        # The additive fit ignores a channel the clamp has pinned to 255, which is where
        # a saturated fill loses the evidence rather than contradicting it -- but two
        # channels have to survive, or white itself fits every fill with nothing left to
        # contradict it.
        free = [(a, b) for a, b in zip(colour, base) if a < 255]
        offset = sum(a - b for a, b in free) / len(free) if free else 0.0
        lifted = (
            max(abs(a - b - offset) for a, b in free)
            if len(free) >= 2
            else FACE_TOLERANCE + 1.0
        )
        if not 0.05 <= factor <= 1.6:
            continue
        # A lift is a front face, and a front face is the fill nudged rather than
        # repainted; without the bound, one series' *shaded* face reads as another
        # series' fill plus a large constant whenever two fills happen to differ by one.
        if scaled > FACE_TOLERANCE and abs(offset) > MAX_FRONT_LIFT:
            continue
        if min(scaled, lifted) <= FACE_TOLERANCE:
            out.append((factor, count, colour, scaled, offset, lifted))
    out.sort(key=lambda row: -row[0])
    return out


def ink_box(obj) -> tuple[float, float, float, float] | None:
    """The drawn scene's own bounding box inside its raster, in page points.

    The image object covers the plot with room to spare; what the scene *occupies* is the
    non-background ink in it, and for a ``pie3DChart`` -- which draws no axis and so
    leaves no vector text to measure -- that box is the only reading there is.
    """
    array = obj.get_bitmap().to_numpy()
    height, width = array.shape[:2]
    # The export leaves the transparent surround pure black and the scene's own paper
    # white; anything else is ink.
    lo = array.min(axis=2).astype(int)
    hi = array.max(axis=2).astype(int)
    mask = (hi > 8) & (lo < 247)
    rows = mask.any(axis=1).nonzero()[0]
    cols = mask.any(axis=0).nonzero()[0]
    if not len(rows) or not len(cols):
        return None
    left, bottom, right, top = obj.get_bounds()
    sx = (right - left) / width
    sy = (top - bottom) / height
    return (
        left + cols[0] * sx,
        top - (rows[0]) * sy,
        left + (cols[-1] + 1) * sx,
        top - (rows[-1] + 1) * sy,
    )


def _runs(flags) -> list[tuple[int, int]]:
    """Contiguous ``True`` spans of a boolean row, as inclusive index pairs."""
    out: list[tuple[int, int]] = []
    start = None
    for index, value in enumerate(flags):
        if value and start is None:
            start = index
        elif not value and start is not None:
            out.append((start, index - 1))
            start = None
    if start is not None:
        out.append((start, len(flags) - 1))
    return out


def solids(obj, fill: str, factor: float, minimum: int = 200):
    """Every span of one shaded face of one fill, left to right, in page points.

    A face is found by its **exact** drawn colour -- ``factor`` times the fill per sRGB
    channel, which ``view3d-colour`` measured as a constant -- within the level the export
    rounds by, and split into spans by the columns it occupies.  One span is one solid's
    face: prisms stand apart across the width and a ribbon is one span per series.

    Each span comes back as ``(left, right, top, bottom, rise)`` where the first four are
    the span's own box and ``rise`` is how much higher the face's top edge is at its right
    end than at its left -- which for a side face *is* the depth vector's slope, read
    without reference to anything else on the slide.
    """
    import numpy as np

    array = obj.get_bitmap().to_numpy()
    height, width = array.shape[:2]
    left, bottom, right, top = obj.get_bounds()
    sx = (right - left) / width
    sy = (top - bottom) / height
    base = [int(fill[i : i + 2], 16) for i in (0, 2, 4)]
    want = np.array([round(c * factor) for c in base])
    rgb = array[:, :, [2, 1, 0]].astype(int)
    mask = np.abs(rgb - want).max(axis=2) <= 2
    out = []
    for first, last in _runs(mask.any(axis=0)):
        span = mask[:, first : last + 1]
        if span.sum() < minimum:
            continue
        rows = span.any(axis=1).nonzero()[0]
        tops = [span[:, i].nonzero()[0][0] for i in range(span.shape[1]) if span[:, i].any()]
        out.append(
            (
                left + first * sx,
                left + (last + 1) * sx,
                top - rows[0] * sy,
                top - (rows[-1] + 1) * sy,
                (tops[0] - tops[-1]) * sy,
            )
        )
    return out


def mesh(path: Path, probes: list[dict], only: str | None) -> None:
    """Per probe, where the **solid** stands inside the scene -- the mesh's own reading.

    Everything else here reads the scene's box.  This reads what is drawn inside it: each
    prism's, ribbon's or slab's front face against the front plane the value axis defines,
    and its own depth off its right face.  Two numbers come out of every solid and neither
    needs the scene's box:

    * ``near`` -- how far the solid's front face is drawn *above* where the value it plots
      would put it on the front plane.  A solid standing back in the depth is lifted by
      exactly the depth it stands back, so this is that distance in drawn points.
    * ``deep`` -- the width and the rise of its right face, which is the solid's own depth
      vector.

    ``near / deep`` is where in its row the solid stands and needs no camera at all;
    ``deep`` against the camera is the projection the scene is *drawn* with, which is a
    different number from the one the plot rectangle is *reserved* with.
    """
    doc = pdfium.PdfDocument(path)
    print(
        "# key\tkind\tseries\tgapD\tgapW\trotX\trotY\tdepth%\th%\tframe"
        "\tplot_h\tplot_w\tunit_pt\tzero_y\tfill\tvalue\tfront_l\tfront_r\tfront_t"
        "\tfront_b\tnear\tdeep_dx\tdeep_dy"
    )
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page = doc[index]
        images = scene_images(page)
        read = read_page(page.raw, probe)
        values, positions = read["values"], None
        text = labels(page.raw)
        ticks = []
        for label, _left, y, _height, _width, _x in text:
            try:
                ticks.append((_number(label), y))
            except ValueError:
                pass
        ticks.sort()
        if len(ticks) < 2 or not images:
            print(f"{probe['key']}\t{probe['kind']}\t(nothing to read)")
            continue
        span = ticks[-1][0] - ticks[0][0]
        unit = (ticks[-1][1] - ticks[0][1]) / span if span else 0.0
        zero = ticks[0][1] - ticks[0][0] * unit
        view = probe["view"] or {}
        head = [
            probe["key"],
            probe["kind"],
            str(probe.get("series", 1)),
            str(probe.get("gapDepth", 150)),
            str(probe.get("gapWidth", 150)),
            str(view.get("rotX", 0)),
            str(view.get("rotY", 0)),
            str(view.get("depthPercent", 100)),
            str(view.get("hPercent", "")),
            f"{probe['frame'][1] / EMU:.0f}",
            f"{read['height']:.3f}",
            f"{read['right'] - read['left']:.3f}",
            f"{unit:.5f}",
            f"{zero:.3f}",
        ]
        high = probe["high"]
        shape = probe.get("factors") or ()
        scales = probe.get("scales") or (1.0,)
        for order, fill in enumerate((probe.get("colours") or ())[: probe.get("series", 1)]):
            fronts = solids(images[0], fill, 1.0)
            sides = solids(images[0], fill, 0.6364)
            scale = scales[order] if order < len(scales) else 1.0
            for place, front in enumerate(fronts):
                value = high * shape[place] * scale if place < len(shape) else float("nan")
                side = next(
                    (row for row in sides if row[0] >= front[1] - 0.5), None
                )
                print(
                    "\t".join(
                        head
                        + [
                            fill,
                            f"{value:.4f}",
                            f"{front[0]:.3f}",
                            f"{front[1]:.3f}",
                            f"{front[2]:.3f}",
                            f"{front[3]:.3f}",
                            f"{front[3] - zero:.3f}",
                            f"{side[1] - side[0]:.3f}" if side else "-",
                            f"{side[4]:.3f}" if side else "-",
                        ]
                    )
                )


def centroid(array, colour: tuple[int, int, int]) -> tuple[float, float]:
    """Where a colour's pixels sit in the raster, as a fraction of it: (down, across).

    Which face a factor belongs to is not something the factor itself can say -- if the
    shading moves with the camera the two could even cross -- so the faces are told apart
    by where they are drawn.  A top face sits above the front face it caps and a side face
    beside it, and that ordering holds at every camera this sweeps.
    """
    mask = (
        (array[:, :, 2] == colour[0])
        & (array[:, :, 1] == colour[1])
        & (array[:, :, 0] == colour[2])
    )
    rows, cols = mask.nonzero()
    height, width = array.shape[:2]
    return (float(rows.mean()) / height, float(cols.mean()) / width)


def colours(path: Path, probes: list[dict], only: str | None) -> None:
    """Per probe and per series fill, every face factor the raster shows."""
    doc = pdfium.PdfDocument(path)
    print(
        "# key\trotX\trotY\tfill\tfactor\tdrawn\tpixels\tscaled_by\toffset"
        "\tlifted_by\tdown\tacross"
    )
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        view = probe["view"] or {}
        images = scene_images(doc[index])
        if not images:
            print(f"{probe['key']}\t(no raster)")
            continue
        array = images[0].get_bitmap().to_numpy()
        found = census(images[0])
        for fill in probe.get("colours") or ():
            for factor, count, colour, scaled, offset, lifted in faces(found, fill):
                down, across = centroid(array, colour)
                print(
                    f"{probe['key']}\t{view.get('rotX', 0)}\t{view.get('rotY', 0)}\t"
                    f"{fill}\t{factor:.4f}\t#{colour[0]:02X}{colour[1]:02X}{colour[2]:02X}\t"
                    f"{count}\t{scaled:.2f}\t{offset:+.2f}\t{lifted:.2f}"
                    f"\t{down:.4f}\t{across:.4f}"
                )


def shapes(path: Path, probes: list[dict], only: str | None) -> None:
    """Per probe, the scene raster's box and the ink inside it, in page points.

    ROADMAP.md 3.4 checked the camera against this box on the gallery: the face plus its
    depth vector landed within two points of the image object's own corners.  That makes
    the box a reading of the whole scene for every group element, including the two whose
    axis is not measured here and the one that has no axis at all.
    """
    doc = pdfium.PdfDocument(path)
    print(
        "# key\tkind\tframe\tseries\tview\timg_l\timg_t\timg_w\timg_h"
        "\tink_l\tink_t\tink_w\tink_h"
    )
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page = doc[index]
        images = scene_images(page)
        view = probe["view"] or {}
        shown = ",".join(f"{k}={v}" for k, v in view.items())
        if not images:
            print(f"{probe['key']}\t{probe['kind']}\t-\t-\t{shown}\t(no raster)")
            continue
        left, bottom, right, top = images[0].get_bounds()
        ink = ink_box(images[0])
        row = [
            probe["key"],
            probe["kind"],
            f"{probe['frame'][1] / EMU:.0f}",
            str(probe.get("series", 1)),
            shown,
            f"{left:.3f}",
            f"{top:.3f}",
            f"{right - left:.3f}",
            f"{top - bottom:.3f}",
        ]
        if ink:
            row += [
                f"{ink[0]:.3f}",
                f"{ink[1]:.3f}",
                f"{ink[2] - ink[0]:.3f}",
                f"{ink[1] - ink[3]:.3f}",
            ]
        print("\t".join(row))


def main() -> int:
    path = Path(sys.argv[1]).expanduser()
    rest = sys.argv[2:]
    check = "--check" in rest
    only = next((arg for arg in rest if not arg.startswith("--")), None)
    probes = probes_for(path)
    if "--solve" in rest:
        solve(path, probes)
        return 0
    if "--colours" in rest:
        colours(path, probes, only)
        return 0
    if "--shapes" in rest:
        shapes(path, probes, only)
        return 0
    if "--mesh" in rest:
        mesh(path, probes, only)
        return 0
    doc = pdfium.PdfDocument(path)
    mine = ours(path.with_suffix(".pptx"), probes) if check else None

    print(
        "# key\tkind\tframe\tview\tdrawn\tplot_h\ttop\tbottom\tleft\tright"
        + ("\tours\tour_h\tour_top\tour_bottom\td_top\td_bottom\td_left" if check else "")
    )
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page = doc[index]
        read = read_page(page.raw, probe)
        box = insets(read, probe, page.get_height())
        view = probe["view"]
        shown = (
            "absent"
            if view is None
            else ",".join(f"{k}={v}" for k, v in view.items()) or "empty"
        )
        row = [
            probe["key"],
            probe["kind"],
            f"{probe['frame'][1] / EMU:.0f}",
            shown,
            _axis(read),
            f"{box['height']:.3f}",
            f"{box['top']:.3f}",
            f"{box['bottom']:.3f}",
            f"{box['left']:.3f}",
            f"{box['right']:.3f}",
        ]
        if check:
            our = mine[index]
            our_box = insets(our, probe, page.get_height())
            row += [
                _axis(our),
                f"{our_box['height']:.3f}",
                f"{our_box['top']:.3f}",
                f"{our_box['bottom']:.3f}",
                f"{box['top'] - our_box['top']:+.3f}",
                f"{box['bottom'] - our_box['bottom']:+.3f}",
                f"{box['left'] - our_box['left']:+.3f}",
            ]
        print("\t".join(row))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
