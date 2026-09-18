#!/usr/bin/env python3
"""Build decks of **3-D** charts, to measure what ``c:view3D`` does to the axis.

ROADMAP.md 3.4 records that PowerPoint rasterises 3-D chart geometry but leaves the text
vector, so the axis, its tick labels and therefore the plot rectangle are exactly
measurable even though the scene itself is a picture.  These decks are that measurement.

The defect these exist for is gallery slide 16, whose ``area3DChart`` PowerPoint draws
**0..50 by 5** where the flat fallback draws 0..60 by 10.  The stated cause was a depth
reservation shrinking the plot, which would lower the interval count through
:func:`~pptx2svg.resolve.chart.side_axis_intervals` -- so each probe reads back two
things at once:

* the **axis** PowerPoint chose, from the tick labels' values, and
* the **plot rectangle**, from the extreme tick labels' own centres, which is where the
  depth reservation would show.

Both decks draw one chart per slide, no title and no legend, so the frame the axis
divides is the whole frame and nothing else has to be subtracted from it.

``view3d-meter``
    The N-meter (see ``tools/make_axis_probe.py``) on a ``bar3DChart`` over seven frame
    heights, plus ``F`` (0..50) and ``G`` (0..96), which are the two datasets that
    separate "the range is padded 5% at each end" from "it is not".  A ``col`` control
    and a ``line3DChart``/``area3DChart``/``pie3DChart`` row say whether the answer is a
    property of 3-D or of one group element.
``view3d-view``
    One dataset, one frame, and ``c:view3D`` swept one element at a time --
    ``depthPercent``, ``hPercent``, ``rotX``, ``rotY``, ``rAngAx``, ``perspective``, and
    the element absent altogether.  The axis is held at the interval cap so that anything
    that moves is geometry rather than the count; a short-frame block at the end puts the
    count back in play at the extremes of the depth.
``view3d-colour``
    The **raster** rather than the text, which is the other half of what a 3-D slide holds.
    A four-series ``bar3DChart`` with its fills stated, over twelve pitches, fifteen yaws
    and eight diagonals, so that ``tools/read_view3d_probe.py --colours`` can read each
    prism's three faces off each slide's own image object.  What it settles: the face
    factors are **constants** and do not move with the camera at all.
``view3d-shape`` / ``view3d-aspect``
    The same instrument turned on the scene's *box*, which the image object's own bounds
    give directly -- so a ``pie3DChart``, which draws no axis, is measurable too.
    ``view3d-shape`` sweeps eleven cameras per group element and two more frames;
    ``view3d-aspect`` gives every series count from one to four four cameras each, two of
    them width-bound and two height-bound, which is what separates the scene's aspect from
    its depth.  See ``three_d_scene_shape``.
``view3d-cat`` / ``view3d-count`` / ``view3d-band``
    The **category count against the series count**, which every deck above holds at five
    -- and holding it at five is what hid ``1 / categories`` inside a fitted constant and
    made the scene's aspect look like a ladder in the series count.  ``view3d-cat`` moves
    the category count at one to three series, ``view3d-count`` takes the series count to
    six at two category counts, and ``view3d-band`` settles what those two leave open:
    three more category counts, an ``area3DChart`` spelled ``crossBetween="between"``, the
    stacked shape at two more counts, and ``depthPercent``/``hPercent``/``gapDepth`` away
    from their defaults.  Read with the plain (text) reader rather than ``--shapes``: the
    value axis' extreme tick labels *are* the drawn face's edges.  What they settle:
    ``aspect = floor((across + series) / 2) / categories`` and ``depth = series /
    categories``.  See ``three_d_scene_shape``.
``view3d-mesh``
    The **solid inside** the scene rather than the scene's box: each prism's own three
    faces, found in the raster by their exact drawn colours, against the front plane its
    value axis defines.  That gives where in its row of depth a bar stands and how deep it
    is, which is what a renderer drawing the scene needs and what no reading of the box
    can say.  What it settles: a 3-D bar is as deep as it is **wide**, so the depth's
    unidentified constant is the category count and its ``1.5`` is ``c:gapWidth``'s
    default.  Read back with ``--mesh``.

Usage -- the deck's file name picks its probe table::

    python3 tools/make_view3d_probe.py ~/pptx2svg-oracle/view3d-meter.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/view3d-meter.pptx ~/pptx2svg-oracle/view3d-meter.pdf
    python3 tools/read_view3d_probe.py ~/pptx2svg-oracle/view3d-meter.pdf

The decks and their exports are throwaway and are **not** committed; ``~/pptx2svg-oracle``
is the only directory PowerPoint is allowed to write to, and it is left holding its
eighteen corpus files.  See ROADMAP.md section 0.1 before blaming a failed export.
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from make_axis_probe import A, C, R, write_deck  # noqa: E402

#: The five N-meter datasets, whose drawn units together name the interval count, plus
#: the two that separate the padded rule from the bare one.  Under the shipped 2-D rule
#: ``F`` draws 0..60 by 10 and ``G`` 0..120 by 20 at ten intervals; with no padding at all
#: they draw 0..50 by 5 and 0..100 by 10, which is what gallery slide 16 shows.
METER = (("A", 5.0), ("B", 9.0), ("C", 3.8), ("D", 6.6), ("E", 8.5))
PADDING = (("F", 50.0), ("G", 96.0))

#: ``c:view3D`` as PowerPoint writes it for a chart inserted from the 3-D gallery.
DEFAULT_VIEW = {"rotX": 15, "rotY": 20, "depthPercent": 100, "rAngAx": 1}

FRAME_WIDTH = 8686800  # 684 pt, the width every axis deck has used
HEIGHTS = (60, 90, 120, 150, 195, 250, 330)

METER_PROBES: list[dict] = [
    *[
        {
            "key": f"m{height}-{tag}",
            "high": high,
            "frame": (FRAME_WIDTH, height * 12700),
            "kind": "bar3D",
            "view": DEFAULT_VIEW,
        }
        for height in HEIGHTS
        for tag, high in METER + PADDING
    ],
    # The same two padding probes drawn **flat**, on the same deck and the same frames:
    # the control that says the export machine still draws the 2-D rule this is compared
    # against, rather than that the rule has moved under everyone's feet.
    *[
        {
            "key": f"flat{height}-{tag}",
            "high": high,
            "frame": (FRAME_WIDTH, height * 12700),
            "kind": "col",
            "view": None,
        }
        for height in (120, 195)
        for tag, high in PADDING
    ],
    # Is the answer a property of "3-D" or of ``bar3DChart``?  Same data, same frame.
    *[
        {
            "key": f"type-{kind}-{tag}",
            "high": high,
            "frame": (FRAME_WIDTH, 195 * 12700),
            "kind": kind,
            "view": DEFAULT_VIEW,
        }
        for kind in ("line3D", "area3D", "area3Dstack")
        for tag, high in PADDING
    ],
]

#: The view sweep.  One dataset (``B``, 0..9) on a frame tall enough that the interval
#: count is at its cap, so the drawn axis is the same on every slide and every difference
#: read back is geometry.
VIEW_FRAME = (FRAME_WIDTH, 195 * 12700)


def _view(**overrides) -> dict:
    view = dict(DEFAULT_VIEW)
    view.update(overrides)
    return {key: value for key, value in view.items() if value is not None}


VIEW_CASES: list[tuple[str, dict | None]] = [
    ("absent", None),
    ("default", DEFAULT_VIEW),
    ("empty", {}),
    *[(f"d{value}", _view(depthPercent=value)) for value in (20, 50, 200, 500, 1000, 2000)],
    *[(f"h{value}", _view(hPercent=value)) for value in (20, 50, 100, 200, 500)],
    *[(f"rx{value}", _view(rotX=value)) for value in (0, 5, 30, 45, 60, 90, -15, -45)],
    *[(f"ry{value}", _view(rotY=value)) for value in (0, 45, 90, 135, 180, 270, 340)],
    ("ra0", _view(rAngAx=0)),
    ("ra0-rx30", _view(rAngAx=0, rotX=30)),
    *[(f"ra0-p{value}", _view(rAngAx=0, perspective=value)) for value in (0, 30, 60, 120)],
    # `rAngAx=1` with a perspective stated: the schema says the perspective is ignored
    # when the axes are at right angles, and this is the probe that says whether it is.
    ("ra1-p120", _view(perspective=120)),
]

VIEW_PROBES: list[dict] = [
    *[
        {
            "key": f"v-{name}",
            "high": 9.0,
            "frame": VIEW_FRAME,
            "kind": "bar3D",
            "view": view,
        }
        for name, view in VIEW_CASES
    ],
    # The count back in play: a 120 pt frame takes six intervals flat, so if the depth
    # reserve comes off the band the count is read from, these differ from each other.
    *[
        {
            "key": f"n{name}-{tag}",
            "high": high,
            "frame": (FRAME_WIDTH, 120 * 12700),
            "kind": "bar3D",
            "view": view,
        }
        for name, view in (
            ("d20", _view(depthPercent=20)),
            ("d100", DEFAULT_VIEW),
            ("d2000", _view(depthPercent=2000)),
            ("rx0", _view(rotX=0)),
            ("rx60", _view(rotX=60)),
            ("h500", _view(hPercent=500)),
        )
        for tag, high in (("A", 5.0), ("B", 9.0), ("F", 50.0))
    ],
]

#: The group element as a variable.  ``view3d-view`` swept ``c:view3D`` on a
#: ``bar3DChart`` alone, and the three readings beside it on ``view3d-meter`` say the
#: other 3-D group elements do *not* share its scene: on one frame and one view a
#: ``bar3DChart`` drew a 127.68 pt axis where ``line3DChart`` drew 88.56 and
#: ``area3DChart`` 59.04.  This repeats the sweep per element, and adds the series count,
#: which is the depth's own row divisor and the obvious candidate for the difference.
TYPE_CASES: list[tuple[str, dict | None, dict]] = [
    ("base", DEFAULT_VIEW, {}),
    ("d20", _view(depthPercent=20), {}),
    ("d500", _view(depthPercent=500), {}),
    ("rx0", _view(rotX=0), {}),
    ("rx60", _view(rotX=60), {}),
    ("ry0", _view(rotY=0), {}),
    ("ry90", _view(rotY=90), {}),
    ("h50", _view(hPercent=50), {}),
    ("h200", _view(hPercent=200), {}),
    ("s2", DEFAULT_VIEW, {"series": 2}),
    ("s3", DEFAULT_VIEW, {"series": 3}),
    ("f120", DEFAULT_VIEW, {"frame": (FRAME_WIDTH, 120 * 12700)}),
    ("f330", DEFAULT_VIEW, {"frame": (FRAME_WIDTH, 330 * 12700)}),
]

TYPE_PROBES: list[dict] = [
    {
        "key": f"t-{kind}-{name}",
        "high": 9.0,
        "frame": VIEW_FRAME,
        "kind": kind,
        "view": view,
        **extra,
    }
    for kind in ("bar3D", "line3D", "area3D")
    for name, view, extra in TYPE_CASES
]

#: The **series count** as a variable, which ``view3d-type`` found and could not settle.
#: A ``bar3DChart``'s depth reservation *shrinks* as series are added -- one, two and
#: three series reserved 0.0600, 0.0459 and 0.0373 of the scene's width on one frame --
#: and `2.5 / (1.5 + n)` and `1 / sqrt(n)` both reproduce those three to 3%.  They part
#: company at six (0.333 against 0.408), which is what this deck is for.  ``gapDepth`` is
#: swept beside it because the first of those two laws is written in terms of it.
SERIES_PROBES: list[dict] = [
    *[
        {
            "key": f"n{count}-f{height}",
            "high": 9.0,
            "frame": (FRAME_WIDTH, height * 12700),
            "kind": "bar3D",
            "view": DEFAULT_VIEW,
            "series": count,
        }
        for height in (195, 120)
        for count in (1, 2, 3, 4, 5, 6)
    ],
    *[
        {
            "key": f"g{gap}-s{count}",
            "high": 9.0,
            "frame": VIEW_FRAME,
            "kind": "bar3D",
            "view": DEFAULT_VIEW,
            "series": count,
            "gapDepth": gap,
        }
        for count in (1, 2, 4)
        for gap in (0, 50, 300, 500)
    ],
    *[
        {
            "key": f"{kind}-n{count}",
            "high": 9.0,
            "frame": VIEW_FRAME,
            "kind": kind,
            "view": DEFAULT_VIEW,
            "series": count,
        }
        for kind in ("line3D", "area3D")
        for count in (1, 2, 3, 4)
    ],
]


#: The colour deck's fills, stated rather than themed.  Four series whose channels are
#: spread as widely as possible, because what separates "the face factor multiplies the
#: sRGB triple" from "it scales the light and converts back" is an offset that only shows
#: on a channel far from the others: scaling in linear light is ``k*c - 0.055*(1-k)`` in
#: sRGB, an affine map whose offset is invisible on a bright channel and a tenth of a dark
#: one.  ``FF3300`` carries a zero channel and a dark one, ``103070`` is dark throughout.
PROBE_COLOURS = ("4472C4", "ED7D31", "FF3300", "103070")

#: The pie deck's slice fills -- the same four plus a mid grey, which is the control that
#: says whether a face factor reads the fill's hue at all.
PIE_COLOURS = ("4472C4", "ED7D31", "FF3300", "103070", "808080")

#: Three categories and four series on a 684 x 330 pt frame: twelve prisms about 41 pt
#: wide, which is enough face for a colour census to find the top and the side even at the
#: shallow pitches where they are a few pixels tall.
COLOUR_FRAME = (FRAME_WIDTH, 330 * 12700)
COLOUR_CATEGORIES = ("C1", "C2", "C3")
#: Descending, and that is not decoration: the series of a clustered ``bar3DChart`` stand
#: side by side and touching, so a bar taller than its right-hand neighbour is the only
#: one whose own right face is not behind it.  Stepping the heights down keeps all four
#: side faces in the picture.
COLOUR_SCALES = (1.0, 0.85, 0.7, 0.55)


def _colour_probe(name: str, view: dict | None, **extra) -> dict:
    probe = {
        "key": f"c-{name}",
        "high": 9.0,
        "frame": COLOUR_FRAME,
        "kind": "bar3D",
        "view": view,
        "series": 4,
        "colours": PROBE_COLOURS,
        "categories": COLOUR_CATEGORIES,
        "scales": COLOUR_SCALES,
    }
    probe.update(extra)
    return probe


#: **The colour sweep.**  ROADMAP.md 3.4 read a prism's three faces off one camera --
#: the front face the series colour exactly, the top ``0.758x`` it and the right
#: ``0.632x`` -- and nothing there says whether those are constants or a cosine evaluated
#: at fifteen degrees.  This sweeps the pitch and the yaw one at a time and reads the
#: faces off each slide's own raster, which is the measurement that tells the two apart.
COLOUR_PROBES: list[dict] = [
    *[_colour_probe(f"rx{value}", _view(rotX=value)) for value in
      (0, 5, 10, 15, 20, 25, 30, 40, 50, 60, 75, 90)],
    *[_colour_probe(f"ry{value}", _view(rotY=value)) for value in
      (0, 5, 10, 20, 30, 45, 60, 75, 90, 120, 150, 180, 225, 270, 315)],
    *[
        _colour_probe(f"x{x}y{y}", _view(rotX=x, rotY=y))
        for x, y in (
            (0, 0), (30, 30), (45, 45), (60, 45),
            (45, 60), (90, 90), (-15, 20), (-45, 20),
        )
    ],
    *[_colour_probe(f"d{value}", _view(depthPercent=value)) for value in (20, 500)],
    *[_colour_probe(f"h{value}", _view(hPercent=value)) for value in (50, 200)],
    # Does the series count move the factors?  One and two series, same cameras.
    _colour_probe("s1", DEFAULT_VIEW, series=1, colours=PROBE_COLOURS[:1], scales=(1.0,)),
    _colour_probe(
        "s2", DEFAULT_VIEW, series=2, colours=PROBE_COLOURS[:2], scales=(1.0, 0.7)
    ),
    _colour_probe(
        "s1-rx45", _view(rotX=45), series=1, colours=PROBE_COLOURS[:1], scales=(1.0,)
    ),
]

#: **Part C: the scene per group element.**  Only ``bar3DChart`` has a camera, because
#: only its scene's shape is measured; ``line3DChart`` and ``area3DChart`` read 0.6 and
#: 0.4 of the region's aspect at one series on one frame, which is two readings and not a
#: law.  Every probe here carries stated fills, so the drawn scene's own ink is findable
#: in the raster -- which is the only reading a ``pie3DChart`` has, having no axis.
SHAPE_VIEWS: list[tuple[str, dict]] = [
    ("base", DEFAULT_VIEW),
    ("rx0", _view(rotX=0)),
    ("rx30", _view(rotX=30)),
    ("rx60", _view(rotX=60)),
    ("ry0", _view(rotY=0)),
    ("ry45", _view(rotY=45)),
    ("ry90", _view(rotY=90)),
    ("d20", _view(depthPercent=20)),
    ("d500", _view(depthPercent=500)),
    ("h50", _view(hPercent=50)),
    ("h200", _view(hPercent=200)),
]

SHAPE_KINDS = ("bar3D", "line3D", "area3D", "pie3D")


def _shape_probe(kind: str, name: str, view: dict, **extra) -> dict:
    probe = {
        "key": f"s-{kind}-{name}",
        "high": 9.0,
        "frame": VIEW_FRAME,
        "kind": kind,
        "view": view,
        "series": 1,
        "colours": PIE_COLOURS if kind == "pie3D" else PROBE_COLOURS,
    }
    probe.update(extra)
    return probe


SHAPE_PROBES: list[dict] = [
    *[
        _shape_probe(kind, name, view)
        for kind in SHAPE_KINDS
        for name, view in SHAPE_VIEWS
    ],
    # The second frame, which is what separates the scene's aspect from its depth.
    *[
        _shape_probe(
            kind,
            f"f{height}",
            DEFAULT_VIEW,
            frame=(FRAME_WIDTH, height * 12700),
        )
        for kind in SHAPE_KINDS
        for height in (120, 330)
    ],
    # The series count, on both frames: a ``line3DChart``'s depth *grows* with it.
    *[
        _shape_probe(
            kind,
            f"n{count}-f{height}",
            DEFAULT_VIEW,
            series=count,
            frame=(FRAME_WIDTH, height * 12700),
        )
        for kind in ("line3D", "area3D")
        for count in (2, 3, 4)
        for height in (195, 120)
    ],
    # A pie folds ``rotY`` into the camera or into its own seam, and those look different
    # past a quarter turn.
    *[
        _shape_probe("pie3D", f"ry{value}", _view(rotY=value))
        for value in (135, 180, 270)
    ],
]


#: **The scene aspect against the series count.**  ``view3d-shape`` measured each group
#: element's scene over eleven cameras at one series and found the aspect a clean multiple
#: of the region's -- 1.0 for a ``bar3DChart``, 0.6 for a ``line3DChart``, 0.4 for an
#: ``area3DChart`` -- but it swept the series count at one camera only, where the scene is
#: bound by the region's height and a single reading cannot separate the aspect from the
#: depth.  This gives every series count four cameras, which is what makes each of them
#: solvable on its own: two of them pitch the scene into the width-bound branch.
ASPECT_CAMERAS: list[tuple[str, dict]] = [
    ("base", DEFAULT_VIEW),
    ("rx0", _view(rotX=0)),
    ("rx45", _view(rotX=45)),
    ("ry90", _view(rotY=90)),
]

ASPECT_PROBES: list[dict] = [
    *[
        {
            "key": f"a-{kind}-n{count}-{name}",
            "high": 9.0,
            "frame": VIEW_FRAME,
            "kind": kind,
            "view": view,
            "series": count,
            "colours": PROBE_COLOURS,
            }
        for kind in ("bar3D", "line3D", "area3D")
        for count in (1, 2, 3, 4)
        for name, view in ASPECT_CAMERAS
    ],
    # Gallery slide 16's own shape: stacked, two series.  A stacked group stands its series
    # on top of each other rather than beside them, and whether that changes the scene is
    # exactly the question the aspect law has to answer for that slide.
    *[
        {
            "key": f"a-area3Dstack-{name}",
            "high": 9.0,
            "frame": VIEW_FRAME,
            "kind": "area3Dstack",
            "view": view,
            "series": 2,
            "colours": PROBE_COLOURS,
        }
        for name, view in ASPECT_CAMERAS
    ],
]

#: **The mesh.**  Every deck above reads where the scene's *box* is; this one reads where
#: the **solid inside it** stands, which is what a renderer that draws the scene has to
#: know and what nothing here had measured.  One series with a stated fill, three
#: categories whose values descend so that no prism's right face is hidden behind its
#: neighbour, and a value axis whose ticks give the front plane's own ``value -> y`` on
#: every slide.  ``tools/read_view3d_probe.py --mesh`` reads each prism's three faces off
#: the raster by their exact colours and prints, in page points:
#:
#: * the **near offset** -- how far back from the front plane the solid begins, as the
#:   drawn displacement between the value the bar plots and where its front face is
#:   actually painted, and
#: * the **solid's own depth**, as the width and slope of its right face.
#:
#: Those two, against the camera, are the whole of what the mesh needs: the depth vector
#: the scene is drawn with (which is *not* the same number as the reservation's, and this
#: is the deck that says so) and where in a row of it a bar stands.  ``c:gapDepth`` is
#: swept because it is what should move the second, and ``c:gapWidth`` because the bar's
#: *width* rule in 3-D was assumed to be the flat one and never checked.
MESH_FRAME = (FRAME_WIDTH, 250 * 12700)
MESH_CATEGORIES = ("C1", "C2", "C3")
#: Descending: a prism's right face is hidden by the next prism along whenever that one
#: is taller and the gap between them is small, and the ``gapWidth=0`` probe has no gap.
MESH_FACTORS = (1.0, 0.72, 0.44)


def _mesh_probe(name: str, view: dict | None = None, **extra) -> dict:
    probe = {
        "key": f"m-{name}",
        "high": 9.0,
        "frame": MESH_FRAME,
        "kind": "bar3D",
        "view": view or DEFAULT_VIEW,
        "series": 1,
        "colours": PROBE_COLOURS,
        "categories": MESH_CATEGORIES,
        "factors": MESH_FACTORS,
        "scales": (1.0,),
    }
    probe.update(extra)
    return probe


MESH_PROBES: list[dict] = [
    # The camera, one element at a time.  ``rotX`` and ``rotY`` each project the depth on
    # their own axis, so a sweep of either separates the projection constant from the
    # sine it multiplies.
    *[_mesh_probe(f"rx{value}", _view(rotX=value)) for value in (5, 10, 15, 20, 30, 45, 60)],
    *[_mesh_probe(f"ry{value}", _view(rotY=value)) for value in (5, 10, 20, 30, 45, 60, 90)],
    *[_mesh_probe(f"d{value}", _view(depthPercent=value)) for value in (20, 50, 200, 500)],
    *[_mesh_probe(f"h{value}", _view(hPercent=value)) for value in (50, 200)],
    # The row.  ``gapDepth`` is the whole of where a bar stands in its own depth: at 0 it
    # should fill the row and at 500 it should be a sixth of it.
    *[_mesh_probe(f"gd{value}", gapDepth=value) for value in (0, 50, 300, 500)],
    # The bar's width across the category, which is the flat rule or is not.
    *[_mesh_probe(f"gw{value}", gapWidth=value) for value in (0, 50, 300)],
    # Series stand side by side across the width and share one row of depth -- or they do
    # not, and this is what says which.
    # Descending, so that no series' right face is hidden behind the one beside it.
    *[
        _mesh_probe(f"s{count}", series=count, scales=(1.0, 0.78, 0.56, 0.34))
        for count in (2, 3, 4)
    ],
    *[
        _mesh_probe(f"f{height}", frame=(FRAME_WIDTH, height * 12700))
        for height in (150, 330)
    ],
    # **Stacked**, which the clustered sweep can only infer: series that stand on top of
    # each other take one slot of the band rather than one each, so the bar is as wide --
    # and therefore as deep -- as a single-series chart's.
    *[
        _mesh_probe(
            f"{name}{count}",
            series=count,
            grouping=grouping,
            scales=(0.5, 0.3, 0.2, 0.15),
        )
        for name, grouping in (("stack", "stacked"), ("pct", "percentStacked"))
        for count in (2, 3)
    ],
    # A horizontal 3-D bar, whose categories run *up* the scene: the same solids in the
    # same scene, and a different order to paint them in.
    _mesh_probe("bardir", barDir="bar"),
    # The other two group elements that draw a scene.  A ``line3DChart`` stands a ribbon
    # per series and an ``area3DChart`` a slab, both one row of depth per series, and
    # where in its row each of those sits is the same question the bar's ``gapDepth``
    # sweep asks.
    *[
        _mesh_probe(
            f"{kind}-n{count}", kind=kind, series=count, scales=(0.45, 0.72, 1.0)
        )
        for kind in ("line3D", "area3D")
        for count in (1, 2, 3)
    ],
    *[
        _mesh_probe(f"{kind}-{name}", view, kind=kind)
        for kind in ("line3D", "area3D")
        for name, view in (("rx45", _view(rotX=45)), ("ry60", _view(rotY=60)))
    ],
    *[
        _mesh_probe(f"{kind}-gd{value}", kind=kind, gapDepth=value)
        for kind in ("line3D", "area3D")
        for value in (0, 500)
    ],
]

#: **The scene aspect against the *category* count**, which every earlier deck held at
#: five.  ``three_d_scene_shape``'s ladder ``0.4 + 0.2 * ceil(m / 2)`` was fitted on
#: five-category decks -- which is exactly where the depth law's old form
#: (``VIEW_3D_DEPTH_PROJECTION`` standing in for ``1 / categories``) and its corrected one
#: agree -- and ``view3d-mesh``'s three-category ``area3DChart`` contradicts it.  So the
#: aspect has to be re-derived rather than extended, and that needs the two counts varied
#: **independently**: this deck moves the category count at a fixed series count and
#: ``view3d-count`` moves the series count past where the ladder was read.
#:
#: Four cameras per cell, two of them width-bound (``rx0``, ``ry90``) and two height-bound
#: (``base``, ``rx45``), which is what separates the scene's aspect from its depth without
#: needing an estimate of the region at all: the ink box's own height over its width is
#: ``(aspect + margin + depth*|sin rotX|) / (1 + margin + depth*|sin rotY|)``, in which the
#: region cancels.  A ``bar3DChart`` on the same frame and the same category count is the
#: control that says what the region's aspect is, measured the same way.
CAT_COUNTS = (2, 3, 5, 8)

#: A value shape per category count, so that no category is the maximum twice and a deck
#: with eight categories is not the five-category table run off its end.
CAT_SHAPES: dict[int, tuple[float, ...]] = {
    2: (0.45, 1.0),
    3: (0.45, 1.0, 0.7),
    4: (0.3, 1.0, 0.55, 0.8),
    5: (0.3, 1.0, 0.55, 0.8, 0.45),
    6: (0.3, 1.0, 0.55, 0.8, 0.45, 0.9),
    7: (0.3, 1.0, 0.55, 0.8, 0.45, 0.9, 0.6),
    8: (0.3, 1.0, 0.55, 0.8, 0.45, 0.9, 0.6, 0.35),
}

#: Descending, so a slab or a ribbon standing further back is never entirely hidden by the
#: one in front of it -- which matters for the *ink* box, since a scene whose back rows are
#: covered still draws its floor and its walls but its ink box is read off whatever is
#: painted.
CAT_SCALES = (1.0, 0.8, 0.62, 0.48, 0.36, 0.26)

CAT_FILLS = ("4472C4", "ED7D31", "FF3300", "103070", "70AD47", "7030A0")


def _cat_probe(kind: str, cats: int, count: int, name: str, view: dict, **extra) -> dict:
    probe = {
        "key": f"{kind}-c{cats}-n{count}-{name}",
        "high": 9.0,
        "frame": VIEW_FRAME,
        "kind": kind,
        "view": view,
        "series": count,
        "colours": CAT_FILLS,
        "categories": tuple(f"C{i + 1}" for i in range(cats)),
        "factors": CAT_SHAPES[cats],
        "scales": CAT_SCALES[:count],
        "cats": cats,
    }
    probe.update(extra)
    return probe


CAT_PROBES: list[dict] = [
    *[
        _cat_probe(kind, cats, count, name, view)
        for kind in ("line3D", "area3D")
        for cats in CAT_COUNTS
        for count in (1, 2, 3)
        for name, view in ASPECT_CAMERAS
    ],
    # The control: a ``bar3DChart``'s scene is the region's own, so the same reading on the
    # same frame says what the region is and what this instrument's own bias is.
    *[
        _cat_probe("bar3D", cats, 1, name, view)
        for cats in CAT_COUNTS
        for name, view in ASPECT_CAMERAS
    ],
]

#: **The series count past the ladder**, at two category counts, plus the stacked
#: ``area3DChart`` gallery slide 16 is.  Five and six series are the counts
#: ``VIEW_3D_SHAPE_MAX_SERIES`` refuses for want of a reading.
COUNT_PROBES: list[dict] = [
    *[
        _cat_probe(kind, cats, count, name, view)
        for kind in ("line3D", "area3D")
        for cats in (3, 5)
        for count in (4, 5, 6)
        for name, view in ASPECT_CAMERAS
    ],
    *[
        _cat_probe("area3Dstack", cats, count, name, view)
        for cats in (3, 5)
        for count in (2, 3)
        for name, view in ASPECT_CAMERAS
    ],
    *[
        _cat_probe("bar3D", 5, count, name, view)
        for count in (2, 4)
        for name, view in ASPECT_CAMERAS
    ],
]

#: **The band deck**, which settles what ``view3d-cat`` and ``view3d-count`` leave open
#: once their law is written down.  Those two say the scene's aspect is
#: ``floor((across + series) / 2) / categories`` of the region's, in which *across* is the
#: category count for a ``line3DChart`` and one less for an ``area3DChart`` -- and every
#: area probe on those decks draws ``c:crossBetween="midCat"``, which is the other thing
#: that separates the two group elements.  So four questions, at two cameras a cell
#: rather than four, the depth being settled (``series / categories``) and one
#: height-bound camera therefore enough to read the aspect:
#:
#: * an ``area3DChart`` with ``crossBetween="between"`` -- gallery slide 16's own spelling
#:   -- which says whether *across* is the axis' or the group element's;
#: * three more category counts, 4, 6 and 7, against a law whose floor makes the even and
#:   the odd counts behave differently;
#: * a stacked ``area3DChart`` at two more counts, whose aspect reads ``categories - 1``
#:   over the two it has been read at; and
#: * ``depthPercent`` and ``hPercent`` away from their defaults at two category counts,
#:   which is the multiplication those two are assumed to be.
BAND_PROBES: list[dict] = [
    *[
        _cat_probe(kind, cats, count, name, view)
        for kind in ("line3D", "area3D")
        for cats in (4, 6, 7)
        for count in (1, 2, 3)
        for name, view in ASPECT_CAMERAS[:1] + ASPECT_CAMERAS[2:3]
    ],
    *[
        _cat_probe("area3D", cats, count, f"btw-{name}", view, crossBetween="between")
        for cats in (3, 5, 8)
        for count in (1, 2)
        for name, view in ASPECT_CAMERAS[:1] + ASPECT_CAMERAS[2:3]
    ],
    *[
        _cat_probe("area3Dstack", cats, 2, name, view)
        for cats in (2, 8)
        for name, view in ASPECT_CAMERAS[:1] + ASPECT_CAMERAS[2:3]
    ],
    *[
        _cat_probe("area3Dstack", cats, 2, f"btw-{name}", view, crossBetween="between")
        for cats in (3, 5)
        for name, view in ASPECT_CAMERAS[:1] + ASPECT_CAMERAS[2:3]
    ],
    *[
        _cat_probe(kind, cats, 2, f"{tag}-{name}", dict(view, **override))
        for kind in ("line3D", "area3D")
        for cats in (3, 8)
        for tag, override in (
            ("d500", {"depthPercent": 500}),
            ("h50", {"hPercent": 50}),
        )
        for name, view in ASPECT_CAMERAS[:1] + ASPECT_CAMERAS[2:3]
    ],
    *[
        _cat_probe(kind, 5, 2, f"gd{gap}-{name}", view, gapDepth=gap)
        for kind in ("line3D", "area3D")
        for gap in (0, 500)
        for name, view in ASPECT_CAMERAS[:1] + ASPECT_CAMERAS[2:3]
    ],
]

DECKS = {
    "view3d-mesh": MESH_PROBES,
    "view3d-cat": CAT_PROBES,
    "view3d-count": COUNT_PROBES,
    "view3d-band": BAND_PROBES,
    "view3d-meter": METER_PROBES,
    "view3d-view": VIEW_PROBES,
    "view3d-type": TYPE_PROBES,
    "view3d-series": SERIES_PROBES,
    "view3d-colour": COLOUR_PROBES,
    "view3d-shape": SHAPE_PROBES,
    "view3d-aspect": ASPECT_PROBES,
}

CATEGORIES = ("C1", "C2", "C3", "C4", "C5")
#: The same shape every axis deck has used, so a value is never the maximum twice.
FACTORS = (0.3, 1.0, 0.55, 0.8, 0.45)

VIEW_ORDER = ("rotX", "hPercent", "rotY", "depthPercent", "rAngAx", "perspective")


def view_xml(view: dict | None) -> str:
    """``c:view3D``, in schema order.  ``None`` writes no element at all."""
    if view is None:
        return ""
    body = "".join(
        f"<c:{name} val='{view[name]}'/>" for name in VIEW_ORDER if name in view
    )
    return f"<c:view3D>{body}</c:view3D>"


def series_xml(
    high: float,
    index: int = 0,
    scale: float = 1.0,
    colour: str | None = None,
    categories: tuple[str, ...] = CATEGORIES,
    factors: tuple[float, ...] | None = None,
) -> str:
    """One ``c:ser``.  *colour* states its fill as an explicit ``srgbClr``.

    The colour deck needs the fill stated rather than inherited, because what it measures
    is the ratio between a face's drawn colour and the fill it came from, and a theme
    colour would leave the denominator to be looked up rather than known.
    """
    shape = factors or FACTORS
    values = [high * factor * scale for factor in shape[: len(categories)]]
    if factors is None:
        values[1] = high * scale
    points = "".join(f"<c:pt idx='{i}'><c:v>{v!r}</c:v></c:pt>" for i, v in enumerate(values))
    cats = "".join(
        f"<c:pt idx='{i}'><c:v>{name}</c:v></c:pt>" for i, name in enumerate(categories)
    )
    fill = (
        f"<c:spPr><a:solidFill><a:srgbClr val='{colour}'/></a:solidFill>"
        "<a:ln><a:noFill/></a:ln></c:spPr>"
        if colour
        else ""
    )
    return (
        f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        f"<c:pt idx='0'><c:v>S{index + 1}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        + fill
        + "<c:cat><c:strRef><c:strCache>"
        f"<c:ptCount val='{len(categories)}'/>{cats}</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef></c:val>"
        "</c:ser>"
    )


def axes_xml(kind: str, cross_between: str | None = None) -> str:
    """The category, value and (for a 3-D group) series axes this chart needs.

    The value axis carries gridlines and ten-point labels exactly as every earlier axis
    deck's does, so the reading is comparable to them.  The ``c:serAx`` is deleted --
    PowerPoint still draws the depth, it just prints nothing along it.
    """
    if kind == "pie3D":
        return ""
    depth = (
        "<c:serAx><c:axId val='100004'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='1'/><c:axPos val='b'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/>"
        "<c:crossAx val='100003'/><c:crosses val='autoZero'/></c:serAx>"
        if "3D" in kind
        else ""
    )
    cross_between = cross_between or ("midCat" if kind.startswith("area") else "between")
    return (
        "<c:catAx><c:axId val='100002'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='b'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100003'/>"
        "<c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/>"
        "<c:lblOffset val='100'/></c:catAx>"
        "<c:valAx><c:axId val='100003'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='l'/><c:majorGridlines/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100002'/>"
        f"<c:crosses val='autoZero'/><c:crossBetween val='{cross_between}'/></c:valAx>"
        + depth
    )


def group_xml(
    kind: str,
    high: float,
    series: int = 1,
    gap_depth: int = 150,
    colours: tuple[str, ...] | None = None,
    categories: tuple[str, ...] = CATEGORIES,
    scales: tuple[float, ...] | None = None,
    gap_width: int = 150,
    factors: tuple[float, ...] | None = None,
    bar_dir: str = "col",
    grouping: str = "clustered",
) -> str:
    """One ``c:*Chart`` group.  A 3-D group states three ``c:axId`` children, exactly.

    *series* is the number of ``c:ser`` children, which is what a 3-D chart lays out
    along its **depth** -- one row per series -- and therefore the lever that separates
    "the scene's depth is ``depthPercent`` of its width" from "it is that per row".
    """
    body = "".join(
        series_xml(
            high,
            index,
            (scales[index] if scales else (1.0 if series == 1 else 0.6 - 0.15 * index)),
            colours[index] if colours else None,
            categories,
            factors,
        )
        for index in range(series)
    )
    depth_gap = f"<c:gapDepth val='{gap_depth}'/>"
    ids3 = "<c:axId val='100002'/><c:axId val='100003'/><c:axId val='100004'/>"
    ids2 = "<c:axId val='100002'/><c:axId val='100003'/>"
    if kind == "bar3D":
        return (
            f"<c:bar3DChart><c:barDir val='{bar_dir}'/>"
            f"<c:grouping val='{grouping}'/>"
            "<c:varyColors val='0'/>" + body + f"<c:gapWidth val='{gap_width}'/>"
            + depth_gap + "<c:shape val='box'/>" + ids3 + "</c:bar3DChart>"
        )
    if kind == "col":
        return (
            "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/>"
            "<c:varyColors val='0'/>" + body + "<c:gapWidth val='150'/>"
            + ids2 + "</c:barChart>"
        )
    if kind == "line3D":
        return (
            "<c:line3DChart><c:grouping val='standard'/><c:varyColors val='0'/>"
            + body
            + depth_gap + ids3 + "</c:line3DChart>"
        )
    if kind == "area3D":
        return (
            "<c:area3DChart><c:grouping val='standard'/><c:varyColors val='0'/>"
            + body
            + depth_gap + ids3 + "</c:area3DChart>"
        )
    if kind == "area3Dstack":
        # Gallery slide 16 is stacked and two-series, which is the one shape the single
        # series above cannot rule out on its own, so **two** is the floor here: a probe
        # that states no count still gets the pair.  The fills and the gap are the
        # probe's, so a colour reading of this kind has a denominator.
        count = max(series, 2)
        shape = scales or (
            (0.6, 0.4) if count == 2 else tuple(0.6 - 0.12 * i for i in range(count))
        )
        stacked = "".join(
            series_xml(
                high,
                index,
                shape[index] if index < len(shape) else 0.2,
                colours[index] if colours else None,
                categories,
                factors,
            )
            for index in range(count)
        )
        return (
            "<c:area3DChart><c:grouping val='stacked'/><c:varyColors val='0'/>"
            + stacked
            + depth_gap + ids3 + "</c:area3DChart>"
        )
    if kind == "pie3D":
        # A pie has no axes and one series, so its slices carry the colours: ``c:dPt``
        # per point, stated rather than inherited for the same reason the bars' are.
        points = "".join(
            f"<c:dPt><c:idx val='{i}'/><c:bubble3D val='0'/>"
            f"<c:spPr><a:solidFill><a:srgbClr val='{(colours or PIE_COLOURS)[i]}'/>"
            "</a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:dPt>"
            for i in range(len(categories))
        )
        ser = series_xml(high, 0, 1.0, None, categories)
        ser = ser.replace("<c:cat>", points + "<c:cat>", 1)
        return "<c:pie3DChart><c:varyColors val='1'/>" + ser + "</c:pie3DChart>"
    raise SystemExit(f"unknown probe kind {kind!r}")


def chart_part(probe: dict) -> bytes:
    size = probe.get("size", 1000)
    kind = probe["kind"]
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:date1904 val='0'/><c:lang val='en-US'/><c:roundedCorners val='0'/>"
        "<c:chart>"
        + view_xml(probe["view"])
        + "<c:autoTitleDeleted val='1'/>"
        "<c:plotArea><c:layout/>"
        + group_xml(
            kind,
            probe["high"],
            probe.get("series", 1),
            probe.get("gapDepth", 150),
            probe.get("colours"),
            probe.get("categories", CATEGORIES),
            probe.get("scales"),
            probe.get("gapWidth", 150),
            probe.get("factors"),
            probe.get("barDir", "col"),
            probe.get("grouping", "clustered"),
        )
        + axes_xml(kind, probe.get("crossBetween"))
        + "</c:plotArea>"
        "<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr sz='{size}'/></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
        "</c:chartSpace>"
    ).encode()


def probes_for(target: Path) -> list[dict]:
    for name, table in DECKS.items():
        if name in target.stem:
            return table
    raise SystemExit(f"name the deck after one of {sorted(DECKS)}, not {target.stem!r}")


def main() -> int:
    target = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else Path("view3d-meter.pptx")
    probes = probes_for(target)
    write_deck(target, probes, part_for=chart_part)
    for index, probe in enumerate(probes):
        view = probe["view"]
        shown = "absent" if view is None else ",".join(f"{k}={v}" for k, v in view.items())
        print(
            f"{index + 1:3d} {probe['key']:16s} {probe['kind']:12s} 0..{probe['high']:<6g} "
            f"frame={probe['frame'][1] / 12700:.0f}pt  {shown}"
        )
    print(f"\n{len(probes)} probes -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
