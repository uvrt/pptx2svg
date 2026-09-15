#!/usr/bin/env python3
"""Read the drawn geometry back out of the group-shear probe's PDF export.

Rasterising and eyeballing cannot separate a parallelogram from a rotated rectangle to
the precision the question needs, so this walks PDF page objects instead and reports the
actual device-space vertices PowerPoint emitted.  For text it reports the per-character
text matrix, which carries the effective angle and -- crucially -- whether there is any
skew term in it at all.

For a geometry probe it prints the two factors the drawn rectangle says were applied to
the child's *own* axes -- matched by edge direction, so the reading does not depend on
which corner PowerPoint chose to start the path at -- together with the interior angle,
which is the shear test, and the centre against the one a plain scale predicts.

Usage::

    python3 tools/read_group_shear_probe.py ~/pptx2svg-oracle/group-shear.pdf [substring]

The probe table is chosen by the PDF's file name (``group-shear``, ``group-sweep``,
``group-tie``); the optional substring filters to the probes whose key contains it.
"""

from __future__ import annotations

import ctypes
import math
import os
import sys
from pathlib import Path

import pypdfium2 as pdfium
import pypdfium2.raw as raw

sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_group_shear_probe import BOX, BOX_OFF, DECKS, ORIGIN  # noqa: E402

EMU_PER_PT = 12700
SLIDE_H_PT = 405.0


def _matrix(obj):
    m = raw.FS_MATRIX()
    raw.FPDFPageObj_GetMatrix(obj, ctypes.byref(m))
    return (m.a, m.b, m.c, m.d, m.e, m.f)


def _mul(outer, inner):
    """``inner`` then ``outer``, in PDF's row-vector convention."""
    a1, b1, c1, d1, e1, f1 = inner
    a2, b2, c2, d2, e2, f2 = outer
    return (
        a1 * a2 + b1 * c2,
        a1 * b2 + b1 * d2,
        c1 * a2 + d1 * c2,
        c1 * b2 + d1 * d2,
        e1 * a2 + f1 * c2 + e2,
        e1 * b2 + f1 * d2 + f2,
    )


def _apply(m, x, y):
    a, b, c, d, e, f = m
    return (a * x + c * y + e, b * x + d * y + f)


def _stroke(obj, parent):
    """The pen width in points.

    ``FPDFPageObj_GetStrokeWidth`` answers in the object's own space, and PowerPoint
    draws these shapes in EMU -- it returns 76200 for a 6 pt line -- so the object's
    matrix has to be applied before the number means anything, the same trap as the
    ``/FontSize 1`` that makes ``FPDFText_GetFontSize`` useless on this output.
    """
    w = ctypes.c_float()
    raw.FPDFPageObj_GetStrokeWidth(obj, ctypes.byref(w))
    m = _mul(parent, _matrix(obj))
    return round(w.value * math.sqrt(abs(m[0] * m[3] - m[1] * m[2])), 4)


def _fill(obj):
    vals = [ctypes.c_uint() for _ in range(4)]
    raw.FPDFPageObj_GetFillColor(obj, *[ctypes.byref(v) for v in vals])
    return tuple(v.value for v in vals[:3])


def _path_points(obj, parent):
    m = _mul(parent, _matrix(obj))
    count = raw.FPDFPath_CountSegments(obj)
    points = []
    for i in range(count):
        seg = raw.FPDFPath_GetPathSegment(obj, i)
        x, y = ctypes.c_float(), ctypes.c_float()
        raw.FPDFPathSegment_GetPoint(seg, ctypes.byref(x), ctypes.byref(y))
        points.append(_apply(m, x.value, y.value))
    return points


def walk(page_or_form, parent=(1, 0, 0, 1, 0, 0), form=False):
    """Yield ``(object, accumulated matrix)`` for every object, descending into forms."""
    if form:
        count = raw.FPDFFormObj_CountObjects(page_or_form)
        get = lambda i: raw.FPDFFormObj_GetObject(page_or_form, i)  # noqa: E731
    else:
        count = raw.FPDFPage_CountObjects(page_or_form)
        get = lambda i: raw.FPDFPage_GetObject(page_or_form, i)  # noqa: E731
    for i in range(count):
        obj = get(i)
        kind = raw.FPDFPageObj_GetType(obj)
        if kind == raw.FPDF_PAGEOBJ_FORM:
            yield from walk(obj, _mul(parent, _matrix(obj)), form=True)
        else:
            yield obj, kind, parent


def polygon(points, tol=0.05):
    """Drop the duplicate closing point and any repeats, keeping order."""
    out = []
    for p in points:
        if not out or abs(p[0] - out[-1][0]) > tol or abs(p[1] - out[-1][1]) > tol:
            out.append(p)
    if len(out) > 1 and abs(out[0][0] - out[-1][0]) <= tol and abs(out[0][1] - out[-1][1]) <= tol:
        out.pop()
    return out


def describe(points):
    """Side lengths, diagonals, interior angle and the edge directions of a quad."""
    if len(points) != 4:
        return {"n": len(points), "points": [(round(x, 3), round(y, 3)) for x, y in points]}
    (x0, y0), (x1, y1), (x2, y2), (x3, y3) = points
    side = lambda a, b: math.hypot(b[0] - a[0], b[1] - a[1])  # noqa: E731
    e1 = (x1 - x0, y1 - y0)
    e2 = (x3 - x0, y3 - y0)
    dot = e1[0] * e2[0] + e1[1] * e2[1]
    n1 = math.hypot(*e1)
    n2 = math.hypot(*e2)
    interior = math.degrees(math.acos(max(-1.0, min(1.0, dot / (n1 * n2))))) if n1 and n2 else 0
    return {
        "points": [(round(x, 3), round(y, 3)) for x, y in points],
        "sides": [round(side(points[i], points[(i + 1) % 4]), 3) for i in range(4)],
        "diagonals": [round(side(points[0], points[2]), 3), round(side(points[1], points[3]), 3)],
        "interior_deg": round(interior, 4),
        # PDF y runs up the page; negate so angles read clockwise-from-east like OOXML.
        "edge1_deg": round(math.degrees(math.atan2(-e1[1], e1[0])), 4),
        "edge2_deg": round(math.degrees(math.atan2(-e2[1], e2[0])), 4),
        "len1": round(n1, 3),
        "len2": round(n2, 3),
    }


def reduce_quad(points, probe):
    """What the drawn quad says about the factors applied to the shape's own axes.

    Model-free: the two edge directions of a drawn *rectangle* always come out as the
    pair ``{theta, theta + 90}`` modulo 180, so each edge can be matched to the shape's
    own x or y axis by its direction alone, without trusting the order PowerPoint chose
    to emit the vertices in.  The edge matched to the local x axis, divided by the
    authored ``cx``, is the factor that axis was actually scaled by.
    """
    if len(points) != 4:
        return {"n": len(points)}
    theta = probe["rot"]
    cx_pt = probe.get("cx", BOX) / EMU_PER_PT
    cy_pt = probe.get("cy", BOX) / EMU_PER_PT
    edges = [(points[(i + 1) % 4][0] - points[i][0], points[(i + 1) % 4][1] - points[i][1])
             for i in range(4)]
    out = {}
    for dx, dy in edges:
        length = math.hypot(dx, dy)
        # PDF y runs up the page; negate so the angle reads clockwise-from-east as OOXML
        # measures rotation.  Modulo 180 because an edge and its reverse are one axis.
        angle = math.degrees(math.atan2(-dy, dx)) % 180
        for axis, expected, extent in (("x", theta % 180, cx_pt),
                                       ("y", (theta + 90) % 180, cy_pt)):
            if min(abs(angle - expected), 180 - abs(angle - expected)) < 0.05:
                out[axis] = round(length / extent, 4)
    (x0, y0), (x1, y1), (x2, y2), (x3, y3) = points
    dot = (x1 - x0) * (x3 - x0) + (y1 - y0) * (y3 - y0)
    n = math.hypot(x1 - x0, y1 - y0) * math.hypot(x3 - x0, y3 - y0)
    out["interior"] = round(math.degrees(math.acos(max(-1.0, min(1.0, dot / n)))), 3) if n else 0
    out["centre"] = (round(sum(p[0] for p in points) / 4, 2),
                     round(SLIDE_H_PT - sum(p[1] for p in points) / 4, 2))
    sx, sy = probe["scale"]
    out["centre_pred"] = (
        round((ORIGIN[0] + sx * (BOX_OFF + probe.get("cx", BOX) / 2)) / EMU_PER_PT, 2),
        round((ORIGIN[1] + sy * (BOX_OFF + probe.get("cy", BOX) / 2)) / EMU_PER_PT, 2),
    )
    return out


def text_readout(page):
    textpage = raw.FPDFText_LoadPage(page)
    n = raw.FPDFText_CountChars(textpage)
    rows = []
    for i in range(n):
        ch = chr(raw.FPDFText_GetUnicode(textpage, i))
        m = raw.FS_MATRIX()
        raw.FPDFText_GetMatrix(textpage, i, ctypes.byref(m))
        box = [ctypes.c_double() for _ in range(4)]
        raw.FPDFText_GetCharBox(textpage, i, *[ctypes.byref(v) for v in box])
        left, right, bottom, top = (v.value for v in box)
        rows.append({
            "i": i,
            "ch": ch,
            "matrix": (round(m.a, 5), round(m.b, 5), round(m.c, 5), round(m.d, 5),
                       round(m.e, 3), round(m.f, 3)),
            "box": (round(left, 3), round(right, 3), round(bottom, 3), round(top, 3)),
        })
    raw.FPDFText_ClosePage(textpage)
    return rows


def main() -> int:
    path = Path(sys.argv[1] if len(sys.argv) > 1
                else os.path.expanduser("~/pptx2svg-oracle/group-shear.pdf"))
    doc = pdfium.PdfDocument(path)
    probes = next((d for n, d in DECKS.items() if n in path.name), DECKS["probe"])
    only = sys.argv[2] if len(sys.argv) > 2 else None
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page = doc[index]
        raw_page = page.raw
        print(f"=== slide {index + 1}: {probe['key']}  scale={probe['scale']} "
              f"rot={probe['rot']} {'flipH' if probe.get('flip_h') else ''}"
              f"{'flipV' if probe.get('flip_v') else ''}"
              f"{' grot=' + str(probe['group_rot']) if probe.get('group_rot') else ''}"
              f"{' nest=' + probe['nest'] if probe.get('nest') else ''}")
        if probe["key"].startswith("text-"):
            for row in text_readout(raw_page):
                print("   ", row)
        else:
            for obj, kind, parent in walk(raw_page):
                if kind != raw.FPDF_PAGEOBJ_PATH:
                    continue
                pts = polygon(_path_points(obj, parent))
                if len(pts) < 3 or _fill(obj) == (255, 255, 255):
                    continue  # the slide backdrop, not a probe
                print(f"    fill={_fill(obj)} pen={_stroke(obj, parent)} "
                      f"{reduce_quad(pts, probe)}")
                print(f"        {describe(pts)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
