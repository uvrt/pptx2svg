"""OOXML DrawingML preset shapes (ECMA-376 §20.1.10.56 ``prst``).

Each generator turns a shape's pixel width/height and its adjustment values into a
single SVG element -- ``<rect>``, ``<ellipse>``, ``<polygon>`` or ``<path>``.  The caller
splices fill and stroke attributes into whatever element comes back, which is why every
generator returns exactly one element (or one ``<g>``).

Adjustment values arrive as OOXML 1/1000-percent integers, so ``adj=50000`` means 50%.
Each generator supplies PowerPoint's default when the shape omits an adjustment.

These are approximations, not the full ECMA-376 guide-formula geometry: a ``cloud`` is a
rounded rectangle, not nine overlapping arcs.  Shapes whose real outline matters are
usually authored as custom geometry anyway, which :mod:`pptx2svg.parse.drawing` evaluates
exactly.  An unknown preset falls back to a rectangle.
"""

from __future__ import annotations

import math
from typing import Callable

from .. import model as m
from ..guides import arc_segments, evaluate_guides, resolve_value
from .preset_specs import PRESET_SPECS

Generator = Callable[[float, float, dict], str]


def _n(value: float) -> str:
    """Format a coordinate compactly: ``12.0`` -> ``12``, ``12.3456`` -> ``12.346``."""
    rounded = round(value, 3)
    if rounded == int(rounded):
        return str(int(rounded))
    return f"{rounded:g}"


def _pts(*points: tuple[float, float]) -> str:
    return " ".join(f"{_n(x)},{_n(y)}" for x, y in points)


def _adj(adj: dict, name: str, default: float) -> float:
    """Adjustment value as a 0..1 ratio (OOXML stores 1/1000 of a percent)."""
    return adj.get(name, default) / 100000



def _regular_polygon(w: float, h: float, sides: int) -> str:
    cx, cy = w / 2, h / 2
    points = [
        (
            cx + cx * math.cos(math.tau * i / sides - math.pi / 2),
            cy + cy * math.sin(math.tau * i / sides - math.pi / 2),
        )
        for i in range(sides)
    ]
    return f'<polygon points="{_pts(*points)}"/>'



# --------------------------------------------------------------------------------------
# Basic shapes
# --------------------------------------------------------------------------------------


def _rect(w, h, adj):
    return f'<rect width="{_n(w)}" height="{_n(h)}"/>'


def _ellipse(w, h, adj):
    return f'<ellipse cx="{_n(w/2)}" cy="{_n(h/2)}" rx="{_n(w/2)}" ry="{_n(h/2)}"/>'


def _round_rect(w, h, adj):
    ratio = min(0.5, max(0.0, _adj(adj, "adj", 16667)))
    r = ratio * min(w, h)
    return f'<rect width="{_n(w)}" height="{_n(h)}" rx="{_n(r)}" ry="{_n(r)}"/>'


def _triangle(w, h, adj):
    top_x = _adj(adj, "adj", 50000) * w
    return f'<polygon points="{_pts((top_x, 0), (w, h), (0, h))}"/>'


def _rt_triangle(w, h, adj):
    return f'<polygon points="{_pts((0, 0), (w, h), (0, h))}"/>'


def _diamond(w, h, adj):
    return f'<polygon points="{_pts((w/2, 0), (w, h/2), (w/2, h), (0, h/2))}"/>'




def _up_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 50000) * w
    head_l = _adj(adj, "adj2", 50000) * h
    left = (w - head_w) / 2
    right = w - left
    return f'<polygon points="{_pts((left, head_l), (0, head_l), (w/2, 0), (w, head_l), (right, head_l), (right, h), (left, h))}"/>'


def _down_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 50000) * w
    head_l = _adj(adj, "adj2", 50000) * h
    left = (w - head_w) / 2
    right = w - left
    shaft = h - head_l
    return f'<polygon points="{_pts((left, 0), (right, 0), (right, shaft), (w, shaft), (w/2, h), (0, shaft), (left, shaft))}"/>'


def _line(w, h, adj):
    return f'<line x1="0" y1="0" x2="{_n(w)}" y2="{_n(h)}"/>'


# --------------------------------------------------------------------------------------
# Connectors
# --------------------------------------------------------------------------------------


def _straight_connector1(w, h, adj):
    return f'<path d="M 0 0 L {_n(w)} {_n(h)}"/>'










def _cloud(w, h, adj):
    return f'<rect width="{_n(w)}" height="{_n(h)}" rx="{_n(min(w, h) * 0.15)}"/>'


def _heart(w, h, adj):
    cx = w / 2
    return (
        f'<path d="M {_n(cx)} {_n(h*0.35)} C {_n(cx)} {_n(h*0.1)}, 0 0, 0 {_n(h*0.35)} '
        f'C 0 {_n(h*0.65)}, {_n(cx)} {_n(h*0.85)}, {_n(cx)} {_n(h)} '
        f'C {_n(cx)} {_n(h*0.85)}, {_n(w)} {_n(h*0.65)}, {_n(w)} {_n(h*0.35)} '
        f'C {_n(w)} 0, {_n(cx)} {_n(h*0.1)}, {_n(cx)} {_n(h*0.35)} Z"/>'
    )


def _irregular_seal1(w, h, adj):
    ratios = [
        (0.15, 0.35), (0.27, 0.03), (0.38, 0.28), (0.5, 0.0), (0.6, 0.23), (0.73, 0.08),
        (0.72, 0.35), (1.0, 0.35), (0.78, 0.5), (0.95, 0.7), (0.73, 0.65), (0.65, 1.0),
        (0.5, 0.72), (0.35, 0.95), (0.32, 0.65), (0.05, 0.7), (0.18, 0.5), (0.0, 0.35),
    ]
    return f'<polygon points="{_pts(*[(w*x, h*y) for x, y in ratios])}"/>'


def _irregular_seal2(w, h, adj):
    ratios = [
        (0.1, 0.4), (0.18, 0.08), (0.32, 0.3), (0.45, 0.0), (0.55, 0.18), (0.72, 0.05),
        (0.68, 0.32), (1.0, 0.3), (0.82, 0.5), (0.98, 0.68), (0.75, 0.65), (0.8, 0.92),
        (0.55, 0.75), (0.42, 1.0), (0.38, 0.72), (0.12, 0.88), (0.22, 0.6), (0.0, 0.55),
    ]
    return f'<polygon points="{_pts(*[(w*x, h*y) for x, y in ratios])}"/>'


# --------------------------------------------------------------------------------------
# Additional arrows
# --------------------------------------------------------------------------------------



def _up_down_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 50000) * w
    head_l = _adj(adj, "adj2", 50000) * h
    left = (w - head_w) / 2
    right = w - left
    return f'<polygon points="{_pts((left, head_l), (0, head_l), (w/2, 0), (w, head_l), (right, head_l), (right, h - head_l), (w, h - head_l), (w/2, h), (0, h - head_l), (left, h - head_l))}"/>'




def _home_plate(w, h, adj):
    offset = _adj(adj, "adj", 50000) * min(w, h)
    return f'<polygon points="{_pts((0, 0), (w - offset, 0), (w, h/2), (w - offset, h), (0, h))}"/>'





# --------------------------------------------------------------------------------------
# Flowchart
# --------------------------------------------------------------------------------------


def _flow_alternate_process(w, h, adj):
    r = min(w, h) / 6
    return f'<rect width="{_n(w)}" height="{_n(h)}" rx="{_n(r)}" ry="{_n(r)}"/>'


def _flow_decision(w, h, adj):
    return f'<polygon points="{_pts((w/2, 0), (w, h/2), (w/2, h), (0, h/2))}"/>'


def _flow_input_output(w, h, adj):
    offset = w / 5
    return f'<polygon points="{_pts((offset, 0), (w, 0), (w - offset, h), (0, h))}"/>'


def _flow_predefined_process(w, h, adj):
    d = w / 8
    return (
        f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h)} L 0 {_n(h)} Z '
        f'M {_n(d)} 0 L {_n(d)} {_n(h)} M {_n(w-d)} 0 L {_n(w-d)} {_n(h)}"/>'
    )


def _flow_internal_storage(w, h, adj):
    dx, dy = w / 8, h / 8
    return (
        f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h)} L 0 {_n(h)} Z '
        f'M {_n(dx)} 0 L {_n(dx)} {_n(h)} M 0 {_n(dy)} L {_n(w)} {_n(dy)}"/>'
    )





def _flow_preparation(w, h, adj):
    offset = w / 5
    return f'<polygon points="{_pts((offset, 0), (w - offset, 0), (w, h/2), (w - offset, h), (offset, h), (0, h/2))}"/>'


def _flow_manual_input(w, h, adj):
    return f'<polygon points="{_pts((0, h/5), (w, 0), (w, h), (0, h))}"/>'


def _flow_manual_operation(w, h, adj):
    offset = w / 5
    return f'<polygon points="{_pts((0, 0), (w, 0), (w - offset, h), (offset, h))}"/>'


def _flow_offpage_connector(w, h, adj):
    arrow_h = h * 0.2
    return f'<polygon points="{_pts((0, 0), (w, 0), (w, h - arrow_h), (w/2, h), (0, h - arrow_h))}"/>'


def _flow_punched_card(w, h, adj):
    cut = min(w, h) * 0.2
    return f'<polygon points="{_pts((cut, 0), (w, 0), (w, h), (0, h), (0, cut))}"/>'



def _flow_collate(w, h, adj):
    return f'<polygon points="{_pts((0, 0), (w, 0), (w/2, h/2), (w, h), (0, h), (w/2, h/2))}"/>'


def _flow_sort(w, h, adj):
    return (
        f'<path d="M {_n(w/2)} 0 L {_n(w)} {_n(h/2)} L 0 {_n(h/2)} Z '
        f'M 0 {_n(h/2)} L {_n(w)} {_n(h/2)} '
        f'M {_n(w/2)} {_n(h)} L {_n(w)} {_n(h/2)} L 0 {_n(h/2)} Z"/>'
    )


def _flow_extract(w, h, adj):
    return f'<polygon points="{_pts((w/2, 0), (w, h), (0, h))}"/>'


def _flow_merge(w, h, adj):
    return f'<polygon points="{_pts((0, 0), (w, 0), (w/2, h))}"/>'



def _flow_delay(w, h, adj):
    arc_w = w * 0.35
    return (
        f'<path d="M 0 0 L {_n(w-arc_w)} 0 A {_n(arc_w)} {_n(h/2)} 0 0 1 {_n(w-arc_w)} {_n(h)} '
        f'L 0 {_n(h)} Z"/>'
    )


def _flow_display(w, h, adj):
    left_w = w * 0.15
    arc_w = w * 0.35
    return (
        f'<path d="M {_n(left_w)} 0 L {_n(w-arc_w)} 0 A {_n(arc_w)} {_n(h/2)} 0 0 1 {_n(w-arc_w)} {_n(h)} '
        f'L {_n(left_w)} {_n(h)} L 0 {_n(h/2)} Z"/>'
    )






# --------------------------------------------------------------------------------------
# Callouts
# --------------------------------------------------------------------------------------


def _wedge_rect_callout(w, h, adj):
    tip_x = w / 2 + _adj(adj, "adj1", -20833) * w
    tip_y = h / 2 + _adj(adj, "adj2", 62500) * h
    bx = w / 2
    wedge = w * 0.06
    return (
        f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h)} L {_n(bx+wedge)} {_n(h)} '
        f'L {_n(tip_x)} {_n(tip_y)} L {_n(bx-wedge)} {_n(h)} L 0 {_n(h)} Z"/>'
    )


def _wedge_round_rect_callout(w, h, adj):
    tip_x = w / 2 + _adj(adj, "adj1", -20833) * w
    tip_y = h / 2 + _adj(adj, "adj2", 62500) * h
    r = _adj(adj, "adj3", 16667) * min(w, h)
    bx = w / 2
    wedge = w * 0.06
    return (
        f'<path d="M {_n(r)} 0 L {_n(w-r)} 0 A {_n(r)} {_n(r)} 0 0 1 {_n(w)} {_n(r)} '
        f'L {_n(w)} {_n(h-r)} A {_n(r)} {_n(r)} 0 0 1 {_n(w-r)} {_n(h)} '
        f'L {_n(bx+wedge)} {_n(h)} L {_n(tip_x)} {_n(tip_y)} L {_n(bx-wedge)} {_n(h)} '
        f'L {_n(r)} {_n(h)} A {_n(r)} {_n(r)} 0 0 1 0 {_n(h-r)} L 0 {_n(r)} '
        f'A {_n(r)} {_n(r)} 0 0 1 {_n(r)} 0 Z"/>'
    )


def _wedge_ellipse_callout(w, h, adj):
    tip_x = w / 2 + _adj(adj, "adj1", -20833) * w
    tip_y = h / 2 + _adj(adj, "adj2", 62500) * h
    cx, cy = w / 2, h / 2
    rx, ry = w / 2, h / 2
    angle = math.atan2(tip_y - cy, tip_x - cx)
    wedge_angle = 0.15
    x1 = cx + rx * math.cos(angle - wedge_angle)
    y1 = cy + ry * math.sin(angle - wedge_angle)
    x2 = cx + rx * math.cos(angle + wedge_angle)
    y2 = cy + ry * math.sin(angle + wedge_angle)
    return (
        f'<path d="M {_n(x1)} {_n(y1)} L {_n(tip_x)} {_n(tip_y)} L {_n(x2)} {_n(y2)} '
        f'A {_n(rx)} {_n(ry)} 0 1 1 {_n(x1)} {_n(y1)} Z"/>'
    )


def _cloud_callout(w, h, adj):
    tip_x = w / 2 + _adj(adj, "adj1", -20833) * w
    tip_y = h / 2 + _adj(adj, "adj2", 62500) * h
    r = min(w, h) * 0.15
    bx, by = w / 2, h / 2
    dx, dy = tip_x - bx, tip_y - by
    d1x, d1y = bx + dx * 0.33, by + dy * 0.33
    d2x, d2y = bx + dx * 0.66, by + dy * 0.66
    return (
        f'<path d="M {_n(r)} {_n(h)} A {_n(r)} {_n(r)} 0 0 1 0 {_n(h-r)} '
        f'A {_n(r)} {_n(r)} 0 0 1 {_n(r)} {_n(h - 2*r)} L {_n(r)} {_n(r)} '
        f'A {_n(r)} {_n(r)} 0 0 1 {_n(2*r)} 0 L {_n(w - 2*r)} 0 '
        f'A {_n(r)} {_n(r)} 0 0 1 {_n(w-r)} {_n(r)} L {_n(w-r)} {_n(h - 2*r)} '
        f'A {_n(r)} {_n(r)} 0 0 1 {_n(w - 2*r)} {_n(h-r)} A {_n(r)} {_n(r)} 0 0 1 {_n(w - 2*r)} {_n(h)} Z '
        f'M {_n(d1x)} {_n(d1y)} m {_n(r*0.25)} 0 a {_n(r*0.25)} {_n(r*0.25)} 0 1 1 {_n(-r*0.5)} 0 '
        f'a {_n(r*0.25)} {_n(r*0.25)} 0 1 1 {_n(r*0.5)} 0 Z '
        f'M {_n(d2x)} {_n(d2y)} m {_n(r*0.15)} 0 a {_n(r*0.15)} {_n(r*0.15)} 0 1 1 {_n(-r*0.3)} 0 '
        f'a {_n(r*0.15)} {_n(r*0.15)} 0 1 1 {_n(r*0.3)} 0 Z"/>'
    )


def _border_callout1(w, h, adj):
    y1 = _adj(adj, "adj1", 18750) * h
    x1 = _adj(adj, "adj2", -8333) * w
    y2 = _adj(adj, "adj3", 112500) * h
    x2 = _adj(adj, "adj4", -38333) * w
    return (
        f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h)} L 0 {_n(h)} Z '
        f'M {_n(x1)} {_n(y1)} L {_n(x2)} {_n(y2)}"/>'
    )


def _border_callout2(w, h, adj):
    y1 = _adj(adj, "adj1", 18750) * h
    x1 = _adj(adj, "adj2", -8333) * w
    y2 = _adj(adj, "adj3", 18750) * h
    x2 = _adj(adj, "adj4", -16667) * w
    y3 = _adj(adj, "adj5", 112500) * h
    x3 = _adj(adj, "adj6", -46667) * w
    return (
        f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h)} L 0 {_n(h)} Z '
        f'M {_n(x1)} {_n(y1)} L {_n(x2)} {_n(y2)} L {_n(x3)} {_n(y3)}"/>'
    )


def _border_callout3(w, h, adj):
    y1 = _adj(adj, "adj1", 18750) * h
    x1 = _adj(adj, "adj2", -8333) * w
    y2 = _adj(adj, "adj3", 18750) * h
    x2 = _adj(adj, "adj4", -16667) * w
    y3 = _adj(adj, "adj5", 100000) * h
    x3 = _adj(adj, "adj6", -16667) * w
    y4 = _adj(adj, "adj7", 112963) * h
    x4 = _adj(adj, "adj8", -46667) * w
    return (
        f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h)} L 0 {_n(h)} Z '
        f'M {_n(x1)} {_n(y1)} L {_n(x2)} {_n(y2)} L {_n(x3)} {_n(y3)} L {_n(x4)} {_n(y4)}"/>'
    )


# --------------------------------------------------------------------------------------
# Arcs
# --------------------------------------------------------------------------------------





# --------------------------------------------------------------------------------------
# Math
# --------------------------------------------------------------------------------------



# --------------------------------------------------------------------------------------
# Misc
# --------------------------------------------------------------------------------------



def _folded_corner(w, h, adj):
    fold = _adj(adj, "adj", 16667) * min(w, h)
    return (
        f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h-fold)} L {_n(w-fold)} {_n(h)} L 0 {_n(h)} Z '
        f'M {_n(w-fold)} {_n(h)} L {_n(w-fold)} {_n(h-fold)} L {_n(w)} {_n(h-fold)}"/>'
    )



def _frame(w, h, adj):
    t = _adj(adj, "adj1", 12500) * min(w, h)
    return (
        f'<path fill-rule="evenodd" d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h)} L 0 {_n(h)} Z '
        f'M {_n(t)} {_n(t)} L {_n(t)} {_n(h-t)} L {_n(w-t)} {_n(h-t)} L {_n(w-t)} {_n(t)} Z"/>'
    )


def _bevel(w, h, adj):
    t = _adj(adj, "adj", 12500) * min(w, h)
    return (
        f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h)} L 0 {_n(h)} Z '
        f'M {_n(t)} {_n(t)} L {_n(w-t)} {_n(t)} L {_n(w-t)} {_n(h-t)} L {_n(t)} {_n(h-t)} Z '
        f'M 0 0 L {_n(t)} {_n(t)} M {_n(w)} 0 L {_n(w-t)} {_n(t)} '
        f'M {_n(w)} {_n(h)} L {_n(w-t)} {_n(h-t)} M 0 {_n(h)} L {_n(t)} {_n(h-t)}"/>'
    )



def _snip1_rect(w, h, adj):
    d = _adj(adj, "adj", 16667) * min(w, h)
    return f'<polygon points="{_pts((0, 0), (w - d, 0), (w, d), (w, h), (0, h))}"/>'


def _snip2_same_rect(w, h, adj):
    d1 = _adj(adj, "adj1", 16667) * min(w, h)
    d2 = _adj(adj, "adj2", 0) * min(w, h)
    return f'<polygon points="{_pts((d1, 0), (w - d1, 0), (w, d1), (w, h - d2), (w - d2, h), (d2, h), (0, h - d2), (0, d1))}"/>'


def _snip2_diag_rect(w, h, adj):
    d1 = _adj(adj, "adj1", 16667) * min(w, h)
    d2 = _adj(adj, "adj2", 0) * min(w, h)
    return f'<polygon points="{_pts((d1, 0), (w, 0), (w, h - d2), (w - d2, h), (0, h), (0, d1))}"/>'


def _snip_round_rect(w, h, adj):
    r = _adj(adj, "adj1", 16667) * min(w, h)
    d = _adj(adj, "adj2", 16667) * min(w, h)
    return (
        f'<path d="M {_n(r)} 0 L {_n(w-d)} 0 L {_n(w)} {_n(d)} L {_n(w)} {_n(h)} L 0 {_n(h)} '
        f'L 0 {_n(r)} A {_n(r)} {_n(r)} 0 0 1 {_n(r)} 0 Z"/>'
    )


def _round1_rect(w, h, adj):
    r = min(0.5, max(0.0, _adj(adj, "adj", 16667))) * min(w, h)
    return (
        f'<path d="M 0 0 L {_n(w-r)} 0 A {_n(r)} {_n(r)} 0 0 1 {_n(w)} {_n(r)} '
        f'L {_n(w)} {_n(h)} L 0 {_n(h)} Z"/>'
    )


def _round2_same_rect(w, h, adj):
    r1 = min(0.5, max(0.0, _adj(adj, "adj1", 16667))) * min(w, h)
    r2 = min(0.5, max(0.0, _adj(adj, "adj2", 0))) * min(w, h)
    return (
        f'<path d="M {_n(r1)} 0 L {_n(w-r1)} 0 A {_n(r1)} {_n(r1)} 0 0 1 {_n(w)} {_n(r1)} '
        f'L {_n(w)} {_n(h-r2)} A {_n(r2)} {_n(r2)} 0 0 1 {_n(w-r2)} {_n(h)} L {_n(r2)} {_n(h)} '
        f'A {_n(r2)} {_n(r2)} 0 0 1 0 {_n(h-r2)} L 0 {_n(r1)} A {_n(r1)} {_n(r1)} 0 0 1 {_n(r1)} 0 Z"/>'
    )


def _round2_diag_rect(w, h, adj):
    r1 = min(0.5, max(0.0, _adj(adj, "adj1", 16667))) * min(w, h)
    r2 = min(0.5, max(0.0, _adj(adj, "adj2", 0))) * min(w, h)
    return (
        f'<path d="M {_n(r1)} 0 L {_n(w)} 0 L {_n(w)} {_n(h-r2)} '
        f'A {_n(r2)} {_n(r2)} 0 0 1 {_n(w-r2)} {_n(h)} L 0 {_n(h)} L 0 {_n(r1)} '
        f'A {_n(r1)} {_n(r1)} 0 0 1 {_n(r1)} 0 Z"/>'
    )


def _left_bracket(w, h, adj):
    r = _adj(adj, "adj", 8333) * h
    return (
        f'<path d="M {_n(w)} 0 L {_n(r)} 0 A {_n(r)} {_n(r)} 0 0 0 0 {_n(r)} L 0 {_n(h-r)} '
        f'A {_n(r)} {_n(r)} 0 0 0 {_n(r)} {_n(h)} L {_n(w)} {_n(h)}"/>'
    )


def _right_bracket(w, h, adj):
    r = _adj(adj, "adj", 8333) * h
    return (
        f'<path d="M 0 0 L {_n(w-r)} 0 A {_n(r)} {_n(r)} 0 0 1 {_n(w)} {_n(r)} L {_n(w)} {_n(h-r)} '
        f'A {_n(r)} {_n(r)} 0 0 1 {_n(w-r)} {_n(h)} L 0 {_n(h)}"/>'
    )


def _left_brace(w, h, adj):
    r = _adj(adj, "adj1", 8333) * h
    mid = _adj(adj, "adj2", 50000) * h
    return (
        f'<path d="M {_n(w)} 0 A {_n(w/2)} {_n(r)} 0 0 0 {_n(w/2)} {_n(r)} L {_n(w/2)} {_n(mid-r)} '
        f'A {_n(w/2)} {_n(r)} 0 0 1 0 {_n(mid)} A {_n(w/2)} {_n(r)} 0 0 1 {_n(w/2)} {_n(mid+r)} '
        f'L {_n(w/2)} {_n(h-r)} A {_n(w/2)} {_n(r)} 0 0 0 {_n(w)} {_n(h)}"/>'
    )


def _right_brace(w, h, adj):
    r = _adj(adj, "adj1", 8333) * h
    mid = _adj(adj, "adj2", 50000) * h
    return (
        f'<path d="M 0 0 A {_n(w/2)} {_n(r)} 0 0 1 {_n(w/2)} {_n(r)} L {_n(w/2)} {_n(mid-r)} '
        f'A {_n(w/2)} {_n(r)} 0 0 0 {_n(w)} {_n(mid)} A {_n(w/2)} {_n(r)} 0 0 0 {_n(w/2)} {_n(mid+r)} '
        f'L {_n(w/2)} {_n(h-r)} A {_n(w/2)} {_n(r)} 0 0 1 0 {_n(h)}"/>'
    )






def _wave(w, h, adj):
    dy = _adj(adj, "adj1", 12500) * h
    dx = _adj(adj, "adj2", 0) * w
    return (
        f'<path d="M {_n(dx)} {_n(dy)} C {_n(dx + w*0.25)} 0, {_n(dx + w*0.5)} 0, {_n(w)} {_n(dy)} '
        f'L {_n(w-dx)} {_n(h-dy)} C {_n(w - dx - w*0.25)} {_n(h)}, {_n(w - dx - w*0.5)} {_n(h)}, 0 {_n(h-dy)} Z"/>'
    )


def _double_wave(w, h, adj):
    dy = _adj(adj, "adj1", 6250) * h
    dx = _adj(adj, "adj2", 0) * w
    return (
        f'<path d="M {_n(dx)} {_n(dy)} C {_n(dx + w*0.167)} 0, {_n(dx + w*0.333)} {_n(dy*2)}, {_n(w/2)} {_n(dy)} '
        f'C {_n(w/2 + w*0.167)} 0, {_n(w/2 + w*0.333)} {_n(dy*2)}, {_n(w)} {_n(dy)} '
        f'L {_n(w-dx)} {_n(h-dy)} C {_n(w - dx - w*0.167)} {_n(h)}, {_n(w - dx - w*0.333)} {_n(h - dy*2)}, {_n(w/2)} {_n(h-dy)} '
        f'C {_n(w/2 - w*0.167)} {_n(h)}, {_n(w/2 - w*0.333)} {_n(h - dy*2)}, 0 {_n(h-dy)} Z"/>'
    )




# --------------------------------------------------------------------------------------
# Presets taken verbatim from the specification
#
# The generators above are hand-written approximations, which is the right trade for a
# shape whose outline is a rounded rectangle or an ellipse.  It is the wrong trade for a
# callout with nineteen guides and eleven vertices, or an action button whose symbol is a
# dozen line segments: approximating those by eye produces something recognisably not the
# shape PowerPoint draws, and offers no way to check it short of looking at a render.
#
# So these presets carry their ECMA-376 Appendix D definition as data -- the same
# ``avLst`` / ``gdLst`` / ``pathLst`` the specification publishes -- and evaluate it with
# :mod:`pptx2svg.guides`, the evaluator ``a:custGeom`` already uses.  The geometry is then
# exact by construction rather than by judgement, and adding a preset becomes a
# transcription that can be diffed against the spec instead of a drawing exercise.
#
# Evaluation happens in pixels, so ``precise=True`` keeps the fractional part of every
# guide.  The spec's integer rounding assumes Office's EMU-sized coordinate space, where
# it is invisible; at pixel scale it would be a visible half-pixel error per vertex.
# --------------------------------------------------------------------------------------

#: Path-level fill modes (ECMA-376 §20.1.10.36) that shade the *shape's own* fill rather
#: than naming a colour: the underside of a curved arrow, the curl of a scroll, the bevel
#: of an action button.  Nothing at this layer knows what that fill is -- a generator sees
#: only width, height and adjustments -- so the shading is a neutral overlay of comparable
#: strength, which reads correctly over any base colour.
#:
#: The overlay is painted *over a copy of the path in the shape's own colour*, not over
#: whatever happens to be behind.  That distinction is the whole game: an action button's
#: bevel lies on top of the already-filled button, so shading alone would look right, but
#: a curved arrow's shaded underside is a region **no other path covers** -- shading the
#: slide background there gives a washed-out grey wedge instead of a darker arrow.
_SHADE_OVERLAYS = {
    "lighten": 'fill="#ffffff" fill-opacity="0.4"',
    "lightenLess": 'fill="#ffffff" fill-opacity="0.2"',
    "darken": 'fill="#000000" fill-opacity="0.4"',
    "darkenLess": 'fill="#000000" fill-opacity="0.2"',
}


class _P:
    """One ``a:path`` of a preset definition: its commands, and how it is painted."""

    __slots__ = ("commands", "fill", "stroke", "space")

    def __init__(
        self,
        *commands,
        fill: str = "norm",
        stroke: bool = True,
        space: "tuple[float, float] | None" = None,
    ):
        self.commands = commands
        self.fill = fill
        self.stroke = stroke
        #: ``a:path@w``/``@h``.  A path may be authored in its own coordinate space and
        #: scaled onto the shape -- the chart markers are drawn on a 10x10 grid, and
        #: ``flowChartOfflineStorage`` on a 2x4 one.  Without this they render as a
        #: few-pixel smudge in the corner of whatever box they are given.
        self.space = space

    def attributes(self) -> str:
        """Presentation attributes that override what the caller splices onto the
        wrapping ``<g>``.  A normally-painted path overrides nothing and inherits both."""
        parts = []
        if self.fill == "none":
            parts.append('fill="none"')
        if not self.stroke:
            parts.append('stroke="none"')
        return " ".join(parts)

    @property
    def overlay(self) -> str:
        """The shading pass for this path, if it has one; see :data:`_SHADE_OVERLAYS`."""
        return _SHADE_OVERLAYS.get(self.fill, "")


class _Spec:
    """A preset's specification: adjustment defaults, guides, and paths."""

    __slots__ = ("adjustments", "guides", "paths")

    def __init__(self, *, adjustments=(), guides=(), paths=()):
        self.adjustments = adjustments
        self.guides = guides
        self.paths = paths


def _spec_geometry(spec: "_Spec", w: float, h: float, adj: dict) -> str:
    """Evaluate one spec-defined preset into SVG.

    A shape's own ``a:avLst`` overrides the specification's defaults by name, and the
    values pass through raw: the guides consume them in the spec's own 1/1000-percent
    units (``*/ h adj1 100000``), so scaling them here would apply the division twice.
    """
    adjustments = [
        (name, f"val {adj.get(name, default)}") for name, default in spec.adjustments
    ]
    variables = evaluate_guides([adjustments, list(spec.guides)], w, h, precise=True)

    rendered = []
    for path in spec.paths:
        data = _spec_path_data(path.commands, variables, scale=_path_scale(path, w, h))
        if not data:
            continue
        attributes = path.attributes()
        rendered.append(f'<path d="{data}"' + (f" {attributes}" if attributes else "") + "/>")
        # A shaded path is drawn twice: once inheriting the shape's fill, then again
        # with the overlay on top.  See _SHADE_OVERLAYS for why the first pass matters.
        if path.overlay:
            rendered.append(f'<path d="{data}" {path.overlay} stroke="none"/>')

    if not rendered:
        return f'<rect width="{_n(w)}" height="{_n(h)}"/>'
    if len(rendered) == 1 and not spec.paths[0].attributes():
        return rendered[0]
    # More than one path, or one painted differently from the shape: wrap them so the
    # caller still has a single element to splice fill and stroke onto, and let the
    # children inherit it or override it.
    return f"<g>{''.join(rendered)}</g>"


def _path_scale(path: "_P", w: float, h: float) -> tuple[float, float]:
    """How a path's own coordinate space maps onto the shape box.

    Most preset paths are authored directly in shape coordinates and need no scaling.
    A few declare ``a:path@w``/``@h`` and are authored in a space of that size instead:
    the chart markers on a 10x10 grid, ``flowChartOfflineStorage`` on a 2x2 one.

    The scale is applied to the coordinates rather than emitted as an SVG ``transform``,
    because a transform would scale the *stroke width* with them -- an outline authored
    2 units wide becomes a 160-pixel slab once the path is blown up to fill a shape.
    Scaling the numbers keeps the stroke in shape units, where the caller set it.
    """
    if not path.space:
        return 1.0, 1.0
    space_width, space_height = path.space
    return (
        w / space_width if space_width else 1.0,
        h / space_height if space_height else 1.0,
    )


def _spec_path_data(commands, variables: dict, scale: tuple[float, float] = (1.0, 1.0)) -> str:
    """Build an SVG ``d`` string, tracking the pen so ``arcTo`` can be converted.

    DrawingML's ``arcTo`` is relative to wherever the pen already is -- it names a sweep,
    not an end point -- so the conversion needs the position the preceding command left.

    ``scale`` maps a path-local coordinate space onto the shape; see :func:`_path_scale`.
    The pen is tracked in the path's *own* coordinates and only scaled on the way out,
    which matters for arcs: ``stAng`` is a geometric angle measured in the space the path
    was authored in, so converting it to the ellipse's parametric angle has to use the
    unscaled radii.  Scaling them first silently rotates every arc that does not start on
    an axis -- invisible in a square box, where the two radii scale equally, and worth
    0.19 of silhouette overlap against PowerPoint on a stretched `cloud`.

    Scaling afterwards is exact: the scale is axis-aligned and so are the ellipse axes,
    DrawingML having no rotated ``arcTo``.
    """
    scale_x, scale_y = scale

    def value(token) -> float:
        return resolve_value(token, variables)

    parts: list[str] = []
    x = y = start_x = start_y = 0.0

    for command in commands:
        kind = command[0]
        if kind in ("M", "L"):
            x, y = value(command[1]), value(command[2])
            parts.append(f"{kind} {_n(x * scale_x)} {_n(y * scale_y)}")
            if kind == "M":
                start_x, start_y = x, y
        elif kind == "Z":
            parts.append("Z")
            x, y = start_x, start_y
        elif kind == "A":
            width_radius, height_radius = value(command[1]), value(command[2])
            start_angle, sweep_angle = value(command[3]), value(command[4])
            if sweep_angle == 0 or (width_radius == 0 and height_radius == 0):
                continue
            for x, y, large, sweep in arc_segments(
                x, y, width_radius, height_radius, start_angle, sweep_angle
            ):
                parts.append(
                    f"A {_n(width_radius * scale_x)} {_n(height_radius * scale_y)} 0 "
                    f"{large} {sweep} {_n(x * scale_x)} {_n(y * scale_y)}"
                )
        elif kind in ("Q", "C"):
            points = [
                (value(command[i]), value(command[i + 1]))
                for i in range(1, len(command), 2)
            ]
            parts.append(
                f"{kind} "
                + ", ".join(f"{_n(px * scale_x)} {_n(py * scale_y)}" for px, py in points)
            )
            x, y = points[-1]

    return " ".join(parts)


def _spec_generator(name: str) -> Generator:
    """Adapt a spec entry to the ``(w, h, adj) -> str`` shape of every other generator."""

    def generate(w: float, h: float, adj: dict) -> str:
        return _spec_geometry(SPEC_PRESETS[name], w, h, adj)

    return generate


#: The specification's definitions, inflated into the objects :func:`_spec_geometry`
#: draws from.  The data itself is generated -- see :mod:`pptx2svg.render.preset_specs`
#: and ``tools/derive_preset_geometry.py`` -- so a shape that looks wrong is diffed
#: against ECMA-376 rather than argued about.
SPEC_PRESETS: dict[str, "_Spec"] = {
    name: _Spec(
        adjustments=adjustments,
        guides=guides,
        paths=tuple(
            _P(*commands, fill=fill, stroke=stroke, space=space)
            for fill, stroke, space, commands in paths
        ),
    )
    for name, (adjustments, guides, paths) in PRESET_SPECS.items()
}


PRESET_GEOMETRIES: dict[str, Generator] = {
    # Basic
    "rect": _rect,
    "ellipse": _ellipse,
    "roundRect": _round_rect,
    "triangle": _triangle,
    "rtTriangle": _rt_triangle,
    "diamond": _diamond,
    "decagon": lambda w, h, adj: _regular_polygon(w, h, 10),
    "dodecagon": lambda w, h, adj: _regular_polygon(w, h, 12),
    "line": _line,
    # Stars and seals
    "irregularSeal1": _irregular_seal1,
    "irregularSeal2": _irregular_seal2,
    # Arrows
    "upArrow": _up_arrow,
    "downArrow": _down_arrow,
    "upDownArrow": _up_down_arrow,
    "homePlate": _home_plate,
    # Connectors
    "straightConnector1": _straight_connector1,
    # Flowchart
    "flowChartProcess": _rect,
    "flowChartAlternateProcess": _flow_alternate_process,
    "flowChartDecision": _flow_decision,
    "flowChartInputOutput": _flow_input_output,
    "flowChartPredefinedProcess": _flow_predefined_process,
    "flowChartInternalStorage": _flow_internal_storage,
    "flowChartPreparation": _flow_preparation,
    "flowChartManualInput": _flow_manual_input,
    "flowChartManualOperation": _flow_manual_operation,
    "flowChartConnector": _ellipse,
    "flowChartOffpageConnector": _flow_offpage_connector,
    "flowChartPunchedCard": _flow_punched_card,
    "flowChartCollate": _flow_collate,
    "flowChartSort": _flow_sort,
    "flowChartExtract": _flow_extract,
    "flowChartMerge": _flow_merge,
    "flowChartDelay": _flow_delay,
    "flowChartDisplay": _flow_display,
    # Callouts
    "wedgeRectCallout": _wedge_rect_callout,
    "wedgeRoundRectCallout": _wedge_round_rect_callout,
    "wedgeEllipseCallout": _wedge_ellipse_callout,
    "cloudCallout": _cloud_callout,
    "borderCallout1": _border_callout1,
    "borderCallout2": _border_callout2,
    "borderCallout3": _border_callout3,
    # Arcs
    # Math
    # Misc
    "cloud": _cloud,
    "heart": _heart,
    "foldedCorner": _folded_corner,
    "frame": _frame,
    "bevel": _bevel,
    "snip1Rect": _snip1_rect,
    "snip2SameRect": _snip2_same_rect,
    "snip2DiagRect": _snip2_diag_rect,
    "snipRoundRect": _snip_round_rect,
    "round1Rect": _round1_rect,
    "round2SameRect": _round2_same_rect,
    "round2DiagRect": _round2_diag_rect,
    "leftBracket": _left_bracket,
    "rightBracket": _right_bracket,
    "leftBrace": _left_brace,
    "rightBrace": _right_brace,
    "wave": _wave,
    "doubleWave": _double_wave,
}


# Specification-derived presets are registered last so that, if a name ever appears in
# both tables, the exact geometry wins over the hand-written approximation.
PRESET_GEOMETRIES.update({name: _spec_generator(name) for name in SPEC_PRESETS})

#: Misspellings that appear in real files, pointed at the shape they meant.  Resolving
#: them here rather than giving them their own entry means an alias cannot drift away
#: from the geometry it aliases -- which it did: promoting `bentUpArrow` to the
#: specification left `bendUpArrow` still drawing the old approximation.
PRESET_ALIASES = {"bendUpArrow": "bentUpArrow"}
PRESET_GEOMETRIES.update(
    {alias: PRESET_GEOMETRIES[target] for alias, target in PRESET_ALIASES.items()}
)

def preset_geometry_svg(preset: str, width: float, height: float, adjust: dict) -> str:
    """One SVG element for a preset shape; unknown presets fall back to a rectangle."""
    generator = PRESET_GEOMETRIES.get(preset)
    if generator is not None:
        return generator(width, height, adjust)
    return f'<rect width="{_n(width)}" height="{_n(height)}"/>'


def render_geometry(geometry: m.Geometry, width: float, height: float) -> str:
    if isinstance(geometry, m.PresetGeometry):
        return preset_geometry_svg(geometry.preset, width, height, geometry.adjust_values)
    if isinstance(geometry, m.CustomGeometry) and geometry.paths:
        return _render_custom_geometry(geometry.paths, width, height)
    return f'<rect width="{_n(width)}" height="{_n(height)}"/>'


def _render_custom_geometry(
    paths: list[m.CustomGeometryPath], shape_width: float, shape_height: float
) -> str:
    if len(paths) == 1:
        return _render_custom_path(paths[0], shape_width, shape_height)
    inner = "".join(_render_custom_path(path, shape_width, shape_height) for path in paths)
    return f"<g>{inner}</g>"


def _render_custom_path(
    path: m.CustomGeometryPath, shape_width: float, shape_height: float
) -> str:
    """Custom paths are authored in their own coordinate space; scale it onto the shape."""
    scale_x = shape_width / path.width if path.width > 0 else 1.0
    scale_y = shape_height / path.height if path.height > 0 else 1.0
    return f'<path d="{path.commands}" transform="scale({_n(scale_x)}, {_n(scale_y)})"/>'
