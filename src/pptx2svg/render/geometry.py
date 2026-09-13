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
from ..guides import arc_endpoint, evaluate_guides, resolve_value

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


def _angle(adj: dict, name: str, default: float) -> float:
    """OOXML angle (1/60,000 degrees) -> radians."""
    return math.radians(adj.get(name, default) / 60000)


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


def _star_polygon(w: float, h: float, points: int, inner_ratio: float) -> str:
    cx, cy = w / 2, h / 2
    coords = []
    for i in range(points * 2):
        angle = math.tau * i / (points * 2) - math.pi / 2
        radius = 1.0 if i % 2 == 0 else inner_ratio
        coords.append((cx + cx * radius * math.cos(angle), cy + cy * radius * math.sin(angle)))
    return f'<polygon points="{_pts(*coords)}"/>'


def _elliptical_arc_endpoints(
    w: float, h: float, start_angle: float, end_angle: float
) -> tuple[float, float, float, float, int]:
    """Shared setup for arc/chord/pie/blockArc.

    OOXML measures angles clockwise from 3 o'clock; SVG's y axis points down, so the
    sine is negated and the sweep flag is 0 (counter-clockwise in SVG terms).
    """
    rx, ry = w / 2, h / 2
    cx, cy = rx, ry
    x1 = cx + rx * math.cos(start_angle)
    y1 = cy - ry * math.sin(start_angle)
    x2 = cx + rx * math.cos(end_angle)
    y2 = cy - ry * math.sin(end_angle)
    sweep = start_angle - end_angle
    if sweep < 0:
        sweep += math.tau
    return x1, y1, x2, y2, (1 if sweep > math.pi else 0)


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


def _parallelogram(w, h, adj):
    offset = _adj(adj, "adj", 25000) * w
    return f'<polygon points="{_pts((offset, 0), (w, 0), (w - offset, h), (0, h))}"/>'


def _trapezoid(w, h, adj):
    offset = _adj(adj, "adj", 25000) * w
    return f'<polygon points="{_pts((offset, 0), (w - offset, 0), (w, h), (0, h))}"/>'


def _hexagon(w, h, adj):
    offset = _adj(adj, "adj", 25000) * w
    return (
        f'<polygon points="{_pts((offset, 0), (w - offset, 0), (w, h/2), (w - offset, h), (offset, h), (0, h/2))}"/>'
    )


def _star4(w, h, adj):
    cx, cy = w / 2, h / 2
    ir = 0.38
    return f'<polygon points="{_pts((cx, 0), (cx + cx*ir, cy - cy*ir), (w, cy), (cx + cx*ir, cy + cy*ir), (cx, h), (cx - cx*ir, cy + cy*ir), (0, cy), (cx - cx*ir, cy - cy*ir))}"/>'


def _right_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 50000) * h
    head_l = _adj(adj, "adj2", 50000) * w
    top = (h - head_w) / 2
    bottom = h - top
    shaft = w - head_l
    return f'<polygon points="{_pts((0, top), (shaft, top), (shaft, 0), (w, h/2), (shaft, h), (shaft, bottom), (0, bottom))}"/>'


def _left_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 50000) * h
    head_l = _adj(adj, "adj2", 50000) * w
    top = (h - head_w) / 2
    bottom = h - top
    return f'<polygon points="{_pts((head_l, top), (head_l, 0), (0, h/2), (head_l, h), (head_l, bottom), (w, bottom), (w, top))}"/>'


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


def _bent_connector2(w, h, adj):
    return f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h)}"/>'


def _bent_connector3(w, h, adj):
    mid_x = _adj(adj, "adj1", 50000) * w
    return f'<path d="M 0 0 L {_n(mid_x)} 0 L {_n(mid_x)} {_n(h)} L {_n(w)} {_n(h)}"/>'


def _bent_connector4(w, h, adj):
    mid_x = _adj(adj, "adj1", 50000) * w
    mid_y = _adj(adj, "adj2", 50000) * h
    return (
        f'<path d="M 0 0 L {_n(mid_x)} 0 L {_n(mid_x)} {_n(mid_y)} '
        f'L {_n(w)} {_n(mid_y)} L {_n(w)} {_n(h)}"/>'
    )


def _bent_connector5(w, h, adj):
    x1 = _adj(adj, "adj1", 50000) * w
    mid_y = _adj(adj, "adj2", 50000) * h
    x2 = _adj(adj, "adj3", 50000) * w
    return (
        f'<path d="M 0 0 L {_n(x1)} 0 L {_n(x1)} {_n(mid_y)} L {_n(x2)} {_n(mid_y)} '
        f'L {_n(x2)} {_n(h)} L {_n(w)} {_n(h)}"/>'
    )


def _curved_connector2(w, h, adj):
    return f'<path d="M 0 0 C {_n(w)} 0 0 {_n(h)} {_n(w)} {_n(h)}"/>'


def _curved_connector3(w, h, adj):
    mid_x = _adj(adj, "adj1", 50000) * w
    return f'<path d="M 0 0 C {_n(mid_x)} 0 {_n(mid_x)} {_n(h)} {_n(w)} {_n(h)}"/>'


def _curved_connector4(w, h, adj):
    mid_x = _adj(adj, "adj1", 50000) * w
    mid_y = _adj(adj, "adj2", 50000) * h
    return (
        f'<path d="M 0 0 C {_n(mid_x)} 0 {_n(mid_x)} {_n(mid_y)} {_n(mid_x)} {_n(mid_y)} '
        f'S {_n(w)} {_n(mid_y)} {_n(w)} {_n(h)}"/>'
    )


def _curved_connector5(w, h, adj):
    x1 = _adj(adj, "adj1", 50000) * w
    mid_y = _adj(adj, "adj2", 50000) * h
    x2 = _adj(adj, "adj3", 50000) * w
    return (
        f'<path d="M 0 0 C {_n(x1)} 0 {_n(x1)} {_n(mid_y)} {_n(x1)} {_n(mid_y)} '
        f'S {_n(x2)} {_n(mid_y)} {_n(x2)} {_n(h)} S {_n(w)} {_n(h)} {_n(w)} {_n(h)}"/>'
    )


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


def _left_right_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 50000) * h
    head_l = _adj(adj, "adj2", 50000) * w
    top = (h - head_w) / 2
    bottom = h - top
    return f'<polygon points="{_pts((head_l, top), (head_l, 0), (0, h/2), (head_l, h), (head_l, bottom), (w - head_l, bottom), (w - head_l, h), (w, h/2), (w - head_l, 0), (w - head_l, top))}"/>'


def _up_down_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 50000) * w
    head_l = _adj(adj, "adj2", 50000) * h
    left = (w - head_w) / 2
    right = w - left
    return f'<polygon points="{_pts((left, head_l), (0, head_l), (w/2, 0), (w, head_l), (right, head_l), (right, h - head_l), (w, h - head_l), (w/2, h), (0, h - head_l), (left, h - head_l))}"/>'


def _notched_right_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 50000) * h
    head_l = _adj(adj, "adj2", 50000) * w
    top = (h - head_w) / 2
    bottom = h - top
    shaft = w - head_l
    notch = head_l * 0.5
    return f'<polygon points="{_pts((0, top), (shaft, top), (shaft, 0), (w, h/2), (shaft, h), (shaft, bottom), (0, bottom), (notch, h/2))}"/>'


def _striped_right_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 50000) * h
    head_l = _adj(adj, "adj2", 50000) * w
    top = (h - head_w) / 2
    bottom = h - top
    shaft = w - head_l
    sw = w * 0.05
    return (
        f'<path d="M 0 {_n(top)} L {_n(sw)} {_n(top)} L {_n(sw)} {_n(bottom)} L 0 {_n(bottom)} Z '
        f'M {_n(sw*1.5)} {_n(top)} L {_n(sw*2.5)} {_n(top)} L {_n(sw*2.5)} {_n(bottom)} L {_n(sw*1.5)} {_n(bottom)} Z '
        f'M {_n(sw*3)} {_n(top)} L {_n(shaft)} {_n(top)} L {_n(shaft)} 0 L {_n(w)} {_n(h/2)} '
        f'L {_n(shaft)} {_n(h)} L {_n(shaft)} {_n(bottom)} L {_n(sw*3)} {_n(bottom)} Z"/>'
    )


def _chevron(w, h, adj):
    offset = _adj(adj, "adj", 50000) * w
    return f'<polygon points="{_pts((0, 0), (w - offset, 0), (w, h/2), (w - offset, h), (0, h), (offset, h/2))}"/>'


def _home_plate(w, h, adj):
    offset = _adj(adj, "adj", 50000) * min(w, h)
    return f'<polygon points="{_pts((0, 0), (w - offset, 0), (w, h/2), (w - offset, h), (0, h))}"/>'


def _left_right_up_arrow(w, h, adj):
    side = min(w, h)
    head_w = _adj(adj, "adj1", 25000) * side
    head_l = _adj(adj, "adj2", 25000) * side
    body_w = _adj(adj, "adj3", 25000) * side
    cx = w / 2
    bh = body_w / 2
    body_mid = h - head_l - body_w
    arm_y = h / 2 + body_mid / 2
    return f'<polygon points="{_pts((cx, 0), (cx + head_w/2, head_l), (cx + bh, head_l), (cx + bh, body_mid), (w - head_l, body_mid), (w - head_l, arm_y - head_w/2), (w, arm_y), (w - head_l, arm_y + head_w/2), (w - head_l, body_mid + body_w), (cx - bh, body_mid + body_w), (head_l, body_mid + body_w), (head_l, arm_y + head_w/2), (0, arm_y), (head_l, arm_y - head_w/2), (head_l, body_mid), (cx - bh, body_mid), (cx - bh, head_l), (cx - head_w/2, head_l))}"/>'


def _quad_arrow(w, h, adj):
    side = min(w, h)
    head_w = _adj(adj, "adj1", 22500) * side
    head_l = _adj(adj, "adj2", 22500) * side
    body_w = _adj(adj, "adj3", 11250) * side
    cx, cy = w / 2, h / 2
    bh = body_w / 2
    return f'<polygon points="{_pts((cx, 0), (cx + head_w/2, head_l), (cx + bh, head_l), (cx + bh, cy - bh), (w - head_l, cy - bh), (w - head_l, cy - head_w/2), (w, cy), (w - head_l, cy + head_w/2), (w - head_l, cy + bh), (cx + bh, cy + bh), (cx + bh, h - head_l), (cx + head_w/2, h - head_l), (cx, h), (cx - head_w/2, h - head_l), (cx - bh, h - head_l), (cx - bh, cy + bh), (head_l, cy + bh), (head_l, cy + head_w/2), (0, cy), (head_l, cy - head_w/2), (head_l, cy - bh), (cx - bh, cy - bh), (cx - bh, head_l), (cx - head_w/2, head_l))}"/>'


def _bent_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 25000) * h
    head_l = _adj(adj, "adj2", 25000) * w
    body_w = _adj(adj, "adj3", 25000) * h
    shaft = w - head_l
    top = head_w / 2 - body_w / 2
    bottom = head_w / 2 + body_w / 2
    return f'<polygon points="{_pts((shaft, top), (shaft, 0), (w, head_w/2), (shaft, head_w), (shaft, bottom), (body_w, bottom), (body_w, h), (0, h), (0, h - body_w))}"/>'


def _bend_up_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 25000) * w
    head_l = _adj(adj, "adj2", 25000) * h
    body_w = _adj(adj, "adj3", 25000) * w
    cx = w - head_w / 2
    left = cx - body_w / 2
    right = cx + body_w / 2
    return f'<polygon points="{_pts((cx - head_w/2, head_l), (cx, 0), (cx + head_w/2, head_l), (right, head_l), (right, h - body_w), (body_w, h - body_w), (body_w, h), (0, h), (0, h - body_w), (left, h - body_w), (left, head_l))}"/>'


def _left_up_arrow(w, h, adj):
    side = min(w, h)
    head_w = _adj(adj, "adj1", 25000) * side
    head_l = _adj(adj, "adj2", 25000) * side
    body_w = _adj(adj, "adj3", 25000) * side
    bh = body_w / 2
    top_cx = w - head_w / 2
    left_cy = h - head_w / 2
    return f'<polygon points="{_pts((top_cx, 0), (top_cx + head_w/2, head_l), (top_cx + bh, head_l), (top_cx + bh, left_cy - bh), (head_l, left_cy - bh), (head_l, left_cy - head_w/2), (0, left_cy), (head_l, left_cy + head_w/2), (head_l, left_cy + bh), (top_cx - bh, left_cy + bh), (top_cx - bh, head_l), (top_cx - head_w/2, head_l))}"/>'


def _uturn_arrow(w, h, adj):
    head_w = _adj(adj, "adj1", 25000) * w
    head_l = _adj(adj, "adj2", 25000) * h
    body_w = _adj(adj, "adj3", 25000) * w
    arc_r = w * 0.35
    cx = w / 2
    arrow_start = h - head_l
    body_right = w - (head_w / 2 - body_w / 2)
    body_left = w - (head_w / 2 + body_w / 2)
    return (
        f'<path d="M {_n(w - head_w)} {_n(arrow_start)} L {_n(w - head_w/2)} {_n(h)} '
        f'L {_n(w)} {_n(arrow_start)} L {_n(body_right)} {_n(arrow_start)} L {_n(body_right)} {_n(arc_r)} '
        f'A {_n(arc_r)} {_n(arc_r)} 0 0 0 {_n(body_w)} {_n(arc_r)} L {_n(body_w)} {_n(arrow_start)} '
        f'L 0 {_n(arrow_start)} L 0 {_n(arc_r)} A {_n(cx)} {_n(cx)} 0 0 1 {_n(body_left)} {_n(arc_r)} '
        f'L {_n(body_left)} {_n(arrow_start)} Z"/>'
    )


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


def _flow_document(w, h, adj):
    bh = h * 0.83
    return (
        f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(bh)} '
        f'C {_n(w*0.75)} {_n(h)}, {_n(w*0.25)} {_n(h*0.66)}, 0 {_n(bh)} Z"/>'
    )


def _flow_multidocument(w, h, adj):
    dx, dy = w * 0.1, h * 0.1
    bw = w - dx
    bh = (h - dy) * 0.83
    return (
        f'<path d="M {_n(dx)} {_n(dy)} L {_n(w)} {_n(dy)} L {_n(w)} {_n(dy+bh)} '
        f'C {_n(w - bw*0.25)} {_n(h)}, {_n(dx + bw*0.25)} {_n(dy + (h-dy)*0.66)}, {_n(dx)} {_n(dy+bh)} Z '
        f'M {_n(dx/2)} {_n(dy/2)} L {_n(dx)} {_n(dy/2)} L {_n(dx)} {_n(dy)} '
        f'M 0 0 L {_n(dx/2)} 0 L {_n(dx/2)} {_n(dy/2)}"/>'
    )


def _flow_terminator(w, h, adj):
    r = h / 2
    return f'<rect width="{_n(w)}" height="{_n(h)}" rx="{_n(r)}" ry="{_n(r)}"/>'


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


def _flow_punched_tape(w, h, adj):
    wave = h * 0.1
    return (
        f'<path d="M 0 {_n(wave)} C {_n(w*0.25)} {_n(-wave)}, {_n(w*0.75)} {_n(wave*3)}, {_n(w)} {_n(wave)} '
        f'L {_n(w)} {_n(h-wave)} C {_n(w*0.75)} {_n(h+wave)}, {_n(w*0.25)} {_n(h-wave*3)}, 0 {_n(h-wave)} Z"/>'
    )


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


def _flow_online_storage(w, h, adj):
    arc_w = w * 0.15
    return (
        f'<path d="M {_n(arc_w)} 0 L {_n(w)} 0 L {_n(w)} {_n(h)} L {_n(arc_w)} {_n(h)} '
        f'A {_n(arc_w)} {_n(h/2)} 0 0 1 {_n(arc_w)} 0 Z"/>'
    )


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


def _flow_magnetic_tape(w, h, adj):
    r = min(w, h) / 2
    cx, cy = w / 2, h / 2
    return (
        f'<path d="M {_n(cx+r)} {_n(cy)} A {_n(r)} {_n(r)} 0 1 1 {_n(cx+r-0.01)} {_n(cy+0.01)} '
        f'L {_n(w)} {_n(cy)} L {_n(w)} {_n(h)} L {_n(w - r*0.3)} {_n(h)} '
        f'L {_n(cx + r*math.cos(math.pi/6))} {_n(cy + r*math.sin(math.pi/6))}"/>'
    )


def _flow_magnetic_disk(w, h, adj):
    ry = h * 0.15
    return (
        f'<path d="M 0 {_n(ry)} A {_n(w/2)} {_n(ry)} 0 0 1 {_n(w)} {_n(ry)} L {_n(w)} {_n(h-ry)} '
        f'A {_n(w/2)} {_n(ry)} 0 0 1 0 {_n(h-ry)} Z '
        f'M 0 {_n(ry)} A {_n(w/2)} {_n(ry)} 0 0 0 {_n(w)} {_n(ry)}"/>'
    )


def _flow_magnetic_drum(w, h, adj):
    rx = w * 0.15
    return (
        f'<path d="M {_n(rx)} 0 A {_n(rx)} {_n(h/2)} 0 0 0 {_n(rx)} {_n(h)} L {_n(w-rx)} {_n(h)} '
        f'A {_n(rx)} {_n(h/2)} 0 0 0 {_n(w-rx)} 0 Z '
        f'M {_n(w-rx)} 0 A {_n(rx)} {_n(h/2)} 0 0 1 {_n(w-rx)} {_n(h)}"/>'
    )


def _flow_summing_junction(w, h, adj):
    cx, cy = w / 2, h / 2
    rx, ry = w / 2, h / 2
    d = 0.707
    return (
        f'<path d="M {_n(cx+rx)} {_n(cy)} A {_n(rx)} {_n(ry)} 0 1 1 {_n(cx+rx-0.01)} {_n(cy-0.01)} Z '
        f'M {_n(cx - rx*d)} {_n(cy - ry*d)} L {_n(cx + rx*d)} {_n(cy + ry*d)} '
        f'M {_n(cx + rx*d)} {_n(cy - ry*d)} L {_n(cx - rx*d)} {_n(cy + ry*d)}"/>'
    )


def _flow_or(w, h, adj):
    cx, cy = w / 2, h / 2
    rx, ry = w / 2, h / 2
    return (
        f'<path d="M {_n(cx+rx)} {_n(cy)} A {_n(rx)} {_n(ry)} 0 1 1 {_n(cx+rx-0.01)} {_n(cy-0.01)} Z '
        f'M {_n(cx)} 0 L {_n(cx)} {_n(h)} M 0 {_n(cy)} L {_n(w)} {_n(cy)}"/>'
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


def _arc(w, h, adj):
    x1, y1, x2, y2, large = _elliptical_arc_endpoints(
        w, h, _angle(adj, "adj1", 16200000), _angle(adj, "adj2", 0)
    )
    return f'<path d="M {_n(x1)} {_n(y1)} A {_n(w/2)} {_n(h/2)} 0 {large} 0 {_n(x2)} {_n(y2)}"/>'


def _chord(w, h, adj):
    x1, y1, x2, y2, large = _elliptical_arc_endpoints(
        w, h, _angle(adj, "adj1", 2700000), _angle(adj, "adj2", 16200000)
    )
    return f'<path d="M {_n(x1)} {_n(y1)} A {_n(w/2)} {_n(h/2)} 0 {large} 0 {_n(x2)} {_n(y2)} Z"/>'


def _pie(w, h, adj):
    x1, y1, x2, y2, large = _elliptical_arc_endpoints(
        w, h, _angle(adj, "adj1", 0), _angle(adj, "adj2", 16200000)
    )
    return (
        f'<path d="M {_n(w/2)} {_n(h/2)} L {_n(x1)} {_n(y1)} '
        f'A {_n(w/2)} {_n(h/2)} 0 {large} 0 {_n(x2)} {_n(y2)} Z"/>'
    )


def _block_arc(w, h, adj):
    start = _angle(adj, "adj1", 10800000)
    end = _angle(adj, "adj2", 0)
    thickness = _adj(adj, "adj3", 25000)
    rx, ry = w / 2, h / 2
    cx, cy = rx, ry
    irx, iry = rx * (1 - thickness), ry * (1 - thickness)
    ox1, oy1, ox2, oy2, large = _elliptical_arc_endpoints(w, h, start, end)
    ix1 = cx + irx * math.cos(start)
    iy1 = cy - iry * math.sin(start)
    ix2 = cx + irx * math.cos(end)
    iy2 = cy - iry * math.sin(end)
    return (
        f'<path d="M {_n(ox1)} {_n(oy1)} A {_n(rx)} {_n(ry)} 0 {large} 0 {_n(ox2)} {_n(oy2)} '
        f'L {_n(ix2)} {_n(iy2)} A {_n(irx)} {_n(iry)} 0 {large} 1 {_n(ix1)} {_n(iy1)} Z"/>'
    )


# --------------------------------------------------------------------------------------
# Math
# --------------------------------------------------------------------------------------


def _math_plus(w, h, adj):
    t = _adj(adj, "adj1", 23520)
    tw, th = t * w, t * h
    lx = (w - tw) / 2
    rx = lx + tw
    ty = (h - th) / 2
    by = ty + th
    return f'<polygon points="{_pts((lx, 0), (rx, 0), (rx, ty), (w, ty), (w, by), (rx, by), (rx, h), (lx, h), (lx, by), (0, by), (0, ty), (lx, ty))}"/>'


def _math_minus(w, h, adj):
    th = _adj(adj, "adj1", 23520) * h
    ty = (h - th) / 2
    return f'<rect x="0" y="{_n(ty)}" width="{_n(w)}" height="{_n(th)}"/>'


def _math_multiply(w, h, adj):
    d = _adj(adj, "adj1", 23520) * min(w, h) * 0.5
    cx, cy = w / 2, h / 2
    return f'<polygon points="{_pts((cx, cy - d), (w - d, 0), (w, d), (cx + d, cy), (w, h - d), (w - d, h), (cx, cy + d), (d, h), (0, h - d), (cx - d, cy), (0, d), (d, 0))}"/>'


def _math_divide(w, h, adj):
    t = _adj(adj, "adj1", 23520)
    th = t * h
    ty = (h - th) / 2
    by = ty + th
    dot_r = min(w, h) * t * 0.5
    cx = w / 2
    top_dot_y = ty / 2
    bottom_dot_y = h - ty / 2
    return (
        f'<path d="M 0 {_n(ty)} L {_n(w)} {_n(ty)} L {_n(w)} {_n(by)} L 0 {_n(by)} Z '
        f'M {_n(cx+dot_r)} {_n(top_dot_y)} A {_n(dot_r)} {_n(dot_r)} 0 1 1 {_n(cx+dot_r-0.01)} {_n(top_dot_y-0.01)} Z '
        f'M {_n(cx+dot_r)} {_n(bottom_dot_y)} A {_n(dot_r)} {_n(dot_r)} 0 1 1 {_n(cx+dot_r-0.01)} {_n(bottom_dot_y-0.01)} Z"/>'
    )


def _math_equal(w, h, adj):
    t = _adj(adj, "adj1", 23520)
    gap = bar = t * h
    y1 = (h - gap) / 2 - bar
    y2 = (h + gap) / 2
    return (
        f'<path d="M 0 {_n(y1)} L {_n(w)} {_n(y1)} L {_n(w)} {_n(y1+bar)} L 0 {_n(y1+bar)} Z '
        f'M 0 {_n(y2)} L {_n(w)} {_n(y2)} L {_n(w)} {_n(y2+bar)} L 0 {_n(y2+bar)} Z"/>'
    )


def _math_not_equal(w, h, adj):
    t = _adj(adj, "adj1", 23520)
    gap = bar = t * h
    y1 = (h - gap) / 2 - bar
    y2 = (h + gap) / 2
    slash_w = w * 0.15
    sx = w / 2 - slash_w / 2
    return (
        f'<path d="M 0 {_n(y1)} L {_n(w)} {_n(y1)} L {_n(w)} {_n(y1+bar)} L 0 {_n(y1+bar)} Z '
        f'M 0 {_n(y2)} L {_n(w)} {_n(y2)} L {_n(w)} {_n(y2+bar)} L 0 {_n(y2+bar)} Z '
        f'M {_n(sx+slash_w)} {_n(y1-bar)} L {_n(sx + 2*slash_w)} {_n(y1-bar)} '
        f'L {_n(sx)} {_n(y2 + 2*bar)} L {_n(sx-slash_w)} {_n(y2 + 2*bar)} Z"/>'
    )


# --------------------------------------------------------------------------------------
# Misc
# --------------------------------------------------------------------------------------


def _plus(w, h, adj):
    t = _adj(adj, "adj", 25000)
    lx, rx = t * w, w - t * w
    ty, by = t * h, h - t * h
    return f'<polygon points="{_pts((lx, 0), (rx, 0), (rx, ty), (w, ty), (w, by), (rx, by), (rx, h), (lx, h), (lx, by), (0, by), (0, ty), (lx, ty))}"/>'


def _corner(w, h, adj):
    cx = _adj(adj, "adj1", 50000) * w
    cy = _adj(adj, "adj2", 50000) * h
    return f'<polygon points="{_pts((0, 0), (cx, 0), (cx, cy), (w, cy), (w, h), (0, h))}"/>'


def _diag_stripe(w, h, adj):
    d = _adj(adj, "adj", 50000) * min(w, h)
    return f'<polygon points="{_pts((0, d), (d, 0), (w, 0), (0, h))}"/>'


def _folded_corner(w, h, adj):
    fold = _adj(adj, "adj", 16667) * min(w, h)
    return (
        f'<path d="M 0 0 L {_n(w)} 0 L {_n(w)} {_n(h-fold)} L {_n(w-fold)} {_n(h)} L 0 {_n(h)} Z '
        f'M {_n(w-fold)} {_n(h)} L {_n(w-fold)} {_n(h-fold)} L {_n(w)} {_n(h-fold)}"/>'
    )


def _plaque(w, h, adj):
    r = _adj(adj, "adj", 16667) * min(w, h)
    return (
        f'<path d="M 0 {_n(r)} A {_n(r)} {_n(r)} 0 0 1 {_n(r)} 0 L {_n(w-r)} 0 '
        f'A {_n(r)} {_n(r)} 0 0 1 {_n(w)} {_n(r)} L {_n(w)} {_n(h-r)} '
        f'A {_n(r)} {_n(r)} 0 0 1 {_n(w-r)} {_n(h)} L {_n(r)} {_n(h)} '
        f'A {_n(r)} {_n(r)} 0 0 1 0 {_n(h-r)} Z"/>'
    )


def _can(w, h, adj):
    ry = _adj(adj, "adj", 25000) * h * 0.5
    return (
        f'<path d="M 0 {_n(ry)} A {_n(w/2)} {_n(ry)} 0 0 1 {_n(w)} {_n(ry)} L {_n(w)} {_n(h-ry)} '
        f'A {_n(w/2)} {_n(ry)} 0 0 1 0 {_n(h-ry)} Z '
        f'M 0 {_n(ry)} A {_n(w/2)} {_n(ry)} 0 0 0 {_n(w)} {_n(ry)}"/>'
    )


def _cube(w, h, adj):
    d = _adj(adj, "adj", 25000) * min(w, h)
    return (
        f'<path d="M 0 {_n(d)} L {_n(d)} 0 L {_n(w)} 0 L {_n(w)} {_n(h-d)} L {_n(w-d)} {_n(h)} L 0 {_n(h)} Z '
        f'M 0 {_n(d)} L {_n(w-d)} {_n(d)} L {_n(w)} 0 M {_n(w-d)} {_n(d)} L {_n(w-d)} {_n(h)}"/>'
    )


def _donut(w, h, adj):
    t = _adj(adj, "adj", 25000)
    rx, ry = w / 2, h / 2
    irx, iry = rx * (1 - t), ry * (1 - t)
    cx, cy = rx, ry
    return (
        f'<path fill-rule="evenodd" d="M {_n(cx+rx)} {_n(cy)} A {_n(rx)} {_n(ry)} 0 1 1 {_n(cx-rx)} {_n(cy)} '
        f'A {_n(rx)} {_n(ry)} 0 1 1 {_n(cx+rx)} {_n(cy)} Z '
        f'M {_n(cx+irx)} {_n(cy)} A {_n(irx)} {_n(iry)} 0 1 0 {_n(cx-irx)} {_n(cy)} '
        f'A {_n(irx)} {_n(iry)} 0 1 0 {_n(cx+irx)} {_n(cy)} Z"/>'
    )


def _no_smoking(w, h, adj):
    t = _adj(adj, "adj", 18750)
    rx, ry = w / 2, h / 2
    cx, cy = rx, ry
    irx, iry = rx * (1 - t), ry * (1 - t)
    angle = math.pi / 4
    lx1 = cx + irx * math.cos(angle)
    ly1 = cy - iry * math.sin(angle)
    lx2 = cx - irx * math.cos(angle)
    ly2 = cy + iry * math.sin(angle)
    return (
        f'<path fill-rule="evenodd" d="M {_n(cx+rx)} {_n(cy)} A {_n(rx)} {_n(ry)} 0 1 1 {_n(cx-rx)} {_n(cy)} '
        f'A {_n(rx)} {_n(ry)} 0 1 1 {_n(cx+rx)} {_n(cy)} Z '
        f'M {_n(cx+irx)} {_n(cy)} A {_n(irx)} {_n(iry)} 0 1 0 {_n(cx-irx)} {_n(cy)} '
        f'A {_n(irx)} {_n(iry)} 0 1 0 {_n(cx+irx)} {_n(cy)} Z '
        f'M {_n(lx1)} {_n(ly1)} L {_n(lx2)} {_n(ly2)}"/>'
    )


def _smiley_face(w, h, adj):
    smile = _adj(adj, "adj", 4653)
    rx, ry = w / 2, h / 2
    cx, cy = rx, ry
    eye_rx, eye_ry = w * 0.06, h * 0.06
    eye_y = h * 0.35
    left_eye_x, right_eye_x = w * 0.35, w * 0.65
    mouth_y = h * 0.6
    mouth_w = w * 0.3
    curve = smile * h
    return (
        f'<path d="M {_n(cx+rx)} {_n(cy)} A {_n(rx)} {_n(ry)} 0 1 1 {_n(cx-rx)} {_n(cy)} '
        f'A {_n(rx)} {_n(ry)} 0 1 1 {_n(cx+rx)} {_n(cy)} Z '
        f'M {_n(left_eye_x+eye_rx)} {_n(eye_y)} A {_n(eye_rx)} {_n(eye_ry)} 0 1 1 {_n(left_eye_x-eye_rx)} {_n(eye_y)} '
        f'A {_n(eye_rx)} {_n(eye_ry)} 0 1 1 {_n(left_eye_x+eye_rx)} {_n(eye_y)} Z '
        f'M {_n(right_eye_x+eye_rx)} {_n(eye_y)} A {_n(eye_rx)} {_n(eye_ry)} 0 1 1 {_n(right_eye_x-eye_rx)} {_n(eye_y)} '
        f'A {_n(eye_rx)} {_n(eye_ry)} 0 1 1 {_n(right_eye_x+eye_rx)} {_n(eye_y)} Z '
        f'M {_n(cx-mouth_w)} {_n(mouth_y)} C {_n(cx - mouth_w*0.5)} {_n(mouth_y+curve)}, '
        f'{_n(cx + mouth_w*0.5)} {_n(mouth_y+curve)}, {_n(cx+mouth_w)} {_n(mouth_y)}"/>'
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


def _half_frame(w, h, adj):
    adj_x = _adj(adj, "adj1", 33333) * w
    adj_y = _adj(adj, "adj2", 33333) * h
    return f'<polygon points="{_pts((0, 0), (w, 0), (w, adj_y), (adj_x, adj_y), (adj_x, h), (0, h))}"/>'


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


def _bracket_pair(w, h, adj):
    r = _adj(adj, "adj", 16667) * min(w, h)
    return (
        f'<path d="M {_n(r)} 0 A {_n(r)} {_n(r)} 0 0 0 0 {_n(r)} L 0 {_n(h-r)} '
        f'A {_n(r)} {_n(r)} 0 0 0 {_n(r)} {_n(h)} '
        f'M {_n(w-r)} 0 A {_n(r)} {_n(r)} 0 0 1 {_n(w)} {_n(r)} L {_n(w)} {_n(h-r)} '
        f'A {_n(r)} {_n(r)} 0 0 1 {_n(w-r)} {_n(h)}"/>'
    )


def _brace_pair(w, h, adj):
    r = _adj(adj, "adj", 8333) * min(w, h)
    return (
        f'<path d="M {_n(r)} 0 A {_n(r)} {_n(r)} 0 0 0 0 {_n(r)} L 0 {_n(h/2-r)} '
        f'A {_n(r)} {_n(r)} 0 0 1 {_n(-r)} {_n(h/2)} A {_n(r)} {_n(r)} 0 0 1 0 {_n(h/2+r)} '
        f'L 0 {_n(h-r)} A {_n(r)} {_n(r)} 0 0 0 {_n(r)} {_n(h)} '
        f'M {_n(w-r)} 0 A {_n(r)} {_n(r)} 0 0 1 {_n(w)} {_n(r)} L {_n(w)} {_n(h/2-r)} '
        f'A {_n(r)} {_n(r)} 0 0 0 {_n(w+r)} {_n(h/2)} A {_n(r)} {_n(r)} 0 0 0 {_n(w)} {_n(h/2+r)} '
        f'L {_n(w)} {_n(h-r)} A {_n(r)} {_n(r)} 0 0 1 {_n(w-r)} {_n(h)}"/>'
    )


def _lightning_bolt(w, h, adj):
    ratios = [(0.55, 0.0), (0.3, 0.4), (0.52, 0.4), (0.25, 1.0), (0.75, 0.5), (0.52, 0.5), (0.85, 0.0)]
    return f'<polygon points="{_pts(*[(w*x, h*y) for x, y in ratios])}"/>'


def _moon(w, h, adj):
    t = _adj(adj, "adj", 50000) * w
    rx, ry = w / 2, h / 2
    irx = t / 2
    return (
        f'<path d="M {_n(w)} 0 A {_n(rx)} {_n(ry)} 0 1 0 {_n(w)} {_n(h)} '
        f'A {_n(irx)} {_n(ry)} 0 1 1 {_n(w)} 0 Z"/>'
    )


def _teardrop(w, h, adj):
    d = _adj(adj, "adj", 100000) * min(w, h) * 0.5
    rx, ry = w / 2, h / 2
    cx, cy = rx, ry
    return (
        f'<path d="M {_n(cx)} 0 L {_n(cx+d)} 0 L {_n(w)} {_n(cy - d + ry)} '
        f'A {_n(rx)} {_n(ry)} 0 1 1 {_n(cx)} 0 Z"/>'
    )


def _sun(w, h, adj):
    cx, cy = w / 2, h / 2
    inner = 0.35
    parts = [
        f'M {_n(cx + cx*inner)} {_n(cy)} '
        f'A {_n(cx*inner)} {_n(cy*inner)} 0 1 1 {_n(cx - cx*inner)} {_n(cy)} '
        f'A {_n(cx*inner)} {_n(cy*inner)} 0 1 1 {_n(cx + cx*inner)} {_n(cy)} Z'
    ]
    for i in range(8):
        angle = math.tau * i / 8
        x1 = cx + cx * inner * 1.15 * math.cos(angle)
        y1 = cy + cy * inner * 1.15 * math.sin(angle)
        x2 = cx + cx * math.cos(angle)
        y2 = cy + cy * math.sin(angle)
        parts.append(f"M {_n(x1)} {_n(y1)} L {_n(x2)} {_n(y2)}")
    return f'<path d="{" ".join(parts)}"/>'


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


def _ribbon(w, h, adj):
    tab_h = _adj(adj, "adj1", 16667) * h
    tab_w = _adj(adj, "adj2", 50000) * w
    fold = tab_w * 0.3
    return (
        f'<path d="M 0 {_n(tab_h)} L {_n(fold)} {_n(tab_h*1.5)} L {_n(fold)} {_n(h)} '
        f'L {_n(tab_w)} {_n(h-tab_h)} L {_n(w-tab_w)} {_n(h-tab_h)} L {_n(w-fold)} {_n(h)} '
        f'L {_n(w-fold)} {_n(tab_h*1.5)} L {_n(w)} {_n(tab_h)} L {_n(w)} 0 L {_n(w-tab_w)} 0 '
        f'L {_n(w-tab_w)} {_n(tab_h)} L {_n(tab_w)} {_n(tab_h)} L {_n(tab_w)} 0 L 0 0 Z"/>'
    )


def _ribbon2(w, h, adj):
    tab_h = _adj(adj, "adj1", 16667) * h
    tab_w = _adj(adj, "adj2", 50000) * w
    fold = tab_w * 0.3
    return (
        f'<path d="M 0 {_n(h-tab_h)} L {_n(fold)} {_n(h - tab_h*1.5)} L {_n(fold)} 0 '
        f'L {_n(tab_w)} {_n(tab_h)} L {_n(w-tab_w)} {_n(tab_h)} L {_n(w-fold)} 0 '
        f'L {_n(w-fold)} {_n(h - tab_h*1.5)} L {_n(w)} {_n(h-tab_h)} L {_n(w)} {_n(h)} '
        f'L {_n(w-tab_w)} {_n(h)} L {_n(w-tab_w)} {_n(h-tab_h)} L {_n(tab_w)} {_n(h-tab_h)} '
        f'L {_n(tab_w)} {_n(h)} L 0 {_n(h)} Z"/>'
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
#: than naming a colour.  Nothing at this layer knows what that fill is -- a generator
#: sees only width, height and adjustments -- so they are approximated by a neutral
#: overlay of comparable strength, which reads correctly over any base colour and is what
#: makes an action button's bevel visible at all.
_SHADE_ATTRIBUTES = {
    "none": 'fill="none"',
    "lighten": 'fill="#ffffff" fill-opacity="0.4"',
    "lightenLess": 'fill="#ffffff" fill-opacity="0.2"',
    "darken": 'fill="#000000" fill-opacity="0.4"',
    "darkenLess": 'fill="#000000" fill-opacity="0.2"',
}


class _P:
    """One ``a:path`` of a preset definition: its commands, and how it is painted."""

    __slots__ = ("commands", "fill", "stroke")

    def __init__(self, *commands, fill: str = "norm", stroke: bool = True):
        self.commands = commands
        self.fill = fill
        self.stroke = stroke

    def attributes(self) -> str:
        """Presentation attributes that override what the caller splices onto the
        wrapping ``<g>``.  A normally-painted path overrides nothing and inherits both."""
        parts = []
        shade = _SHADE_ATTRIBUTES.get(self.fill)
        if shade:
            parts.append(shade)
        if not self.stroke:
            parts.append('stroke="none"')
        return " ".join(parts)


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
        data = _spec_path_data(path.commands, variables)
        if not data:
            continue
        attributes = path.attributes()
        rendered.append(f'<path d="{data}"' + (f" {attributes}" if attributes else "") + "/>")

    if not rendered:
        return f'<rect width="{_n(w)}" height="{_n(h)}"/>'
    if len(rendered) == 1 and not spec.paths[0].attributes():
        return rendered[0]
    # More than one path, or one painted differently from the shape: wrap them so the
    # caller still has a single element to splice fill and stroke onto, and let the
    # children inherit it or override it.
    return f"<g>{''.join(rendered)}</g>"


def _spec_path_data(commands, variables: dict) -> str:
    """Build an SVG ``d`` string, tracking the pen so ``arcTo`` can be converted.

    DrawingML's ``arcTo`` is relative to wherever the pen already is -- it names a sweep,
    not an end point -- so the conversion needs the position the preceding command left.
    """

    def value(token) -> float:
        return resolve_value(token, variables)

    parts: list[str] = []
    x = y = start_x = start_y = 0.0

    for command in commands:
        kind = command[0]
        if kind in ("M", "L"):
            x, y = value(command[1]), value(command[2])
            parts.append(f"{kind} {_n(x)} {_n(y)}")
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
            x, y, large, sweep = arc_endpoint(
                x, y, width_radius, height_radius, start_angle, sweep_angle
            )
            parts.append(
                f"A {_n(width_radius)} {_n(height_radius)} 0 {large} {sweep} {_n(x)} {_n(y)}"
            )
        elif kind in ("Q", "C"):
            points = [
                (value(command[i]), value(command[i + 1])) for i in range(1, len(command), 2)
            ]
            parts.append(f"{kind} " + ", ".join(f"{_n(px)} {_n(py)}" for px, py in points))
            x, y = points[-1]

    return " ".join(parts)


def _spec_generator(name: str) -> Generator:
    """Adapt a spec entry to the ``(w, h, adj) -> str`` shape of every other generator."""

    def generate(w: float, h: float, adj: dict) -> str:
        return _spec_geometry(SPEC_PRESETS[name], w, h, adj)

    return generate


#: Preset definitions transcribed from ECMA-376 Appendix D (``presetShapeDefinitions.xml``).
#: Cross-checked against two independent copies -- a published dump of the spec file and
#: OnlyOffice's C++ transcription -- which agree byte for byte, including the oddities
#: (``accentCallout1`` really does close an empty subpath between its ``moveTo`` and its
#: ``lnTo``, and really does place its accent bar at ``x1`` rather than at the left edge).
SPEC_PRESETS: dict[str, "_Spec"] = {
    "callout1": _Spec(
        adjustments=(("adj1", 18750), ("adj2", -8333), ("adj3", 112500), ("adj4", -38333),),
        guides=(
            ("y1", "*/ h adj1 100000"),
            ("x1", "*/ w adj2 100000"),
            ("y2", "*/ h adj3 100000"),
            ("x2", "*/ w adj4 100000"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "r", "t"),
                ("L", "r", "b"),
                ("L", "l", "b"),
                ("Z",),
                stroke=False,
            ),
            _P(
                ("M", "x1", "y1"),
                ("L", "x2", "y2"),
                fill="none",
            ),
        ),
    ),
    "callout2": _Spec(
        adjustments=(
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 112500),
            ("adj6", -46667),
        ),
        guides=(
            ("y1", "*/ h adj1 100000"),
            ("x1", "*/ w adj2 100000"),
            ("y2", "*/ h adj3 100000"),
            ("x2", "*/ w adj4 100000"),
            ("y3", "*/ h adj5 100000"),
            ("x3", "*/ w adj6 100000"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "r", "t"),
                ("L", "r", "b"),
                ("L", "l", "b"),
                ("Z",),
                stroke=False,
            ),
            _P(
                ("M", "x1", "y1"),
                ("L", "x2", "y2"),
                ("L", "x3", "y3"),
                fill="none",
            ),
        ),
    ),
    "callout3": _Spec(
        adjustments=(
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
        guides=(
            ("y1", "*/ h adj1 100000"),
            ("x1", "*/ w adj2 100000"),
            ("y2", "*/ h adj3 100000"),
            ("x2", "*/ w adj4 100000"),
            ("y3", "*/ h adj5 100000"),
            ("x3", "*/ w adj6 100000"),
            ("y4", "*/ h adj7 100000"),
            ("x4", "*/ w adj8 100000"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "r", "t"),
                ("L", "r", "b"),
                ("L", "l", "b"),
                ("Z",),
                stroke=False,
            ),
            _P(
                ("M", "x1", "y1"),
                ("L", "x2", "y2"),
                ("L", "x3", "y3"),
                ("L", "x4", "y4"),
                fill="none",
            ),
        ),
    ),
    "accentCallout1": _Spec(
        adjustments=(("adj1", 18750), ("adj2", -8333), ("adj3", 112500), ("adj4", -38333),),
        guides=(
            ("y1", "*/ h adj1 100000"),
            ("x1", "*/ w adj2 100000"),
            ("y2", "*/ h adj3 100000"),
            ("x2", "*/ w adj4 100000"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "r", "t"),
                ("L", "r", "b"),
                ("L", "l", "b"),
                ("Z",),
                stroke=False,
            ),
            _P(
                ("M", "x1", "t"),
                ("Z",),
                ("L", "x1", "b"),
                fill="none",
            ),
            _P(
                ("M", "x1", "y1"),
                ("L", "x2", "y2"),
                fill="none",
            ),
        ),
    ),
    "accentCallout2": _Spec(
        adjustments=(
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 112500),
            ("adj6", -46667),
        ),
        guides=(
            ("y1", "*/ h adj1 100000"),
            ("x1", "*/ w adj2 100000"),
            ("y2", "*/ h adj3 100000"),
            ("x2", "*/ w adj4 100000"),
            ("y3", "*/ h adj5 100000"),
            ("x3", "*/ w adj6 100000"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "r", "t"),
                ("L", "r", "b"),
                ("L", "l", "b"),
                ("Z",),
                stroke=False,
            ),
            _P(
                ("M", "x1", "t"),
                ("Z",),
                ("L", "x1", "b"),
                fill="none",
            ),
            _P(
                ("M", "x1", "y1"),
                ("L", "x2", "y2"),
                ("L", "x3", "y3"),
                fill="none",
            ),
        ),
    ),
    "accentCallout3": _Spec(
        adjustments=(
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
        guides=(
            ("y1", "*/ h adj1 100000"),
            ("x1", "*/ w adj2 100000"),
            ("y2", "*/ h adj3 100000"),
            ("x2", "*/ w adj4 100000"),
            ("y3", "*/ h adj5 100000"),
            ("x3", "*/ w adj6 100000"),
            ("y4", "*/ h adj7 100000"),
            ("x4", "*/ w adj8 100000"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "r", "t"),
                ("L", "r", "b"),
                ("L", "l", "b"),
                ("Z",),
                stroke=False,
            ),
            _P(
                ("M", "x1", "t"),
                ("Z",),
                ("L", "x1", "b"),
                fill="none",
            ),
            _P(
                ("M", "x1", "y1"),
                ("L", "x2", "y2"),
                ("L", "x3", "y3"),
                ("L", "x4", "y4"),
                fill="none",
            ),
        ),
    ),
    "accentBorderCallout1": _Spec(
        adjustments=(("adj1", 18750), ("adj2", -8333), ("adj3", 112500), ("adj4", -38333),),
        guides=(
            ("y1", "*/ h adj1 100000"),
            ("x1", "*/ w adj2 100000"),
            ("y2", "*/ h adj3 100000"),
            ("x2", "*/ w adj4 100000"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "r", "t"),
                ("L", "r", "b"),
                ("L", "l", "b"),
                ("Z",),
            ),
            _P(
                ("M", "x1", "t"),
                ("Z",),
                ("L", "x1", "b"),
                fill="none",
            ),
            _P(
                ("M", "x1", "y1"),
                ("L", "x2", "y2"),
                fill="none",
            ),
        ),
    ),
    "accentBorderCallout2": _Spec(
        adjustments=(
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 112500),
            ("adj6", -46667),
        ),
        guides=(
            ("y1", "*/ h adj1 100000"),
            ("x1", "*/ w adj2 100000"),
            ("y2", "*/ h adj3 100000"),
            ("x2", "*/ w adj4 100000"),
            ("y3", "*/ h adj5 100000"),
            ("x3", "*/ w adj6 100000"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "r", "t"),
                ("L", "r", "b"),
                ("L", "l", "b"),
                ("Z",),
            ),
            _P(
                ("M", "x1", "t"),
                ("Z",),
                ("L", "x1", "b"),
                fill="none",
            ),
            _P(
                ("M", "x1", "y1"),
                ("L", "x2", "y2"),
                ("L", "x3", "y3"),
                fill="none",
            ),
        ),
    ),
    "accentBorderCallout3": _Spec(
        adjustments=(
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
        guides=(
            ("y1", "*/ h adj1 100000"),
            ("x1", "*/ w adj2 100000"),
            ("y2", "*/ h adj3 100000"),
            ("x2", "*/ w adj4 100000"),
            ("y3", "*/ h adj5 100000"),
            ("x3", "*/ w adj6 100000"),
            ("y4", "*/ h adj7 100000"),
            ("x4", "*/ w adj8 100000"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "r", "t"),
                ("L", "r", "b"),
                ("L", "l", "b"),
                ("Z",),
            ),
            _P(
                ("M", "x1", "t"),
                ("Z",),
                ("L", "x1", "b"),
                fill="none",
            ),
            _P(
                ("M", "x1", "y1"),
                ("L", "x2", "y2"),
                ("L", "x3", "y3"),
                ("L", "x4", "y4"),
                fill="none",
            ),
        ),
    ),
    "leftArrowCallout": _Spec(
        adjustments=(("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 64977),),
        guides=(
            ("maxAdj2", "*/ 50000 h ss"),
            ("a2", "pin 0 adj2 maxAdj2"),
            ("maxAdj1", "*/ a2 2 1"),
            ("a1", "pin 0 adj1 maxAdj1"),
            ("maxAdj3", "*/ 100000 w ss"),
            ("a3", "pin 0 adj3 maxAdj3"),
            ("q2", "*/ a3 ss w"),
            ("maxAdj4", "+- 100000 0 q2"),
            ("a4", "pin 0 adj4 maxAdj4"),
            ("dy1", "*/ ss a2 100000"),
            ("dy2", "*/ ss a1 200000"),
            ("y1", "+- vc 0 dy1"),
            ("y2", "+- vc 0 dy2"),
            ("y3", "+- vc dy2 0"),
            ("y4", "+- vc dy1 0"),
            ("x1", "*/ ss a3 100000"),
            ("dx2", "*/ w a4 100000"),
            ("x2", "+- r 0 dx2"),
            ("x3", "+/ x2 r 2"),
        ),
        paths=(
            _P(
                ("M", "l", "vc"),
                ("L", "x1", "y1"),
                ("L", "x1", "y2"),
                ("L", "x2", "y2"),
                ("L", "x2", "t"),
                ("L", "r", "t"),
                ("L", "r", "b"),
                ("L", "x2", "b"),
                ("L", "x2", "y3"),
                ("L", "x1", "y3"),
                ("L", "x1", "y4"),
                ("Z",),
            ),
        ),
    ),
    "rightArrowCallout": _Spec(
        adjustments=(("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 64977),),
        guides=(
            ("maxAdj2", "*/ 50000 h ss"),
            ("a2", "pin 0 adj2 maxAdj2"),
            ("maxAdj1", "*/ a2 2 1"),
            ("a1", "pin 0 adj1 maxAdj1"),
            ("maxAdj3", "*/ 100000 w ss"),
            ("a3", "pin 0 adj3 maxAdj3"),
            ("q2", "*/ a3 ss w"),
            ("maxAdj4", "+- 100000 0 q2"),
            ("a4", "pin 0 adj4 maxAdj4"),
            ("dy1", "*/ ss a2 100000"),
            ("dy2", "*/ ss a1 200000"),
            ("y1", "+- vc 0 dy1"),
            ("y2", "+- vc 0 dy2"),
            ("y3", "+- vc dy2 0"),
            ("y4", "+- vc dy1 0"),
            ("dx3", "*/ ss a3 100000"),
            ("x3", "+- r 0 dx3"),
            ("x2", "*/ w a4 100000"),
            ("x1", "*/ x2 1 2"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "x2", "t"),
                ("L", "x2", "y2"),
                ("L", "x3", "y2"),
                ("L", "x3", "y1"),
                ("L", "r", "vc"),
                ("L", "x3", "y4"),
                ("L", "x3", "y3"),
                ("L", "x2", "y3"),
                ("L", "x2", "b"),
                ("L", "l", "b"),
                ("Z",),
            ),
        ),
    ),
    "upArrowCallout": _Spec(
        adjustments=(("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 64977),),
        guides=(
            ("maxAdj2", "*/ 50000 w ss"),
            ("a2", "pin 0 adj2 maxAdj2"),
            ("maxAdj1", "*/ a2 2 1"),
            ("a1", "pin 0 adj1 maxAdj1"),
            ("maxAdj3", "*/ 100000 h ss"),
            ("a3", "pin 0 adj3 maxAdj3"),
            ("q2", "*/ a3 ss h"),
            ("maxAdj4", "+- 100000 0 q2"),
            ("a4", "pin 0 adj4 maxAdj4"),
            ("dx1", "*/ ss a2 100000"),
            ("dx2", "*/ ss a1 200000"),
            ("x1", "+- hc 0 dx1"),
            ("x2", "+- hc 0 dx2"),
            ("x3", "+- hc dx2 0"),
            ("x4", "+- hc dx1 0"),
            ("y1", "*/ ss a3 100000"),
            ("dy2", "*/ h a4 100000"),
            ("y2", "+- b 0 dy2"),
            ("y3", "+/ y2 b 2"),
        ),
        paths=(
            _P(
                ("M", "l", "y2"),
                ("L", "x2", "y2"),
                ("L", "x2", "y1"),
                ("L", "x1", "y1"),
                ("L", "hc", "t"),
                ("L", "x4", "y1"),
                ("L", "x3", "y1"),
                ("L", "x3", "y2"),
                ("L", "r", "y2"),
                ("L", "r", "b"),
                ("L", "l", "b"),
                ("Z",),
            ),
        ),
    ),
    "downArrowCallout": _Spec(
        adjustments=(("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 64977),),
        guides=(
            ("maxAdj2", "*/ 50000 w ss"),
            ("a2", "pin 0 adj2 maxAdj2"),
            ("maxAdj1", "*/ a2 2 1"),
            ("a1", "pin 0 adj1 maxAdj1"),
            ("maxAdj3", "*/ 100000 h ss"),
            ("a3", "pin 0 adj3 maxAdj3"),
            ("q2", "*/ a3 ss h"),
            ("maxAdj4", "+- 100000 0 q2"),
            ("a4", "pin 0 adj4 maxAdj4"),
            ("dx1", "*/ ss a2 100000"),
            ("dx2", "*/ ss a1 200000"),
            ("x1", "+- hc 0 dx1"),
            ("x2", "+- hc 0 dx2"),
            ("x3", "+- hc dx2 0"),
            ("x4", "+- hc dx1 0"),
            ("dy3", "*/ ss a3 100000"),
            ("y3", "+- b 0 dy3"),
            ("y2", "*/ h a4 100000"),
            ("y1", "*/ y2 1 2"),
        ),
        paths=(
            _P(
                ("M", "l", "t"),
                ("L", "r", "t"),
                ("L", "r", "y2"),
                ("L", "x3", "y2"),
                ("L", "x3", "y3"),
                ("L", "x4", "y3"),
                ("L", "hc", "b"),
                ("L", "x1", "y3"),
                ("L", "x2", "y3"),
                ("L", "x2", "y2"),
                ("L", "l", "y2"),
                ("Z",),
            ),
        ),
    ),
    "leftRightArrowCallout": _Spec(
        adjustments=(("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 48123),),
        guides=(
            ("maxAdj2", "*/ 50000 h ss"),
            ("a2", "pin 0 adj2 maxAdj2"),
            ("maxAdj1", "*/ a2 2 1"),
            ("a1", "pin 0 adj1 maxAdj1"),
            ("maxAdj3", "*/ 50000 w ss"),
            ("a3", "pin 0 adj3 maxAdj3"),
            ("q2", "*/ a3 ss wd2"),
            ("maxAdj4", "+- 100000 0 q2"),
            ("a4", "pin 0 adj4 maxAdj4"),
            ("dy1", "*/ ss a2 100000"),
            ("dy2", "*/ ss a1 200000"),
            ("y1", "+- vc 0 dy1"),
            ("y2", "+- vc 0 dy2"),
            ("y3", "+- vc dy2 0"),
            ("y4", "+- vc dy1 0"),
            ("x1", "*/ ss a3 100000"),
            ("x4", "+- r 0 x1"),
            ("dx2", "*/ w a4 200000"),
            ("x2", "+- hc 0 dx2"),
            ("x3", "+- hc dx2 0"),
        ),
        paths=(
            _P(
                ("M", "l", "vc"),
                ("L", "x1", "y1"),
                ("L", "x1", "y2"),
                ("L", "x2", "y2"),
                ("L", "x2", "t"),
                ("L", "x3", "t"),
                ("L", "x3", "y2"),
                ("L", "x4", "y2"),
                ("L", "x4", "y1"),
                ("L", "r", "vc"),
                ("L", "x4", "y4"),
                ("L", "x4", "y3"),
                ("L", "x3", "y3"),
                ("L", "x3", "b"),
                ("L", "x2", "b"),
                ("L", "x2", "y3"),
                ("L", "x1", "y3"),
                ("L", "x1", "y4"),
                ("Z",),
            ),
        ),
    ),
    "upDownArrowCallout": _Spec(
        adjustments=(("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 48123),),
        guides=(
            ("maxAdj2", "*/ 50000 w ss"),
            ("a2", "pin 0 adj2 maxAdj2"),
            ("maxAdj1", "*/ a2 2 1"),
            ("a1", "pin 0 adj1 maxAdj1"),
            ("maxAdj3", "*/ 50000 h ss"),
            ("a3", "pin 0 adj3 maxAdj3"),
            ("q2", "*/ a3 ss hd2"),
            ("maxAdj4", "+- 100000 0 q2"),
            ("a4", "pin 0 adj4 maxAdj4"),
            ("dx1", "*/ ss a2 100000"),
            ("dx2", "*/ ss a1 200000"),
            ("x1", "+- hc 0 dx1"),
            ("x2", "+- hc 0 dx2"),
            ("x3", "+- hc dx2 0"),
            ("x4", "+- hc dx1 0"),
            ("y1", "*/ ss a3 100000"),
            ("y4", "+- b 0 y1"),
            ("dy2", "*/ h a4 200000"),
            ("y2", "+- vc 0 dy2"),
            ("y3", "+- vc dy2 0"),
        ),
        paths=(
            _P(
                ("M", "l", "y2"),
                ("L", "x2", "y2"),
                ("L", "x2", "y1"),
                ("L", "x1", "y1"),
                ("L", "hc", "t"),
                ("L", "x4", "y1"),
                ("L", "x3", "y1"),
                ("L", "x3", "y2"),
                ("L", "r", "y2"),
                ("L", "r", "y3"),
                ("L", "x3", "y3"),
                ("L", "x3", "y4"),
                ("L", "x4", "y4"),
                ("L", "hc", "b"),
                ("L", "x1", "y4"),
                ("L", "x2", "y4"),
                ("L", "x2", "y3"),
                ("L", "l", "y3"),
                ("Z",),
            ),
        ),
    ),
    "quadArrowCallout": _Spec(
        adjustments=(("adj1", 18515), ("adj2", 18515), ("adj3", 18515), ("adj4", 48123),),
        guides=(
            ("a2", "pin 0 adj2 50000"),
            ("maxAdj1", "*/ a2 2 1"),
            ("a1", "pin 0 adj1 maxAdj1"),
            ("maxAdj3", "+- 50000 0 a2"),
            ("a3", "pin 0 adj3 maxAdj3"),
            ("q2", "*/ a3 2 1"),
            ("maxAdj4", "+- 100000 0 q2"),
            ("a4", "pin a1 adj4 maxAdj4"),
            ("dx2", "*/ ss a2 100000"),
            ("dx3", "*/ ss a1 200000"),
            ("ah", "*/ ss a3 100000"),
            ("dx1", "*/ w a4 200000"),
            ("dy1", "*/ h a4 200000"),
            ("x8", "+- r 0 ah"),
            ("x2", "+- hc 0 dx1"),
            ("x7", "+- hc dx1 0"),
            ("x3", "+- hc 0 dx2"),
            ("x6", "+- hc dx2 0"),
            ("x4", "+- hc 0 dx3"),
            ("x5", "+- hc dx3 0"),
            ("y8", "+- b 0 ah"),
            ("y2", "+- vc 0 dy1"),
            ("y7", "+- vc dy1 0"),
            ("y3", "+- vc 0 dx2"),
            ("y6", "+- vc dx2 0"),
            ("y4", "+- vc 0 dx3"),
            ("y5", "+- vc dx3 0"),
        ),
        paths=(
            _P(
                ("M", "l", "vc"),
                ("L", "ah", "y3"),
                ("L", "ah", "y4"),
                ("L", "x2", "y4"),
                ("L", "x2", "y2"),
                ("L", "x4", "y2"),
                ("L", "x4", "ah"),
                ("L", "x3", "ah"),
                ("L", "hc", "t"),
                ("L", "x6", "ah"),
                ("L", "x5", "ah"),
                ("L", "x5", "y2"),
                ("L", "x7", "y2"),
                ("L", "x7", "y4"),
                ("L", "x8", "y4"),
                ("L", "x8", "y3"),
                ("L", "r", "vc"),
                ("L", "x8", "y6"),
                ("L", "x8", "y5"),
                ("L", "x7", "y5"),
                ("L", "x7", "y7"),
                ("L", "x5", "y7"),
                ("L", "x5", "y8"),
                ("L", "x6", "y8"),
                ("L", "hc", "b"),
                ("L", "x3", "y8"),
                ("L", "x4", "y8"),
                ("L", "x4", "y7"),
                ("L", "x2", "y7"),
                ("L", "x2", "y5"),
                ("L", "ah", "y5"),
                ("L", "ah", "y6"),
                ("Z",),
            ),
        ),
    ),
}


PRESET_GEOMETRIES: dict[str, Generator] = {
    # Basic
    "rect": _rect,
    "ellipse": _ellipse,
    "roundRect": _round_rect,
    "triangle": _triangle,
    "rtTriangle": _rt_triangle,
    "diamond": _diamond,
    "parallelogram": _parallelogram,
    "trapezoid": _trapezoid,
    "pentagon": lambda w, h, adj: _regular_polygon(w, h, 5),
    "hexagon": _hexagon,
    "heptagon": lambda w, h, adj: _regular_polygon(w, h, 7),
    "octagon": lambda w, h, adj: _regular_polygon(w, h, 8),
    "decagon": lambda w, h, adj: _regular_polygon(w, h, 10),
    "dodecagon": lambda w, h, adj: _regular_polygon(w, h, 12),
    "line": _line,
    # Stars and seals
    "star4": _star4,
    "star5": lambda w, h, adj: _star_polygon(w, h, 5, 0.38),
    "star6": lambda w, h, adj: _star_polygon(w, h, 6, 0.5),
    "star7": lambda w, h, adj: _star_polygon(w, h, 7, 0.38),
    "star8": lambda w, h, adj: _star_polygon(w, h, 8, 0.38),
    "star10": lambda w, h, adj: _star_polygon(w, h, 10, 0.38),
    "star12": lambda w, h, adj: _star_polygon(w, h, 12, 0.38),
    "star16": lambda w, h, adj: _star_polygon(w, h, 16, 0.38),
    "star24": lambda w, h, adj: _star_polygon(w, h, 24, 0.38),
    "star32": lambda w, h, adj: _star_polygon(w, h, 32, 0.38),
    "irregularSeal1": _irregular_seal1,
    "irregularSeal2": _irregular_seal2,
    # Arrows
    "rightArrow": _right_arrow,
    "leftArrow": _left_arrow,
    "upArrow": _up_arrow,
    "downArrow": _down_arrow,
    "leftRightArrow": _left_right_arrow,
    "upDownArrow": _up_down_arrow,
    "notchedRightArrow": _notched_right_arrow,
    "stripedRightArrow": _striped_right_arrow,
    "chevron": _chevron,
    "homePlate": _home_plate,
    "leftRightUpArrow": _left_right_up_arrow,
    "quadArrow": _quad_arrow,
    "bentArrow": _bent_arrow,
    "bentUpArrow": _bend_up_arrow,
    "bendUpArrow": _bend_up_arrow,
    "leftUpArrow": _left_up_arrow,
    "uturnArrow": _uturn_arrow,
    # Connectors
    "straightConnector1": _straight_connector1,
    "bentConnector2": _bent_connector2,
    "bentConnector3": _bent_connector3,
    "bentConnector4": _bent_connector4,
    "bentConnector5": _bent_connector5,
    "curvedConnector2": _curved_connector2,
    "curvedConnector3": _curved_connector3,
    "curvedConnector4": _curved_connector4,
    "curvedConnector5": _curved_connector5,
    # Flowchart
    "flowChartProcess": _rect,
    "flowChartAlternateProcess": _flow_alternate_process,
    "flowChartDecision": _flow_decision,
    "flowChartInputOutput": _flow_input_output,
    "flowChartPredefinedProcess": _flow_predefined_process,
    "flowChartInternalStorage": _flow_internal_storage,
    "flowChartDocument": _flow_document,
    "flowChartMultidocument": _flow_multidocument,
    "flowChartTerminator": _flow_terminator,
    "flowChartPreparation": _flow_preparation,
    "flowChartManualInput": _flow_manual_input,
    "flowChartManualOperation": _flow_manual_operation,
    "flowChartConnector": _ellipse,
    "flowChartOffpageConnector": _flow_offpage_connector,
    "flowChartPunchedCard": _flow_punched_card,
    "flowChartPunchedTape": _flow_punched_tape,
    "flowChartCollate": _flow_collate,
    "flowChartSort": _flow_sort,
    "flowChartExtract": _flow_extract,
    "flowChartMerge": _flow_merge,
    "flowChartOnlineStorage": _flow_online_storage,
    "flowChartDelay": _flow_delay,
    "flowChartDisplay": _flow_display,
    "flowChartMagneticTape": _flow_magnetic_tape,
    "flowChartMagneticDisk": _flow_magnetic_disk,
    "flowChartMagneticDrum": _flow_magnetic_drum,
    "flowChartSummingJunction": _flow_summing_junction,
    "flowChartOr": _flow_or,
    # Callouts
    "wedgeRectCallout": _wedge_rect_callout,
    "wedgeRoundRectCallout": _wedge_round_rect_callout,
    "wedgeEllipseCallout": _wedge_ellipse_callout,
    "cloudCallout": _cloud_callout,
    "borderCallout1": _border_callout1,
    "borderCallout2": _border_callout2,
    "borderCallout3": _border_callout3,
    "callout1": _spec_generator("callout1"),
    "callout2": _spec_generator("callout2"),
    "callout3": _spec_generator("callout3"),
    "accentCallout1": _spec_generator("accentCallout1"),
    "accentCallout2": _spec_generator("accentCallout2"),
    "accentCallout3": _spec_generator("accentCallout3"),
    "accentBorderCallout1": _spec_generator("accentBorderCallout1"),
    "accentBorderCallout2": _spec_generator("accentBorderCallout2"),
    "accentBorderCallout3": _spec_generator("accentBorderCallout3"),
    "leftArrowCallout": _spec_generator("leftArrowCallout"),
    "rightArrowCallout": _spec_generator("rightArrowCallout"),
    "upArrowCallout": _spec_generator("upArrowCallout"),
    "downArrowCallout": _spec_generator("downArrowCallout"),
    "leftRightArrowCallout": _spec_generator("leftRightArrowCallout"),
    "upDownArrowCallout": _spec_generator("upDownArrowCallout"),
    "quadArrowCallout": _spec_generator("quadArrowCallout"),
    # Arcs
    "arc": _arc,
    "chord": _chord,
    "pie": _pie,
    "blockArc": _block_arc,
    # Math
    "mathPlus": _math_plus,
    "mathMinus": _math_minus,
    "mathMultiply": _math_multiply,
    "mathDivide": _math_divide,
    "mathEqual": _math_equal,
    "mathNotEqual": _math_not_equal,
    # Misc
    "cloud": _cloud,
    "heart": _heart,
    "plus": _plus,
    "corner": _corner,
    "diagStripe": _diag_stripe,
    "foldedCorner": _folded_corner,
    "plaque": _plaque,
    "can": _can,
    "cube": _cube,
    "donut": _donut,
    "noSmoking": _no_smoking,
    "smileyFace": _smiley_face,
    "frame": _frame,
    "bevel": _bevel,
    "halfFrame": _half_frame,
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
    "bracketPair": _bracket_pair,
    "bracePair": _brace_pair,
    "lightningBolt": _lightning_bolt,
    "moon": _moon,
    "teardrop": _teardrop,
    "sun": _sun,
    "wave": _wave,
    "doubleWave": _double_wave,
    "ribbon": _ribbon,
    "ribbon2": _ribbon2,
}


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
