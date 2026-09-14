"""The DrawingML guide-formula language (ECMA-376 §20.1.9.11).

A shape's geometry is not a list of coordinates; it is a tiny prefix expression language.
``<a:gd name="dx2" fmla="*/ ss 3 8"/>`` means "dx2 is three eighths of the shortest
side", and a path then refers to ``dx2`` by name.  The same language appears in two
places, which is why it lives here rather than inside either of them:

* ``a:custGeom`` on a shape, evaluated by :mod:`pptx2svg.parse.drawing`;
* the preset shape definitions of ECMA-376 Appendix D, evaluated by
  :mod:`pptx2svg.render.geometry` for presets whose outline is too intricate to
  hand-approximate.

**Rounding.** ECMA-376 defines every guide as an integer, because Office evaluates them
in EMU -- a unit so small that rounding is invisible.  We evaluate custom geometry in its
own authored coordinate space, which is EMU-sized, so rounding there matches Office.
Preset geometry, though, is evaluated directly in *pixels*, where rounding to whole
numbers is up to a half-pixel error on every vertex.  Hence :func:`evaluate_formula`
takes a ``precise`` flag: off for custom geometry, which keeps its output byte-identical;
on for presets, where the rounding would show.
"""

from __future__ import annotations

import math

#: One degree, in the 1/60000-degree unit DrawingML uses for every angle.
DEGREE = 60000.0

#: Named angles.  ``cd`` is "circle divided": ``cd4`` is a quarter turn.
_ANGLE_CONSTANTS = {
    "cd8": 2700000.0,
    "cd4": 5400000.0,
    "cd2": 10800000.0,
    "3cd8": 8100000.0,
    "3cd4": 16200000.0,
    "5cd8": 13500000.0,
    "7cd8": 18900000.0,
}

#: Denominators for which the spec defines ``wd<n>`` / ``hd<n>`` / ``ssd<n>`` shorthands.
#: The list is not decorative: a guide referring to a name that is missing resolves to
#: zero rather than raising, so an incomplete set here shows up as a shape that renders
#: subtly wrong -- the worst kind of bug to find. Keep it complete.
_DIVISORS = (2, 3, 4, 5, 6, 8, 10, 12, 16, 20, 24, 32)


def builtin_variables(width: float, height: float) -> dict[str, float]:
    """The names every guide expression may use without declaring them."""
    shortest = min(width, height)
    variables: dict[str, float] = {
        "w": width,
        "h": height,
        "l": 0.0,
        "t": 0.0,
        "r": width,
        "b": height,
        # Centres: "horizontal center" and "vertical center".
        "hc": width / 2,
        "vc": height / 2,
        # Shortest and longest side -- how a preset keeps a feature square-ish when the
        # shape is stretched (an action button's symbol, a bevel's depth).
        "ss": shortest,
        "ls": max(width, height),
    }
    for divisor in _DIVISORS:
        variables[f"wd{divisor}"] = width / divisor
        variables[f"hd{divisor}"] = height / divisor
        variables[f"ssd{divisor}"] = shortest / divisor
    variables.update(_ANGLE_CONSTANTS)
    return variables


def evaluate_guides(
    guide_lists: list[list[tuple[str, str]]],
    width: float,
    height: float,
    *,
    precise: bool = False,
) -> dict[str, float]:
    """Evaluate guide lists, in order, into one namespace.

    Order matters and is not alphabetical: a ``gdLst`` guide may refer to an ``avLst``
    adjustment and to any guide declared before it, so each result must be visible to
    everything that follows.
    """
    variables = builtin_variables(width, height)
    for guides in guide_lists:
        for name, formula in guides:
            variables[name] = evaluate_formula(formula, variables, precise=precise)
    return variables


def evaluate_formula(
    formula: str, variables: dict[str, float], *, precise: bool = False
) -> float:
    """Evaluate one ``a:gd`` formula.

    The grammar is prefix: an operator name followed by up to three operands, each either
    a literal or the name of an already-evaluated guide.  An unknown operator yields 0
    rather than raising -- geometry is cosmetic, and a shape drawn wrong beats a deck
    that will not convert at all.
    """
    tokens = formula.strip().split()
    if not tokens:
        return 0.0
    operator = tokens[0]

    def operand(index: int) -> float:
        return resolve_value(tokens[index], variables) if index < len(tokens) else 0.0

    def quantize(value: float) -> float:
        return value if precise else round(value)

    if operator == "val":
        return operand(1)
    if operator == "+-":
        return operand(1) + operand(2) - operand(3)
    if operator == "*/":
        return quantize((operand(1) * operand(2)) / (operand(3) or 1))
    if operator == "+/":
        return quantize((operand(1) + operand(2)) / (operand(3) or 1))
    if operator == "pin":
        # Clamp.  `pin x y z` is y limited to [x, z] -- note that the value being
        # clamped is the *middle* operand, not the first.
        return max(operand(1), min(operand(2), operand(3)))
    if operator == "min":
        return min(operand(1), operand(2))
    if operator == "max":
        return max(operand(1), operand(2))
    if operator == "abs":
        return abs(operand(1))
    if operator == "sqrt":
        return quantize(math.sqrt(max(0.0, operand(1))))
    if operator == "sin":
        return quantize(operand(1) * math.sin(math.radians(operand(2) / DEGREE)))
    if operator == "cos":
        return quantize(operand(1) * math.cos(math.radians(operand(2) / DEGREE)))
    if operator == "tan":
        return quantize(operand(1) * math.tan(math.radians(operand(2) / DEGREE)))
    if operator == "at2":
        return quantize(math.degrees(math.atan2(operand(2), operand(1))) * DEGREE)
    if operator == "mod":
        return quantize(math.sqrt(operand(1) ** 2 + operand(2) ** 2 + operand(3) ** 2))
    if operator == "cat2":
        return quantize(operand(1) * math.cos(math.atan2(operand(3), operand(2))))
    if operator == "sat2":
        return quantize(operand(1) * math.sin(math.atan2(operand(3), operand(2))))
    if operator == "?:":
        return operand(2) if operand(1) > 0 else operand(3)
    return 0.0


def resolve_value(token: str, variables: dict[str, float]) -> float:
    try:
        return float(token)
    except ValueError:
        return variables.get(token, 0.0)


def _parametric_angle(geometric: float, width_radius: float, height_radius: float) -> float:
    """Geometric angle -> the ellipse's parametric angle, in radians.

    This is the subtlety that makes elliptical arcs hard.  DrawingML's ``stAng`` is a
    *geometric* angle: the direction of a ray from the ellipse centre, measured in the
    shape's own coordinates.  The parametric angle ``t`` that generates a point --
    ``(wR·cos t, hR·sin t)`` -- is a different number unless the ellipse is a circle,
    related by ``tan t = (wR/hR)·tan θ``.

    Treating one as the other is a silent error: it vanishes when ``wR == hR``, so
    circles and every arc in a square shape look right, and it grows with the aspect
    ratio until a stretched arc visibly fails to meet the segment it should join.

    The revolution is preserved: ``atan2`` returns a value in (-pi, pi], so a sweep that
    crosses the 12 o'clock mark would otherwise jump a full turn and flip the arc.
    """
    if width_radius == height_radius:
        return geometric
    t = math.atan2(width_radius * math.sin(geometric), height_radius * math.cos(geometric))
    return t + math.tau * round((geometric - t) / math.tau)


def arc_segments(
    current_x: float,
    current_y: float,
    width_radius: float,
    height_radius: float,
    start_angle: float,
    sweep_angle: float,
) -> list[tuple[float, float, int, int]]:
    """Convert a DrawingML ``a:arcTo`` into one or more SVG ``A`` commands.

    The two models disagree about what an arc *is*.  DrawingML gives a start angle and a
    sweep about an implied ellipse centre, with the current point already on that
    ellipse; SVG gives radii and an explicit end point and infers the centre.  So the
    centre is reconstructed from the current point and the start angle, and each end
    point is projected from it.  Both angles are geometric and must be converted to the
    ellipse's parametric angle first -- see :func:`_parametric_angle`.

    **Why more than one segment.**  SVG defines an arc whose endpoints coincide as
    drawing *nothing*, and DrawingML draws a circle as a single ``arcTo`` sweeping a full
    360 degrees -- which is exactly that case.  Taken literally, every circle in the
    preset catalogue disappears: the face of ``smileyFace``, its eyes, the middle of
    ``sun``, the hole in ``donut``.  So a sweep is cut into pieces of at most half a
    turn, which is representable, exact, and what every other converter does.

    Returns a list of ``(end_x, end_y, large_arc_flag, sweep_flag)``, in order.  Angles
    arrive in 1/60000 of a degree, and y grows downwards in both models, so no sign flip
    is needed.
    """
    start_geometric = math.radians(start_angle / DEGREE)
    end_geometric = math.radians((start_angle + sweep_angle) / DEGREE)
    start_parametric = _parametric_angle(start_geometric, width_radius, height_radius)
    end_parametric = _parametric_angle(end_geometric, width_radius, height_radius)

    center_x = current_x - width_radius * math.cos(start_parametric)
    center_y = current_y - height_radius * math.sin(start_parametric)

    total = end_parametric - start_parametric
    if total == 0:
        return []
    count = max(1, math.ceil(abs(total) / math.pi - 1e-9))
    step = total / count

    segments = []
    for index in range(1, count + 1):
        angle = start_parametric + step * index
        end_x = center_x + width_radius * math.cos(angle)
        end_y = center_y + height_radius * math.sin(angle)
        # Each piece is at most half a turn by construction, so large-arc is never set;
        # the direction still follows the sign of the sweep.
        segments.append((end_x, end_y, 1 if abs(step) > math.pi else 0, 1 if step > 0 else 0))
    return segments


def arc_endpoint(
    current_x: float,
    current_y: float,
    width_radius: float,
    height_radius: float,
    start_angle: float,
    sweep_angle: float,
) -> tuple[float, float, int, int]:
    """The final point of an ``a:arcTo``; see :func:`arc_segments` for the detail."""
    segments = arc_segments(
        current_x, current_y, width_radius, height_radius, start_angle, sweep_angle
    )
    if not segments:
        return current_x, current_y, 0, 0
    return segments[-1]
