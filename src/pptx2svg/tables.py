"""A table's rows as the renderer lays them out, for a caller that checks a slide without
drawing it.

``a:tr@h`` is a *minimum*.  PowerPoint makes a row taller when its text needs the space and
moves everything below it down; it never draws a row shorter than the stored height.  So a
table filled with wrapping text can run far past its frame -- and past the slide's bottom
edge, where PowerPoint cuts it off -- while every stored height still adds up to a frame
that fits.  :func:`table_row_heights` is the height of each row as the renderer draws it,
from the same text measurement it wraps with: the stored height, or what the tallest
single-row cell's text needs (its lines, paragraph spacing and the cell's top and bottom
margins) when that is more.  A cell spanning several rows does not grow any of them.
"""

from __future__ import annotations

from . import model as m
from .render.context import RenderContext

__all__ = ["table_row_heights"]


def table_row_heights(
    table: "m.TableElement | m.TableData", context: RenderContext | None = None
) -> list[float]:
    """Each row's height, EMU, as pptx2svg draws the table (see the module docstring).

    ``table`` is a resolved table -- an element of :func:`pptx2svg.convert_pptx_to_model`'s
    slides, or its ``table`` -- and ``context`` the measuring context to lay its text out
    with (a :class:`~pptx2svg.RenderContext` with the measurer the render would use; a
    default one when ``None``).

    For example::

        model = convert_pptx_to_model("deck.pptx", ConvertOptions(slide_numbers=[3]))
        table = next(e for e in model.slides[0].elements if isinstance(e, m.TableElement))
        heights = table_row_heights(table)          # [370840.0, 1005840.0, ...]
    """
    from .render.shape import _row_heights

    data = table.table if isinstance(table, m.TableElement) else table
    return list(_row_heights(data, context or RenderContext()))
