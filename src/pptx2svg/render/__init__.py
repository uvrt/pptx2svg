"""SVG generation from the render model."""

from .context import RenderContext
from .geometry import preset_geometry_svg, render_geometry
from .svg import render_slide_to_svg

__all__ = ["RenderContext", "preset_geometry_svg", "render_geometry", "render_slide_to_svg"]
