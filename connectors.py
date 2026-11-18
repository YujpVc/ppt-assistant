"""Single-path connectors for native, editable mechanism diagrams.

A relationship is one polyline with one arrowhead at its end. Do not build a
bend from several straight connectors: the intermediate arrowheads read as
broken/extra edges at the corner.
"""
from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR  # noqa: F401  (kept for parity with builders)
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


def _rgb(value: str | RGBColor) -> RGBColor:
    if isinstance(value, RGBColor):
        return value
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def add_polyline(slide, points, *, color="#1F3A5F", width=1.2, dash=None,
                 arrow_end=True, arrow_start=False, name="connector:polyline"):
    """Draw one continuous orthogonal polyline.

    points: [(x_in, y_in), ...] in inches, at least two points.
    Only the last vertex gets an arrowhead; corners never do.
    """
    if len(points) < 2:
        raise ValueError("add_polyline needs at least two points")
    x0, y0 = points[0]
    builder = slide.shapes.build_freeform(Inches(x0), Inches(y0), scale=1.0)
    builder.add_line_segments([(Inches(x), Inches(y)) for x, y in points[1:]], close=False)
    shape = builder.convert_to_shape()
    shape.name = name
    shape.fill.background()
    shape.line.color.rgb = _rgb(color)
    shape.line.width = Pt(width)
    ln = shape.line._get_or_add_ln()
    if dash:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash}))
    if arrow_start:
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if arrow_end:
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    return shape
