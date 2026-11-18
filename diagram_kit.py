"""Native, editable flowchart primitives for custom PPT builders.

Everything here is a normal PowerPoint shape/connector, so the result stays
editable: nodes, ports, orthogonal edges, lanes, legends, notes, braces and
timeline rails. Build a relationship with one `edge()` call; never splice
straight connectors at a corner.
"""
from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from connectors import add_polyline

INK = "1F2937"
MUTED = "5D6877"
ACCENT = "1F3A5F"
ACCENT2 = "B58A55"
ACCENT3 = "6A7F99"
LINE = "D9E0E8"
PANEL = "FFFFFF"
TINT = "E8EDF3"
SUCCESS = "3F6B57"
WARNING = "B07D4F"
DANGER = "A33A2A"

_SHAPES = {
    "box": MSO_SHAPE.RECTANGLE,
    "round": MSO_SHAPE.ROUNDED_RECTANGLE,
    "pill": MSO_SHAPE.ROUNDED_RECTANGLE,
    "diamond": MSO_SHAPE.DIAMOND,
    "hexagon": MSO_SHAPE.HEXAGON,
    "circle": MSO_SHAPE.OVAL,
    "cylinder": MSO_SHAPE.CAN,
    "document": MSO_SHAPE.FLOWCHART_DOCUMENT,
    "start": MSO_SHAPE.ROUNDED_RECTANGLE,
    "end": MSO_SHAPE.ROUNDED_RECTANGLE,
    "subprocess": MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS,
    "manual": MSO_SHAPE.FLOWCHART_MANUAL_OPERATION,
    "delay": MSO_SHAPE.FLOWCHART_DELAY,
    "note": MSO_SHAPE.FOLDED_CORNER,
}


def _rgb(value: str | RGBColor) -> RGBColor:
    if isinstance(value, RGBColor):
        return value
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def node(slide, x, y, w, h, text="", *, kind="box", fill=PANEL, line=LINE,
         color=INK, size=10, bold=False, sub="", sub_size=8, radius=0.08,
         align=PP_ALIGN.CENTER, name=None):
    """One native node. `kind` is shape or a semantic role."""
    shape = slide.shapes.add_shape(_SHAPES.get(kind, MSO_SHAPE.RECTANGLE),
                                   Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = name or f"flow:{kind}"
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill)
    shape.line.color.rgb = _rgb(line)
    shape.line.width = Pt(1.0)
    if kind in ("round", "pill", "start", "end"):
        try:
            shape.adjustments[0] = min(0.5, radius / max(min(w, h), 0.01))
        except Exception:
            pass
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = 1.05
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold
    r.font.color.rgb = _rgb(color)
    r.font.name = "Noto Sans CJK SC"
    if sub:
        p2 = tf.add_paragraph(); p2.alignment = align
        r2 = p2.add_run(); r2.text = sub
        r2.font.size = Pt(sub_size); r2.font.color.rgb = _rgb(MUTED)
        r2.font.name = "Noto Sans CJK SC"
    return shape


def port(shape, side="right"):
    """Anchor in inches on a native shape: left/right/top/bottom/center."""
    x = shape.left / 914400; y = shape.top / 914400
    w = shape.width / 914400; h = shape.height / 914400
    return {
        "left": (x, y + h / 2), "right": (x + w, y + h / 2),
        "top": (x + w / 2, y), "bottom": (x + w / 2, y + h),
        "center": (x + w / 2, y + h / 2),
    }[side]


def _side_points(a, b, from_side, to_side):
    p1 = port(a, from_side); p2 = port(b, to_side)
    # one extra midpoint gives a clean orthogonal elbow
    if from_side in ("left", "right") and to_side in ("left", "right"):
        mx = (p1[0] + p2[0]) / 2
        return [p1, (mx, p1[1]), (mx, p2[1]), p2]
    if from_side in ("top", "bottom") and to_side in ("top", "bottom"):
        my = (p1[1] + p2[1]) / 2
        return [p1, (p1[0], my), (p2[0], my), p2]
    if from_side in ("left", "right"):
        return [p1, (p2[0], p1[1]), p2]
    return [p1, (p1[0], p2[1]), p2]


def edge(slide, a, b, *, from_side="right", to_side="left", color=ACCENT,
         width=1.2, dash=None, label="", label_size=8.5, label_dx=0,
         label_dy=-0.22, arrow_end=True, arrow_start=False, name="flow:edge"):
    """One continuous orthogonally routed relationship, one end arrowhead."""
    pts = _side_points(a, b, from_side, to_side)
    shape = add_polyline(slide, pts, color=color, width=width, dash=dash,
                         arrow_end=arrow_end, arrow_start=arrow_start, name=name)
    if label:
        mx = (pts[1][0] + pts[-2][0]) / 2
        my = (pts[1][1] + pts[-2][1]) / 2
        text(slide, label, mx + label_dx - 0.45, my + label_dy, 0.9, 0.18,
             label_size, MUTED, align=PP_ALIGN.CENTER)
    return shape


def text(slide, value, x, y, w, h, size=9, color=INK, *, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.15):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame; tf.clear(); tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]; p.alignment = align; p.line_spacing = line_spacing
    r = p.add_run(); r.text = value
    r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = _rgb(color)
    r.font.name = "Noto Sans CJK SC"
    return box


def badge(slide, x, y, w, value, *, fill=TINT, color=ACCENT, size=8, bold=True):
    return node(slide, x, y, w, 0.24, value, kind="pill", fill=fill, line=fill,
                color=color, size=size, bold=bold, name="flow:badge")


def lane(slide, x, y, w, h, title, *, fill=None, line=LINE, color=MUTED):
    """A labelled swimlane/container behind nodes."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
                                   Inches(w), Inches(h))
    shape.name = "flow:lane"
    shape.fill.solid(); shape.fill.fore_color.rgb = _rgb(fill or "F8FAFB")
    shape.line.color.rgb = _rgb(line); shape.line.width = Pt(0.8)
    text(slide, title, x + 0.12, y + 0.08, w - 0.24, 0.2, 8.5, color, bold=True)
    return shape


def legend(slide, x, y, items, *, color=ACCENT, size=8):
    """Small line samples + labels; items: [(label, color, dash), ...]."""
    for i, item in enumerate(items):
        lab, col, dash = (item + (None,))[:3]
        yy = y + i * 0.18
        add_polyline(slide, [(x, yy + 0.08), (x + 0.28, yy + 0.08)],
                     color=col or color, width=1.3, dash=dash, arrow_end=True,
                     name="flow:legend")
        text(slide, lab, x + 0.36, yy, 0.9, 0.16, size, MUTED)


def bracket(slide, x1, x2, y, label="", *, color=ACCENT3, drop=0.12, size=8):
    """A horizontal grouping bracket over equal-weight nodes."""
    pts = [(x1, y - drop), (x1, y), (x2, y), (x2, y - drop)]
    add_polyline(slide, pts, color=color, width=1.0, arrow_end=False, name="flow:bracket")
    if label:
        text(slide, label, (x1 + x2) / 2 - 0.9, y - drop - 0.22, 1.8, 0.18,
             size, MUTED, align=PP_ALIGN.CENTER)


def note(slide, x, y, w, h, value, *, fill="F3EBDD", line=ACCENT2, size=8):
    shape = slide.shapes.add_shape(MSO_SHAPE.FOLDED_CORNER, Inches(x),
                                               Inches(y), Inches(w), Inches(h))
    shape.name = "flow:note"
    shape.fill.solid(); shape.fill.fore_color.rgb = _rgb(fill)
    shape.line.color.rgb = _rgb(line); shape.line.width = Pt(0.8)
    text(slide, value, x + 0.1, y + 0.08, w - 0.2, h - 0.16, size, INK)
    return shape


def axis(slide, x1, x2, y, ticks, *, color=ACCENT3, size=8):
    """Timeline rail with labelled ticks; ticks: [(x, label), ...]."""
    add_polyline(slide, [(x1, y), (x2, y)], color=color, width=1.2,
                 arrow_end=True, name="flow:axis")
    for tx, lab in ticks:
        add_polyline(slide, [(tx, y - 0.07), (tx, y + 0.07)], color=color,
                     width=1.0, arrow_end=False, name="flow:tick")
        text(slide, lab, tx - 0.45, y + 0.1, 0.9, 0.18, size, MUTED,
             align=PP_ALIGN.CENTER)
