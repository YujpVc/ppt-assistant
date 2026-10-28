#!/usr/bin/env python3
"""Build an editable deck from JSON and explicit assets, independent of cwd.

Relative image paths resolve beside the specification. Layout and rendering
code live in this package; workspace scripts are never imported or executed.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageOps
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from charts import draw_chart, resolve_chart

from grasp_evidence import (candidate_projections, depth_colors, depth_pair,
                            gated_candidate_projections, gripper_projection)

LAYOUTS = {"cover", "cards", "process", "comparison", "table", "chart",
           "image_split", "image_grid", "chart_dashboard", "closing"}
THEMES = {
    "midnight": {"bg": "0B1324", "panel": "17243B", "text": "F7FBFF",
                 "muted": "B2C2D9", "accent": "48CAE4"},
    "paper": {"bg": "F4F6FA", "panel": "FFFFFF", "text": "15253F",
              "muted": "4D627A", "accent": "176C88", "danger": "C0392B"},
}


def validate(spec: dict, base: Path) -> None:
    """Reject unsupported layouts, missing assets and content beyond capacity."""
    if not isinstance(spec, dict) or not isinstance(spec.get("meta"), dict):
        raise ValueError("meta must be an object")
    if not spec["meta"].get("title"):
        raise ValueError("meta.title is required")
    if spec["meta"].get("theme", "midnight") not in THEMES:
        raise ValueError("meta.theme must be midnight or paper")
    slides = spec.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("slides must be a non-empty array")
    for index, slide in enumerate(slides, 1):
        if not isinstance(slide, dict) or slide.get("layout") not in LAYOUTS:
            raise ValueError(f"slide {index}: unsupported layout")
        layout = slide["layout"]
        if not slide.get("title") and layout not in {"cover", "closing"}:
            raise ValueError(f"slide {index}: title is required")
        if layout in {"cards", "process", "comparison", "image_grid"}:
            items = slide.get("items")
            minimum, maximum = (2, 2) if layout == "comparison" else (2, 4)
            if not isinstance(items, list) or not minimum <= len(items) <= maximum:
                raise ValueError(f"slide {index}: items requires {minimum}–{maximum} entries")
            if any(not isinstance(item, dict) or not item.get("title") for item in items):
                raise ValueError(f"slide {index}: each item requires a title")
        assets = [slide] if layout in {"cover", "image_split"} else []
        if layout == "image_grid":
            assets = slide["items"]
        for asset in assets:
            image = asset.get("image")
            if not image and layout != "cover":
                raise ValueError(f"slide {index}: image is required")
            if image:
                path = (base / image).resolve()
                if not path.is_file():
                    raise ValueError(f"slide {index}: image not found: {path}")
                with Image.open(path) as picture:
                    picture.verify()
        if layout == "table":
            headers, rows = slide.get("headers"), slide.get("rows")
            if (not isinstance(headers, list) or not 2 <= len(headers) <= 5
                    or not isinstance(rows, list) or not 1 <= len(rows) <= 6
                    or any(not isinstance(row, list) or len(row) != len(headers) for row in rows)):
                raise ValueError(f"slide {index}: table requires 2–5 columns and 1–6 complete rows")
        if layout in {"chart", "chart_dashboard"}:
            validate_chart_page(slide, base)


def validate_chart_page(page: dict, base: Path) -> None:
    """Validate chart payloads before the output can be replaced."""
    if page["layout"] == "chart":
        charts = [page.get("chart", {})]
    else:
        charts = page.get("charts")
        if not isinstance(charts, list) or not 2 <= len(charts) <= 4:
            raise ValueError("chart_dashboard requires 2–4 chart objects")
    for chart in charts:
        resolve_chart(chart, base)
        for field in ("title", "unit", "source", "takeaway"):
            if field in chart and not isinstance(chart[field], str):
                raise ValueError(f"chart.{field} must be a string")
    if "takeaway" in page and not isinstance(page["takeaway"], str):
        raise ValueError("takeaway must be a string")


class DeckBuilder:
    """Own presentation state and layout primitives for one specification."""

    def __init__(self, spec: dict, base: Path):
        self.meta = spec["meta"]
        self.base = base
        self.colors = THEMES[self.meta.get("theme", "midnight")]
        self.font = self.meta.get("font", "Noto Sans CJK SC")
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(13.333333), Inches(7.5)

    def box(self, slide, x, y, w, h, color="panel", rounded=False):
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
            Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        color_value = self.colors.get(color, color)
        shape.fill.fore_color.rgb = RGBColor.from_string(color_value)
        shape.line.fill.background()
        return shape

    def text(self, slide, value, x, y, w, h, size=21, color="text", bold=False):
        """Fit text conservatively; reject dense content instead of shrinking it."""
        value = str(value)
        if not value.strip():
            return None
        lines = sum(max(1, math.ceil(sum(1 if ord(c) > 0x2E80 else .55 for c in line)
                                    / (w * 72 / size))) for line in value.split("\n"))
        if lines * size * 1.25 + max(0, len(value.split("\n")) - 1) * 8 > h * 72:
            raise ValueError(f"text exceeds layout capacity; split or shorten: {value[:60]}")
        shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        frame = shape.text_frame
        frame.word_wrap = True
        frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
        for index, line in enumerate(value.split("\n")):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            paragraph.space_after = Pt(8)
            paragraph.line_spacing = 1.2
            run = paragraph.add_run()
            run.text = line
            run.font.name, run.font.size, run.font.bold = self.font, Pt(size), bold
            color_value = self.colors.get(color, color)
            run.font.color.rgb = RGBColor.from_string(color_value)
            east_asian = run._r.get_or_add_rPr().makeelement(qn("a:ea"), {"typeface": self.font})
            run._r.get_or_add_rPr().append(east_asian)
        return shape

    def picture(self, slide, filename, x, y, w, h, crop=None):
        with Image.open(self.base / filename) as source:
            picture = ImageOps.exif_transpose(source).convert("RGB")
            if crop is not None:
                if (len(crop) != 4 or any(type(v) is not int for v in crop)
                        or not 0 <= crop[0] < crop[2] <= picture.width
                        or not 0 <= crop[1] < crop[3] <= picture.height):
                    raise ValueError(f"invalid pixel crop for image: {filename}")
                picture = picture.crop(tuple(crop))
            return self._picture_image(slide, picture, x, y, w, h)

    def _picture_image(self, slide, picture, x, y, w, h):
        payload = io.BytesIO()
        picture.save(payload, "PNG")
        payload.seek(0)
        ratio = min(w / picture.width, h / picture.height)
        width, height = picture.width * ratio, picture.height * ratio
        left, top = x + (w - width) / 2, y + (h - height) / 2
        slide.shapes.add_picture(payload, Inches(left), Inches(top), Inches(width), Inches(height))
        return left, top, ratio

    def new(self, page, index):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(self.colors["bg"])
        self.text(slide, page.get("title", self.meta["title"]), .6, .38, 12.1, .95, 30, bold=True)
        if page.get("subtitle"):
            self.text(slide, page["subtitle"], .6, 1.35, 12.1, .65, 17, "muted")
        self.text(slide, self.meta["title"], .6, 7.06, 10.8, .24, 10, "muted")
        self.text(slide, f"{index:02}", 12.0, 7.02, .7, .3, 12, "muted")
        slide.notes_slide.notes_text_frame.text = page.get("notes", "")
        return slide

    def card(self, slide, item, x, y, w, h, number=None):
        self.box(slide, x, y, w, h, rounded=True)
        offset = .25
        if number is not None:
            self.text(slide, f"{number:02}", x + .25, y + .25, w - .5, .55, 28, "accent", True)
            offset = 1.0
        self.text(slide, item["title"], x + .25, y + offset, w - .5, .85, 23, "accent", True)
        body_height = h - offset - 1.25
        # Four-card pages intentionally use compact cards.  Keep their body
        # copy readable while leaving enough vertical room for two wrapped
        # Chinese lines instead of allowing the strict capacity check to fail.
        body_size = 15 if body_height < 1.0 else 20
        self.text(slide, item.get("body", ""), x + .25, y + offset + 1.0,
                  w - .5, body_height, body_size)

    def compose(self, slide, page):
        layout = page["layout"]
        if layout in {"cover", "closing"}:
            image = page.get("image")
            width = 7.1 if image else 11.7
            self.text(slide, page.get("body", self.meta.get("subtitle", "")),
                      .8, 2.35, width, 2.5, 32, "accent", True)
            self.text(slide, self.meta.get("author", ""), .8, 5.5, width, .7, 20, "muted")
            if image:
                self.picture(slide, image, 8.3, 1.95, 4.3, 4.75)
        elif layout in {"cards", "comparison"}:
            items = page["items"]
            columns = len(items) if len(items) <= 3 else 2
            rows = math.ceil(len(items) / columns)
            width = (12.1 - .25 * (columns - 1)) / columns
            height = (4.55 - .25 * (rows - 1)) / rows
            for index, item in enumerate(items):
                self.card(slide, item, .6 + (index % columns) * (width + .25),
                          2.05 + (index // columns) * (height + .25), width, height)
        elif layout == "process":
            items = page["items"]
            width = (12.1 - .4 * (len(items) - 1)) / len(items)
            for index, item in enumerate(items):
                x = .6 + index * (width + .4)
                self.card(slide, item, x, 2.25, width, 4.1, index + 1)
                if index < len(items) - 1:
                    shape = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x + width + .07),
                                                  Inches(3.9), Inches(.25), Inches(.45))
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = RGBColor.from_string(self.colors["accent"])
                    shape.line.fill.background()
        elif layout == "image_split":
            self.picture(slide, page["image"], .6, 2.05, 7.1, 4.55)
            self.box(slide, 8.0, 2.05, 4.7, 4.55, rounded=True)
            self.text(slide, page.get("body", ""), 8.3, 2.4, 4.1, 3.8, 22)
        elif layout == "image_grid":
            items = page["items"]
            rows = math.ceil(len(items) / 2)
            height = (4.55 - .2 * (rows - 1)) / rows
            for index, item in enumerate(items):
                x, y = .6 + (index % 2) * 6.2, 2.05 + (index // 2) * (height + .2)
                self.picture(slide, item["image"], x, y, 5.9, height - .6)
                self.text(slide, item["title"], x, y + height - .48, 5.9, .45, 18, "accent", True)
        elif layout in {"chart", "chart_dashboard"}:
            top = 2.05 if page.get("subtitle") else 1.65
            self.chart_page(slide, page, top=top, height=6.6 - top)
        elif layout == "table":
            headers, rows = page["headers"], page["rows"]
            table = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(.6), Inches(2.1),
                                           Inches(12.1), Inches(4.4)).table
            for row_index, row in enumerate([headers] + rows):
                for column_index, value in enumerate(row):
                    cell = table.cell(row_index, column_index)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string(self.colors["accent" if row_index == 0 else "panel"])
                    cell.text = str(value)
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.name = self.font
                        paragraph.font.size = Pt(18)
                        paragraph.font.bold = row_index == 0
                        paragraph.font.color.rgb = RGBColor.from_string(self.colors["bg" if row_index == 0 else "text"])
                        for run in paragraph.runs:
                            run._r.get_or_add_rPr().append(run._r.makeelement(qn("a:ea"), {"typeface": self.font}))
        if page.get("source"):
            self.text(slide, page["source"], .6, 6.7, 12.1, .27, 11, "muted")

    def chart_page(self, slide, page, top=1.5, height=5.15):
        """Compose one chart or 2–4 coordinated charts with native annotations."""
        validate_chart_page(page, self.base)
        charts = [page["chart"]] if page["layout"] == "chart" else page["charts"]
        count = len(charts)
        columns = 1 if count == 1 else 2
        rows = math.ceil(count / columns)
        panel_w = (12.1 - .25 * (columns - 1)) / columns
        has_summary = bool(page.get("takeaway"))
        available_h = height - (.55 if has_summary else 0)
        panel_h = (available_h - .2 * (rows - 1)) / rows
        for index, spec in enumerate(charts):
            x = .6 + (index % columns) * (panel_w + .25)
            y = top + (index // columns) * (panel_h + .2)
            # A third chart spans the bottom row rather than leaving a blank cell.
            width = 12.1 if count == 3 and index == 2 else panel_w
            self.box(slide, x, y, width, panel_h)
            heading = spec.get("title", "")
            if spec.get("unit"):
                heading += (" · " if heading else "") + spec["unit"]
            heading_h = .36 if heading else 0
            if heading:
                self.text(slide, heading, x + .2, y + .12, width - .4,
                          heading_h, 13 if rows > 1 else 17, "text", True)
            tail = spec.get("takeaway", "")
            source = spec.get("source", "")
            tail_h = (.33 if tail else 0) + (.24 if source else 0)
            chart_y = y + .12 + heading_h
            chart_h = panel_h - .24 - heading_h - tail_h
            draw_chart(slide, spec, self.base, (x + .12, chart_y, width - .24, chart_h),
                       self.font, self.colors)
            cursor = chart_y + chart_h
            if tail:
                self.text(slide, tail, x + .2, cursor, width - .4, .33,
                          11 if rows > 1 else 13, "accent", True)
                cursor += .33
            if source:
                self.text(slide, source, x + .2, cursor, width - .4, .24, 9, "muted")
        if has_summary:
            self.text(slide, page["takeaway"], .8, top + height - .4,
                      11.7, .4, 16, "accent", True)

    # The probation v2 deck uses native PowerPoint shapes for its explanatory
    # pages.  These primitives deliberately keep the template's restrained
    # paper palette while giving each page a different composition.
    def _probation_bg(self, slide):
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string("F4F7FB")

    def _probation_header(self, slide, title, subtitle, index):
        self._probation_bg(slide)
        self.text(slide, title, .62, .36, 10.8, .45, 25, "text", True)
        self.box(slide, .62, 1.0, 1.15, .06, "accent")
        if subtitle:
            self.text(slide, subtitle, 1.92, .85, 10.7, .28, 10.5, "muted")
        self.text(slide, "人形机器人（上海）有限公司  |  轮臂-应用开发  |  赵文辉",
                  .55, 7.12, 10.5, .18, 8.5, "muted")
        self.text(slide, f"{index:02d}", 12.25, 7.08, .45, .22, 9, "muted")

    def _probation_panel(self, slide, x, y, w, h, rounded=True):
        return self.box(slide, x, y, w, h, "panel", rounded=rounded)

    def _probation_image(self, slide, filename, x, y, w, h, used_assets, crop=None):
        path = (self.base / filename).resolve()
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if digest in used_assets:
            raise ValueError(f"probation deck reuses image asset: {path}")
        used_assets.add(digest)
        self._probation_panel(slide, x, y, w, h, rounded=True)
        if path.suffix.lower() == ".gif" and crop is None:
            self._direct_picture(slide, path, x + .12, y + .12, w - .24, h - .24)
        else:
            self.picture(slide, filename, x + .12, y + .12, w - .24, h - .24, crop=crop)

    def _direct_picture(self, slide, path, x, y, w, h):
        """Insert the original media bytes so animated GIFs remain animated."""
        with Image.open(path) as source:
            width, height = source.size
        ratio = min(w / width, h / height)
        display_w, display_h = width * ratio, height * ratio
        left, top = x + (w - display_w) / 2, y + (h - display_h) / 2
        slide.shapes.add_picture(str(path), Inches(left), Inches(top),
                                 Inches(display_w), Inches(display_h))

    def _probation_arrow(self, slide, x, y, w=.28, h=.38):
        shape = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(self.colors["accent"])
        shape.line.fill.background()

    def _layered_photo(self, slide, item, x, y, w, h, used_assets):
        """Place a photo and aligned RGBA layers as separate editable pictures."""
        image_path = self.base / item["image"]
        digest = hashlib.sha256(image_path.read_bytes()).hexdigest()
        if digest in used_assets:
            raise ValueError(f"probation deck reuses image asset: {image_path}")
        used_assets.add(digest)
        with Image.open(image_path) as source:
            photo = ImageOps.exif_transpose(source).convert("RGB")
            source_size = photo.size
            crop = item["crop"]
            if (len(crop) != 4 or any(type(v) is not int for v in crop)
                    or not 0 <= crop[0] < crop[2] <= photo.width
                    or not 0 <= crop[1] < crop[3] <= photo.height):
                raise ValueError("invalid layered photo crop")
            photo = photo.crop(tuple(crop))
        left, top, ratio = self._picture_image(slide, photo, x, y, w, h)
        for filename in item["layers"]:
            with Image.open(self.base / filename) as source:
                if source.size != source_size or source.mode != "RGBA":
                    raise ValueError("overlay must be an RGBA layer matching the source photo dimensions")
                layer = source.crop(tuple(crop))
                payload = io.BytesIO()
                layer.save(payload, "PNG")
                payload.seek(0)
                slide.shapes.add_picture(payload, Inches(left), Inches(top),
                    Inches(layer.width * ratio), Inches(layer.height * ratio))

    def _probation_tag(self, slide, value, x, y, w, color="accent"):
        self.box(slide, x, y, w, .28, color, rounded=True)
        self.text(slide, value, x + .08, y + .035, w - .16, .18, 9.5,
                  "bg" if color == "accent" else "text", True)

    def _pose_photo(self, slide, frame, projections, x, y, w, h):
        """Draw recorded candidates and selected poses using one native style."""
        crop = frame["crop"]
        with Image.open(self.base / frame["image"]) as source:
            original = ImageOps.exif_transpose(source).convert("RGB")
            if not (0 <= crop[0] < crop[2] <= original.width
                    and 0 <= crop[1] < crop[3] <= original.height):
                raise ValueError("gripper photo crop exceeds source dimensions")
            picture = original.crop(tuple(crop))
        # A full candidate set can contain dozens of poses.  Rasterize only
        # that dense overlay in memory so the deck remains responsive while
        # preserving every candidate and the same geometry/colors.
        if len(projections) > 8:
            composite = picture.convert("RGBA")
            draw = ImageDraw.Draw(composite)
            for projection in projections:
                color = (23, 108, 136, 235) if projection.get("status") != "rejected" else (192, 57, 43, 235)
                for start, end in projection["segments"]:
                    points = [(round(start[0] - crop[0]), round(start[1] - crop[1])),
                              (round(end[0] - crop[0]), round(end[1] - crop[1]))]
                    draw.line(points, fill=color, width=max(2, round(picture.width / 180)))
                cx, cy = projection["center_uv"]
                radius = max(2, round(picture.width / 180))
                draw.ellipse((round(cx - crop[0] - radius), round(cy - crop[1] - radius),
                              round(cx - crop[0] + radius), round(cy - crop[1] + radius)), fill=color)
            self._picture_image(slide, composite.convert("RGB"), x, y, w, h)
            return
        left, top, ratio = self._picture_image(slide, picture, x, y, w, h)
        def to_slide(point):
            u, v = point
            if not crop[0] <= u <= crop[2] or not crop[1] <= v <= crop[3]:
                raise ValueError("gripper projection lies outside comparison crop")
            return left + (u - crop[0]) * ratio, top + (v - crop[1]) * ratio
        for projection in projections:
            color = self.colors["accent"] if projection.get("status") != "rejected" else self.colors["danger"]
            for start, end in projection["segments"]:
                sx, sy = to_slide(start)
                ex, ey = to_slide(end)
                connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                    Inches(sx), Inches(sy), Inches(ex), Inches(ey))
                connector.line.color.rgb = RGBColor.from_string(color)
                connector.line.width = Pt(1.25)
            cx, cy = to_slide(projection["center_uv"])
            center = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                Inches(cx - .027), Inches(cy - .027), Inches(.054), Inches(.054))
            center.fill.solid()
            center.fill.fore_color.rgb = RGBColor.from_string(color)
            center.line.fill.background()

    def _grasp_panel(self, slide, frame, phase, x, y, w, h):
        projection = gripper_projection(self.base / frame[phase], self.base / frame["calibration"])
        expected = frame.get(f"{phase}_opening_mm")
        if expected is not None and abs(projection["opening_mm"] - expected) > .02:
            raise ValueError("recorded gripper opening does not match comparison statistics")
        self._pose_photo(slide, frame, [projection], x, y, w, h)
        self.text(slide, f"所选开合 {projection['opening_mm']:.2f} mm", x, y + h + .08,
                  w, .26, 13, "text", True)

    def _probation_card(self, slide, title, body, x, y, w, h, number=None,
                        title_color="accent", body_size=15):
        self._probation_panel(slide, x, y, w, h, rounded=True)
        if h < 1.0:
            self.text(slide, title, x + .22, y + .18, 2.35, .35, 16, title_color, True)
            self.text(slide, body, x + 2.8, y + .18, w - 3.05, h - .3, body_size)
            return
        if number is not None and w < 3.5:
            self.text(slide, f"{number:02d}", x + .22, y + .15, .55, .38, 21, "accent", True)
            self.text(slide, title, x + .22, y + .68, w - .44, .42, 17, title_color, True)
            self.text(slide, body, x + .22, y + 1.28, w - .44, h - 1.45, body_size)
            return
        if number is not None:
            self.text(slide, f"{number:02d}", x + .22, y + .18, .55, .38, 21,
                      "accent", True)
            title_x, title_y = x + .92, y + .2
            title_w = w - 1.15
        else:
            title_x, title_y = x + .22, y + .2
            title_w = w - .44
        self.text(slide, title, title_x, title_y, title_w, .42, 18,
                  title_color, True)
        self.text(slide, body, x + .22, y + .78, w - .44, h - .96,
                  body_size, "text")

    def compose_probation_v2(self, slide, page, index, used_assets):
        """Compose a native, varied page for the v2 probation presentation."""
        layout = page["layout"]
        if layout == "thanks":
            # Match the original cover's composition at the end: artwork on
            # the right and a quiet, deliberately typeset thank-you on the left.
            self.text(slide, page["headline"], .78, 2.15, 6.2, .85, 40, "accent", True)
            self.box(slide, .82, 3.25, 1.1, .055, "accent")
            self.text(slide, page["body"], .82, 3.72, 6.2, 1.2, 20, "text")
            self.text(slide, page["footer"], .82, 5.8, 6.2, .45, 17, "muted")
            return
        self._probation_header(slide, page["title"], page.get("subtitle", ""), index)

        if layout in {"chart", "chart_dashboard"}:
            self.chart_page(slide, page)
        elif layout == "section_divider":
            # Section pages reuse the directory's restrained typography but
            # give the audience a clear reset before a dense block of work.
            self.text(slide, "SECTION", .84, 1.78, 2.0, .32, 14, "muted", True)
            self.text(slide, page.get("section_no", "01"), .82, 2.23, 1.15, .92,
                      46, "accent", True)
            self.box(slide, 2.28, 2.2, .08, 1.35, "accent")
            self.text(slide, page["section_title"], 2.7, 2.23, 7.6, .58,
                      31, "text", True)
            self.text(slide, page.get("section_subtitle", ""), 2.72, 2.92,
                      8.8, .4, 17, "muted")
            self.box(slide, 2.72, 3.62, 9.55, .012, "muted")
            for i, item in enumerate(page.get("items", [])):
                y = 4.0 + i * .58
                self.text(slide, f"{i + 1:02d}", 2.72, y, .55, .27,
                          13, "accent", True)
                self.text(slide, item, 3.5, y, 8.4, .3, 16, "text")
            self.text(slide, page.get("transition", "下面进入本节内容"),
                      2.72, 6.35, 4.5, .28, 13, "muted")
        elif layout == "directory":
            entries = page["entries"]
            self.text(slide, "CONTENTS", .82, 1.7, 2.4, .45, 22, "muted", True)
            for i, item in enumerate(entries):
                y = 1.72 + i * 1.12
                self.text(slide, f"{i + 1:02d}", 4.25, y, .8, .6, 30, "accent", True)
                self.text(slide, item["title"], 5.35, y, 6.7, .46, 23, "text", True)
                self.text(slide, item["body"], 5.35, y + .54, 6.7, .3, 14, "muted")
                if i < 3:
                    self.box(slide, 4.25, y + .98, 8.05, .01, "muted")
        elif layout == "overview":
            items = page["items"]
            self.text(slide, "工作方向", .9, 1.7, 2.5, .36, 17, "muted", True)
            self.text(slide, "主要负责内容", 4.05, 1.7, 7.6, .36, 17, "muted", True)
            for i, item in enumerate(items):
                y = 2.27 + i * 1.35
                self.box(slide, .8, y, 2.65, 1.1, "accent")
                self.text(slide, item["title"], 1.06, y + .32, 2.1, .43, 22, "bg", True)
                self.text(slide, item["body"].replace("\n", "；"), 4.05, y + .21, 8.0, .76, 18)
                self.box(slide, 4.05, y + 1.13, 8.15, .012, "muted")
        elif layout == "timeline":
            self.box(slide, .95, 3.35, 11.35, .06, "accent")
            for i, item in enumerate(page["items"]):
                x = 1.0 + i * 3.9
                self.box(slide, x, 3.18, .38, .38, "accent", rounded=True)
                self._probation_card(slide, item["title"], item["body"], x - .32, 1.62,
                                     3.35, 1.35, body_size=13)
                self._probation_card(slide, item["detail_title"], item["detail"], x - .32,
                                     3.72, 3.35, 1.72, body_size=13)
        elif layout == "pipeline":
            items = page["items"]
            width = 2.2
            for i, item in enumerate(items):
                x = .55 + i * 2.55
                self._probation_card(slide, item["title"], item["body"], x, 1.7, width, 2.2,
                                     i + 1, body_size=12.5)
                if i < len(items) - 1:
                    self._probation_arrow(slide, x + width + .08, 2.58, .25, .36)
            self._probation_panel(slide, .72, 4.48, 11.88, 1.65)
            self.text(slide, page["bottom_title"], .98, 4.76, 2.0, .32, 17, "accent", True)
            self.text(slide, page["bottom"], 3.0, 4.7, 9.1, .72, 16, "text")
        elif layout == "root_cause":
            images = page["images"]
            self._probation_image(slide, images[0]["path"], .72, 1.65, 2.85, 2.2, used_assets,
                                  crop=images[0].get("crop"))
            self.text(slide, images[0]["label"], .82, 3.92, 2.65, .3, 12, "muted", True)
            self._probation_image(slide, images[1]["path"], .72, 4.35, 2.85, 1.85, used_assets,
                                  crop=images[1].get("crop"))
            self.text(slide, images[1]["label"], .82, 6.28, 2.65, .3, 12, "muted", True)
            self.text(slide, page["result"], 3.95, 1.68, 8.45, .68, 21, "accent", True)
            for i, item in enumerate(page["items"]):
                self._probation_card(slide, item["title"], item["body"], 3.95,
                                     2.55 + i * 1.18, 8.45, .93, i + 1, body_size=12.5)
        elif layout == "depth_pair":
            images = depth_pair(self.base / page["before_array"], self.base / page["after_array"],
                                page["crop"], page["limits_m"])
            for i, image in enumerate(images):
                x = .85 + i * 6.1
                self.text(slide, page["labels"][i], x, 1.65, 5.4, .4, 21, "accent", True)
                self._picture_image(slide, image, x, 2.18, 5.4, 2.75)
            # Native legend uses the identical palette/range as both images.
            for i, rgb in enumerate(depth_colors([i / 59 for i in range(60)])):
                shape = self.box(slide, 3.7 + i * .1, 5.25, .102, .18)
                shape.fill.fore_color.rgb = RGBColor(*[int(v) for v in rgb])
            self.text(slide, f"{page['limits_m'][0]:.2f} m", 2.8, 5.2, .85, .3, 13, "muted")
            self.text(slide, f"{page['limits_m'][1]:.2f} m", 9.8, 5.2, .9, .3, 13, "muted")
            self.text(slide, page["note"], .85, 5.85, 11.6, .9, 17)
        elif layout == "depth_case":
            digest = hashlib.sha256((self.base / page["image"]).read_bytes()).hexdigest()
            if digest in used_assets:
                raise ValueError("depth case photo is reused on another slide")
            used_assets.add(digest)
            depth_images = depth_pair(self.base / page["before_array"],
                self.base / page["after_array"], page["depth_crop"], page["limits_m"])
            for i, phase in enumerate(("before", "after")):
                x = .72 if i == 0 else 8.13
                result = candidate_projections(self.base / page[f"{phase}_response"],
                    self.base / page["calibration"], page.get("display_count", 3))
                if result["total"] != page[f"{phase}_count"]:
                    raise ValueError("depth case candidate count does not match source")
                self.text(slide, "优化前" if i == 0 else "优化后", x, 1.6, 4.48, .4,
                          21, "accent", True)
                self._pose_photo(slide, page, result["projections"], x, 2.18, 4.48, 2.25)
                self.text(slide, "原始深度" if i == 0 else "补全深度", x, 4.63, 2.3, .25,
                          12, "muted")
                self._picture_image(slide, depth_images[i], x, 5.0, 2.3, 1.05)
                self.text(slide, f"{result['total']} 个候选\n{page[f'{phase}_points']} 个目标点",
                          x + 2.55, 4.98, 1.93, 1.1, 16, "text", True)
            self.text(slide, "优化方案", 5.44, 1.6, 2.38, .4, 21, "accent", True)
            self.text(slide, page["solution"], 5.44, 2.4, 2.38, 2.6, 17)
            for i, rgb in enumerate(depth_colors([i / 59 for i in range(60)])):
                shape = self.box(slide, 5.44 + i * .039, 5.36, .04, .14)
                shape.fill.fore_color.rgb = RGBColor(*[int(v) for v in rgb])
            self.text(slide, f"{page['limits_m'][0]:.2f}–{page['limits_m'][1]:.2f} m",
                      5.44, 5.68, 2.38, .25, 12, "muted")
            self.text(slide, page["note"], .82, 6.5, 11.6, .32, 12, "muted")
        elif layout == "pose_case":
            if len(page["frames"]) != 3:
                raise ValueError("pose case requires three paired rounds")
            for label, x, w in [("优化前", .72, 4.78), ("优化方案", 5.62, 1.95),
                                 ("优化后", 7.72, 4.78)]:
                self.text(slide, label, x, 1.58, w, .4, 21, "accent", True)
            # The middle column is intentionally narrow; keep the root-cause
            # wording readable without forcing it into a separate slide.
            self.text(slide, page["solution"], 5.62, 2.48, 1.95, 3.18, 13.5)
            for i, frame in enumerate(page["frames"]):
                digest = hashlib.sha256((self.base / frame["image"]).read_bytes()).hexdigest()
                if digest in used_assets:
                    raise ValueError("comparison frame is reused on another slide")
                used_assets.add(digest)
                y = 2.18 + i * 1.36
                for phase, x in (("before", .72), ("after", 7.72)):
                    gate = gated_candidate_projections(self.base / frame[phase],
                        self.base / frame["calibration"])
                    projection = gripper_projection(self.base / frame[phase],
                        self.base / frame["calibration"])
                    if abs(projection["opening_mm"] - frame[f"{phase}_opening_mm"]) > .02:
                        raise ValueError("recorded gripper opening does not match comparison statistics")
                    self._pose_photo(slide, frame, gate["projections"], x, y, 2.15, 1.08)
                    self.text(slide,
                        f"第 {i+1} 轮 · {projection['opening_mm']:.2f} mm\n"
                        f"蓝 {gate['accepted']} · 红 {gate['rejected']}",
                        x + 2.36, y + .14, 2.42, .81, 14)
            self.text(slide, page["note"], .82, 6.55, 11.6, .3, 12, "muted")
        elif layout == "efficiency_case":
            for label, x in (("优化前", .82), ("优化方案", 4.98), ("优化后", 9.02)):
                self.text(slide, label, x, 1.7, 3.5, .4, 21, "accent", True)
            self.text(slide, page["before"], .82, 2.48, 3.3, .62, 32, "text", True)
            self.text(slide, page["solution"], 4.98, 2.42, 3.5, 1.45, 17)
            self.text(slide, page["after"], 9.02, 2.48, 3.3, .62, 32, "accent", True)
            self.text(slide, page["delta"], 9.02, 3.22, 3.3, .4, 17, "accent")
            self.text(slide, "优化后耗时构成", .82, 3.72, 3.6, .4, 21, "text", True)
            self.text(slide, page["after_detail"], 4.98, 3.68, 7.4, .62, 14)
            self.text(slide, "夹爪几何检查", .82, 4.34, 3.6, .4, 21, "text", True)
            self.text(slide, page["check_solution"], 4.98, 4.34, 7.4, .65, 14)
            table = slide.shapes.add_table(len(page["rows"]) + 1, 4, Inches(.82),
                Inches(5.08), Inches(11.7), Inches(1.0)).table
            for ri, row in enumerate([page["headers"]] + page["rows"]):
                for ci, value in enumerate(row):
                    cell = table.cell(ri, ci)
                    cell.text = str(value)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string(self.colors["accent" if ri == 0 else "panel"])
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.name = self.font
                        paragraph.font.size = Pt(16)
                        paragraph.font.color.rgb = RGBColor.from_string(self.colors["bg" if ri == 0 else "text"])
                        for run in paragraph.runs:
                            run._r.get_or_add_rPr().append(run._r.makeelement(qn("a:ea"), {"typeface": self.font}))
            self.text(slide, page["note"], .82, 6.58, 11.6, .3, 12, "muted")
        elif layout == "grasp_comparison":
            frames = page["frames"]
            if len(frames) != 3:
                raise ValueError("grasp comparison requires three paired rounds")
            self.text(slide, "优化前", .55, 2.55, .9, .65, 18, "muted", True)
            self.text(slide, "优化后", .55, 4.82, .9, .65, 18, "accent", True)
            for i, frame in enumerate(frames):
                digest = hashlib.sha256((self.base / frame["image"]).read_bytes()).hexdigest()
                if digest in used_assets:
                    raise ValueError("comparison frame is reused on another slide")
                used_assets.add(digest)
                x = 1.55 + i * 3.72
                self.text(slide, frame["label"], x, 1.63, 3.4, .35, 17, "accent", True)
                self._grasp_panel(slide, frame, "before", x, 2.08, 3.45, 1.64)
                self._grasp_panel(slide, frame, "after", x, 4.35, 3.45, 1.64)
            self.text(slide, page["note"], .82, 6.54, 11.6, .3, 12, "muted")
        elif layout == "evidence":
            has_image = bool(page.get("image"))
            if has_image:
                self._probation_image(slide, page["image"], .7, 1.72, 3.8, 4.65, used_assets,
                                      crop=page.get("image_crop"))
            for i, metric in enumerate(page["charts"]):
                x, y, w, h = ((4.85, 1.6 + i * 2.1, 7.4, 1.55) if has_image
                               else (.82 + i * 6.12, 1.82, 5.55, 3.65))
                self.text(slide, metric["title"], x, y, w, .4, 19, "accent", True)
                data = CategoryChartData()
                data.categories = ["第 1 轮", "第 2 轮", "第 3 轮"]
                data.add_series("优化前", metric["before"])
                data.add_series("优化后", metric["after"])
                chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                    Inches(x), Inches(y + .55), Inches(w), Inches(h), data).chart
                chart.has_legend = True
                chart.legend.font.size = Pt(10)
                chart.legend.font.name = self.font
                chart.legend.include_in_layout = False
                chart.plots[0].has_data_labels = True
                chart.plots[0].data_labels.font.size = Pt(11)
                chart.plots[0].data_labels.font.name = self.font
                for axis in (chart.category_axis, chart.value_axis):
                    axis.tick_labels.font.size = Pt(10)
                    axis.tick_labels.font.name = self.font
                for j, series in enumerate(chart.series):
                    series.format.fill.solid()
                    series.format.fill.fore_color.rgb = RGBColor.from_string("D97706" if j == 0 else "176C88")
            note_x, note_y, note_w = (4.95, 5.9, 7.25) if has_image else (.82, 6.2, 11.7)
            self.text(slide, page["note"], note_x, note_y, note_w, .42, 13, "text")
        elif layout == "performance":
            self.text(slide, "同一 1280×720 场景的完整感知调用", .82, 1.72, 6.2, .32, 15, "muted")
            self._probation_panel(slide, .82, 2.2, 11.7, 1.52)
            self.text(slide, page["before"], 1.22, 2.62, 2.7, .55, 31, "muted", True)
            self._probation_arrow(slide, 4.15, 2.68, .5, .5)
            self.text(slide, page["after"], 5.05, 2.62, 2.7, .55, 31, "accent", True)
            self._probation_tag(slide, page["delta"], 8.6, 2.78, 2.5)
            self.text(slide, page["caption"], 1.22, 3.35, 10.3, .25, 12, "muted")
            self._probation_card(slide, "为什么变快", page["reason"], .82, 4.22, 5.55, 1.55, 1, body_size=14)
            self._probation_card(slide, "工程侧同步", page["engineering"], 6.95, 4.22, 5.55, 1.55, 2, body_size=14)
        elif layout == "field":
            images = page["images"]
            if len(images) == 1:
                item = images[0]
                self._probation_image(slide, item["path"], .72, 1.7, 7.15, 4.2, used_assets)
                self.text(slide, item["label"], 8.2, 2.05, 4.25, .55, 20, "accent", True)
                self.text(slide, item["caption"], 8.2, 2.85, 4.25, 1.25, 17, "text")
                self._probation_panel(slide, 8.2, 4.55, 4.25, 1.15, rounded=False)
                self.text(slide, page["flow"], 8.45, 4.82, 3.75, .78, 14, "text", True)
            else:
                for i, item in enumerate(images):
                    x = .72 + i * 6.15
                    self._probation_image(slide, item["path"], x, 1.65, 5.65, 3.45, used_assets)
                    self.text(slide, item["label"], x + .12, 5.22, 5.4, .3, 14, "accent", True)
                    self.text(slide, item["caption"], x + .12, 5.57, 5.4, .36, 12, "muted")
                self.text(slide, page["flow"], .82, 6.35, 11.8, .35, 17, "text", True)
        elif layout == "field_story":
            for i, item in enumerate(page["images"]):
                x = .72 + i * 4.17
                self._probation_image(slide, item["path"], x, 1.82, 3.75, 2.4, used_assets)
                self.text(slide, item["label"], x, 4.5, 3.75, .4, 20, "accent", True)
                self.text(slide, item["caption"], x, 5.12, 3.75, 1.05, 16)
            self.text(slide, page["note"], .82, 6.55, 11.6, .3, 12, "muted")
        elif layout == "visual_gallery":
            if len(page["items"]) != 3:
                raise ValueError("visual gallery requires three source captures")
            for i, item in enumerate(page["items"]):
                x = .72 + i * 4.17
                self.text(slide, item["title"], x, 1.65, 3.75, .4, 20, "accent", True)
                self._layered_photo(slide, item, x, 2.25, 3.75, 2.8, used_assets)
                self.text(slide, item["body"], x, 5.32, 3.75, .95, 15)
            self.text(slide, page["note"], .82, 6.55, 11.6, .3, 12, "muted")
        elif layout == "native_gallery":
            if len(page["items"]) != 3:
                raise ValueError("native gallery requires three source captures")
            for i, item in enumerate(page["items"]):
                x = .72 + i * 4.17
                digest = hashlib.sha256((self.base / item["image"]).read_bytes()).hexdigest()
                if digest in used_assets:
                    raise ValueError("native gallery photo is reused on another slide")
                used_assets.add(digest)
                projection = gripper_projection(self.base / item["pose"],
                    self.base / item["calibration"])
                self.text(slide, item["title"], x, 1.65, 3.75, .4, 20, "accent", True)
                self._pose_photo(slide, item, [projection], x, 2.25, 3.75, 2.8)
                self.text(slide, item["body"], x, 5.32, 3.75, .95, 15)
            self.text(slide, page["note"], .82, 6.55, 11.6, .3, 12, "muted")
        elif layout == "data_table":
            self.compose(slide, {"layout": "table", "headers": page["headers"],
                                 "rows": page["rows"], "source": page["note"]})
        elif layout == "wrc":
            self._probation_image(slide, page["image"], .72, 1.65, 4.2, 4.75, used_assets)
            for i, item in enumerate(page["items"]):
                self._probation_card(slide, item["title"], item["body"], 5.35, 1.65 + i * 1.52,
                                     6.95, 1.2, i + 1, body_size=13)
        elif layout == "voice":
            width = 2.78
            for i, item in enumerate(page["items"]):
                x = .58 + i * 3.18
                self._probation_card(slide, item["title"], item["body"], x, 2.0, width, 2.1,
                                     i + 1, body_size=12.5)
                if i < len(page["items"]) - 1:
                    self._probation_arrow(slide, x + width + .1, 2.78, .22, .34)
            if page.get("bottom"):
                self._probation_panel(slide, .82, 4.72, 11.7, 1.18)
                self.text(slide, page["bottom"], 1.1, 5.05, 11.1, .42, 17, "accent", True)
        elif layout == "engineering":
            for i, item in enumerate(page["items"]):
                y = 1.7 + i * 1.02
                self.text(slide, f"{i+1:02d}", .9, y + .13, .7, .46, 23, "accent", True)
                self.text(slide, item["title"], 1.9, y + .13, 2.35, .4, 20, "text", True)
                self.text(slide, item["body"], 5.0, y + .14, 7.15, .58, 17)
                self.box(slide, .9, y + .86, 11.35, .013, "muted")
            self.text(slide, page["bottom"], .86, 6.35, 11.6, .45, 13, "muted")
        elif layout == "contribution":
            self.text(slide, page["lead"], .84, 1.48, 11.4, .62, 17.5, "accent", True)
            items = page["items"]
            width = 3.78
            for i, item in enumerate(items):
                x = .82 + i * 4.18
                self._probation_panel(slide, x, 2.28, width, 3.55, rounded=True)
                self.text(slide, f"{i + 1:02d}", x + .22, 2.52, .6, .4, 22, "accent", True)
                self.text(slide, item["title"], x + .22, 3.03, width - .44, .42,
                          19, "text", True)
                if item.get("tag"):
                    self._probation_tag(slide, item["tag"], x + .22, 3.58,
                                        min(1.55, width - .44))
                self.text(slide, item["body"], x + .22, 4.08, width - .44, 1.35,
                          14.5)
            self._probation_panel(slide, .82, 6.02, 11.7, .62, rounded=False)
            self.text(slide, page.get("bottom", "从代码产出到现场验证，形成可复用的机器人能力。"),
                      1.05, 6.18, 11.2, .35, 12.5, "text", True)
        elif layout == "improvement":
            self.text(slide, page["lead"], .84, 1.57, 11.4, .42, 19, "accent", True)
            self.box(slide, .8, 2.25, 11.7, .62, "accent")
            for label, x, w in [("改进方向", 1.0, 2.1), ("当前不足", 3.25, 3.8), ("具体行动", 7.65, 4.3)]:
                self.text(slide, label, x, 2.43, w, .31, 16, "bg", True)
            for i, item in enumerate(page["items"]):
                y = 2.92 + i * 1.08
                self.box(slide, .8, y, 11.7, 1.03, "panel")
                before, action = item["body"].split("\n", 1)
                self.text(slide, item["title"], 1.0, y + .25, 2.1, .43, 19, "accent", True)
                self.text(slide, before.removeprefix("现状："), 3.25, y + .18, 3.8, .77, 17)
                self.text(slide, action.removeprefix("动作："), 7.65, y + .18, 4.3, .77, 17)
        elif layout == "roadmap":
            # Match the supplied template's growth page: one horizontal
            # timeline, three colored milestones, and aligned bullet tracks.
            milestone_colors = ["2DBE62", "4285E8", "F36C2B"]
            xs = [.82, 4.65, 8.48]
            centers = [1.82, 5.65, 9.48]
            self.box(slide, 1.82, 2.43, 9.75, .035, "text")
            for i, (item, x, center, color) in enumerate(
                    zip(page["items"], xs, centers, milestone_colors)):
                outer = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(center - .23),
                                               Inches(2.19), Inches(.46), Inches(.46))
                outer.fill.solid()
                outer.fill.fore_color.rgb = RGBColor.from_string("0B1324")
                outer.line.color.rgb = RGBColor.from_string(color)
                outer.line.width = Pt(3.2)
                self.text(slide, f"{i + 1:02d}", center - .16, 2.28, .32, .24,
                          9, color, True)
                title = item["title"].replace("｜", " | ")
                self.text(slide, title, x + .52, 1.84, 2.75, .42, 18, color, True)
                self.box(slide, x, 3.58, 3.2, .012, color)
                parts = [part.strip() for part in item["body"].replace("；", "\n").split("\n")
                         if part.strip()]
                for j, part in enumerate(parts[:3]):
                    y = 3.86 + j * .72
                    bullet = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y + .08),
                                                    Inches(.09), Inches(.09))
                    bullet.fill.solid()
                    bullet.fill.fore_color.rgb = RGBColor.from_string(color)
                    bullet.line.fill.background()
                    self.text(slide, part, x + .26, y, 3.0, .5, 13.5, "text")
            self.box(slide, .82, 6.48, 11.7, .012, "text")
            self.text(slide, page.get("bottom", "按阶段推进，从当前工作延伸到通用机器人应用。"),
                      3.05, 6.67, 8.5, .25, 12.5, "muted")
        elif layout == "research_plan":
            for i, item in enumerate(page["items"]):
                x, y = .82 + (i % 2) * 6.12, 1.65 + (i // 2) * 2.45
                self._probation_panel(slide, x, y, 5.55, 2.05, rounded=False)
                self.box(slide, x, y, .065, 2.05, "accent")
                self.text(slide, item["title"], x + .23, y + .18, 5.0, .4, 21, "accent", True)
                self.text(slide, item["body"], x + .23, y + .8, 5.0, 1.1, 16)
        elif layout == "application_plan":
            for i, item in enumerate(page["models"]):
                x = .82 + i * 6.12
                self._probation_card(slide, item["title"], item["body"], x, 1.65,
                                     5.55, 1.4, body_size=15)
            self.text(slide, "部署与推理服务", .9, 3.33, 3.1, .42, 20, "accent", True)
            self.text(slide, page["integration"], 4.15, 3.3, 8.1, .65, 17)
            for i, item in enumerate(page["steps"]):
                x = .82 + i * 4.15
                self._probation_card(slide, item["title"], item["body"], x, 4.22,
                                     3.7, 1.65, body_size=15)
                if i < len(page["steps"]) - 1:
                    self._probation_arrow(slide, x + 3.78, 4.85, .23, .38)
            self.text(slide, page["note"], .9, 6.43, 11.45, .32, 14, "muted")
        elif layout == "summary":
            self._probation_card(slide, page["left_title"], page["left_body"], .82, 1.8, 5.6, 3.65,
                                 1, body_size=15)
            self._probation_card(slide, page["right_title"], page["right_body"], 6.92, 1.8, 5.6, 3.65,
                                 2, body_size=15)
            self.text(slide, page["bottom"], .92, 6.0, 11.5, .42, 18, "accent", True)
        else:
            raise ValueError(f"unsupported probation layout: {layout}")


def build(spec: dict, base: Path, output: Path) -> None:
    """Validate and build; only spec assets are read and output is written."""
    validate(spec, base)
    builder = DeckBuilder(spec, base)
    for index, page in enumerate(spec["slides"], 1):
        slide = builder.new(page, index)
        builder.compose(slide, page)
    output.parent.mkdir(parents=True, exist_ok=True)
    builder.prs.save(str(output))


def _template_page_map(spec: dict, base: Path, template: Presentation) -> list[dict]:
    """Validate a one-to-one page map for template-preserving generation."""
    pages = spec.get("slides")
    if not isinstance(pages, list) or len(pages) != len(template.slides):
        raise ValueError("template mode requires one slide spec for every template slide")
    seen: set[int] = set()
    for index, page in enumerate(pages, 1):
        if not isinstance(page, dict):
            raise ValueError(f"slide {index}: template page spec must be an object")
        source_index = page.get("template_slide")
        if type(source_index) is not int or not 1 <= source_index <= len(template.slides):
            raise ValueError(f"slide {index}: template_slide must be a valid 1-based slide number")
        if source_index in seen:
            raise ValueError(f"slide {index}: template_slide {source_index} is used more than once")
        seen.add(source_index)
        for field in ("text", "images"):
            values = page.get(field, {})
            if not isinstance(values, dict) or any(not isinstance(k, str) for k in values):
                raise ValueError(f"slide {index}: {field} must map shape names to values")
        slide = template.slides[source_index - 1]
        shapes = {shape.name: shape for shape in slide.shapes}
        for name, value in page.get("text", {}).items():
            if name not in shapes or not getattr(shapes[name], "has_text_frame", False):
                raise ValueError(f"slide {index}: text target is not a text shape: {name}")
            if not isinstance(value, str):
                raise ValueError(f"slide {index}: text for {name} must be a string")
        for name, filename in page.get("images", {}).items():
            if name not in shapes or shapes[name].shape_type != 13:
                raise ValueError(f"slide {index}: image target is not a picture shape: {name}")
            image_path = (Path(filename) if Path(filename).is_absolute() else base / filename).resolve()
            if not image_path.is_file():
                raise ValueError(f"slide {index}: image not found: {image_path}")
            with Image.open(image_path) as picture:
                picture.verify()
        notes = page.get("notes", "")
        if not isinstance(notes, str):
            raise ValueError(f"slide {index}: notes must be a string")
    if seen != set(range(1, len(template.slides) + 1)):
        raise ValueError("template_slide values must cover every template slide exactly once")
    return pages


def _replace_text(shape, value: str) -> None:
    """Replace text while retaining the first run's font and paragraph styling."""
    frame = shape.text_frame
    runs = [run for paragraph in frame.paragraphs for run in paragraph.runs]
    if runs:
        runs[0].text = value
        for run in runs[1:]:
            run.text = ""
        for paragraph in frame.paragraphs[1:]:
            for run in paragraph.runs:
                run.text = ""
    else:
        frame.paragraphs[0].add_run().text = value


def _replace_text_with_style(shape, value: str) -> None:
    """Replace all text while preserving the existing paragraph/run styling."""
    frame = shape.text_frame
    paragraphs = frame.paragraphs
    if not paragraphs:
        frame.text = value
        return
    first_run = next((run for paragraph in paragraphs for run in paragraph.runs), None)
    if first_run is None:
        frame.text = value
        return
    first_run.text = value
    for paragraph in paragraphs:
        for run in paragraph.runs:
            if run is not first_run:
                run.text = ""
    for paragraph in paragraphs[1:]:
        paragraph.text = ""


def build_from_template(spec: dict, base: Path, template_path: Path, output: Path) -> None:
    """Edit named text and picture shapes without rebuilding template slides."""
    template = Presentation(str(template_path))
    pages = _template_page_map(spec, base, template)
    for page in pages:
        slide = template.slides[page["template_slide"] - 1]
        shapes = {shape.name: shape for shape in slide.shapes}
        for name, value in page.get("text", {}).items():
            _replace_text_with_style(shapes[name], value)
        for name, filename in page.get("images", {}).items():
            shape = shapes[name]
            image_path = (Path(filename) if Path(filename).is_absolute() else base / filename).resolve()
            image_part, _ = slide.part.get_or_add_image_part(str(image_path))
            blip = shape._element.blipFill.blip
            blip.set(qn("r:embed"), slide.part.relate_to(image_part, RT.IMAGE))
        if "notes" in page:
            slide.notes_slide.notes_text_frame.text = page["notes"]
    output.parent.mkdir(parents=True, exist_ok=True)
    template.save(str(output))


def build_probation_deck_v2(spec: dict, base: Path, template_path: Path, output: Path) -> None:
    """Keep the reference cover untouched and compose varied native pages."""
    template = Presentation(str(template_path))
    if len(template.slides) < 1:
        raise ValueError("probation mode requires a template with at least one slide")
    if (abs(template.slide_width / Inches(1) - 13.333333) > .02
            or abs(template.slide_height / Inches(1) - 7.5) > .02):
        raise ValueError("native_v2 layouts require a 13.333 × 7.5 inch template")
    prs = Presentation(str(template_path))
    # Slide 1 is the supplied cover.  Remove only the template's example
    # content pages; this leaves the original cover XML, artwork and wording
    # unchanged while retaining its masters and theme for new pages.
    while len(prs.slides) > 1:
        slide_id = prs.slides._sldIdLst[-1]
        prs.part.drop_rel(slide_id.rId)
        prs.slides._sldIdLst.remove(slide_id)
    builder = DeckBuilder(spec, base)
    builder.prs = prs
    builder.font = spec.get("meta", {}).get("font", "Microsoft YaHei")
    builder.prs.slide_width = template.slide_width
    builder.prs.slide_height = template.slide_height
    used_assets: set[str] = set()
    pages = spec["slides"]
    if not pages or pages[0].get("type") != "cover":
        raise ValueError("native_v2 spec must start with a cover page")
    for index, page in enumerate(pages[1:], 2):
        preferred_layout = (0 if page["layout"] == "thanks" else
                            3 if len(builder.prs.slide_layouts) > 3 else len(builder.prs.slide_layouts) - 1)
        slide = builder.prs.slides.add_slide(builder.prs.slide_layouts[preferred_layout])
        for shape in list(slide.shapes):
            if shape.is_placeholder:
                shape._element.getparent().remove(shape._element)
        builder.compose_probation_v2(slide, page, index, used_assets)
        slide.notes_slide.notes_text_frame.text = page.get("notes", "")
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))


def build_probation_deck(spec: dict, base: Path, template_path: Path, output: Path) -> None:
    """Build a complete editable review deck using the supplied template DNA.

    The template contributes its masters, theme, background and fonts. Content
    pages are composed by this repository-owned engine from explicit JSON and
    local assets; no workspace generator is imported or executed.
    """
    if spec.get("meta", {}).get("deck_mode") == "native_v2":
        build_probation_deck_v2(spec, base, template_path, output)
        return
    template = Presentation(str(template_path))
    if len(template.slides) < 1:
        raise ValueError("probation mode requires a template with at least one slide")
    reference = template.slides[0]
    prs = Presentation(str(template_path))
    while len(prs.slides):
        slide_id = prs.slides._sldIdLst[-1]
        prs.part.drop_rel(slide_id.rId)
        prs.slides._sldIdLst.remove(slide_id)
    builder = DeckBuilder(spec, base)
    builder.prs = prs
    builder.font = spec.get("meta", {}).get("font", "Microsoft YaHei")
    builder.prs.slide_width = template.slide_width
    builder.prs.slide_height = template.slide_height
    builder.reference = reference
    for index, page in enumerate(spec["slides"], 1):
        # The supplied reference deck has only five layouts.  Select an
        # existing layout so the generated pages retain the template master
        # without assuming a blank layout at a fixed index.
        preferred_layout = (
            0 if page.get("type") == "cover"
            else 3 if page.get("layout") == "closing"
            else 1
        )
        layout_index = min(preferred_layout, len(builder.prs.slide_layouts) - 1)
        slide = builder.prs.slides.add_slide(builder.prs.slide_layouts[layout_index])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string("F4F7FB")
        builder.text(slide, page.get("title", spec["meta"]["title"]), .62, .38, 10.8, .45, 25, "text", True)
        builder.box(slide, .62, 1.0, 1.15, .06, "accent")
        if page.get("subtitle") and page.get("type") not in {"cover", "closing"}:
            builder.text(slide, page["subtitle"], 1.92, .85, 10.7, .28, 10.5, "muted")
        builder.text(slide, "人形机器人（上海）有限公司  |  轮臂-应用开发  |  赵文辉", .55, 7.12, 10.5, .18, 8.5, "muted")
        builder.text(slide, f"{index:02d}", 12.25, 7.08, .45, .22, 9, "muted")
        if page.get("type") == "cover":
            builder.text(slide, page.get("body", ""), .8, 1.8, 7.0, 2.5, 27, "accent", True)
            builder.text(slide, page.get("subtitle", ""), .8, 4.45, 7.0, 1.2, 15, "muted")
            if page.get("image"):
                builder.picture(slide, page["image"], 8.4, 1.4, 4.0, 4.9)
        elif page.get("type") == "closing":
            builder.text(slide, page.get("body", "谢谢聆听"), .8, 2.2, 8.5, 1.5, 30, "accent", True)
            builder.text(slide, page.get("subtitle", ""), .8, 3.65, 8.5, 1.2, 17, "muted")
        else:
            builder.compose(slide, page)
        slide.notes_slide.notes_text_frame.text = page.get("notes", "")
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("spec", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--template", type=Path,
                        help="reuse every slide from this PPTX, replacing named text/image shapes")
    parser.add_argument("--probation", action="store_true",
                        help="compose a complete probation deck while retaining the template theme and masters")
    args = parser.parse_args()
    try:
        spec = json.loads(args.spec.read_text(encoding="utf-8"))
        base = args.spec.resolve().parent
        if args.probation:
            if not args.template:
                parser.error("--probation requires --template")
            build_probation_deck(spec, base, args.template.resolve(), args.output.resolve())
        elif args.template:
            build_from_template(spec, base, args.template.resolve(), args.output.resolve())
        else:
            build(spec, base, args.output.resolve())
    except (ValueError, OSError, KeyError, TypeError) as exc:
        parser.exit(2, f"build_deck: {exc}\n")
    print(f"Generated {len(spec['slides'])} editable slides: {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
