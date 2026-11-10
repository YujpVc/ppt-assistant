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
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

import media
from charts import draw_chart, resolve_chart

from grasp_evidence import (candidate_projections, depth_colors, depth_pair,
                            gated_candidate_projections, gripper_projection)

LAYOUTS = {"cover", "cards", "process", "architecture", "sequence", "pillars", "comparison", "table", "chart",
           "image_split", "image_grid", "image_triptych", "image_band", "image_panel", "image_panel_strip", "ingredient_board", "ingredient_split", "state_triptych", "state_band", "state_gallery", "prep_stove", "heat_bars", "chart_dashboard", "closing",
           "image_hero", "image_overlay", "image_wall",
           "section_number", "image_statement",
           "split_diagonal", "glass_overlay", "quote_stage", "kpi_showcase",
           "cutout_statement", "image_annotated", "image_arch"}
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
    theme = spec["meta"].get("theme", "midnight")
    if theme not in THEMES:
        raise ValueError(f"meta.theme must be one of {sorted(THEMES)}")
    overrides = spec["meta"].get("colors") or {}
    if not isinstance(overrides, dict):
        raise ValueError("meta.colors must be an object")
    for key, value in overrides.items():
        if key not in THEMES[theme]:
            raise ValueError(f"meta.colors.{key} is not a {theme} token")
        text = str(value).strip().lstrip("#")
        if len(text) != 6 or any(char not in "0123456789ABCDEFabcdef" for char in text):
            raise ValueError(f"meta.colors.{key} must be a 6-digit hex color")
    title_font = spec["meta"].get("titleFont")
    if title_font is not None and (not isinstance(title_font, str) or not title_font.strip()):
        raise ValueError("meta.titleFont must be a non-empty string")
    effect = spec["meta"].get("transition")
    if effect not in {None, "fade", "push", "morph"}:
        raise ValueError("meta.transition must be fade, push or morph")
    accent = spec["meta"].get("accent")
    if accent is not None:
        text = str(accent).strip().lstrip("#")
        if len(text) != 6 or any(char not in "0123456789ABCDEFabcdef" for char in text):
            raise ValueError("meta.accent must be a 6-digit hex color")
    slides = spec.get("slides")
    seen_ids: set[str] = set()
    if not isinstance(slides, list) or not slides:
        raise ValueError("slides must be a non-empty array")
    for index, slide in enumerate(slides, 1):
        if not isinstance(slide, dict) or slide.get("layout") not in LAYOUTS:
            raise ValueError(f"slide {index}: unsupported layout")
        stable_id = slide.get("id", f"P{index:02d}")
        if (not isinstance(stable_id, str) or not 1 <= len(stable_id) <= 24
                or any(char not in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-"
                       for char in stable_id)):
            raise ValueError(f"slide {index}: id must be 1-24 word characters")
        if stable_id in seen_ids:
            raise ValueError(f"slide {index}: duplicate id {stable_id}")
        seen_ids.add(stable_id)
        layout = slide["layout"]
        if not slide.get("title") and layout not in {"cover", "closing"}:
            raise ValueError(f"slide {index}: title is required")
        if layout in {"sequence", "pillars"}:
            items = slide.get("items")
            if not isinstance(items, list) or not 3 <= len(items) <= 4:
                raise ValueError(f"slide {index}: {layout} requires 3–4 items")
            if any(not isinstance(item, dict) or not item.get("title") or not item.get("body")
                   for item in items):
                raise ValueError(f"slide {index}: every {layout} item requires title and body")
        if layout == "architecture":
            items = slide.get("items")
            if not isinstance(items, list) or len(items) != 3:
                raise ValueError(f"slide {index}: architecture requires exactly 3 items")
            if any(not isinstance(item, dict) or not item.get("title") or not item.get("body")
                   for item in items):
                raise ValueError(f"slide {index}: each architecture item requires title and body")
            if slide.get("bottom") and not isinstance(slide["bottom"], str):
                raise ValueError(f"slide {index}: architecture bottom must be text")
        if layout in {"cards", "process", "comparison", "image_grid"}:
            items = slide.get("items")
            if layout == "comparison":
                minimum, maximum = 2, 2
            elif layout == "process":
                minimum, maximum = 2, 5
            else:
                minimum, maximum = 2, 4
            if not isinstance(items, list) or not minimum <= len(items) <= maximum:
                raise ValueError(f"slide {index}: items requires {minimum}–{maximum} entries")
            if any(not isinstance(item, dict) or not item.get("title") for item in items):
                raise ValueError(f"slide {index}: each item requires a title")
        assets = [slide] if layout in {"cover", "image_split", "image_hero",
                                       "image_overlay", "image_statement",
                                       "glass_overlay", "split_diagonal"} else []
        if layout == "heat_bars":
            items = slide.get("items")
            if not isinstance(items, list) or not 2 <= len(items) <= 5:
                raise ValueError(f"slide {index}: heat_bars requires 2–5 items")
            for item in items:
                if not isinstance(item, dict) or not item.get("title") or not item.get("body"):
                    raise ValueError(f"slide {index}: every heat_bars item requires title and body")
                if item.get("fire") not in {"大火", "中火", "小火", "微火"}:
                    raise ValueError(f"slide {index}: heat_bars fire must be 大火/中火/小火/微火")
                if not isinstance(item.get("minutes"), (int, float)) or float(item["minutes"]) <= 0:
                    raise ValueError(f"slide {index}: heat_bars minutes must be positive")
        if layout == "ingredient_split":
            if not slide.get("image"):
                raise ValueError(f"slide {index}: ingredient_split requires an image")
            items = slide.get("items")
            if not isinstance(items, list) or not 3 <= len(items) <= 6:
                raise ValueError(f"slide {index}: ingredient_split requires 3–6 items")
            for item in items:
                if not isinstance(item, dict) or not item.get("label") or not item.get("value"):
                    raise ValueError(f"slide {index}: every ingredient item requires label and value")
            assets = [slide]
        if layout == "state_gallery":
            items = slide.get("items")
            if not isinstance(items, list) or len(items) != 3:
                raise ValueError(f"slide {index}: state_gallery requires exactly 3 items")
            if any(not isinstance(item, dict) or not item.get("image") for item in items):
                raise ValueError(f"slide {index}: every state item requires an image")
            assets = items
        if layout == "state_band":
            items = slide.get("items")
            if not isinstance(items, list) or len(items) != 3:
                raise ValueError(f"slide {index}: state_band requires exactly 3 items")
            if any(not isinstance(item, dict) or not item.get("image") for item in items):
                raise ValueError(f"slide {index}: every state item requires an image")
            assets = items
        if layout == "state_triptych":
            items = slide.get("items")
            if not isinstance(items, list) or len(items) != 3:
                raise ValueError(f"slide {index}: state_triptych requires exactly 3 items")
            if any(not isinstance(item, dict) or not item.get("image") for item in items):
                raise ValueError(f"slide {index}: every state item requires an image")
            assets = items
        if layout == "ingredient_board":
            if not slide.get("image"):
                raise ValueError(f"slide {index}: ingredient_board requires an image")
            items = slide.get("items")
            if not isinstance(items, list) or not 3 <= len(items) <= 6:
                raise ValueError(f"slide {index}: ingredient_board requires 3–6 items")
            for item in items:
                if not isinstance(item, dict) or not item.get("label") or not item.get("value"):
                    raise ValueError(f"slide {index}: every ingredient item requires label and value")
            assets = [slide]
        if layout == "prep_stove":
            items = slide.get("items")
            if not isinstance(items, list) or not 3 <= len(items) <= 6:
                raise ValueError(f"slide {index}: prep_stove requires 3–6 items")
            for item in items:
                if not isinstance(item, dict) or item.get("lane") not in {"prep", "stove"}:
                    raise ValueError(f"slide {index}: prep_stove lane must be prep or stove")
                if not isinstance(item.get("minutes"), (int, float)) or float(item["minutes"]) <= 0:
                    raise ValueError(f"slide {index}: prep_stove minutes must be positive")
                if not item.get("title") or not item.get("body"):
                    raise ValueError(f"slide {index}: prep_stove item requires title and body")
        if layout == "image_triptych":
            items = slide.get("items")
            if not isinstance(items, list) or len(items) != 3:
                raise ValueError(f"slide {index}: image_triptych requires exactly 3 items")
            if any(not isinstance(item, dict) or not item.get("image") for item in items):
                raise ValueError(f"slide {index}: every triptych item requires an image")
            assets = items
        if layout == "image_grid":
            assets = slide["items"]
        if layout == "section_number" and slide.get("image"):
            assets = [slide]
            if slide.get("art") != "image":
                raise ValueError(f"slide {index}: section_number art must be image when image is set")
        if layout in {"section_number", "image_statement"}:
            if slide.get("art", "gradient") not in {"gradient", "outline", "image"}:
                raise ValueError(f"slide {index}: art must be gradient, outline or image")
        if layout in {"cutout_statement", "image_arch"}:
            if not slide.get("image"):
                raise ValueError(f"slide {index}: {layout} requires an image")
            assets = [slide]
        if layout == "image_band":
            if not slide.get("image"):
                raise ValueError(f"slide {index}: image_band requires an image")
            items = slide.get("items")
            if not isinstance(items, list) or len(items) != 3:
                raise ValueError(f"slide {index}: image_band requires exactly 3 items")
            if any(not isinstance(item, dict) or not item.get("title") or not item.get("body")
                   for item in items):
                raise ValueError(f"slide {index}: every image_band item requires title and body")
            assets = [slide]
        if layout in {"image_panel", "image_panel_strip"}:
            if not slide.get("image"):
                raise ValueError(f"slide {index}: image_panel requires an image")
            if slide.get("side", "left") not in {"left", "right"}:
                raise ValueError(f"slide {index}: image_panel side must be left or right")
            bullets = slide.get("bullets") or []
            if not isinstance(bullets, list) or not 2 <= len(bullets) <= 5:
                raise ValueError(f"slide {index}: image_panel requires 2–4 bullets")
            stats = slide.get("stats") or []
            if stats and (not isinstance(stats, list) or len(stats) > 3):
                raise ValueError(f"slide {index}: image_panel stats must be <= 3")
            assets = [slide]
            if slide.get("inset"):
                assets.append({"image": slide["inset"]})
        if layout == "image_annotated":
            points = slide.get("points")
            if not isinstance(points, list) or not 2 <= len(points) <= 4:
                raise ValueError(f"slide {index}: image_annotated requires 2–4 points")
            for point in points:
                if not isinstance(point, dict) or not point.get("text"):
                    raise ValueError(f"slide {index}: every point requires text")
                if not 0 <= float(point.get("x", -1)) <= 1 or not 0 <= float(point.get("y", -1)) <= 1:
                    raise ValueError(f"slide {index}: point x/y must be 0–1")
            assets = [slide]
        if layout == "split_diagonal":
            images = slide.get("images")
            if not isinstance(images, list) or len(images) != 2:
                raise ValueError(f"slide {index}: split_diagonal requires images: [a, b]")
            assets = [{"image": item} for item in images]
        if layout == "quote_stage" and not str(slide.get("quote", "")).strip():
            raise ValueError(f"slide {index}: quote_stage requires quote text")
        if layout == "kpi_showcase":
            stats = slide.get("stats")
            if not isinstance(stats, list) or not 2 <= len(stats) <= 3:
                raise ValueError(f"slide {index}: kpi_showcase requires 2–3 stats")
            if any(not isinstance(item, dict) or not item.get("value") for item in stats):
                raise ValueError(f"slide {index}: every stat requires value")
        if layout == "image_wall":
            items = slide.get("items")
            if not isinstance(items, list) or not 2 <= len(items) <= 4:
                raise ValueError(f"slide {index}: image_wall requires 2–4 items")
            if any(not isinstance(item, dict) or not item.get("image")
                   or not item.get("title") for item in items):
                raise ValueError(f"slide {index}: every wall item requires image and title")
            assets = items
        if layout in {"image_hero", "image_overlay"}:
            if not slide.get("image"):
                raise ValueError(f"slide {index}: {layout} requires an image")
            if slide.get("ink", "light") not in {"light", "dark"}:
                raise ValueError(f"slide {index}: ink must be light or dark")
            allowed = {"left", "right", "bottom", "top", "diagonal"}
            if layout == "image_hero" and slide.get("scrim", "left") not in allowed:
                raise ValueError(f"slide {index}: scrim must be one of {sorted(allowed)}")
            if layout == "image_overlay" and slide.get("side", "left") not in {"left", "right"}:
                raise ValueError(f"slide {index}: image_overlay side must be left or right")
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


def _soft_shadow(shape, blur_pt=14, distance_pt=6, opacity=34000,
                 color="000000") -> None:
    """Attach a restrained outer shadow to a picture (DrawingML, editable)."""
    sp_pr = shape._element.spPr
    for tag in ("a:effectLst", "a:effectDag"):
        for old in sp_pr.findall(qn(tag)):
            sp_pr.remove(old)
    effect = sp_pr.makeelement(qn("a:effectLst"), {})
    shadow = sp_pr.makeelement(qn("a:outerShdw"), {
        "blurRad": str(int(blur_pt * 12700)),
        "dist": str(int(distance_pt * 12700)),
        "dir": "5400000", "algn": "ctr", "rotWithShape": "0"})
    shade = sp_pr.makeelement(qn("a:srgbClr"), {"val": color})
    shade.append(sp_pr.makeelement(qn("a:alpha"), {"val": str(opacity)}))
    shadow.append(shade)
    effect.append(shadow)
    sp_pr.append(effect)


class DeckBuilder:
    """Own presentation state and layout primitives for one specification."""

    def __init__(self, spec: dict, base: Path):
        self.meta = spec["meta"]
        self.base = base
        theme = THEMES[self.meta.get("theme", "midnight")]
        self.colors = dict(theme)
        for key, value in (self.meta.get("colors") or {}).items():
            self.colors[key] = str(value).strip().lstrip("#").upper()
        if self.meta.get("accent"):
            self.colors["accent"] = str(self.meta["accent"]).strip().lstrip("#").upper()
        self.font = self.meta.get("font", "Noto Sans CJK SC")
        self.head_font = self.meta.get("titleFont") or self.font
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
        shape.shadow.inherit = False
        # Drop the theme style reference: an empty effectLst is not enough for
        # every renderer, and the theme shadow made thin rules look embossed.
        style = shape._element.find(qn("p:style"))
        if style is not None:
            shape._element.remove(style)
        return shape

    def text(self, slide, value, x, y, w, h, size=21, color="text", bold=False,
             anchor=None, align=None, tracking=None, font=None):
        """Fit text conservatively; reject dense content instead of shrinking it."""
        value = str(value)
        if not value.strip():
            return None
        plain = value.replace("**", "")
        lines = sum(max(1, math.ceil(sum(1 if ord(c) > 0x2E80 else .55 for c in line)
                                    / (w * 72 / size))) for line in plain.split("\n"))
        if lines * size * 1.25 + max(0, len(value.split("\n")) - 1) * 8 > h * 72:
            raise ValueError(f"text exceeds layout capacity; split or shorten: {value[:60]}")
        shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        frame = shape.text_frame
        frame.word_wrap = True
        if anchor is not None:
            frame.vertical_anchor = anchor
        frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
        base_color = self.colors.get(color, color)
        family = font or self.font
        for index, line in enumerate(value.split("\n")):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            paragraph.space_after = Pt(8)
            paragraph.line_spacing = 1.2
            if align is not None:
                paragraph.alignment = align
            for part_index, part in enumerate(line.split("**")):
                if not part:
                    continue
                emphasis = part_index % 2 == 1
                run = paragraph.add_run()
                run.text = part
                run.font.name, run.font.size = family, Pt(size)
                run.font.bold = bold or emphasis
                if tracking is not None:
                    run._r.get_or_add_rPr().set("spc", str(int(tracking)))
                run.font.color.rgb = RGBColor.from_string(
                    self.colors["accent"] if emphasis else base_color)
                east_asian = run._r.get_or_add_rPr().makeelement(qn("a:ea"), {"typeface": family})
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

    def _load(self, filename):
        with Image.open(self.base / filename) as source:
            return ImageOps.exif_transpose(source).convert("RGB")

    def artwork(self, slide, filename, x, y, w, h, *, radius=0.0,
                scrim=None, shadow=False, dpi=160, treatment=None, mask=None):
        """Place a picture as designed artwork: crop-to-fill, never letterboxed.

        ``radius`` rounds the corners (ratio of the short edge), ``scrim`` is a
        media.apply_scrim kwargs dict and ``shadow`` adds a soft outer shadow.
        """
        picture = media.cover(self._load(filename),
                              (max(1, int(w * dpi)), max(1, int(h * dpi))))
        if mask == "arch":
            picture = media.arch_mask(picture)
        elif mask == "circle":
            picture = media.circle_mask(picture)
        if treatment == "mono":
            picture = ImageOps.grayscale(picture).convert("RGB")
        elif treatment == "duotone":
            picture = media.duotone(picture, dark=self.colors["bg"],
                                    light="E8F6FA", strength=0.85)
        elif treatment == "halftone":
            picture = media.halftone(picture, spacing=6)
        elif treatment == "glitch":
            picture = media.glitch(picture, shift=5)
        if scrim:
            picture = media.apply_scrim(picture, **scrim)
        if radius:
            picture = media.rounded(picture, radius)
        payload = io.BytesIO()
        picture.save(payload, "PNG")
        payload.seek(0)
        shape = slide.shapes.add_picture(payload, Inches(x), Inches(y), Inches(w), Inches(h))
        if shadow:
            _soft_shadow(shape)
        return shape

    def _bleed_image(self, filename, *, scrim=None, color=None,
                     start=0.0, end=0.82, gamma=1.2, opacity=0.94,
                     vignette=0.0, focus=(0.5, 0.5), veil=0.0, grain=0.03):
        picture = media.cover(self._load(filename), (1920, 1080), centering=focus)
        tone = (media.parse_color(color) if color
                else media.sample_ink(picture, mix=0.45, mix_color=self.colors["bg"]))
        if veil:
            picture = media.veil(picture, color=tone, opacity=veil)
        if scrim:
            picture = media.apply_scrim(picture, color=tone, direction=scrim,
                                        start=start, end=end, gamma=gamma,
                                        opacity=opacity)
        if vignette:
            picture = media.vignette(picture, strength=vignette)
        if grain:
            picture = media.grain(picture, grain)
        return picture

    def full_bleed(self, slide, filename, *, scrim=None, color=None,
                   start=0.0, end=0.82, gamma=1.2, opacity=0.94,
                   vignette=0.0, focus=(0.5, 0.5), veil=0.0, grain=0.03):
        """A 16:9 picture that covers the whole slide and melts into a color."""
        picture = self._bleed_image(filename, scrim=scrim, color=color, start=start,
                                    end=end, gamma=gamma, opacity=opacity,
                                    vignette=vignette, focus=focus, veil=veil,
                                    grain=grain)
        self.add_full(slide, picture)
        tone = (color if isinstance(color, str) and len(color) == 6
                else media.to_hex(media.sample_ink(picture, mix=0.45,
                                                   mix_color=self.colors["bg"])))
        return tone.upper()

    def blank(self, page, index):
        """A slide without the standard title band (full-bleed compositions)."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(self.colors["bg"])
        slide.notes_slide.notes_text_frame.text = page.get("notes", "")
        return slide

    def hero_slide(self, slide, page):
        """Full-bleed image + gradient scrim + oversized headline."""
        ink = page.get("ink", "light")
        side = page.get("scrim", "left")
        # Dark text needs a light veil; light text needs the picture's own dark
        # tone. Sampling a dark scrim under dark text is unreadable by design.
        color = page.get("scrim_color")
        if color is None and ink == "dark":
            color = "F7FBFF"
        self.full_bleed(slide, page["image"], scrim=side,
                        color=color,
                        start=0.16 if side == "bottom" else 0.0,
                        end=page.get("scrim_end", 1.0 if side == "bottom" else 0.84),
                        opacity=page.get("scrim_opacity",
                                         0.90 if ink == "dark" else (0.96 if side == "bottom" else 0.95)),
                        vignette=page.get("vignette", 0.16),
                        veil=page.get("veil", 0.42 if ink == "dark" else 0.0))
        light = ink == "light"
        title_color = "FFFFFF" if light else self.colors["text"]
        muted = "CBD9EA" if light else self.colors["muted"]
        accent = self._accent_on(light)
        if side == "bottom":
            x, y, w = 0.95, 3.35, 11.45
        elif side == "right":
            x, y, w = 6.55, 1.95, 6.2
        else:
            x, y, w = 0.95, 1.95, 6.6
        kicker = page.get("kicker")
        if kicker:
            self.text(slide, str(kicker).upper(), x, y - 0.55, w, 0.32, 12, accent, True, tracking=160)
            self.box(slide, x, y - 0.16, 0.62, 0.045, accent)
        title = self.text(slide, page["title"], x, y, w, 2.15,
                          page.get("title_size", 42 if side != "bottom" else 38),
                          title_color, True, font=self.head_font)
        if title is not None:
            title.name = "title"
        if page.get("body"):
            self.text(slide, page["body"], x, y + 2.28, w, 1.35, 19, muted)
        meta = page.get("meta") or self.meta.get("author", "")
        if meta:
            self.text(slide, meta, x, 6.88 if side == "bottom" else 6.55, w, 0.35,
                      12, muted)
        if page.get("page_number", False):
            self.text(slide, f"{page.get('index', '')}", 12.0, 7.02, 0.7, 0.3, 11, muted)

    def overlay_slide(self, slide, page):
        """Full-bleed image + one-side scrim panel with title, body and stats."""
        side = page.get("side", "left")
        ink = page.get("ink", "light")
        self.full_bleed(slide, page["image"], scrim=side,
                        color=page.get("scrim_color"),
                        end=page.get("scrim_end", 0.60),
                        opacity=page.get("scrim_opacity", 0.93),
                        vignette=0.12)
        light = ink == "light"
        title_color = "FFFFFF" if light else self.colors["text"]
        body_color = "DCE8F6" if light else self.colors["text"]
        muted = "AFC3DA" if light else self.colors["muted"]
        accent = self._accent_on(light)
        x = 0.95 if side != "right" else 7.05
        w = 5.35
        top = 1.55
        if page.get("kicker"):
            self.text(slide, str(page["kicker"]).upper(), x, top, w, 0.3, 11, accent, True)
            top += 0.42
        title = self.text(slide, page["title"], x, top, w, 1.35,
                          page.get("title_size", 32), title_color, True,
                          font=self.head_font)
        if title is not None:
            title.name = "title"
        cursor = top + 1.62
        if page.get("body"):
            self.text(slide, page["body"], x, cursor, w, 1.35, 17, body_color)
            cursor += 1.05
        bullets = page.get("bullets") or []
        for position, bullet in enumerate(bullets[:4]):
            y = cursor + position * 0.62
            self.box(slide, x, y + 0.1, 0.1, 0.1, accent)
            self.text(slide, bullet, x + 0.28, y, w - 0.28, 0.5, 15, body_color)
        stats = page.get("stats") or []
        if stats:
            stat_top = 6.05
            span = w / len(stats)
            for position, stat in enumerate(stats[:3]):
                sx = x + position * span
                self.text(slide, str(stat.get("value", "")), sx, stat_top, span - 0.2, 0.55,
                          26, accent, True)
                self.text(slide, str(stat.get("label", "")), sx, stat_top + 0.5, span - 0.2,
                          0.3, 11, muted)
        if page.get("source"):
            self.text(slide, page["source"], x, 6.85, w, 0.28, 10, muted)

    def wall_slide(self, slide, page):
        """Staggered 2–4 image wall: equal widths, uneven heights, soft depth."""
        items = page["items"]
        count = len(items)
        gap = 0.3
        width = (12.1 - gap * (count - 1)) / count
        offsets = (0.0, 0.62, 0.26, 0.84)
        for index, item in enumerate(items):
            top = 1.92 + offsets[index % len(offsets)]
            height = 6.35 - top - 0.5
            left = 0.6 + index * (width + gap)
            self.artwork(slide, item["image"], left, top, width, height,
                         radius=0.055, shadow=True)
            self.text(slide, item["title"], left, top + height + 0.16, width, 0.3,
                      15, "text", True)
            if item.get("body"):
                self.text(slide, item["body"], left, top + height + 0.46, width, 0.3,
                          11, "muted")


    # ── editorial system: mesh backgrounds, type art, layered compositions ──

    def _dark(self) -> bool:
        return self.meta.get("theme", "midnight") == "midnight"

    def _blend(self, first, second, ratio: float) -> str:
        a, b = media.parse_color(first), media.parse_color(second)
        return media.to_hex(tuple(int(round(x * (1 - ratio) + y * ratio))
                                  for x, y in zip(a, b)))
    def _accent_on(self, light_text: bool) -> str:
        """Accent that survives on a dark image/theme without fixed cyan."""
        if light_text:
            return self._blend(self.colors["accent"], "FFFFFF", .42)
        return self.colors["accent"]


    def _mesh_colors(self) -> list[str]:
        accent, background = self.colors["accent"], self.colors["bg"]
        if self._dark():
            return [background, self._blend(background, accent, 0.14),
                    self._blend(accent, background, 0.70),
                    self._blend(background, "000000", 0.55)]
        return [background, self._blend(background, accent, 0.05),
                self._blend(accent, "FFFFFF", 0.78),
                self._blend(background, "FFFFFF", 0.58)]

    def background_image(self, *, grid: bool = False, seed: int = 7,
                         palette: list[str] | None = None):
        image = media.mesh_gradient((1920, 1080), palette or self._mesh_colors(),
                                    seed=seed, blur=170)
        if grid:
            image = media.rule_grid(image, spacing=88,
                                    color="FFFFFF" if self._dark() else "0B1324",
                                    opacity=0.035)
        image = media.vignette(image, 0.30 if self._dark() else 0.10)
        return media.grain(image, 0.028, seed=seed)

    def add_full(self, slide, picture):
        # Full-bleed art has no transparency; JPEG keeps a 9-page deck small.
        return self._add_picture(slide, picture, 0, 0, 13.333333, 7.5, jpeg=True)

    def _add_picture(self, slide, picture, x, y, w, h, shadow=False, jpeg=False):
        payload = io.BytesIO()
        if jpeg:
            picture.convert("RGB").save(payload, "JPEG", quality=90)
        else:
            picture.save(payload, "PNG")
        payload.seek(0)
        shape = slide.shapes.add_picture(payload, Inches(x), Inches(y),
                                         Inches(w), Inches(h))
        if shadow:
            _soft_shadow(shape)
        return shape

    def type_art(self, slide, text, box, *, mode="gradient", image=None,
                 opacity: float = 1.0, color=None):
        """Place decorative display type (numerals / short words) inside a box."""
        font_size = 520
        if isinstance(image, str):
            image = self._load(image)
        accent = color or self.colors["accent"]
        if mode == "outline":
            art = media.text_art(text, font_size, stroke=max(3, font_size // 90),
                                 stroke_color=accent)
        elif mode == "image" and image is not None:
            art = media.text_art(text, font_size, image=image, outline=2,
                                 outline_color=self.colors["accent"])
        else:
            art = media.text_art(text, font_size,
                                 fill=[accent, self._blend(accent, "FFFFFF", 0.55)])
        if opacity < 1:
            channel = art.getchannel("A").point(lambda value: int(value * opacity))
            art.putalpha(channel)
        x, y, w, h = box
        ratio = min(w / art.width, h / art.height)
        width, height = art.width * ratio, art.height * ratio
        return self._add_picture(slide, art, x + (w - width) / 2,
                                 y + (h - height) / 2, width, height)

    def section_number_slide(self, slide, page):
        dark = self._dark()
        index = page.get("index", 1)
        self.add_full(slide, self.background_image(grid=True, seed=index))
        number = str(page.get("number", f"{index:02d}"))
        art_mode = page.get("art", "gradient")
        self.type_art(slide, number, (-0.55, 1.05, 6.3, 5.45),
                      mode=art_mode,
                      image=page.get("image") if art_mode == "image" else None,
                      opacity=page.get("art_opacity", 0.92))
        text_color = self.colors["text"]
        muted = "AFC3DA" if dark else self.colors["muted"]
        accent = self._accent_on(dark)
        x, w = 6.55, 6.05
        kicker = str(page.get("kicker", f"SECTION {number}")).upper()
        self.text(slide, kicker, x, 1.42, w, 0.3, 12, accent, True, tracking=180)
        self.text(slide, page["title"], x, 1.86, w, 2.0,
                  page.get("title_size", 40), text_color, True, font=self.head_font)
        self.box(slide, x, 4.06, 0.72, 0.05, accent)
        cursor = 4.3
        if page.get("body"):
            self.text(slide, page["body"], x, cursor, w, 0.9, 17, muted)
            cursor += 0.95
        for position, item in enumerate((page.get("items") or [])[:4]):
            y = cursor + position * 0.5
            self.text(slide, f"{position + 1:02d}", x, y, 0.5, 0.26, 11, accent, True)
            self.text(slide, item, x + 0.55, y - 0.02, w - 0.55, 0.34, 15, text_color)
        if page.get("transition"):
            self.text(slide, page["transition"], x, 6.75, w, 0.3, 12, muted)
        total = page.get("sections")
        if isinstance(total, int) and total > 1:
            current = int(page.get("section_no", 1))
            for dot in range(total):
                filled = dot == current - 1
                self.box(slide, 11.35 + dot * 0.28, 6.86, 0.12, 0.12,
                         accent if filled else ("3A4E66" if dark else "D9E2EE"))

    def statement_slide(self, slide, page):
        side = page.get("scrim", "left")
        ink = page.get("ink", "light")
        color = page.get("scrim_color")
        if color is None and ink == "dark":
            color = "F7FBFF"
        self.full_bleed(slide, page["image"], scrim=side, color=color,
                        end=page.get("scrim_end", 0.82),
                        opacity=page.get("scrim_opacity",
                                         0.90 if ink == "dark" else 0.94),
                        vignette=0.15,
                        veil=page.get("veil", 0.42 if ink == "dark" else 0.0))
        text_light = ink == "light"
        title_color = "FFFFFF" if text_light else self.colors["text"]
        muted = "CBD9EA" if text_light else self.colors["muted"]
        accent = self._accent_on(text_light)
        left_text = side == "left"
        if side == "bottom":
            x, w, y = 0.95, 9.6, 3.55
            art_box = (9.55, 1.05, 3.3, 2.55)
        elif left_text:
            x, w, y = 0.95, 5.9, 1.95
            art_box = (-0.65, 1.25, 6.0, 4.95)
        else:
            x, w, y = 7.0, 5.45, 1.95
            art_box = (8.4, 1.25, 6.0, 4.95)
        self.type_art(slide, str(page.get("art_text", page.get("number", ""))),
                      art_box, mode=page.get("art", "gradient"),
                      image=page["image"] if page.get("art") == "image" else None,
                      opacity=page.get("art_opacity", 0.20))
        if page.get("kicker"):
            self.text(slide, str(page["kicker"]).upper(), x, y - 0.52, w, 0.3, 12, accent, True, tracking=160)
            self.box(slide, x, y - 0.16, 0.62, 0.045, accent)
        self.text(slide, page["title"], x, y, w, 2.0,
                  page.get("title_size", 40 if side != "bottom" else 36),
                  title_color, True, font=self.head_font)
        if page.get("body"):
            self.text(slide, page["body"], x, y + 2.2, w, 1.3, 18, muted)
        meta = page.get("meta") or self.meta.get("author", "")
        if meta:
            self.text(slide, meta, x, 6.78, w, 0.32, 12, muted)

    def stack_slide(self, slide, page):
        items = page["items"]
        count = len(items)
        layouts = {
            2: ((0.85, 2.0, 5.5, 3.9), (6.75, 2.4, 5.5, 3.9), (-4, 3)),
            3: ((0.65, 2.05, 4.15, 3.7), (4.6, 1.85, 4.15, 3.95), (8.55, 2.35, 4.15, 3.65),
                (-5, 2, -3)),
            4: ((0.55, 2.1, 3.05, 3.1), (3.85, 1.9, 3.05, 3.35),
                (7.15, 2.2, 3.05, 3.15), (10.45, 1.95, 3.05, 3.3),
                (-6, 2, -4, 3)),
        }
        boxes = layouts[count][:count]
        rotations = layouts[count][-1]
        for index, item in enumerate(items):
            x, y, w, h = boxes[index]
            shape = self.artwork(slide, item["image"], x, y, w, h,
                                 radius=0.035, shadow=True,
                                 treatment=item.get("treatment"),
                                 mask=item.get("mask"))
            shape.rotation = rotations[index]
            if item.get("title"):
                self.text(slide, item["title"], x + 0.05, 6.42,
                          w - 0.1, 0.3, 13, "text", True)
            if item.get("body"):
                self.text(slide, item["body"], x + 0.05, 6.72,
                          w - 0.1, 0.26, 9, "muted")

    def diagonal_slide(self, slide, page):
        left, right = page["images"]
        background = media.diagonal_split(
            self._load(left), self._load(right), (1920, 1080),
            split=page.get("split", 0.58), lean=page.get("lean", 0.16),
            seam=page.get("seam", self.colors["accent"]), seam_width=5)
        background = media.apply_scrim(background, color="0B1224",
                                       direction="bottom", start=0.42, end=1.0,
                                       gamma=1.4, opacity=0.78)
        background = media.grain(background, 0.03)
        self.add_full(slide, background)
        x, w = 0.95, 8.6
        if page.get("kicker"):
            self.text(slide, str(page["kicker"]).upper(), x, 4.02, w, 0.3, 12, self._accent_on(True), True)
        self.text(slide, page["title"], x, 4.35, w, 1.5,
                  page.get("title_size", 38), "FFFFFF", True)
        if page.get("body"):
            self.text(slide, page["body"], x, 5.92, 6.5, 0.55, 16, "CBD9EA")
        labels = page.get("labels") or []
        for index, label in enumerate(labels[:2]):
            lx = 0.95 if index == 0 else 10.0
            self.box(slide, lx, 6.55, 0.34, 0.05, self._accent_on(True))
            self.text(slide, label, lx + 0.48, 6.42, 2.9, 0.3, 12, "FFFFFF", True)

    def glass_slide(self, slide, page):
        ink = page.get("ink", "light" if self._dark() else "dark")
        background = self._bleed_image(page["image"], scrim=None, vignette=0.10,
                                       grain=page.get("grain", 0.03))
        self.add_full(slide, background)
        x, y, w, h = 7.55, 1.35, 5.15, 5.05
        scale = 1920 / 13.333333
        dark = self._dark()
        panel = media.frosted(background, (int(x * scale), int(y * scale),
                                           int((x + w) * scale), int((y + h) * scale)),
                              radius=0.055,
                              tint=page.get("glass_tint",
                                            self.colors["bg"] if dark else "FFFFFF"),
                              opacity=page.get("glass_opacity", 0.34 if dark else 0.22),
                              blur=30, border=0.35)
        self._add_picture(slide, panel, x, y, w, h)
        light = ink == "light"
        title_color = "FFFFFF" if light else self.colors["text"]
        body_color = "DCE8F6" if light else self.colors["text"]
        muted = "AFC3DA" if light else self.colors["muted"]
        accent = self._accent_on(light)
        tx, tw = x + 0.38, w - 0.76
        self.box(slide, tx, y + 0.34, 0.55, 0.045, accent)
        cursor = y + 0.55
        if page.get("kicker"):
            self.text(slide, str(page["kicker"]).upper(), tx, cursor, tw, 0.3, 11, accent, True)
            cursor += 0.38
        self.text(slide, page["title"], tx, cursor, tw, 1.35,
                  page.get("title_size", 29), title_color, True, font=self.head_font)
        cursor += 1.42
        if page.get("body"):
            self.text(slide, page["body"], tx, cursor, tw, 1.15, 16, body_color)
            cursor += 1.2
        stats = page.get("stats") or []
        if stats:
            stat_y = y + h - 1.35
            span = tw / len(stats)
            for index, stat in enumerate(stats[:3]):
                sx = tx + index * span
                self.text(slide, str(stat.get("value", "")), sx, stat_y,
                          span - 0.15, 0.6, 26, accent, True)
                self.text(slide, str(stat.get("label", "")), sx, stat_y + 0.55,
                          span - 0.15, 0.3, 10, muted)
        if page.get("source"):
            self.text(slide, page["source"], x + 0.38, y + h - 0.32, w - 0.76, 0.24,
                      9, muted)

    def quote_slide(self, slide, page):
        dark = self._dark()
        background = self.background_image(grid=page.get("grid", False),
                                           seed=page.get("index", 1))
        accent = self.colors["accent"]
        layer = Image.new("RGBA", background.size, (0, 0, 0, 0))
        font = media.display_font(560)
        if font is not None:
            ImageDraw.Draw(layer).text((70, 20), "“", font=font,
                                       fill=media.parse_color(accent) + (52,))
            background = Image.alpha_composite(background.convert("RGBA"), layer)
        self.add_full(slide, background)
        text_color = self.colors["text"]
        muted = "AFC3DA" if dark else self.colors["muted"]
        self.text(slide, page["quote"], 1.55, 2.05, 10.2, 2.7,
                  page.get("quote_size", 34), text_color, True,
                  align=PP_ALIGN.CENTER, font=self.head_font)
        self.box(slide, 5.95, 5.12, 1.45, 0.04, accent)
        if page.get("by"):
            self.text(slide, page["by"], 1.55, 5.4, 10.2, 0.4, 15, muted,
                      align=PP_ALIGN.CENTER)
        if page.get("title"):
            self.text(slide, page["title"], 1.55, 1.35, 10.2, 0.42, 11,
                      accent, True, align=PP_ALIGN.CENTER)

    def kpi_slide(self, slide, page):
        dark = self._dark()
        background = self.background_image(grid=False, seed=page.get("index", 1))
        background = media.dot_grid(background, spacing=26,
                                    color="FFFFFF" if dark else "0B1324",
                                    opacity=0.07)
        self.add_full(slide, background)
        text_color = self.colors["text"]
        muted = "AFC3DA" if dark else self.colors["muted"]
        accent = self._accent_on(dark)
        self.text(slide, str(page.get("kicker", "")).upper(), 0.95, 1.0, 6.0, 0.3,
                  12, accent, True, tracking=160)
        self.text(slide, page["title"], 0.95, 1.4, 10.6, 1.2,
                  page.get("title_size", 32), text_color, True, font=self.head_font)
        stats = page["stats"]
        top = 3.0
        span = 12.0 / len(stats)
        for index, stat in enumerate(stats):
            x = 0.95 + index * span
            raw = str(stat.get("value", ""))
            size = 64 if len(raw) <= 4 else 50 if len(raw) <= 7 else 42
            self.text(slide, raw, x, top, span - 0.6, 1.15, size, accent, True)
            self.text(slide, str(stat.get("label", "")), x, top + 1.18, span - 0.6,
                      0.34, 14, text_color)
            if stat.get("note"):
                self.text(slide, str(stat["note"]), x, top + 1.58, span - 0.6,
                          0.6, 11, muted)
            share = stat.get("share")
            if isinstance(share, (int, float)):
                ratio = min(1.0, max(0.0, float(share)))
                self.box(slide, x, top + 2.28, span - 0.8, 0.05,
                         "2A3E55" if dark else "D9E2EE")
                self.box(slide, x, top + 2.28,
                         max(0.2, (span - 0.8) * ratio), 0.05, accent)
            if index:
                self.box(slide, x - 0.35, top + 0.05, 0.012, 1.7, "3A4E66" if dark else "D9E2EE")
        if page.get("source"):
            self.text(slide, page["source"], 0.95, 6.65, 11.4, 0.26, 10, muted)

    def transition(self, slide, kind: str):
        element = slide._element
        for old in element.findall(qn("p:transition")):
            element.remove(old)
        transition = element.makeelement(qn("p:transition"), {})
        transition.append(element.makeelement(qn(f"p:{kind}"), {}))
        extension = element.find(qn("p:extLst"))
        if extension is not None:
            extension.addprevious(transition)
        else:
            element.append(transition)


    def _load_alpha(self, filename):
        with Image.open(self.base / filename) as source:
            return ImageOps.exif_transpose(source).convert("RGBA")

    def _set_fill_alpha(self, shape, percent: float) -> None:
        fill = shape.fill._xPr.find(qn("a:solidFill"))
        if fill is None:
            return
        color = fill.find(qn("a:srgbClr"))
        if color is None:
            return
        alpha = color.makeelement(qn("a:alpha"), {"val": str(int(percent * 100000))})
        color.append(alpha)

    def cutout_slide(self, slide, page):
        dark = self._dark()
        index = page.get("index", 1)
        self.add_full(slide, self.background_image(grid=True, seed=index))
        side = page.get("side", "right")
        subject = self._load_alpha(page["image"])
        if page.get("flip"):
            subject = ImageOps.mirror(subject)
        subject = media.defringe(subject, 1)
        box = subject.getbbox()
        if box:
            subject = subject.crop(box)
        height = float(page.get("height", 6.4))
        width = height * subject.width / subject.height
        x = 12.85 - width - (0.6 if side == "right" else 0) if side == "right" else 0.45
        y = 7.55 - height
        # A quiet ring ties the cutout to the accent system.
        ring_x = x + width * 0.02 - 0.35
        ring = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(ring_x), Inches(y + 0.5),
                                      Inches(width + 0.7), Inches(width + 0.7))
        ring.fill.background()
        ring.line.color.rgb = RGBColor.from_string(self.colors["accent"])
        ring.line.width = Pt(1.1)
        ring.shadow.inherit = False
        text_light = True
        title_color = "FFFFFF"
        muted = "B9CCE2"
        accent = self._accent_on(True)
        tx, tw = (0.95, 8.55) if side == "right" else (4.55, 7.9)
        if page.get("kicker"):
            self.text(slide, str(page["kicker"]).upper(), tx, 1.45, tw, 0.3, 12, accent,
                      True, tracking=160)
            self.box(slide, tx, 1.82, 0.62, 0.045, accent)
        self.text(slide, page["title"], tx, 2.02, tw, 2.4,
                  page.get("title_size", 46), title_color, True, font=self.head_font)
        if page.get("body"):
            self.text(slide, page["body"], tx, 4.62, tw - 0.4, 1.3, 18, muted)
        stats = page.get("stats") or []
        for position, stat in enumerate(stats[:3]):
            sx = tx + position * 2.25
            self.text(slide, str(stat.get("value", "")), sx, 6.02, 1.9, 0.55, 28, accent, True)
            self.text(slide, str(stat.get("label", "")), sx, 6.55, 1.9, 0.26, 10, muted)
        subject = media.drop_shadow(subject, blur=26, offset=(0, 16), opacity=0.34)
        self._add_picture(slide, subject, x, y, width, height)

    def annotated_slide(self, slide, page):
        background = self._bleed_image(page["image"], scrim="bottom", end=0.42,
                                       opacity=0.55, vignette=0.22, grain=0.025,
                                       veil=0.14)
        self.add_full(slide, background)
        self.text(slide, str(page.get("kicker", "")).upper(), 0.95, 0.85, 6.0, 0.3,
                  12, self._accent_on(True), True, tracking=160)
        self.text(slide, page["title"], 0.95, 1.2, 9.0, 0.9, 30, "FFFFFF", True,
                  font=self.head_font)
        accent = self.colors["accent"]
        for index, point in enumerate(page["points"], 1):
            px, py = float(point["x"]) * 13.333, float(point["y"]) * 7.5
            dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px - 0.11), Inches(py - 0.11),
                                         Inches(0.22), Inches(0.22))
            dot.fill.solid()
            dot.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
            dot.line.color.rgb = RGBColor.from_string(accent)
            dot.line.width = Pt(2)
            dot.shadow.inherit = False
            self.text(slide, f"{index:02d}", px - 0.09, py - 0.075, 0.2, 0.2, 7, "0B1324", True,
                      align=PP_ALIGN.CENTER)
            right = px < 6.8
            label_w = 3.1
            lx = px + 0.45 if right else px - 0.45 - label_w
            ly = min(max(py - 0.42, 0.55), 6.3)
            line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                              Inches(px), Inches(py),
                                              Inches(lx if right else lx + label_w),
                                              Inches(ly + 0.3))
            line.line.color.rgb = RGBColor.from_string(accent)
            line.line.width = Pt(1.1)
            chip = self.box(slide, lx, ly, label_w, 0.62, "0B1324", rounded=True)
            self._set_fill_alpha(chip, 0.72)
            self.text(slide, point["text"], lx + 0.18, ly + 0.12, label_w - 0.3, 0.42,
                      12 if len(point["text"]) > 10 else 14, "FFFFFF", True)
        if page.get("source"):
            self.text(slide, page["source"], 0.95, 6.85, 11.4, 0.25, 10, "B9CCE2")

    def arch_slide(self, slide, page):
        dark = self._dark()
        background = media.mesh_gradient((1920, 1080), [self.colors["bg"],
                                                        self._blend(self.colors["bg"],
                                                                    self.colors["accent"], 0.10)],
                                         seed=page.get("index", 1), blur=220)
        background = media.grain(background, 0.03)
        self.add_full(slide, background)
        picture = media.cover(self._load(page["image"]), (900, 1100))
        picture = media.arch_mask(picture, radius_ratio=0.5)
        picture = media.drop_shadow(picture, blur=24, offset=(0, 16), opacity=0.45)
        self._add_picture(slide, picture, 0.75, 0.72, 5.15, 6.3)
        text_color = self.colors["text"]
        muted = "AFC3DA" if dark else self.colors["muted"]
        accent = self._accent_on(dark)
        x, w = 6.65, 5.95
        if page.get("kicker"):
            self.text(slide, str(page["kicker"]).upper(), x, 1.55, w, 0.3, 12, accent, True,
                      tracking=160)
        self.text(slide, page["title"], x, 1.95, w, 1.8, page.get("title_size", 36),
                  text_color, True, font=self.head_font)
        self.box(slide, x, 3.92, 0.7, 0.05, accent)
        if page.get("body"):
            self.text(slide, page["body"], x, 4.18, w, 1.4, 17, muted)
        for position, item in enumerate((page.get("items") or [])[:3]):
            self.text(slide, item, x, 5.72 + position * 0.34, w, 0.3, 13, text_color)

    def new(self, page, index):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(self.colors["bg"])
        title = self.text(slide, page.get("title", self.meta["title"]),
                          .6, .38, 12.1, .95, 30, bold=True, font=self.head_font)
        if title is not None:
            title.name = "title"
        if page.get("subtitle"):
            subtitle = self.text(slide, page["subtitle"], .6, 1.35, 12.1, .65, 17, "muted")
            if subtitle is not None:
                subtitle.name = "subtitle"
        footer = self.text(slide, self.meta["title"], .6, 7.06, 10.8, .24, 10, "muted")
        if footer is not None:
            footer.name = "chrome:footer"
        number = self.text(slide, f"{index:02}", 12.0, 7.02, .7, .3, 12, "muted")
        if number is not None:
            number.name = "chrome:page"
        slide.notes_slide.notes_text_frame.text = page.get("notes", "")
        return slide

    def _text_lines(self, value, width, size):
        return sum(max(1, math.ceil(sum(1 if ord(char) > 0x2E80 else .55 for char in line)
                                    / (width * 72 / size)))
                   for line in str(value).split("\n"))

    def _card_need(self, item, width):
        """Estimate the height a card needs for its copy, in inches."""
        body = self._text_lines(item.get("body", ""), width - .5, 20) * 20 * 1.25 / 72
        # title band + top pad (1.15) + bottom band for rule/foot (1.02)
        base = 1.15 + body + 1.02
        return max(2.6, min(4.55, base))

    def card(self, slide, item, x, y, w, h, number=None):
        self.box(slide, x, y, w, h, rounded=True)
        offset = .25
        if number is not None:
            self.text(slide, f"{number:02}", x + .25, y + .25, w - .5, .55, 28, "accent", True)
            offset = 1.0
        self.text(slide, item["title"], x + .25, y + offset, w - .5, .85, 23, "accent", True)
        compact = h < 3.0
        numbered = number is not None
        body_height = h - offset - (1.55 if compact else (1.7 if numbered else 1.9))
        # Four-card pages intentionally use compact cards.  Keep their body
        # copy readable while leaving enough vertical room for two wrapped
        # Chinese lines instead of allowing the strict capacity check to fail.
        body_size = 15 if body_height < 1.0 or w < 2.4 else 20
        # Un-numbered cards centre short copy; a numbered step card keeps the
        # body tight under its title so the number stays part of the header.
        body_y = y + offset + (1.0 if numbered else 0.85)
        self.text(slide, item.get("body", ""), x + .25, body_y,
                  w - .5, body_height, body_size,
                  anchor=None if numbered else MSO_ANCHOR.MIDDLE)
        # A restrained base rule anchors short copy instead of leaving a card
        # that looks like an unfilled template placeholder.
        rule_y = y + h - (0.38 if compact else 0.66)
        self.box(slide, x + .25, rule_y, min(w - .5, 1.0), .045, "accent")
        if item.get("foot") and not compact:
            self.text(slide, item["foot"], x + .25, y + h - .52, w - .5, .26, 11, "muted")

    def compose(self, slide, page):
        layout = page["layout"]
        if layout == "image_hero":
            self.hero_slide(slide, page)
        elif layout == "section_number":
            self.section_number_slide(slide, page)
        elif layout == "image_statement":
            self.statement_slide(slide, page)
        elif layout == "split_diagonal":
            self.diagonal_slide(slide, page)
        elif layout == "glass_overlay":
            self.glass_slide(slide, page)
        elif layout == "quote_stage":
            self.quote_slide(slide, page)
        elif layout == "kpi_showcase":
            self.kpi_slide(slide, page)
        elif layout == "cutout_statement":
            self.cutout_slide(slide, page)
        elif layout == "image_annotated":
            self.annotated_slide(slide, page)
        elif layout == "image_arch":
            self.arch_slide(slide, page)
        elif layout == "image_overlay":
            self.overlay_slide(slide, page)
        elif layout == "image_wall":
            self.wall_slide(slide, page)
        elif layout in {"cover", "closing"}:
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
            takeaway = page.get("takeaway")
            area = 4.55 - (0.55 if takeaway else 0)
            slot = (area - .25 * (rows - 1)) / rows
            # Fit the card to its copy instead of stretching a short card into
            # an empty 4.5-inch shell; the group stays vertically centred.
            height = min(slot, max(self._card_need(item, width) for item in items))
            top = 2.05 + (area - (height * rows + .25 * (rows - 1))) / 2
            for index, item in enumerate(items):
                self.card(slide, item, .6 + (index % columns) * (width + .25),
                          top + (index // columns) * (height + .25), width, height)
            if takeaway:
                self.box(slide, .6, 6.24, .06, .5, "accent")
                self.text(slide, takeaway, .88, 6.22, 11.7, 0.55, 15, "text", True)
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
        elif layout == "architecture":
            items = page["items"]
            # Ink band: the page's one-sentence argument, not another title.
            self.box(slide, 0, 1.32, 13.333, .78, "1C1B19")
            self.text(slide, page.get("band", "三段不是三个模块，是一条状态链"),
                      .65, 1.47, 9.7, .36, 18, "FFFFFF", True,
                      font=self.head_font)
            self.text(slide, "01—03 / 同一任务状态", 9.55, 1.55, 3.15, .22,
                      8, "B98A63", True, align=PP_ALIGN.RIGHT)
            # Technical canvas: a single field, not three cards.
            self.box(slide, .45, 2.05, 12.43, 3.35, "F1E7DB")
            for x in (4.45, 8.55):
                self.box(slide, x, 2.25, .008, 3.0, "D8C6B4")
            for y in (2.25, 5.25):
                self.box(slide, .60, y, 12.1, .008, "D8C6B4")
            for index, item in enumerate(items):
                x = .65 + index * 4.10
                self.box(slide, x, 2.25, 3.40, .030,
                         "B04A2A" if index == 0 else "D8C6B4")
                self.text(slide, item.get("kicker", f"STAGE {index + 1:02d}"),
                          x, 2.36, 1.80, .18, 7.5, "muted", True)
                self.text(slide, f"{index + 1:02d}", x, 2.52, 1.0, .76,
                          40, "accent", True, font=self.head_font)
                self.text(slide, item["title"], x, 3.30, 3.35, .50,
                          19, "text", True)
                self.text(slide, item["body"], x, 3.84, 3.35, 1.15,
                          13, "muted")
                if index < 2:
                    y = 3.52
                    self.box(slide, x + 3.45, y, .50, .008, "B04A2A")
                    tri = slide.shapes.add_shape(
                        MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(x + 3.88),
                        Inches(y - .10), Inches(.18), Inches(.22))
                    tri.rotation = 90
                    tri.fill.solid()
                    tri.fill.fore_color.rgb = RGBColor.from_string("B04A2A")
                    tri.line.fill.background()
            # Evidence write-back loop: a thin, technical return path.
            loop_y = 5.30
            self.box(slide, 10.55, 5.02, .008, .28, "B04A2A")
            self.box(slide, 2.35, loop_y, 8.20, .008, "B04A2A")
            self.box(slide, 2.35, 5.02, .008, .28, "B04A2A")
            tri = slide.shapes.add_shape(
                MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(2.27), Inches(4.94),
                Inches(.18), Inches(.22))
            tri.fill.solid()
            tri.fill.fore_color.rgb = RGBColor.from_string("B04A2A")
            tri.line.fill.background()
            self.text(slide, page.get("loop", "执行证据回写计划"),
                      4.95, 5.06, 4.2, .22, 8, "muted")
            # Validation runway: four grades, one dark strip.
            self.box(slide, .60, 5.58, 12.10, 1.06, "1C1B19")
            self.text(slide, page.get("bottom_title", "验证递进"),
                      .85, 5.72, 1.55, .26, 11, "B98A63", True)
            steps = page.get("steps", ["单项操作", "跨设备流程",
                                       "完整培养周期", "人工对照"])
            for index, step in enumerate(steps[:4]):
                x = 2.55 + index * 2.45
                self.text(slide, f"0{index + 1}", x, 5.70, .38, .18,
                          8, "B98A63", True)
                self.text(slide, step, x, 5.90, 2.20, .24, 11, "FFFFFF")
                if index:
                    self.box(slide, x - .16, 5.70, .006, .62, "5A4A40")
        elif layout == "sequence":
            items = page["items"]
            rail_y = 4.08
            self.box(slide, 1.00, rail_y, 11.35, .012, "B04A2A")
            n = len(items)
            if n == 4:
                positions = [1.70, 4.82, 7.94, 11.06]
                bw, half = 2.95, 1.42
                title_size, body_size, foot_size = 16.5, 12.5, 8.5
            else:
                positions = [2.20, 6.60, 11.00]
                bw, half = 3.25, 1.55
                title_size, body_size, foot_size = 18, 13, 9
            for index, item in enumerate(items):
                x = positions[index]
                above = index % 2 == 0
                self.box(slide, x - .08, rail_y - .08, .16, .16, "B04A2A")
                self.box(slide, x - .006, 3.58 if above else rail_y,
                         .012, .50, "B04A2A")
                sx = x - half
                y = 1.95 if above else 4.55
                self.box(slide, sx - .15, y, .03,
                         1.62 if above else 1.75, "B04A2A")
                self.text(slide, f"{index + 1:02d}", sx, y - .04, .75, .58,
                          30, "accent", True, font=self.head_font)
                self.text(slide, item["title"], sx, y + .50, bw, .40,
                          title_size, "text", True)
                self.text(slide, item["body"], sx, y + .94, bw, .78,
                          11.5 if n == 4 else body_size, "muted")
                if item.get("foot"):
                    self.text(slide, item["foot"], sx, y + 1.80, bw, .22,
                              foot_size, "accent", True)
            tri = slide.shapes.add_shape(
                MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(12.32),
                Inches(rail_y - .10), Inches(.18), Inches(.22))
            tri.rotation = 90
            tri.fill.solid()
            tri.fill.fore_color.rgb = RGBColor.from_string("B04A2A")
            tri.line.fill.background()
            if page.get("bottom"):
                self.text(slide, page["bottom"], .65, 6.55, 12.0, .30,
                          12, "muted")
        elif layout == "pillars":
            items = page["items"]
            n = len(items)
            width = (12.10 - .25 * (n - 1)) / n
            for index, item in enumerate(items):
                x = .60 + index * (width + .25)
                self.box(slide, x, 2.15, width, 1.05, "1C1B19")
                self.text(slide, f"{index + 1:02d}", x + .18, 2.28, .55, .22,
                          9, "B98A63", True)
                self.text(slide, item["title"], x + .18, 2.56, width - .36, .55,
                          16 if n == 4 else 18, "FFFFFF", True)
                self.text(slide, item["body"], x + .18, 3.42, width - .36,
                          1.40 if n == 4 else 2.10,
                          12.5 if n == 4 else 14, "muted")
                if item.get("foot"):
                    self.text(slide, item["foot"], x + .18,
                              4.82 if n == 4 else 5.62, width - .36, .25,
                              8.5 if n == 4 else 9, "accent", True)
                self.box(slide, x, 6.02, width, .012, "D8C6B4")
            if page.get("bottom"):
                if n == 4:
                    self.box(slide, .65, 4.82, .8, .03, "accent")
                    self.text(slide, page["bottom"], .65, 5.05, 11.8, .80,
                              18, "text", True, font=self.head_font)
                else:
                    self.text(slide, page["bottom"], .65, 6.28, 12.0, .38,
                              13, "muted")
        elif layout == "image_band":
            items = page["items"]
            self.artwork(slide, page["image"], 0, 0, 13.333, 4.55, radius=0.0,
                         scrim={"direction": "bottom", "color": "0B1324",
                                "start": .28, "end": 1.0, "opacity": .94})
            if page.get("kicker"):
                self.text(slide, str(page["kicker"]).upper(), .65, 1.05, 6.0, .26,
                          9.5, "B98A63", True)
            self.text(slide, page["title"], .65, 1.44, 8.6, 1.16,
                      page.get("title_size", 30), "FFFFFF", True, font=self.head_font)
            if page.get("body"):
                self.text(slide, page["body"], .65, 2.78, 7.4, .78, 13.5, "DCE8F6")
            self.box(slide, .60, 4.55, 12.1, .012, "B04A2A")
            for index, item in enumerate(items):
                x = .65 + index * 4.05
                self.text(slide, f"{index + 1:02d}", x, 4.74, .55, .24,
                          9, "accent", True)
                self.text(slide, item["title"], x, 4.98, 3.65, .36,
                          16.5, "text", True)
                self.text(slide, item["body"], x, 5.40, 3.65, .86,
                          12.5, "muted")
                if item.get("value"):
                    self.text(slide, str(item["value"]), x, 6.28, 1.9, .44,
                              20, "accent", True, font=self.head_font)
                    if item.get("label"):
                        self.text(slide, str(item["label"]), x + 1.72, 6.38,
                                  1.8, .26, 8.5, "muted")
            if page.get("source"):
                self.text(slide, page["source"], .65, 6.95, 11.9, .24, 8.5, "muted")
        elif layout in {"image_panel", "image_panel_strip"}:
            side = page.get("side", "left")
            dense = bool(page.get("dense"))
            image_w = 5.15 if dense else 6.50
            panel_w = 13.333 - image_w
            if side == "left":
                image_x, panel_x = 0.0, image_w
            else:
                image_x, panel_x = panel_w, 0.0
            self.artwork(slide, page["image"], image_x, 0, image_w, 7.5, radius=0.0)
            if page.get("inset"):
                iw, ih = 1.95, 1.32
                ix = image_x + image_w - iw - .24
                iy = 7.50 - ih - .24
                self.box(slide, ix - .06, iy - .06, iw + .12, ih + .12, "FBF6EE")
                self.artwork(slide, page["inset"], ix, iy, iw, ih, radius=0.0)
            self.box(slide, panel_x, 0, panel_w, 7.5, "FBF6EE")
            seam = panel_x if side == "left" else image_x - .012
            self.box(slide, seam, 0, .012, 7.5, "B04A2A")
            px = panel_x + .52
            pw = panel_w - 1.04
            if page.get("kicker"):
                self.text(slide, str(page["kicker"]).upper(), px, .62, pw, .24,
                          9.5, "accent", True)
            self.text(slide, page["title"], px, 1.02, pw, 1.40,
                      page.get("title_size", 26 if dense else 28), "text", True,
                      font=self.head_font)
            cursor = 2.42
            if page.get("body"):
                self.text(slide, page["body"], px, cursor, pw,
                          .92 if not dense else 1.02,
                          14 if dense else 15, "muted")
                cursor += 1.05 if not dense else 1.18
            for index, bullet in enumerate(page.get("bullets") or []):
                step = .40 if dense else .62
                y = cursor + index * step
                self.box(slide, px, y + .10, .09, .09, "accent")
                self.text(slide, bullet, px + .24, y, pw - .24,
                          .34 if dense else .50, 12.0 if dense else 13.5, "text")
            stats = page.get("stats") or []
            if stats:
                stat_y = 5.58 if dense else 5.88
                span = pw / len(stats)
                for index, stat in enumerate(stats):
                    sx = px + index * span
                    self.text(slide, str(stat.get("value", "")), sx, stat_y,
                              span - .2, .48, 21 if dense else 24, "accent", True,
                              font=self.head_font)
                    self.text(slide, str(stat.get("label", "")), sx, stat_y + .46,
                              span - .2, .32, 9, "muted")
                    if index:
                        self.box(slide, sx - .16, stat_y, .008, .72, "D8C6B4")
            if page.get("strip"):
                self.box(slide, px, 6.35, pw, .010, "D8C6B4")
                labels = page["strip"][:3]
                span = pw / len(labels)
                for index, label in enumerate(labels):
                    self.text(slide, label, px + index * span, 6.45,
                              span - .12, .30, 8.5, "muted")
            if page.get("source"):
                self.text(slide, page["source"], px, 6.95, pw, .24, 8.5, "muted")
        elif layout == "image_split":
            self.artwork(slide, page["image"], .6, 2.05, 7.1, 4.55,
                         radius=.05, shadow=True)
            self.box(slide, 8.0, 2.05, 4.7, 4.55, rounded=True)
            self.text(slide, page.get("body", ""), 8.3, 2.4, 4.1, 3.8, 22)
        elif layout == "image_grid":
            items = page["items"]
            if len(items) == 3:
                # Three equal tiles across, flat and square-cornered: one clean row.
                gap, top = .3, 2.02
                width = (12.1 - gap * 2) / 3
                for index, item in enumerate(items):
                    x = .6 + index * (width + gap)
                    self.artwork(slide, item["image"], x, top, width, 3.70, radius=0.0)
                    self.text(slide, item["title"], x, 5.84, width, .34, 17, "accent", True)
                    if item.get("body"):
                        self.text(slide, item["body"], x, 6.22, width, .30, 12, "muted")
            else:
                rows = math.ceil(len(items) / 2)
                height = (4.55 - .2 * (rows - 1)) / rows
                for index, item in enumerate(items):
                    x, y = .6 + (index % 2) * 6.2, 2.05 + (index // 2) * (height + .2)
                    self.artwork(slide, item["image"], x, y, 5.9, height - .62,
                                 radius=.04, shadow=True)
                    self.text(slide, item["title"], x, y + height - .48, 5.9, .45, 18, "accent", True)
        elif layout == "heat_bars":
            items = page["items"]
            fire_colors = {"大火": "8C2F22", "中火": "B4602A",
                           "小火": "E2A35E", "微火": "A9744F"}
            levels = {"大火": 4, "中火": 3, "小火": 2, "微火": 1}
            total = float(page.get("total_minutes", 90))
            x0, x1 = 5.45, 11.95
            span = x1 - x0
            top = 2.52
            row_h = .82
            for i in range(len(items)):
                if i % 2 == 0:
                    self.box(slide, .55, top + .04 + i * row_h, 11.9, row_h - .02, "F6EEE4")
            for tick in range(0, int(total), 15):
                x = x0 + span * tick / total
                self.box(slide, x, top - .10, .012, 4.10, "E4D2C0")
                self.text(slide, str(tick), x - .18, top - .38, .36, .20, 8, "muted",
                          align=PP_ALIGN.CENTER)
            self.text(slide, "分钟", x1 + .02, top - .38, .5, .20, 8, "muted")
            starts = []
            cursor = 0.0
            for item in items:
                starts.append(cursor)
                cursor += float(item["minutes"])
            for i, item in enumerate(items):
                y = top + .14 + i * row_h
                fire = item["fire"]
                level = int(item.get("level", levels.get(fire, 2)))
                self.text(slide, item.get("number", f"{i + 1:02d}"), .78, y - .02,
                          .55, .30, 14, "accent", True)
                self.text(slide, item["title"], 1.34, y - .04, 2.6, .40, 20, "text", True)
                duration = item.get("time", f"{int(float(item['minutes']))} min")
                self.text(slide, f"{fire} · {duration}", 1.36, y + .36, 2.3, .22, 10, "muted")
                self.box(slide, .62, y + .04, .035, .54, fire_colors[fire])
                minutes = float(item["minutes"])
                bar_w = max(.08, span * minutes / total)
                bar_x = x0 + span * starts[i] / total
                by = y + .18
                self.box(slide, bar_x, by, bar_w, .20, fire_colors[fire], rounded=True)
                body_x = bar_x
                body_w = x1 - bar_x
                if body_w < 2.6:
                    body_x = max(x0, bar_x - 2.8)
                    body_w = 2.8
                self.box(slide, bar_x, by - .04, .012, .28, "text")
                self.text(slide, item["body"], body_x, y + .43, body_w, .26, 10, "muted")
            if page.get("source"):
                self.text(slide, page["source"], .6, 6.70, 11.8, .20, 9, "muted")
        elif layout == "ingredient_split":
            items = page["items"]
            self.artwork(slide, page["image"], 5.55, 0, 7.78, 7.5, radius=0.0)
            self.box(slide, 0, 0, 5.55, 7.5, "140C08")
            self.box(slide, 5.54, 0, .012, 7.5, "B98A63")
            self.text(slide, page.get("kicker", "备料 / MISE EN PLACE"), .65, .55,
                      4.5, .22, 9, "B98A63", True)
            self.text(slide, "800", .60, 1.00, 2.6, 1.16, 64, "FFFFFF", True,
                      font=self.head_font)
            self.text(slide, "g", 3.08, 1.48, .5, .4, 20, "B98A63", True)
            self.text(slide, page.get("fish_label", "草鱼（或黑鱼 750 g）"), .65, 2.18,
                      4.35, .25, 9.5, "E8D5C4")
            self.text(slide, "300", .60, 2.54, 2.6, 1.16, 64, "FFFFFF", True,
                      font=self.head_font)
            self.text(slide, "g", 3.08, 3.02, .5, .4, 20, "B98A63", True)
            self.text(slide, page.get("acid_label", "老坛酸菜"), .65, 3.72,
                      4.35, .25, 9.5, "E8D5C4")
            self.text(slide, "8 : 3", 4.15, 1.10, 1.1, .34, 18, "B98A63", True,
                      align=PP_ALIGN.RIGHT)
            self.text(slide, "鱼 : 酸菜", 4.15, 1.42, 1.1, .20, 8, "E8D5C4",
                      align=PP_ALIGN.RIGHT)
            for i, item in enumerate(items):
                y = 4.05 + i * .38
                self.text(slide, item["label"], .65, y, .95, .22, 8, "B98A63", True)
                self.text(slide, item["value"], 1.65, y - .02, 3.55, .30, 9.5, "F4E9DC")
            if page.get("source"):
                self.text(slide, page["source"], .65, 6.52, 4.55, .52, 8.5, "C9B3A0")
        elif layout == "state_gallery":
            items = page["items"]
            self.box(slide, .60, 2.00, 7.15, 4.45, "B04A2A")
            self.artwork(slide, items[0]["image"], .70, 2.10, 6.95, 4.25, radius=0.0)
            self.text(slide, "01", .60, 6.48, .7, .40, 22, "accent", True)
            self.text(slide, items[0]["title"], 1.28, 6.50, 6.4, .32, 17, "text", True)
            self.text(slide, items[0]["body"], 1.28, 6.78, 6.4, .22, 11, "muted")
            for idx, (y, item) in enumerate(((2.00, items[1]), (4.52, items[2])), start=1):
                self.box(slide, 8.00, y, 4.70, 2.00, "B04A2A")
                self.artwork(slide, item["image"], 8.06, y + .06, 4.58, 1.88, radius=0.0)
                self.text(slide, f"0{idx+1}", 8.00, y + 2.04, .55, .28, 13, "accent", True)
                self.text(slide, item["title"], 8.52, y + 2.02, 4.0, .26, 12.5, "text", True)
                self.text(slide, item["body"], 8.52, y + 2.26, 4.0, .18, 8.8, "muted")
        elif layout == "state_band":
            items = page["items"]
            self.box(slide, 0, 0, 13.333, 1.18, "FBF6EE")
            self.text(slide, page["title"], .6, .28, 12.1, .52, 26, "text", True,
                      font=self.head_font)
            if page.get("subtitle"):
                self.text(slide, page["subtitle"], .6, .82, 12.1, .26, 13, "muted")
            top, bottom = 1.18, 7.02
            h2 = (bottom - top - .04) / 2
            self.artwork(slide, items[0]["image"], 0, top, 8.2, bottom - top, radius=0.0)
            self.artwork(slide, items[1]["image"], 8.2, top, 5.133, h2, radius=0.0)
            self.artwork(slide, items[2]["image"], 8.2, top + h2 + .04, 5.133, h2, radius=0.0)
            # left caption band
            self.box(slide, 0, 6.02, 8.2, 1.0, "FBF6EE")
            self.box(slide, 0, 6.02, 8.2, .018, "B04A2A")
            self.text(slide, "01", .55, 6.12, .7, .38, 20, "accent", True)
            self.text(slide, items[0]["title"], 1.28, 6.12, 6.5, .34, 18, "text", True)
            self.text(slide, items[0]["body"], 1.28, 6.46, 6.5, .26, 11.5, "muted")
            # right top band
            self.box(slide, 8.2, 3.30, 5.133, .68, "FBF6EE")
            self.box(slide, 8.2, 3.30, 5.133, .016, "B04A2A")
            self.text(slide, "02", 8.42, 3.38, .55, .28, 13, "accent", True)
            self.text(slide, items[1]["title"], 9.02, 3.36, 4.0, .26, 12.5, "text", True)
            self.text(slide, items[1]["body"], 9.02, 3.60, 4.0, .22, 9.5, "muted")
            # right bottom band
            self.box(slide, 8.2, 6.34, 5.133, .68, "FBF6EE")
            self.box(slide, 8.2, 6.34, 5.133, .016, "B04A2A")
            self.text(slide, "03", 8.42, 6.42, .55, .28, 13, "accent", True)
            self.text(slide, items[2]["title"], 9.02, 6.40, 4.0, .26, 12.5, "text", True)
            self.text(slide, items[2]["body"], 9.02, 6.64, 4.0, .22, 9.5, "muted")
        elif layout == "state_triptych":
            items = page["items"]
            self.box(slide, 0, 0, 13.333, 1.32, "FBF6EE")
            self.text(slide, page["title"], .6, .36, 12.1, .55, 28, "text", True,
                      font=self.head_font)
            if page.get("subtitle"):
                self.text(slide, page["subtitle"], .6, .94, 12.1, .28, 14, "muted")
            top = 1.32
            bottom = 7.02
            h = (bottom - top) / 2 - .02
            scrim = {"direction": "bottom", "color": "120B08", "start": .42,
                     "end": .98, "opacity": .88}
            self.artwork(slide, items[0]["image"], 0, top, 8.2, bottom - top,
                         radius=0.0, scrim=scrim)
            self.artwork(slide, items[1]["image"], 8.2, top, 5.133, h,
                         radius=0.0, scrim=scrim)
            self.artwork(slide, items[2]["image"], 8.2, top + h + .04, 5.133, h,
                         radius=0.0, scrim=scrim)
            self.text(slide, "01", .55, top + .28, .8, .44, 24, "FFFFFF", True)
            self.text(slide, items[0]["title"], .55, bottom - 1.00, 7.0, .38, 20, "FFFFFF", True)
            self.text(slide, items[0]["body"], .55, bottom - .58, 7.0, .24, 12, "FFF6EC")
            self.text(slide, "02", 8.65, top + .26, .8, .36, 20, "FFFFFF", True)
            self.text(slide, items[1]["title"], 8.65, top + h - .60, 4.3, .28, 14, "FFFFFF", True)
            self.text(slide, items[1]["body"], 8.65, top + h - .32, 4.3, .22, 10.5, "FFF6EC")
            y2 = top + h + .04
            self.text(slide, "03", 8.65, y2 + .26, .8, .36, 20, "FFFFFF", True)
            self.text(slide, items[2]["title"], 8.65, bottom - .66, 4.3, .28, 14, "FFFFFF", True)
            self.text(slide, items[2]["body"], 8.65, bottom - .38, 4.3, .22, 10.5, "FFF6EC")
        elif layout == "ingredient_board":
            items = page["items"]
            top = 2.02
            self.artwork(slide, page["image"], .6, top, 7.2, 4.86, radius=0.0)
            self.text(slide, page.get("fish_value", "800 g"), 8.0, 1.98, 2.25, .55,
                      30, "accent", True, font=self.head_font)
            self.text(slide, page.get("fish_label", "草鱼"), 8.0, 2.52, 2.2, .22,
                      10.5, "muted")
            self.text(slide, page.get("acid_value", "300 g"), 10.3, 1.98, 2.25, .55,
                      30, "accent", True, font=self.head_font)
            self.text(slide, page.get("acid_label", "酸菜"), 10.3, 2.52, 2.2, .22,
                      10.5, "muted")
            self.box(slide, 8.0, 2.82, 4.65, .012, "D8C4B2")
            for i, item in enumerate(items):
                y = 2.98 + i * .54
                self.text(slide, item["label"], 8.0, y, 1.0, .22, 9.5, "muted", True)
                self.text(slide, item["value"], 9.05, y - .03, 3.55, .42, 11.5, "text")
                if i < len(items) - 1:
                    self.box(slide, 8.0, y + .46, 4.65, .008, "E7D8C8")
            if page.get("source"):
                self.text(slide, page["source"], 8.0, 6.25, 4.65, .55, 9.5, "muted")
        elif layout == "prep_stove":
            items = page["items"]
            total = float(page.get("total_minutes", 40))
            x0, x1 = 2.85, 12.55
            span = x1 - x0
            for tick in (0, 10, 20, 30, 40):
                x = x0 + span * tick / total
                self.box(slide, x, 2.16, .008, 3.45, "E6D5C5")
                self.text(slide, str(tick), x - .14, 1.92, .28, .18, 8, "muted",
                          align=PP_ALIGN.CENTER)
            self.text(slide, "分钟", x1 + .02, 1.92, .4, .18, 8, "muted")
            for lane, y, label, hint, color in (
                    ("prep", 2.42, "提前备料", "可提前做", "B98A63"),
                    ("stove", 4.05, "灶前", "连续操作", "B04A2A")):
                entries = [item for item in items if item.get("lane") == lane]
                self.text(slide, label, .6, y - .04, 2.1, .30, 14, "text", True)
                self.text(slide, hint, .6, y + .28, 2.1, .20, 9.5, "muted")
                offset = 16.0 if lane == "stove" else 0.0
                cursor = 0.0
                for item in entries:
                    minutes = float(item["minutes"])
                    bx = x0 + span * (offset + cursor) / total
                    bw = max(.55, span * minutes / total - .05)
                    self.box(slide, bx + .025, y, bw, .92, color, rounded=True)
                    title_size = 11 if bw > 1.15 else (9.5 if bw > .8 else 9)
                    self.text(slide, item["title"], bx + .10, y + .08, bw - .16, .26,
                              title_size, "FFFFFF", True)
                    self.text(slide, f"{minutes:g} 分钟", bx + .10, y + .32, bw - .16,
                              .20, 8.5, "FFF6EC")
                    self.text(slide, item["body"], bx + .10, y + .52, bw - .16, .36,
                              8.2, "FFF6EC")
                    cursor += minutes
            xh = x0 + span * 16 / total
            self.box(slide, xh, 2.16, .012, 3.45, "B04A2A")
            self.text(slide, "开火", xh - .18, 1.92, .5, .18, 8.5, "accent", True,
                      align=PP_ALIGN.CENTER)
            self.text(slide, "16", .6, 5.62, .75, .55, 28, "accent", True)
            self.text(slide, "分钟 · 备料", 1.35, 5.76, 1.8, .22, 10, "muted")
            self.text(slide, "24", 3.4, 5.62, .75, .55, 28, "accent", True)
            self.text(slide, "分钟 · 灶前", 4.15, 5.76, 1.8, .22, 10, "muted")
            self.text(slide, "40", 6.2, 5.62, .75, .55, 28, "accent", True)
            self.text(slide, "分钟 · 上桌", 6.95, 5.76, 1.8, .22, 10, "muted")
        elif layout == "image_triptych":
            items = page["items"]
            top = 2.06
            self.artwork(slide, items[0]["image"], .6, top, 7.55, 4.38, radius=0.0)
            self.artwork(slide, items[1]["image"], 8.40, top, 4.30, 1.98, radius=0.0)
            self.artwork(slide, items[2]["image"], 8.40, 4.60, 4.30, 1.84, radius=0.0)
            self.box(slide, 8.24, top, .012, 4.38, "D8C6B4")
            self.text(slide, f"01  {items[0]['title']}", .6, 6.46, 7.55, .31, 17, "accent", True)
            self.text(slide, items[0]["body"], .6, 6.80, 7.55, .22, 12, "muted")
            self.text(slide, f"02  {items[1]['title']}", 8.40, 4.08, 4.30, .26, 14, "accent", True)
            self.text(slide, items[1]["body"], 8.40, 4.32, 4.30, .20, 11, "muted")
            self.text(slide, f"03  {items[2]['title']}", 8.40, 6.48, 4.30, .26, 14, "accent", True)
            self.text(slide, items[2]["body"], 8.40, 6.74, 4.30, .20, 11, "muted")
        elif layout in {"chart", "chart_dashboard"}:
            top = 2.05 if page.get("subtitle") else 1.65
            self.chart_page(slide, page, top=top, height=6.6 - top)
        elif layout == "table":
            headers, rows = page["headers"], page["rows"]
            table = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(.6), Inches(2.1),
                                           Inches(12.1), Inches(4.4)).table
            weights = [max([len(str(headers[index]))] +
                           [len(str(row[index])) for row in rows])
                       for index in range(len(headers))]
            total = sum(weights) or 1
            for index, weight in enumerate(weights):
                table.columns[index].width = Inches(12.1 * weight / total)
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
        if page.get("source") and layout not in {"glass_overlay", "kpi_showcase",
                                                 "image_annotated", "ingredient_board",
                                                 "ingredient_split", "image_panel"}:
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
        page.setdefault("index", index)
        if page["layout"] in {"image_hero", "image_overlay", "section_number",
                              "image_statement", "split_diagonal", "glass_overlay",
                              "quote_stage", "kpi_showcase", "cutout_statement",
                              "image_annotated", "image_arch"}:
            slide = builder.blank(page, index)
        else:
            slide = builder.new(page, index)
        # Stable page identity lives in the slide name: it survives reordering
        # and gives the edit ledger something position-independent to match on.
        slide._element.cSld.set("name", page.get("id", f"P{index:02d}"))
        builder.compose(slide, page)
        effect = spec["meta"].get("transition")
        if effect:
            builder.transition(slide, effect)
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
