#!/usr/bin/env python3
"""Deck-ready image compositing primitives.

Everything is baked with Pillow/NumPy so a slide renders identically in
PowerPoint, LibreOffice and an exported PDF, with no font- or renderer-specific
gradient behaviour. The functions here are deliberately small and pure: they
take an image and return an image, so ``build_deck.py`` stays about layout.
"""
from __future__ import annotations

import math
from pathlib import Path
from typing import Iterable, Sequence

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageOps

RGB = tuple[int, int, int]


def parse_color(value) -> RGB:
    """Accept ``#RRGGBB``, ``RRGGBB`` or an ``(r, g, b)`` triple."""
    if isinstance(value, (tuple, list)) and len(value) == 3:
        return tuple(int(max(0, min(255, c))) for c in value)  # type: ignore[return-value]
    text = str(value).strip().lstrip("#")
    if len(text) != 6:
        raise ValueError(f"unsupported color: {value!r}")
    return int(text[0:2], 16), int(text[2:4], 16), int(text[4:6], 16)


def to_hex(color: Iterable[int]) -> str:
    return "".join(f"{max(0, min(255, int(c))):02X}" for c in color)


def cover(image: Image.Image, size: tuple[int, int],
          centering: tuple[float, float] = (0.5, 0.5)) -> Image.Image:
    """Crop-to-fill ``size`` without distortion (the only legal image fit)."""
    width, height = size
    return ImageOps.fit(image.convert("RGB"), (width, height), method=Image.LANCZOS,
                        centering=centering)


def sample_ink(image: Image.Image, percentile: float = 12.0,
               mix: float = 0.0, mix_color: str = "0B1324") -> tuple[int, int, int]:
    """Return the dominant dark tone of ``image``.

    Gradient scrims should continue the picture's own palette instead of a dead
    black; sampling the darkest band is stable and cheap. ``mix`` blends that
    tone toward a fixed ink color (used to keep a deck's theme consistent).
    """
    small = image.convert("RGB").resize((160, max(1, round(160 * image.height / max(1, image.width))))
                                        , Image.BOX)
    data = np.asarray(small, dtype=np.float32)
    luminance = data @ np.array([0.299, 0.587, 0.114], dtype=np.float32)
    threshold = np.percentile(luminance, percentile)
    darkest = data[luminance <= threshold]
    tone = darkest.mean(axis=0) if len(darkest) else data.reshape(-1, 3).mean(axis=0)
    if mix:
        blend = np.array(parse_color(mix_color), dtype=np.float32)
        tone = tone * (1 - mix) + blend * mix
    return tuple(int(round(c)) for c in tone)  # type: ignore[return-value]


def _axis_alpha(size: tuple[int, int], direction: str, start: float, end: float,
                gamma: float) -> np.ndarray:
    width, height = size
    if direction in {"left", "right"}:
        axis = np.linspace(0.0, 1.0, width, dtype=np.float32)
        if direction == "right":
            axis = 1.0 - axis
        grid = np.tile(axis, (height, 1))
    elif direction in {"top", "bottom"}:
        axis = np.linspace(0.0, 1.0, height, dtype=np.float32)
        # alpha is 1 where the axis is 0, so "bottom" must measure from the
        # bottom edge: that is what makes a photo fade *up* into the page.
        if direction == "bottom":
            axis = 1.0 - axis
        grid = np.tile(axis[:, None], (1, width))
    elif direction == "diagonal":
        # Dark toward the lower-left corner: a safe zone for left-aligned text.
        x = np.linspace(0.0, 1.0, width, dtype=np.float32)
        y = np.linspace(0.0, 1.0, height, dtype=np.float32)
        grid = np.maximum(np.tile(x, (height, 1)), np.tile(1.0 - y[:, None], (1, width)))
    else:
        raise ValueError(f"unsupported gradient direction: {direction}")
    span = max(1e-6, end - start)
    alpha = np.clip((end - grid) / span, 0.0, 1.0) ** gamma
    return alpha.astype(np.float32)


def apply_scrim(image: Image.Image, color="0B1324", direction: str = "left",
                start: float = 0.0, end: float = 0.78, gamma: float = 1.15,
                opacity: float = 0.92) -> Image.Image:
    """Composite a linear-color scrim used to keep overlaid text legible."""
    base = image.convert("RGBA")
    alpha = _axis_alpha(base.size, direction, start, end, gamma) * float(np.clip(opacity, 0, 1))
    layer = np.asarray(base, dtype=np.float32).copy()
    tone = np.array(parse_color(color), dtype=np.float32)
    layer[..., :3] = layer[..., :3] * (1 - alpha[..., None]) + tone * alpha[..., None]
    return Image.fromarray(np.clip(layer, 0, 255).astype(np.uint8), "RGBA")


def veil(image: Image.Image, color="FFFFFF", opacity: float = 0.4) -> Image.Image:
    """A uniform translucent wash; keeps dark text legible on any picture."""
    base = image.convert("RGBA")
    layer = np.asarray(base, dtype=np.float32).copy()
    tone = np.array(parse_color(color), dtype=np.float32)
    alpha = float(np.clip(opacity, 0, 1))
    layer[..., :3] = layer[..., :3] * (1 - alpha) + tone * alpha
    return Image.fromarray(np.clip(layer, 0, 255).astype(np.uint8), "RGBA")


def blend_edge(image: Image.Image, color, edge: str = "right",
               depth: float = 0.38, gamma: float = 1.6,
               opacity: float = 1.0) -> Image.Image:
    """Melt one edge of a picture into a flat background color.

    This is what makes a photo look designed into the page instead of pasted on:
    the picture bleeds off one side and its far edge dissolves into the theme.
    """
    return apply_scrim(image, color=color, direction=edge,
                       start=0.0, end=max(0.02, min(1.0, depth)),
                       gamma=gamma, opacity=opacity)


def vignette(image: Image.Image, strength: float = 0.32,
             color="0B1220") -> Image.Image:
    """Subtle radial darkening; keeps a full-bleed photo from looking flat."""
    base = image.convert("RGBA")
    width, height = base.size
    y, x = np.ogrid[:height, :width]
    cx, cy = width / 2.0, height / 2.0
    distance = np.sqrt(((x - cx) / cx) ** 2 + ((y - cy) / cy) ** 2) / math.sqrt(2)
    alpha = np.clip((distance - 0.35) / 0.65, 0, 1) ** 1.5 * float(np.clip(strength, 0, 1))
    layer = np.asarray(base, dtype=np.float32).copy()
    tone = np.array(parse_color(color), dtype=np.float32)
    layer[..., :3] = layer[..., :3] * (1 - alpha[..., None]) + tone * alpha[..., None]
    return Image.fromarray(np.clip(layer, 0, 255).astype(np.uint8), "RGBA")


def rounded(image: Image.Image, radius_ratio: float = 0.06,
            supersample: int = 3) -> Image.Image:
    """Round the corners with a supersampled mask (clean antialiasing)."""
    base = image.convert("RGBA")
    width, height = base.size
    radius = max(0, int(round(min(width, height) * radius_ratio)))
    if radius == 0:
        return base
    mask = Image.new("L", (width * supersample, height * supersample), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, width * supersample - 1, height * supersample - 1],
        radius=radius * supersample, fill=255)
    mask = mask.resize((width, height), Image.LANCZOS)
    base.putalpha(mask)
    return base


def duotone(image: Image.Image, dark="0B1324", light="F7FBFF",
            strength: float = 1.0) -> Image.Image:
    """Map luminance onto two brand colors; use for section art, not evidence."""
    gray = ImageOps.grayscale(image).convert("L")
    data = np.asarray(gray, dtype=np.float32) / 255.0
    lo = np.array(parse_color(dark), dtype=np.float32)
    hi = np.array(parse_color(light), dtype=np.float32)
    mapped = lo[None, None, :] * (1 - data[..., None]) + hi[None, None, :] * data[..., None]
    original = np.asarray(image.convert("RGB"), dtype=np.float32)
    mixed = original * (1 - float(np.clip(strength, 0, 1))) + mapped * float(np.clip(strength, 0, 1))
    return Image.fromarray(np.clip(mixed, 0, 255).astype(np.uint8), "RGB")


def blur(image: Image.Image, radius: float = 12.0) -> Image.Image:
    return image.convert("RGB").filter(ImageFilter.GaussianBlur(radius))


def strip_color_cast(image: Image.Image, keep: float = 0.35) -> Image.Image:
    """Reduce saturation while keeping a hint of the original color."""
    base = image.convert("RGB")
    gray = ImageOps.grayscale(base).convert("RGB")
    return Image.blend(gray, base, float(np.clip(keep, 0, 1)))
