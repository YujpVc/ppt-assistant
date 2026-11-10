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


# ── editorial backgrounds and type art ────────────────────────────────────────

def display_font(size: int = 400, bold: bool = True):
    """Load a heavy CJK display face for raster type; None when unavailable."""
    candidates = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
    ]
    from PIL import ImageFont
    for path in candidates:
        if Path(path).is_file():
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return None


def mesh_gradient(size: tuple[int, int], colors: Sequence[str],
                  seed: int = 7, blur: float = 150.0) -> Image.Image:
    """A soft multi-color field: the premium alternative to a flat background."""
    import random
    width, height = size
    rng = random.Random(seed)
    base = Image.new("RGB", size, parse_color(colors[0]))
    layer = Image.new("RGB", size, parse_color(colors[0]))
    draw = ImageDraw.Draw(layer)
    palette = list(colors)
    for index in range(7):
        color = parse_color(palette[index % len(palette)])
        blob_w = rng.randint(int(width * 0.45), int(width * 1.1))
        blob_h = rng.randint(int(height * 0.45), int(height * 1.1))
        left = rng.randint(-int(width * 0.25), width - int(blob_w * 0.3))
        top = rng.randint(-int(height * 0.25), height - int(blob_h * 0.3))
        draw.ellipse([left, top, left + blob_w, top + blob_h], fill=color)
    layer = layer.filter(ImageFilter.GaussianBlur(blur))
    return Image.blend(base, layer, 0.88)


def grain(image: Image.Image, amount: float = 0.045, seed: int = 3) -> Image.Image:
    """Fine film grain; keeps large color fields from looking flat."""
    rng = np.random.default_rng(seed)
    base = np.asarray(image.convert("RGB"), dtype=np.float32)
    noise = rng.normal(0.0, 255.0 * float(np.clip(amount, 0, 0.2)), base.shape[:2])
    base = np.clip(base + noise[..., None], 0, 255)
    return Image.fromarray(base.astype(np.uint8), "RGB")


def dot_grid(image: Image.Image, spacing: int = 28, color="FFFFFF",
             opacity: float = 0.10) -> Image.Image:
    """A quiet dot matrix behind content pages."""
    base = image.convert("RGBA")
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    tone = parse_color(color) + (int(255 * np.clip(opacity, 0, 1)),)
    for y in range(spacing // 2, base.height, spacing):
        for x in range(spacing // 2, base.width, spacing):
            draw.ellipse([x, y, x + 2, y + 2], fill=tone)
    return Image.alpha_composite(base, layer)


def rule_grid(image: Image.Image, spacing: int = 96, color="FFFFFF",
              opacity: float = 0.06) -> Image.Image:
    base = image.convert("RGBA")
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    tone = parse_color(color) + (int(255 * np.clip(opacity, 0, 1)),)
    for x in range(0, base.width, spacing):
        draw.line([(x, 0), (x, base.height)], fill=tone, width=1)
    for y in range(0, base.height, spacing):
        draw.line([(0, y), (base.width, y)], fill=tone, width=1)
    return Image.alpha_composite(base, layer)


def frosted(background: Image.Image, box: tuple[int, int, int, int],
            radius: float = 0.08, tint="FFFFFF", opacity: float = 0.20,
            blur: float = 22.0, border: float = 0.35) -> Image.Image:
    """A frosted-glass panel cut from the page's own background."""
    left, top, right, bottom = [int(value) for value in box]
    region = background.convert("RGB").crop((left, top, right, bottom))
    region = region.filter(ImageFilter.GaussianBlur(blur))
    # Desaturating before the wash is what keeps frosted glass from turning
    # muddy when it sits over a warm highlight.
    gray = ImageOps.grayscale(region).convert("RGB")
    region = Image.blend(region, gray, 0.32)
    wash = Image.new("RGB", region.size, parse_color(tint))
    region = Image.blend(region, wash, float(np.clip(opacity, 0, 1)))
    region = grain(region, 0.02)
    region = rounded(region, radius)
    if border > 0:
        draw = ImageDraw.Draw(region)
        width = max(1, int(border * 4))
        draw.rounded_rectangle([width, width, region.width - width, region.height - width],
                               radius=int(min(region.width, region.height) * radius),
                               outline=parse_color(tint) + (90,), width=width)
    return region


def diagonal_split(left: Image.Image, right: Image.Image, size: tuple[int, int],
                   split: float = 0.60, lean: float = 0.16,
                   seam: str | None = None, seam_width: int = 5) -> Image.Image:
    """Two photographs joined by a diagonal cut instead of a straight gutter."""
    width, height = size
    left_image = cover(left, size)
    canvas = cover(right, size).convert("RGB")
    mask = Image.new("L", size, 0)
    top_x = int(width * split)
    bottom_x = int(width * (split - lean))
    ImageDraw.Draw(mask).polygon(
        [(0, 0), (top_x, 0), (bottom_x, height), (0, height)], fill=255)
    canvas.paste(left_image, (0, 0), mask)
    if seam:
        draw = ImageDraw.Draw(canvas)
        draw.line([(top_x, 0), (bottom_x, height)], fill=parse_color(seam),
                  width=max(1, seam_width))
    return canvas


def text_art(text: str, font_size: int, *, fill: str | Sequence[str] | None = None,
             image: Image.Image | None = None, stroke: int = 0,
             stroke_color="FFFFFF", shadow: float = 0.0,
             padding: float = 0.12, outline: int = 0,
             outline_color="FFFFFF") -> Image.Image:
    """Rasterise display type filled with a color, gradient or photograph.

    Used for decorative numerals and word-art only; semantic copy stays as real
    text boxes. With ``stroke`` > 0 the glyph becomes a hollow outline in
    ``stroke_color`` (``fill``/``image`` are ignored).
    """
    font = display_font(font_size)
    if font is None:
        raise ValueError("no display font available for text art")
    bbox = font.getbbox(text)
    text_w, text_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pad = max(8, int(font_size * padding))
    size = (text_w + pad * 2, text_h + pad * 2)
    pos = (pad - bbox[0], pad - bbox[1])

    mask = Image.new("L", size, 0)
    draw = ImageDraw.Draw(mask)
    if stroke > 0:
        draw.text(pos, text, font=font, fill=255,
                  stroke_width=stroke, stroke_fill=255)
        draw.text(pos, text, font=font, fill=0)
        art = Image.new("RGBA", size, parse_color(stroke_color) + (255,))
    else:
        draw.text(pos, text, font=font, fill=255)
        if image is not None:
            canvas = Image.new("RGB", size, parse_color("0B1324"))
            canvas.paste(cover(image, size), (0, 0))
            art = canvas.convert("RGBA")
        elif isinstance(fill, (list, tuple)) and len(fill) > 1:
            gradient = np.zeros((size[1], size[0], 3), dtype=np.float32)
            stops = [parse_color(color) for color in fill]
            for x in range(size[0]):
                at = x / max(1, size[0] - 1)
                segment = at * (len(stops) - 1)
                index = min(int(segment), len(stops) - 2)
                local = segment - index
                gradient[:, x] = (np.array(stops[index]) * (1 - local)
                                  + np.array(stops[index + 1]) * local)
            art = Image.fromarray(gradient.astype(np.uint8), "RGB").convert("RGBA")
        else:
            color = parse_color(fill if isinstance(fill, str) else "FFFFFF")
            art = Image.new("RGBA", size, color + (255,))
    art.putalpha(mask)

    if outline > 0 and image is not None:
        # A thin rim keeps a photo-filled glyph readable against a busy page.
        grown = mask.filter(ImageFilter.MaxFilter(2 * outline + 1))
        ring = Image.fromarray(np.clip(np.asarray(grown, dtype=np.int16)
                                       - np.asarray(mask, dtype=np.int16), 0, 255).astype(np.uint8))
        rim = Image.new("RGBA", art.size, parse_color(outline_color) + (255,))
        rim.putalpha(ring)
        rim.alpha_composite(art, (0, 0))
        art = rim

    if shadow > 0:
        offset = int(font_size * 0.035)
        canvas = Image.new("RGBA", (size[0] + offset * 2, size[1] + offset * 2), (0, 0, 0, 0))
        shade_mask = Image.new("L", canvas.size, 0)
        shade_mask.paste(mask, (offset, offset))
        shade_mask = shade_mask.filter(ImageFilter.GaussianBlur(max(1, int(font_size * 0.02))))
        shade = Image.new("RGBA", canvas.size,
                          (0, 0, 0, int(255 * float(np.clip(shadow, 0, 1)))))
        shade.putalpha(shade_mask)
        shade.alpha_composite(art, (offset, offset))
        return shade
    return art


def halftone(image: Image.Image, spacing: int = 7, angle: float = 0.0) -> Image.Image:
    """Print-style dot screen; use only for section art, never for evidence."""
    import math
    gray = ImageOps.grayscale(image).convert("RGB")
    gray = gray.rotate(angle, expand=False)
    width, height = gray.size
    out = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(out)
    pixels = gray.load()
    for y in range(0, height, spacing):
        for x in range(0, width, spacing):
            value = pixels[min(x, width - 1), min(y, height - 1)][0]
            radius = max(0.4, (1 - value / 255) * spacing * 0.55)
            draw.ellipse([x - radius, y - radius, x + radius, y + radius], fill=(0, 0, 0))
    return out.rotate(-angle, expand=False)
