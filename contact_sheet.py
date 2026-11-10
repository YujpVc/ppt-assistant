#!/usr/bin/env python3
"""Compose rendered slides into one contact sheet for a thumbnail review.

Visual direction requires looking at the deck as a whole before inspecting a
single page at full size. This tool is deliberately small: it only arranges the
PNGs produced by ``render_ppt.py`` and writes one image plus optional page
labels, so the reviewer can judge rhythm and density in one glance.
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def natural_key(path: Path):
    parts = re.split(r"(\d+)", path.stem)
    return [int(part) if part.isdigit() else part for part in parts]


def load_pages(source: Path) -> list[Path]:
    if source.is_dir():
        candidates = sorted(source.glob("*.png"), key=natural_key)
    elif source.is_file() and source.suffix.lower() in {".png", ".jpg", ".jpeg"}:
        candidates = [source]
    else:
        raise SystemExit("contact_sheet: pass a render directory (or one PNG)")
    pages = [path for path in candidates if path.is_file()]
    if not pages:
        raise SystemExit(f"contact_sheet: no PNG pages in {source}")
    return pages


def label_font(size: int):
    for path in ("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"):
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("source", type=Path, help="directory from render_ppt.py")
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--cols", type=int, default=4)
    parser.add_argument("--width", type=int, default=480, help="thumbnail width in px")
    parser.add_argument("--gap", type=int, default=14)
    parser.add_argument("--background", default="#0B1324")
    parser.add_argument("--labels", action="store_true", help="draw page numbers")
    args = parser.parse_args()

    pages = load_pages(args.source)
    thumbs = []
    for page in pages:
        with Image.open(page) as source:
            image = source.convert("RGB")
        height = max(1, round(image.height * args.width / image.width))
        thumbs.append(image.resize((args.width, height), Image.LANCZOS))
    label_h = 24 if args.labels else 0
    cell_h = max(image.height for image in thumbs) + label_h
    cols = max(1, min(args.cols, len(thumbs)))
    rows = (len(thumbs) + cols - 1) // cols
    width = cols * args.width + (cols + 1) * args.gap
    height = rows * cell_h + (rows + 1) * args.gap
    sheet = Image.new("RGB", (width, height), args.background)
    draw = ImageDraw.Draw(sheet)
    font = label_font(max(12, args.width // 32))
    for index, image in enumerate(thumbs):
        left = args.gap + (index % cols) * (args.width + args.gap)
        top = args.gap + (index // cols) * (cell_h + args.gap)
        sheet.paste(image, (left, top + label_h + (cell_h - label_h - image.height) // 2))
        if args.labels:
            text = f"{index + 1:02d}"
            draw.text((left + 2, top + 3), text, fill="#FFFFFF", font=font)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(args.out, "PNG")
    print(f"contact_sheet: {len(thumbs)} pages -> {args.out} ({sheet.width}x{sheet.height})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
