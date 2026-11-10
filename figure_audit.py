#!/usr/bin/env python3
"""Flag dense technical figures that will be unreadable when projected.

A 4800 px-wide paper composite placed at 7 in has ~686 px/in.  A 30 px label
inside it is then only ~3 pt on the slide -- high resolution does not help.
This audit estimates that effective label size from the source pixels and the
displayed size, and flags pictures that are both dense and large enough to be
the main evidence on a page.
"""
from __future__ import annotations

import argparse
import io
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

PAGE_W_IN = 13.333
PAGE_H_IN = 7.5
TESSERACT = shutil.which("tesseract")


def ocr_text_height(image: Image.Image) -> tuple[float, int]:
    """Median OCR word height in px; 0 if OCR is unavailable/no text."""
    if not TESSERACT:
        return 0.0, 0
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        image.save(tmp.name)
        tmp_path = tmp.name
    try:
        proc = subprocess.run(
            [TESSERACT, tmp_path, "stdout", "-l", "chi_sim+eng", "--psm", "11", "tsv"],
            capture_output=True, text=True, timeout=30)
        heights = []
        for line in proc.stdout.splitlines()[1:]:
            cols = line.split("	")
            if len(cols) < 12:
                continue
            text = cols[11].strip()
            try:
                conf = float(cols[10])
                height = int(cols[9])
            except ValueError:
                continue
            if text and conf >= 35 and height >= 3:
                heights.append(height)
        if len(heights) < 5:
            return 0.0, len(heights)
        heights.sort()
        return float(heights[len(heights) // 2]), len(heights)
    except Exception:
        return 0.0, 0
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def audit(path: Path, max_ppi: float = 450.0, min_area: float = 0.12,
          use_ocr: bool = True) -> dict:
    prs = Presentation(str(path))
    rows = []
    for page_no, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            try:
                with Image.open(io.BytesIO(shape.image.blob)) as image:
                    px_w, px_h = image.size
                    image_rgb = image.convert("RGB")
            except Exception:
                continue
            disp_w = shape.width / 914400
            disp_h = shape.height / 914400
            if disp_w <= 0 or disp_h <= 0:
                continue
            ppi = px_w / disp_w
            area = (disp_w * disp_h) / (PAGE_W_IN * PAGE_H_IN)
            gray = image_rgb.convert("L")
            white_pct = sum(1 for value in gray.getdata() if value > 235) / (px_w * px_h) * 100
            text_px, word_count = ocr_text_height(image_rgb) if use_ocr else (0.0, 0)
            text_pt = text_px / ppi * 72.0 if text_px else 0.0
            if text_px:
                dense = (min_area <= area < .85 and text_pt < 13.0
                         and word_count >= 5 and white_pct >= 12.0)
            else:
                # No OCR evidence: only flag very high pixel density, where a
                # pasted paper figure is almost certainly over-detailed.
                dense = max_ppi <= ppi and min_area <= area < .85
            rows.append({
                "page": page_no,
                "name": shape.name,
                "source_px": [px_w, px_h],
                "display_in": [round(disp_w, 2), round(disp_h, 2)],
                "ppi": round(ppi, 0),
                "area_pct": round(area * 100, 1),
                "text_pt": round(text_pt, 1) if text_px else None,
                "word_count": word_count,
                "white_pct": round(white_pct, 1),
                "dense": dense,
            })
    return {"file": str(path), "pictures": rows,
            "dense": [row for row in rows if row["dense"]]}


def main() -> int:
    parser = argparse.ArgumentParser(description="Flag dense figures / unreadable slide images")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--max-ppi", type=float, default=450.0)
    parser.add_argument("--min-area", type=float, default=0.12,
                        help="minimum picture area ratio of the slide (0-1)")
    parser.add_argument("--strict", action="store_true", help="exit 2 when dense figures exist")
    parser.add_argument("--no-ocr", action="store_true", help="skip OCR text-height estimate")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    if not args.pptx.is_file():
        print(f"figure_audit: file not found: {args.pptx}", file=sys.stderr)
        return 2
    result = audit(args.pptx, args.max_ppi, args.min_area, use_ocr=not args.no_ocr)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"figure_audit: {len(result['pictures'])} picture(s)")
        for row in result["pictures"]:
            flag = "DENSE" if row["dense"] else "ok"
            print(f"  P{row['page']:02d} {row['name']}: "
                  f"{row['source_px'][0]}x{row['source_px'][1]}px -> "
                  f"{row['display_in'][0]}x{row['display_in'][1]}in, "
                  f"{row['ppi']:.0f} ppi, area {row['area_pct']}%, "
                  f"OCR text ~{row['text_pt']}pt, white {row['white_pct']}% [{flag}]")
        if result["dense"]:
            print("\nFix: crop to one mechanism, redraw as 3-5 editable boxes, "
                  "or split the figure across pages. Do not paste the whole paper composite.")
        else:
            print("No dense large figure detected.")
    return 2 if args.strict and result["dense"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
