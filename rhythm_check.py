#!/usr/bin/env python3
"""Catch cross-page repetition that page-level density checks miss.

- three or more consecutive pages in the same layout family are a template trail
- image-dominant layouts must not exceed 60% of an 8+ page deck
- an 8+ page mechanism deck needs at least two editable diagram pages
  (architecture / route_map / flow_chart / diagram_page / state_* / blocks graph)
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path

FAMILY = {
    "image_panel": "image_panel", "image_panel_strip": "image_panel",
    "image_hero": "image_hero", "image_statement": "image_hero",
    "image_band": "image_hero", "image_overlay": "image_hero",
    "glass_overlay": "glass", "glass_canvas": "glass",
    "matrix": "matrix", "compare_grid": "matrix", "diagonal_compare": "matrix",
    "academic_table": "table", "table": "table",
    "sequence": "sequence", "pillars": "pillars", "briefing": "briefing",
}
MECH = {"architecture", "route_map", "flow_chart", "diagram_page",
        "state_overview", "task_swimlane", "safety_mechanism", "ledger_rails",
        "state_band", "state_gallery", "state_triptych", "process"}
IMAGE_DOMINANT = {"image_hero", "image_statement", "image_band", "image_overlay",
                  "image_panel", "image_panel_strip", "image_annotated",
                  "evidence_grid", "image_grid", "image_mosaic", "image_wall"}

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    args = ap.parse_args()
    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    slides = spec.get("slides") or []
    if len(slides) < 8:
        print("rhythm_check: ok — <8 pages, rhythm gate skipped")
        return 0
    families, problems = [], []
    missing_skeleton = [str(p.get("id") or i + 1) for i, p in enumerate(slides) if not str(p.get("skeleton") or "").strip()]
    if missing_skeleton:
        problems.append("8+ 页必须给每页写 skeleton（编号列/左图右列/状态账本/链式/时间轴），缺：" + "、".join(missing_skeleton))
    for p in slides:
        layout = str(p.get("layout") or "")
        fam = str(p.get("skeleton") or "").strip() or FAMILY.get(layout, layout)
        if layout == "blocks":
            types = {b.get("type") for b in (p.get("blocks") or []) if isinstance(b, dict)}
            fam = "blocks:diagram" if types & {"graph", "flow", "timeline"} else "blocks"
        families.append(fam)
    from collections import Counter
    counts = Counter(families)
    repeated = [f for f, n in counts.items() if n >= 2]
    if len(repeated) >= 2:
        problems.append("同一骨架出现两对以上：" + "、".join(f"{f}×{counts[f]}" for f in repeated))
    if counts and max(counts.values()) > 3:
        worst = max(counts, key=counts.get)
        problems.append(f"单一骨架 {worst} 超过 3 页（{counts[worst]} 页）")
    run = 1
    for i in range(1, len(families)):
        run = run + 1 if families[i] == families[i - 1] else 1
        if run >= 3:
            problems.append(f"连续同族版式：{families[i]} 连续 {run} 页（{i-run+2}–{i+1}）")
            run = 1
    image_n = sum(1 for p in slides if str(p.get("layout")) in IMAGE_DOMINANT)
    if image_n > int(len(slides) * 0.6):
        problems.append(f"图片主导页 {image_n}/{len(slides)}（>60%），机制图被照片替代")
    mech_n = 0
    for p in slides:
        layout = str(p.get("layout"))
        if layout in MECH:
            mech_n += 1
        elif layout == "blocks":
            types = {b.get("type") for b in (p.get("blocks") or []) if isinstance(b, dict)}
            if types & {"graph", "flow", "timeline"}:
                mech_n += 1
    if mech_n < 2:
        problems.append(f"机制图只有 {mech_n} 页；8+ 页至少 2 页可编辑机制图")
    if problems:
        print("rhythm_check: below floor")
        for p in problems:
            print("  -", p)
        return 2
    print(f"rhythm_check: ok — {len(slides)} pages, {mech_n} mechanism pages")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
