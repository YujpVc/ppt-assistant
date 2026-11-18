#!/usr/bin/env python3
"""Catch text-heavy pages that pass density_check but still read as a wall of text.

Non-table/non-mechanism pages: <=420 visible chars, <=28 text blocks.
Table/mechanism pages: <=700 chars, <=80 blocks.
Cover: <=250 chars, <=20 blocks.
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path
from pptx import Presentation

WIDE = {"academic_table", "table", "briefing", "route_map", "architecture",
        "ledger_rails", "diagram_page", "safety_mechanism", "state_overview",
        "task_swimlane", "state_ledger", "trilogy", "pillars"}

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("spec")
    args = ap.parse_args()
    prs = Presentation(args.deck)
    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    pages = spec.get("slides") or spec.get("pages") or []
    problems = []
    for i, slide in enumerate(prs.slides, 1):
        texts = []
        for sh in slide.shapes:
            if sh.has_text_frame:
                t = sh.text_frame.text.strip()
                if t:
                    texts.append(t)
        chars = sum(len(t) for t in texts)
        blocks = len(texts)
        meta = pages[i - 1] if i - 1 < len(pages) else {}
        layout = str(meta.get("layout") or "")
        role = str(meta.get("role") or "")
        wide = (layout in WIDE or "mechanism" in layout or meta.get("mechanism")
                or role in ("mechanism", "table"))
        if i == 1 and layout == "cover":
            lim_c, lim_b = 250, 20
        elif wide:
            lim_c, lim_b = 700, 80
        else:
            lim_c, lim_b = 420, 32
        if chars > lim_c or blocks > lim_b:
            problems.append(f"P{i:02d} {layout or '?'}: {chars} 字 / {blocks} 块（上限 {lim_c}/{lim_b}）")
    if problems:
        print("text_budget_check: below floor")
        for p in problems:
            print("  -", p)
        return 2
    print("text_budget_check: ok")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
