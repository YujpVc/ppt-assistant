#!/usr/bin/env python3
"""Refuse to call a deck final unless an independent visual review exists.

A zero shot is a draft. At least one `visual_review_round1.md` must be in the
output directory, and the latest review must say it passed (or the user has
explicitly accepted the remaining findings in `review_waiver.md`).

Usage:
  bash run_py.sh review_gate.py <output-dir>
"""
from __future__ import annotations
import json, sys
from pathlib import Path

def main() -> int:
    if len(sys.argv) != 2:
        print("usage: review_gate.py <output-dir>", file=sys.stderr)
        return 64
    out = Path(sys.argv[1])
    if not out.is_dir():
        print(f"review_gate: not a directory: {out}", file=sys.stderr)
        return 66
    reviews = sorted(out.glob("visual_review_round*.md"),
                     key=lambda p: (p.stat().st_mtime, p.name))
    final = out / "visual_review_final.md"
    waiver = out / "review_waiver.md"
    if not reviews:
        print("review_gate: FAIL — zero shot is a draft; no visual_review_round*.md found")
        print("next: spawn an independent read-only reviewer, then rebuild if it finds must-fix")
        return 2
    latest = final if final.exists() and (not reviews or final.stat().st_mtime >= reviews[-1].stat().st_mtime) else reviews[-1]
    text = latest.read_text(encoding="utf-8", errors="replace")
    passed = ("通过" in text and "不通过" not in text and "仍需改" not in text)
    if not passed and waiver.exists():
        print(f"review_gate: WAIVED by {waiver.name}; latest={latest.name}")
        return 0
    if not passed:
        print(f"review_gate: FAIL — latest review {latest.name} still has must-fix")
        print("next: fix spec/builder, rebuild, render, then run a NEW independent review")
        return 2
    report = {"status": "pass", "review": latest.name,
              "rounds": [p.name for p in reviews]}
    (out / "review_status.json").write_text(json.dumps(report, ensure_ascii=False, indent=1) + "\n")
    print(f"review_gate: PASS — {latest.name}; rounds={len(reviews)}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
