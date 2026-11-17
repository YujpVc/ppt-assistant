#!/usr/bin/env python3
"""Refuse to call a deck final unless a signed independent visual review exists.

The latest `visual_review_round*.md` / `visual_review_final.md` must contain a
fenced `review-meta` JSON block:

```review-meta
{"reviewer_session_id":"session-...", "builder_session_id":"session-...",
 "status":"pass", "must_fix":[{"id":"M1","status":"fixed"}]}
```

Rules:
- reviewer_session_id != builder_session_id
- status must be `pass`
- every must_fix item must be `fixed` (or `waived` + review_waiver.md)
- no block = FAIL; self review / same id = FAIL
"""
from __future__ import annotations
import json, re, sys
from pathlib import Path

META_RE = re.compile(r"```review-meta\s*(\{.*?\})\s*```", re.S)

def fail(msg: str) -> int:
    print("review_gate: FAIL — " + msg)
    return 2

def load_meta(path: Path):
    text = path.read_text(encoding="utf-8", errors="replace")
    m = META_RE.search(text)
    if not m:
        return None, text
    try:
        meta = json.loads(m.group(1))
    except json.JSONDecodeError as e:
        return {"_error": f"review-meta is not JSON: {e}"}, text
    return meta, text

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
        return fail("no visual_review_round*.md; zero shot is a draft and must be reviewed")
    latest = final if final.exists() and (not reviews or final.stat().st_mtime >= reviews[-1].stat().st_mtime) else reviews[-1]
    meta, text = load_meta(latest)
    if meta is None:
        return fail(f"{latest.name} has no ```review-meta``` block; self-written prose is not enough")
    if "_error" in meta:
        return fail(f"{latest.name}: {meta['_error']}")
    reviewer = str(meta.get("reviewer_session_id") or "").strip()
    builder = str(meta.get("builder_session_id") or "").strip()
    if not reviewer or not builder:
        return fail(f"{latest.name}: reviewer_session_id and builder_session_id are required")
    if reviewer == builder:
        return fail(f"{latest.name}: reviewer_session_id equals builder_session_id; self review is not independent")
    if meta.get("status") != "pass":
        return fail(f"{latest.name}: status={meta.get('status')!r}; must be 'pass'")
    must = meta.get("must_fix") or []
    if not isinstance(must, list):
        return fail(f"{latest.name}: must_fix must be a list")
    open_items = []
    for item in must:
        if not isinstance(item, dict):
            open_items.append(str(item))
            continue
        status = str(item.get("status") or "open")
        if status not in ("fixed", "waived"):
            open_items.append(item.get("id") or "?")
        if status == "waived" and not waiver.exists():
            open_items.append(f"{item.get('id') or '?'} (waived without review_waiver.md)")
    if open_items:
        return fail(f"{latest.name}: must_fix not closed: {', '.join(map(str, open_items))}")
    report = {"status": "pass", "review": latest.name,
              "reviewer_session_id": reviewer, "builder_session_id": builder,
              "must_fix": must, "rounds": [p.name for p in reviews]}
    (out / "review_status.json").write_text(json.dumps(report, ensure_ascii=False, indent=1) + "\n")
    print(f"review_gate: PASS — {latest.name}; reviewer={reviewer}; rounds={len(reviews)}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
