#!/usr/bin/env python3
"""Baseline snapshot and page-by-page change ledger for scoped deck edits.

A local edit is only trustworthy when it can be proven. This tool implements the
contract used by careful human editors:

  1. ``snapshot`` before touching the deck: stable page id + position + title +
     body lines + numbers + notes.
  2. edit a working copy, never the original.
  3. ``compare`` after rebuilding: exactly which pages changed, which moved,
     which protected pages were left untouched, and whether any change fell
     outside the authorised scope.

Stable ids are the slide names written by ``build_deck.py`` (``cSld/@name`` =
``P01``, ``P02`` ...). They survive reordering; the page number does not. Decks
without stable names fall back to positional ids and the ledger says so.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

NUMBER = re.compile(
    r"-?\d[\d,，]*(?:\.\d+)?"
    r"(?:\s*(?:%|‰|mm|cm|km|kg|mg|N·m|N|元|万元|亿元|美元|万|亿|ms|s|Hz|°))?"
)
CHROME = "chrome:"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _max_font(shape) -> float:
    sizes = []
    for paragraph in shape.text_frame.paragraphs:
        if paragraph.font.size:
            sizes.append(paragraph.font.size.pt)
        for run in paragraph.runs:
            if run.font.size:
                sizes.append(run.font.size.pt)
    return max(sizes) if sizes else 0.0


def _title(slide, texts: list[dict]) -> str:
    named = next((item for item in texts if item["shape"] == "title"), None)
    if named:
        return " ".join(named["text"].split())
    placeholder = slide.shapes.title
    if placeholder is not None and placeholder.text.strip():
        return " ".join(placeholder.text.split())
    if not texts:
        return ""
    # Largest type first, then the top-most shape: that is the visual heading.
    ranked = sorted(texts, key=lambda item: (-item["size"], item.get("top", 0)))
    return " ".join(ranked[0]["text"].split())


def _page_record(slide, index: int) -> dict[str, Any]:
    texts = []
    lines: list[str] = []
    numbers: list[str] = []
    pictures = charts = tables = 0
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            pictures += 1
        elif getattr(shape, "has_chart", False):
            charts += 1
        elif getattr(shape, "has_table", False):
            tables += 1
        if not getattr(shape, "has_text_frame", False):
            continue
        text = shape.text.strip()
        if not text:
            continue
        chrome = shape.name.startswith(CHROME)
        texts.append({
            "shape": shape.name, "shape_id": shape.shape_id,
            "chrome": chrome, "size": round(_max_font(shape), 1),
            "top": round(shape.top / 914400, 2) if shape.top is not None else 0,
            "text": text,
        })
        if chrome:
            continue
        for line in text.split("\n"):
            clean = " ".join(line.split())
            if clean:
                lines.append(clean)
                numbers.extend(NUMBER.findall(clean))
    numbers = ["".join(token.split()).replace("，", "").replace(",", "") for token in numbers]
    notes = ""
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
    name = slide._element.cSld.get("name")
    return {
        "id": name or f"P{index:02d}",
        "stable": bool(name),
        "index": index,
        "title": _title(slide, texts),
        "texts": texts,
        "lines": lines,
        "numbers": numbers,
        "pictures": pictures,
        "charts": charts,
        "tables": tables,
        "notes": notes,
    }


def snapshot(pptx: Path) -> dict[str, Any]:
    prs = Presentation(str(pptx))
    pages = [_page_record(slide, index) for index, slide in enumerate(prs.slides, 1)]
    return {
        "file": str(pptx.resolve()),
        "sha256": sha256(pptx),
        "slides": len(pages),
        "size_inches": [round(prs.slide_width / 914400, 3),
                        round(prs.slide_height / 914400, 3)],
        "stable_ids": any(page["stable"] for page in pages),
        "pages": pages,
    }


def _diff(before: list[str], after: list[str]) -> tuple[list[str], list[str]]:
    left, right = Counter(before), Counter(after)
    return list((left - right).elements()), list((right - left).elements())


def compare(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    before_pages = before["pages"]
    after_pages = after["pages"]
    before_ids = {page["id"] for page in before_pages}
    after_by_id: dict[str, dict] = {}
    for page in after_pages:
        after_by_id.setdefault(page["id"], page)
    results = []
    matched: set[str] = set()
    for page in before_pages:
        current = after_by_id.get(page["id"])
        if current is None and not before["stable_ids"]:
            current = after_pages[page["index"] - 1] if page["index"] <= len(after_pages) else None
        if current is None:
            results.append({"id": page["id"], "index": page["index"], "status": "missing",
                            "title_before": page["title"], "title_after": ""})
            continue
        matched.add(current["id"])
        removed_lines, added_lines = _diff(page["lines"], current["lines"])
        removed_numbers, added_numbers = _diff(page["numbers"], current["numbers"])
        content_changed = bool(removed_lines or added_lines or removed_numbers
                               or added_numbers or page["notes"] != current["notes"]
                               or page["pictures"] != current["pictures"])
        moved = page["index"] != current["index"]
        status = "changed" if content_changed else ("moved" if moved else "unchanged")
        results.append({
            "id": page["id"],
            "index_before": page["index"],
            "index_after": current["index"],
            "status": status,
            "moved": moved,
            "title_before": page["title"],
            "title_after": current["title"],
            "title_changed": page["title"] != current["title"],
            "lines_added": added_lines,
            "lines_removed": removed_lines,
            "numbers_added": added_numbers,
            "numbers_removed": removed_numbers,
            "pictures_before": page["pictures"],
            "pictures_after": current["pictures"],
            "notes_changed": page["notes"] != current["notes"],
        })
    added_pages = [page for page in after_pages if page["id"] not in before_ids
                   and page["id"] not in matched]
    order_before = [page["id"] for page in before_pages]
    order_after = [page["id"] for page in after_pages]
    return {
        "before": {"file": before["file"], "sha256": before["sha256"], "slides": before["slides"]},
        "after": {"file": after["file"], "sha256": after["sha256"], "slides": after["slides"]},
        "stable_ids": before["stable_ids"] and after["stable_ids"],
        "order_changed": order_before != order_after,
        "order_before": order_before,
        "order_after": order_after,
        "pages": results,
        "added": [{"id": page["id"], "index": page["index"], "title": page["title"]}
                  for page in added_pages],
        "changed": [page["id"] for page in results if page["status"] == "changed"],
        "moved": [page["id"] for page in results if page.get("moved")],
        "missing": [page["id"] for page in results if page["status"] == "missing"],
    }


def violations(report: dict[str, Any], protect: set[str], expect: set[str],
               allow_reorder: bool) -> list[str]:
    problems: list[str] = []
    for page in report["pages"]:
        page_id = page["id"]
        touched = page["status"] in {"changed", "missing"}
        if page_id in protect and touched:
            problems.append(f"{page_id} 是保护页，但内容发生了变化")
        if page_id in protect and page.get("moved"):
            problems.append(f"{page_id} 是保护页，但位置从 {page['index_before']} "
                            f"移到了 {page['index_after']}")
        if expect and touched and page_id not in expect:
            problems.append(f"{page_id} 不在授权修改范围内，却发生了内容变化")
        if expect and page.get("moved") and not allow_reorder and page_id not in expect:
            problems.append(f"{page_id} 未获授权就调整了页序")
    for page in report["added"]:
        if expect and page["id"] not in expect:
            problems.append(f"{page['id']} 是新增页，不在授权范围内")
    return problems


def markdown(report: dict[str, Any], protect: set[str], expect: set[str],
             problems: list[str]) -> str:
    lines = [
        "# 变更台账",
        "",
        f"- 基线：`{report['before']['file']}`  ·  {report['before']['slides']} 页",
        f"- 当前：`{report['after']['file']}`  ·  {report['after']['slides']} 页",
        f"- 稳定页名：{'是' if report['stable_ids'] else '否（按位置比较，仅供参照）'}",
        f"- 页序变化：{'有' if report['order_changed'] else '无'}",
        f"- 授权页：{', '.join(sorted(expect)) if expect else '未限定'}",
        f"- 保护页：{', '.join(sorted(protect)) if protect else '未限定'}",
        f"- 结论：{'**通过**' if not problems else '**未通过**'}",
        "",
        "| 页 | 位置 | 状态 | 标题 | 文字 +/− | 数字 +/− | 图片 |",
        "| --- | --- | --- | --- | --- | --- | --- |",
    ]
    status_text = {"unchanged": "未变", "changed": "有改动",
                   "moved": "仅换位", "missing": "缺失"}
    for page in report["pages"]:
        if page["status"] == "missing":
            lines.append(f"| {page['id']} | {page['index']} → — | 缺失 | {page['title_before']} "
                         "| — | — | — |")
            continue
        added = len(page["lines_added"]) + len(page["numbers_added"])
        removed = len(page["lines_removed"]) + len(page["numbers_removed"])
        title = page["title_after"] or "—"
        if page["title_changed"]:
            title = f"{page['title_before']} → {page['title_after']}"
        lines.append(
            f"| {page['id']} | {page['index_before']} → {page['index_after']} "
            f"| {status_text.get(page['status'], page['status'])} | {title} "
            f"| +{added} / −{removed} | +{len(page['numbers_added'])} / "
            f"−{len(page['numbers_removed'])} | {page['pictures_before']} → "
            f"{page['pictures_after']} |")
    if report["added"]:
        lines.extend(["", "## 新增页", ""])
        for page in report["added"]:
            lines.append(f"- {page['id']}（第 {page['index']} 页）：{page['title']}")
    lines.extend(["", "## 明细", ""])
    for page in report["pages"]:
        if page["status"] != "changed":
            continue
        lines.append(f"### {page['id']} · {page['title_after'] or page['title_before']}")
        for item in page["lines_removed"]:
            lines.append(f"- − {item}")
        for item in page["lines_added"]:
            lines.append(f"- + {item}")
        for item in page["numbers_removed"]:
            lines.append(f"- 数字 − {item}")
        for item in page["numbers_added"]:
            lines.append(f"- 数字 + {item}")
        if page["notes_changed"]:
            lines.append("- 备注讲稿有变化")
        lines.append("")
    if problems:
        lines.extend(["## 未通过项", ""])
        lines.extend(f"- {item}" for item in problems)
    return "\n".join(lines) + "\n"


def parse_ids(value: str | None) -> set[str]:
    return {token.strip() for token in (value or "").split(",") if token.strip()}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    snap = sub.add_parser("snapshot", help="write the pre-edit baseline")
    snap.add_argument("pptx", type=Path)
    snap.add_argument("out", type=Path)
    cmp_ = sub.add_parser("compare", help="diff the baseline against the rebuilt deck")
    cmp_.add_argument("baseline", type=Path)
    cmp_.add_argument("pptx", type=Path)
    cmp_.add_argument("--out", type=Path, help="markdown ledger path")
    cmp_.add_argument("--json", type=Path, help="machine-readable ledger path")
    cmp_.add_argument("--protect", help="comma-separated page ids that must not change")
    cmp_.add_argument("--expect", help="comma-separated page ids allowed to change")
    cmp_.add_argument("--allow-reorder", action="store_true",
                      help="treat page moves as expected even outside --expect")
    args = parser.parse_args()

    if args.command == "snapshot":
        data = snapshot(args.pptx)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        ids = ", ".join(page["id"] for page in data["pages"])
        print(f"baseline: {args.out} ({data['slides']} pages; {ids})")
        return 0

    before = json.loads(args.baseline.read_text(encoding="utf-8"))
    after = snapshot(args.pptx)
    report = compare(before, after)
    protect, expect = parse_ids(args.protect), parse_ids(args.expect)
    problems = violations(report, protect, expect, args.allow_reorder)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(markdown(report, protect, expect, problems), encoding="utf-8")
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(json.dumps({**report, "protect": sorted(protect),
                                         "expect": sorted(expect),
                                         "violations": problems},
                                        ensure_ascii=False, indent=2), encoding="utf-8")
    changed = len(report["changed"])
    print(f"changed={changed} moved={len(report['moved'])} "
          f"added={len(report['added'])} missing={len(report['missing'])} "
          f"violations={len(problems)}")
    for problem in problems:
        print(f"VIOLATION: {problem}", file=sys.stderr)
    return 2 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
