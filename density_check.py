#!/usr/bin/env python3
"""Check that each layout carries enough information modules, not just layout variety.

This is a spec-level gate: it catches the "beautiful but empty" page before build.
It is intentionally structural (modules, labels, evidence, sources), not a style score.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def count(values) -> int:
    return len(values) if isinstance(values, list) else 0


def text_len(value) -> int:
    return len(str(value).strip()) if value is not None else 0


def page_text(page: dict) -> str:
    """Text that actually appears on the page (notes excluded)."""
    parts: list[str] = []

    def walk(value):
        if isinstance(value, str):
            parts.append(value)
        elif isinstance(value, list):
            for item in value:
                walk(item)
        elif isinstance(value, dict):
            for key, item in value.items():
                if key in {"notes", "id", "layout", "image", "assets"}:
                    continue
                walk(item)

    walk(page)
    return "".join(parts)


def module_count(page: dict) -> int:
    total = 1  # the page's own claim/title
    for key in ("items", "steps", "bullets", "stats", "chips", "flow", "step_notes", "blocks"):
        total += count(page.get(key))
    total += count(page.get("rows")) * max(1, count(page.get("headers")))
    return total


def checks(page: dict) -> list[str]:
    layout = page.get("layout", "")
    out: list[str] = []
    if layout == "frost_number_panel":
        if count(page.get("items")) < 3:
            out.append("frost_number_panel: 至少 3 条短标签")
        if text_len(page.get("body")) > 90:
            out.append("frost_number_panel: 主正文超过 90 字")
        for item in page.get("items", []):
            if text_len(item) > 40:
                out.append("frost_number_panel: 单条标签超过 40 字")
        if not page.get("source"):
            out.append("frost_number_panel: 需要来源/依据")
    elif layout == "section_number":
        if count(page.get("items")) < 3:
            out.append("section_number: 至少 3 条本章要点")
        if not (page.get("body") or page.get("transition")):
            out.append("section_number: 需要本章判断或过渡语")
        if text_len(page.get("body")) > 90:
            out.append("section_number: 主正文超过 90 字，改成短判断 + 3 条短要点")
        if layout == "section_number":
            for item in page.get("items", []):
                if text_len(item) > 40:
                    out.append("section_number: 单条要点超过 40 字，改成短标签")
    elif layout == "frost_number_panel":
        # Chapter transition (双磨砂章节页): the grant recipe sets its floor at
        # "一句短判断 + 3 条短标签, 主正文 ≤90 字, ≥150 字", not the content-page floor.
        if count(page.get("items")) < 3:
            out.append("frost_number_panel: 章节过渡至少 3 条短标签")
        if not page.get("body"):
            out.append("frost_number_panel: 需要一句短判断")
        if text_len(page.get("body")) > 90:
            out.append("frost_number_panel: 主正文超过 90 字，压成一句短判断")
        for item in page.get("items", []):
            if text_len(item) > 40:
                out.append("frost_number_panel: 单条标签超过 40 字，改成短标签")
    elif layout == "pillars":
        if not 3 <= count(page.get("items")) <= 4:
            out.append("pillars: 3–4 个支柱")
        if any(text_len(x.get("body")) < 18 for x in page.get("items", [])):
            out.append("pillars: 每个支柱正文至少 18 字")
        if any(not (x.get("foot") or x.get("value") or x.get("tag")) for x in page.get("items", [])):
            out.append("pillars: 每个支柱需要 foot/value/tag，不能只有正文")
        if not page.get("bottom"):
            out.append("pillars: 需要底部结论")
    elif layout == "architecture":
        if count(page.get("items")) != 3:
            out.append("architecture: 3 个阶段")
        if count(page.get("steps")) < 4:
            out.append("architecture: 至少 4 级验证")
        if count(page.get("step_notes")) < 4:
            out.append("architecture: 4 级验证每级补一行说明/证据，不能只有标签")
    elif layout in {"image_panel", "image_panel_strip"}:
        if count(page.get("bullets")) < 3 and count(page.get("steps")) < 3:
            out.append(f"{layout}: 至少 3 条要点或 3 步状态梯")
        if not page.get("image"):
            out.append(f"{layout}: 需要图像")
        if not page.get("body"):
            out.append(f"{layout}: 需要一句实质正文")
    elif layout == "glass_overlay":
        if count(page.get("bullets")) < 4:
            out.append("glass_overlay: 内容页至少 4 条机制/要点；否则改用 glass_canvas")
        if not page.get("body"):
            out.append("glass_overlay: 需要一句实质正文")
    elif layout == "glass_canvas":
        if count(page.get("items")) < 3:
            out.append("glass_canvas: 至少 3 条机制")
        if count(page.get("stats")) < 2:
            out.append("glass_canvas: 至少 2 个坐标/判定标记")
        if count(page.get("chips")) < 2:
            out.append("glass_canvas: 至少 2 条边界/纪律 chips")
    elif layout == "image_band":
        if count(page.get("items")) != 3:
            out.append("image_band: 3 栏证据")
        if not page.get("body"):
            out.append("image_band: 需要一句主张")
    elif layout == "figure_annotated":
        if count(page.get("items")) != 3:
            out.append("figure_annotated: 3 个标注")
        if any(not x.get("next") for x in page.get("items", [])):
            out.append("figure_annotated: 每个标注补 next 接法")
    elif layout == "academic_table":
        if count(page.get("rows")) < 3:
            out.append("academic_table: 至少 3 行")
        if count(page.get("headers")) < 3:
            out.append("academic_table: 至少 3 列")
        if not page.get("caption"):
            out.append("academic_table: 需要表题")
    elif layout == "sequence":
        if count(page.get("items")) < 4:
            out.append("sequence: 至少 4 个节点；3 步改用 route_map/pillars")
        if not page.get("bottom"):
            out.append("sequence: 需要底部结论")
    elif layout in {"route_map", "flow_chart"}:
        if count(page.get("items")) < 3:
            out.append(f"{layout}: 至少 3 个节点")
    elif layout == "blocks":
        blocks = page.get("blocks") or []
        if len(blocks) < 4:
            out.append("blocks: 至少 4 个要素")
        kinds = {str(b.get("type", "text")) for b in blocks if isinstance(b, dict)}
        if len(kinds) < 2:
            out.append("blocks: 至少两种要素类型，不能整页同一种卡")
        if not any(k in kinds for k in {"flow", "stat", "chips", "image", "table"}):
            out.append("blocks: 至少一个 flow/stat/chips/image/table 视觉要素")
        if not page.get("source"):
            out.append("blocks: 需要来源/依据")
        for b in blocks:
            if not isinstance(b, dict):
                continue
            body = str(b.get("body", "")).strip()
            if len(body) > 90 and not b.get("bullets"):
                out.append("blocks: 单个正文段超过 90 字，改成 2–4 条短 bullet")
            for bullet in b.get("bullets", [])[:6]:
                if len(str(bullet).strip()) > 60:
                    out.append("blocks: 单条 bullet 超过 60 字，拆短")
        occupied = set()
        for b in blocks:
            if not isinstance(b, dict):
                continue
            bx = max(0, int(float(b.get("x", 0))))
            by = max(0, int(float(b.get("y", 0))))
            bw = max(1, int(float(b.get("w", 1))))
            bh = max(1, int(float(b.get("h", 1))))
            for yy in range(by, min(10, by + bh)):
                for xx in range(bx, min(12, bx + bw)):
                    occupied.add((xx, yy))
        coverage = len(occupied) / 120.0
        if coverage < .68:
            out.append(f"blocks: 版面覆盖 {coverage:.0%}，低于 68%；下半页不能空着")
        if not any(yy >= 5 for _, yy in occupied):
            out.append("blocks: 下半页没有任何要素")
        if sum(1 for _, yy in occupied if yy >= 8) < 4:
            out.append("blocks: 最后两行（y=8–9）覆盖不足，底部还有空白")
        cards = [b for b in blocks if isinstance(b, dict) and b.get("type") == "card"]
        if len(cards) >= 3:
            sizes = [(round(float(b.get("w", 0)), 1), round(float(b.get("h", 0)), 1)) for b in cards]
            if len(set(sizes)) == 1:
                out.append("blocks: 不要三张等宽等高的 card，至少换一处宽窄/高低")
        for b in cards:
            h = float(b.get("h", 0))
            bullets = b.get("bullets") or []
            body = str(b.get("body", "")).strip()
            if h >= 3.2 and len(bullets) < 3 and len(body) < 60:
                out.append("blocks: card 高但内容少，底部会空；缩高度或补 2–3 条短 bullet")
        flows = [b for b in blocks if isinstance(b, dict) and b.get("type") == "flow"]
        for b in flows:
            items = b.get("items") or []
            if len(items) >= 4:
                shapes = {str(it.get("shape", "round")) for it in items if isinstance(it, dict)}
                roles = {str(it.get("role", "")) for it in items if isinstance(it, dict) and it.get("role")}
                if len(shapes) < 2 and len(roles) < 2:
                    out.append("blocks: 4 节点以上流程仍是同一种矩形；至少换一个 diamond/circle/hex 或 role")
    elif layout == "image_hero":
        if text_len(page.get("body")) < 18 and count(page.get("chips")) < 2:
            out.append("image_hero: 需要一句实质正文或 2 个信息 chip")
    elif layout == "matrix":
        if count(page.get("items")) != 4:
            out.append("matrix: 4 个象限")
        if any(not item.get("tag") and not item.get("value") for item in page.get("items", [])):
            out.append("matrix: 每格需要 tag/value，不能只有标题+正文")
        if not page.get("bottom"):
            out.append("matrix: 需要底部结论带")
    elif layout == "briefing":
        if count(page.get("items")) < 3 or count(page.get("stats")) < 2:
            out.append("briefing: 3 个模块 + 2 个坐标")
    elif layout == "academic_table":
        pass

    if layout not in {"section_number", "cover", "closing", "blocks"}:
        if text_len(page.get("body")) > 160:
            out.append("文字墙: 页面主正文超过 160 字，拆成机制/证据/判据")
        for item in page.get("items", []) if isinstance(page.get("items"), list) else []:
            if isinstance(item, dict) and text_len(item.get("body")) > 160:
                out.append("文字墙: 单个条目正文超过 160 字，拆成短句或分条")

    # Page-level floor: modules alone are not enough; the page must carry enough
    # actual copy and a source. Section pages get a lower copy floor.
    chars = len(page_text(page).replace("\n", ""))
    modules = module_count(page)
    if layout == "section_number":
        if chars < 120:
            out.append(f"section_number: 页面文字量 {chars} 字，低于 120")
    elif layout == "frost_number_panel":
        if chars < 150:
            out.append(f"frost_number_panel: 章节过渡文字量 {chars} 字，低于 150")
        if modules < 4:
            out.append(f"frost_number_panel: 信息模块 {modules} 个，低于 4（页题 + 判断句 + 3 条标签）")
    elif layout not in {"cover", "closing"}:
        if chars < 220:
            out.append(f"内容密度: 页面文字量 {chars} 字，低于 220")
        if modules < 5:
            out.append(f"内容密度: 信息模块 {modules} 个，低于 5")
        if not page.get("source"):
            out.append("内容密度: 需要来源/依据")
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("spec", type=Path)
    parser.add_argument("--strict", action="store_true", help="exit 2 if any page is below floor")
    args = parser.parse_args()
    spec = json.loads(args.spec.read_text(encoding="utf-8"))
    problems: dict[str, list[str]] = {}
    for index, page in enumerate(spec.get("slides", []), 1):
        issues = checks(page)
        if issues:
            problems[f"P{index:02d} {page.get('layout')}"] = issues
    # Deck-level repeat: same element recipe on too many pages is still a template.
    sigs = []
    for page in spec.get("slides", []):
        if page.get("layout") == "blocks":
            kinds = tuple(sorted({str(b.get("type", "text")) for b in page.get("blocks", [])
                                  if isinstance(b, dict)}))
            if kinds:
                sigs.append(kinds)
    blocks_pages = len([p for p in spec.get("slides", []) if p.get("layout") == "blocks"])
    if blocks_pages > 4:
        problems["deck blocks"] = ["blocks 用了 {} 页，超过 4 页；至少换回 2 页侧图清单/四道闸/玻璃/表格原型".format(blocks_pages)]
    if len(sigs) >= 4:
        counts = {}
        for sig in sigs:
            counts[sig] = counts.get(sig, 0) + 1
        if max(counts.values()) > 3:
            problems["deck blocks"] = ["同一组要素组合用了超过 3 页，换宽窄/顺序/要素类型"]
        if len(set(sigs)) < 3:
            problems["deck blocks"] = ["blocks 要素组合少于 3 种，整章会像同一模板"]

    if not problems:
        print("density_check: ok — every layout has enough modules")
        return 0
    print("density_check: below floor")
    for page, issues in problems.items():
        print(f"  {page}")
        for issue in issues:
            print(f"    - {issue}")
    return 2 if args.strict else 0


if __name__ == "__main__":
    raise SystemExit(main())
