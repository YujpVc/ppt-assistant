#!/usr/bin/env python3
"""
restyle.py —— 把已有 PPT 原地「套用模板风格」：统一字体与配色，输出新文件。
（排版级重排/内容改动走 ingest → 规格 → generate 的工作流。）

规则：
  - 字体：字号 >= 20pt 视为标题 → 华文中宋；否则正文 → Noto Serif SC。
  - 配色：黑色及非白文本 → 主色藏蓝 #192D77；白色文本保留（避免深色背景上文字消失）。

用法：
    python3 restyle.py <input.pptx> <output.pptx>
"""
import sys
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

NAVY = RGBColor(0x19, 0x2D, 0x77)
TITLE_FONT = "华文中宋"
BODY_FONT = "Noto Serif SC"

# 模板调色板（这些颜色保留，不归一化）
PALETTE = {'192D77', 'C00000', 'D24726'}

def _set_font(run, name):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set('typeface', name)

def _rgb_str(run):
    """返回 run 当前实色 'RRGGBB'，theme/自动色等返回 None。"""
    try:
        s = str(run.font.color.rgb)
        return s if len(s) == 6 else None
    except Exception:
        return None

def _is_whiteish(hex6):
    r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
    return r > 220 and g > 220 and b > 220

def restyle(in_path, out_path):
    prs = Presentation(in_path)
    stats = {'runs': 0, 'font': 0, 'color': 0, 'title': 0, 'body': 0}
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if not run.text.strip():
                        continue
                    stats['runs'] += 1
                    size = run.font.size.pt if run.font.size else 18
                    name = TITLE_FONT if size >= 20 else BODY_FONT
                    if size >= 20:
                        stats['title'] += 1
                    else:
                        stats['body'] += 1
                    _set_font(run, name)
                    stats['font'] += 1
                    hex6 = _rgb_str(run)
                    if hex6 and not _is_whiteish(hex6) and hex6.upper() not in PALETTE:
                        run.font.color.rgb = NAVY
                        stats['color'] += 1
    prs.save(out_path)
    return stats

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("用法: python3 restyle.py <input.pptx> <output.pptx>")
        sys.exit(1)
    s = restyle(sys.argv[1], sys.argv[2])
    print(f"已重排: {sys.argv[1]} -> {sys.argv[2]}")
    print(f"  处理文本段 {s['runs']}，标题 {s['title']}，正文 {s['body']}，改字体 {s['font']}，改配色 {s['color']}")
