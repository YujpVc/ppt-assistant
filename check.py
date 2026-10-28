#!/usr/bin/env python3
"""
check.py —— PPT 检查器：排版一致性 + 文字遮挡/溢出。

检查项：
  1. 字体一致性：是否出现模板外的字体。
  2. 配色一致性：文字是否出现模板调色板外的颜色。
  3. 标题对齐：各页主标题（最大字号）横向位置是否一致。
  4. 文字遮挡：按「估算文字实际占用高度」检测文本框/图片重叠。
  5. 文本溢出：按字号+框宽高估算文本是否放不下（启发式，非像素级精确）。
  6. 出界：形状是否超出幻灯片边界。

用法：
    python3 check.py <input.pptx> [--json]
"""
import sys, math, json
from pptx import Presentation

# 模板允许的字体（含主题引用）
OK_FONTS = {"华文中宋", "Noto Serif SC", "Noto Serif SC Bold", "微软雅黑", "等线",
            "Microsoft YaHei", "Arial", "Calibri", "Calibri Light", "Wingdings"}
# 允许的文字色（hex 大写）：中性色 + 全部 palette + 页面底色/封面辅助色
OK_COLORS = {
    "000000", "595959", "FFFFFF", "A5A5A5", "F4F6F9",
    # original
    "192D77", "C00000", "D24726", "E7E6E6", "44546A", "5B9BD5",
    # coastal
    "1B4965", "E07A5F", "62B6CB", "E3EEF6", "5FA8D3", "81B29A",
    # ocean
    "1D3557", "E63946", "457B9D", "F1FAEE", "A8DADC", "F4A261",
    # aurora
    "2A9D8F", "E76F51", "48CAE4", "E8F4F2", "90E0EF",
    # indigo
    "312E81", "E11D48", "6366F1", "EEF2FF", "818CF8", "F59E0B",
    # institutional 体制内庄重
    "1F3A5F", "C0392B", "C9A227", "EEF1F5", "6B7A99", "A8B4C8",
    # rose 暖玫瑰
    "8C3B4C", "C2576B", "E08E6F", "F6ECE7", "D9A08A", "9C7A6B",
    # 封面辅助
    "9FB3DB", "FFC799", "E6ECF6", "CFD9E9", "E2CE8C",
    # 图片署名/辅助灰蓝/投影
    "8A93A6", "C7CDD9",
    # probation native_v2 paper theme
    "15253F", "4D627A", "176C88", "F4F6FA",
    "2DBE62", "4285E8", "F36C2B",
}

SLIDE_W = 12192000
SLIDE_H = 6858000

def _run_info(shape):
    out = []
    if not shape.has_text_frame:
        return out
    for para in shape.text_frame.paragraphs:
        for r in para.runs:
            size = r.font.size.pt if r.font.size else 18
            font = r.font.name or ''
            color = None
            try:
                s = str(r.font.color.rgb)
                if len(s) == 6:
                    color = s.upper()
            except Exception:
                pass
            out.append((font, size, color))
    return out

def _bbox(shape):
    try:
        return (shape.left, shape.top, shape.width, shape.height)
    except Exception:
        return None

def _disp_width(text):
    """估算文本显示宽度（em 数）：CJK 等全角字符≈1，拉丁/数字≈0.55。"""
    w = 0.0
    for ch in text:
        w += 1.0 if ord(ch) > 0x2E80 else 0.55
    return w


def _est_extent(shape):
    """估算形状实际占用框：文字用估算高度，图片用全框。空文本返回 None。"""
    bb = _bbox(shape)
    if not bb:
        return None
    if not shape.has_text_frame:
        return bb
    txt = shape.text_frame.text
    if not txt.strip():
        return None
    font_pt = 0
    for _, size, _ in _run_info(shape):
        if size > font_pt:
            font_pt = size
    if font_pt <= 0:
        font_pt = 18
    w_pt = bb[2] / 914400 * 72
    cpl_em = max(0.5, w_pt / font_pt)   # 每行可容纳的 em 数
    n_lines = 0
    for line in txt.split('\n'):
        n_lines += max(1, math.ceil(_disp_width(line) / cpl_em))
    line_h_emu = int(font_pt * 1.25 / 72 * 914400)
    est_h = n_lines * line_h_emu
    return (bb[0], bb[1], bb[2], est_h)

def _overlap(a, b):
    if a[2] < 10000 or a[3] < 10000 or b[2] < 10000 or b[3] < 10000:
        return False
    ax2, ay2 = a[0]+a[2], a[1]+a[3]
    bx2, by2 = b[0]+b[2], b[1]+b[3]
    return not (ax2 <= b[0] or bx2 <= a[0] or ay2 <= b[1] or by2 <= a[1])

def _overflow_est(text, font_pt, box):
    if not text.strip() or not box:
        return None
    w_pt = box[2] / 914400 * 72
    h_pt = box[3] / 914400 * 72
    if font_pt <= 0:
        font_pt = 18
    cpl_em = max(0.5, w_pt / font_pt)
    n_lines = 0
    for line in text.split('\n'):
        n_lines += max(1, math.ceil(_disp_width(line) / cpl_em))
    fit_lines = max(1, int(h_pt / (font_pt * 1.25)))
    if n_lines > fit_lines:
        return {'need_lines': n_lines, 'fit_lines': fit_lines, 'font_pt': font_pt}
    return None

def check(path):
    prs = Presentation(path)
    report = {
        'slides': len(prs.slides),
        'fonts': {'unexpected': []},
        'colors': {'unexpected': []},
        'titles': [],
        'overlap': [],
        'overflow': [],
        'out_of_bounds': [],
    }
    for i, slide in enumerate(prs.slides, 1):
        shapes = list(slide.shapes)
        # 主标题：本页最大字号文本（>=24pt）
        best = None
        for sh in shapes:
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            for _, size, _ in _run_info(sh):
                if size >= 24 and (best is None or size > best[1]):
                    best = (sh, size)
        if best:
            bb = _bbox(best[0])
            if bb:
                report['titles'].append({'slide': i, 'x_in': round(bb[0]/914400, 2),
                                         'size': best[1], 'shape': best[0].name})
        # 字体/颜色/出界
        for sh in shapes:
            name = sh.name or ''
            for font, size, color in _run_info(sh):
                if font and font not in OK_FONTS and not font.startswith('+'):
                    report['fonts']['unexpected'].append((i, name, font, size))
                if color and color not in OK_COLORS:
                    report['colors']['unexpected'].append((i, name, color))
            bb = _bbox(sh)
            if bb and (sh.has_text_frame or sh.shape_type == 13):
                if bb[0] < 0 or bb[1] < 0 or bb[0]+bb[2] > SLIDE_W or bb[1]+bb[3] > SLIDE_H:
                    report['out_of_bounds'].append((i, name))
        # 重叠（用估算占用框；跳过图片，因为图片作为背景/图文混排是刻意的）
        extents = [(_est_extent(sh), sh) for sh in shapes]
        for a in range(len(extents)):
            for b in range(a+1, len(extents)):
                ea, sa = extents[a]
                eb, sb = extents[b]
                if not ea or not eb:
                    continue
                if sa.shape_type == 13 or sb.shape_type == 13:
                    continue
                # A wireframe's segments intentionally meet and cross.  That
                # is not text obstruction; keep line/text and text/text checks.
                has_text_a = sa.has_text_frame and sa.text_frame.text.strip()
                has_text_b = sb.has_text_frame and sb.text_frame.text.strip()
                if not has_text_a and not has_text_b:
                    continue
                if _overlap(ea, eb):
                    report['overlap'].append((i, sa.name, sb.name))
        # 溢出
        for sh in shapes:
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            font_pt = 18
            for _, size, _ in _run_info(sh):
                font_pt = size
                break
            est = _overflow_est(sh.text_frame.text, font_pt, _bbox(sh))
            if est:
                report['overflow'].append((i, sh.name, est))
    return report

def _fmt(report):
    L = []
    L.append("=== PPT 检查报告 ===")
    L.append(f"页数: {report['slides']}")
    L.append(f"\n【1. 字体】非模板字体: {len(report['fonts']['unexpected'])} 处")
    for slide, sh, font, size in report['fonts']['unexpected']:
        L.append(f"  第{slide}页 [{sh}] 字体={font} ({size}pt)")
    L.append(f"\n【2. 配色】非模板文字色: {len(report['colors']['unexpected'])} 处")
    for slide, sh, color in report['colors']['unexpected']:
        L.append(f"  第{slide}页 [{sh}] 颜色=#{color}")
    L.append("\n【3. 标题对齐】各页主标题横向位置:")
    xs = [t['x_in'] for t in report['titles']]
    for t in report['titles']:
        L.append(f"  第{t['slide']}页 x={t['x_in']}in ({t['size']}pt)")
    if xs:
        mode = max(set(xs), key=xs.count)
        dev = [t for t in report['titles'] if abs(t['x_in'] - mode) > 0.5]
        L.append(f"  主流位置 {mode}in，偏离>0.5in: {len(dev)} 页")
        for t in dev:
            L.append(f"    第{t['slide']}页 x={t['x_in']}in")
    L.append(f"\n【4. 文字遮挡】重叠: {len(report['overlap'])} 处")
    for slide, na, nb in report['overlap']:
        L.append(f"  第{slide}页: [{na}] 与 [{nb}] 重叠")
    L.append(f"\n【5. 文本溢出】可能溢出: {len(report['overflow'])} 处")
    for slide, sh, info in report['overflow']:
        L.append(f"  第{slide}页 [{sh}] 需 {info['need_lines']} 行 > 可容 {info['fit_lines']} 行 ({info['font_pt']}pt)")
    L.append(f"\n【6. 出界】: {len(report['out_of_bounds'])} 处")
    for slide, sh in report['out_of_bounds']:
        L.append(f"  第{slide}页 [{sh}] 超出边界")
    return "\n".join(L)

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python3 check.py <input.pptx> [--json]")
        sys.exit(1)
    rep = check(sys.argv[1])
    print(_fmt(rep) if '--json' not in sys.argv else json.dumps(rep, ensure_ascii=False, indent=2))
