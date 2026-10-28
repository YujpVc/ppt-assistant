#!/usr/bin/env python3
"""
PPT 生成引擎 —— 以指定模板为基底，套用其视觉规范（藏蓝 #192D77 / 暗红 #C00000 /
章节标 #D24726，标题华文中宋，正文 Noto Serif SC），按 JSON 内容规格生成可编辑 .pptx。

版式原则：饱满充实、无过多空白、不杂乱 —— 内容卡片网格铺满正文区，底部统一页脚+页码，
可选「要点提炼」条；标题/章节/页脚三区对齐一致。

用法：
    python3 generate_ppt.py spec.json 输出.pptx [--template 模板.pptx]

内容规格 JSON 见 README 或同目录 spec 示例。
"""
import sys, os, json, math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

# ------------------------- 配色系统 -------------------------
# 主色/强调/章节标/浅底 会按所选 palette 切换；BLACK/GREY 为固定中性色。
# 默认 original = 模板原配色。
NAVY   = RGBColor(0x19, 0x2D, 0x77)   # primary 主色：深藏蓝
RED    = RGBColor(0xC0, 0x00, 0x00)   # accent 强调：暗红
ORANGE = RGBColor(0xD2, 0x47, 0x26)   # marker 章节标：红橙
BLACK  = RGBColor(0x00, 0x00, 0x00)
GREY   = RGBColor(0x59, 0x59, 0x59)
LIGHT  = RGBColor(0xE7, 0xE6, 0xE6)   # light 浅底（卡片/表格斑马纹）
BG     = RGBColor(0xF4, 0xF6, 0xF9)   # 页面底（覆盖模板母版背景图）

PALETTES = {
    'original': {
        'name': '原模板色', 'primary': '192D77', 'accent': 'C00000', 'marker': 'D24726',
        'light': 'E7E6E6', 'series': ['192D77', 'C00000', 'D24726', '44546A', '5B9BD5'],
    },
    'coastal': {  # 海岸蓝调
        'name': '海岸蓝调', 'primary': '1B4965', 'accent': 'E07A5F', 'marker': '62B6CB',
        'light': 'E3EEF6', 'series': ['1B4965', '62B6CB', '5FA8D3', 'E07A5F', '81B29A'],
    },
    'ocean': {  # 深海蓝
        'name': '深海蓝', 'primary': '1D3557', 'accent': 'E63946', 'marker': '457B9D',
        'light': 'F1FAEE', 'series': ['1D3557', '457B9D', 'A8DADC', 'E63946', 'F4A261'],
    },
    'aurora': {  # 极光青绿
        'name': '极光青绿', 'primary': '2A9D8F', 'accent': 'E76F51', 'marker': '48CAE4',
        'light': 'E8F4F2', 'series': ['2A9D8F', '48CAE4', '90E0EF', 'E76F51', 'F4A261'],
    },
    'indigo': {  # 靛蓝紫
        'name': '靛蓝紫', 'primary': '312E81', 'accent': 'E11D48', 'marker': '6366F1',
        'light': 'EEF2FF', 'series': ['312E81', '6366F1', '818CF8', 'E11D48', 'F59E0B'],
    },
    'institutional': {  # 体制内庄重：深藏蓝 + 中国红 + 金
        'name': '体制内庄重', 'primary': '1F3A5F', 'accent': 'C0392B', 'marker': 'C9A227',
        'light': 'EEF1F5', 'series': ['1F3A5F', 'C0392B', 'C9A227', '6B7A99', 'A8B4C8'],
    },
    'rose': {  # 暖玫瑰：酒红 + 玫瑰 + 杏橘 + 奶油（适合温暖情感主题）
        'name': '暖玫瑰', 'primary': '8C3B4C', 'accent': 'C2576B', 'marker': 'E08E6F',
        'light': 'F6ECE7', 'series': ['8C3B4C', 'C2576B', 'E08E6F', 'D9A08A', '9C7A6B'],
    },
}

_ACTIVE_SERIES = [NAVY, ORANGE, RED, GREY]

def _apply_palette(name):
    global NAVY, RED, ORANGE, LIGHT, _ACTIVE_SERIES
    p = PALETTES.get(name, PALETTES['original'])
    NAVY = RGBColor.from_string(p['primary'])
    RED = RGBColor.from_string(p['accent'])
    ORANGE = RGBColor.from_string(p['marker'])
    LIGHT = RGBColor.from_string(p['light'])
    _ACTIVE_SERIES = [RGBColor.from_string(s) for s in p['series']]
    return p

TITLE_FONT = "华文中宋"
BODY_FONT  = "Noto Serif SC"
YAHEI      = "微软雅黑"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

SIZE_CHAPTER = 18        # 章节标记
SIZE_TITLE   = 30        # 页标题
SIZE_HEAD    = 20        # 小节标题
SIZE_BODY    = 16        # 正文
SIZE_SMALL   = 13        # 注脚

# 正文版心（英寸）
MARGIN_X = 0.45
CONTENT_TOP = 2.0
CONTENT_BOTTOM = 6.9           # 无 takeaway 时的正文下界
CONTENT_BOTTOM_TAKEAWAY = 6.26 # 有 takeaway 时的正文下界
FOOTER_Y = 6.98

DECK_TITLE = ""


def _set_run(run, text, font=TITLE_FONT, size=16, bold=False, color=NAVY, italic=False):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', font)


def _add_textbox(slide, x, y, w, h, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    return tb, tf


def _add_para(tf, text, font=TITLE_FONT, size=16, bold=False, color=NAVY,
              align=PP_ALIGN.LEFT, first=False, space_after=6, line_spacing=1.15):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    r = p.add_run()
    _set_run(r, text, font=font, size=size, bold=bold, color=color)
    return p


def _add_rect(slide, x, y, w, h, fill=None, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp


# ------------------------- 公共版式元素 -------------------------

def _add_background(slide):
    """覆盖模板母版背景图，铺干净浅色底。"""
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=BG)


_BG_TEXTURES = ['theme_rose', 'theme_peach', 'theme_blush', 'theme_apricot',
                'theme_mauve', 'theme_sunset', 'theme_coral', 'theme_glow']
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _add_rich_background(slide, page=None):
    """满幅暖色纹理背景（渐变+爱心+光晕若隐若现）+ 角落装饰，填满空白、避免单调。"""
    _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=BG)
    if page is not None and _BG_TEXTURES:
        tex = os.path.join(_BASE_DIR, 'assets', _BG_TEXTURES[(page - 1) % len(_BG_TEXTURES)] + '.jpg')
        if os.path.exists(tex):
            crop = _cover_crop(tex, SLIDE_W, SLIDE_H)
            slide.shapes.add_picture(crop, 0, 0, SLIDE_W, SLIDE_H)
            # 高亮度蒙层：让纹理淡入，文字仍清晰，同时消除“大白板”感
            _add_gradient_rect(slide, 0, 0, SLIDE_W, SLIDE_H, 'FFFFFF', angle=0, a1=73, a2=73)
    # 右下角柔和装饰圆 + 小爱心点（纯色淡影，填补角落）
    for (cx, cy, r, col) in [(12.9, 7.0, 2.5, RGBColor(0xF3, 0xDE, 0xD2)),
                              (-0.7, -0.7, 2.0, RGBColor(0xF6, 0xE6, 0xDB))]:
        c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx) - Inches(r), Inches(cy) - Inches(r),
                                   Inches(r * 2), Inches(r * 2))
        c.fill.solid(); c.fill.fore_color.rgb = col; c.line.fill.background(); c.shadow.inherit = False


def _add_topbar(slide):
    """顶部主色条 + 金色细线（体制内双色带）。"""
    _add_rect(slide, 0, 0, SLIDE_W, Inches(0.12), fill=NAVY)
    _add_rect(slide, 0, Inches(0.12), SLIDE_W, Inches(0.035), fill=ORANGE)


def _add_chapter_marker(slide, chapter):
    """章节标记（红橙小字）+ 右上角小菱形与短线点缀。"""
    if chapter:
        tb, tf = _add_textbox(slide, Inches(MARGIN_X), Inches(0.55), Inches(11.0), Inches(0.35))
        _add_para(tf, chapter, font=BODY_FONT, size=SIZE_CHAPTER, bold=False,
                  color=ORANGE, first=True, space_after=0)
    d = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(12.28), Inches(0.6), Inches(0.12), Inches(0.12))
    d.fill.solid(); d.fill.fore_color.rgb = ORANGE; d.line.fill.background(); d.shadow.inherit = False
    _add_rect(slide, Inches(12.46), Inches(0.645), Inches(0.42), Inches(0.03), fill=ORANGE)


def _add_title(slide, title, size=SIZE_TITLE, y=Inches(1.22)):
    """标题 + 左侧强调竖条 + 底部短线。"""
    _add_rect(slide, Inches(MARGIN_X), y + Inches(0.03), Inches(0.09), Inches(0.58), fill=RED)
    tb, tf = _add_textbox(slide, Inches(0.72), y, Inches(11.9), Inches(0.7))
    _add_para(tf, title, font=TITLE_FONT, size=size, bold=True, color=NAVY,
              first=True, space_after=0)
    _add_rect(slide, Inches(0.72), y + Inches(0.7), Inches(1.4), Pt(3), fill=ORANGE)


def _add_footer(slide, page):
    """底部页脚：细线 + 金色菱形点缀 + 短标题 + 页码。"""
    _add_rect(slide, Inches(MARGIN_X), Inches(FOOTER_Y), Inches(12.43), Pt(1), fill=LIGHT)
    d = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(MARGIN_X), Inches(7.045), Inches(0.1), Inches(0.1))
    d.fill.solid(); d.fill.fore_color.rgb = ORANGE; d.line.fill.background(); d.shadow.inherit = False
    if DECK_TITLE:
        tb, tf = _add_textbox(slide, Inches(0.62), Inches(7.02), Inches(10.5), Inches(0.34))
        _add_para(tf, DECK_TITLE, font=YAHEI, size=11.5, bold=True, color=RGBColor(0x9A, 0x5A, 0x6A),
                  first=True, space_after=0)
    if page is not None:
        tb2, tf2 = _add_textbox(slide, Inches(12.2), Inches(7.02), Inches(0.7), Inches(0.34))
        _add_para(tf2, str(page), font=YAHEI, size=11.5, bold=True, color=RGBColor(0x9A, 0x5A, 0x6A),
                  align=PP_ALIGN.RIGHT, first=True, space_after=0)


def _add_takeaway(slide, text):
    """底部「结语」条：深玫瑰底 + 白字 + 金色爱心点缀，视觉重量足、不单薄。"""
    if not text:
        return
    y = Inches(6.24)
    _add_rect(slide, Inches(MARGIN_X), y, Inches(12.43), Inches(0.56), fill=NAVY)
    _add_rect(slide, Inches(MARGIN_X), y, Inches(0.09), Inches(0.56), fill=ORANGE)
    # 金色菱形点缀（代替爱心）
    d = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(0.72), y + Inches(0.2), Inches(0.16), Inches(0.16))
    d.fill.solid(); d.fill.fore_color.rgb = ORANGE; d.line.fill.background(); d.shadow.inherit = False
    tb, tf = _add_textbox(slide, Inches(1.02), y + Inches(0.05), Inches(11.5), Inches(0.48))
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _add_para(tf, "「" + text + "」", font=BODY_FONT, size=14, bold=True,
              color=RGBColor(0xFF, 0xFF, 0xFF), first=True, space_after=0, line_spacing=1.05)


def _add_bullet_icon(slide, x, y, size=0.12):
    """要点前几何小圆点（统一主色，避免杂色）。"""
    ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, Inches(size), Inches(size))
    ic.fill.solid(); ic.fill.fore_color.rgb = NAVY; ic.line.fill.background()
    ic.shadow.inherit = False
    return ic


def _card(slide, x, y, w, h, accent=NAVY):
    """内容卡片：白底 + 细边 + 顶部强调条（层次更清晰）。"""
    _add_rect(slide, x, y, w, h, fill=RGBColor(0xFF, 0xFF, 0xFF), line=RGBColor(0xD8, 0xDE, 0xE8))
    _add_rect(slide, x, y, w, Inches(0.06), fill=accent)


# ------------------------- 图片「花活」工具 -------------------------

def _cover_crop(src, tw, th):
    """把图片中心裁剪到 tw×th 比例，返回临时文件路径（避免拉伸变形）。"""
    from PIL import Image
    import tempfile
    im = Image.open(src)
    w, h = im.size
    target = tw / th
    cur = w / h
    if cur > target:            # 太宽，裁左右
        new_w = int(h * target)
        left = (w - new_w) // 2
        im = im.crop((left, 0, left + new_w, h))
    elif cur < target:          # 太高，裁上下
        new_h = int(w / target)
        top = (h - new_h) // 2
        im = im.crop((0, top, w, top + new_h))
    tmp = tempfile.NamedTemporaryFile(suffix='.jpg', delete=False)
    im.convert('RGB').save(tmp.name, 'JPEG', quality=90)
    return tmp.name


def _add_gradient_rect(slide, x, y, w, h, rgb, angle=0, a1=100, a2=0):
    """渐变矩形（同一颜色透明度渐变，用于图片融入背景/封面压暗）。angle 0=左→右, 90=上→下。"""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.line.fill.background(); shp.shadow.inherit = False
    spPr = shp._element.spPr
    solid = spPr.find(qn('a:solidFill'))
    grad = spPr.makeelement(qn('a:gradFill'), {})
    gsLst = spPr.makeelement(qn('a:gsLst'), {})
    for pos, a in ((0, a1), (100000, a2)):
        gs = spPr.makeelement(qn('a:gs'), {'pos': str(pos)})
        srgb = spPr.makeelement(qn('a:srgbClr'), {'val': rgb})
        if a < 100:
            al = spPr.makeelement(qn('a:alpha'), {'val': str(int(a * 1000))})
            srgb.append(al)
        gs.append(srgb); gsLst.append(gs)
    grad.append(gsLst)
    lin = spPr.makeelement(qn('a:lin'), {'ang': str(int(angle * 60000)), 'scaled': '1'})
    grad.append(lin)
    if solid is not None:
        solid.addprevious(grad)
        spPr.remove(solid)
    else:
        spPr.append(grad)
    return shp


def _set_picture_fill(shape, rId):
    """把 shape 的填充替换为图片填充（拉伸填满）。"""
    spPr = shape._element.spPr
    for tag in ('a:solidFill', 'a:noFill', 'a:blipFill', 'a:gradFill'):
        f = spPr.find(qn(tag))
        if f is not None:
            spPr.remove(f)
    bf = spPr.makeelement(qn('a:blipFill'), {})
    blip = spPr.makeelement(qn('a:blip'), {qn('r:embed'): rId})
    stretch = spPr.makeelement(qn('a:stretch'), {})
    stretch.append(spPr.makeelement(qn('a:fillRect'), {}))
    bf.append(blip); bf.append(stretch)
    spPr.append(bf)


def _add_image_rounded(slide, img_path, x, y, w, h, radius=0.08):
    """圆角图片：中心裁剪到目标比例后，用圆角矩形图片填充。"""
    from PIL import Image
    im = Image.open(img_path)
    crop = _cover_crop(img_path, w, h)
    pic = slide.shapes.add_picture(crop, x, y, w, h)
    rId = pic._element.find(qn('p:blipFill')).find(qn('a:blip')).get(qn('r:embed'))
    pic._element.getparent().remove(pic._element)
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    shp.line.fill.background()
    shp.shadow.inherit = False
    _set_picture_fill(shp, rId)
    return shp


def _add_image_circle(slide, img_path, x, y, d):
    """圆形图片（头像式）。"""
    crop = _cover_crop(img_path, d, d)
    pic = slide.shapes.add_picture(crop, x, y, d, d)
    rId = pic._element.find(qn('p:blipFill')).find(qn('a:blip')).get(qn('r:embed'))
    pic._element.getparent().remove(pic._element)
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    shp.line.fill.background()
    shp.shadow.inherit = False
    _set_picture_fill(shp, rId)
    return shp


def _add_rounded_rect(slide, x, y, w, h, radius=0.08, fill=None, line=None, line_w=1.0):
    """圆角矩形（可作投影底 / 描边框）。"""
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    return shp


def _add_image_framed(slide, img_path, x, y, w, h, radius=0.09, border=None):
    """带柔和投影 + 细描边的圆角图片（更精致、更有质感）。"""
    _add_rounded_rect(slide, x + Inches(0.04), y + Inches(0.05), w, h, radius,
                      fill=RGBColor(0xDF, 0xE4, 0xEC))
    _add_image_rounded(slide, img_path, x, y, w, h, radius)
    if border is not None:
        _add_rounded_rect(slide, x, y, w, h, radius, fill=None, line=border, line_w=1.0)


def _add_fullbleed(slide, img_path):
    """整页铺满的背景图（中心裁剪，不变形）。"""
    crop = _cover_crop(img_path, SLIDE_W, SLIDE_H)
    return slide.shapes.add_picture(crop, 0, 0, SLIDE_W, SLIDE_H)


def _add_caption_chip(slide, x, y, text):
    """图片署名小徽章。"""
    if not text:
        return
    tb, tf = _add_textbox(slide, x, y, Inches(4.0), Inches(0.3))
    _add_para(tf, text, font=YAHEI, size=10, bold=False, color=RGBColor(0x8A, 0x93, 0xA6),
              first=True, space_after=0)


def _resolve_image(d):
    """从 dict 取本地 image 路径（或 keyword 联网取图），返回 (path, attribution)。"""
    img = d.get('image', '')
    if img:
        if not os.path.isabs(img):
            img = os.path.join(os.path.dirname(os.path.abspath(__file__)), img)
        if os.path.exists(img):
            return img, d.get('attr', '')
    kw = d.get('keyword', '')
    if kw:
        import image_fetch
        cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output', 'img')
        os.makedirs(cache_dir, exist_ok=True)
        safe = "".join(c for c in kw if c.isalnum() or c in '-_') or 'img'
        out = os.path.join(cache_dir, f"{safe}.jpg")
        try:
            return image_fetch.fetch(kw, out)
        except Exception:
            return '', ''
    return '', ''


def _grid(n):
    """按卡片数量选网格列数/行数，尽量铺满。"""
    if n <= 1:
        return 1, 1
    if n == 2:
        return 2, 1
    if n == 3:
        return 3, 1
    if n == 4:
        return 4, 1
    if n <= 6:
        return 3, 2
    return 3, math.ceil(n / 3)


def _grid_cells(n, top, bottom, cols, rows, gap=Inches(0.25)):
    """密集网格：铺满正文区，末行卡片自动加宽填满（不留空槽）。返回 [(x,y,w,h)]。"""
    avail_w = Inches(12.43)
    avail_h = bottom - top
    cells = []
    for i in range(n):
        r = i // cols
        items_in_row = min(cols, n - r * cols)
        cw = (avail_w - gap * (items_in_row - 1)) / items_in_row
        ch = (avail_h - gap * (rows - 1)) / rows
        c = i - r * cols
        x = Inches(MARGIN_X) + c * (cw + gap)
        y = top + r * (ch + gap)
        cells.append((x, y, cw, ch))
    return cells


def _card_need_h(b):
    """估算一张 section 卡片内容所需高度（英寸）：图文卡按图+标题+要点，文字卡按标题+要点。"""
    pts = len(b.get('points', [])) or 0
    if b.get('image') or b.get('keyword'):
        return 1.9 + pts * 0.55 + 0.25
    return 0.72 + pts * 0.68 + 0.4


# ------------------------- 各版式构建器 -------------------------

def build_cover(prs, meta, page=None):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_background(slide)
    # 整页英雄图 + 左侧渐变压暗（图文并茂的封面「花活」）
    hero, _ = _resolve_image(meta)
    dark = bool(hero)
    if dark:
        _add_fullbleed(slide, hero)
        # 全页轻蒙层，保证文字在暖色图上仍清晰（可用 cover_overlay 覆盖为暖深色）
        ov = meta.get('cover_overlay', str(NAVY))
        _add_gradient_rect(slide, 0, 0, SLIDE_W, SLIDE_H, ov, angle=0, a1=14, a2=14)
        _add_gradient_rect(slide, 0, 0, Inches(9.6), SLIDE_H, ov, angle=0, a1=93, a2=0)
    title_color = RGBColor(0xFF, 0xFF, 0xFF) if dark else NAVY
    sub_color = RGBColor(0xFF, 0xFF, 0xFF) if dark else GREY
    author_color = RGBColor(0xFF, 0xFF, 0xFF) if dark else BLACK
    date_color = RGBColor(0xCF, 0xD9, 0xE9) if dark else GREY
    org_color = RGBColor(0xFF, 0xC7, 0x99) if dark else ORANGE
    # 顶部组织/品类标签
    if meta.get('org'):
        tb, tf = _add_textbox(slide, Inches(0.7), Inches(0.5), Inches(11.9), Inches(0.4))
        _add_para(tf, meta['org'], font=YAHEI, size=14, bold=True, color=org_color, first=True, space_after=0)
    # 主标题
    tb, tf = _add_textbox(slide, Inches(0.7), Inches(1.15), Inches(11.9), Inches(1.2))
    _add_para(tf, meta.get('title', ''), font=TITLE_FONT, size=44, bold=True,
              color=title_color, first=True, space_after=0)
    # 标题下强调线
    _add_rect(slide, Inches(0.72), Inches(2.4), Inches(2.2), Pt(3), fill=RED)
    # 副标题
    if meta.get('subtitle'):
        tb2, tf2 = _add_textbox(slide, Inches(0.72), Inches(2.6), Inches(11.9), Inches(0.8))
        _add_para(tf2, meta['subtitle'], font=BODY_FONT, size=21, bold=False,
                  color=sub_color, first=True, space_after=0)
    # 作者 / 日期
    y = Inches(3.6)
    if meta.get('author'):
        tb3, tf3 = _add_textbox(slide, Inches(0.72), y, Inches(11.9), Inches(0.5))
        _add_para(tf3, meta['author'], font=BODY_FONT, size=16, bold=False,
                  color=author_color, first=True, space_after=0)
        y = y + Inches(0.45)
    if meta.get('date'):
        tb4, tf4 = _add_textbox(slide, Inches(0.72), y, Inches(11.9), Inches(0.5))
        _add_para(tf4, meta['date'], font=YAHEI, size=14, bold=False,
                  color=date_color, first=True, space_after=0)
    # 底部看点条（一行多列，编号 + 关键词，铺满）
    highlights = meta.get('highlights', [])
    if highlights:
        # 底部柔和渐隐，自然过渡到看点条（消除硬切边）
        _add_gradient_rect(slide, 0, Inches(4.55), SLIDE_W, Inches(0.8), str(NAVY), angle=90, a1=0, a2=100)
        band_y = Inches(5.35)
        band_h = Inches(2.15)
        _add_rect(slide, 0, band_y, SLIDE_W, band_h, fill=NAVY)
        tb5, tf5 = _add_textbox(slide, Inches(0.7), band_y + Inches(0.16), Inches(11.9), Inches(0.35))
        _add_para(tf5, meta.get('highlights_label', '本指南看点'), font=YAHEI, size=13, bold=True,
                  color=RGBColor(0xE2, 0xCE, 0x8C), first=True, space_after=0)
        ncol = len(highlights)
        cw = Inches(11.9) / ncol
        for i, h in enumerate(highlights):
            cx = Inches(0.7) + i * cw
            cy = band_y + Inches(0.6)
            tb1, tf1 = _add_textbox(slide, cx, cy, cw - Inches(0.2), Inches(0.5))
            _add_para(tf1, f"{i + 1:02d}", font=TITLE_FONT, size=22, bold=True,
                      color=ORANGE, first=True, space_after=0)
            tb2, tf2 = _add_textbox(slide, cx, cy + Inches(0.52), cw - Inches(0.3), Inches(0.8))
            _add_para(tf2, h, font=YAHEI, size=13.5, bold=False,
                      color=RGBColor(0xFF, 0xFF, 0xFF), first=True, space_after=0, line_spacing=1.1)
    return slide


def build_toc(prs, s, page=None):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_rich_background(slide, page)
    _add_topbar(slide)
    _add_title(slide, s.get('title', '目录'), size=34, y=Inches(1.0))
    items = s.get('items', [])
    n = len(items)
    if n == 0:
        _add_footer(slide, page)
        return slide
    # 右侧配图（可选，页级 image）
    img, _ = _resolve_image(s)
    if img:
        img_w = Inches(4.2)
        img_x = Inches(12.88) - img_w
        _add_image_framed(slide, img, img_x, Inches(2.0), img_w, Inches(4.65), radius=0.06, border=ORANGE)
    list_w = Inches(7.4) if img else Inches(12.43)
    top = Inches(2.0)
    bottom = Inches(6.85)
    y = top
    intro = s.get('intro', '')
    if intro:
        tb, tf = _add_textbox(slide, Inches(MARGIN_X), y, list_w, Inches(0.5))
        _add_para(tf, intro, font=BODY_FONT, size=14, color=GREY, first=True, space_after=0)
        y = y + Inches(0.6)
    row_h = (bottom - y) / max(n, 1)
    for i, item in enumerate(items):
        iy = y + i * row_h
        # 序号圆
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(MARGIN_X), iy + (row_h - Inches(0.5)) / 2, Inches(0.5), Inches(0.5))
        circ.fill.solid(); circ.fill.fore_color.rgb = ORANGE; circ.line.fill.background()
        circ.shadow.inherit = False
        ctf = circ.text_frame; ctf.word_wrap = False
        cp = ctf.paragraphs[0]; cp.alignment = PP_ALIGN.CENTER
        cr = cp.add_run(); _set_run(cr, str(i + 1), font=YAHEI, size=15, bold=True,
                                    color=RGBColor(0xFF, 0xFF, 0xFF))
        # 条目文字
        tb2, tf2 = _add_textbox(slide, Inches(0.78), iy + (row_h - Inches(0.5)) / 2, list_w - Inches(0.4), Inches(0.5))
        _add_para(tf2, item, font=BODY_FONT, size=19, bold=False, color=NAVY, first=True, space_after=0)
        # 行间细线
        if i < n - 1:
            _add_rect(slide, Inches(0.78), iy + row_h - Inches(0.02), list_w - Inches(0.4), Pt(1), fill=RGBColor(0xE7, 0xD3, 0xC8))
    _add_footer(slide, page)
    return slide


def build_section(prs, s, page=None):
    """按 layout 分派到不同版式：cards / split(rev) / quote / list / banner / big / mosaic / comparison。"""
    layout = s.get('layout', 'cards')
    if layout == 'split':
        return _build_split(prs, s, page)
    if layout == 'quote':
        return _build_quote(prs, s, page)
    if layout == 'list':
        return _build_list(prs, s, page)
    if layout == 'banner':
        return _build_banner(prs, s, page)
    if layout == 'big':
        return _build_big(prs, s, page)
    if layout == 'mosaic':
        return _build_mosaic(prs, s, page)
    if layout == 'comparison':
        return _build_comparison(prs, s, page)
    return _build_cards(prs, s, page)


def _build_cards(prs, s, page=None):
    """卡片网格（图文卡，铺满版心，正文加粗加大）。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_rich_background(slide, page)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    blocks = s.get('blocks', [])
    n = len(blocks)
    if n == 0:
        _add_footer(slide, page)
        return slide
    has_takeaway = bool(s.get('takeaway'))
    top = Inches(1.95)
    bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if has_takeaway else CONTENT_BOTTOM)
    cols, rows = _grid(n)
    cells = _grid_cells(n, top, bottom, cols, rows)
    for i, b in enumerate(blocks):
        x, y, cw, ch = cells[i]
        img, attr = _resolve_image(b)
        _card(slide, x, y, cw, ch, accent=NAVY)
        pad = Inches(0.16)
        img_h = Emu(int(ch * 0.5))
        if img_h > Inches(2.2):
            img_h = Inches(2.2)
        if img:
            ix, iy = x + pad, y + Inches(0.18)
            iw = cw - pad * 2
            _add_image_framed(slide, img, ix, iy, iw, img_h, radius=0.09, border=ORANGE)
            _add_gradient_rect(slide, ix, iy + img_h - Inches(0.16), iw, Inches(0.16),
                               str(LIGHT), angle=90, a1=0, a2=100)
        ty = y + Inches(0.18) + img_h + Inches(0.12)
        tb, tf = _add_textbox(slide, x + Inches(0.26), ty, cw - Inches(0.52), Inches(0.5))
        _add_para(tf, b.get('heading', ''), font=BODY_FONT, size=18, bold=True,
                  color=NAVY, first=True, space_after=0)
        note = b.get('note', '')
        pts = b.get('points', [])
        npts = len(pts)
        note_reserve = Inches(0.46) if note else Inches(0.0)
        if npts:
            body_top = ty + Inches(0.52)
            body_h = (y + ch) - body_top - Inches(0.14) - note_reserve
            step = body_h / npts
            for idx, pt in enumerate(pts):
                py = body_top + idx * step
                _add_bullet_icon(slide, x + Inches(0.28), py + Inches(0.06), size=0.11)
                tb2, tf2 = _add_textbox(slide, x + Inches(0.52), py, cw - Inches(0.78), step - Inches(0.02))
                _add_para(tf2, pt, font=BODY_FONT, size=15, bold=False,
                          color=NAVY, first=True, space_after=0, line_spacing=1.18)
        # 卡片底部：手写感短句 或 细线+菱形点缀（填补下半留白）
        if note:
            ny = y + ch - Inches(0.44)
            _add_rect(slide, x + Inches(0.24), ny - Inches(0.06), cw - Inches(0.48), Pt(1), fill=RGBColor(0xE7, 0xD3, 0xC8))
            tn, tnf = _add_textbox(slide, x + Inches(0.26), ny + Inches(0.02), cw - Inches(0.52), Inches(0.38))
            _add_para(tnf, note, font=YAHEI, size=12, bold=False, color=RGBColor(0x9A, 0x5A, 0x6A),
                      first=True, space_after=0, line_spacing=1.1)
        elif attr:
            _add_caption_chip(slide, x + Inches(0.26), y + ch - Inches(0.32), attr)
        else:
            ly = y + ch - Inches(0.3)
            _add_rect(slide, x + Inches(0.24), ly, cw - Inches(0.48), Pt(1), fill=RGBColor(0xE7, 0xD3, 0xC8))
            dm = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, x + Inches(0.26), ly + Inches(0.04), Inches(0.1), Inches(0.1))
            dm.fill.solid(); dm.fill.fore_color.rgb = ORANGE; dm.line.fill.background(); dm.shadow.inherit = False
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def _build_split(prs, s, page=None):
    """左右分栏：左侧标题+要点，右侧整高大图。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_rich_background(slide, page)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    blocks = s.get('blocks', [])
    has_takeaway = bool(s.get('takeaway'))
    top = Inches(1.98)
    bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if has_takeaway else CONTENT_BOTTOM)
    # 右侧（或左侧）整高大图（优先取页面级 image，否则取首个 block 的图）
    img, attr = _resolve_image(s if s.get('image') else (blocks[0] if blocks else {}))
    reverse = bool(s.get('reverse'))
    img_w = Inches(4.9)
    if reverse:
        img_x = Inches(0.45)
        tx = Inches(5.7)
    else:
        img_x = Inches(12.88) - img_w
        tx = Inches(0.45)
    if img:
        _add_image_framed(slide, img, img_x, top, img_w, bottom - top, radius=0.06, border=ORANGE)
        _add_gradient_rect(slide, img_x, top + (bottom - top) - Inches(0.3), img_w, Inches(0.3),
                           str(LIGHT), angle=90, a1=0, a2=100)
    # 文字（紧凑排布，自动适配块数与要点数，不压到页脚）
    lw = Inches(6.75)
    total_pts = sum(len(b.get('points', [])) for b in blocks)
    n_blocks = len(blocks)
    avail = bottom - top - Inches(0.1)
    # 动态行距：内容多则收紧
    head_h = Inches(0.46)
    gap_h = Inches(0.16)
    pt_h = (avail - n_blocks * head_h - (n_blocks - 1) * gap_h) / max(total_pts, 1)
    pt_h = max(Inches(0.36), min(Inches(0.55), pt_h))
    y = top + Inches(0.05)
    for bi, b in enumerate(blocks):
        heading = b.get('heading', '')
        points = b.get('points', [])
        _add_rect(slide, tx, y + Inches(0.05), Inches(0.08), Inches(0.4), fill=ORANGE)
        tb, tf = _add_textbox(slide, tx + Inches(0.27), y, lw - Inches(0.25), Inches(0.44))
        _add_para(tf, heading, font=BODY_FONT, size=19, bold=True, color=NAVY, first=True, space_after=0)
        y += head_h
        for pt in points:
            _add_bullet_icon(slide, tx + Inches(0.27), y + Inches(0.07), size=0.1)
            tb2, tf2 = _add_textbox(slide, tx + Inches(0.53), y, lw - Inches(0.5), pt_h + Inches(0.05))
            _add_para(tf2, pt, font=BODY_FONT, size=15, color=NAVY, first=True, space_after=0, line_spacing=1.18)
            y += pt_h
        y += gap_h
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def _build_quote(prs, s, page=None):
    """整页引言：满幅背景图 + 暖色蒙层 + 居中大字。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_background(slide)
    img, _ = _resolve_image(s)
    if img:
        _add_fullbleed(slide, img)
        _add_gradient_rect(slide, 0, 0, SLIDE_W, SLIDE_H, '3A1B20', angle=0, a1=46, a2=46)
    else:
        _add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=NAVY)
    text = s.get('quote', '') or s.get('title', '')
    tb, tf = _add_textbox(slide, Inches(0.9), Inches(1.75), Inches(11.53), Inches(2.7))
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _add_para(tf, f"“{text}”", font=TITLE_FONT, size=31, bold=True,
              color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER, first=True, space_after=0, line_spacing=1.28)
    sub = s.get('sub', '')
    if sub:
        tb2, tf2 = _add_textbox(slide, Inches(1.4), Inches(4.55), Inches(10.53), Inches(1.1))
        _add_para(tf2, sub, font=BODY_FONT, size=18, color=RGBColor(0xFF, 0xE3, 0xC7),
                  align=PP_ALIGN.CENTER, first=True, space_after=0, line_spacing=1.3)
    # 底部短横线点缀
    _add_rect(slide, Inches(6.17), Inches(4.35), Inches(1.0), Pt(2.5), fill=ORANGE)
    if page is not None:
        tb3, tf3 = _add_textbox(slide, Inches(12.2), Inches(7.0), Inches(0.7), Inches(0.3))
        _add_para(tf3, str(page), font=YAHEI, size=10, color=RGBColor(0xFF, 0xFF, 0xFF),
                  align=PP_ALIGN.RIGHT, first=True, space_after=0)
    return slide


def _build_list(prs, s, page=None):
    """大序号横条列表：每行 = 序号色块 + 标题要点(内联) + 右侧缩略图。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_rich_background(slide, page)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    blocks = s.get('blocks', [])
    n = len(blocks)
    has_takeaway = bool(s.get('takeaway'))
    top = Inches(1.98)
    bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if has_takeaway else CONTENT_BOTTOM)
    gap = Inches(0.2)
    row_h = (bottom - top - gap * (n - 1)) / n if n else Inches(1.0)
    for i, b in enumerate(blocks):
        y = top + i * (row_h + gap)
        # 行底色 + 左侧序号色块
        _add_rect(slide, Inches(0.45), y, Inches(12.43), row_h, fill=RGBColor(0xFF, 0xFF, 0xFF),
                  line=RGBColor(0xE4, 0xD0, 0xC6))
        _add_rect(slide, Inches(0.45), y, Inches(1.4), row_h, fill=ORANGE)
        tbn = slide.shapes.add_textbox(Inches(0.45), y + (row_h - Inches(0.6)) / 2, Inches(1.4), Inches(0.6)).text_frame
        tbn.word_wrap = False
        pn = tbn.paragraphs[0]; pn.alignment = PP_ALIGN.CENTER
        rn = pn.add_run(); _set_run(rn, f"{i + 1:02d}", font=TITLE_FONT, size=27, bold=True,
                                    color=RGBColor(0xFF, 0xFF, 0xFF))
        # 缩略图（右侧）
        img, _ = _resolve_image(b)
        thumb_w = Inches(1.8) if img else Inches(0.0)
        tx = Inches(2.15)
        tw = Inches(12.88) - tx - (thumb_w + Inches(0.42) if img else Inches(0.32))
        # 标题(粗) + 要点(内联)，垂直居中
        head = b.get('heading', '')
        pts = b.get('points', [])
        tb, tf = _add_textbox(slide, tx, y, tw, row_h)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.line_spacing = 1.18
        r1 = p.add_run(); _set_run(r1, head, font=BODY_FONT, size=17, bold=True, color=NAVY)
        if pts:
            r2 = p.add_run(); _set_run(r2, "　" + " · ".join(pts), font=BODY_FONT, size=14.5,
                                       bold=False, color=NAVY)
        if img:
            _add_image_framed(slide, img, Inches(12.88) - thumb_w - Inches(0.12), y + Inches(0.12),
                              thumb_w, row_h - Inches(0.24), radius=0.1, border=ORANGE)
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def _build_banner(prs, s, page=None):
    """整宽大图横幅 + 下方横排要点（区别于卡片网格）。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_rich_background(slide, page)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    blocks = s.get('blocks', [])
    has_takeaway = bool(s.get('takeaway'))
    img, _ = _resolve_image(s)
    if img:
        _add_image_framed(slide, img, Inches(0.45), Inches(1.92), Inches(12.43), Inches(2.45), radius=0.05, border=ORANGE)
        _add_gradient_rect(slide, Inches(0.45), Inches(4.37) - Inches(0.22), Inches(12.43), Inches(0.22),
                           str(LIGHT), angle=90, a1=0, a2=100)
    top = Inches(4.6)
    bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if has_takeaway else CONTENT_BOTTOM)
    n = len(blocks)
    row_h = (bottom - top) / max(n, 1)
    for i, b in enumerate(blocks):
        y = top + i * row_h
        head = b.get('heading', '')
        pts = b.get('points', [])
        _add_rect(slide, Inches(0.45), y + Inches(0.1), Inches(0.09), row_h - Inches(0.2), fill=ORANGE)
        tb, tf = _add_textbox(slide, Inches(0.72), y, Inches(11.9), row_h)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.line_spacing = 1.15
        r1 = p.add_run(); _set_run(r1, head, font=BODY_FONT, size=18, bold=True, color=NAVY)
        if pts:
            r2 = p.add_run(); _set_run(r2, "　" + " · ".join(pts), font=BODY_FONT, size=15, color=NAVY)
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def _build_big(prs, s, page=None):
    """大字强调页：左侧超大关键词 + 右侧要点卡片。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_rich_background(slide, page)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    big = s.get('big', '')
    sub = s.get('sub', '')
    points = s.get('points', [])
    # 左侧超大词
    tb, tf = _add_textbox(slide, Inches(0.6), Inches(1.7), Inches(6.3), Inches(3.0))
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _add_para(tf, big, font=TITLE_FONT, size=100, bold=True, color=NAVY, first=True, space_after=0)
    _add_rect(slide, Inches(0.66), Inches(4.05), Inches(1.6), Pt(3), fill=ORANGE)
    if sub:
        tb2, tf2 = _add_textbox(slide, Inches(0.62), Inches(4.25), Inches(6.1), Inches(0.7))
        _add_para(tf2, sub, font=BODY_FONT, size=18, color=GREY, first=True, space_after=0, line_spacing=1.2)
    # 右侧要点竖排卡片
    x = Inches(7.35)
    w = Inches(5.55)
    top = Inches(1.98)
    bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if s.get('takeaway') else CONTENT_BOTTOM)
    n = len(points)
    gap = Inches(0.22)
    row_h = (bottom - top - gap * (n - 1)) / max(n, 1)
    for i, pt in enumerate(points):
        y = top + i * (row_h + gap)
        _add_rect(slide, x, y, w, row_h, fill=RGBColor(0xFF, 0xFF, 0xFF), line=RGBColor(0xE4, 0xD0, 0xC6))
        _add_rect(slide, x, y, Inches(0.08), row_h, fill=ORANGE)
        tb3, tf3 = _add_textbox(slide, x + Inches(0.3), y, w - Inches(0.5), row_h)
        tf3.vertical_anchor = MSO_ANCHOR.MIDDLE
        _add_para(tf3, pt, font=BODY_FONT, size=16, color=NAVY, first=True, space_after=0, line_spacing=1.2)
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def _build_mosaic(prs, s, page=None):
    """图片拼贴：左侧大图 + 右侧上下两张小图，各带标题。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_rich_background(slide, page)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    blocks = s.get('blocks', [])
    has_takeaway = bool(s.get('takeaway'))
    top = Inches(1.95)
    bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if has_takeaway else CONTENT_BOTTOM)
    H = bottom - top
    # 左大图
    if len(blocks) >= 1:
        img, _ = _resolve_image(blocks[0])
        if img:
            _add_image_framed(slide, img, Inches(0.45), top, Inches(7.0), H, radius=0.05, border=ORANGE)
            cap = blocks[0].get('heading', '')
            if cap:
                _add_rect(slide, Inches(0.45), top + H - Inches(0.46), Inches(7.0), Inches(0.46), fill=NAVY)
                tb, tf = _add_textbox(slide, Inches(0.62), top + H - Inches(0.42), Inches(6.6), Inches(0.36))
                _add_para(tf, cap, font=BODY_FONT, size=14, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), first=True, space_after=0)
    # 右侧两小图
    small = blocks[1:3]
    sw = Inches(5.0)
    sx = Inches(7.75)
    sh = (H - Inches(0.25)) / max(len(small), 1)
    for i, b in enumerate(small):
        y = top + i * (sh + Inches(0.25))
        img, _ = _resolve_image(b)
        if img:
            _add_image_framed(slide, img, sx, y, sw, sh, radius=0.06, border=ORANGE)
            cap = b.get('heading', '')
            if cap:
                _add_rect(slide, sx, y + sh - Inches(0.4), sw, Inches(0.4), fill=NAVY)
                tb, tf = _add_textbox(slide, sx + Inches(0.18), y + sh - Inches(0.36), sw - Inches(0.36), Inches(0.32))
                _add_para(tf, cap, font=BODY_FONT, size=13, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), first=True, space_after=0)
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def _build_comparison(prs, s, page=None):
    """双栏对比：冷调 vs 暖调（两种语气 / 两种做法）。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_rich_background(slide, page)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    blocks = s.get('blocks', [])
    has_takeaway = bool(s.get('takeaway'))
    top = Inches(1.95)
    bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if has_takeaway else CONTENT_BOTTOM)
    col_w = Inches(6.0)
    gap = Inches(0.43)
    for i, b in enumerate(blocks[:2]):
        x = Inches(0.45) + i * (col_w + gap)
        cold = (i == 0)
        bg = RGBColor(0xED, 0xE7, 0xE6) if cold else RGBColor(0xFB, 0xEC, 0xE3)
        accent = RGBColor(0x8A, 0x93, 0xA6) if cold else ORANGE
        _add_rect(slide, x, top, col_w, bottom - top, fill=bg, line=RGBColor(0xE2, 0xD5, 0xCF))
        _add_rect(slide, x, top, col_w, Inches(0.12), fill=accent)
        tb, tf = _add_textbox(slide, x + Inches(0.32), top + Inches(0.32), col_w - Inches(0.64), Inches(0.5))
        _add_para(tf, b.get('heading', ''), font=BODY_FONT, size=20, bold=True, color=NAVY, first=True, space_after=0)
        py = top + Inches(0.95)
        for pt in b.get('points', []):
            tb2, tf2 = _add_textbox(slide, x + Inches(0.32), py, col_w - Inches(0.64), Inches(0.9))
            _add_para(tf2, "“" + pt + "”", font=BODY_FONT, size=15.5, color=NAVY, first=True, space_after=0, line_spacing=1.22)
            py += Inches(0.85)
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def build_bullets(prs, s, page=None):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_background(slide)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    has_takeaway = bool(s.get('takeaway'))
    top = Inches(1.95)
    bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if has_takeaway else CONTENT_BOTTOM)
    y = top
    intro = s.get('intro', '')
    if intro:
        tb, tf = _add_textbox(slide, Inches(MARGIN_X), y, Inches(12.43), Inches(0.5))
        _add_para(tf, intro, font=BODY_FONT, size=14, color=GREY, first=True, space_after=0)
        y = y + Inches(0.62)
    items = s.get('items', [])
    n = len(items)
    if n:
        avail = bottom - y
        step = avail / n
        if n >= 5:
            # 网格模式（5+ 条）：两列卡片，铺满宽度与高度
            cols = 2
            rows = math.ceil(n / cols)
            gap = Inches(0.25)
            cw = (Inches(12.43) - gap * (cols - 1)) / cols
            ch = (bottom - y - gap * (rows - 1)) / rows
            for i, it in enumerate(items):
                c = i % cols
                r = i // cols
                x = Inches(MARGIN_X) + c * (cw + gap)
                iy = y + r * (ch + gap)
                head = it.get('head', '')
                text = it.get('text', '')
                _add_rect(slide, x, iy, cw, ch, fill=LIGHT)
                _add_rect(slide, x, iy, Inches(0.09), ch, fill=NAVY)
                circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.28), iy + (ch - Inches(0.42)) / 2, Inches(0.42), Inches(0.42))
                circ.fill.solid(); circ.fill.fore_color.rgb = NAVY; circ.line.fill.background()
                circ.shadow.inherit = False
                ctf = circ.text_frame; ctf.word_wrap = False
                cp = ctf.paragraphs[0]; cp.alignment = PP_ALIGN.CENTER
                cr = cp.add_run(); _set_run(cr, str(i + 1), font=YAHEI, size=13, bold=True,
                                            color=RGBColor(0xFF, 0xFF, 0xFF))
                tb, tf = _add_textbox(slide, x + Inches(0.88), iy + Inches(0.12), cw - Inches(1.05), Inches(0.4))
                _add_para(tf, head, font=BODY_FONT, size=14.5, bold=True, color=NAVY, first=True, space_after=0)
                tb2, tf2 = _add_textbox(slide, x + Inches(0.88), iy + Inches(0.52), cw - Inches(1.05), ch - Inches(0.64))
                _add_para(tf2, text, font=BODY_FONT, size=12, color=NAVY,
                          first=True, space_after=0, line_spacing=1.1)
        else:
            # 卡片模式（<=4 条）
            for i, it in enumerate(items):
                iy = y + i * step
                head = it.get('head', '')
                text = it.get('text', '')
                ch = step - Inches(0.14)
                _add_rect(slide, Inches(MARGIN_X), iy, Inches(12.43), ch, fill=LIGHT)
                _add_rect(slide, Inches(MARGIN_X), iy, Inches(0.09), ch, fill=NAVY)
                circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.88), iy + (ch - Inches(0.5)) / 2, Inches(0.5), Inches(0.5))
                circ.fill.solid(); circ.fill.fore_color.rgb = NAVY; circ.line.fill.background()
                circ.shadow.inherit = False
                ctf = circ.text_frame; ctf.word_wrap = False
                cp = ctf.paragraphs[0]; cp.alignment = PP_ALIGN.CENTER
                cr = cp.add_run(); _set_run(cr, str(i + 1), font=YAHEI, size=16, bold=True,
                                            color=RGBColor(0xFF, 0xFF, 0xFF))
                tb, tf = _add_textbox(slide, Inches(1.65), iy + Inches(0.12), Inches(11.0), Inches(0.42))
                _add_para(tf, head, font=BODY_FONT, size=16, bold=True, color=NAVY, first=True, space_after=0)
                tb2, tf2 = _add_textbox(slide, Inches(1.65), iy + Inches(0.56), Inches(11.0), ch - Inches(0.66))
                _add_para(tf2, text, font=BODY_FONT, size=13, color=NAVY,
                          first=True, space_after=0, line_spacing=1.15)
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def build_table(prs, s, page=None):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_background(slide)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    headers = s.get('headers', [])
    rows = s.get('rows', [])
    n_rows = len(rows) + 1
    n_cols = len(headers)
    if n_cols == 0:
        _add_footer(slide, page)
        return slide
    left, top = Inches(MARGIN_X), Inches(1.95)
    width, height = Inches(12.43), Inches(0.58) * n_rows
    gtable = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    table = gtable.table
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = h
        cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                _set_run(r, h, font=YAHEI, size=13, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row):
            cell = table.cell(ri, ci)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT if ri % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
                for r in p.runs:
                    _set_run(r, str(val), font=YAHEI, size=12.5, bold=False, color=BLACK)
    note = s.get('note', '')
    if note:
        tb, tf = _add_textbox(slide, Inches(MARGIN_X), top + Inches(0.58) * n_rows + Inches(0.06),
                              Inches(12.43), Inches(0.4))
        _add_para(tf, note, font=YAHEI, size=11, color=GREY, first=True, space_after=0)
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def build_process(prs, s, page=None):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_background(slide)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    steps = s.get('steps', [])
    n = len(steps)
    if n == 0:
        _add_footer(slide, page)
        return slide
    has_takeaway = bool(s.get('takeaway'))
    top = Inches(CONTENT_TOP)
    bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if has_takeaway else CONTENT_BOTTOM)
    if n <= 4:
        cols, rows = n, 1
    else:
        cols, rows = 3, math.ceil(n / 3)
    # 步骤卡片内容偏少时收缩高度并垂直居中
    need_h = Inches(2.15 if n <= 4 else 3.0)
    avail = bottom - top
    if need_h < avail:
        off = (avail - need_h) / 2
        top += off
        bottom -= off
    cells = _grid_cells(n, top, bottom, cols, rows)
    for i, st in enumerate(steps):
        x, y, cw, ch = cells[i]
        _card(slide, x, y, cw, ch, accent=ORANGE)
        tb, tf = _add_textbox(slide, x + Inches(0.22), y + Inches(0.16), cw - Inches(0.44), Inches(0.5))
        _add_para(tf, st.get('n', str(i + 1)), font=TITLE_FONT, size=22, bold=True,
                  color=ORANGE, first=True, space_after=0)
        tb2, tf2 = _add_textbox(slide, x + Inches(0.22), y + Inches(0.66), cw - Inches(0.44), Inches(0.5))
        _add_para(tf2, st.get('head', ''), font=BODY_FONT, size=16, bold=True,
                  color=NAVY, first=True, space_after=0)
        tb3, tf3 = _add_textbox(slide, x + Inches(0.22), y + Inches(1.2), cw - Inches(0.44), ch - Inches(1.42))
        _add_para(tf3, st.get('text', ''), font=BODY_FONT, size=13.5, color=NAVY,
                  first=True, space_after=0, line_spacing=1.15)
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def build_chart(prs, s, page=None):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_background(slide)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    cs = s.get('chart', {})
    categories = cs.get('categories', [])
    series_list = cs.get('series', [])
    if not categories or not series_list:
        _add_footer(slide, page)
        return slide
    cd = CategoryChartData()
    cd.categories = categories
    for ser in series_list:
        cd.add_series(ser.get('name', ''), ser.get('values', []))
    type_map = {
        'column': XL_CHART_TYPE.COLUMN_CLUSTERED,
        'bar': XL_CHART_TYPE.BAR_CLUSTERED,
        'line': XL_CHART_TYPE.LINE_MARKERS,
        'pie': XL_CHART_TYPE.PIE,
    }
    kind = cs.get('kind', 'column')
    chart_type = type_map.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED)
    has_takeaway = bool(s.get('takeaway'))
    note = s.get('note', '')
    chart_top = Inches(1.95)
    chart_bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if has_takeaway else CONTENT_BOTTOM)
    if note:
        chart_bottom = chart_bottom - Inches(0.38)
    gframe = slide.shapes.add_chart(chart_type, Inches(0.6), chart_top, Inches(12.2), chart_bottom - chart_top, cd)
    chart = gframe.chart
    chart.has_legend = (len(series_list) > 1) or (kind == 'pie')
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    palette = _ACTIVE_SERIES
    for idx, ser in enumerate(chart.series):
        try:
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = palette[idx % len(palette)]
        except Exception:
            pass
    if note:
        tb, tf = _add_textbox(slide, Inches(0.6), chart_bottom + Inches(0.02), Inches(12.2), Inches(0.34))
        _add_para(tf, note, font=YAHEI, size=11, color=GREY, first=True, space_after=0)
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def build_image(prs, s, page=None):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_background(slide)
    _add_topbar(slide)
    _add_chapter_marker(slide, s.get('chapter'))
    _add_title(slide, s.get('title'))
    has_takeaway = bool(s.get('takeaway'))
    bottom = Inches(CONTENT_BOTTOM_TAKEAWAY if has_takeaway else CONTENT_BOTTOM)
    img_path, attr = _resolve_image(s)
    img_bottom = Inches(2.0)
    if img_path and os.path.exists(img_path):
        area_top = Inches(2.0)
        iw = Inches(12.43)
        ih = bottom - area_top - Inches(0.5)
        # 全宽横幅图（中心裁剪不变形）+ 投影描边 + 底部渐变融入
        _add_image_framed(slide, img_path, Inches(MARGIN_X), Inches(2.0), iw, ih, radius=0.03, border=ORANGE)
        _add_gradient_rect(slide, Inches(MARGIN_X), Inches(2.0) + ih - Inches(0.3), iw, Inches(0.3),
                           str(BG), angle=90, a1=0, a2=100)
        img_bottom = Inches(2.0) + ih
    caption = s.get('caption', '')
    if attr:
        caption = f"{caption}  |  {attr}".strip(' | ')
    if caption:
        tb, tf = _add_textbox(slide, Inches(0.9), img_bottom + Inches(0.06), Inches(11.5), Inches(0.4))
        _add_para(tf, caption, font=YAHEI, size=11.5, color=GREY,
                  align=PP_ALIGN.CENTER, first=True, space_after=0)
    _add_takeaway(slide, s.get('takeaway'))
    _add_footer(slide, page)
    return slide


def build_closing(prs, s, page=None):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_background(slide)
    # 可选：整页温暖背景图 + 浅色蒙层（保持文字可读，同时温暖有图）
    img_path, _ = _resolve_image(s)
    if img_path and os.path.exists(img_path):
        _add_fullbleed(slide, img_path)
        _add_gradient_rect(slide, 0, 0, SLIDE_W, SLIDE_H, 'FFFFFF', angle=0, a1=80, a2=80)
    _add_topbar(slide)
    _add_rect(slide, 0, Inches(7.38), SLIDE_W, Inches(0.12), fill=NAVY)
    text = s.get('text', '谢谢观看')
    sub = s.get('subtitle', '')
    # 主文字
    tb, tf = _add_textbox(slide, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.9))
    _add_para(tf, text, font=TITLE_FONT, size=40, bold=True, color=NAVY,
              align=PP_ALIGN.CENTER, first=True, space_after=0)
    if sub:
        tb2, tf2 = _add_textbox(slide, Inches(0.5), Inches(2.1), Inches(12.3), Inches(0.6))
        _add_para(tf2, sub, font=BODY_FONT, size=18, color=GREY,
                  align=PP_ALIGN.CENTER, first=True, space_after=0)
    # 红 + 金 双色分隔线
    _add_rect(slide, Inches(5.9), Inches(1.98), Inches(1.5), Pt(2.5), fill=RED)
    _add_rect(slide, Inches(5.9), Inches(2.1), Inches(1.5), Pt(1.2), fill=ORANGE)
    # 核心回顾数字卡（铺满中段）
    stats = s.get('stats', [])
    if stats:
        sy = Inches(2.55)
        ncol = len(stats)
        gap = Inches(0.25)
        cw = (Inches(12.43) - gap * (ncol - 1)) / ncol
        ch = Inches(1.3)
        for i, st in enumerate(stats):
            x = Inches(MARGIN_X) + i * (cw + gap)
            _card(slide, x, sy, cw, ch, accent=NAVY)
            n = st.get('n', '') if isinstance(st, dict) else str(st)
            label = st.get('label', '') if isinstance(st, dict) else ''
            tb, tf = _add_textbox(slide, x + Inches(0.15), sy + Inches(0.1), cw - Inches(0.3), Inches(0.6))
            _add_para(tf, n, font=TITLE_FONT, size=30, bold=True, color=ORANGE,
                      align=PP_ALIGN.CENTER, first=True, space_after=0)
            tb2, tf2 = _add_textbox(slide, x + Inches(0.15), sy + Inches(0.74), cw - Inches(0.3), Inches(0.42))
            _add_para(tf2, label, font=YAHEI, size=12, color=GREY,
                      align=PP_ALIGN.CENTER, first=True, space_after=0)
    # 关键行动条
    actions = s.get('actions', [])
    if actions:
        y = Inches(4.15)
        h = Inches(1.95)
        _add_rect(slide, Inches(MARGIN_X), y, Inches(12.43), h, fill=LIGHT)
        cols = len(actions)
        cw = Inches(12.43) / cols
        for i, a in enumerate(actions):
            x = Inches(MARGIN_X) + i * cw
            if i > 0:
                _add_rect(slide, x, y + Inches(0.4), Pt(1), h - Inches(0.8), fill=RGBColor(0xC9, 0xD2, 0xDE))
            head = a.get('head', '') if isinstance(a, dict) else ''
            atext = a.get('text', '') if isinstance(a, dict) else a
            tb, tf = _add_textbox(slide, x + Inches(0.32), y + Inches(0.24), cw - Inches(0.62), Inches(0.5))
            _add_para(tf, f"{i + 1:02d}  {head}", font=BODY_FONT, size=16, bold=True,
                      color=ORANGE, first=True, space_after=0)
            tb2, tf2 = _add_textbox(slide, x + Inches(0.32), y + Inches(0.78), cw - Inches(0.62), Inches(0.95))
            _add_para(tf2, atext, font=BODY_FONT, size=13, color=NAVY,
                      first=True, space_after=0, line_spacing=1.2)
    # 底部署名/说明
    note = s.get('note', '')
    if note:
        tb, tf = _add_textbox(slide, Inches(0.5), Inches(6.35), Inches(12.3), Inches(0.5))
        _add_para(tf, note, font=YAHEI, size=11.5, color=GREY,
                  align=PP_ALIGN.CENTER, first=True, space_after=0)
    _add_footer(slide, page)
    return slide


BUILDERS = {
    'toc': build_toc,
    'section': build_section,
    'bullets': build_bullets,
    'table': build_table,
    'chart': build_chart,
    'image': build_image,
    'process': build_process,
    'closing': build_closing,
}


def _blank_layout(prs):
    """选一个接近空白的 layout（去掉标题占位符干扰）。"""
    for ly in prs.slide_layouts:
        if 'blank' in (ly.name or '').lower() or '空白' in (ly.name or ''):
            return ly
    return prs.slide_layouts[-1]


def _delete_all_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.get(qn('r:id'))
        if rId:
            try:
                prs.part.drop_rel(rId)
            except Exception:
                pass
        sldIdLst.remove(sldId)


def generate(spec, template_path, out_path):
    meta = spec.get('meta', {})
    _apply_palette(meta.get('palette', 'original'))
    global DECK_TITLE
    DECK_TITLE = (meta.get('title') or '').strip()

    prs = Presentation(template_path)
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    _delete_all_slides(prs)

    slides = spec.get('slides', [])
    for page, s in enumerate(slides, 1):
        stype = s.get('type', 'section')
        if stype == 'cover':
            build_cover(prs, meta, page)
        else:
            builder = BUILDERS.get(stype, build_section)
            builder(prs, s, page)
        notes = s.get('notes') if stype != 'cover' else meta.get('notes')
        if notes:
            try:
                prs.slides[-1].notes_slide.notes_text_frame.text = notes
            except Exception:
                pass

    prs.save(out_path)
    return out_path


def _load_config():
    """读取本目录 config.json（可选），失败返回空 dict。"""
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
    try:
        return json.load(open(p, encoding='utf-8'))
    except Exception:
        return {}


def _resolve_path(p):
    """把 config 里的相对路径解析到脚本所在目录；空值返回空串。"""
    if not p:
        return ''
    if os.path.isabs(p):
        return p
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), p)


def _default_template():
    """默认模板：config.json 的 template → 内置 assets/base_template.pptx。"""
    t = _resolve_path(_load_config().get('template', ''))
    if t and os.path.exists(t):
        return t
    t = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'base_template.pptx')
    return t if os.path.exists(t) else ''


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("用法: python3 generate_ppt.py spec.json out.pptx [--template t.pptx]")
        sys.exit(1)
    spec_path = sys.argv[1]
    out_path = sys.argv[2]
    template = _default_template()
    if '--template' in sys.argv:
        template = sys.argv[sys.argv.index('--template') + 1]
    if not template or not os.path.exists(template):
        print("错误: 未找到模板 .pptx。请用 --template 指定，或在 config.json 的 template 里配置。")
        sys.exit(1)
    spec = json.load(open(spec_path, encoding='utf-8'))
    generate(spec, template, out_path)
    print(f"已生成: {out_path}")
