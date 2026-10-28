#!/usr/bin/env python3
"""
庄重学术风 PPT 生成器 —— 藏蓝 #1F3A5F 主色 + 金 #C9A227 辅色 + 白/浅灰背景 + 微软雅黑，少图多几何。
适合机构/学术/医学等正式汇报场景。
"""
import sys, os, json
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn


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

# ---------- 国企风格配色（藏蓝 + 金 + 中国红，庄重大气） ----------
TEAL  = RGBColor(0x1F, 0x3A, 0x5F)   # 主色 藏蓝
TEAL_D = RGBColor(0x16, 0x2F, 0x4A)  # 藏蓝(暗，标题用)
AMBER = RGBColor(0xC9, 0xA2, 0x27)   # 辅色 金色
AMBER_L = RGBColor(0xF6, 0xF0, 0xE0) # 浅金底
RED   = RGBColor(0xC0, 0x39, 0x2B)   # 强调 中国红
DARK  = RGBColor(0x2B, 0x2B, 0x2B)   # 正文深灰
GREY  = RGBColor(0x3D, 0x4C, 0x5E)   # 次要深蓝灰
LIGHT = RGBColor(0xEE, 0xF1, 0xF4)   # 浅灰底
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEAL_L = RGBColor(0xE6, 0xEC, 0xF2)  # 浅蓝灰底

TITLE_FONT = "微软雅黑"
BODY_FONT  = "微软雅黑"
EN_FONT    = "Arial"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FOOTER = ""   # 页脚文字默认留空；可在 spec 的 meta.footer 里指定

# ---------- 基础 ----------
def _set_run(run, text, font=BODY_FONT, size=16, bold=False, color=DARK, italic=False):
    run.text = text
    f = run.font
    f.name = EN_FONT          # latin（英文/数字）
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', font)  # ea（中文）

def _tb(slide, x, y, w, h, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    for m in ('margin_left','margin_right','margin_top','margin_bottom'):
        setattr(tf, m, 0)
    return tb, tf

def _para(tf, text, font=BODY_FONT, size=16, bold=False, color=DARK,
          align=PP_ALIGN.LEFT, first=False, space_after=6, line_spacing=1.2):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    r = p.add_run()
    _set_run(r, text, font=font, size=size, bold=bold, color=color)
    return p

def _rect(slide, x, y, w, h, fill=None, line=None, line_w=1.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
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

def _round_rect(slide, x, y, w, h, fill=None, line=None, radius=0.06):
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
        shp.line.color.rgb = line; shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp

def _shape(slide, kind, x, y, w, h, fill=None, line=None):
    shp = slide.shapes.add_shape(kind, x, y, w, h)
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp

def _blank(prs):
    """选接近空白的版式（模板背景由母版继承）。"""
    for ly in prs.slide_layouts:
        if 'blank' in (ly.name or '').lower() or '空白' in (ly.name or ''):
            return ly
    return prs.slide_layouts[-1]

BASE_BG = RGBColor(0xF6, 0xF7, 0xF9)   # 极浅灰底（覆盖模板装饰底纹，白色卡片可浮起）

def _background(slide):
    """干净浅蓝背景（覆盖模板复杂底纹，保留浅蓝基调，避免装饰干扰内容）。"""
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=BASE_BG)


def _logo_path():
    """logo 图片路径：config.json 的 logo → 内置 assets/logo.png。"""
    p = _resolve_path(_load_config().get('logo', ''))
    if p:
        return p
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'logo.png')

LOGO_PATH = _logo_path()

def _add_logo(slide):
    """左上角 logo（清晰、不被遮挡）；仅当 LOGO_PATH 存在时放置。"""
    if os.path.exists(LOGO_PATH):
        from PIL import Image
        im = Image.open(LOGO_PATH)
        iw, ih = im.size
        h = Inches(0.46)
        w = Emu(int(iw / ih * 0.46 * 914400))
        slide.shapes.add_picture(LOGO_PATH, Inches(0.42), Inches(0.3), width=w, height=h)

def _new_slide(prs):
    """新建空白页，用极浅灰底覆盖模板装饰底纹，得到干净学术背景。"""
    s = prs.slides.add_slide(_blank(prs))
    _background(s)
    _add_logo(s)   # 可选 logo：仅当 LOGO_PATH 存在时自动置于左上角
    return s

# ---------- 页眉/页脚 ----------
def _header(slide, title, chapter=None, page=None):
    # 章节标签（下移避开模板左上角 logo）
    if chapter:
        tb, tf = _tb(slide, Inches(0.5), Inches(0.98), Inches(11.0), Inches(0.3))
        _para(tf, chapter, size=13, bold=True, color=AMBER, first=True, space_after=0)
    y = Inches(1.24)
    _rect(slide, Inches(0.5), y, Inches(0.09), Inches(0.62), fill=TEAL)
    tb, tf = _tb(slide, Inches(0.72), y, Inches(11.6), Inches(0.7))
    _para(tf, title, size=30, bold=True, color=TEAL_D, first=True, space_after=0)
    # 标题下方金/红菱形点缀（国企风格三色点缀）
    _shape(slide, MSO_SHAPE.DIAMOND, Inches(0.47), Inches(1.9), Inches(0.13), Inches(0.13), fill=AMBER)
    _shape(slide, MSO_SHAPE.DIAMOND, Inches(0.66), Inches(1.9), Inches(0.13), Inches(0.13), fill=RED)
    _footer(slide, page)

def _footer(slide, page=None):
    # 底部浅灰条垫底，页脚文字清晰
    _rect(slide, 0, Inches(7.02), SLIDE_W, Inches(0.48), fill=LIGHT)
    y = Inches(7.08)
    _rect(slide, Inches(0.5), y, Inches(12.33), Pt(1), fill=RGBColor(0xD8, 0xDE, 0xE4))
    tb, tf = _tb(slide, Inches(0.5), Inches(7.16), Inches(11.0), Inches(0.3))
    _para(tf, FOOTER, size=10.5, color=RGBColor(0x31, 0x3F, 0x4E), first=True, space_after=0)
    if page is not None:
        tb2, tf2 = _tb(slide, Inches(12.3), Inches(7.16), Inches(0.6), Inches(0.3))
        _para(tf2, str(page), size=10.5, color=RGBColor(0x31, 0x3F, 0x4E), align=PP_ALIGN.RIGHT, first=True, space_after=0)

def _icon(slide, x, y, idx=0, size=0.16, kind='circle'):
    colors = [TEAL, AMBER, TEAL_D]
    c = colors[idx % 3]
    if kind == 'diamond':
        shp = _shape(slide, MSO_SHAPE.DIAMOND, x, y, Inches(size), Inches(size), fill=c)
    elif kind == 'square':
        shp = _shape(slide, MSO_SHAPE.RECTANGLE, x, y, Inches(size), Inches(size), fill=c)
    else:
        shp = _shape(slide, MSO_SHAPE.OVAL, x, y, Inches(size), Inches(size), fill=c)
    return shp

def _bullet(slide, x, y, w, text, idx=0, size=16, color=DARK, kind='circle'):
    _icon(slide, x, y + Inches(0.05), idx, 0.14, kind)
    tb, tf = _tb(slide, x + Inches(0.28), y, w - Inches(0.28), Inches(0.5))
    _para(tf, text, size=size, color=color, first=True, space_after=0, line_spacing=1.15)
    return tf

IMG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output', 'img')

def _add_framed_image(slide, name, x, y, w, h, caption=None):
    """插入切题插图：等比居中，不加边框（自然融入）。"""
    path = os.path.join(IMG_DIR, name + '.jpg')
    if not os.path.exists(path):
        return
    from PIL import Image
    im = Image.open(path)
    iw, ih = im.size
    ratio = min(w / iw, h / ih)
    pw, ph = Emu(int(iw * ratio)), Emu(int(ih * ratio))
    px = Emu(int(x + (w - pw) / 2)); py = Emu(int(y + (h - ph) / 2))
    slide.shapes.add_picture(path, px, py, width=pw, height=ph)

def _gradient_rect(slide, x, y, w, h, rgb1, rgb2, angle=90, a1=100, a2=0):
    """线性渐变矩形（rgb 为 'RRGGBB'，a 为不透明度 0-100，angle 度）。"""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.line.fill.background(); shp.shadow.inherit = False
    spPr = shp._element.spPr
    solid = spPr.find(qn('a:solidFill'))
    grad = spPr.makeelement(qn('a:gradFill'), {})
    gsLst = spPr.makeelement(qn('a:gsLst'), {})
    for pos, rgb, a in ((0, rgb1, a1), (100000, rgb2, a2)):
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

def _add_image_fade(slide, name, x, y, w, h, bg='FFFFFF', direction='top'):
    """图片：固定宽度、等比缩放（不裁剪、不拉伸、不加渐变），自然清晰。"""
    path = os.path.join(IMG_DIR, name + '.jpg')
    if not os.path.exists(path):
        return
    from PIL import Image
    im = Image.open(path); iw, ih = im.size
    # 等比缩放并约束在 (w,h) 框内（contain，不裁剪不变形，永不溢出边界）
    ratio = min(w / iw, h / ih)
    pw = Emu(int(iw * ratio)); ph = Emu(int(ih * ratio))
    px = Emu(int(x + (w - pw) / 2)); py = Emu(int(y + (h - ph) / 2))
    slide.shapes.add_picture(path, px, py, width=pw, height=ph)

def _add_side_image(slide, name, x, y, w, h, caption=None, accent='gold'):
    """右侧配图：圆角 + 偏移强调框 + 底部图注条（有设计感，不生硬）。"""
    path = os.path.join(IMG_DIR, name + '.jpg')
    if not os.path.exists(path):
        return
    from PIL import Image, ImageDraw
    im = Image.open(path).convert('RGB')
    iw, ih = im.size
    box_ratio = w / h; img_ratio = iw / ih
    if img_ratio > box_ratio:
        cw = int(ih * box_ratio); left = (iw - cw) // 2; im = im.crop((left, 0, left + cw, ih))
    else:
        ch = int(iw / box_ratio); top = (ih - ch) // 2; im = im.crop((0, top, iw, top + ch))
    # 圆角（透明角）
    rad = max(6, int(min(im.size) * 0.04))
    mask = Image.new('L', im.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, im.size[0], im.size[1]], radius=rad, fill=255)
    out = Image.new('RGBA', im.size, (0, 0, 0, 0))
    out.paste(im, (0, 0), mask)
    tmp = os.path.join('/tmp', f'ppt_rnd_{name}_{int(x/914400)}.png')
    out.save(tmp, 'PNG')
    # 偏移强调框（右下偏移金色衬底，形成层叠“相框”感）
    _rect(slide, x + Inches(0.15), y + Inches(0.15), w, h, fill=AMBER)
    slide.shapes.add_picture(tmp, x, y, width=w, height=h)
    # 底部图注条（圆角 + 白字）
    if caption:
        cy = y + h - Inches(0.52)
        _round_rect(slide, x + Inches(0.12), cy, w - Inches(0.24), Inches(0.4), fill=TEAL_D, radius=0.12)
        tb, tf = _tb(slide, x + Inches(0.28), cy + Inches(0.03), w - Inches(0.55), Inches(0.34))
        _para(tf, caption, size=11, bold=True, color=WHITE, first=True, space_after=0)

def _add_bottom_image(slide, name, x, y, w, h, caption=None, accent='gold', sub=None):
    """底部横幅：左侧圆角照片 + 右侧藏蓝图注面板（图注+要点，不覆盖照片主体）。"""
    path = os.path.join(IMG_DIR, name + '.jpg')
    if not os.path.exists(path):
        return
    from PIL import Image, ImageDraw
    im = Image.open(path).convert('RGB')
    iw, ih = im.size
    photo_w = w * 0.58
    box_ratio = photo_w / h; img_ratio = iw / ih
    if img_ratio > box_ratio:
        cw = int(ih * box_ratio); left = (iw - cw) // 2; im = im.crop((left, 0, left + cw, ih))
    else:
        ch = int(iw / box_ratio); top = (ih - ch) // 2; im = im.crop((0, top, iw, top + ch))
    rad = max(6, int(min(im.size) * 0.04))
    mask = Image.new('L', im.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, im.size[0], im.size[1]], radius=rad, fill=255)
    out = Image.new('RGBA', im.size, (0, 0, 0, 0))
    out.paste(im, (0, 0), mask)
    tmp = os.path.join('/tmp', f'ppt_rnd_{name}_{int(x/914400)}_{int(y/914400)}.png')
    out.save(tmp, 'PNG')
    slide.shapes.add_picture(tmp, x, y, width=photo_w, height=h)
    # 右侧藏蓝图注面板（金色/藏蓝竖条 + 白字图注，垂直居中）
    px = x + photo_w + Inches(0.22)
    pw = w - photo_w - Inches(0.22)
    _round_rect(slide, px, y, pw, h, fill=TEAL_D, radius=0.08)
    accent_c = AMBER if accent == 'gold' else TEAL
    _rect(slide, px, y, Inches(0.09), h, fill=accent_c)
    if caption:
        tb, tf = _tb(slide, px + Inches(0.34), y, pw - Inches(0.6), h)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        _para(tf, caption, size=14, bold=True, color=WHITE, first=True, space_after=2, line_spacing=1.1)
        if sub:
            _para(tf, sub, size=11, color=RGBColor(0xBE, 0xCE, 0xE4), space_after=0, line_spacing=1.15)

def _add_cover_bleed(slide, name, bg='FFFFFF'):
    """整页配图 + 多段渐变蒙层（左半可读、右半见图、平滑过渡）。"""
    path = os.path.join(IMG_DIR, name + '.jpg')
    if not os.path.exists(path):
        return
    from PIL import Image
    im = Image.open(path); iw, ih = im.size
    ratio = max(SLIDE_W / iw, SLIDE_H / ih)
    pw = Emu(int(iw * ratio)); ph = Emu(int(ih * ratio))
    px = Emu(int((SLIDE_W - pw) / 2)); py = Emu(int((SLIDE_H - ph) / 2))
    slide.shapes.add_picture(path, px, py, width=pw, height=ph)
    # 多段渐变蒙层：左 42% 不透明 → 65% 处全透明
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    shp.line.fill.background(); shp.shadow.inherit = False
    spPr = shp._element.spPr
    solid = spPr.find(qn('a:solidFill'))
    grad = spPr.makeelement(qn('a:gradFill'), {})
    gsLst = spPr.makeelement(qn('a:gsLst'), {})
    for pos, a in ((0, 100), (42000, 100), (68000, 0), (100000, 0)):
        gs = spPr.makeelement(qn('a:gs'), {'pos': str(pos)})
        srgb = spPr.makeelement(qn('a:srgbClr'), {'val': bg})
        if a < 100:
            al = spPr.makeelement(qn('a:alpha'), {'val': str(int(a * 1000))})
            srgb.append(al)
        gs.append(srgb); gsLst.append(gs)
    grad.append(gsLst)
    lin = spPr.makeelement(qn('a:lin'), {'ang': '0', 'scaled': '1'})
    grad.append(lin)
    if solid is not None:
        solid.addprevious(grad); spPr.remove(solid)
    else:
        spPr.append(grad)

# ---------- 版式 ----------
def build_cover(prs, meta):
    s = _new_slide(prs)
    # 顶部与底部藏蓝细条（原版纯文字封面）
    _rect(s, 0, 0, SLIDE_W, Inches(0.14), fill=TEAL)
    _rect(s, 0, Inches(7.36), SLIDE_W, Inches(0.14), fill=TEAL)
    # 左侧金色竖条装饰
    _rect(s, Inches(0.9), Inches(1.5), Inches(0.14), Inches(3.6), fill=AMBER)
    # 主标题（全宽，避开左上 logo，33pt 两行）
    tb, tf = _tb(s, Inches(1.35), Inches(1.5), Inches(11.3), Inches(2.0))
    for li, line in enumerate(meta['title'].split('\n')):
        _para(tf, line, font=TITLE_FONT, size=33, bold=True, color=TEAL_D,
              first=(li == 0), space_after=4, line_spacing=1.15)
    # 副标题
    tb, tf = _tb(s, Inches(1.35), Inches(3.6), Inches(11.0), Inches(0.6))
    _para(tf, meta['subtitle'], size=20, bold=False, color=DARK, first=True, space_after=0)
    # 金色分隔线
    _rect(s, Inches(1.35), Inches(4.45), Inches(2.2), Pt(2.5), fill=AMBER)
    # 来源
    y = Inches(4.7)
    for line in meta.get('source', []):
        tb, tf = _tb(s, Inches(1.35), y, Inches(11.0), Inches(0.4))
        _para(tf, line, size=13, color=DARK, first=True, space_after=0)
        y += Inches(0.4)
    # 占位
    tb, tf = _tb(s, Inches(1.35), Inches(6.4), Inches(11.0), Inches(0.4))
    _para(tf, meta.get('placeholder', '汇报人：＿＿＿    日期：＿＿＿'), size=13, color=DARK, first=True, space_after=0)
    # 右侧几何点缀（国企三色菱形阵列，平衡左侧文字，保持无图纯文字风格）
    _dots = [
        (11.9, 1.7, AMBER, 0.2), (11.35, 2.45, TEAL, 0.15), (11.85, 3.2, RED, 0.18),
        (11.25, 3.95, AMBER, 0.13), (11.7, 4.7, TEAL, 0.2), (11.1, 5.45, RED, 0.15),
        (11.6, 6.15, AMBER, 0.18),
    ]
    for dx, dy, dc, ds in _dots:
        _shape(s, MSO_SHAPE.DIAMOND, Inches(dx), Inches(dy), Inches(ds), Inches(ds), fill=dc)
    return s

def build_toc(prs, slide_spec):
    s = _new_slide(prs)
    _header(s, slide_spec.get('title', '目录'), page=slide_spec.get('page', 2))
    items = slide_spec.get('items', [])
    pages = slide_spec.get('pages', [])
    # 两栏，纵向舒展
    col_w = Inches(5.9)
    for i, it in enumerate(items):
        col = i % 2
        row = i // 2
        x = Inches(0.6) + col * Inches(6.3)
        y = Inches(2.35) + row * Inches(1.32)
        # 序号方块
        box = _rect(s, x, y, Inches(0.55), Inches(0.55), fill=TEAL)
        btf = box.text_frame; btf.word_wrap = False
        bp = btf.paragraphs[0]; bp.alignment = PP_ALIGN.CENTER
        br = bp.add_run(); _set_run(br, f"{i+1:02d}", size=17, bold=True, color=WHITE)
        # 标题 + 右侧页码
        tb, tf = _tb(s, x + Inches(0.75), y + Inches(0.02), col_w - Inches(1.6), Inches(0.55))
        _para(tf, it, size=18, bold=True, color=TEAL_D, first=True, space_after=0)
        if i < len(pages):
            tb2, tf2 = _tb(s, x + col_w - Inches(1.15), y + Inches(0.06), Inches(0.95), Inches(0.4))
            _para(tf2, str(pages[i]), size=15, bold=True, color=AMBER, align=PP_ALIGN.RIGHT, first=True, space_after=0)
    # 底部说明（点明模块数、用途与时长）
    _rect(s, Inches(0.6), Inches(6.25), Inches(12.13), Pt(1), fill=RGBColor(0xD8, 0xDE, 0xE4))
    tb, tf = _tb(s, Inches(0.6), Inches(6.42), Inches(12.13), Inches(0.35))
    _para(tf, "6 大模块 · 本场约 15~20 分钟", size=12.5, color=GREY, align=PP_ALIGN.CENTER, first=True, space_after=0)
    return s

def build_two_col(prs, s):
    slide = _new_slide(prs)
    _header(slide, s.get('title'), s.get('chapter'), s.get('page'))
    cols = s.get('cols', [])
    n = len(cols)
    if n == 0:
        return
    # 左文右图：有配图时，右侧整高配图（圆角+偏移框+图注）+ 左侧单栏上下堆叠内容
    if s.get('image'):
        _add_side_image(slide, s['image'], Inches(8.45), Inches(2.1), Inches(4.35), Inches(4.7),
                        caption=s.get('caption'), accent=s.get('accent', 'gold'))
        content_w = Inches(7.55)
        x0 = Inches(0.5)
        y = Inches(2.15)
        for ci, col in enumerate(cols):
            heading = col.get('heading', '')
            points = col.get('points', [])
            _rect(slide, x0, y, Inches(0.08), Inches(0.3), fill=AMBER)
            tb, tf = _tb(slide, x0 + Inches(0.22), y, content_w - Inches(0.22), Inches(0.4))
            _para(tf, heading, size=17, bold=True, color=TEAL_D, first=True, space_after=0)
            py = y + Inches(0.46)
            for pi, pt in enumerate(points):
                _bullet(slide, x0 + Inches(0.22), py, content_w - Inches(0.22), pt, pi, size=13)
                py += Inches(0.46)
            y = py + Inches(0.14)
        return slide
    # 无配图：左右两栏
    col_w = Inches((12.33 - 0.3 * (n - 1)) / n)
    x0 = Inches(0.5)
    y0 = Inches(2.2)
    for ci, col in enumerate(cols):
        x = x0 + ci * (col_w + Inches(0.3))
        heading = col.get('heading', '')
        points = col.get('points', [])
        _rect(slide, x, y0, Inches(0.08), Inches(0.34), fill=AMBER)
        tb, tf = _tb(slide, x + Inches(0.22), y0, col_w - Inches(0.22), Inches(0.45))
        _para(tf, heading, size=19, bold=True, color=TEAL_D, first=True, space_after=0)
        py = y0 + Inches(0.62)
        for pi, pt in enumerate(points):
            _bullet(slide, x + Inches(0.22), py, col_w - Inches(0.22), pt, pi)
            py += Inches(0.92)
    return slide

def build_data_band(prs, s):
    """左侧数据带 + 右侧要点（可选右图，三栏版式）。"""
    slide = _new_slide(prs)
    _header(slide, s.get('title'), s.get('chapter'), s.get('page'))
    band = s.get('band', [])   # 数据带：[(数字, 标签)]
    right = s.get('right', {}) # {'heading','points'}
    has_img = bool(s.get('image'))
    if has_img:
        _add_side_image(slide, s['image'], Inches(8.45), Inches(2.1), Inches(4.35), Inches(4.7),
                        caption=s.get('caption'), accent=s.get('accent', 'gold'))
        band_x, band_w = Inches(0.5), Inches(3.0)
        rx, rw = Inches(3.85), Inches(4.35)
    else:
        band_x, band_w = Inches(0.5), Inches(3.6)
        rx, rw = Inches(4.5), Inches(8.2)
    # 左侧藏蓝数据带
    _rect(slide, band_x, Inches(2.2), band_w, Inches(4.5), fill=TEAL)
    y = Inches(2.45)
    for num, label in band:
        tb, tf = _tb(slide, band_x + Inches(0.25), y, band_w - Inches(0.5), Inches(0.55))
        _para(tf, num, size=21 if has_img else 26, bold=True, color=WHITE, first=True, space_after=0)
        tb, tf = _tb(slide, band_x + Inches(0.25), y + Inches(0.44), band_w - Inches(0.5), Inches(0.9))
        _para(tf, label, size=10.5 if has_img else 12, color=RGBColor(0xDD,0xEE,0xEC), first=True, space_after=0, line_spacing=1.1)
        y += Inches(1.45)
    # 右侧要点
    heading = right.get('heading', '')
    if heading:
        _rect(slide, rx, Inches(2.25), Inches(0.08), Inches(0.34), fill=AMBER)
        tb, tf = _tb(slide, rx + Inches(0.22), Inches(2.25), rw - Inches(0.22), Inches(0.45))
        _para(tf, heading, size=17 if has_img else 19, bold=True, color=TEAL_D, first=True, space_after=0)
    py = Inches(2.9)
    for pi, pt in enumerate(right.get('points', [])):
        _bullet(slide, rx + Inches(0.22), py, rw - Inches(0.22), pt, pi, size=12.5 if has_img else 16)
        py += Inches(0.72)
    # 数据来源脚注
    if s.get('note'):
        tb, tf = _tb(slide, rx, Inches(6.4), rw + Inches(0.2), Inches(0.3))
        _para(tf, "注：" + s['note'], size=9.5 if has_img else 10.5, color=GREY, first=True, space_after=0)
    return slide

def build_timeline(prs, s):
    """横向时间线/流程 + 可选右侧配图。"""
    slide = _new_slide(prs)
    _header(slide, s.get('title'), s.get('chapter'), s.get('page'))
    steps = s.get('steps', [])
    n = len(steps)
    if n == 0:
        return
    has_img = bool(s.get('image'))
    if has_img:
        _add_side_image(slide, s['image'], Inches(8.6), Inches(2.1), Inches(4.2), Inches(4.7),
                        caption=s.get('caption'), accent=s.get('accent', 'gold'))
        lx, lw = Inches(0.7), Inches(7.6)
        data_x0, data_w = Inches(0.5), Inches(7.9)
    else:
        lx, lw = Inches(0.7), Inches(11.9)
        data_x0, data_w = Inches(0.5), Inches(12.33)
    # 横向轴线
    y_line = Inches(2.75)
    _rect(slide, lx, y_line, lw, Pt(2.5), fill=LIGHT)
    step_w = lw / n
    for i, st in enumerate(steps):
        cx = lx + step_w * i + step_w / 2
        # 节点圆
        _shape(slide, MSO_SHAPE.OVAL, cx - Inches(0.14), y_line - Inches(0.14) + Inches(0.05), Inches(0.28), Inches(0.28), fill=TEAL if i % 2 == 0 else AMBER)
        # 时间文字
        tb, tf = _tb(slide, cx - Inches(1.2), Inches(3.05), Inches(2.4), Inches(0.4))
        _para(tf, st.get('time', ''), size=13 if has_img else 14, bold=True, color=TEAL_D, align=PP_ALIGN.CENTER, first=True, space_after=0)
        # 说明
        tb, tf = _tb(slide, cx - Inches(1.4), Inches(3.5), Inches(2.8), Inches(0.9))
        _para(tf, st.get('text', ''), size=12 if has_img else 12.5, color=DARK, align=PP_ALIGN.CENTER, first=True, space_after=0, line_spacing=1.15)
    # 底部补充（方法学依据）
    if s.get('note'):
        tb, tf = _tb(slide, lx, Inches(4.4), lw, Inches(0.5))
        _para(tf, s['note'], size=11.5 if has_img else 13.5, color=GREY, first=True, space_after=0, line_spacing=1.25)
    # 底部委员会数据栏
    if s.get('data'):
        y2 = Inches(5.35)
        _rect(slide, data_x0, y2, data_w, Inches(0.06), fill=AMBER)
        data = s['data']; dn = len(data)
        for di, (val, label) in enumerate(data):
            x = data_x0 + Inches(0.1) + di * (data_w / dn)
            tb, tf = _tb(slide, x, y2 + Inches(0.2), data_w / dn - Inches(0.25), Inches(0.4))
            _para(tf, val, size=15 if has_img else 18, bold=True, color=TEAL_D, first=True, space_after=0)
            tb, tf = _tb(slide, x, y2 + Inches(0.56), data_w / dn - Inches(0.25), Inches(0.9))
            _para(tf, label, size=10.5 if has_img else 11.5, color=GREY, first=True, space_after=0, line_spacing=1.1)
    return slide

def build_funnel(prs, s):
    """漏斗图 + 构成。"""
    slide = _new_slide(prs)
    _header(slide, s.get('title'), s.get('chapter'), s.get('page'))
    levels = s.get('levels', [])   # [('27 问题', 宽比例), ...]
    n = len(levels)
    if n == 0:
        return
    y = Inches(2.2)
    h = Inches(0.62)
    max_w = Inches(6.0)
    for i, (label, ratio) in enumerate(levels):
        w = max_w * ratio
        x = Inches(0.7) + (max_w - w) / 2
        _rect(slide, x, y, w, h, fill=TEAL if i % 2 == 0 else AMBER)
        tbf = slide.shapes.add_textbox(x, y + Inches(0.12), w, Inches(0.4)).text_frame
        tbf.word_wrap = False
        p = tbf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); _set_run(r, label, size=15, bold=True, color=WHITE)
        y += Inches(0.8)
    # 右侧构成
    if s.get('composition'):
        tb, tf = _tb(slide, Inches(7.4), Inches(2.2), Inches(5.4), Inches(0.4))
        _para(tf, s['composition'].get('heading', '文献构成'), size=18, bold=True, color=TEAL_D, first=True, space_after=0)
        py = Inches(2.8)
        for pi, (label, val) in enumerate(s['composition'].get('items', [])):
            _icon(slide, Inches(7.4), py + Inches(0.05), pi, 0.14)
            tb, tf = _tb(slide, Inches(7.7), py, Inches(5.1), Inches(0.5))
            _para(tf, f"{label}", size=15, color=DARK, first=True, space_after=0)
            tb2, tf2 = _tb(slide, Inches(10.6), py, Inches(1.9), Inches(0.5))
            _para(tf2, val, size=15, bold=True, color=AMBER, align=PP_ALIGN.RIGHT, first=True, space_after=0)
            py += Inches(0.5)
    # 底部切题插图（左照片 + 右图注面板）
    if s.get('image'):
        _add_bottom_image(slide, s['image'], Inches(0.5), Inches(5.05), Inches(12.33), Inches(1.8),
                          caption=s.get('caption'), accent=s.get('accent', 'gold'), sub=s.get('sub'))
    return slide

def build_process(prs, s):
    """N 步流程卡片 + 关键数据 + 可选底部横幅插图。"""
    slide = _new_slide(prs)
    _header(slide, s.get('title'), s.get('chapter'), s.get('page'))
    steps = s.get('steps', [])
    n = len(steps)
    has_img = bool(s.get('image'))
    card_w = Inches(11.9 / n - 0.25)
    card_h = Inches(2.3) if has_img else Inches(2.6)
    y = Inches(2.2)
    for i, st in enumerate(steps):
        x = Inches(0.5) + i * Inches(12.0 / n)
        _round_rect(slide, x, y, card_w, card_h, fill=None, line=LIGHT, radius=0.06)
        _rect(slide, x, y, card_w, Inches(0.16), fill=TEAL)
        tb, tf = _tb(slide, x + Inches(0.16), y + Inches(0.28), card_w - Inches(0.32), Inches(0.4))
        _para(tf, st.get('n', str(i+1)), size=18 if has_img else 20, bold=True, color=AMBER, first=True, space_after=0)
        tb, tf = _tb(slide, x + Inches(0.16), y + Inches(0.68), card_w - Inches(0.32), Inches(0.5))
        _para(tf, st.get('head', ''), size=14 if has_img else 15, bold=True, color=TEAL_D, first=True, space_after=0)
        tb, tf = _tb(slide, x + Inches(0.16), y + Inches(1.22), card_w - Inches(0.32), Inches(1.0))
        _para(tf, st.get('text', ''), size=11.5 if has_img else 12, color=DARK, first=True, space_after=0, line_spacing=1.12)
    # 关键数据栏
    if s.get('data'):
        y2 = Inches(4.7) if has_img else Inches(5.0)
        _rect(slide, Inches(0.5), y2, Inches(12.33), Inches(0.06), fill=AMBER)
        for di, (val, label) in enumerate(s['data']):
            x = Inches(0.6) + di * Inches(3.1)
            tb, tf = _tb(slide, x, y2 + Inches(0.18), Inches(2.9), Inches(0.4))
            _para(tf, val, size=16 if has_img else 17, bold=True, color=RED if di == 0 else TEAL_D, first=True, space_after=0)
            tb, tf = _tb(slide, x, y2 + Inches(0.55), Inches(2.9), Inches(0.9))
            _para(tf, label, size=11 if has_img else 11.5, color=GREY, first=True, space_after=0, line_spacing=1.1)
    # 底部横幅插图
    if has_img:
        _add_bottom_image(slide, s['image'], Inches(0.5), Inches(5.72), Inches(12.33), Inches(1.2),
                          caption=s.get('caption'), accent=s.get('accent', 'gold'), sub=s.get('sub'))
    # 数据来源脚注
    if s.get('note') and not has_img:
        tb, tf = _tb(slide, Inches(0.5), Inches(6.35), Inches(12.3), Inches(0.3))
        _para(tf, "注：" + s['note'], size=10.5, color=GREY, first=True, space_after=0)
    return slide

def build_matrix(prs, s):
    """主题色块矩阵/条形分布 + 可选右侧配图。"""
    slide = _new_slide(prs)
    _header(slide, s.get('title'), s.get('chapter'), s.get('page'))
    intro = s.get('intro', '')
    has_img = bool(s.get('image'))
    if has_img:
        _add_side_image(slide, s['image'], Inches(8.45), Inches(2.1), Inches(4.35), Inches(4.7),
                        caption=s.get('caption'), accent=s.get('accent', 'gold'))
        area_w = Inches(7.55); x0 = Inches(0.5)
    else:
        area_w = Inches(12.33); x0 = Inches(0.5)
    y = Inches(2.05)
    if intro:
        tb, tf = _tb(slide, x0, y, area_w, Inches(0.5))
        _para(tf, intro, size=14 if has_img else 15, color=GREY, first=True, space_after=0)
        y = Inches(2.6) if has_img else Inches(2.95)
    cells = s.get('cells', [])  # [(主题, 条数)]
    if has_img:
        cell_w = (area_w - Inches(0.2)) / 3
        row_h = Inches(1.18); cell_h = Inches(1.04)
        gap = Inches(0.1)
    else:
        cell_w = Inches(3.9); row_h = Inches(1.32); cell_h = Inches(1.15)
        gap = Inches(0.2)
    # 3列网格
    for i, (label, val) in enumerate(cells):
        col = i % 3
        row = i // 3
        x = x0 + col * (cell_w + gap)
        yy = y + row * row_h
        _round_rect(slide, x, yy, cell_w, cell_h, fill=TEAL_L if (i % 2 == 0) else LIGHT, radius=0.08)
        _rect(slide, x, yy, Inches(0.09), cell_h, fill=TEAL if val and int(val) > 1 else AMBER)
        tb, tf = _tb(slide, x + Inches(0.2), yy + Inches(0.07), cell_w - Inches(0.35), Inches(0.45))
        _para(tf, label, size=14 if has_img else 15, bold=True, color=TEAL_D, first=True, space_after=0)
        tb, tf = _tb(slide, x + Inches(0.2), yy + Inches(0.5), cell_w - Inches(0.35), Inches(0.45))
        _para(tf, f"{val} 条", size=18 if has_img else 20, bold=True, color=AMBER, first=True, space_after=0)
    return slide

def build_table(prs, s):
    slide = _new_slide(prs)
    _header(slide, s.get('title'), s.get('chapter'), s.get('page'))
    headers = s.get('headers', [])
    rows = s.get('rows', [])
    nr = len(rows) + 1
    nc = len(headers)
    if nc == 0:
        return
    gt = slide.shapes.add_table(nr, nc, Inches(0.5), Inches(2.2), Inches(12.33), Inches(0.55))
    table = gt.table
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = h
        cell.fill.solid(); cell.fill.fore_color.rgb = TEAL
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                _set_run(r, h, size=13, bold=True, color=WHITE)
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row):
            cell = table.cell(ri, ci)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT if ri % 2 == 0 else WHITE
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
                for r in p.runs:
                    _set_run(r, str(val), size=12, color=DARK)
    # 表注
    if s.get('note'):
        tb, tf = _tb(slide, Inches(0.5), Inches(6.3), Inches(12.3), Inches(0.5))
        _para(tf, s['note'], size=12, color=GREY, first=True, space_after=0)
    return slide

def build_cards(prs, s):
    """策略卡片（两栏或三栏）+ 可选底部横幅插图。"""
    slide = _new_slide(prs)
    _header(slide, s.get('title'), s.get('chapter'), s.get('page'))
    cards = s.get('cards', [])
    n = len(cards)
    if n == 0:
        return
    has_img = bool(s.get('image'))
    card_h = Inches(2.7) if has_img else Inches(3.7)
    card_w = Inches(12.33 / n - 0.25)
    y = Inches(2.2)
    for i, c in enumerate(cards):
        x = Inches(0.5) + i * Inches(12.4 / n)
        _round_rect(slide, x, y, card_w, card_h, fill=WHITE, line=LIGHT, radius=0.06)
        _rect(slide, x, y, card_w, Inches(0.14), fill=TEAL)
        tb, tf = _tb(slide, x + Inches(0.2), y + Inches(0.28), card_w - Inches(0.4), Inches(0.5))
        _para(tf, c.get('head', ''), size=16 if has_img else 17, bold=True, color=TEAL_D, first=True, space_after=0)
        py = y + Inches(0.9)
        for pi, pt in enumerate(c.get('points', [])):
            _bullet(slide, x + Inches(0.2), py, card_w - Inches(0.4), pt, pi, size=13 if has_img else 14.5)
            py += Inches(0.56)
    # 底部横幅插图（左照片 + 右图注面板）
    if has_img:
        _add_bottom_image(slide, s['image'], Inches(0.5), Inches(5.05), Inches(12.33), Inches(1.8),
                          caption=s.get('caption'), accent=s.get('accent', 'gold'), sub=s.get('sub'))
    return slide

def build_checklist(prs, s):
    """清单 + 沟通要点 + 可选底部横幅插图。"""
    slide = _new_slide(prs)
    _header(slide, s.get('title'), s.get('chapter'), s.get('page'))
    left = s.get('left', {})
    right = s.get('right', {})
    has_img = bool(s.get('image'))
    box_h = Inches(3.2) if has_img else Inches(4.6)
    # 左：清单
    _round_rect(slide, Inches(0.5), Inches(2.2), Inches(6.0), box_h, fill=TEAL_L, radius=0.05)
    tb, tf = _tb(slide, Inches(0.75), Inches(2.32), Inches(5.5), Inches(0.45))
    _para(tf, left.get('heading', ''), size=16 if has_img else 17, bold=True, color=TEAL_D, first=True, space_after=0)
    py = Inches(2.9)
    for pi, pt in enumerate(left.get('points', [])):
        _icon(slide, Inches(0.75), py + Inches(0.04), pi, 0.13)
        tb, tf = _tb(slide, Inches(1.05), py, Inches(5.2), Inches(0.55))
        _para(tf, pt, size=13 if has_img else 13.5, color=DARK, first=True, space_after=0)
        py += Inches(0.74 if has_img else 0.82)
    # 右：沟通要点
    _round_rect(slide, Inches(6.8), Inches(2.2), Inches(6.03), box_h, fill=LIGHT, radius=0.05)
    tb, tf = _tb(slide, Inches(7.05), Inches(2.32), Inches(5.5), Inches(0.45))
    _para(tf, right.get('heading', ''), size=16 if has_img else 17, bold=True, color=TEAL_D, first=True, space_after=0)
    py = Inches(2.9)
    for pi, pt in enumerate(right.get('points', [])):
        _bullet(slide, Inches(7.05), py, Inches(5.4), pt, pi, size=13.5 if has_img else 14)
        py += Inches(0.74 if has_img else 0.82)
    # 底部横幅插图
    if has_img:
        _add_bottom_image(slide, s['image'], Inches(0.5), Inches(5.6), Inches(12.33), Inches(1.3),
                          caption=s.get('caption'), accent=s.get('accent', 'gold'), sub=s.get('sub'))
    return slide

def build_summary(prs, s):
    """三句话 + 医生行动清单（2×2 网格），可选右侧配图。"""
    slide = _new_slide(prs)
    _header(slide, s.get('title'), s.get('chapter'), s.get('page'))
    has_img = bool(s.get('image'))
    if has_img:
        _add_side_image(slide, s['image'], Inches(8.45), Inches(2.1), Inches(4.35), Inches(4.7),
                        caption=s.get('caption'), accent=s.get('accent', 'gold'))
        cw = Inches(7.55); x0 = Inches(0.5)
    else:
        cw = Inches(12.33); x0 = Inches(0.5)
    # 三句话
    y = Inches(2.3)
    for i, sent in enumerate(s.get('sentences', [])):
        _icon(slide, x0, y + Inches(0.06), i, 0.18, 'diamond')
        tb, tf = _tb(slide, x0 + Inches(0.35), y, cw - Inches(0.35), Inches(0.55))
        _para(tf, sent, size=15.5 if has_img else 17, bold=True, color=TEAL_D, first=True, space_after=0, line_spacing=1.1)
        y += Inches(0.66 if has_img else 0.7)
    # 分隔线 + 行动清单标题
    _rect(slide, x0, y + Inches(0.03), cw, Inches(0.06), fill=AMBER)
    tb, tf = _tb(slide, x0 + Inches(0.1), y + Inches(0.22), cw - Inches(0.2), Inches(0.4))
    _para(tf, s.get('actions_title', '医生行动清单'), size=16.5 if has_img else 18, bold=True, color=TEAL_D, first=True, space_after=0)
    # 2×2 行动卡片（统一浅蓝底，编号方块统一藏蓝）
    ay = y + Inches(0.7)
    actions = s.get('actions', [])
    col_w = (cw - Inches(0.1) - Inches(0.3)) / 2
    for pi, act in enumerate(actions):
        col = pi % 2
        row = pi // 2
        x = x0 + col * (col_w + Inches(0.3))
        yy = ay + row * Inches(0.82 if has_img else 0.88)
        _round_rect(slide, x, yy, col_w, Inches(0.68 if has_img else 0.74), fill=TEAL_L, radius=0.14)
        _rect(slide, x + Inches(0.15), yy + Inches(0.19), Inches(0.3), Inches(0.3), fill=TEAL)
        tbb = slide.shapes.add_textbox(x + Inches(0.15), yy + Inches(0.2), Inches(0.3), Inches(0.3)).text_frame
        tbb.word_wrap = False
        pp = tbb.paragraphs[0]; pp.alignment = PP_ALIGN.CENTER
        rr = pp.add_run(); _set_run(rr, str(pi+1), size=12 if has_img else 13, bold=True, color=WHITE)
        tb, tf = _tb(slide, x + Inches(0.58), yy + Inches(0.1), col_w - Inches(0.72), Inches(0.5))
        _para(tf, act, size=13.5 if has_img else 15, color=DARK, first=True, space_after=0, line_spacing=1.1)
    return slide

def build_thanks(prs, s):
    slide = _new_slide(prs)
    # 原版纯文字结尾
    _rect(slide, 0, 0, SLIDE_W, Inches(0.14), fill=TEAL)
    _rect(slide, 0, Inches(7.36), SLIDE_W, Inches(0.14), fill=TEAL)
    tb, tf = _tb(slide, Inches(0.5), Inches(2.7), Inches(12.3), Inches(1.0))
    _para(tf, s.get('text', '感谢聆听'), font=TITLE_FONT, size=44, bold=True, color=TEAL_D, align=PP_ALIGN.CENTER, first=True, space_after=0)
    if s.get('subtitle'):
        tb, tf = _tb(slide, Inches(0.5), Inches(3.8), Inches(12.3), Inches(0.5))
        _para(tf, s['subtitle'], size=20, color=DARK, align=PP_ALIGN.CENTER, first=True, space_after=0)
    # 琥珀色强调分隔线
    _rect(slide, Inches(6.17), Inches(4.45), Inches(1.0), Pt(2.5), fill=AMBER)
    y = Inches(4.6)
    for line in s.get('source', []):
        tb, tf = _tb(slide, Inches(0.5), y, Inches(12.3), Inches(0.4))
        _para(tf, line, size=12, color=DARK, align=PP_ALIGN.CENTER, first=True, space_after=0)
        y += Inches(0.4)
    return slide

BUILDERS = {
    'cover': build_cover, 'toc': build_toc, 'thanks': build_thanks,
    'two_col': build_two_col, 'data_band': build_data_band,
    'timeline': build_timeline, 'funnel': build_funnel, 'process': build_process,
    'matrix': build_matrix, 'table': build_table, 'cards': build_cards,
    'checklist': build_checklist, 'summary': build_summary,
}

def _default_template():
    """默认模板：config.json 的 template → 内置 assets/base_template.pptx。"""
    t = _resolve_path(_load_config().get('template', ''))
    if t and os.path.exists(t):
        return t
    t = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'base_template.pptx')
    return t if os.path.exists(t) else ''

TPL = _default_template()


def generate(spec, out_path, template=TPL):
    # 以模板为基底，保留其浅蓝母版背景
    prs = Presentation(template)
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    # 删除模板原始页（保留母版/主题/背景）
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.get(qn('r:id'))
        if rId:
            try:
                prs.part.drop_rel(rId)
            except Exception:
                pass
        sldIdLst.remove(sldId)
    meta = spec.get('meta', {})
    global FOOTER
    FOOTER = meta.get('footer', '') or ''
    slides = spec.get('slides', [])
    for s in slides:
        stype = s.get('type', 'two_col')
        if stype == 'cover':
            build_cover(prs, meta)
        else:
            BUILDERS.get(stype, build_two_col)(prs, s)
        notes = s.get('notes') if stype != 'cover' else meta.get('notes')
        if notes:
            try:
                prs.slides[-1].notes_slide.notes_text_frame.text = notes
            except Exception:
                pass
    prs.save(out_path)
    return out_path

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("用法: python3 generate_consensus.py spec.json out.pptx")
        sys.exit(1)
    spec = json.load(open(sys.argv[1], encoding='utf-8'))
    print("已生成:", generate(spec, sys.argv[2]))
