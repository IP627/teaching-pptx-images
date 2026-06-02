# -*- coding: utf-8 -*-
"""模板04：中心放射联动 (Center-Radial / Hub-Spoke Layout)
来源：国二 思政-明大德 + 国一 视觉传感小车 Slide 1 + 寻路立心 Slide 6
适用场景：四驱联动策略、教学模式核心图、多维协同机制
结构：中心核心节点 → 四周N个驱动卡片 → 连接线/箭头
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os

CN_FONT = 'PingFang SC'

BG_OUTER  = RGBColor(0xF8, 0xFB, 0xFE)
BG_MAIN   = RGBColor(0xF5, 0xFA, 0xFF)
CARD_LT   = RGBColor(0xD5, 0xF4, 0xFF)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK      = RGBColor(0x1A, 0x1A, 0x2E)
MED       = RGBColor(0x55, 0x55, 0x55)
DARK_BLUE = RGBColor(0x1A, 0x3C, 0x5E)
MED_BLUE  = RGBColor(0x2C, 0x5F, 0x8A)
LINE      = RGBColor(0xD0, 0xDE, 0xEB)
LINE2     = RGBColor(0xE5, 0xEC, 0xF3)
ORANGE    = RGBColor(0xF4, 0x79, 0x20)

B1 = RGBColor(0x1E, 0x57, 0x93)
B2 = RGBColor(0x2E, 0x78, 0xB5)
B3 = RGBColor(0x4B, 0x9B, 0xD4)
B4 = RGBColor(0x7D, 0xC5, 0xED)

def cn(run, name=CN_FONT, size=12, bold=False, color=None):
    run.font.name = name; run.font.size = Pt(size); run.font.bold = bold
    if color: run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {}); rPr.append(ea)
    ea.set('typeface', name)

def rrect(s, l, t, w, h, fill, border=None, radius=0.04):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if border: sh.line.color.rgb = border; sh.line.width = Pt(0.5)
    else: sh.line.fill.background()
    sh.adjustments[0] = radius
    return sh

def rect(s, l, t, w, h, fill, border=None):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if border: sh.line.color.rgb = border; sh.line.width = Pt(0.5)
    else: sh.line.fill.background()
    return sh

def tb(s, l, t, w, h, text, size=10, bold=False, color=DARK, align=PP_ALIGN.LEFT):
    bx = s.shapes.add_textbox(Inches(l),Inches(t),Inches(w),Inches(h))
    bx.text_frame.word_wrap = True; bx.text_frame.auto_size = None
    p = bx.text_frame.paragraphs[0]
    p.alignment = align; p.space_after = Pt(0); p.space_before = Pt(0)
    r = p.add_run(); r.text = text
    cn(r, size=size, bold=bold, color=color)
    return bx

def mtb(s, l, t, w, h, lines):
    bx = s.shapes.add_textbox(Inches(l),Inches(t),Inches(w),Inches(h))
    bx.text_frame.word_wrap = True; bx.text_frame.auto_size = None
    for i, (text, sz, bd, cl, al) in enumerate(lines):
        p = bx.text_frame.paragraphs[0] if i==0 else bx.text_frame.add_paragraph()
        p.alignment = al; p.space_after = Pt(1); p.space_before = Pt(0)
        r = p.add_run(); r.text = text
        cn(r, size=sz, bold=bd, color=cl)
    return bx

def arrow_r(s, x, y, w, h, color):
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x),Inches(y),Inches(w),Inches(h))
    a.fill.solid(); a.fill.fore_color.rgb = color; a.line.fill.background()

def arrow_d(s, x, y, w, h, color):
    a = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(x),Inches(y),Inches(w),Inches(h))
    a.fill.solid(); a.fill.fore_color.rgb = color; a.line.fill.background()

# ============================================================
# ==== 可修改区 ==============================================
# ============================================================

PAGE_TITLE = '▎图X："四驱联动"教学策略框架'
PAGE_SUBTITLE = '以学生为中心、产出为导向  |  课程名称'

# 中心节点
CENTER_TITLE = '教学实施\n核心'
CENTER_SUB = '六模块递进'
CENTER_COLOR = DARK_BLUE

# 四周驱动卡片 (上、右、下、左 四方向)
DRIVE_CARDS = [
    {
        'name': '驱动一',
        'label': '01',
        'sub': '驱动副标题',
        'details': ['详细点 A', '详细点 B', '详细点 C'],
        'color': B1,
        'position': 'top',  # top/right/bottom/left
    },
    {
        'name': '驱动二',
        'label': '02',
        'sub': '驱动副标题',
        'details': ['详细点 A', '详细点 B', '详细点 C'],
        'color': B2,
        'position': 'right',
    },
    {
        'name': '驱动三',
        'label': '03',
        'sub': '驱动副标题',
        'details': ['详细点 A', '详细点 B', '详细点 C'],
        'color': B3,
        'position': 'bottom',
    },
    {
        'name': '驱动四',
        'label': '04',
        'sub': '驱动副标题',
        'details': ['详细点 A', '详细点 B', '详细点 C'],
        'color': B4,
        'position': 'left',
    },
]

# 底部产出标语
OUTPUT_SLOGAN = '"核心标语文字"'
OUTPUT_SUB = '复合型XXX技术人才'

# ============================================================
# ==== 生成逻辑 ==============================================
# ============================================================

# 位置计算 (以中心为原点)
# 中心区域: x=4.5~8.8, y=2.5~5.0
CX, CY = 6.65, 3.75
CW, CH = 2.0, 2.0  # 中心卡片尺寸

# 四周卡片位置 (相对于中心)
POSITIONS = {
    'top':    (CX + CW/2 - 1.3, CY - 1.8, 2.6, 1.3),   # 上方
    'right':  (CX + CW + 0.3, CY + CH/2 - 0.85, 2.8, 1.5),  # 右侧
    'bottom': (CX + CW/2 - 1.3, CY + CH + 0.5, 2.6, 1.3),   # 下方
    'left':   (CX - 2.8 - 0.3, CY + CH/2 - 0.85, 2.8, 1.5), # 左侧
}

def create_center_radial():
    prs = Presentation()
    prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = BG_OUTER
    rrect(slide, 0.35, 0.08, 12.65, 7.35, BG_MAIN, border=LINE, radius=0.03)

    # 标题
    tb(slide, 0.65, 0.18, 8.0, 0.32, PAGE_TITLE, size=18, bold=True, color=DARK_BLUE)
    tb(slide, 0.65, 0.52, 10.0, 0.18, PAGE_SUBTITLE, size=8, color=MED)
    rect(slide, 0.65, 0.76, 12.0, 0.01, LINE)

    # ---- 中心节点 ----
    # 大背景圆 (装饰)
    rrect(slide, CX-0.3, CY-0.3, CW+0.6, CH+0.6, BG_MAIN, border=LINE, radius=0.5)
    # 中心卡片
    rrect(slide, CX, CY, CW, CH, CENTER_COLOR, radius=0.08)
    mtb(slide, CX+0.1, CY+0.35, CW-0.2, CH-0.7, [
        (CENTER_TITLE, 14, True, WHITE, PP_ALIGN.CENTER),
        ('', 4, False, WHITE, PP_ALIGN.CENTER),
        (CENTER_SUB, 9, False, RGBColor(0xB0,0xCC,0xE5), PP_ALIGN.CENTER),
    ])

    # ---- 四周驱动卡片 ----
    for card in DRIVE_CARDS:
        x, y, w, h = POSITIONS[card['position']]
        ac = card['color']

        # 卡片背景
        rrect(slide, x, y, w, h, WHITE, border=LINE)
        # 顶部色条
        rect(slide, x, y, w, 0.04, ac)
        # 编号
        rrect(slide, x+0.1, y+0.15, 0.36, 0.36, ac, radius=0.5)
        tb(slide, x+0.1, y+0.19, 0.36, 0.28, card['label'],
           size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # 标题
        tb(slide, x+0.55, y+0.1, w-0.7, 0.22, card['name'],
           size=10, bold=True, color=DARK)
        tb(slide, x+0.55, y+0.32, w-0.7, 0.14, card['sub'],
           size=7, color=MED)
        # 详情
        for j, det in enumerate(card['details']):
            dy = y + 0.55 + j*0.22
            rrect(slide, x+0.2, dy+0.04, 0.07, 0.07, ac, radius=0.5)
            tb(slide, x+0.35, dy, w-0.55, 0.16, det, size=7, color=MED)

    # ---- 连接箭头 (中心 ←→ 四周) ----
    # 上→中心
    arrow_d(slide, CX+CW/2-0.06, CY-0.45, 0.12, 0.3, B1)
    # 中心→右
    arrow_r(slide, CX+CW+0.05, CY+CH/2-0.06, 0.25, 0.12, B2)
    # 中心→下
    arrow_d(slide, CX+CW/2-0.06, CY+CH+0.05, 0.12, 0.3, B3)
    # 左→中心
    arrow_r(slide, CX-0.4, CY+CH/2-0.06, 0.25, 0.12, B4)

    # ---- 底部产出 ----
    out_y = 6.5
    rrect(slide, 3.0, out_y, 7.3, 0.8, WHITE, border=LINE)
    rect(slide, 3.0, out_y, 7.3, 0.04, ORANGE)
    mtb(slide, 3.15, out_y+0.1, 7.0, 0.6, [
        (OUTPUT_SLOGAN, 14, True, ORANGE, PP_ALIGN.CENTER),
        (OUTPUT_SUB, 10, True, DARK_BLUE, PP_ALIGN.CENTER),
    ])

    return prs

if __name__ == '__main__':
    out_dir = os.path.dirname(os.path.abspath(__file__))
    prs = create_center_radial()
    prs.save(os.path.join(out_dir, '模板04-中心放射联动.pptx'))
    print('已生成: 模板04-中心放射联动.pptx')
