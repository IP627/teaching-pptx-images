# -*- coding: utf-8 -*-
"""模板05：岗课赛证融通 (Job-Course-Competition-Certificate Integration)
来源：国一 建筑工程 + 国一 重构框架 + 国二 智能摆渡车
适用场景：岗课赛证四维融通图、人才培养路径、课程体系对证
结构：左侧岗位/课/赛/证 四张纵向卡片 → 右侧融通链接 → 底部培养目标
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

# ============================================================
# ==== 可修改区 ==============================================
# ============================================================

PAGE_TITLE = '▎图X："岗课赛证"融通培养体系'
PAGE_SUBTITLE = '岗位需求驱动 · 课程体系支撑 · 技能竞赛检验 · 证书标准对接'

# 岗课赛证四个卡片
GKSZ_ITEMS = [
    {
        'label': '岗',
        'title': '岗位名称',
        'desc': '对标XX企业等\n岗位能力标准要求',
        'color': B1,
    },
    {
        'label': '课',
        'title': '课程名称 (X学时)',
        'desc': 'N模块·四阶递进\n项目驱动·AI赋能',
        'color': B2,
    },
    {
        'label': '赛',
        'title': '技能大赛赛项',
        'desc': 'XX赛项\n核心任务拆解融入',
        'color': B3,
    },
    {
        'label': '证',
        'title': '职业技能等级证书',
        'desc': 'XX职业技能等级(级别)\nXX执照/证书',
        'color': B4,
    },
]

# 融通描述 (中间列)
INTEGRATION_TEXT = [
    '课程内容对标岗位能力标准',
    '赛项任务拆解为教学模块',
    '证书考核点融入教学过程',
]

# 底部培养目标
GOAL_TITLE = '培养目标'
GOAL_SLOGAN = '"核心标语"'
GOAL_SUB = '复合型XXX技术人才'

# ============================================================
# ==== 生成逻辑 ==============================================
# ============================================================

def create_gksz():
    prs = Presentation()
    prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = BG_OUTER
    rrect(slide, 0.35, 0.08, 12.65, 7.35, BG_MAIN, border=LINE, radius=0.03)

    # 标题
    tb(slide, 0.65, 0.18, 8.0, 0.32, PAGE_TITLE, size=18, bold=True, color=DARK_BLUE)
    tb(slide, 0.65, 0.52, 10.0, 0.18, PAGE_SUBTITLE, size=8, color=MED)
    rect(slide, 0.65, 0.76, 12.0, 0.01, LINE)

    # ---- 左列：岗课赛证四张卡片 ----
    lx, lw = 0.55, 3.65
    card_h = 1.3; card_gap = 0.12
    card_top = 0.95

    for i, item in enumerate(GKSZ_ITEMS):
        cy = card_top + i*(card_h + card_gap)
        ac = item['color']

        # 卡片背景
        rrect(slide, lx, cy, lw, card_h, WHITE, border=LINE)
        rect(slide, lx, cy, lw, 0.04, ac)

        # 大字标签
        rrect(slide, lx+0.1, cy+0.12, 0.55, 0.55, ac, radius=0.1)
        tb(slide, lx+0.1, cy+0.22, 0.55, 0.4, item['label'],
           size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # 标题与描述
        tb(slide, lx+0.8, cy+0.12, lw-1.0, 0.24, item['title'],
           size=10, bold=True, color=DARK)
        mtb(slide, lx+0.8, cy+0.42, lw-1.0, 0.8,
            [(item['desc'], 7.5, False, MED, PP_ALIGN.LEFT)])

    # ---- 中列：融通箭头和描述 ----
    mx, mw = 4.4, 2.4
    arrow_y_start = card_top + card_h/2
    arrow_y_end = card_top + 3*(card_h+card_gap) + card_h/2

    for i, item in enumerate(GKSZ_ITEMS):
        ay = card_top + i*(card_h+card_gap) + card_h/2
        arrow_r(slide, lx+lw+0.05, ay-0.06, 0.3, 0.12, item['color'])

    # 融通文字说明
    rrect(slide, mx, card_top+0.5, mw, card_h*3.8, CARD_LT, border=None, radius=0.03)
    mtb(slide, mx+0.15, card_top+0.7, mw-0.3, card_h*3.5, [
        ('融通机制', 10, True, DARK_BLUE, PP_ALIGN.CENTER),
        ('', 6, False, DARK, PP_ALIGN.CENTER),
    ] + [(item, 7.5, False, MED, PP_ALIGN.CENTER) for item in INTEGRATION_TEXT])

    # ---- 右列：岗课赛证横排连标 ----
    rx, rw = 7.0, 1.8
    for i, item in enumerate(GKSZ_ITEMS):
        ry = card_top + i*(card_h+card_gap) + card_h/2
        rrect(slide, rx, ry-0.22, rw, 0.44, item['color'], radius=0.05)
        tb(slide, rx, ry-0.15, rw, 0.3, f'{item["label"]}：{item["title"]}',
           size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # ---- 右侧：培养目标 ----
    goal_l, goal_w = 9.1, 3.7
    goal_y = card_top
    goal_h = card_h + card_gap
    rrect(slide, goal_l, goal_y, goal_w, goal_h*3.5, WHITE, border=LINE)
    rect(slide, goal_l, goal_y, goal_w, 0.05, ORANGE)

    mtb(slide, goal_l+0.15, goal_y+0.25, goal_w-0.3, goal_h*3.0-0.5, [
        (GOAL_TITLE, 9, False, MED_BLUE, PP_ALIGN.CENTER),
        ('', 8, False, DARK, PP_ALIGN.CENTER),
        (GOAL_SLOGAN, 16, True, ORANGE, PP_ALIGN.CENTER),
        ('', 4, False, DARK, PP_ALIGN.CENTER),
        (GOAL_SUB, 10, True, DARK_BLUE, PP_ALIGN.CENTER),
    ])

    # 底部四字横条
    bar_y = goal_y + goal_h*3.5 + 0.15
    bar_items = ['岗', '课', '赛', '证']
    bar_colors = [B1, B2, B3, B4]
    for i, (bl, bc) in enumerate(zip(bar_items, bar_colors)):
        bx = lx + i*1.15
        rrect(slide, bx, bar_y, 1.0, 0.35, bc, radius=0.04)
        tb(slide, bx, bar_y+0.03, 1.0, 0.28, bl,
           size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        if i < 3:
            arrow_r(slide, bx+1.02, bar_y+0.1, 0.1, 0.15, MED_BLUE)

    return prs

if __name__ == '__main__':
    out_dir = os.path.dirname(os.path.abspath(__file__))
    prs = create_gksz()
    prs.save(os.path.join(out_dir, '模板05-岗课赛证融通.pptx'))
    print('已生成: 模板05-岗课赛证融通.pptx')
