# -*- coding: utf-8 -*-
"""模板10：课程体系全景图 (Course System Panorama)
来源：国一实施报告重构框架 Slides 1-4 + 国一 建筑工程 Slide 1-2 + 国一 视觉传感小车 Slide 2
适用场景：课程体系全景、岗课赛证融通架构、人才培养路径总览
结构：顶部岗位/证书/大赛输入→ 中间课程内容重构→ 底部项目模块展开
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

PAGE_TITLE = '▎图X：课程体系全景图'
PAGE_SUBTITLE = '岗课赛证融通 · 项目驱动 · 能力递进  |  课程名称（X学时）'

# 输入源 (顶部驱动)
INPUT_SOURCES = [
    {'name': '岗位能力标准', 'items': ['国家标准/行业规范', '人才培养方案', '企业岗位需求'], 'color': B1},
    {'name': '技能大赛赛项', 'items': ['XX技能大赛', '竞赛规程分析', '赛项任务拆解'], 'color': B2},
    {'name': '职业技能证书', 'items': ['1+X证书(级别)', '行业执照要求', '证书考核大纲'], 'color': B3},
    {'name': '国规/省规教材', 'items': ['国家规划教材', '校企合作教材', '数字化资源'], 'color': B4},
]

# 课程内容重构 (中间层)
COURSE_RECONSTRUCTION = {
    'name': '课程内容重构',
    'desc': '按照技术逻辑选取教学载体\n简单→复杂 · 单一→综合',
    'color': DARK_BLUE,
}

# 教学项目模块 (底部展开)
PROJECTS = [
    {
        'name': '项目一',
        'subtitle': '基础认知模块',
        'tasks': ['任务一：XXX基础认知', '任务二：XXX基本操作', '任务三：XXX技能训练'],
        'hours': 'X学时',
        'color': B1,
    },
    {
        'name': '项目二',
        'subtitle': '核心技能模块',
        'tasks': ['任务一：XXX方案设计', '任务二：XXX核心实操', '任务三：XXX综合训练'],
        'hours': 'X学时',
        'color': B2,
    },
    {
        'name': '项目三',
        'subtitle': '综合应用模块',
        'tasks': ['任务一：XXX项目实战', '任务二：XXX综合调测', '任务三：XXX成果验收'],
        'hours': 'X学时',
        'color': B3,
    },
    {
        'name': '项目四',
        'subtitle': '创新拓展模块',
        'tasks': ['任务一：XXX创新设计', '任务二：XXX企业实战', '任务三：XXX成果转化'],
        'hours': 'X学时',
        'color': B4,
    },
]

# 底部培养目标
GOAL_SLOGAN = '"核心标语"'
GOAL_SUB = '复合型XXX技术技能人才'

# ============================================================
# ==== 生成逻辑 ==============================================
# ============================================================

def create_course_panorama():
    prs = Presentation()
    prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = BG_OUTER
    rrect(slide, 0.35, 0.08, 12.65, 7.35, BG_MAIN, border=LINE, radius=0.03)

    # 标题
    tb(slide, 0.65, 0.18, 8.0, 0.32, PAGE_TITLE, size=18, bold=True, color=DARK_BLUE)
    tb(slide, 0.65, 0.52, 10.0, 0.18, PAGE_SUBTITLE, size=8, color=MED)
    rect(slide, 0.65, 0.76, 12.0, 0.01, LINE)

    # ---- 顶部：四大输入源 ----
    source_top = 0.88
    source_w = 2.85; source_gap = 0.18
    source_start_x = 0.55

    for i, src in enumerate(INPUT_SOURCES):
        sx = source_start_x + i*(source_w + source_gap)
        sc = src['color']

        # 源卡片
        rrect(slide, sx, source_top, source_w, 1.15, WHITE, border=LINE)
        # 顶部色条
        rect(slide, sx, source_top, source_w, 0.06, sc)
        # 标题
        rrect(slide, sx+0.06, source_top+0.15, source_w-0.12, 0.32, sc, radius=0.03)
        tb(slide, sx+0.06, source_top+0.18, source_w-0.12, 0.26, src['name'],
           size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # 条目
        for j, item in enumerate(src['items']):
            iy = source_top + 0.58 + j*0.18
            rrect(slide, sx+0.12, iy+0.03, 0.08, 0.08, sc, radius=0.5)
            tb(slide, sx+0.28, iy, source_w-0.42, 0.15, item, size=7, color=MED)

        # 下指箭头
        arrow_d(slide, sx+source_w/2-0.06, source_top+1.2, 0.12, 0.1, sc)

    # ---- 中间：课程内容重构 ----
    recon_top = 2.25
    recon_w = 11.75
    recon_h = 0.55

    rrect(slide, source_start_x, recon_top, recon_w, recon_h, DARK_BLUE, radius=0.04)
    tb(slide, source_start_x+0.15, recon_top+0.05, 3.0, 0.22, COURSE_RECONSTRUCTION['name'],
       size=12, bold=True, color=WHITE)
    mtb(slide, source_start_x+3.5, recon_top+0.06, recon_w-3.8, recon_h-0.12,
        [(COURSE_RECONSTRUCTION['desc'], 8, False, RGBColor(0xB0,0xCC,0xE5), PP_ALIGN.LEFT)])

    # 多个箭头指向项目
    arrow_d(slide, source_start_x+recon_w/2-0.06, recon_top+recon_h, 0.12, 0.1, MED_BLUE)

    # ---- 底部：四项目展开 ----
    proj_top = 3.1
    proj_w = 2.85; proj_gap = 0.18
    proj_h = 2.2

    for i, proj in enumerate(PROJECTS):
        px = source_start_x + i*(proj_w + proj_gap)
        pc = proj['color']

        # 项目卡片
        rrect(slide, px, proj_top, proj_w, proj_h, WHITE, border=LINE)
        # 项目标题
        rrect(slide, px, proj_top, proj_w, 0.4, pc, radius=0.03)
        tb(slide, px+0.08, proj_top+0.05, proj_w-0.16, 0.2, proj['name'],
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, px+0.08, proj_top+0.24, proj_w-0.16, 0.14, proj['subtitle'],
           size=7, color=RGBColor(0xB0,0xCC,0xE5), align=PP_ALIGN.CENTER)

        # 任务列表
        task_top = proj_top + 0.55
        for j, task in enumerate(proj['tasks']):
            ty = task_top + j*0.45
            # 任务背景
            rrect(slide, px+0.08, ty, proj_w-0.16, 0.37, CARD_LT, border=None, radius=0.03)
            # 序号
            rrect(slide, px+0.14, ty+0.07, 0.22, 0.22, pc, radius=0.5)
            tb(slide, px+0.14, ty+0.08, 0.22, 0.2, str(j+1),
               size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            tb(slide, px+0.44, ty+0.06, proj_w-0.6, 0.24, task, size=7.5, color=DARK)

        # 学时标签
        rrect(slide, px+proj_w-0.75, proj_top+proj_h-0.36, 0.6, 0.24, CARD_LT, radius=0.03)
        tb(slide, px+proj_w-0.75, proj_top+proj_h-0.34, 0.6, 0.2, proj['hours'],
           size=7.5, bold=True, color=MED_BLUE, align=PP_ALIGN.CENTER)

    # ---- 右侧纵向：能力递进标注 ----
    right_bar_x = 12.6; right_bar_w = 0.18
    rrect(slide, right_bar_x, proj_top, right_bar_w, proj_h, B1, radius=0.02)
    mtb(slide, right_bar_x+0.03, proj_top+0.2, 0.12, proj_h-0.4, [
        ('递', 7, True, WHITE, PP_ALIGN.CENTER),
        ('', 4, False, WHITE, PP_ALIGN.CENTER),
        ('进', 7, True, WHITE, PP_ALIGN.CENTER),
    ])

    # ---- 底部培养目标 ----
    goal_y = 5.55
    rrect(slide, 3.0, goal_y, 7.3, 0.7, WHITE, border=LINE)
    rect(slide, 3.0, goal_y, 7.3, 0.04, ORANGE)
    mtb(slide, 3.15, goal_y+0.1, 7.0, 0.5, [
        (GOAL_SLOGAN, 13, True, ORANGE, PP_ALIGN.CENTER),
        (GOAL_SUB, 10, True, DARK_BLUE, PP_ALIGN.CENTER),
    ])

    # 底部岗课赛证横条
    bar_y = goal_y + 0.85
    bar_w = 2.85; bar_gap = 0.18
    bar_labels = ['岗：岗位导向', '课：课程支撑', '赛：大赛检验', '证：证书认证']
    bar_colors = [B1, B2, B3, B4]

    for i, (bl, bc) in enumerate(zip(bar_labels, bar_colors)):
        bx = source_start_x + i*(bar_w + bar_gap)
        rrect(slide, bx, bar_y, bar_w, 0.35, bc, radius=0.04)
        tb(slide, bx, bar_y+0.04, bar_w, 0.27, bl,
           size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    return prs

if __name__ == '__main__':
    out_dir = os.path.dirname(os.path.abspath(__file__))
    prs = create_course_panorama()
    prs.save(os.path.join(out_dir, '模板10-课程体系全景图.pptx'))
    print('已生成: 模板10-课程体系全景图.pptx')
