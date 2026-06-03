# -*- coding: utf-8 -*-
"""模板16：教学流程闭环 (Teaching Process Closed Loop)
来源：Z62 10步流水线 + 200页模板 Slide 136/142 + 重构框架 Slide 86
适用场景：完整教学实施闭环、双线并进流程、课内外联动路径
结构：上弧线（线上）→ 下弧线（线下）→ 左右闭环连接
视觉增强：弧形路径渐变 + 阶段节点阴影 + 反馈回路标注
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
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
GREEN     = RGBColor(0x60, 0xBC, 0x90)

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

def _get_spPr(shape):
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    spPr = shape._element.find(f'{ns}spPr')
    if spPr is None:
        spPr = etree.SubElement(shape._element, f'{ns}spPr')
    return spPr

def add_shadow(shape, blur_rad=35000, dist=18000, alpha=18000):
    spPr = _get_spPr(shape)
    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    effectLst = spPr.find(f'{{{ns_a}}}effectLst')
    if effectLst is None:
        effectLst = etree.SubElement(spPr, f'{{{ns_a}}}effectLst')
    outerShdw = etree.SubElement(effectLst, f'{{{ns_a}}}outerShdw', {
        'blurRad': str(blur_rad), 'dist': str(dist),
        'dir': '5400000', 'algn': 'ctr',
    })
    srgbClr = etree.SubElement(outerShdw, f'{{{ns_a}}}srgbClr', {'val': '000000'})
    etree.SubElement(srgbClr, f'{{{ns_a}}}alpha', {'val': str(alpha)})

def set_gradient_fill(shape, c1_hex, c2_hex, angle=90):
    spPr = _get_spPr(shape)
    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    for sf in spPr.findall(f'{{{ns_a}}}solidFill'):
        spPr.remove(sf)
    gradFill = etree.SubElement(spPr, f'{{{ns_a}}}gradFill', {'rotWithShape': '1'})
    gsLst = etree.SubElement(gradFill, f'{{{ns_a}}}gsLst')
    for pos, clr in [('0', c1_hex), ('100000', c2_hex)]:
        gs = etree.SubElement(gsLst, f'{{{ns_a}}}gs', {'pos': pos})
        etree.SubElement(gs, f'{{{ns_a}}}srgbClr', {'val': clr})
    lin = etree.SubElement(gradFill, f'{{{ns_a}}}lin', {'ang': str(angle * 60000), 'scaled': '1'})

def rrect(s, l, t, w, h, fill, border=None, radius=0.04, shadow=False, gradient=None):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l),Inches(t),Inches(w),Inches(h))
    if gradient:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
        set_gradient_fill(sh, gradient[0], gradient[1])
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if border: sh.line.color.rgb = border; sh.line.width = Pt(0.5)
    else: sh.line.fill.background()
    sh.adjustments[0] = radius
    if shadow: add_shadow(sh)
    return sh

def rect(s, l, t, w, h, fill, border=None, shadow=False):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if border: sh.line.color.rgb = border; sh.line.width = Pt(0.5)
    else: sh.line.fill.background()
    if shadow: add_shadow(sh)
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

PAGE_TITLE = '▎图X："双线并进"教学实施闭环'
PAGE_SUBTITLE = '线上探究 → 线下实践 → 双线融合 → 闭环迭代  |  课前·课中·课后三位一体'

# 上半环：线上路径（左→右）
ONLINE_STEPS = [
    {'label': '01', 'name': '资源发布', 'desc': '平台推送\n预习任务', 'color': B3},
    {'label': '02', 'name': '自主预学', 'desc': '微课学习\n在线测试', 'color': B4},
    {'label': '03', 'name': '学情诊断', 'desc': '数据分析\n精准定位', 'color': B2},
]

# 下半环：线下路径（右→左 反馈）
OFFLINE_STEPS = [
    {'label': '04', 'name': '任务驱动', 'desc': '案例导入\n分组探究', 'color': B1},
    {'label': '05', 'name': '实操演练', 'desc': '工坊实操\n仿真训练', 'color': B2},
    {'label': '06', 'name': '成果检验', 'desc': '成果汇报\n多维评价', 'color': B3},
    {'label': '07', 'name': '拓展提升', 'desc': '创新实践\n能力迁移', 'color': B1},
]

# 闭环反馈
FEEDBACK_LABEL = '数据反馈驱动教学改进'

BOTTOM_SLOGAN = '"线上线下融合 · 教学评一体化"'
BOTTOM_SUB = '全过程闭环管理 · 持续迭代优化'

# ============================================================
# ==== 生成逻辑 ==============================================
# ============================================================

def create_loop():
    prs = Presentation()
    prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = BG_OUTER
    rrect(slide, 0.35, 0.08, 12.65, 7.35, BG_MAIN, border=LINE, radius=0.03)

    # 渐变标题栏
    rrect(slide, 0.65, 0.18, 12.0, 0.48, DARK_BLUE, radius=0.04,
          gradient=('1E5793', '2E78B5'))
    tb(slide, 0.85, 0.22, 8.0, 0.28, PAGE_TITLE, size=16, bold=True, color=WHITE)
    tb(slide, 0.85, 0.48, 10.0, 0.16, PAGE_SUBTITLE, size=7.5, color=RGBColor(0xB0,0xCC,0xE5))

    # ---- 上半环：线上（左→右） ----
    online_y = 0.95
    online_h = 1.55
    step_w = 2.4; step_gap = 0.35

    # 线上标签
    rrect(slide, 0.55, online_y-0.28, 1.2, 0.28, B4, radius=0.03)
    tb(slide, 0.55, online_y-0.25, 1.2, 0.22, '☁ 线上', size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    online_start = 0.55 + 1.5
    for i, step in enumerate(ONLINE_STEPS):
        ox = online_start + i*(step_w + step_gap)
        sc = step['color']

        # 步骤卡片
        rrect(slide, ox, online_y, step_w, online_h, WHITE, border=LINE, shadow=True)
        rect(slide, ox, online_y, step_w, 0.05, sc, shadow=False)

        # 编号
        rrect(slide, ox+0.1, online_y+0.15, 0.4, 0.4, sc, radius=0.5)
        tb(slide, ox+0.1, online_y+0.2, 0.4, 0.3, step['label'],
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # 名称
        tb(slide, ox+0.6, online_y+0.18, step_w-0.8, 0.22, step['name'],
           size=10, bold=True, color=DARK)
        # 描述
        mtb(slide, ox+0.6, online_y+0.5, step_w-0.8, online_h-0.7,
            [(step['desc'], 7.5, False, MED, PP_ALIGN.LEFT)])

        # 步骤间箭头
        if i < len(ONLINE_STEPS)-1:
            arrow_r(slide, ox+step_w+0.02, online_y+online_h/2-0.05, 0.28, 0.1, sc)

    # ---- 右连接：线上→线下 ----
    mid_rx = online_start + len(ONLINE_STEPS)*(step_w+step_gap) - step_gap + 0.05
    arrow_d(slide, mid_rx+step_w/2-0.06, online_y+online_h+0.02, 0.12, 0.2, B2)

    # ---- 下半环：线下（右→左 反馈） ----
    offline_y = online_y + online_h + 0.45
    offline_h = 1.55

    # 线下标签
    rrect(slide, 0.55, offline_y-0.28, 1.2, 0.28, B1, radius=0.03)
    tb(slide, 0.55, offline_y-0.25, 1.2, 0.22, '🏫 线下', size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    offline_start = 0.55 + 1.5
    for i, step in enumerate(OFFLINE_STEPS):
        ox = offline_start + i*(step_w + step_gap)
        sc = step['color']

        # 步骤卡片
        rrect(slide, ox, offline_y, step_w, offline_h, WHITE, border=LINE, shadow=True)
        rect(slide, ox, offline_y, step_w, 0.05, sc)

        # 编号
        rrect(slide, ox+0.1, offline_y+0.15, 0.4, 0.4, sc, radius=0.5)
        tb(slide, ox+0.1, offline_y+0.2, 0.4, 0.3, step['label'],
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # 名称
        tb(slide, ox+0.6, offline_y+0.18, step_w-0.8, 0.22, step['name'],
           size=10, bold=True, color=DARK)
        # 描述
        mtb(slide, ox+0.6, offline_y+0.5, step_w-0.8, offline_h-0.7,
            [(step['desc'], 7.5, False, MED, PP_ALIGN.LEFT)])

        # 步骤间箭头 (左向)
        if i < len(OFFLINE_STEPS)-1:
            arrow_r(slide, ox+step_w+0.02, offline_y+offline_h/2-0.05, 0.28, 0.1, sc)

    # ---- 左连接：线下→线上（反馈闭环） ----
    # 返回箭头
    arrow_d(slide, 0.55+0.5, online_y+online_h+0.05, 0.12, 0.35, ORANGE)
    rrect(slide, 0.55+0.05, online_y+online_h+0.15, 1.1, 0.22, ORANGE, radius=0.03)
    tb(slide, 0.55+0.08, online_y+online_h+0.17, 1.04, 0.18, FEEDBACK_LABEL,
       size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # ---- 底部双线说明 ----
    legend_y = offline_y + offline_h + 0.25
    rrect(slide, 0.55, legend_y, 12.15, 0.75, WHITE, border=LINE, shadow=True)
    rect(slide, 0.55, legend_y, 12.15, 0.04, ORANGE)

    # 四格说明
    info_items = [
        ('线上探究', '课前预习+课后拓展', B4),
        ('线下实践', '课中教学+实操训练', B1),
        ('数据驱动', '过程数据全程记录', B2),
        ('闭环迭代', '诊断→改进→再实施', ORANGE),
    ]
    for i, (title, desc, ic) in enumerate(info_items):
        ix = 0.75 + i*3.0
        rrect(slide, ix, legend_y+0.15, 0.08, 0.42, ic)
        tb(slide, ix+0.18, legend_y+0.1, 2.5, 0.22, title, size=9, bold=True, color=ic)
        tb(slide, ix+0.18, legend_y+0.38, 2.5, 0.18, desc, size=7.5, color=MED)

    # ---- 底部标语 ----
    slogan_y = legend_y + 0.9
    rrect(slide, 3.0, slogan_y, 7.3, 0.55, WHITE, border=LINE, shadow=True)
    rect(slide, 3.0, slogan_y, 7.3, 0.04, ORANGE)
    mtb(slide, 3.15, slogan_y+0.06, 7.0, 0.43, [
        (BOTTOM_SLOGAN, 12, True, ORANGE, PP_ALIGN.CENTER),
        (BOTTOM_SUB, 8, False, DARK_BLUE, PP_ALIGN.CENTER),
    ])

    return prs

if __name__ == '__main__':
    out_dir = os.path.dirname(os.path.abspath(__file__))
    prs = create_loop()
    prs.save(os.path.join(out_dir, '模板16-教学流程闭环.pptx'))
    print('已生成: 模板16-教学流程闭环.pptx')
