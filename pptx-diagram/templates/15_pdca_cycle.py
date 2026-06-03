# -*- coding: utf-8 -*-
"""模板15：PDCA循环改进环 (PDCA Continuous Improvement Cycle)
来源：200页模板 Slide 142/143 + Z68 BOPPPS + 重构框架 Slide 86
适用场景：教学诊断改进、质量管理循环、课程迭代优化、教改实施路径
结构：4象限环形布局 + 中心驱动 + 循环箭头 + 每象限展开细则
视觉增强：渐变环形色块 + 阴影节点 + 旋转箭头连接
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
import os, math

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
PURPLE    = RGBColor(0x8E, 0x6B, 0xC4)

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

PAGE_TITLE = '▎图X：PDCA教学质量持续改进循环'
PAGE_SUBTITLE = 'Plan → Do → Check → Act  |  数据驱动 · 闭环迭代 · 螺旋上升'

# PDCA 四象限
PDCA_QUADRANTS = [
    {
        'label': 'P',
        'name': '计划',
        'full': 'Plan',
        'color': B1,
        'details': [
            '学情分析诊断',
            '教学目标设定',
            '教学策略设计',
            '资源准备评估',
        ],
    },
    {
        'label': 'D',
        'name': '执行',
        'full': 'Do',
        'color': B2,
        'details': [
            '课堂教学实施',
            '项目任务驱动',
            '小组协作探究',
            '过程数据采集',
        ],
    },
    {
        'label': 'C',
        'name': '检查',
        'full': 'Check',
        'color': ORANGE,
        'details': [
            '多维评价分析',
            '学习效果检测',
            '达成度对比',
            '问题诊断归因',
        ],
    },
    {
        'label': 'A',
        'name': '改进',
        'full': 'Act',
        'color': GREEN,
        'details': [
            '教学策略优化',
            '资源迭代升级',
            '个性化补救',
            '标准化推广',
        ],
    },
]

# 中心驱动
CENTER_TITLE = '持续改进\n质量文化'
CENTER_SUB = '闭环迭代\n螺旋上升'

BOTTOM_SLOGAN = '"诊断 · 改进 · 提升"'
BOTTOM_SUB = '基于PDCA的教学质量持续改进机制'

# ============================================================
# ==== 生成逻辑 ==============================================
# ============================================================

def create_pdca():
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

    # ---- PDCA 四象限环 ----
    cx, cy = 6.65, 4.15  # 中心
    quad_w = 4.2; quad_h = 2.3

    # 四象限位置（左上/右上/右下/左下）
    positions = [
        (cx-quad_w-0.2, cy-quad_h-0.35, quad_w, quad_h),   # P - 左上
        (cx+0.2, cy-quad_h-0.35, quad_w, quad_h),            # D - 右上
        (cx+0.2, cy+0.35, quad_w, quad_h),                   # C - 右下
        (cx-quad_w-0.2, cy+0.35, quad_w, quad_h),            # A - 左下
    ]

    for i, quad in enumerate(PDCA_QUADRANTS):
        qx, qy, qw, qh = positions[i]
        qc = quad['color']

        # 象限背景
        rrect(slide, qx, qy, qw, qh, WHITE, border=LINE, shadow=True)
        # 象限标题
        rrect(slide, qx, qy, qw, 0.42, qc, radius=0.03,
              gradient=(_hh(qc), _hh(qc, 1.15)))
        # 大字标签
        rrect(slide, qx+0.1, qy+0.06, 0.32, 0.32, WHITE, radius=0.5)
        tb(slide, qx+0.1, qy+0.1, 0.32, 0.24, quad['label'],
           size=15, bold=True, color=qc, align=PP_ALIGN.CENTER)
        tb(slide, qx+0.52, qy+0.06, 1.5, 0.2, quad['name'],
           size=12, bold=True, color=WHITE)
        tb(slide, qx+0.52, qy+0.24, 1.5, 0.14, quad['full'],
           size=7, color=RGBColor(0xB0,0xCC,0xE5))

        # 象限内容
        for j, det in enumerate(quad['details']):
            dy = qy + 0.55 + j*0.4
            rrect(slide, qx+0.15, dy+0.04, 0.1, 0.1, qc, radius=0.5)
            tb(slide, qx+0.35, dy, qw-0.55, 0.2, det, size=8, color=DARK)

    # ---- 中心圆 ----
    center_r = 0.85
    rrect(slide, cx-center_r, cy-center_r, center_r*2, center_r*2, DARK_BLUE, radius=0.5, shadow=True,
          gradient=('1E5793', '2E78B5'))
    mtb(slide, cx-center_r+0.1, cy-0.35, center_r*2-0.2, center_r*2-0.2, [
        (CENTER_TITLE, 10, True, WHITE, PP_ALIGN.CENTER),
        ('', 3, False, WHITE, PP_ALIGN.CENTER),
        (CENTER_SUB, 7, False, RGBColor(0xB0,0xCC,0xE5), PP_ALIGN.CENTER),
    ])

    # ---- 循环箭头 ----
    arrow_colors = [B1, B2, ORANGE, GREEN]
    arrow_positions = [
        # 上: P → D (右箭头)
        (cx-1.2, cy-quad_h-0.45, 2.4, 0.1),
        # 右: D → C (下箭头)
        (cx+quad_w+0.1, cy-0.9, 0.1, 1.8),
        # 下: C → A (左箭头) - use right arrow flipped isn't easy, use smaller arrows
        (cx-1.2, cy+quad_h+0.35, 2.4, 0.1),
        # 左: A → P (上箭头)
        (cx-quad_w-0.3, cy-0.9, 0.1, 1.8),
    ]
    for i, (ax, ay, aw, ah) in enumerate(arrow_positions):
        ac = arrow_colors[i]
        if i in [0, 2]:  # horizontal
            arrow_r(slide, ax, ay+0.02, aw, ah-0.04, ac)
        else:  # vertical
            arrow_d(slide, ax+0.02, ay, aw-0.04, ah, ac)

    # ---- 底部标语 ----
    slogan_y = cy + quad_h + 0.55
    rrect(slide, 3.0, slogan_y, 7.3, 0.6, WHITE, border=LINE, shadow=True)
    rect(slide, 3.0, slogan_y, 7.3, 0.04, ORANGE)
    mtb(slide, 3.15, slogan_y+0.07, 7.0, 0.47, [
        (BOTTOM_SLOGAN, 13, True, ORANGE, PP_ALIGN.CENTER),
        (BOTTOM_SUB, 8, False, DARK_BLUE, PP_ALIGN.CENTER),
    ])

    return prs

def _hh(rgb, factor=1.0):
    clamp = lambda v: min(255, max(0, int(v*factor)))
    if isinstance(rgb, tuple):
        return f'{clamp(rgb[0]):02X}{clamp(rgb[1]):02X}{clamp(rgb[2]):02X}'
    return f'{clamp(rgb.red):02X}{clamp(rgb.green):02X}{clamp(rgb.blue):02X}'

if __name__ == '__main__':
    out_dir = os.path.dirname(os.path.abspath(__file__))
    prs = create_pdca()
    prs.save(os.path.join(out_dir, '模板15-PDCA循环改进环.pptx'))
    print('已生成: 模板15-PDCA循环改进环.pptx')
