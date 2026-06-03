# -*- coding: utf-8 -*-
"""模板11：教学数据仪表盘 (Teaching Data Dashboard)
来源：实施报告12图模板200页 Slide 59/106/114 + 劳动教育 Slide 5 + 明大德 Slide 13
适用场景：教学效果数据展示、评价权重分布、KPI达成率、多维评分
结构：顶部KPI数字卡片 → 中部多维评价矩阵 → 底部进度条对比
视觉增强：阴影卡片 + 渐变标题栏 + 比例进度条 + 装饰分割线
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
RED       = RGBColor(0xD9, 0x4F, 0x4F)

B1 = RGBColor(0x1E, 0x57, 0x93)
B2 = RGBColor(0x2E, 0x78, 0xB5)
B3 = RGBColor(0x4B, 0x9B, 0xD4)
B4 = RGBColor(0x7D, 0xC5, 0xED)

# ─── Enhanced Helpers ───────────────────────────────────────

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
    """Add subtle outer shadow to a shape."""
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
    """Replace solid fill with 2-stop linear gradient."""
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
        sh.fill.solid(); sh.fill.fore_color.rgb = fill  # fallback
        set_gradient_fill(sh, gradient[0], gradient[1])
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if border: sh.line.color.rgb = border; sh.line.width = Pt(0.5)
    else: sh.line.fill.background()
    sh.adjustments[0] = radius
    if shadow: add_shadow(sh)
    return sh

def rect(s, l, t, w, h, fill, border=None, shadow=False, gradient=None):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l),Inches(t),Inches(w),Inches(h))
    if gradient:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
        set_gradient_fill(sh, gradient[0], gradient[1])
    else:
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

def divider(s, x, y, w, color=LINE):
    """Decorative divider line with gradient dots."""
    rect(s, x, y, w, 0.012, color)
    for i in range(0, 6):
        dx = x + w*0.15 + i*w*0.14
        rrect(s, dx, y-0.03, 0.06, 0.06, color, radius=0.5)

# ============================================================
# ==== 可修改区 ==============================================
# ============================================================

PAGE_TITLE = '▎图X：教学效果多维数据仪表盘'
PAGE_SUBTITLE = '过程评价 · 结果评价 · 增值评价  |  数据驱动教学改进'

# 顶部 KPI 数字卡片
KPI_CARDS = [
    {'label': '课前测试\n平均分', 'value': '75.2', 'unit': '分', 'trend': '↑ 8.5%', 'color': B1},
    {'label': '课后测试\n平均分', 'value': '82.4', 'unit': '分', 'trend': '↑ 12.3%', 'color': B2},
    {'label': '实操技能\n合格率', 'value': '93.8', 'unit': '%', 'trend': '↑ 5.2%', 'color': B3},
    {'label': '学生\n满意度', 'value': '96.5', 'unit': '%', 'trend': '↑ 3.1%', 'color': GREEN},
]

# 多维评价矩阵
EVAL_DIMENSIONS = [
    {'name': '知识掌握', 'score': 82, 'max': 100, 'color': B1},
    {'name': '技能操作', 'score': 88, 'max': 100, 'color': B2},
    {'name': '职业素养', 'score': 90, 'max': 100, 'color': B3},
    {'name': '创新能力', 'score': 76, 'max': 100, 'color': B4},
    {'name': '团队协作', 'score': 85, 'max': 100, 'color': GREEN},
    {'name': '工程思维', 'score': 79, 'max': 100, 'color': ORANGE},
]

# 评价权重分布
WEIGHT_ITEMS = [
    {'name': '过程评价（课前+课中）', 'weight': 60, 'color': B1},
    {'name': '结果评价（期末+证书）', 'weight': 40, 'color': B2},
]

# 底部标语
BOTTOM_SLOGAN = '"数据驱动 · 精准教学"'
BOTTOM_SUB = '全过程 · 全维度 · 全主体评价体系'

# ============================================================
# ==== 生成逻辑 ==============================================
# ============================================================

def create_dashboard():
    prs = Presentation()
    prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = BG_OUTER
    rrect(slide, 0.35, 0.08, 12.65, 7.35, BG_MAIN, border=LINE, radius=0.03)

    # 标题栏（渐变）
    rrect(slide, 0.65, 0.18, 12.0, 0.48, DARK_BLUE, radius=0.04,
          gradient=('1E5793', '2E78B5'))
    tb(slide, 0.85, 0.22, 8.0, 0.28, PAGE_TITLE, size=16, bold=True, color=WHITE)
    tb(slide, 0.85, 0.48, 10.0, 0.16, PAGE_SUBTITLE, size=7.5, color=RGBColor(0xB0,0xCC,0xE5))

    # 装饰分割线
    divider(slide, 0.65, 0.78, 12.0, LINE)

    # ---- 顶部 KPI 数字卡片 ----
    kpi_top = 0.92
    kpi_w = 2.85; kpi_gap = 0.18
    kpi_h = 1.3

    for i, kpi in enumerate(KPI_CARDS):
        kx = 0.55 + i*(kpi_w + kpi_gap)
        kc = kpi['color']

        # 阴影卡片
        rrect(slide, kx, kpi_top, kpi_w, kpi_h, WHITE, border=LINE, shadow=True)
        # 顶部渐变色条
        rrect(slide, kx, kpi_top, kpi_w, 0.06, kc, gradient=(rgb_to_hex(kc), rgb_to_hex(kc)))

        # KPI 标签
        mtb(slide, kx+0.12, kpi_top+0.15, kpi_w-0.24, 0.3,
            [(kpi['label'], 7, False, MED, PP_ALIGN.LEFT)])
        # KPI 数值（大字号）
        tb(slide, kx+0.12, kpi_top+0.42, kpi_w-0.75, 0.5, kpi['value'],
           size=28, bold=True, color=kc)
        # 单位
        tb(slide, kx+kpi_w-0.6, kpi_top+0.55, 0.45, 0.22, kpi['unit'],
           size=10, color=MED)
        # 趋势标签
        rrect(slide, kx+0.12, kpi_top+kpi_h-0.35, 0.85, 0.24, CARD_LT, radius=0.03)
        tb(slide, kx+0.16, kpi_top+kpi_h-0.33, 0.77, 0.2, kpi['trend'],
           size=7.5, bold=True, color=GREEN, align=PP_ALIGN.CENTER)

        # 装饰迷你进度条
        pct = min(1.0, float(kpi['value'].replace('%',''))/100 if '%' in kpi['value'] else float(kpi['value'])/100)
        rect(slide, kx+0.12, kpi_top+kpi_h-0.48, kpi_w-0.24, 0.04, LINE2)
        rect(slide, kx+0.12, kpi_top+kpi_h-0.48, (kpi_w-0.24)*pct, 0.04, kc)

    # ---- 中部：多维评价雷达条 + 权重分布 ----
    dim_top = kpi_top + kpi_h + 0.25

    # 左：评分进度条
    left_w = 6.2
    rrect(slide, 0.55, dim_top, left_w, 3.65, WHITE, border=LINE, shadow=True)
    tb(slide, 0.7, dim_top+0.08, 4.0, 0.22, '多维能力评分', size=10, bold=True, color=DARK_BLUE)

    bar_top = dim_top + 0.38
    bar_h = 0.3; bar_gap = 0.16

    for i, dim in enumerate(EVAL_DIMENSIONS):
        by = bar_top + i*(bar_h + bar_gap)
        dc = dim['color']

        # 维度名称
        tb(slide, 0.7, by, 1.4, 0.24, dim['name'], size=8, bold=True, color=DARK)
        # 进度条背景
        rrect(slide, 2.15, by+0.02, 3.6, 0.22, LINE2, border=None, radius=0.03)
        # 进度条填充
        pct = dim['score']/dim['max']
        if pct > 0:
            rrect(slide, 2.15, by+0.02, 3.6*pct, 0.22, dc, radius=0.03,
                  gradient=(rgb_to_hex(dc), adjust_brightness(dc, 1.25)))
        # 分数
        rrect(slide, 5.9, by-0.02, 0.65, 0.28, dc, radius=0.03)
        tb(slide, 5.9, by, 0.65, 0.24, str(dim['score']),
           size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # 右：评价权重
    right_x = 7.0; right_w = 5.65
    rrect(slide, right_x, dim_top, right_w, 1.45, WHITE, border=LINE, shadow=True)
    tb(slide, right_x+0.12, dim_top+0.08, 3.0, 0.22, '评价权重分布', size=10, bold=True, color=DARK_BLUE)

    for i, wi in enumerate(WEIGHT_ITEMS):
        wy = dim_top + 0.42 + i*0.42
        tb(slide, right_x+0.2, wy, 2.5, 0.2, wi['name'], size=8, color=DARK)
        # 权重数字
        rrect(slide, right_x+right_w-1.1, wy-0.02, 0.85, 0.28, wi['color'], radius=0.04)
        tb(slide, right_x+right_w-1.1, wy, 0.85, 0.24, f'{wi["weight"]}%',
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # 权重块可视化
        rect(slide, right_x+0.2, wy+0.28, right_w-1.5, 0.08, LINE2)
        rect(slide, right_x+0.2, wy+0.28, (right_w-1.5)*wi['weight']/100, 0.08, wi['color'])

    # 下部右侧：评价主体
    subj_top = dim_top + 1.65
    rrect(slide, right_x, subj_top, right_w, 2.1, WHITE, border=LINE, shadow=True)
    tb(slide, right_x+0.12, subj_top+0.08, 3.0, 0.22, '多元评价主体', size=10, bold=True, color=DARK_BLUE)

    subjects = [
        ('教师评价', '30%', B1), ('学生自评', '20%', B2),
        ('小组互评', '20%', B3), ('企业导师', '20%', B4),
        ('平台数据', '10%', GREEN),
    ]
    for i, (sn, sw, sc) in enumerate(subjects):
        col = i % 2; row = i // 2
        sx = right_x + 0.18 + col*2.65
        sy = subj_top + 0.42 + row*0.65

        rrect(slide, sx, sy, 2.45, 0.5, CARD_LT, border=None, radius=0.03)
        rrect(slide, sx+0.08, sy+0.1, 0.3, 0.3, sc, radius=0.5)
        tb(slide, sx+0.08, sy+0.13, 0.3, 0.24, str(i+1),
           size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, sx+0.48, sy+0.06, 1.2, 0.2, sn, size=9, bold=True, color=DARK)
        rrect(slide, sx+1.7, sy+0.1, 0.6, 0.24, WHITE, radius=0.03)
        tb(slide, sx+1.7, sy+0.12, 0.6, 0.2, sw,
           size=9, bold=True, color=sc, align=PP_ALIGN.CENTER)

    # ---- 底部标语 ----
    slogan_y = 6.55
    rrect(slide, 3.0, slogan_y, 7.3, 0.65, WHITE, border=LINE, shadow=True)
    rect(slide, 3.0, slogan_y, 7.3, 0.05, ORANGE, gradient=(rgb_to_hex(ORANGE), 'FAB86C'))
    mtb(slide, 3.15, slogan_y+0.08, 7.0, 0.5, [
        (BOTTOM_SLOGAN, 13, True, ORANGE, PP_ALIGN.CENTER),
        (BOTTOM_SUB, 9, False, DARK_BLUE, PP_ALIGN.CENTER),
    ])

    return prs

def rgb_to_hex(rgb):
    if hasattr(rgb, 'red'):
        return f'{rgb.red:02X}{rgb.green:02X}{rgb.blue:02X}'
    if isinstance(rgb, (tuple, list)):
        return f'{int(rgb[0]):02X}{int(rgb[1]):02X}{int(rgb[2]):02X}'
    return str(rgb)

def adjust_brightness(rgb, factor):
    clamp = lambda v: min(255, max(0, int(v)))
    if hasattr(rgb, 'red'):
        r, g, b = rgb.red, rgb.green, rgb.blue
    elif isinstance(rgb, (tuple, list)):
        r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
    else:
        return rgb
    return f'{clamp(r*factor):02X}{clamp(g*factor):02X}{clamp(b*factor):02X}'

if __name__ == '__main__':
    out_dir = os.path.dirname(os.path.abspath(__file__))
    prs = create_dashboard()
    prs.save(os.path.join(out_dir, '模板11-教学数据仪表盘.pptx'))
    print('已生成: 模板11-教学数据仪表盘.pptx')
