# -*- coding: utf-8 -*-
"""模板12：多维能力雷达图 (Multi-Dimensional Capability Radar)
来源：实施报告12图模板200页 Slide 57/195 + 明大德 Slide 11 + 建筑工程 Slide 3
适用场景：学情分析雷达图、能力画像、教学目标达成度、多维素质评估
结构：中心五/六边形 → N条能力轴 → 评分点连线 → 填充区域
视觉增强：渐变填充多边形 + 阴影节点 + 轴标签装饰
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
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

# ============================================================
# ==== 可修改区 ==============================================
# ============================================================

PAGE_TITLE = '▎图X：学生多维能力画像雷达图'
PAGE_SUBTITLE = '知识 · 技能 · 素养 三维评估  |  班级平均 vs 个人画像'

# 雷达图维度定义（6维度）
RADAR_DIMENSIONS = [
    {'name': '理论知识', 'score': 82, 'max': 100, 'color': B1},
    {'name': '实操技能', 'score': 88, 'max': 100, 'color': B2},
    {'name': '创新能力', 'score': 72, 'max': 100, 'color': B3},
    {'name': '团队协作', 'score': 90, 'max': 100, 'color': B4},
    {'name': '职业素养', 'score': 85, 'max': 100, 'color': GREEN},
    {'name': '工程思维', 'score': 76, 'max': 100, 'color': ORANGE},
]

# 对比组（虚线参考）
BENCHMARK_SCORES = [75, 78, 65, 80, 78, 68]

# 右侧能力详情
ABILITY_DETAILS = [
    {'icon': '知', 'name': '理论知识', 'detail': '专业基础扎实\n概念理解清晰', 'color': B1},
    {'icon': '技', 'name': '实操技能', 'detail': '动手能力强\n工具运用熟练', 'color': B2},
    {'icon': '创', 'name': '创新能力', 'detail': '善于发现问题\n提出改进方案', 'color': B3},
    {'icon': '合', 'name': '团队协作', 'detail': '沟通表达优秀\n团队配合默契', 'color': B4},
    {'icon': '德', 'name': '职业素养', 'detail': '工匠精神突出\n责任心强', 'color': GREEN},
    {'icon': '思', 'name': '工程思维', 'detail': '系统思考能力\n还需加强', 'color': ORANGE},
]

BOTTOM_SLOGAN = '"精准画像 · 因材施教"'
BOTTOM_SUB = '数据驱动的学情诊断与个性化培养'

# ============================================================
# ==== 生成逻辑 ==============================================
# ============================================================

def create_radar():
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

    # ---- 雷达图主体（左侧） ----
    radar_cx, radar_cy = 4.0, 4.2  # 中心（英寸）
    radar_r = 2.1  # 半径（英寸）
    n = len(RADAR_DIMENSIONS)

    # 雷达背景卡片
    rrect(slide, 0.45, 0.82, 7.1, 5.6, WHITE, border=LINE, shadow=True)
    tb(slide, 0.6, 0.92, 4.0, 0.2, '多维能力雷达图', size=10, bold=True, color=DARK_BLUE)

    # 绘制同心环（刻度线）
    for ring in [0.25, 0.5, 0.75, 1.0]:
        rr = radar_r * ring
        # 用正N边形近似圆
        pts = []
        for j in range(n):
            a = -math.pi/2 + j*2*math.pi/n
            x = radar_cx + rr*math.cos(a)
            y = radar_cy + rr*math.sin(a)
            pts.append((x, y))
        _draw_polygon(slide, pts, None, LINE2, 0.3)

    # 绘制轴线
    for i in range(n):
        a = -math.pi/2 + i*2*math.pi/n
        x = radar_cx + radar_r*math.cos(a)
        y = radar_cy + radar_r*math.sin(a)
        # 轴线
        _draw_line(slide, radar_cx, radar_cy, x, y, LINE2, 0.3)
        # 轴端点标签
        label_x = radar_cx + (radar_r+0.4)*math.cos(a)
        label_y = radar_cy + (radar_r+0.4)*math.sin(a)
        rrect(slide, label_x-0.55, label_y-0.15, 1.1, 0.3, RADAR_DIMENSIONS[i]['color'], radius=0.03)
        tb(slide, label_x-0.5, label_y-0.12, 1.0, 0.24, RADAR_DIMENSIONS[i]['name'],
           size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # 绘制基准多边形（虚线效果用浅色多边形近似）
    bench_pts = []
    for i in range(n):
        a = -math.pi/2 + i*2*math.pi/n
        r = radar_r * (BENCHMARK_SCORES[i]/100)
        x = radar_cx + r*math.cos(a)
        y = radar_cy + r*math.sin(a)
        bench_pts.append((x, y))
    _draw_polygon(slide, bench_pts, RGBColor(0xE0,0xE8,0xF0), LINE, 0.5)

    # 绘制实际评分的填充多边形
    score_pts = []
    for i in range(n):
        a = -math.pi/2 + i*2*math.pi/n
        r = radar_r * (RADAR_DIMENSIONS[i]['score']/100)
        x = radar_cx + r*math.cos(a)
        y = radar_cy + r*math.sin(a)
        score_pts.append((x, y))
    _draw_polygon(slide, score_pts, RGBColor(0x1E,0x57,0x93,), B1, 1.2, alpha_fill=0.3)

    # 评分节点圆
    for i in range(n):
        a = -math.pi/2 + i*2*math.pi/n
        r = radar_r * (RADAR_DIMENSIONS[i]['score']/100)
        x = radar_cx + r*math.cos(a)
        y = radar_cy + r*math.sin(a)
        nd = rrect(slide, x-0.1, y-0.1, 0.2, 0.2, B1, radius=0.5, shadow=True)
        tb(slide, x-0.06, y-0.06, 0.12, 0.12, str(RADAR_DIMENSIONS[i]['score']),
           size=6, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # 中心标签
    rrect(slide, radar_cx-0.55, radar_cy-0.25, 1.1, 0.5, DARK_BLUE, radius=0.06, shadow=True)
    mtb(slide, radar_cx-0.5, radar_cy-0.2, 1.0, 0.4, [
        ('综合', 8, False, RGBColor(0xB0,0xCC,0xE5), PP_ALIGN.CENTER),
        (f'{sum(d["score"] for d in RADAR_DIMENSIONS)//len(RADAR_DIMENSIONS)}', 14, True, WHITE, PP_ALIGN.CENTER),
    ])

    # ---- 右侧：能力详情卡片 ----
    right_x = 7.8; right_w = 5.0
    # 图例
    rrect(slide, right_x, 0.85, 0.28, 0.14, RGBColor(0xE0,0xE8,0xF0), border=LINE)
    tb(slide, right_x+0.34, 0.83, 2.0, 0.16, '班级平均', size=7, color=MED)
    rrect(slide, right_x+2.2, 0.85, 0.28, 0.14, B1, border=None)
    tb(slide, right_x+2.54, 0.83, 2.0, 0.16, '个人得分', size=7, color=MED)

    detail_top = 1.1
    for i, ab in enumerate(ABILITY_DETAILS):
        dy = detail_top + i*0.72
        ac = ab['color']

        rrect(slide, right_x, dy, right_w, 0.6, WHITE, border=LINE)
        rect(slide, right_x, dy+0.04, 0.05, 0.52, ac)
        # 图标
        rrect(slide, right_x+0.15, dy+0.1, 0.38, 0.38, ac, radius=0.5)
        tb(slide, right_x+0.15, dy+0.15, 0.38, 0.28, ab['icon'],
           size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # 名称 + 详情
        tb(slide, right_x+0.62, dy+0.05, 1.5, 0.2, ab['name'],
           size=9, bold=True, color=DARK)
        mtb(slide, right_x+0.62, dy+0.26, right_w-0.8, 0.3,
            [(ab['detail'], 7, False, MED, PP_ALIGN.LEFT)])
        # 评分条
        pct = RADAR_DIMENSIONS[i]['score']/100
        rect(slide, right_x+0.62, dy+0.52, right_w-0.8, 0.04, LINE2)
        rect(slide, right_x+0.62, dy+0.52, (right_w-0.8)*pct, 0.04, ac)

    # ---- 底部标语 ----
    slogan_y = 6.6
    rrect(slide, 3.0, slogan_y, 7.3, 0.6, WHITE, border=LINE, shadow=True)
    rect(slide, 3.0, slogan_y, 7.3, 0.04, ORANGE)
    mtb(slide, 3.15, slogan_y+0.07, 7.0, 0.46, [
        (BOTTOM_SLOGAN, 13, True, ORANGE, PP_ALIGN.CENTER),
        (BOTTOM_SUB, 8, False, DARK_BLUE, PP_ALIGN.CENTER),
    ])

    return prs

def _draw_line(slide, x1, y1, x2, y2, color, width_pt=0.5):
    """Draw a connector line between two points."""
    conn = slide.shapes.add_connector(
        1,  # straight connector
        Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    conn.line.color.rgb = color
    conn.line.width = Pt(width_pt)

def _draw_polygon(slide, pts, fill_color, line_color, line_width=0.5, alpha_fill=None):
    """Draw a freeform polygon."""
    if len(pts) < 3: return
    builder = slide.shapes.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]))
    for x, y in pts[1:]:
        builder.add_line_segments([(Inches(x), Inches(y))], close=False)
    builder.add_line_segments([(Inches(pts[0][0]), Inches(pts[0][1]))], close=True)
    shape = builder.convert_to_shape()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    shape.line.color.rgb = line_color
    shape.line.width = Pt(line_width)

if __name__ == '__main__':
    out_dir = os.path.dirname(os.path.abspath(__file__))
    prs = create_radar()
    prs.save(os.path.join(out_dir, '模板12-多维能力雷达图.pptx'))
    print('已生成: 模板12-多维能力雷达图.pptx')
