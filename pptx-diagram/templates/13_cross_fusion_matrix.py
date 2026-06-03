# -*- coding: utf-8 -*-
"""模板13：交叉融合矩阵 (Cross-Fusion Matrix)
来源：29优秀模板 Slide 9 + 200页模板 Slide 185 + 重构框架 Slide 2
适用场景：多维度交叉融合、岗课赛证矩阵、课程内容交叉映射、SWOT分析
结构：横轴×纵轴交叉矩阵 → 交叉节点标注 → 底部融合说明
视觉增强：阴影卡片 + 双色区隔 + 交叉亮点标注
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

# ============================================================
# ==== 可修改区 ==============================================
# ============================================================

PAGE_TITLE = '▎图X：多维度交叉融合矩阵'
PAGE_SUBTITLE = '横轴：教学维度  ×  纵轴：能力层级  |  交叉节点映射课程目标'

# 横轴标签（列标题）
COL_HEADERS = ['知识目标', '技能目标', '素养目标']
COL_COLORS = [B1, B2, B3]

# 纵轴标签（行标题）
ROW_HEADERS = [
    {'name': '基础认知', 'color': B1},
    {'name': '理解应用', 'color': B2},
    {'name': '综合创新', 'color': B3},
    {'name': '实践评价', 'color': B4},
]

# 矩阵交叉节点内容（row×col）
MATRIX_CELLS = [
    # 第1行：基础认知
    [
        {'text': '概念理解\n术语掌握', 'highlight': False},
        {'text': '操作规范\n流程认知', 'highlight': False},
        {'text': '职业道德\n法规意识', 'highlight': False},
    ],
    # 第2行：理解应用
    [
        {'text': '原理分析\n知识迁移', 'highlight': True},
        {'text': '工具运用\n方法选择', 'highlight': False},
        {'text': '价值判断\n伦理分析', 'highlight': False},
    ],
    # 第3行：综合创新
    [
        {'text': '系统设计\n批判思维', 'highlight': False},
        {'text': '方案创新\n技术整合', 'highlight': True},
        {'text': '匠心精神\n追求卓越', 'highlight': False},
    ],
    # 第4行：实践评价
    [
        {'text': '评估反思\n知识建构', 'highlight': False},
        {'text': '成果检验\n过程优化', 'highlight': False},
        {'text': '社会责任\n持续改进', 'highlight': True},
    ],
]

BOTTOM_SLOGAN = '"纵横贯通 · 融合育人"'
BOTTOM_SUB = '三维教学目标 × 四层能力递进  =  12节点课程映射矩阵'

# ============================================================
# ==== 生成逻辑 ==============================================
# ============================================================

def create_matrix():
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

    # ---- 矩阵布局 ----
    n_rows = len(ROW_HEADERS)
    n_cols = len(COL_HEADERS)

    # 左上角留白
    corner_w = 1.8; corner_h = 0.55
    rrect(slide, 0.55, 0.85, corner_w, corner_h, DARK_BLUE, radius=0.04)
    tb(slide, 0.55, 0.92, corner_w, 0.3, '能力层级', size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, 0.55, 1.15, corner_w, 0.16, '知识维度', size=7, color=RGBColor(0xB0,0xCC,0xE5), align=PP_ALIGN.CENTER)

    # 列标题
    col_w = 3.3; col_start_x = 0.55 + corner_w + 0.12
    col_header_h = 0.55
    for ci, (ch, cc) in enumerate(zip(COL_HEADERS, COL_COLORS)):
        cx = col_start_x + ci*col_w
        rrect(slide, cx, 0.85, col_w, col_header_h, cc, radius=0.04,
              gradient=(_hh(cc), _hh(cc, 1.15)))
        tb(slide, cx, 0.92, col_w, 0.3, ch,
           size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # 行标题 + 矩阵单元格
    row_h = 1.3; row_start_y = 0.85 + corner_h + 0.1

    for ri, rh in enumerate(ROW_HEADERS):
        ry = row_start_y + ri*row_h
        rc = rh['color']

        # 行标题
        rrect(slide, 0.55, ry, corner_w, row_h, rc, radius=0.04,
              gradient=(_hh(rc), _hh(rc, 1.2)))
        tb(slide, 0.55, ry+row_h/2-0.15, corner_w, 0.3, rh['name'],
           size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # 矩阵单元格
        for ci in range(n_cols):
            cx = col_start_x + ci*col_w
            cell = MATRIX_CELLS[ri][ci]

            cell_fill = WHITE
            cell_border = LINE
            if cell['highlight']:
                cell_fill = RGBColor(0xE8, 0xF4, 0xFD)

            rrect(slide, cx, ry, col_w, row_h, cell_fill, border=cell_border, shadow=cell['highlight'])
            # 左侧色条
            rect(slide, cx, ry+0.08, 0.05, row_h-0.16, rc)
            # 顶部色条
            rect(slide, cx+0.1, ry, col_w-0.2, 0.03, COL_COLORS[ci])

            mtb(slide, cx+0.2, ry+0.15, col_w-0.4, row_h-0.3,
                [(cell['text'], 8, cell['highlight'], DARK if not cell['highlight'] else B1, PP_ALIGN.CENTER)])

            # 如果高亮，加标注角标
            if cell['highlight']:
                rrect(slide, cx+col_w-0.35, ry+0.06, 0.28, 0.2, ORANGE, radius=0.03)
                tb(slide, cx+col_w-0.35, ry+0.07, 0.28, 0.18, '融合',
                   size=6.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # ---- 矩阵下方说明 ----
    legend_y = row_start_y + n_rows*row_h + 0.15

    # 说明卡片
    rrect(slide, 0.55, legend_y, 12.15, 0.7, WHITE, border=LINE, shadow=True)
    rect(slide, 0.55, legend_y, 0.07, 0.7, ORANGE)

    # 图例
    rrect(slide, 0.8, legend_y+0.15, 0.2, 0.2, RGBColor(0xE8,0xF4,0xFD), border=LINE)
    tb(slide, 1.1, legend_y+0.15, 6.0, 0.2, '高亮节点 = 课程核心融合点（岗课赛证交汇处）', size=7.5, color=MED)
    tb(slide, 0.8, legend_y+0.4, 10.0, 0.2, '12 个交叉节点覆盖全部教学目标 · 4 个核心融合点对接岗位能力与证书标准', size=7.5, color=MED_BLUE)

    # ---- 底部标语 ----
    slogan_y = legend_y + 0.9
    rrect(slide, 3.0, slogan_y, 7.3, 0.55, WHITE, border=LINE, shadow=True)
    rect(slide, 3.0, slogan_y, 7.3, 0.04, ORANGE)
    mtb(slide, 3.15, slogan_y+0.06, 7.0, 0.43, [
        (BOTTOM_SLOGAN, 12, True, ORANGE, PP_ALIGN.CENTER),
        (BOTTOM_SUB, 7.5, False, DARK_BLUE, PP_ALIGN.CENTER),
    ])

    return prs

def _hh(rgb, factor=1.0):
    """RGBColor to hex string, optionally lightened."""
    clamp = lambda v: min(255, max(0, int(v*factor)))
    return f'{clamp(rgb[0]):02X}{clamp(rgb[1]):02X}{clamp(rgb[2]):02X}' if isinstance(rgb, tuple) else f'{clamp(rgb.red):02X}{clamp(rgb.green):02X}{clamp(rgb.blue):02X}'

if __name__ == '__main__':
    out_dir = os.path.dirname(os.path.abspath(__file__))
    prs = create_matrix()
    prs.save(os.path.join(out_dir, '模板13-交叉融合矩阵.pptx'))
    print('已生成: 模板13-交叉融合矩阵.pptx')
