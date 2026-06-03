---
name: pptx-diagram
description: Use python-pptx to programmatically generate editable PPTX block diagrams, flowcharts, and framework charts with native PowerPoint shapes. Use when the user asks to create diagrams for teaching competitions, presentations, or reports — especially Chinese teaching ability contest (教学能力大赛) materials. Covers three-column pipeline layouts, card-based information architecture, color system design, and iterative refinement (v1→v4 workflow).
---

# PPTX Diagram Generator (python-pptx)

Generate fully editable PPTX block diagrams programmatically. Output is native PowerPoint shapes (not images), so text, colors, and layout remain editable in PowerPoint.

## When to Use

- User asks to create block diagrams, flowcharts, framework charts, or knowledge graphs as `.pptx`
- User mentions "框图", "流程图", "架构图", "教学实施报告配图", "比赛PPT"
- User wants editable PPTX output (not PNG/image export)
- User needs to batch-generate diagrams with consistent styling

## Core Boilerplate

Every script starts with this skeleton:

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os

CN_FONT = 'PingFang SC'

# ---- Slide setup ----
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

# ---- Background ----
bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = BG_OUTER

# ---- Save ----
out_dir = os.path.dirname(os.path.abspath(__file__))
prs.save(os.path.join(out_dir, 'output.pptx'))
```

## Chinese Font Handling (CRITICAL)

Chinese text in python-pptx needs explicit East Asian font assignment via XML manipulation. Always use this `cn()` helper for every text run:

```python
def cn(run, name=CN_FONT, size=12, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', name)
```

Without this, Chinese characters will fall back to Calibri and render poorly in PowerPoint.

## Color System (Award-Winning Reference Palette)

Based on analysis of national second-prize winning (国二) teaching contest PPTX. Core principle: **unified blue gradient system, orange only for critical emphasis.**

```python
# ---- Base colors ----
BG_OUTER  = RGBColor(0xF8, 0xFB, 0xFE)  # page background
BG_MAIN   = RGBColor(0xF5, 0xFA, 0xFF)  # main container fill
CARD_LT   = RGBColor(0xD5, 0xF4, 0xFF)  # light blue card fill
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK      = RGBColor(0x1A, 0x1A, 0x2E)   # primary text
MED       = RGBColor(0x55, 0x55, 0x55)   # secondary text
DARK_BLUE = RGBColor(0x1A, 0x3C, 0x5E)   # headers, exits, strong emphasis
MED_BLUE  = RGBColor(0x2C, 0x5F, 0x8A)   # arrows, labels
LINE      = RGBColor(0xD0, 0xDE, 0xEB)   # card borders
LINE2     = RGBColor(0xE5, 0xEC, 0xF3)   # subtle dividers
ORANGE    = RGBColor(0xF4, 0x79, 0x20)   # ONLY for key emphasis (slogan, accent bar)

# ---- Blue gradient (dark→light, for 4-stage sequences) ----
BLUE1 = RGBColor(0x1E, 0x57, 0x93)  # deepest
BLUE2 = RGBColor(0x2E, 0x78, 0xB5)
BLUE3 = RGBColor(0x4B, 0x9B, 0xD4)
BLUE4 = RGBColor(0x7D, 0xC5, 0xED)  # lightest

# ---- Light backgrounds for cards (matching each blue) ----
L_BLUE1 = RGBColor(0xE6, 0xF0, 0xF8)
L_BLUE2 = RGBColor(0xEC, 0xF4, 0xFA)
L_BLUE3 = RGBColor(0xF2, 0xF7, 0xFC)
L_BLUE4 = RGBColor(0xF6, 0xFA, 0xFD)
```

**Color rules:**
- Use `BLUE1→BLUE4` for sequences with natural progression (stages, levels, time)
- Only use `ORANGE` in 2-3 places per slide (slogan text, key accent bar on most important card)
- All cards use `WHITE` fill with `LINE` border; `CARD_LT` for secondary/background cards
- `DARK_BLUE` for headers, exit nodes, main titles
- `MED_BLUE` for arrows, connecting labels, module tags

## Shape Helper Functions

```python
def rrect(slide, l, t, w, h, fill, border=None, radius=0.04):
    """Rounded rectangle. radius=0.5 for circles, 0.03-0.04 for cards."""
    sh = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if border:
        sh.line.color.rgb = border; sh.line.width = Pt(0.5)
    else:
        sh.line.fill.background()
    sh.adjustments[0] = radius
    return sh

def rect(slide, l, t, w, h, fill, border=None):
    """Rectangle. Thin rectangles (h < 0.1) used as accent bars / dividers."""
    sh = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if border:
        sh.line.color.rgb = border; sh.line.width = Pt(0.5)
    else:
        sh.line.fill.background()
    return sh

def tb(slide, l, t, w, h, text, size=10, bold=False, color=DARK, align=PP_ALIGN.LEFT):
    """Single-line text box."""
    bx = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    bx.text_frame.word_wrap = True; bx.text_frame.auto_size = None
    p = bx.text_frame.paragraphs[0]
    p.alignment = align; p.space_after = Pt(0); p.space_before = Pt(0)
    r = p.add_run(); r.text = text
    cn(r, size=size, bold=bold, color=color)
    return bx

def mtb(slide, l, t, w, h, lines):
    """Multi-line text box. lines = [(text, size, bold, color, align), ...]"""
    bx = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    bx.text_frame.word_wrap = True; bx.text_frame.auto_size = None
    for i, (text, sz, bd, cl, al) in enumerate(lines):
        p = bx.text_frame.paragraphs[0] if i == 0 else bx.text_frame.add_paragraph()
        p.alignment = al; p.space_after = Pt(1); p.space_before = Pt(0)
        r = p.add_run(); r.text = text
        cn(r, size=sz, bold=bd, color=cl)
    return bx

def arrow_r(slide, x, y, w, h, color):
    """Right-pointing arrow."""
    a = slide.shapes.add_shape(
        MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    a.fill.solid(); a.fill.fore_color.rgb = color; a.line.fill.background()

def arrow_d(slide, x, y, w, h, color):
    """Down-pointing arrow."""
    a = slide.shapes.add_shape(
        MSO_SHAPE.DOWN_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    a.fill.solid(); a.fill.fore_color.rgb = color; a.line.fill.background()
```

## Key Layout Patterns

### 1. Page Shell (always present)

```python
# Full-page light blue rounded container
rrect(slide, 0.25, 0.08, 12.85, 7.35, BG_MAIN, border=LINE, radius=0.03)

# Top title bar
tb(slide, 0.55, 0.18, 8.0, 0.32, '▎Page Title', size=18, bold=True, color=DARK_BLUE)
tb(slide, 0.55, 0.52, 10.0, 0.18, 'Subtitle line', size=8, color=MED)
rect(slide, 0.55, 0.76, 12.0, 0.01, LINE)  # horizontal divider
```

### 2. Three-Column Layout (most common pattern)

```python
# Column headers (dark blue bars spanning each column)
for col_l, col_w, title, subtitle in [
    (L_X, L_W, 'Left Title', 'Subtitle'),
    (M_X, M_W, 'Middle Title', 'Subtitle'),
    (R_X, R_W, 'Right Title', 'Subtitle'),
]:
    rect(slide, col_l, y, col_w, 0.32, DARK_BLUE)
    tb(slide, col_l, y + 0.02, col_w, 0.17, title, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, col_l, y + 0.18, col_w, 0.13, subtitle, size=7, color=RGBColor(0xB0,0xCC,0xE5), align=PP_ALIGN.CENTER)
```

Typical column widths for 13.333" canvas:
```
Left:   0.55 + 2.8  = 3.35   (context/input)
Middle: 3.65 + 5.9  = 9.55   (core content)
Right:  9.75 + 3.0  = 12.75  (output/goal)
```

### 3. Card with Left Accent Bar

The most common visual unit — a white card with colored left strip:

```python
def add_accent_card(slide, l, t, w, h, accent_color):
    rrect(slide, l, t, w, h, WHITE, border=LINE)
    rect(slide, l + 0.06, t + 0.08, 0.05, h - 0.16, accent_color)
```

### 4. Numbered Badge (circular)

```python
rrect(slide, x, y, 0.38, 0.38, color, radius=0.5)  # radius=0.5 = circle
tb(slide, x, y + 0.04, 0.38, 0.26, '01', size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
```

### 5. Stage Flow with Down Arrows

Used for sequential steps (four-stage progressive, module pipeline):

```python
for i, (label, name, desc, color) in enumerate(stages):
    sy = top_y + i * (stage_h + gap)
    rrect(slide, l, sy, w, stage_h, WHITE, border=LINE)
    # Left color bar
    rect(slide, l, sy + 0.05, 0.07, stage_h - 0.1, color)
    # Stage label badge
    rrect(slide, l + 0.15, sy + 0.1, 0.6, 0.28, color, radius=0.03)
    tb(slide, l + 0.15, sy + 0.12, 0.6, 0.24, label, size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # Name + description...
    # Arrow between stages
    if i < len(stages) - 1:
        arrow_d(slide, l + 0.18, sy + stage_h + 0.02, 0.12, 0.11, color)
```

### 6. Keyword Tags Row

Render keywords as small colored labels at bottom of a card:

```python
keywords = desc.split(' · ')
kx = start_x
for kw in keywords:
    kw_w = len(kw) * 0.11 + 0.18  # approximate width from character count
    rrect(slide, kx, tag_y, kw_w, 0.22, CARD_LT, radius=0.02)
    tb(slide, kx + 0.04, tag_y + 0.01, kw_w - 0.08, 0.20, kw, size=6.5, color=MED_BLUE)
    kx += kw_w + 0.06
```

### 7. Entry/Exit Nodes

Flow endpoints use light blue fill (entry) or dark blue fill (exit):

```python
# Entry
rrect(slide, l, y, w, 0.38, CARD_LT, border=None, radius=0.03)
tb(slide, l + 0.1, y + 0.05, 0.4, 0.18, '入口', size=7, color=MED_BLUE)

# Exit (dark blue, white text)
rrect(slide, l, y, w, 0.38, DARK_BLUE, border=None, radius=0.03)
tb(slide, l + 0.1, y + 0.05, 0.4, 0.18, '出口', size=7, color=RGBColor(0xB0,0xCC,0xE5))
```

### 8. Evaluation/Proportion Bar

Horizontal segmented bar showing percentages:

```python
for etitle, ec, eratio in eval_data:
    seg_w = bar_w * eratio
    rrect(slide, seg_x, y, seg_w - 0.03, bar_h, ec, radius=0.03)
    tb(slide, seg_x, y + 0.02, seg_w - 0.03, bar_h - 0.04,
        etitle, size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    seg_x += seg_w
```

## Design Language (from National Award-Winning Reference)

Extracted from analyzing "零碳工厂" (Zero-Carbon Factory) second-prize winning PPTX:

1. **Large light-blue container** (`#F5FAFF` rounded rectangle) wraps the entire content
2. **Light blue card fills** (`#D5F4FF`) — not pure white — for secondary/background cards
3. **Left → Middle → Right pipeline** layout: demands → strategy → implementation → output
4. **Horizontal zones** separated by whitespace and thin lines, not explicit borders
5. **Clean rectangles only** — no diamonds, hexagons, or fancy shapes (award-winning work uses pure rectangles)
6. **Dense but clean** — 130-200 shapes looks organized when colors are unified
7. **Fill the canvas** — extend content to ~y=7.0 (of 7.5" height), leave no large empty zones

## Enhanced Visual Effects (v5)

New templates (11-17) include these production-quality effects derived from award-winning PPTX analysis:

### Shadow (outerShdw)
```python
from lxml import etree

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
```

### Gradient Fill (2-stop linear)
```python
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
```

### Updated `rrect()` with shadow + gradient support
```python
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
```

Usage: `rrect(slide, x, y, w, h, WHITE, border=LINE, shadow=True)` or `rrect(slide, x, y, w, h, B1, gradient=('1E5793','3D7AB5'))`

## Template Catalog (17 total)

Each template is a standalone `.py` script in `templates/`. Run individually to generate `.pptx`.

### 基础结构类 (Layout & Structure)
| # | Template | Source | Best For |
|---|----------|--------|----------|
| 01 | `01_three_column_pipeline.py` | 零碳工厂/交通运输/视觉传感小车 | 三栏流水线：输入→过程→产出 |
| 02 | `02_four_stage_progressive.py` | 思政-明大德/交通运输/建筑工程 | 四阶递进课程体系 |
| 03 | `03_ksa_matrix.py` | 交通运输/建筑工程/视觉传感小车 | 知识/技能/素养三维目标 |
| 04 | `04_center_radial.py` | 思政-明大德/视觉传感小车/寻路立心 | 中心放射四驱联动 |
| 05 | `05_gksz_integration.py` | 建筑工程/重构框架/智能摆渡车 | 岗课赛证四维融通 |

### 教学流程类 (Teaching Process)
| # | Template | Source | Best For |
|---|----------|--------|----------|
| 06 | `06_teaching_process_timeline.py` | Z62/视觉传感小车/劳动教育 | 课前→课中→课后三段流程 |
| 07 | `07_boppps_model.py` | Z68/寻路立心/交通运输 | BOPPPS六节点教学设计 |
| 16 | `16_teaching_process_loop.py` | Z62/200页模板/重构框架 | 双线并进教学闭环 |

### 思政与评价类 (Ideology & Evaluation)
| # | Template | Source | Best For |
|---|----------|--------|----------|
| 08 | `08_ideological_integration.py` | 思政-明大德/交通运输 | 课程思政融入渗透体系 |
| 09 | `09_evaluation_dashboard.py` | 劳动教育/明大德/建筑工程 | 多维评价体系仪表盘 |
| 11 | `11_teaching_dashboard.py` ✨ | 200页模板/劳动教育 | KPI仪表盘+进度条+权重 |
| 12 | `12_capability_radar.py` ✨ | 200页模板/明大德/建筑工程 | 六维能力雷达画像 |

### 关系与对比类 (Relationship & Comparison)
| # | Template | Source | Best For |
|---|----------|--------|----------|
| 10 | `10_course_panorama.py` | 重构框架/建筑工程 | 课程体系全景图 |
| 13 | `13_cross_fusion_matrix.py` ✨ | 29模板/200页模板 | 多维度交叉融合矩阵 |
| 14 | `14_comparison_dual_column.py` ✨ | 200页模板/明大德 | 教学重难点对比双栏 |

### 循环与导航类 (Cycle & Navigation)
| # | Template | Source | Best For |
|---|----------|--------|----------|
| 15 | `15_pdca_cycle.py` ✨ | 200页模板/重构框架 | PDCA四象限质量改进 |
| 17 | `17_cover_navigation.py` ✨ | 200页模板/29模板 | 封面/目录/章节/结尾页 |

✨ = New v5 templates with shadow + gradient effects.

## Typography Tiers (from 998-slide analysis)

| Tier | Size | Usage |
|------|------|-------|
| Title | 16-18pt | Page title (页标题) |
| Section | 12-14pt | Section header (区标题) |
| Card Title | 10-11pt | Card/column title (卡片标题) |
| Body | 8-9pt | Body text (正文) |
| Note | 7-7.5pt | Annotation/footer (标注/脚注) |

## Iteration Workflow

Follow this proven 4-version approach:

| Version | Focus | What to do |
|---------|-------|------------|
| **v1** | Layout logic | Get the structure right — correct nodes, correct flow, all text placed. Colors can be rough. |
| **v2** | Reference alignment | Find a real award-winning PPTX in the project, parse its shapes/colors, adopt its design language. Match card fills, fonts, spacing. |
| **v3** | Canvas coverage | Fill empty zones. Add detail rows (sub-processes, supporting info) so content extends to y ≈ 7.0. |
| **v4** | Color unification | Reduce to one color family + one accent. All blues should be from the same gradient. Orange only in 2-3 places. |

**Never stop at v1** — the first version always looks like "programmer art." Iteration is the key difference.

## Anti-Patterns to Avoid

- **Do NOT** use ChatGPT image generation for primary diagrams — PNGs can't be edited later, font/color changes require re-generation
- **Do NOT** use many color families — v3 with orange/cyan/purple/green looks chaotic; always unify to blue gradient in v4
- **Do NOT** leave large empty zones — judges interpret blank space as missing content
- **Do NOT** copy reference pages verbatim with `clone_slide()` — it copies shapes including bugs; instead, write fresh scripts that recreate the design language
- **Do NOT** over-engineer with classes or factories — a single procedural script per diagram is fine
- **Do NOT** use font sizes below 6.5pt — unreadable in projection

## Quick Reference: MSO_SHAPE Types

```python
MSO_SHAPE.RECTANGLE          # rectangle
MSO_SHAPE.ROUNDED_RECTANGLE  # rounded rect (use adjustments[0] for radius)
MSO_SHAPE.RIGHT_ARROW        # right arrow
MSO_SHAPE.DOWN_ARROW         # down arrow
MSO_SHAPE.LEFT_ARROW         # left arrow
MSO_SHAPE.CHEVRON            # chevron/arrowhead
MSO_SHAPE.DIAMOND            # diamond (rarely used — stick to rectangles)
MSO_SHAPE.OVAL               # circle/ellipse
```

## Post-Generation Checklist

After running the script:
1. Open the `.pptx` in PowerPoint — verify all Chinese renders correctly
2. Check that no text overflows its bounding box
3. Verify color consistency — no stray colors from earlier versions
4. Confirm all placeholder text has been replaced with real content
5. Check edge elements don't extend beyond slide boundaries (13.333 × 7.5)
