# pptx-diagram — Claude Code Skill

> 用 python-pptx 生成可编辑的教学比赛 PPT 框图 / Flowchart / 架构图

## 这是什么？

`pptx-diagram` 是一个 [Claude Code](https://claude.ai/code) 的 **Skill**，让 Claude 可以用 python-pptx 程序化生成 PPTX 框图。输出是 **原生 PowerPoint 形状**（不是图片），所以文字、颜色、布局在 PowerPoint 里都可以二次编辑。

主要面向 **教学能力大赛**（全国职业院校技能大赛教学能力比赛）的参赛材料制作。

## 功能特点

- 🎨 **国二获奖作品配色体系** — 基于真实获奖 PPTX 提取的蓝色渐变系统
- 📐 **10 套模板** — 三栏管线、四阶递进、KSA矩阵、BOPPPS模型、课程全景图等
- 🔤 **中文字体完美支持** — 通过 XML 注入东亚字体，不会 fallback 到 Calibri
- ✏️ **完全可编辑** — 生成的是原生 PPTX 形状，不是截图
- 🔄 **v1→v4 迭代流程** — 从结构到配色，四版打磨出精品

## 模板列表

| # | 模板 | 用途 |
|---|------|------|
| 01 | `three_column_pipeline` | 三栏管线图 — 问题→策略→成效 |
| 02 | `four_stage_progressive` | 四阶递进 — 阶段化教学过程 |
| 03 | `ksa_matrix` | KSA 矩阵 — 知识/技能/态度分析 |
| 04 | `center_radial` | 中心辐射图 — 核心理念向外发散 |
| 05 | `gksz_integration` | 岗课赛证融合模型 |
| 06 | `teaching_process_timeline` | 教学流程时间线 |
| 07 | `boppps_model` | BOPPPS 有效教学模型 |
| 08 | `ideological_integration` | 课程思政融入路径 |
| 09 | `evaluation_dashboard` | 多元评价仪表盘 |
| 10 | `course_panorama` | 课程全景图 |

## 安装与使用

### 前置条件

```bash
pip install python-pptx
```

### 在 Claude Code 中使用

将此 skill 放入 `~/.claude/skills/pptx-diagram/` 目录，Claude Code 会自动识别。

也可以在对话中直接对 Claude 说：

> "帮我生成一个三栏管线框图，主题是'课前课中课后三段式教学'"

### 单独使用模板

每个模板都是独立的 Python 脚本，可以直接运行：

```bash
cd templates
python 01_three_column_pipeline.py
# 输出: output.pptx
```

修改模板中的文本内容后重新运行即可生成新图。

## 配色参考

```
BG_OUTER  = #F8FBFE  背景
BG_MAIN   = #F5FAFF  主容器
CARD_LT   = #D5F4FF  卡片浅蓝底
DARK_BLUE = #1A3C5E  标题/强调
MED_BLUE  = #2C5F8A  箭头/标签
LINE      = #D0DEEB  边框
ORANGE    = #F47920  仅用于关键词强调（每页 ≤3 处）

蓝梯度（用于序列/进度）：
BLUE1=#1E5793 → BLUE2=#2E78B5 → BLUE3=#4B9BD4 → BLUE4=#7DC5ED
```

## 设计原则

1. **大浅蓝容器**包裹全部内容
2. **左→中→右**管线布局
3. **纯矩形**为主，不用菱形/六边形等花哨形状
4. **密而不乱** — 130-200 个形状，颜色统一就不显乱
5. **填满画布** — 内容延伸到 y≈7.0（画布高度 7.5"）

## 注意事项

- ❌ 不要用 AI 生图当主体配图 — PNG 不可编辑
- ❌ 不要多色系混用 — 统一到蓝色渐变
- ❌ 不要留大面积空白 — 评委视为内容缺失
- ❌ 不要 `clone_slide()` 抄袭参考页 — 重新写脚本
- ❌ 字号不要低于 6.5pt — 投影看不清

## License

MIT
