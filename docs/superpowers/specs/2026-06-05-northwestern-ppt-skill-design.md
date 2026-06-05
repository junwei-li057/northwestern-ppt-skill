# Northwestern PPT Agent — Skill 设计文档

- 日期：2026-06-05
- 状态：已与用户确认设计，待写实现计划
- 目标：一个 Claude Code skill，输入一篇论文（PDF/文档），产出一份 **Northwestern 风格** 的 PowerPoint 演示文稿（.pptx），风格以本项目 `references/` 内的三份范例为准。

## 1. 目标与范围

### 输入 / 输出
- 输入：一篇论文 / PDF / 文档。
- 输出：一份 `.pptx`，白底 + Northwestern 紫色品牌、Times New Roman、研究汇报风格。

### 工作流程（先大纲、后生成）
1. agent 读论文，**先确认目标时长**，据此规划页数。
2. agent 依据 `STYLE.md` 的分页/选型规则，产出**大纲 JSON**；展示给用户时渲染成**易读的 markdown 大纲**。
3. 暂停 → 用户确认 / 修改大纲。
4. 确认后由 `build_pptx.py` 渲染成 `.pptx`，并自动渲染 PNG 供视觉自检。

### 明确不做（YAGNI）
- 不处理中文论文（西文场景）。
- 不追求与某一份范例像素级一致；取其精华，整体稳定但页与页之间保留差异。
- 不自由发挥配色/字体——视觉身份全部来自模板与集中样式常量。

## 2. 技术方案

**方案 A（已选）：官方母版模板 + python-pptx 组件渲染器。**

- 视觉品牌（白底、紫色页脚条、左侧紫条、"Northwestern" 衬线字标、16:9 画布、主题）来自官方模板 `assets/northwestern_template.pptx`。
- 在品牌画布之上，脚本用 python-pptx **程序化绘制**从范例提炼的组件，颜色/字号/间距来自集中样式常量。
- agent 只决定"内容如何分页、每页用哪种正文类型"，不接触任何颜色值。

已否决：
- 方案 B（纯代码画样式、无模板）——颜色字体散落、易出 AI 味、难还原范例。
- 方案 C（HTML/CSS → pptx 转换）——产物非原生、难编辑、图表处理别扭。

## 3. 文件结构

```
northwestern-ppt/
├── SKILL.md                  # 入口：触发条件 + 工作流程
├── STYLE.md                  # 风格说明书：调色板、字体、组件目录、分页/选型规则
├── assets/
│   └── northwestern_template.pptx   # 官方母版（已就位）
├── scripts/
│   ├── extract_pdf.py        # PDF → 正文文本 + 图表图片 + figures.json
│   ├── build_pptx.py         # 大纲 JSON + 模板 → .pptx
│   └── render_preview.py     # .pptx → 逐页 PNG（视觉自检）
└── references/               # 三份范例（只读，风格来源）
```

最终可安装/软链到 `~/.claude/skills/`。

## 4. 风格系统（来自范例提炼）

视觉风格分两层：**稳定的框**（每页一致）+ **正文处理菜单**（按页内容选，制造差异）。

### 4.1 稳定的框（三份范例的共性）
- Northwestern 品牌画布（模板自带）。
- 顶部**标题**；右下**页码**。
- 统一字体 **Times New Roman**；统一紫色系调色板。
- 专门的**标题页**（标题 + 副标题 + 来源行）与**致谢页**（Thanks!）。

### 4.2 调色板（从 references 实测）
- 主紫 `#542A84`（≈ Northwestern 校紫 4E2A84）、深紫 `#3A1C5C`。
- 浅紫底 `#F6F2FA`、`#E6E1EC`。
- 强调色轮换：蓝 `#4574B8`、绿 `#3D916F`、金 `#CE973E`、珊瑚 `#CF4F61`。
- 卡片浅填充示例：蓝 `#EFF4FC`、绿 `#EEF8F3`、金 `#FCF7ED`（与描边同色系）。
- 正文深灰 `#262626`。

### 4.3 字号基线（实测 part4）
- 眉标 eyebrow：10pt，全大写。
- 标题 title：28pt。
- 卡片标题：15pt；卡片描述：12pt。
- 标题下紫色下划线条：填充 `#542A84`，约 1.15×0.04 in。

### 4.4 正文处理菜单（组件目录）
| 正文类型 | 范例来源 | 用途 / 规格要点 |
|---|---|---|
| **Bullets**（默认、最常用） | hw3 / Bayesian 主力 | 普通论述、罗列；要点 ≤ 5 条，每条 ≤ 2 行 |
| **Card row** | part4 | 2–3 个或网格并列概念/指标；圆角矩形=浅填充+同色描边；标题 15pt+描述 12pt；配色轮换 |
| **Pill flow** | part4 / hw3 / Bayesian | 步骤/管线；小圆角药丸 + 连接线/箭头 |
| **Callout** | part4 "Core claim" / Bayesian "Key Takeaways" | 核心主张 / 一句话总结；整宽圆角框 + 标签 + 一句话 |
| **Mapping / Table** | part4 映射 / hw3 性能表 | 左右对应、结果数字对比 |
| **Figure block** | 三份都有 | 论文图表；大图 + 图注/侧栏卡片 |

可选点缀：part4 的**眉标 + 紫下划线**作为页眉变体，成套启用或成套不用（保持一致），默认对学术汇报启用。

## 5. 分页逻辑 + 正文选型（写入 STYLE.md，agent 出大纲时遵循）

### 5.1 篇幅 = 时长驱动（运行时确认）
- 出大纲前**先问目标时长（分钟）**。
- 换算：**页数 ≈ 时长(min) × 0.7**（约 1.2–1.4 min/页），再按内容微调。
  - 例：10 min → ~8 页；15 min → ~11 页；20 min → ~14 页。
- 反推依据：范例 13 页 ≈ 18 min、8 页 ≈ 9 min。
- 骨架按总页数弹性伸缩：标题 → 背景/问题 → 方法(2–4) → 关键结果(2–3) → 讨论/局限 → 要点总结 → 致谢；时长短压缩方法/结果，时长长则展开。

### 5.2 分页规则
- **一页一论点**：一页只讲一件事。
- 标题**精炼贴题**即可：陈述句（"X 提升了 Y"）、短语或问句都可以，怎么合适怎么来；但**不要机械照搬论文章节名**。
- 论文图/表优先**单独成图块页** + 一句话解读，不塞进要点。

### 5.3 正文类型选择规则（制造差异、避免千篇一律）
- 默认 **Bullets**；但**不连续超过 2 页**用 bullets，第 3 页起强制换结构化处理。
- 2–3 个并列概念/指标 → **Card row**。
- 有顺序的步骤/管线 → **Pill flow**。
- 核心主张 / 一句话总结 → **Callout**。
- 左右对应 / 数字对比 → **Mapping / Table**。
- 有论文配图 → **Figure block**。

### 5.4 大纲数据格式
- 大纲为 **JSON**，每页含：`type`、`eyebrow`、`title`、`body`（按类型而定）、`figure`（图号引用）。
- 向用户展示时渲染成 **markdown 大纲**（易读），确认后再构建。

## 6. 图表抽取 + 构建管线

### 6.1 extract_pdf.py（论文 → 素材）
- 用 **PyMuPDF (fitz)**：
  - 抽**正文文本**（按页/节）供 agent 阅读总结。
  - 抽**图表图片**：① 导出嵌入位图（`page.get_images`）；② 矢量图/图表按图注（"Figure N"/"Table N"）定位区域并 `page.get_pixmap(clip=...)` 渲染为 PNG。
  - 产出 `work/figures/figN.png` + `figures.json`（图号、所在页、图注文字）。
- 兜底：抽不到/不确定的图标记"需人工确认"，不硬塞。

### 6.2 build_pptx.py（大纲 + 模板 → .pptx）
- 打开模板，**删掉自带 3 张示例页**。
- 遍历大纲，每页按 `type` 调对应渲染函数：`render_bullets / render_card_row / render_pill_flow / render_callout / render_mapping / render_figure / render_title / render_thanks`。
- 所有颜色/字号/间距来自集中 `style` 常量表。
- 图片等比缩放、不超边界。
- 输出 `<paper_name>_northwestern.pptx`。

## 7. 错误处理 + 视觉验证

### 7.1 错误处理
- 启动检查 `python-pptx`、`PyMuPDF`、LibreOffice；缺失则由 agent `pip install` / 安装。
- 单页渲染失败 → 跳过该页 + 末尾追加"构建警告"页，不让整份崩。
- 文字超框 → 自动缩字号到下限；仍超则截断 + 警告。
- 素材抽取与构建分离，改大纲后只需重跑 `build_pptx.py`。

### 7.2 视觉验证（闭环自检）
- 安装 **LibreOffice**（`soffice --headless`），构建后 `.pptx → PDF → 逐页 PNG`（pdftoppm 或 PyMuPDF）。
- **agent 用 Read 打开 PNG 视觉核对**：溢出、重叠、留白、配色是否合理；发现问题就调大纲/参数重跑。
- 验证不只看文字抽取是否正确，更要看**整体视觉**是否成立。
- 提供 `make_sample.py`：用内置假大纲生成"组件样张"，单测每种正文类型能正常画出。

## 8. 开放项
- 各组件的精确坐标/间距在实现时以范例为参照标定（spec 已给出关键基线值）。
- 图注区域定位（矢量图）的鲁棒性需在真实 PDF 上调参。
