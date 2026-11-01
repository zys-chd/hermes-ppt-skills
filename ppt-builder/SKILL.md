---
name: ppt-builder
description: "用 python-pptx 制作专业汇报 PPT：原生可编辑图表、多风格设计系统、跨平台字体、问题攻关方法论（5W2H/4M1E/5Why/Weibull/8D）、交付验收。适用于把数据或分析结论做成真实 .pptx 文件的任务。"
version: 2.1.0
author: Hermes Agent
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [ppt, powerpoint, presentation, python-pptx, native-charts, data-storytelling]
---

# PPT Builder — 专业汇报演示文稿制作

方法论 + 带强约束的设计系统（`scripts/pptx_helpers.py`）。不规定每页内容，但**图表实现、字体、配色有硬约束**——这些地方最易出错。

---

## ⚠️ 三条铁律

### 铁律 1：基础图表必须用原生图表，不用图片

| 要画什么 | 必须用 | 
|---------|--------|
| 柱状图 / 条形图 | `native_bar(...)` |
| 折线图 | `native_line(...)` |
| 饼图 / 环形图 | `native_pie(...)` |
| 散点图 | `native_scatter(...)` |
| 流程 / 层级 / 鱼骨图 | `native_flow` / `native_hierarchy` / `native_fishbone` |
| 热力图 | `make_heatmap(...)` + 图片 ✅ |
| Weibull 概率图 / 复杂组合 | matplotlib + 图片 ✅ |
| 照片 / 物证图 | `picture(...)` ✅ |

**为什么**：原生图表用户能右键「编辑数据」，图片是死的。这是交付质量分水岭。

**不要**：自己写 `fig.savefig()` 存图、只 import 部分函数后数据图表仍手搓 matplotlib。

**验收（必须跑）**：
```bash
python3 scripts/verify_native_charts.py deck.pptx --allow-images <物证图数>
```
原生图表 < 3 个 → 不合格重做。这个检查看 pptx 内部 chart XML 数量，绕不过去。

### 铁律 2：字体只用 `CN_FONT`，不自设字体名

- 用 `CN_FONT` 常量（自动检测：macOS→PingFang SC、Windows→微软雅黑、Linux→Noto Sans CJK）
- **禁止**硬编码 `"Songti SC"` / `"Hiragino Sans GB"` / `"STHeiti"` —— 前三个是 macOS 专有，
  交付到 Windows 会缺字（Hiragino 更是直接乱码方框）
- 跨平台交付时：`export PPT_CN_FONT="Microsoft YaHei"`

### 铁律 3：配色走 helper，不自写上色逻辑

- 图表配色：`native_*(..., style=st)` 自动应用色板
- 逐点上色：`paint_points(图, ["#色1","#色2"])`
- **禁止**自己写 `point.format.fill.solid()` 循环（实测有 bug 版本导致配色退化成主题蓝）

---

## 一、代码工具 `scripts/pptx_helpers.py`

```python
st = get_style("tech|minimal|editorial|warm")   # 4 套风格
st = expand_palette("#2E75B6", "#E08A2E")       # 用户给基础色 → 拓展完整色系
```

**能力分层**：
- 底层：`new_deck()` `blank()` `set_bg()` `rect()` `line()` `text()` `picture()`
- 装饰：`big_number()` 大数字 `tag()` 徽章 `title_bar()` `footer()`
- 原生图表：`native_bar/line/pie/scatter` + `paint_points()`
- 原生图形：`native_flow/chevron_flow/hierarchy/fishbone`
- 图片图表：`make_heatmap/bar/line/pie/scatter` + `save_chart()`
- 版式：`cover()` `section()` `closing()`

> 完整 API 签名见 `references/native-charts-api.md`。

---

## 二、设计风格

四套风格按场合选：技术评审→`minimal`/`editorial`；管理层→`warm`；大屏→`tech`。

**⚠️ 反趋同**：默认不要选 `tech`。实测并行产出 12 份时 11 份都选 tech，全长一样。
技术评审用 `minimal` 或 `editorial` 更显克制。**必须在报告里说明为什么选这套风格。**

**反 AI 味**：不要每页都"标题+3个要点"（用大数字/图表/图文混排打破）；不用空话套话
（"赋能""闭环""抓手"）；不万物居中对称。

---

## 三、叙事与图表

**金字塔**：结论先行 → 以上统下 → 归类分组 → 逻辑递进

**标准结构**：封面 → 执行摘要(3-5条结论) → 逐主题分析(每页一条结论) → 建议页 → 封底

**每页骨架**：标题=一句话结论（不是"区域收入分析"）；主体=图表+数字标注；脚注=数据来源

**图表选型**：趋势→折线；对比→条形(横向优先)；构成≤5类→饼图；相关→散点；交叉→热力图

**反模式**：3D图/透视饼图、Y轴不从0开始、双饼图对比、一页超3图、无标注裸图

---

## 四、问题攻关型 PPT

做研发/质量/可靠性汇报时，加载 `problem-diagnosis` skill 先做排查分析，再转 PPT。

要点：5W2H 界定 → 4M1E/鱼骨图排查（警惕单维度假象，需交叉验证）→ Weibull 判失效性质
（β<1 早期失效）→ 失效率用出货量做分母 → 8D 结构收口 → 物证图每张配"证明了什么"。

> 方法细节见 `references/engineering-methods.md`。

---

## 五、工作流程

读数据 → 分析(pandas) → 提炼 3-5 条洞察 → 设计叙事线 → **列图表清单(逐个标注类型+实现方式)**
→ 生成(用 native_*) → **跑 verify_native_charts.py** → 视觉自检 → 交付

---

## 六、交付前自检

1. **结构化（必须）**：`verify_native_charts.py` + 代码重算关键数字并断言（视觉看不出算错）
2. **视觉（建议）**：渲染成图逐页看，查文字溢出/元素重叠/tofu方框/标签压柱/配色统一

> 三平台渲染链路（Keynote / PowerPoint COM / LibreOffice）和平台差异见
> `references/visual-qa.md`。

---

## 七、遇到问题查 references

| 文件 | 什么时候看 |
|------|-----------|
| `references/native-charts-api.md` | 不确定函数签名/参数时 |
| `references/pitfalls.md` | 遇到图表不渲染、颜色不对、尺寸异常等 |
| `references/visual-qa.md` | 交付前逐页检查，或需要跨平台渲染方案 |
| `references/engineering-methods.md` | 做研发/质量攻关汇报 |
| `references/requirements-brief.md` | 用户需求模糊，需要先访谈清楚 |

---

## 八、与其他 agent 协作

方法论与 helper 均为纯 Python（依赖 python-pptx/matplotlib/numpy/pandas/scipy），可移植到
Claude Code、opencode 等。移植时带上 `verify_native_charts.py`，否则模型可能绕过原生图表要求。

**若任务描述与本文冲突，以任务描述为准** —— 但要主动指出冲突让用户知晓。
（实测教训：prompt 里写"用 matplotlib 生成 PNG"曾直接推翻铁律 1，改三轮 skill 都没用——
问题根本不在 skill 里。改 skill 前先并排读 prompt 和 skill 原文。）
