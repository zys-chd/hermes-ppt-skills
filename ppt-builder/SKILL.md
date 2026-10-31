---
name: ppt-builder
description: "用 python-pptx 制作专业汇报 PPT 的方法论与代码工具：叙事结构、问题攻关方法（5W2H/人机料法环/5Why/Weibull/8D）、原生图表、多风格设计系统、跨平台字体、质量验收。适用于把数据或分析结论做成真实 .pptx 文件的任务。"
version: 2.0.0
author: Hermes Agent
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [ppt, powerpoint, presentation, python-pptx, data-storytelling, reliability-engineering, native-charts]
---

# PPT Builder v2 — 专业汇报演示文稿制作

用 python-pptx 制作**专业、可编辑、结论清晰**的 PPT。

**本 skill 的定位**：提供方法论 + 一套带强约束的设计系统（`scripts/pptx_helpers.py`）。
不规定每页内容，但**在关键处（图表实现、字体、配色）有硬性约束**——因为这些地方
实测中极易出错，且错误会显著降低交付质量。

---

## ⚠️ 三条铁律（开工前必读）

这三条是实战中反复踩坑总结的，违反任何一条都会导致交付不合格。

### 铁律 1：基础图表必须用原生图表，不用图片

| 要画的东西 | 必须用 | 禁止 |
|-----------|--------|------|
| 柱状图 / 条形图 | `native_bar(...)` | ❌ matplotlib 存图片 |
| 折线图 | `native_line(...)` | ❌ matplotlib 存图片 |
| 饼图 / 环形图 | `native_pie(...)` | ❌ matplotlib 存图片 |
| 散点图 | `native_scatter(...)` | ❌ matplotlib 存图片 |
| 流程图 / 步骤条 | `native_flow` / `native_chevron_flow` | ❌ 图片 |
| 层级图 / 组织图 | `native_hierarchy(...)` | ❌ 图片 |
| 鱼骨图 / 因果图 | `native_fishbone(...)` | ❌ 图片 |
| 热力图 | `make_heatmap(...)` + 图片 | ✅ 只能图片 |
| Weibull 概率图 / 复杂组合图 | matplotlib + 图片 | ✅ 只能图片 |
| 照片 / 物证图 | `picture(...)` | ✅ 本来就不能矢量化 |

**为什么是铁律**：`native_*` 生成 PowerPoint 原生对象，用户能右键「编辑数据」改数字、
换类型、套主题；图片是死的。这是交付质量的分水岭。

**不要这样做**（实测中模型会犯的错）：
- ❌ 自己写 `def save(fig, name): fig.savefig(...)` 存图 —— 绕过了所有约束
- ❌ 只 import 部分函数（如只导入 `native_fishbone`），数据图表仍自己写 matplotlib
- ❌ 认为"用 matplotlib 更熟悉"就走图片路线

**验收（必须执行）**：
```bash
python3 scripts/verify_native_charts.py <你的deck.pptx> --allow-images <物证图数量>
```
它会数 pptx 内部的 chart XML 数量。**原生图表 < 3 个 → 判定不合格，必须重做。**
这个检查看的是最终文件的客观结构，无法通过改代码绕过。

### 为什么验收要看「最终文件结构」而不是代码（四轮失败换来的）

实测中约束模型用 `native_*` 失败了 4 轮，每一轮都以为是措辞问题：

| 手段 | 结果 | 原因 |
|------|------|------|
| SKILL.md 里写清优先级表 | ❌ 无效 | 位置靠后，且 matplotlib 说明篇幅更大，模型先记住后者 |
| 提到文档最顶部的检查点 | ❌ 无效 | 模型仍照自己的习惯写 |
| helper 里加运行时 `warnings.warn` | ❌ 无效 | **模型自己写了 `def save(fig): fig.savefig()`，压根不经过 helper** |
| **事后检查 pptx 内部 chart XML 数量** | ✅ 有效 | 看的是最终文件的客观事实，改代码绕不过 |

**通用教训（适用于任何"要求模型用我的工具库"的场景）**：
- 运行时检查/警告防不住模型 —— 它能重写你的函数、绕过你的 API、直接调底层库。
- **唯一可靠的约束是「对最终产物做客观结构检查」**，因为产物是事实，代码路径可以任意绕。
- 检查脚本要能一眼判定（数量阈值 + 退出码），并写进流程里"不通过不算完成"。

### 更根本的前提：先确认指令源没有冲突

上述四轮排查里，**真正的主因其实在任务 prompt**（写着"用 matplotlib 生成 PNG"），
与 skill 直接矛盾。模型服从 prompt 是正确的。所以在怀疑 skill 措辞之前，
**第一件事是并排读 prompt 与 skill 的原文**（详见第九节"冲突时怎么办"）。
顺序是：先排除指令冲突 → 再上事后结构验收；两者都做到才闭环。

### 铁律 2：字体只用 `CN_FONT`，不自设字体名

- 一律用 `CN_FONT` 常量（已按平台自动检测：macOS→PingFang SC、Windows→微软雅黑、Linux→Noto Sans CJK）
- 或调用 `text()` 时不传 `font` 参数（默认就是 CN_FONT）
- **禁止**硬编码 `"Songti SC"`、`"Hiragino Sans GB"`、`"STHeiti"`、`"Arial"` 等 ——
  前三个是 macOS 专有字体，交付到 Windows 会缺字（Hiragino 更是直接乱码方框，实测出现过 4 个文件全乱码）
- helper 内有 `UNSAFE_FONTS` 黑名单拦截，但不要依赖拦截，自己别写

**跨平台交付场景**（PPT 要发给 Windows 用户看）：
```bash
export PPT_CN_FONT="Microsoft YaHei"   # 强制用微软雅黑，保证 Windows 上观感一致
```
macOS 自用不需要设，自动选 PingFang SC 更美观。

### 铁律 3：配色走 helper，不自写上色逻辑

- 图表配色：`native_*(..., style=st)` 自动应用风格色板，无需手动设
- 需要逐点上色（如高亮某根柱子）：用 `paint_points(图, ["#色1", "#色2"])`
- **禁止**自己写 `point.format.fill.solid()` 循环 —— 实测有模型自写 `paint_points` 有 bug，
  导致图表退化成 PowerPoint 默认蓝色主题，与全篇风格割裂
- **交付前必查配色是否真的落盘**：用户肉眼能发现"有几张图颜色和主题不统一"，而代码看着没问题。
  逐页审计每个 chart 的系列色，任何 `继承主题` / 非色板色都算不合格：

```python
from pptx import Presentation
p = Presentation("deck.pptx")
for i, s in enumerate(p.slides, 1):
    for sh in s.shapes:
        if sh.has_chart:
            for ser in sh.chart.series:
                try:
                    print(i, ser.format.fill.fore_color.rgb)   # 必须都是色板内的色
                except Exception:
                    print(i, "继承主题 ← 不合格")
```

  饼图还要单独查每个扇区（`series[0].points[i].format.fill`），因为系列级上色对饼图无效。
  最可靠的做法是只用 `native_pie(pct_in_label=True)`，它内部已逐扇区上色。

---

## 一、代码工具：`scripts/pptx_helpers.py`

设计理念：**底层能力 + 风格系统 + 强制一致性**。用 `from pptx_helpers import *`。

### 1. 底层能力（自由布局）
`new_deck()` `blank()` `set_bg()` `rect()`（含圆角/边框）`line()` `text()`（富文本）`picture()`

### 2. 风格系统
```python
st = get_style("tech")       # 科技深色
st = get_style("minimal")    # 极简留白
st = get_style("editorial")  # 杂志图文（米白/墨绿/陶土）
st = get_style("warm")       # 暖色商务
st = expand_palette("#2E75B6", "#E08A2E")   # 从用户给的 1~3 个基础色拓展完整色系
hexcolor("#RRGGBB")          # hex → RGBColor
```
每个 style 含：`bg/panel/primary/accent/text/muted/light/white/series`

### 3. 装饰元素
`big_number()` 大数字焦点 · `tag()` 标签徽章 · `flow()` 简单流程条 · `kicker()` 小字引导
`title_bar()` 标准标题（返回正文起始 y）· `footer()` 页脚

### 4. 原生图表（**铁律 1**）
```python
native_bar(slide, L, T, W, H, categories, {"系列": [值...]}, st, title="...")
native_line(...)  native_pie(..., donut=True)  native_scatter(..., {"名": (xs, ys)})
paint_points(图, ["#C77B3F", "#8AA38F"])   # 逐点上色
```
多系列自动应用完整色板；饼图自动逐扇区上色。

### 5. 原生图形（替代 SmartArt，全矢量可编辑）
`native_flow(steps...)` · `native_chevron_flow(...)` · `native_hierarchy(root, children...)` · `native_fishbone(effect, {类别: [原因]})`

### 6. 图片图表（仅限热力图 / Weibull / 复杂组合）
`make_heatmap()` `make_bar()` `make_line()` `make_pie()` `make_scatter()` + `save_chart()`
（注意：`make_bar/line/pie/scatter` 会触发警告提醒你改用 native_*）

### 7. 参考版式（可选，可自定义）
`cover()` `section()` `closing()`

---

## 二、设计风格

四套内置风格，选风格的依据是**场合和读者**：
- 技术评审 / 工程攻关 → `minimal` 或 `editorial`（克制专业）
- 管理层 / 商务汇报 → `warm`
- 有故事性 / 图多 → `editorial`
- 大屏 / 夜间展示 → `tech`

### ⚠️ 反趋同规则

**默认不要选 `tech`**。实测：并行产出 12 份 PPT 时，11 份都选了 tech 深色，
全部长得一样，失去设计价值。模型会趋同到"看起来最酷"的那套。

决策顺序：
1. 用户指定了风格/配色 → 严格照做
2. 技术评审/工程攻关 → `minimal` 或 `editorial`（tech 深色在评审场合反而显得浮夸）
3. 管理层/商务 → `warm`
4. 需要视觉冲击 → `editorial`
5. 只有明确需要"科技感/夜间/大屏"才用 `tech`

**必须在最终报告里用一句话说明"为什么选这套风格"** —— 让风格成为设计决策，不是默认值。

### 反 AI 味清单
- ✗ 每页都是"标题 + 3~5 个要点" —— 用大数字、图表主导、图文混排、流程条交替打破节奏
- ✗ 空话套话（"赋能""闭环""抓手""深度打通"）—— 用具体数字、具体动作
- ✗ 千篇一律的居中标题 + 对称布局 —— 用不对称构图、色块分区、边栏
- ✗ 图表配空洞总结 —— 图注要么讲数字、要么讲机理、要么讲行动
- ✓ 允许"一页一图一大字"的留白页，也允许"密集证据页"，让汇报有呼吸感

### 字号层级（16:9，13.33×7.5 英寸）
页面标题 24–28pt 加粗 · 小节标题 18–20pt · 正文 14–16pt · 图表/脚注 10–12pt
核心数字 36–54pt（`big_number()`），用强调色单独一行

---

## 三、叙事结构

### 金字塔原理
结论先行 → 以上统下 → 归类分组 → 逻辑递进

### 标准汇报结构（12–14 页）
1. 封面（标题 + 副标题 + 汇报人/日期）
2. 执行摘要（3–5 条核心结论）
3. 逐主题分析页（每页一条结论）
4. 建议/行动项页
5. 封底

### 每页骨架
- **标题 = 一句话结论**（"西南区域同比增长 45%，是唯一增长引擎"，不是"区域收入分析"）
- **主体 = 支撑证据**（图表 + 关键数字标注）
- **脚注 = 数据口径/来源**

---

## 四、问题攻关型 PPT（研发/质量/可靠性）

骨架是"问题 → 根因 → 对策"，不是"现象 → 描述"。

### 5W2H 定义问题
What / Where / When / Who / Why / How / How much（量化：失效率、退回率、成本）

### 人机料法环（4M1E）排查
人 / 机 / 料 / 法 / 环 逐项排查。
**警惕单维度假象**：某维度失效率偏高可能是另一维度分布不均造成的假关联
（如"某产线失效率高"实际是该产线生产了更多问题批次料）。必须做**分层交叉分析**。

### 鱼骨图 + 5Why
鱼骨图按人机料法环分类列出可能原因；5Why 从现象连续追问到可行动的根本原因
（某物料批次、某工艺参数、某设计缺陷），不要停在"环境不好""人员疏忽"。

### 可靠性/失效分析
- **Weibull 形状参数 β**：β<1 早期失效（制造/物料缺陷）、β≈1 随机失效、β>1 磨损老化
- 浴盆曲线：早期失效期 → 偶然失效期 → 磨损期
- 失效模式分布（帕累托）区分主要失效机理
- **失效率必须用出货量做分母**，只看失效绝对数会被市场规模误导

### 证据展示
- 物证图（失效件照片、X-ray、波形、切片）与结论对应放置，**每张配一句"它证明了什么"**
- 数据图讲"统计规律"，物证图讲"物理机理"，两者互证锁定根因
- 排除项也要展示（如"X-ray 空洞在规格内，排除焊接根因"），体现排查严谨性

### 8D 结构
D1 小组 → D2 描述问题(5W2H) → D3 临时措施 → D4 根因分析 → D5 永久对策 →
D6 验证 → D7 预防再发 → D8 总结。
PPT 浓缩为：背景 → 问题定义 → 排查分析 → 根因锁定 → 对策 → 验证

---

## 五、图表选型

| 表达什么 | 图表 | 实现 |
|---------|------|------|
| 趋势 | 折线图 | native_line |
| 对比/排序 | 条形图（横向优先） | native_bar(horizontal=True) |
| 构成/占比 | 饼图(≤5类) / 环形 | native_pie |
| 关系/相关 | 散点图 | native_scatter |
| 交叉分析（料×环） | 热力图 | make_heatmap（图片） |
| 失效分布 | Weibull 概率图 | matplotlib（图片） |
| 根因分类 | 鱼骨图 | native_fishbone |

### 反模式
✗ 3D 图表/透视饼图（扭曲比例） ✗ 柱状图 Y 轴不从 0 开始 ✗ 双饼图对比
✗ 一页超 3 个图表 ✗ 无标注裸图

---

## 六、工作流程

1. **读数据** → 2. **分析**（pandas，算趋势/对比/分层/异常）→ 3. **提炼洞察**（3–5 条，"所以呢？"）
→ 4. **设计叙事线**（金字塔排序）→ 5. **列图表清单**（逐个标注类型 + 实现方式）
→ 6. **生成**（用 native_* 画基础图表）→ 7. **验收**（跑 verify_native_charts.py）
→ 8. **视觉自检**（见第七节）→ 9. **交付**（给文件路径）

---

## 七、交付前的两重自检

### 结构化自检（必须，可自动）
```bash
python3 scripts/verify_native_charts.py deck.pptx --allow-images <物证图数>
```
检查：原生图表 ≥3、图片数不超容差、有嵌入数据。

同时用代码独立复核关键数字（重读 CSV 重算，再断言 pptx 里的字符串）——
视觉检查看不出"数字算错"，它只会说"数字清晰可读"。

### 视觉自检（强烈建议）
代码跑通 ≠ 版面正确。至少藏着 1–3 处"数据逻辑错位"或"标签压字"，只有渲染才看得见。

**macOS**（Keynote 最稳）：
```bash
osascript -e 'tell application "Keynote"
  activate
  set d to open POSIX file "/abs/deck.pptx"
  delay 3
  export d to POSIX file "/abs/preview.pdf" as PDF
  close d saving no
end tell'
```

**Windows**（PowerPoint COM）：
```powershell
$pp = New-Object -ComObject PowerPoint.Application
$p = $pp.Presentations.Open("C:\abs\deck.pptx", $true, $false, $false)
$p.SaveAs("C:\abs\preview.pdf", 32)
$p.Close(); $pp.Quit()
```

**Linux / 三平台通用**（LibreOffice）：
```bash
soffice --headless --convert-to pdf --outdir . deck.pptx
```

之后统一转 PNG 并逐页检查：
```bash
python3 -c "import pymupdf; d=pymupdf.open('preview.pdf'); [p.get_pixmap(dpi=110).save(f'preview/s{i+1:02d}.png') for i,p in enumerate(d)]"
```
> Windows 上 python 命令通常是 `python` 而非 `python3`，路径分隔符用 `\\` 或 `/` 均可；
> 若没装 pymupdf：`pip install pymupdf`。

必查：文字溢出/裁切、元素重叠、tofu 方框、图表标签压柱、图例遮挡、页码缺失、**配色与主题是否统一**。

### 平台差异速查

| 项目 | macOS | Windows | Linux |
|------|-------|---------|-------|
| 中文字体 | PingFang SC | Microsoft YaHei | Noto Sans CJK SC |
| 字体覆盖 | 环境变量 `PPT_CN_FONT` 可强制指定 | 同左 | 同左 |
| PPT→PDF | Keynote (osascript) | PowerPoint COM | LibreOffice (soffice) |
| python 命令 | `python3` | `python` | `python3` |

字体由 `_detect_cn_font()` 自动检测，跨平台兜底时**不会**选其他平台专有字体
（Windows 上不会回退成 PingFang，Linux 上不会回退成微软雅黑）。

---

## 八、常见坑（全部来自实战）

**图表**
- 多系列图表务必确认每个系列都上了色 —— 否则退化成 PowerPoint 默认蓝色主题（实测发生过）
- 逐点上色只用 `paint_points()`，不要自己写循环（实测有 bug 版本导致配色失效）
- 饼图必须逐扇区上色，`_apply_chart_style` 只给「系列」上色对饼图不够
- **`legend.position = XL_LEGEND_POSITION.RIGHT` 会让整张图消失（python-pptx 1.0.2 实测）**：
  它写出空的 `<c:legendPos/>`（缺 val 属性），Keynote / PowerPoint 解析不到图例位置就整图不渲染，
  画面上直接空白且代码不报错。修法：设完手动补属性
  `lp = chart.legend._element.find(qn('c:legendPos')); lp.set('val','r')`（b/t/l/r 四向）。
  凡是设了 RIGHT 图例的图，生成后必须渲染确认图还在。
- **单系列柱图配色**：默认 `varyColors=1` 时 PowerPoint 用调色板逐点着色、系列填充被忽略；
  `paint_points()` 内部已自动设 `vary_by_categories = False`，自己手写设色时也要记得关。
- **折线图颜色要用 `series.format.line.color.rgb`**，`format.fill` 对折线无效（仍显示主题蓝）。
  `native_line()` 已同时设 fill 和 line，无需自己处理。
- **饼图百分比标签不可靠**：部分渲染器忽略 `showPercent`，把 `numFmt` 套到原始值上（显示「463.0%」）。
  稳妥做法：`native_pie(..., pct_in_label=True)` —— 它把百分比写进类别名（`"无输出 48%"`）、
  关掉数据标签，图例即标签，任何渲染器都对。
- **浮点 EMU 污染**：`h * 0.5`、`Inches(x) * i` 会算出 float，写进 XML 成 `cy="123.0"`，
  python-pptx 自己重读时抛 `ValueError: invalid literal for int(): '0.0'`。保存前统一 int 化：
  遍历 `slide.shapes._spTree` 中 `<a:off>/<a:ext>/<a:chOff>/<a:chExt>` 的属性 `int(round(float(v)))`。
- **形状投影**：`rect()` 生成的形状继承主题 `effectRef`，Keynote 渲染出一圈灰投影，显得脏。
  交付前去投影：`sh.shadow.inherit = False` **并** 删除该 `p:sp` 下的 `<p:style>` 节点
  （只设 shadow.inherit 不够，effectRef 仍会生效）。

**布局**
- **单位混算致命坑**：`Inches(img_w - Inches(0.2))` 里 img_w 是 float、Inches(0.2) 是 EMU，
  相减得负数 → 文本框宽度为负，PowerPoint 可能提示修复。统一传英寸数值：`Inches(img_w - 0.2)`
- 大字号装饰文字必须 `text_frame.word_wrap = False`，否则被折行（150pt 的 "2024" 会变成 "202"+ "4"）
- 文本框垂直居中用 `text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE`

**数据**
- 宽表分页粒度陷阱：同一 region-month 有 4 行品类数据时，`groupby("year").store_count.sum()`
  会把门店数放大 4 倍。凡"每店"类指标先按最细粒度去重再聚合
- 瀑布图浮动柱：`base`/`height` 必须同号语义，向下柱用 `base=起点, height=负增量`

**字体**
- 见铁律 2。Hiragino Sans GB / STHeiti 在 Windows 缺字 → 乱码

---

## 九、与其他 agent 协作时的注意事项

本 skill 的方法论与 `pptx_helpers.py` 均为纯文本 + 纯 Python（仅依赖 python-pptx /
matplotlib / numpy / pandas / scipy），可移植到任何遵循 SKILL.md 约定的 agent
（Claude Code、opencode 等）。移植时注意：
- 把 `CN_FONT` 的平台检测保留（跨平台关键）
- `verify_native_charts.py` 要一并带上（否则模型可能绕过原生图表要求）
- 若宿主 agent 的任务描述与本文冲突，**以任务描述为准**，但要主动指出冲突让用户知晓

### 冲突时怎么办（血泪教训）

本 skill 的三条铁律**只在任务没有明确指定时生效**。如果调用方的任务 prompt 里写了相反的
要求（实测遇到过：prompt 里写着"用 matplotlib 生成 PNG 后插入"，直接推翻铁律 1），模型会
服从 prompt，这是正确行为——但你作为 agent 应当主动指出这个冲突，而不是默默照做或反复改
skill 试图压过它。

**改 skill 压不过 prompt。** 如果你发现产出反复不符合 skill 要求，而 skill 已经写清楚了，
先去**并排读 prompt 和 skill 的原文**确认有无反向指令（排查方法见 `systematic-debugging`
的"检查更高优先级的指令源"一节）。改了三轮 skill 都没用，问题往往根本不在 skill 里。

同一份 prompt 若有多处副本（本机 + 仓库公开版 + 模板），改一处必须同步全部。
