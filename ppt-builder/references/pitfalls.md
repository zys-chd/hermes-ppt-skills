# python-pptx 实战踩坑

> 生成 PPT 遇到问题时查阅。

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
