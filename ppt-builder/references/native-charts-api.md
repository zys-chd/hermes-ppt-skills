# 原生图表与原生形状 API 速查

`pptx_helpers.py` 中 `native_*` 系列函数的签名与用法。这些函数生成 PowerPoint 原生对象
（可右键编辑数据），是基础图表和图形化元素的**默认选择**。

## 原生图表（数据可在 PPT 内编辑）

所有函数返回 GraphicFrame，内部是真正的 chart 对象（含嵌入 Excel 数据表）。

```python
from pptx_helpers import *

st = get_style("minimal")   # 或 expand_palette("#2E75B6", "#E08A2E")

# 柱状图 / 条形图
native_bar(slide, left, top, width, height,
           categories=["A-2411", "A-2410", "B-2410"],
           series={"失效率%": [3.20, 0.60, 0.61]},   # dict 或 [(名, 值列表)]
           style=st, title="各批次失效率",
           horizontal=False,        # True 为横向条形图
           data_labels=True,
           num_format='0.00',
           colors=None)             # 可覆盖系列色

# 折线图
native_line(slide, left, top, width, height,
            categories=["1月", "2月", "3月"],
            series={"A-2411": [100, 120, 90], "A-2410": [100, 105, 110]},
            style=st, title="失效趋势", markers=True)

# 饼图 / 环形图
native_pie(slide, left, top, width, height,
           categories=["无输出", "间歇重启", "过热保护", "其他"],
           values=[48, 35, 11, 6],
           style=st, title="失效模式分布", donut=True)

# 散点图（注意 series 的值是 (xs, ys) 元组）
native_scatter(slide, left, top, width, height,
               series={"样本": ([1, 2, 3], [1.1, 1.8, 3.2])},
               style=st, title="相关性")
```

要点：
- `series` 传 dict 或 `[(名, 值列表), ...]` 都可以。
- 单系列时可省略 `show_legend`，多系列自动显示图例。
- 配色自动取 `style["series"]`，可用 `colors=` 覆盖。
- `num_format` 控制数据标签格式（`'0.00'` 两位小数、`'0.0"%"'` 带百分号）。

## 原生形状（替代 SmartArt）

python-pptx 无 SmartArt API（Microsoft 私有格式），以下函数用原生矢量形状实现同等效果，
且可精确控制位置/颜色/字体。

```python
# 横向流程条：方块 + 箭头
native_flow(slide, ["问题定义", "排查分析", "根因锁定", "对策验证"],
            left, top, width, style=st, numbered=True)

# V 形箭头流程（更紧凑、更有方向感）
native_chevron_flow(slide, ["现象", "分层", "交叉", "根因", "对策"],
                    left, top, width, style=st)

# 层级图 / 组织架构图：根节点 + 子节点（自动画连接线）
native_hierarchy(slide, root="早期失效",
                 children=["料：A-2411批次", "环：北方低温", "机：无显著", "法：工艺正常"],
                 left, top, width, style=st)

# 鱼骨图 / 因果图：主脊 + 斜支骨 + 类别标签 + 原因小字（上下交替）
native_fishbone(slide,
                effect="早期失效",
                causes={"料": ["A-2411批次", "ESR异常"],
                        "环": ["北方低温", "加速劣化"],
                        "机": ["产线无差异"],
                        "法": ["工艺正常"]},
                left, top, width, height, style=st)
```

## 组合使用建议

研发攻关 PPT 的典型页面配方：

| 页面 | 组合 |
|------|------|
| 问题定义 | `native_hierarchy`（5W2H 分层）或 `big_number` + 表格 |
| 排查页 | `native_fishbone`（人机料法环）+ 侧边 `tag()` 标注"排除/异常" |
| 分层失效率 | `native_bar`（多系列对比） |
| 交叉分析 | `make_heatmap`（图片，原生无热力图） |
| 失效性质 | matplotlib Weibull 概率图（图片） |
| 证据墙 | `picture()` × 4 物证照 + 每张配图注 |
| 对策页 | `native_flow` 或 `native_chevron_flow`（临时/永久/验证/预防） |

## 验收脚本

**直接跑 skill 自带的脚本，不要手抄**（会抄错、会漏检查）：

```bash
python3 scripts/verify_native_charts.py deck.pptx            # 默认容差 2 张非物证图
python3 scripts/verify_native_charts.py deck.pptx --allow-images 4   # 物证图更多时放宽
```

退出码 0 = 通过，1 = 未通过（会在 stdout 打印具体原因和修正指引）。
判定口径：原生图表 < 3 个 → 失败；图片数 > 容差 → 失败；有 chart 无嵌入 xlsx → 失败。

下面的内联版仅用于理解判定逻辑（脚本内部就是这套）：

```python
import zipfile
from pptx import Presentation

def check_native(deck_path, expected_photos=4):
    prs = Presentation(deck_path)
    with zipfile.ZipFile(deck_path) as z:
        n_charts = len([n for n in z.namelist()
                        if "ppt/charts/chart" in n and n.endswith(".xml")])
    pics = sum(1 for s in prs.slides for sh in s.shapes if sh.shape_type == 13)
    print(f"页数 {len(prs.slides)} | 原生图表 {n_charts} | 图片 {pics}（物证 {expected_photos}）")
    assert n_charts >= 3, "基础图表必须用原生图表"
    assert pics - expected_photos <= 2, f"多出 {pics-expected_photos} 张非物证图，应改 native_*"
    print("✓ 原生性检查通过")
```
