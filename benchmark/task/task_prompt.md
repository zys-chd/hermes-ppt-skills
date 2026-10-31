# 固定评测任务（研发问题攻关场景）

> **路径说明**：文中 `data/`、`scripts/` 均相对于 benchmark 根目录。运行前请把 `__OUTPUT_PPTX__` 替换为实际输出路径。

> 本文件是发给每个模型的**完全相同**的任务 prompt。不得为不同模型修改任何内容。

---

你是资深可靠性工程师兼质量攻关报告专家。请基于给定的失效数据与物证，制作一份面向公司技术评审会的问题攻关汇报 PPT（8D 报告风格）。

## 背景

星云 X1 智能音箱的电源模块在量产上市后出现市场早期失效，退回率超标，需要定位根因并给出优化方案。

## 数据文件

1. `data/failure_records.csv` —— 失效记录明细
   字段：`record_id`、`serial_no`、`production_date`(生产日期)、`failure_date`(失效日期)、
   `time_to_failure_days`(失效天数)、`supplier`(供应商 A/B)、`batch`(批次)、
   `line`(产线 Line1/2/3)、`region`(销售地区 北方/南方)、
   `failure_mode`(失效模式：无输出 / 间歇性重启 / 过热保护 / 其他)
2. `data/shipment_summary.csv` —— 各供应商×批次×地区×产线的出货量
   字段：`supplier`、`batch`、`region`、`line`、`shipped_qty`(出货量)

## 物证图片（需有机嵌入到对应分析章节）

`data/evidence/` 目录下有 4 张物证图：
- `evidence_capacitor_bulge.png`：失效件电解电容顶部鼓包、漏液
- `evidence_xray_void.png`：BGA 焊点 X-ray 空洞
- `evidence_waveform.png`：示波器输出波形对比（正常 vs 异常纹波）
- `evidence_board_corrosion.png`：板卡电解液腐蚀痕迹

## 制作要求

1. **加载并使用 `ppt-builder` skill** 指导制作（含 5W2H、人机料法环、鱼骨图、5Why、Weibull、8D 等方法论）。
2. 代码工具 `pptx_helpers.py` 位于 `scripts/pptx_helpers.py`，请阅读并优先使用它（`sys.path.insert(0, "scripts")` 后 `from pptx_helpers import *`）。
3. 按以下脉络制作：
   **问题定义(5W2H) → 排查(人机料法环/鱼骨图) → 数据分析 → 根因锁定(5Why) → 证据展示 → 对策与验证**
4. 数据分析务必做到：
   - 用 Weibull 分布分析失效时间，判断是否属于早期失效；
   - 用 shipment 做分母计算**失效率**，不要只看失效绝对数（北方市场更大，绝对数会误导）；
   - 做**分层交叉分析**（如料×环、料×机）定位真实根因，警惕单维度假象。
5. 物证图片要嵌入"证据展示 / 根因分析"章节，**每张配一句图注，说明它证明了什么**。
6. 制作约 **14 页**的 PPT，输出为真实 `.pptx` 文件，保存到：
   `__OUTPUT_PPTX__`
   （图表 PNG、中间脚本等临时文件也保存在 `__OUTPUT_PPTX__` 所在的同一目录下，不要写入共享目录。）
7. **硬性要求**：
   - 至少 **4 类数据图表**：**优先用原生图表**（skill 中的 `native_bar`/`native_line`/`native_pie`/`native_scatter`，数据可在 PPT 内编辑）；原生表达不了的（热力图、Weibull 概率图）用 matplotlib 图片
   - 必须包含：封面、问题定义页、根因分析页、证据展示页、对策与验证页、封底
   - 每页标题必须是一句话结论（例如"根因是 A-2411 批次电容在低温下的早期失效"，而不是"根因分析"）
   - 配色、字体、版式全篇统一
8. **图表与数字必须基于真实数据计算，确保准确**。
9. 完成后，对照 skill 中的"质量检查清单"自查，并最终输出生成的 `.pptx` 文件路径。

请开始。
