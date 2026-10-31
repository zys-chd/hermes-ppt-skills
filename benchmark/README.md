# 研发问题攻关 PPT 生成评测（Benchmark）

用同一份任务 prompt + 同一份数据，横向对比不同大模型在「基础 agent + 优质 skill」条件下制作
研发问题攻关汇报 PPT 的能力差异。

## 场景

星云 X1 智能音箱电源模块量产上市后出现市场早期失效，退回率超标（8D 报告风格攻关汇报）。

**数据全部为虚构生成，不含任何真实企业或个人信息。**

## 目录结构

```
benchmark/
├── task/task_prompt.md          # 固定任务 prompt（所有模型收到完全相同的一份）
├── data/
│   ├── failure_records.csv      # 失效记录明细（1330 条）
│   ├── shipment_summary.csv     # 出货量表（失效率分母）
│   ├── evidence/                # 4 张物证图片（电容鼓包/X-ray/波形/板卡腐蚀）
│   └── insights_key.md          # 埋点根因参考答案（评测者用，勿发给模型）
├── rubric/scoring_rubric.md     # 人工盲评评分表 + 流程
├── scripts/
│   ├── data_gen_failure.py      # 失效数据集生成脚本（seed 固定，可复现）
│   ├── gen_evidence.py          # 物证图片生成脚本
│   ├── run_bench.py             # 并行运行脚本
│   └── models.example.txt       # 模型清单模板
└── results/                     # 各模型产出（可选）
```

## 案例设计：埋点与区分度

数据里埋了 6 个可被分析发现的线索，构成一条完整根因链：

| # | 线索 | 正确解读 |
|---|------|---------|
| R1 | Weibull 形状参数 β≈0.93（A-2411） | 早期失效，非随机/磨损 |
| R2 | A-2411 批次失效率 3.2% vs 其他 0.5% | 料是主因（6 倍） |
| R3 | A-2411 北方 4.38% vs 南方 1.42%；其他批次南北无差异 | **料×环交互 —— 真实根因** |
| R4 | Line2 失效率表面偏高，控制批次后消失 | 机的假象（料分布不均） |
| R5 | A-2411 失效以「无输出/间歇重启」为主（83%） | 电容失效指纹 |
| R6 | 北方整体失效率高 | 环是干扰项，不是根因 |

**核心考点**：能否识破「环（低温）」和「机（产线）」两个假象，定位到「料×环交互」。

## 方法要求

任务要求模型运用：5W2H 定义问题、人机料法环（4M1E）排查、鱼骨图、5Why、Weibull 失效分析、
8D 报告结构，并用 shipment 做分母算失效率（而非失效绝对数）、做分层交叉分析。

## 运行

1. 复制模型清单并填写（文件名与具体模型按需修改）：
   ```bash
   cp scripts/models.example.txt scripts/models.txt
   ```
2. 配置各 provider 的 API key 环境变量（DeepSeek: `DEEPSEEK_API_KEY`；
   阿里云百炼: `DASHSCOPE_API_KEY` + `DASHSCOPE_BASE_URL`）。
3. 运行：
   ```bash
   RUNS=3 JOBS=5 python3 scripts/run_bench.py
   ```
4. 产出匿名化为 `P01.pptx …`，映射表在 `results/_mapping.tsv`，按 rubric 盲评。

## 评测维度（详见 rubric）

根因定位 25% / 逻辑分析 20% / 数据可视化 20% / 证据展示与设计 20% / 完整性 15%

## 依赖

```bash
pip install python-pptx matplotlib pandas numpy scipy
```
