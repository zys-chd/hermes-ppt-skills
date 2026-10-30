# Hermes PPT Skills

两个用于制作专业 PPT 的 Hermes Agent skill，配合使用实现「需求访谈 → 生成 PPT」的完整流程。

## Skill 列表

### ppt-builder — 生成 PPT
用 python-pptx 制作专业数据分析 / 研发攻关汇报 PPT 的完整方法论 + 可复用代码工具。

- 叙事结构（金字塔 / SCQA / 单页一观点）
- 研发攻关专用方法（5W2H / 人机料法环 / 鱼骨图 / 5Why / Weibull / 8D）
- 图表选型指南 + 反 AI 味清单
- 4 套内置风格（科技深色 / 极简留白 / 杂志图文 / 暖色商务）
- `expand_palette()` 从 1~3 个基础色拓展出完整自定义配色
- `scripts/pptx_helpers.py` 设计工具库（底层能力 + 风格 + 装饰元素 + 图表）

### ppt-brief — 需求访谈
做 PPT 前的需求访谈器，把模糊需求完善成结构化 brief 再交给 ppt-builder。

- 六类问题清单（主题 / 受众 / 风格配色 / 数据 / 篇幅 / 语言）
- 分批提问原则（每次 3~5 个，不重复问）
- 产出结构化 brief 模板

## 安装

```bash
mkdir -p ~/.hermes/skills/productivity
cp -r ppt-builder ~/.hermes/skills/productivity/
cp -r ppt-brief  ~/.hermes/skills/productivity/
```

## 依赖

```bash
pip install python-pptx matplotlib pandas numpy scipy
```

## 使用流程

1. 加载 `ppt-brief` 访谈需求 → 产出结构化 brief
2. 加载 `ppt-builder`，根据 brief 生成 PPT

## 兼容性

- 内容（方法论 + `pptx_helpers.py`）为纯文本 + 纯 Python，只依赖 python-pptx / matplotlib / pandas / numpy，不依赖 Hermes 私有 API，可移植到 Claude Code 等遵循 SKILL.md 约定的 agent。
- macOS / Windows / Linux 均可用，中文字体按平台自动检测（也可用环境变量 `PPT_CN_FONT` 覆盖）。
