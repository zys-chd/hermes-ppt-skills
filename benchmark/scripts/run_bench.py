#!/usr/bin/env python3
"""并行跑 PPT 生成评测：N 款模型 × R 次，独立工作目录，并发执行。

用法（在 benchmark 根目录下）:
    RUNS=3 JOBS=5 python3 scripts/run_bench.py

环境变量:
    RUNS  每个模型跑几次（默认 3）
    JOBS  并发任务数（默认 5）
    HERMES_BIN  hermes 可执行文件路径（默认从 PATH 查找）

API key 从当前环境变量读取（DeepSeek: DEEPSEEK_API_KEY；
阿里云百炼: DASHSCOPE_API_KEY + DASHSCOPE_BASE_URL）。本脚本不读取任何本地凭据文件。
"""
import os
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
N = int(os.environ.get("RUNS", "3"))
JOBS = int(os.environ.get("JOBS", "5"))
HERMES = os.environ.get("HERMES_BIN", "hermes")

# 1. 读模型清单
models_path = os.path.join(BASE, "scripts", "models.txt")
if not os.path.exists(models_path):
    sys.exit("缺少 scripts/models.txt —— 请先复制 scripts/models.example.txt 为 models.txt 并填写模型")

models = []
with open(models_path, encoding="utf-8") as f:
    for line in f:
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        model, provider = [x.strip() for x in line.split("|")]
        models.append((model, provider))

# 2. 读任务模板
with open(os.path.join(BASE, "task", "task_prompt.md"), encoding="utf-8") as f:
    template = f.read()

# 3. 构建任务列表（run 外层、model 内层，首批各模型各 1 个并发）
tasks = []
idx = 1
for run in range(1, N + 1):
    for model, provider in models:
        tasks.append((f"P{idx:02d}", model, provider, run))
        idx += 1

RESULT_DIR = os.path.join(BASE, "results")
os.makedirs(RESULT_DIR, exist_ok=True)
WORK_ROOT = os.path.join(RESULT_DIR, "work")
os.makedirs(WORK_ROOT, exist_ok=True)

print(f"共 {len(tasks)} 个任务（{len(models)} 模型 × {N} 次），并发 {JOBS}，开始...", flush=True)


def run_task(anon, model, provider, run):
    safe = model.replace("/", "_").replace(":", "_")
    workdir = os.path.join(WORK_ROOT, f"{safe}_run{run}")
    os.makedirs(workdir, exist_ok=True)
    deck = os.path.join(workdir, "deck.pptx")
    task = template.replace("__OUTPUT_PPTX__", deck)
    with open(os.path.join(workdir, "task.md"), "w", encoding="utf-8") as f:
        f.write(task)

    start = time.time()
    cmd = [HERMES, "chat", "-q", task, "-m", model,
           "--provider", provider, "-s", "ppt-builder", "-Q", "--yolo"]
    with open(os.path.join(workdir, "log.txt"), "w", encoding="utf-8") as log:
        subprocess.run(cmd, stdout=log, stderr=subprocess.STDOUT)
    elapsed = int(time.time() - start)

    if os.path.exists(deck):
        os.replace(deck, os.path.join(RESULT_DIR, f"{anon}.pptx"))
        status = "OK"
    else:
        status = "FAILED"

    with open(os.path.join(workdir, "result.txt"), "w", encoding="utf-8") as f:
        f.write(f"{anon}\t{model}\t{provider}\trun{run}\t{elapsed}s\t{status}\n")
    print(f"[{status}] {anon}  {model}  run{run}  {elapsed}s", flush=True)
    return (anon, model, provider, run, elapsed, status)


results = []
with ThreadPoolExecutor(max_workers=JOBS) as ex:
    futs = {ex.submit(run_task, *t): t for t in tasks}
    for fut in as_completed(futs):
        results.append(fut.result())

results.sort(key=lambda r: r[0])
with open(os.path.join(RESULT_DIR, "_mapping.tsv"), "w", encoding="utf-8") as f:
    for r in results:
        f.write("\t".join(str(x) for x in r) + "\n")

ok = sum(1 for r in results if r[5] == "OK")
print(f"\n完成：{ok}/{len(results)} 成功。映射表 results/_mapping.tsv", flush=True)
