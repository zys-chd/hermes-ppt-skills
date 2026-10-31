"""
生成研发攻关失效数据集 —— 星云 X1 智能音箱电源模块市场早期失效

场景：X1 电源模块量产上市后市场退回率超标，需定位根因、给出优化方案。

埋点（6 个可被分析发现的故事点）：
  R1 早期失效：失效时间服从 Weibull，形状参数 β<1（早期失效特征，非随机/磨损）
  R2 料是主因：供应商 A 的 A-2411 批次失效率 ≈3%（正常批次 ≈0.5%，约 6 倍）
  R3 料×环交互（真实根因）：A-2411 在北方(低温)失效率 4.5% vs 南方 1.5%，
     其他批次南北无差异 → 低温加速 A-2411 批次电容 ESR 劣化
  R4 机是假象：Line2 表面失效率偏高，但控制 batch 后消失（Line2 恰好多产 A-2411）
  R5 失效模式指纹：A-2411 以"无输出/间歇性重启"为主（电容失效特征），
     其他批次以"过热保护"等为主
  R6 环是干扰：北方整体退回多（低温），但根因是料×环，不是纯环境问题

输出两个文件：
  shipment_summary.csv   供应商×批次×地区×产线 的出货量（失效率分母）
  failure_records.csv    每条失效记录（含时间/失效模式等）
"""
import csv
import os
import numpy as np
from datetime import date, timedelta

rng = np.random.default_rng(2024)

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, "data")

# ---------------------------------------------------------------- 出货量
# 总产量约 10 万台。供应商 A 6万，B 4万。
A_BATCHES = [("A-2410", 12000), ("A-2411", 30000), ("A-2412", 18000)]
B_BATCHES = [("B-2410", 20000), ("B-2411", 20000)]
REGIONS = ["北方", "南方"]
LINES = ["Line1", "Line2", "Line3"]

# 产线产量占比：A-2411 故意让 Line2 生产最多（制造"机"假象）
def line_share(batch, line):
    if batch == "A-2411":
        return {"Line1": 0.30, "Line2": 0.40, "Line3": 0.30}[line]
    return 1.0 / 3.0

shipment = []
for supplier, batches in (("A", A_BATCHES), ("B", B_BATCHES)):
    for batch, qty in batches:
        for region in REGIONS:
            region_share = 0.60 if region == "北方" else 0.40
            for line in LINES:
                q = int(qty * region_share * line_share(batch, line))
                shipment.append({"supplier": supplier, "batch": batch,
                                 "region": region, "line": line, "shipped_qty": q})

with open(f"{DATA}/shipment_summary.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=["supplier", "batch", "region", "line", "shipped_qty"])
    w.writeheader()
    w.writerows(shipment)

# ---------------------------------------------------------------- 失效率与失效记录
def failure_rate(batch, region):
    if batch == "A-2411":
        return 0.045 if region == "北方" else 0.015   # 料×环交互（低温加速）
    return 0.005                                      # 正常批次

def failure_mode(batch):
    if batch == "A-2411":
        return rng.choice(["无输出", "间歇性重启", "过热保护", "其他"],
                          p=[0.50, 0.35, 0.10, 0.05])
    return rng.choice(["无输出", "间歇性重启", "过热保护", "其他"],
                      p=[0.20, 0.15, 0.40, 0.25])

def time_to_failure(batch):
    if batch == "A-2411":
        ttf = int(rng.weibull(0.9) * 45)    # 早期失效，中位寿命短
    else:
        ttf = int(rng.weibull(1.5) * 200)   # 随机/磨损期
    return max(1, min(ttf, 180))

START = date(2024, 10, 1)
END = date(2025, 3, 31)
DAYS_RANGE = (END - START).days

records = []
rid = 1
for s in shipment:
    rate = failure_rate(s["batch"], s["region"])
    n_fail = int(rng.binomial(s["shipped_qty"], rate))
    for _ in range(n_fail):
        ttf = time_to_failure(s["batch"])
        prod = START + timedelta(days=int(rng.integers(0, DAYS_RANGE + 1)))
        fail = prod + timedelta(days=ttf)
        records.append({
            "record_id": rid,
            "serial_no": f"X1-{rid:06d}",
            "production_date": prod.isoformat(),
            "failure_date": fail.isoformat(),
            "time_to_failure_days": ttf,
            "supplier": s["supplier"],
            "batch": s["batch"],
            "line": s["line"],
            "region": s["region"],
            "failure_mode": failure_mode(s["batch"]),
        })
        rid += 1

rng.shuffle(records)

with open(f"{DATA}/failure_records.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=list(records[0].keys()))
    w.writeheader()
    w.writerows(records)

print(f"failure_records: {len(records)} 条")

# ---------------------------------------------------------------- 验证埋点
import pandas as pd
ship = pd.read_csv(f"{DATA}/shipment_summary.csv")
fail = pd.read_csv(f"{DATA}/failure_records.csv")

ship_fail = fail.groupby(["supplier", "batch", "region", "line"]).size().reset_index(name="fail")
ship = ship.merge(ship_fail, on=["supplier", "batch", "region", "line"], how="left")
ship["fail"] = ship["fail"].fillna(0).astype(int)

print("\n== R2/R3 各批次×地区失效率 ==")
g = ship.groupby(["batch", "region"]).agg(shipped=("shipped_qty", "sum"), fail=("fail", "sum"))
g["rate"] = (g.fail / g.shipped * 100).round(2)
print(g[["shipped", "fail", "rate"]].to_string())

print("\n== R4 各产线失效率（整体 vs 剔除A-2411）==")
for subset, name in [(ship, "整体"), (ship[ship.batch != "A-2411"], "剔除A-2411")]:
    gl = subset.groupby("line").agg(shipped=("shipped_qty", "sum"), fail=("fail", "sum"))
    gl["rate"] = (gl.fail / gl.shipped * 100).round(2)
    print(f"-- {name} --\n{gl[['rate']].to_string()}")

print("\n== R5 失效模式分布 ==")
fm = fail.groupby(["batch", "failure_mode"]).size().unstack(fill_value=0)
fm_pct = fm.div(fm.sum(axis=1), axis=0).mul(100).round(1)
print(fm_pct.to_string())

print("\n== R1 Weibull 形状参数 β（用 fit 估计）==")
from scipy.stats import weibull_min
for batch in ["A-2411", "A-2412", "B-2410"]:
    d = fail[fail.batch == batch].time_to_failure_days
    if len(d) > 10:
        params = weibull_min.fit(d, floc=0)
        beta = params[0]
        print(f"{batch}: n={len(d)}, β≈{beta:.2f} (中位 {d.median():.0f} 天)")
