"""生成物证图片（evidence/）：供模型在 PPT 中"证据展示"环节有机嵌入。"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

plt.rcParams["font.sans-serif"] = ["PingFang SC", "Heiti SC", "Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

EVID = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "evidence")
os.makedirs(EVID, exist_ok=True)


# ---------------------------------------------------------------- 1. 电解电容鼓包
def draw_capacitor():
    fig, ax = plt.subplots(figsize=(7, 4.5))
    ax.set_xlim(0, 10); ax.set_ylim(0, 6); ax.axis("off")
    ax.set_title("失效件物证：A-2411 批次电解电容顶部鼓包、漏液", fontsize=12, color="#1F3864", pad=10)
    # 电容主体（圆柱）
    body = mpatches.FancyBboxPatch((3.6, 0.8), 2.8, 3.6,
                                   boxstyle="round,pad=0.15,rounding_size=0.4",
                                   facecolor="#3A3A3A", edgecolor="#222", linewidth=2)
    ax.add_patch(body)
    # 顶部鼓包（凸起）
    bulge = mpatches.Ellipse((5.0, 4.55), 2.9, 1.15, facecolor="#4A4A4A", edgecolor="#222", linewidth=2)
    ax.add_patch(bulge)
    # 防爆十字刻痕（裂开）
    ax.plot([5.0, 5.0], [4.1, 5.05], color="#111", lw=2.5)
    ax.plot([3.75, 6.25], [4.6, 4.6], color="#111", lw=2.5)
    # 漏液痕迹（顶部流下的黄色）
    leak = mpatches.Polygon([[4.2, 4.6], [4.45, 4.6], [4.35, 2.6], [4.05, 2.6]],
                            facecolor="#C8A24A", edgecolor="none", alpha=0.85)
    ax.add_patch(leak)
    # 引脚
    ax.plot([4.6, 4.2], [0.8, 0.1], color="#777", lw=3)
    ax.plot([5.4, 5.8], [0.8, 0.1], color="#777", lw=3)
    # 标注
    ax.annotate("顶部鼓包\n（内压升高）", xy=(5.0, 5.1), xytext=(7.3, 5.0),
                fontsize=10, color="#E66D2C", ha="center",
                arrowprops=dict(arrowstyle="->", color="#E66D2C"))
    ax.annotate("电解液泄漏", xy=(4.2, 3.5), xytext=(1.2, 3.2),
                fontsize=10, color="#C8A24A",
                arrowprops=dict(arrowstyle="->", color="#C8A24A"))
    fig.tight_layout()
    fig.savefig(f"{EVID}/evidence_capacitor_bulge.png", dpi=150, facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------- 2. X-ray 焊点空洞
def draw_xray():
    fig, ax = plt.subplots(figsize=(6.5, 5))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10); ax.axis("off")
    ax.set_title("物证：X-ray 检测 BGA 焊点空洞（黑色区域）", fontsize=12, color="#1F3864", pad=10)
    rng = np.random.default_rng(7)
    xs = np.repeat(np.arange(2, 10, 1.6), 5)
    ys = np.tile(np.arange(2, 9, 1.6), 5)
    for x, y in zip(xs, ys):
        # 焊球（浅灰圆）
        ax.add_patch(mpatches.Circle((x, y), 0.55, facecolor="#C8C8C8", edgecolor="#888", lw=0.6))
        # 空洞（深色不规则斑，部分焊点有）
        if rng.random() < 0.45:
            for _ in range(int(rng.integers(1, 4))):
                vx = x + rng.uniform(-0.3, 0.3)
                vy = y + rng.uniform(-0.3, 0.3)
                ax.add_patch(mpatches.Circle((vx, vy), rng.uniform(0.06, 0.16),
                                             facecolor="#222", edgecolor="none"))
    ax.annotate("焊点内部空洞\n（焊接不良）", xy=(3.6, 3.6), xytext=(0.6, 6.5),
                fontsize=10, color="#E66D2C",
                arrowprops=dict(arrowstyle="->", color="#E66D2C"))
    fig.tight_layout()
    fig.savefig(f"{EVID}/evidence_xray_void.png", dpi=150, facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------- 3. 示波器波形对比
def draw_waveform():
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(9, 5.2), sharex=True)
    t = np.linspace(0, 10, 2000)
    # 正常：5V 直流 + 小纹波
    v_norm = 5.0 + 0.05 * np.sin(2 * np.pi * 1.2 * t)
    # 异常：5V 直流 + 大幅纹波 + 周期性跌落（电容 ESR 劣化导致）
    v_ab = 5.0 + 0.6 * np.sin(2 * np.pi * 1.2 * t) + 0.5 * (np.sign(np.sin(2 * np.pi * 0.15 * t)) * 0.5)
    for ax, v, label, color in [(ax1, v_norm, "正常批次输出：纹波 ±50mV", "#2E75B6"),
                                 (ax2, v_ab, "A-2411 批次输出：纹波异常、周期性跌落", "#E66D2C")]:
        ax.plot(t, v, color=color, lw=1.3)
        ax.set_ylim(3.0, 7.5)
        ax.grid(True, color="#D9D9D9", lw=0.6)
        ax.set_facecolor("#F7F9FB")
        ax.set_ylabel("输出电压 (V)", fontsize=9)
        ax.text(0.02, 0.9, label, transform=ax.transAxes, fontsize=10, color=color, fontweight="bold")
        for sp in ["top", "right"]:
            ax.spines[sp].set_visible(False)
    ax2.set_xlabel("时间 (ms)", fontsize=9)
    fig.suptitle("物证：示波器实测输出波形对比（环境温度 25℃）", fontsize=12, color="#1F3864", y=0.98)
    fig.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(f"{EVID}/evidence_waveform.png", dpi=150, facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------- 4. 板卡腐蚀/烧蚀
def draw_board():
    fig, ax = plt.subplots(figsize=(7, 4.8))
    ax.set_xlim(0, 12); ax.set_ylim(0, 8); ax.axis("off")
    ax.set_title("物证：失效板卡局部电解液腐蚀痕迹", fontsize=12, color="#1F3864", pad=10)
    # PCB 基板
    ax.add_patch(mpatches.Rectangle((0.8, 0.8), 10.4, 6.4, facecolor="#1E5631", edgecolor="#14401F", lw=2))
    # 走线
    rng = np.random.default_rng(3)
    for _ in range(8):
        x0, y0 = rng.uniform(1.2, 10), rng.uniform(1.2, 6.5)
        ax.plot([x0, x0 + rng.uniform(0.5, 1.5)], [y0, y0], color="#C9A227", lw=1.2)
    # 腐蚀区域（棕黄色斑块）
    cx, cy = 7.5, 4.2
    for _ in range(60):
        x = cx + rng.uniform(-1.5, 1.5)
        y = cy + rng.uniform(-1.0, 1.0)
        r = rng.uniform(0.05, 0.35)
        ax.add_patch(mpatches.Circle((x, y), r, facecolor="#B8860B", edgecolor="#8B6914", alpha=0.7))
    ax.annotate("电解液渗漏腐蚀区域", xy=(7.5, 4.0), xytext=(1.0, 6.8),
                fontsize=10, color="#8B6914",
                arrowprops=dict(arrowstyle="->", color="#8B6914"))
    # 元件
    for (ex, ey) in [(2.5, 2.0), (4.5, 5.0), (2.0, 5.5), (9.8, 2.2)]:
        ax.add_patch(mpatches.Rectangle((ex - 0.4, ey - 0.25), 0.8, 0.5,
                                        facecolor="#2A2A2A", edgecolor="#111", lw=1))
    fig.tight_layout()
    fig.savefig(f"{EVID}/evidence_board_corrosion.png", dpi=150, facecolor="white")
    plt.close(fig)


draw_capacitor()
draw_xray()
draw_waveform()
draw_board()
print("evidence images written to", EVID)
for f in sorted(os.listdir(EVID)):
    print("  ", f, os.path.getsize(f"{EVID}/{f}"), "bytes")
