#!/usr/bin/env python3
"""验收脚本：检查生成的 pptx 是否用了原生图表。

用法：
    python3 verify_native_charts.py <deck.pptx>

判定规则：
  - 统计 pptx 内的原生图表对象（chart XML）和图片数量
  - 原生图表 >= 3 且 图片数 <= (物证图数量 + 2)  → 通过
  - 否则打印失败原因和修正指引

物证图例外：任务给的证据照片本来就该是图片，所以允许少量图片。
热力图 / Weibull 概率图也允许图片，用 --allow-images N 放宽容差。
"""
import sys
import zipfile
import argparse


def check(path, allow_images=2):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        charts = [n for n in names if n.startswith("ppt/charts/chart") and n.endswith(".xml")]
        embeds = [n for n in names if "embeddings" in n and n.endswith(".xlsx")]
        media = [n for n in names if n.startswith("ppt/media/")]

    n_chart = len(charts)
    n_embed = len(embeds)
    n_img = len(media)

    print(f"文件: {path}")
    print(f"  原生图表 (chart XML): {n_chart}")
    print(f"  可编辑数据 (嵌入 xlsx): {n_embed}")
    print(f"  图片 (media): {n_img}")
    print()

    ok = True
    if n_chart < 3:
        print(f"❌ 原生图表只有 {n_chart} 个，要求 >= 3")
        print("   → 柱状/折线/饼/散点必须用 native_bar / native_line / native_pie / native_scatter")
        print("   → 不要自己写 matplotlib 画基础图表（那样是图片，PPT 内不可编辑）")
        ok = False

    if n_img > allow_images:
        print(f"⚠️  图片 {n_img} 张，超出容差 {allow_images} 张")
        print("   → 除物证照片、热力图、Weibull 概率图外，不应有其他图片")
        print("   → 检查是否把数据图表存成了 PNG")
        ok = False

    if n_chart > 0 and n_embed == 0:
        print("⚠️  有原生图表但没有嵌入数据，可能不是真正的原生图表")
        ok = False

    if ok:
        print("✅ 通过：原生图表使用充分")
    else:
        print("\n❌ 未通过 —— 请修正后重新生成")
    return ok


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--allow-images", type=int, default=2,
                    help="允许的图片数（物证图数量），默认 2")
    a = ap.parse_args()
    sys.exit(0 if check(a.pptx, a.allow_images) else 1)
