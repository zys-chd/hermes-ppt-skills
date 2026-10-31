"""
pptx_helpers.py — 专业 PPT 设计工具库（非固定模板）

设计理念：本库提供「底层能力 + 多套风格 + 装饰元素」，不强制你套用某一种版式。
你可以用参考版式函数快速起稿，也可以只用底层函数 + STYLES 自己设计布局——
「长什么样」的决策权在你，用差异化设计体现表达意图。

依赖：python-pptx, matplotlib, numpy, pandas
用法：
    import sys; sys.path.insert(0, "<本文件所在目录>")
    from pptx_helpers import *
"""
from __future__ import annotations

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

# ---------------------------------------------------------------- 基础常量

SLIDE_W = Inches(13.333)      # 16:9
SLIDE_H = Inches(7.5)
MARGIN_L = Inches(0.6)
MARGIN_T = Inches(0.5)

# 中文字体：按平台自动检测（macOS/Windows/Linux），也可用环境变量 PPT_CN_FONT 覆盖
import platform as _platform

_CN_FONT_BY_PLATFORM = {
    "Darwin":  ["PingFang SC", "Heiti SC", "Songti SC", "Hiragino Sans GB"],
    "Windows": ["Microsoft YaHei", "SimHei", "SimSun"],
    "Linux":   ["Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Micro Hei",
                "Noto Sans CJK", "Droid Sans Fallback"],
}

# 安全字体白名单：这些字体在中英文 Windows/macOS 下都不会缺字（避免乱码）
SAFE_CN_FONTS = {"PingFang SC", "Microsoft YaHei", "SimHei", "SimSun",
                 "Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Micro Hei",
                 "Heiti SC", "Songti SC", "Arial Unicode MS"}
# 已知跨平台易缺字/易乱码的字体（禁止使用）
UNSAFE_FONTS = {"Hiragino Sans GB", "Hiragino Sans", "STHeiti", "Droid Sans Fallback"}


def _detect_cn_font():
    """检测当前系统实际可用的中文字体，返回字体名（找不到则回退平台首选）。"""
    try:
        from matplotlib import font_manager
        installed = {f.name for f in font_manager.fontManager.ttflist}
    except Exception:
        installed = set()
    order = _CN_FONT_BY_PLATFORM.get(_platform.system(), [])
    for name in order:
        if name in installed:
            return name
    return order[0] if order else "PingFang SC"


CN_FONT = os.environ.get("PPT_CN_FONT") or _detect_cn_font()

_CN_FONT_FALLBACKS = []
for _lst in _CN_FONT_BY_PLATFORM.values():
    for _f in _lst:
        if _f not in _CN_FONT_FALLBACKS:
            _CN_FONT_FALLBACKS.append(_f)
CN_FONT_FALLBACKS = _CN_FONT_FALLBACKS

# ---------------------------------------------------------------- 多套风格

def _c(hexstr):
    hexstr = hexstr.lstrip("#")
    return RGBColor(*[int(hexstr[i:i+2], 16) for i in (0, 2, 4)])


hexcolor = _c   # 公开别名：hexcolor("#FF6B35") → RGBColor


# 每套风格 = 配色 + 字体 + 设计语言。用 get_style(name) 取用。
STYLES = {
    "tech": dict(
        name="科技深色", dark=True,
        bg="#0E1A2B", panel="#1A2A40", primary="#00D4FF", accent="#FF6B35",
        text="#E8EEF5", muted="#8FA3BF", light="#243650", white="#FFFFFF",
        series=["#00D4FF", "#FF6B35", "#7ED321", "#FFD700", "#9B59B6"],
        font=CN_FONT,
        note="深色底 + 亮青/橙点缀 + 几何网格，适合技术评审/工程攻关",
    ),
    "minimal": dict(
        name="极简留白", dark=False,
        bg="#FFFFFF", panel="#F5F5F3", primary="#111111", accent="#C00000",
        text="#1A1A1A", muted="#8A8A8A", light="#F0F0EE", white="#FFFFFF",
        series=["#444444", "#C00000", "#9A9A9A", "#D6D6D6", "#767676"],
        font=CN_FONT,
        note="白底大量留白 + 单一红色点缀 + 细线，适合严肃汇报/评审答辩",
    ),
    "editorial": dict(
        name="杂志图文", dark=False,
        bg="#FAF7F0", panel="#F0EBE0", primary="#1F3A2D", accent="#C77B3F",
        text="#26221A", muted="#8A8272", light="#EEE8DA", white="#FFFFFF",
        series=["#1F3A2D", "#C77B3F", "#8AA38F", "#D9B98A", "#5C6B63"],
        font=CN_FONT,
        note="米白底 + 墨绿/陶土色 + 大图主导图文混排，适合有故事性的汇报",
    ),
    "warm": dict(
        name="暖色商务", dark=False,
        bg="#FBF7F2", panel="#FFFFFF", primary="#173A5E", accent="#E08A2E",
        text="#2B2B2B", muted="#7A7A7A", light="#F3EADF", white="#FFFFFF",
        series=["#2E75B6", "#E08A2E", "#70AD47", "#C25E5E", "#8B7FC7"],
        font=CN_FONT,
        note="暖灰底 + 深蓝/琥珀 + 圆角色块卡片，适合管理层/商务汇报",
    ),
}

DEFAULT_STYLE = "minimal"


def _resolve(s):
    """把 style dict 里的 hex 颜色解析成 RGBColor（浅拷贝，不改原 STYLES）。"""
    s = dict(s)
    for k in ["bg", "panel", "primary", "accent", "text", "muted", "light", "white"]:
        v = s.get(k)
        if isinstance(v, str) and v.startswith("#"):
            s[k] = _c(v)
    return s


def get_style(name=None):
    """返回风格 dict（含 RGBColor 已解析）。name 为 None 用默认风格。"""
    return _resolve(STYLES.get(name or DEFAULT_STYLE, STYLES[DEFAULT_STYLE]))


def expand_palette(*base_hexes):
    """从 1~3 个基础色（用户从网站找的 hex）拓展出完整配色方案。

    返回 style dict（RGBColor 已解析），用法同 get_style：
        st = expand_palette("#2E75B6", "#E08A2E")
    没传基础色时返回默认风格。"""
    import colorsys

    def _h2rgb(h):
        h = h.lstrip("#")
        return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))

    def _hsl(h, s, l):
        r, g, b = colorsys.hls_to_rgb(h % 1.0, l, s)
        return "#%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))

    bases = [b.strip() for b in base_hexes if b and b.strip()]
    if not bases:
        return get_style()

    h0, l0, s0 = colorsys.rgb_to_hls(*_h2rgb(bases[0]))
    primary = bases[0]
    accent = bases[1] if len(bases) >= 2 else _hsl(h0 + 0.5, max(s0, 0.7), 0.55)

    series = []
    seeds = bases if len(bases) >= 2 else [primary]
    for i in range(5):
        h = colorsys.rgb_to_hls(*_h2rgb(seeds[i % len(seeds)]))[0]
        series.append(_hsl(h + i * 0.10, 0.62, 0.52 + (i % 2) * 0.08))

    return _resolve({
        "name": "自定义配色",
        "dark": False,
        "bg": "#FFFFFF",
        "panel": _hsl(h0, 0.12, 0.97),
        "primary": primary,
        "accent": accent,
        "text": "#262626",
        "muted": "#7A7A7A",
        "light": _hsl(h0, 0.32, 0.94),
        "white": "#FFFFFF",
        "series": series,
        "font": CN_FONT,
        "note": "由用户基础色拓展",
    })


# 兼容旧用法：COLORS 指向默认风格的配色
COLORS = get_style(DEFAULT_STYLE)
SERIES_COLORS = [get_style(DEFAULT_STYLE)["series"][0]]  # 旧代码兜底


def setup_matplotlib_cn(series=None):
    """让 matplotlib 正确显示中文 + 负号。series 可传风格系列色。"""
    plt.rcParams["font.sans-serif"] = CN_FONT_FALLBACKS
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["font.size"] = 11
    if series:
        plt.rcParams["axes.prop_cycle"] = plt.cycler(color=series)


# ---------------------------------------------------------------- 底层工具

def new_deck():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def set_bg(slide, color):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = color


def rect(slide, left, top, width, height, color, round_=False, line=None, radius=0.15):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE,
        left, top, width, height)
    if round_:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    return shp


def line(slide, x1, y1, x2, y2, color, width=1.5):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = Pt(width)
    return c


def text(slide, left, top, width, height, lines, size=14, bold=False,
         color=None, align=PP_ALIGN.LEFT, font=None, space_after=6, line_spacing=1.1):
    """富文本框。lines 是 list[str] 或 list[tuple(text, {size,bold,color,align,font})]。"""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        if isinstance(item, str):
            _run(p.add_run(), item, size, bold, color, font or CN_FONT)
        else:
            t, kw = item
            _run(p.add_run(), t,
                 kw.get("size", size), kw.get("bold", bold),
                 kw.get("color", color), kw.get("font", font or CN_FONT))
            p.alignment = kw.get("align", align)
    return tb


def _run(run, t, size, bold, color, font):
    run.text = t
    run.font.size = Pt(size)
    run.font.bold = bold
    # 字体安全校验：不安全字体一律替换为 CN_FONT，防止跨平台乱码
    if font in UNSAFE_FONTS:
        font = CN_FONT
    run.font.name = font
    if color is not None:
        run.font.color.rgb = color


def picture(slide, path, left, top, width=None, height=None):
    kw = {}
    if width:
        kw["width"] = width
    if height:
        kw["height"] = height
    return slide.shapes.add_picture(path, left, top, **kw)


# ---------------------------------------------------------------- 装饰元素（去模板化关键）

def big_number(slide, left, top, number, label, style, width=Inches(2.6),
               num_size=44, num_color=None, note=""):
    """大数字 KPI——视觉焦点，制造『一页一重点』。number 数字、label 标签、note 补充。"""
    num_color = num_color or style["accent"]
    text(slide, left, top, width, Inches(0.9),
         [(str(number), {"size": num_size, "bold": True, "color": num_color})])
    text(slide, left, top + Inches(0.95), width, Inches(0.4),
         [(label, {"size": 14, "bold": True, "color": style["primary"]})])
    if note:
        text(slide, left, top + Inches(1.35), width, Inches(0.4),
             [(note, {"size": 10, "color": style["muted"]})])
    return top + Inches(1.7)


def tag(slide, left, top, label, color, text_color=None, height=Inches(0.34)):
    """圆角小标签/徽章（如『根因』『证据』『对策』）。"""
    tb = slide.shapes.add_textbox(left, top, Inches(0.2 + 0.5 * len(label)), height)
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _run(p.add_run(), label, 11, True, text_color or RGBColor(0xFF, 0xFF, 0xFF), CN_FONT)
    # 用圆角矩形做底（先加底再加字，保证字在上层）
    from pptx.enum.shapes import MSO_SHAPE as _S
    w = Inches(0.25 + 0.55 * len(label))
    bg = slide.shapes.add_shape(_S.ROUNDED_RECTANGLE, left, top, w, height)
    try:
        bg.adjustments[0] = 0.5
    except Exception:
        pass
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    # 把文字框移到最上层
    from pptx.oxml.ns import qn
    tb._element.getparent().remove(tb._element)
    bg._element.addnext(tb._element)
    return left + w


def kicker(slide, left, top, label, color):
    """标题上方的小字 kicker（如『01 · 问题定义』）。"""
    text(slide, left, top, Inches(6), Inches(0.35),
         [(label, {"size": 11, "bold": True, "color": color})])


def title_bar(slide, title, style, kicker_text=None, page=None):
    """标准标题：kicker + 主标题 + 一条点缀线。返回正文起始 y。"""
    top = Inches(0.45)
    if kicker_text:
        kicker(slide, MARGIN_L, top, kicker_text, style["accent"])
        top += Inches(0.38)
    text(slide, MARGIN_L, top, Inches(12), Inches(0.75),
         [(title, {"size": 25, "bold": True, "color": style["primary"]})])
    line(slide, MARGIN_L, top + Inches(0.78), MARGIN_L + Inches(1.2),
         top + Inches(0.78), style["accent"], 2.5)
    if page is not None:
        footer(slide, page, style)
    return top + Inches(1.05)


def footer(slide, page, style, deck_title=""):
    text(slide, MARGIN_L, SLIDE_H - Inches(0.42), Inches(6), Inches(0.3),
         [(deck_title, {"size": 9, "color": style["muted"]})])
    text(slide, SLIDE_W - Inches(1.2), SLIDE_H - Inches(0.42), Inches(0.8), Inches(0.3),
         [(str(page), {"size": 9, "color": style["muted"]})], align=PP_ALIGN.RIGHT)


def flow(slide, steps, top, style, left=Inches(0.6), width=Inches(12.1)):
    """横向流程条：steps 是 list[str]，自动排成箭头连接的流程。"""
    n = len(steps)
    gap = Inches(0.5)
    box_w = int((width - gap * (n - 1)) / n)
    x = left
    for i, s in enumerate(steps):
        rect(slide, x, top, box_w, Inches(0.7), style["panel"], round_=True)
        text(slide, x, top + Inches(0.15), box_w, Inches(0.4),
             [(s, {"size": 13, "bold": True, "color": style["primary"]})],
             align=PP_ALIGN.CENTER)
        if i < n - 1:
            line(slide, x + box_w, top + Inches(0.35), x + box_w + gap,
                 top + Inches(0.35), style["accent"], 2)
        x += box_w + gap
    return top + Inches(0.9)


# ---------------------------------------------------------------- 图表（matplotlib，统一中文）

def save_chart(fig, path, dpi=150):
    """保存 matplotlib 图表为 PNG。

    ⚠️⚠️ 这是「图片」路线，PPT 里不可编辑。
    只有这三类应该用它：热力图(make_heatmap)、Weibull 概率图、复杂组合图。
    柱状/条形/折线/饼/散点 —— 请改用 native_bar / native_line / native_pie / native_scatter。
    """
    # 热力图等显式标记允许
    if getattr(fig, "_image_chart_ok", False):
        fig.savefig(path, dpi=dpi, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        return path

    # 识别基础图表 → 强提示（打到 stdout，模型能看见）
    try:
        axs = [a for a in fig.axes if hasattr(a, "get_ylim")]
        is_basic = any(
            (a.patches or a.lines or getattr(a, "collections", None))
            and not getattr(a, "images", None)
            for a in axs
        )
        if is_basic:
            msg = (
                "\n" + "!" * 70 + "\n"
                f"!! 警告：正在把基础图表存成图片 -> {path}\n"
                "!! 柱/折/饼/散点 应该用 native_bar / native_line / native_pie / native_scatter\n"
                "!! 这些生成 PPT 原生图表，用户可右键编辑数据；图片不可编辑，属质量降级。\n"
                "!! 只有 热力图 / Weibull概率图 / 复杂组合图 才应该保存为图片。\n"
                + "!" * 70
            )
            print(msg, flush=True)
            import warnings
            warnings.warn(msg, UserWarning, stacklevel=2)
    except Exception:
        pass

    fig.savefig(path, dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path

def _style_ax(ax, title="", ylabel="", xlabel="", title_color="#111111"):
    for s in ["top", "right"]:
        ax.spines[s].set_visible(False)
    for s in ["left", "bottom"]:
        ax.spines[s].set_color("#B0B0B0")
    ax.grid(axis="y", linestyle="--", alpha=0.35, color="#D9D9D9")
    ax.set_axisbelow(True)
    if title:
        ax.set_title(title, fontsize=13, fontweight="bold", color=title_color, loc="left")
    if ylabel:
        ax.set_ylabel(ylabel)
    if xlabel:
        ax.set_xlabel(xlabel)
    return ax


def make_bar(labels, values, title="", ylabel="", xlabel="", horizontal=False,
             figsize=(8.5, 4.6), value_labels=True, highlight_idx=None,
             colors=None, title_color="#111111"):
    setup_matplotlib_cn()
    fig, ax = plt.subplots(figsize=figsize)
    colors = colors or ["#2E75B6"] * len(labels)
    if highlight_idx is not None:
        colors = ["#C9C9C9"] * len(labels)
        colors[highlight_idx] = "#E66D2C"
    if horizontal:
        bars = ax.barh(labels, values, color=colors)
        ax.invert_yaxis()
        if value_labels:
            for b, v in zip(bars, values):
                ax.text(v, b.get_y() + b.get_height() / 2, f" {v:,.1f}",
                        va="center", fontsize=10)
    else:
        bars = ax.bar(labels, values, color=colors)
        if value_labels:
            for b, v in zip(bars, values):
                ax.text(b.get_x() + b.get_width() / 2, b.get_height(), f"{v:,.1f}",
                        ha="center", va="bottom", fontsize=10)
    _style_ax(ax, title, ylabel, xlabel, title_color)
    fig.tight_layout()
    return fig


def make_line(x, series, title="", ylabel="", xlabel="", labels=None,
              figsize=(8.5, 4.6), markers=True, colors=None, title_color="#111111"):
    setup_matplotlib_cn()
    fig, ax = plt.subplots(figsize=figsize)
    if labels is None:
        labels = [None] * len(series)
    for i, s in enumerate(series):
        if s and isinstance(s[0], (list, tuple)):
            xs = [p[0] for p in s]; ys = [p[1] for p in s]
        else:
            xs = x; ys = s
        ax.plot(xs, ys, marker="o" if markers else None, lw=2.2,
                color=(colors[i] if colors else None),
                label=labels[i] if labels[i] else None)
    if any(labels):
        ax.legend(frameon=False, fontsize=10)
    _style_ax(ax, title, ylabel, xlabel, title_color)
    fig.tight_layout()
    return fig


def make_pie(labels, values, title="", figsize=(6.5, 4.6), donut=False,
             colors=None, title_color="#111111"):
    setup_matplotlib_cn()
    fig, ax = plt.subplots(figsize=figsize)
    colors = colors or plt.rcParams["axes.prop_cycle"].by_key()["color"][:len(labels)]
    kw = dict(colors=colors, autopct="%1.1f%%", startangle=90, pctdistance=0.75)
    if donut:
        kw["wedgeprops"] = {"width": 0.42, "edgecolor": "white"}
    else:
        kw["wedgeprops"] = {"edgecolor": "white"}
    ax.pie(values, labels=labels, **kw)
    ax.set_title(title, fontsize=13, fontweight="bold", color=title_color, loc="left")
    fig.tight_layout()
    return fig


def make_heatmap(matrix, row_labels, col_labels, title="", xlabel="", ylabel="",
                 figsize=(8, 4.6), cmap="YlOrRd", annot=True, fmt=".2f",
                 title_color="#111111"):
    """热力图（料×环、料×机等交叉分析必备）。matrix: 二维 list（行×列）。
    注：热力图无原生图表对应，图片是正确选择，save_chart 不会警告。"""
    setup_matplotlib_cn()
    fig, ax = plt.subplots(figsize=figsize)
    mat = np.array(matrix, dtype=float)
    im = ax.imshow(mat, cmap=cmap, aspect="auto")
    ax.set_xticks(range(len(col_labels))); ax.set_xticklabels(col_labels)
    ax.set_yticks(range(len(row_labels))); ax.set_yticklabels(row_labels)
    if xlabel:
        ax.set_xlabel(xlabel)
    if ylabel:
        ax.set_ylabel(ylabel)
    if annot:
        for i in range(mat.shape[0]):
            for j in range(mat.shape[1]):
                ax.text(j, i, format(mat[i, j], fmt), ha="center", va="center",
                        fontsize=10, color="#222" if mat[i, j] < mat.max() * 0.7 else "#fff")
    ax.set_title(title, fontsize=13, fontweight="bold", color=title_color, loc="left")
    fig.colorbar(im, ax=ax, shrink=0.8)
    fig.tight_layout()
    fig._image_chart_ok = True  # 标记：热力图等允许用图片
    return fig


def make_scatter(x, y, title="", xlabel="", ylabel="", figsize=(7.5, 4.6),
                 color="#2E75B6", title_color="#111111"):
    setup_matplotlib_cn()
    fig, ax = plt.subplots(figsize=figsize)
    ax.scatter(x, y, c=color, alpha=0.7, edgecolors="white")
    _style_ax(ax, title, ylabel, xlabel, title_color)
    fig.tight_layout()
    return fig


# ---------------------------------------------------------------- 参考版式（可选，可自定义）

def cover(prs, title, subtitle="", meta="", style=None):
    """封面（参考实现）。想更出彩可用底层函数自己设计。"""
    s = get_style(style)
    sl = blank(prs)
    set_bg(sl, s["bg"])
    rect(sl, Inches(0.9), Inches(2.5), Inches(0.22), Inches(2.3), s["accent"])
    text(sl, Inches(1.35), Inches(2.6), Inches(11), Inches(1.6),
         [(title, {"size": 40, "bold": True, "color": s["text"]})])
    metas = []
    if subtitle:
        metas.append((subtitle, {"size": 18, "color": s["muted"]}))
    if meta:
        metas.append((meta, {"size": 14, "color": s["muted"]}))
    if metas:
        text(sl, Inches(1.35), Inches(4.3), Inches(10), Inches(1.5), metas)
    return sl


def section(prs, num, title, style=None, page=None):
    """章节分隔页。"""
    s = get_style(style)
    sl = blank(prs)
    set_bg(sl, s["bg"] if s.get("dark") else s["light"])
    big_number(sl, Inches(0.9), Inches(2.6), num, title, s, num_size=54,
               num_color=s["accent"])
    if page:
        footer(sl, page, s)
    return sl


def closing(prs, msg="谢谢", sub="", style=None):
    s = get_style(style)
    sl = blank(prs)
    set_bg(sl, s["bg"] if s.get("dark") else s["primary"])
    text(sl, Inches(0), Inches(3.0), SLIDE_W, Inches(1.2),
         [(msg, {"size": 44, "bold": True, "color": s["white"]})], align=PP_ALIGN.CENTER)
    if sub:
        text(sl, Inches(0), Inches(4.3), SLIDE_W, Inches(0.5),
             [(sub, {"size": 16, "color": s["muted"]})], align=PP_ALIGN.CENTER)
    return sl


# ================================================================
# 原生图表模块（Native Charts）—— 数据可在 PowerPoint 内直接编辑
# ================================================================
# 说明：以下函数生成 PowerPoint 原生图表对象（非图片），用户可在 PowerPoint 中
# 「右键 → 编辑数据」修改数据、换图表类型、应用图表样式。优先使用这些函数；
# 只有原生图表无法表达时（如热力图、鱼骨图、Weibull 概率图）才用 matplotlib 图片。

from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION


def _style_native_chart(chart, style, title="", show_legend=False, legend_pos=None,
                        data_labels=False, num_format='0.00', gap_width=60,
                        font_size=11, title_size=14):
    """统一美化原生图表：字体、配色、网格、数据标签。"""
    chart.has_title = bool(title)
    if title:
        tf = chart.chart_title.text_frame
        tf.text = title
        for p in tf.paragraphs:
            for r in p.runs:
                r.font.size = Pt(title_size)
                r.font.bold = True
                r.font.name = CN_FONT
                r.font.color.rgb = style["primary"]

    chart.has_legend = show_legend
    if show_legend:
        chart.legend.position = legend_pos or XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(font_size - 1)
        chart.legend.font.name = CN_FONT

    # 系列配色
    series_colors = style.get("series", [])
    for i, ser in enumerate(chart.series):
        if i < len(series_colors):
            try:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = hexcolor(series_colors[i])
            except Exception:
                pass
        if data_labels:
            try:
                ser.data_labels.show_value = True
                ser.data_labels.number_format = num_format
                ser.data_labels.number_format_is_linked = False
                ser.data_labels.font.size = Pt(font_size - 1)
                ser.data_labels.font.name = CN_FONT
                ser.data_labels.font.color.rgb = style["text"]
                ser.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            except Exception:
                pass

    # 坐标轴
    try:
        va = chart.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = hexcolor("#D9D9D9")
        va.major_gridlines.format.line.width = Pt(0.75)
        va.format.line.color.rgb = hexcolor("#B0B0B0")
        va.tick_labels.font.size = Pt(font_size - 1)
        va.tick_labels.font.name = CN_FONT
        va.tick_labels.font.color.rgb = style["muted"]
    except Exception:
        pass
    try:
        ca = chart.category_axis
        ca.tick_labels.font.size = Pt(font_size - 1)
        ca.tick_labels.font.name = CN_FONT
        ca.tick_labels.font.color.rgb = style["text"]
        ca.format.line.color.rgb = hexcolor("#B0B0B0")
    except Exception:
        pass
    try:
        chart.plots[0].gap_width = gap_width
    except Exception:
        pass
    return chart


def native_bar(slide, left, top, width, height, categories, series,
               style, title="", horizontal=False, show_legend=None,
               data_labels=True, num_format='0.00', colors=None):
    """原生柱状/条形图。series: dict{系列名: [值...]} 或 list[(名, [值...])]。"""
    if isinstance(series, dict):
        series = list(series.items())
    cd = CategoryChartData()
    cd.categories = categories
    for name, vals in series:
        cd.add_series(name, vals)
    ctype = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    gf = slide.shapes.add_chart(ctype, left, top, width, height, cd)
    st = dict(style)
    if colors:
        st["series"] = colors
    _style_native_chart(gf.chart, st, title,
                        show_legend=(len(series) > 1) if show_legend is None else show_legend,
                        data_labels=data_labels, num_format=num_format)
    return gf


def native_line(slide, left, top, width, height, categories, series,
                style, title="", markers=True, show_legend=None,
                data_labels=False, num_format='0.00', colors=None):
    """原生折线图。"""
    if isinstance(series, dict):
        series = list(series.items())
    cd = CategoryChartData()
    cd.categories = categories
    for name, vals in series:
        cd.add_series(name, vals)
    ctype = XL_CHART_TYPE.LINE_MARKERS if markers else XL_CHART_TYPE.LINE
    gf = slide.shapes.add_chart(ctype, left, top, width, height, cd)
    st = dict(style)
    if colors:
        st["series"] = colors
    _style_native_chart(gf.chart, st, title,
                        show_legend=(len(series) > 1) if show_legend is None else show_legend,
                        data_labels=data_labels, num_format=num_format)
    # 折线加粗
    for i, ser in enumerate(gf.chart.series):
        try:
            ser.format.line.width = Pt(2.25)
        except Exception:
            pass
    return gf


def native_pie(slide, left, top, width, height, categories, values,
               style, title="", donut=False, data_labels=True, num_format='0.0"%"'):
    """原生饼图/环形图。"""
    cd = CategoryChartData()
    cd.categories = categories
    cd.add_series("占比", values)
    ctype = XL_CHART_TYPE.DOUGHNUT if donut else XL_CHART_TYPE.PIE
    gf = slide.shapes.add_chart(ctype, left, top, width, height, cd)
    chart = gf.chart
    _style_native_chart(chart, style, title, show_legend=True,
                        legend_pos=XL_LEGEND_POSITION.RIGHT)
    # 逐扇区配色
    series_colors = style.get("series", [])
    try:
        for i, point in enumerate(chart.series[0].points):
            if i < len(series_colors):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = hexcolor(series_colors[i])
    except Exception:
        pass
    if data_labels:
        try:
            chart.plots[0].has_data_labels = True
            dl = chart.plots[0].data_labels
            dl.show_percentage = True
            dl.show_value = False
            dl.font.size = Pt(10)
            dl.font.name = CN_FONT
            dl.number_format = num_format
            dl.number_format_is_linked = False
        except Exception:
            pass
    return gf


def native_scatter(slide, left, top, width, height, series, style, title="",
                   show_legend=None, num_format='0.00'):
    """原生散点图。series: dict{名: (xs, ys)} 或 list[(名, (xs, ys))]。"""
    if isinstance(series, dict):
        series = list(series.items())
    cd = XyChartData()
    for name, (xs, ys) in series:
        s = cd.add_series(name)
        for x, y in zip(xs, ys):
            s.add_data_point(x, y)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, left, top, width, height, cd)
    st = dict(style)
    _style_native_chart(gf.chart, st, title,
                        show_legend=(len(series) > 1) if show_legend is None else show_legend,
                        data_labels=False, num_format=num_format)
    return gf


# ================================================================
# 原生图形模块（Native Shapes）—— 替代 SmartArt，全矢量、可编辑
# ================================================================

def native_flow(slide, steps, left, top, width, style, box_h=Inches(0.75),
                numbered=True, gap=Inches(0.45)):
    """原生横向流程条（替代 SmartArt 流程图）。steps: list[str]。"""
    n = len(steps)
    box_w = int((width - gap * (n - 1)) / n)
    x = left
    for i, txt in enumerate(steps):
        # 圆角矩形块
        shp = rect(slide, x, top, box_w, box_h, style["panel"], round_=True)
        shp.line.color.rgb = style["primary"]
        shp.line.width = Pt(1)
        label = f"{i+1}. {txt}" if numbered else txt
        text(slide, x, top + int(box_h * 0.22), box_w, int(box_h * 0.6),
             [(label, {"size": 12, "bold": True, "color": style["primary"]})],
             align=PP_ALIGN.CENTER)
        if i < n - 1:
            # 箭头
            ax1 = x + box_w
            arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                           ax1 + Inches(0.08), top + int(box_h * 0.3),
                                           gap - Inches(0.16), int(box_h * 0.4))
            arrow.fill.solid(); arrow.fill.fore_color.rgb = style["accent"]
            arrow.line.fill.background()
        x += box_w + gap
    return top + box_h


def native_chevron_flow(slide, steps, left, top, width, style, box_h=Inches(0.7)):
    """V 形箭头流程（替代 SmartArt Chevron List）。"""
    n = len(steps)
    overlap = Inches(0.22)
    box_w = int((width + overlap * (n - 1)) / n)
    x = left
    for i, txt in enumerate(steps):
        shp = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, x, top, box_w, box_h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = (style["primary"] if i % 2 == 0 else style["accent"])
        shp.line.fill.background()
        text(slide, x + int(box_w * 0.18), top + int(box_h * 0.25),
             int(box_w * 0.64), int(box_h * 0.5),
             [(txt, {"size": 11, "bold": True, "color": style["white"]})],
             align=PP_ALIGN.CENTER)
        x += box_w - overlap
    return top + box_h


def native_hierarchy(slide, root, children, left, top, width, style,
                     root_h=Inches(0.62), child_h=Inches(0.52), gap=Inches(0.35)):
    """原生层级图（替代 SmartArt 组织架构图）：根节点 + 子节点。"""
    # 根节点居中
    root_w = Inches(2.2)
    rx = left + int((width - root_w) / 2)
    r = rect(slide, rx, top, root_w, root_h, style["primary"], round_=True)
    text(slide, rx, top + Inches(0.14), root_w, Inches(0.4),
         [(root, {"size": 13, "bold": True, "color": style["white"]})], align=PP_ALIGN.CENTER)
    # 子节点
    n = len(children)
    cw = int((width - gap * (n - 1)) / n) if n else width
    cy = top + root_h + Inches(0.55)
    # 竖线（根到横线）
    mid_x = left + int(width / 2)
    line(slide, mid_x, top + root_h, mid_x, top + root_h + Inches(0.28), style["muted"], 1.5)
    # 横线
    if n > 1:
        first_cx = left + int(cw / 2)
        last_cx = left + (n - 1) * (cw + gap) + int(cw / 2)
        line(slide, first_cx, top + root_h + Inches(0.28), last_cx,
             top + root_h + Inches(0.28), style["muted"], 1.5)
    # 子节点 + 竖线
    for i, c in enumerate(children):
        cx = left + i * (cw + gap)
        line(slide, cx + int(cw / 2), top + root_h + Inches(0.28),
             cx + int(cw / 2), cy, style["muted"], 1.5)
        rect(slide, cx, cy, cw, child_h, style["panel"], round_=True,
             line=style["primary"])
        text(slide, cx, cy + Inches(0.1), cw, Inches(0.35),
             [(c, {"size": 11, "bold": True, "color": style["primary"]})],
             align=PP_ALIGN.CENTER)
    return cy + child_h


def native_fishbone(slide, effect, causes, left, top, width, height, style,
                    spine_color=None):
    """原生鱼骨图（因果图）—— 全矢量形状绘制。
    effect: 鱼头（问题）；causes: dict{类别: [原因...]}，如 {"人":[...], "机":[...]}。
    注：鱼骨图 PPT 无原生图表类型，此处用原生形状绘制，仍可编辑每个文本框。"""
    spine_color = spine_color or style["primary"]
    cy = top + int(height / 2)
    head_w = Inches(1.9)
    head_h = Inches(0.8)
    # 主脊线
    line(slide, left, cy, left + width - head_w, cy, spine_color, 3)
    # 鱼头
    head = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  left + width - head_w, cy - int(head_h / 2),
                                  head_w, head_h)
    head.fill.solid(); head.fill.fore_color.rgb = style["accent"]
    head.line.fill.background()
    text(slide, left + width - head_w, cy - Inches(0.16), head_w, Inches(0.4),
         [(effect, {"size": 12, "bold": True, "color": style["white"]})],
         align=PP_ALIGN.CENTER)
    # 支骨（上下交替）
    cat_items = list(causes.items())
    n = len(cat_items)
    if not n:
        return
    seg = (width - head_w - Inches(1.2)) / max(n, 1)
    for i, (cat, items) in enumerate(cat_items):
        bx = left + Inches(0.9) + int(seg * i) + int(seg / 2)
        up = (i % 2 == 0)
        if up:
            y_bone_end = top + Inches(0.35)
        else:
            y_bone_end = top + height - Inches(0.35)
        # 斜支骨
        line(slide, bx, cy, bx + Inches(0.6), y_bone_end, spine_color, 2)
        # 类别标签（骨端）
        cat_w = Inches(0.9)
        cat_h = Inches(0.36)
        cr = rect(slide, bx + Inches(0.6) - int(cat_w / 2),
                  y_bone_end - (cat_h if up else 0), cat_w, cat_h,
                  style["primary"], round_=True)
        text(slide, cr.left, cr.top + Inches(0.04), cat_w, Inches(0.28),
             [(cat, {"size": 10, "bold": True, "color": style["white"]})],
             align=PP_ALIGN.CENTER)
        # 原因小字
        for j, it in enumerate(items[:3]):
            t = text(slide, cr.left - Inches(0.45),
                     (y_bone_end + Inches(0.12) * j if up else y_bone_end - Inches(0.22) - Inches(0.12) * j),
                     Inches(1.8), Inches(0.24),
                     [(it, {"size": 9, "color": style["text"]})])
