# 视觉自检与平台渲染

> 交付前需要逐页检查版面时查阅。
代码跑通 ≠ 版面正确。至少藏着 1–3 处"数据逻辑错位"或"标签压字"，只有渲染才看得见。

**macOS**（Keynote 最稳）：
```bash
osascript -e 'tell application "Keynote"
  activate
  set d to open POSIX file "/abs/deck.pptx"
  delay 3
  export d to POSIX file "/abs/preview.pdf" as PDF
  close d saving no
end tell'
```

**Windows**（PowerPoint COM）：
```powershell
$pp = New-Object -ComObject PowerPoint.Application
$p = $pp.Presentations.Open("C:\abs\deck.pptx", $true, $false, $false)
$p.SaveAs("C:\abs\preview.pdf", 32)
$p.Close(); $pp.Quit()
```

**Linux / 三平台通用**（LibreOffice）：
```bash
soffice --headless --convert-to pdf --outdir . deck.pptx
```

之后统一转 PNG 并逐页检查：
```bash
python3 -c "import pymupdf; d=pymupdf.open('preview.pdf'); [p.get_pixmap(dpi=110).save(f'preview/s{i+1:02d}.png') for i,p in enumerate(d)]"
```
> Windows 上 python 命令通常是 `python` 而非 `python3`，路径分隔符用 `\\` 或 `/` 均可；
> 若没装 pymupdf：`pip install pymupdf`。

必查：文字溢出/裁切、元素重叠、tofu 方框、图表标签压柱、图例遮挡、页码缺失、**配色与主题是否统一**。

### 平台差异速查

| 项目 | macOS | Windows | Linux |
|------|-------|---------|-------|
| 中文字体 | PingFang SC | Microsoft YaHei | Noto Sans CJK SC |
| 字体覆盖 | 环境变量 `PPT_CN_FONT` 可强制指定 | 同左 | 同左 |
| PPT→PDF | Keynote (osascript) | PowerPoint COM | LibreOffice (soffice) |
| python 命令 | `python3` | `python` | `python3` |

字体由 `_detect_cn_font()` 自动检测，跨平台兜底时**不会**选其他平台专有字体
（Windows 上不会回退成 PingFang，Linux 上不会回退成微软雅黑）。
