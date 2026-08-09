#!/usr/bin/env python3
"""生成技术路线图 PNG（供开题报告「（一）研究思路」插图用）。

零新增依赖：只用 Pillow（已装）+ 系统宋体，不需要 graphviz / mermaid-cli / matplotlib。
生成后用 build_from_template.py 的 image 块嵌入：
    {"image": "路线图.png", "caption": "本研究技术路线图"}

用法：
    python make_route_figure.py --content route.json --output 路线图.png

route.json：
    {
      "nodes": [
        "问题提出：西城区社区行政负担的结构与成因",
        ["台账编码（内容分析）", "深度访谈 + 问卷"],
        "负担结构测量与成因分析",
        "减负对策建议与研究结论"
      ]
    }
nodes 元素为字符串 = 单个框；为数组 = 同一层的并列分支（左右并排，共用上下箭头）。
"""
import argparse
import json
from pathlib import Path

CJK_FONTS = [
    "/System/Library/Fonts/Supplemental/Songti.ttc",   # macOS 宋体
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/usr/share/fonts/truetype/arphic/uming.ttc",      # Linux
    "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
    "C:/Windows/Fonts/simsun.ttc",                     # Windows 宋体
]


def load_font(size):
    from PIL import ImageFont
    for path in CJK_FONTS:
        if Path(path).is_file():
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    raise SystemExit(
        "找不到可用的中文字体，无法生成路线图。\n"
        "  替代方案：自己用任意工具画好图，再用 image 块嵌入：\n"
        '    {"image": "路线图.png", "caption": "本研究技术路线图"}'
    )


def wrap(text, max_w, font, measure):
    """按像素宽度逐字换行（中文没有空格，不能按词断）。"""
    lines, cur = [], ""
    for ch in text:
        if ch == "\n":
            lines.append(cur); cur = ""; continue
        if cur and measure(cur + ch, font) > max_w:
            lines.append(cur); cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    return lines or [""]


def build(nodes, *, font_size=26, box_width=560, margin=30, gap=52,
          pad=16, radius=10, line_w=2, branch_gap=30):
    from PIL import Image, ImageDraw

    font = load_font(font_size)
    probe = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    measure = lambda s, f: probe.textlength(s, font=f)
    line_h = font_size + 10

    # 先排版算尺寸，再画（一趟量、一趟画）
    layout = []
    for node in nodes:
        if isinstance(node, (list, tuple)):
            bw = (box_width - branch_gap * (len(node) - 1)) // len(node)
            wrapped = [wrap(str(s), bw - 2 * pad, font, measure) for s in node]
            h = max(len(w) for w in wrapped) * line_h + 2 * pad
            layout.append(("branch", wrapped, bw, h))
        else:
            wrapped = wrap(str(node), box_width - 2 * pad, font, measure)
            layout.append(("box", [wrapped], box_width, len(wrapped) * line_h + 2 * pad))

    total_h = 2 * margin + sum(h for *_, h in layout) + gap * (len(layout) - 1)
    total_w = box_width + 2 * margin
    im = Image.new("RGB", (total_w, total_h), "white")
    d = ImageDraw.Draw(im)
    cx = margin + box_width // 2

    y = margin
    for i, (kind, wrapped, bw, h) in enumerate(layout):
        for j, lines in enumerate(wrapped):
            x = margin + j * (bw + branch_gap) if kind == "branch" else margin
            d.rounded_rectangle([x, y, x + bw, y + h], radius,
                                outline="black", width=line_w)
            ty = y + pad + (h - 2 * pad - len(lines) * line_h) // 2
            for k, ln in enumerate(lines):
                d.text((x + (bw - measure(ln, font)) / 2, ty + k * line_h),
                       ln, font=font, fill="black")
        y += h
        if i < len(layout) - 1:                      # 层间竖线箭头
            d.line([cx, y, cx, y + gap - 12], fill="black", width=line_w)
            d.polygon([(cx - 7, y + gap - 16), (cx, y + gap - 2),
                       (cx + 7, y + gap - 16)], fill="black")
            y += gap
    return im


def main():
    ap = argparse.ArgumentParser(description="生成技术路线图 PNG")
    ap.add_argument("--content", required=True, help="节点描述 json")
    ap.add_argument("--output", required=True, help="输出 png")
    ap.add_argument("--width-cm", type=float, default=14.0,
                    help="嵌入 docx 后的显示宽度（默认 14cm，不超版心）")
    ap.add_argument("--dpi", type=int, default=200, help="输出分辨率（默认 200）")
    args = ap.parse_args()

    data = json.loads(Path(args.content).expanduser().read_text("utf-8"))
    nodes = data.get("nodes") or []
    if not nodes:
        raise SystemExit("content json 里 nodes 为空，没有可画的节点。")

    im = build(nodes,
               font_size=data.get("font_size", 26),
               box_width=data.get("box_width", 560))

    # 放大到目标物理宽度，保证嵌入 docx 后清晰（image_extent 按像素/dpi 折算 cm）
    target_px = int(args.width_cm / 2.54 * args.dpi)
    if target_px > im.width:
        from PIL import Image
        im = im.resize((target_px, round(im.height * target_px / im.width)),
                       Image.LANCZOS)

    out = Path(args.output).expanduser()
    im.save(out, dpi=(args.dpi, args.dpi))
    disp_w = im.width / args.dpi * 2.54
    disp_h = im.height / args.dpi * 2.54
    print(f"saved: {out}")
    print(f"  尺寸: {im.width}x{im.height}px @ {args.dpi}dpi"
          f"（嵌入后约 {disp_w:.1f}cm × {disp_h:.1f}cm）")
    print(f"  节点: {len(nodes)} 层")
    if disp_h > 20.0:
        print(f"  ⚠️ 高 {disp_h:.1f}cm 可能超出一页版心（A4 约 24cm）。"
              "可减少层数、把并列项合到一层，或调小 box_width/font_size。")
    print(f'  嵌入写法: {{"image": "{out}", "caption": "本研究技术路线图"}}')


if __name__ == "__main__":
    main()
