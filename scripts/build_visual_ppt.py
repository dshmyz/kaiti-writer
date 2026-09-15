#!/usr/bin/env python3
"""学术答辩风汇报 PPT 生成器（参考往届高分答辩稿版式，不依赖外部模板）。

版式语言：
- 封面 / 目录 / 章节分隔页：深蓝底（#0D3567 / #034C9B），白字，金色点缀
- 内容页：白底，顶部 PART 标签 + X.Y 编号小节标题 + 金色下划线
- 页脚：校名·学院·专业 + 页码 N / M
- 文字要点做成带序号徽章的卡片；研究框架/思路/方法/计划等关键页用
  render_diagrams.py 画原生图形（flow/gantt/compare/stats/table）

用法：
    python build_visual_ppt.py --content ppt_content.json --output 开题汇报-<题>.pptx

ppt_content.json 结构：
    title / subtitle / cover{作者姓名,指导教师,日期,专业名称,学院}
    school（默认 北京航空航天大学） / department（页脚学院）
    chapters: [{name, part, part_en, slides:[{title, layout, bullets, extra, notes}]}]
      - part: 所属 PART 名（如 "一、立题依据"）；part_en: 英文副题
      - layout: stats/compare/table/flow/gantt/pipeline/bars/cards/image_center
"""
from __future__ import annotations
import argparse, json
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from render_diagrams import (
    NAVY, BLUE, LIGHT, PALE, GRAY, LINEC, TEXT, WHITE, ACCENT,
    _shape, _fit_text, _wrap,
    draw as draw_diagram,
)

# ── 配色（参考稿：深蓝 + 白 + 金）────────────────────────────
DEEP   = RGBColor(0x0D, 0x35, 0x67)   # 封面深蓝底
MAIN   = RGBColor(0x03, 0x4C, 0x9B)   # 主蓝（章节分隔/标签/标题）
GOLD   = RGBColor(0xC8, 0x9A, 0x4B)   # 金色点缀
INK    = RGBColor(0x1F, 0x2D, 0x3D)   # 深色正文
MUTE   = RGBColor(0x6B, 0x7A, 0x8C)   # 次要灰
PAPER  = RGBColor(0xFF, 0xFF, 0xFF)   # 内容页底

SW = 12192000
SH = 6858000
IN = 914400
MARGIN = int(0.95 * IN)
CONTENT_TOP = int(1.25 * IN)


def _slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _box(shape, left, top, width, height):
    shape.left, shape.top = Emu(left), Emu(top)
    shape.width, shape.height = Emu(width), Emu(height)


def _bg(slide, color):
    from pptx.enum.shapes import MSO_SHAPE
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Emu(SW), Emu(SH))
    r.fill.solid(); r.fill.fore_color.rgb = color
    r.line.fill.background(); r.shadow.inherit = False
    return r


def _text(slide, text, pt, color, *, left, top, width=10000000, height=600000,
          bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, margin=0.02):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(left), Emu(top),
                                 Emu(width), Emu(height))
    box.fill.background(); box.line.fill.background(); box.shadow.inherit = False
    _fit_text(box, text, pt, color=color, bold=bold, align=align,
              margin_pct=margin, anchor=anchor)
    return box


def _gold_line(slide, x, y, w=1000000, h=50000):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y),
                                 Emu(w), Emu(h))
    bar.fill.solid(); bar.fill.fore_color.rgb = GOLD
    bar.line.fill.background(); bar.shadow.inherit = False


def _content_box():
    return (MARGIN, CONTENT_TOP + int(0.5 * IN), SW - 2 * MARGIN,
            SH - CONTENT_TOP - int(0.95 * IN))


# ── 深蓝页（封面/目录/章节分隔/致谢）─────────────────────────
def _draw_cover(slide, data, deck_title):
    _bg(slide, DEEP)
    # 顶部左：校名 · 学院
    _text(slide, f"{data.get('school', '北京航空航天大学')} · {data.get('cover', {}).get('培养学院', '公共管理学院')}",
          12, RGBColor(0xDD, 0xE7, 0xF2), left=MARGIN, top=int(0.55 * IN))
    # 顶部右：开题汇报 · 硕士论文开题
    _text(slide, "开题汇报 · 硕士学位论文开题", 12, RGBColor(0xDD, 0xE7, 0xF2),
          left=SW - MARGIN - int(4 * IN), top=int(0.55 * IN), width=int(4 * IN),
          align=PP_ALIGN.RIGHT)
    # 题目
    _text(slide, data.get("title", ""), 33, WHITE, bold=True,
          left=MARGIN, top=int(2.15 * IN), width=int(10.4 * IN), height=int(1.6 * IN))
    # 副题
    sub = data.get("subtitle", "")
    if sub:
        _text(slide, sub, 16, RGBColor(0xBD, 0xD3, 0xEA),
              left=MARGIN, top=int(3.7 * IN), width=int(9 * IN))
    # 金色线
    _gold_line(slide, MARGIN, int(4.28 * IN), w=int(2.4 * IN), h=int(0.055 * IN))
    # 底部信息
    cover = data.get("cover", {})
    info = [
        ("汇报人", cover.get("作者姓名", "")),
        ("指导教师", cover.get("指导教师", "")),
        ("专业", cover.get("专业名称", "")),
        ("日期", cover.get("日期", "")),
    ]
    iw = (SW - 2 * MARGIN) / 4
    iy = SH - int(1.5 * IN)
    for i, (k, v) in enumerate(info):
        x = int(MARGIN + i * iw)
        _text(slide, f"{k}：{v}" if v else k, 12.5, RGBColor(0xE6, 0xEE, 0xF6),
              left=x, top=iy, width=int(iw))
        if i > 0:
            sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x - int(0.15 * IN)),
                                         Emu(iy + int(0.08 * IN)), Emu(2 * 13716),
                                         Emu(int(0.3 * IN)))
            sep.fill.solid(); sep.fill.fore_color.rgb = RGBColor(0x2A, 0x4E, 0x7F)
            sep.line.fill.background(); sep.shadow.inherit = False


def _draw_toc(slide, parts, deck_title):
    _bg(slide, MAIN)
    _text(slide, "目录", 30, WHITE, bold=True, left=MARGIN, top=int(1.0 * IN))
    _text(slide, "CONTENTS", 13, RGBColor(0x9E, 0xBE, 0xDE),
          left=MARGIN, top=int(1.62 * IN))
    _gold_line(slide, MARGIN, int(2.0 * IN), w=int(1.6 * IN), h=int(0.05 * IN))
    rows = len(parts)
    row_h = int(0.92 * IN)
    gap = int(0.3 * IN)
    start_y = int(2.35 * IN)
    for i, (name, en) in enumerate(parts):
        y = start_y + i * (row_h + gap)
        num = _text(slide, f"{i + 1:02d}", 22, RGBColor(0x9E, 0xBE, 0xDE), bold=True,
                    left=MARGIN, top=y, width=int(1.0 * IN))
        nm = _text(slide, name, 17, WHITE, bold=True,
                   left=MARGIN + int(1.1 * IN), top=y, width=int(6 * IN))
        if en:
            _text(slide, en, 11, RGBColor(0x9E, 0xBE, 0xDE),
                  left=SW - MARGIN - int(5 * IN), top=y + int(0.12 * IN),
                  width=int(5 * IN), align=PP_ALIGN.RIGHT)


def _draw_section(slide, part_no, name, en, deck_title, date):
    _bg(slide, MAIN)
    _text(slide, f"PART {part_no:02d}", 20, RGBColor(0x9E, 0xBE, 0xDE), bold=True,
          left=MARGIN, top=int(1.6 * IN))
    _text(slide, name, 34, WHITE, bold=True, left=MARGIN, top=int(2.2 * IN),
          width=int(10 * IN), height=int(1.0 * IN))
    if en:
        _text(slide, en, 15, RGBColor(0xBD, 0xD3, 0xEA),
              left=MARGIN, top=int(3.3 * IN), width=int(10 * IN))
    _gold_line(slide, MARGIN, int(4.0 * IN), w=int(2.2 * IN), h=int(0.055 * IN))
    if date:
        _text(slide, date, 12, RGBColor(0x9E, 0xBE, 0xDE),
              left=SW - MARGIN - int(3 * IN), top=SH - int(0.9 * IN),
              width=int(3 * IN), align=PP_ALIGN.RIGHT)


# ── 内容页（白底）─────────────────────────────────────────────
def _content_header(slide, part_no, sub_no, title, page_no, total, footer):
    # PART 标签
    _text(slide, f"PART {part_no:02d}", 10.5, MAIN, bold=True,
          left=MARGIN, top=int(0.5 * IN))
    # 小节标题（X.Y 标题）
    _text(slide, f"{part_no}.{sub_no}  {title}", 23, INK, bold=True,
          left=MARGIN, top=int(0.82 * IN), width=int(11 * IN), height=int(0.72 * IN))
    _gold_line(slide, MARGIN, int(1.58 * IN), w=int(1.2 * IN), h=int(0.045 * IN))
    # 页脚
    _footer(slide, footer, page_no, total)


def _footer(slide, footer, page_no, total):
    fy = SH - int(0.5 * IN)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(MARGIN), Emu(fy - int(0.16 * IN)),
                                  Emu(SW - 2 * MARGIN), Emu(2 * 13716))
    line.fill.solid(); line.fill.fore_color.rgb = RGBColor(0xD8, 0xE0, 0xE8)
    line.line.fill.background(); line.shadow.inherit = False
    if footer:
        _text(slide, footer, 9.5, MUTE, left=MARGIN, top=fy - int(0.38 * IN),
              width=int(9 * IN))
    _text(slide, f"{page_no} / {total}", 10, MAIN, bold=True,
          left=SW - MARGIN - int(1.2 * IN), top=fy - int(0.38 * IN),
          width=int(1.2 * IN), align=PP_ALIGN.RIGHT)


def _draw_cards(slide, bullets, part_no, sub_no, title, page_no, total, footer):
    _content_header(slide, part_no, sub_no, title, page_no, total, footer)
    box_l, box_t, box_w, box_h = _content_box()
    items = bullets[:5] or ["（本页要点待补充）"]
    gap = int(0.22 * IN)
    ch = min((box_h - gap * (len(items) - 1)) / len(items), int(1.0 * IN))
    for i, b in enumerate(items):
        y = int(box_t + i * (ch + gap))
        card = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                      box_l, y, box_w, int(ch), fill=PALE, line=LINEC, line_w=1.0)
        try:
            card.adjustments[0] = 0.09
        except Exception:
            pass
        badge = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                       box_l + int(0.25 * IN), y + int(0.22 * IN),
                       int(0.5 * IN), int(0.5 * IN), fill=MAIN, line=None)
        try:
            badge.adjustments[0] = 0.3
        except Exception:
            pass
        _fit_text(badge, f"{i + 1}", 13, color=WHITE, bold=True)
        tb = _shape(slide, MSO_SHAPE.RECTANGLE,
                    box_l + int(1.05 * IN), y, box_w - int(1.2 * IN), int(ch),
                    fill=None, line=None)
        _fit_text(tb, b, 13, color=INK, align=PP_ALIGN.LEFT, margin_pct=0.06)


def _draw_image(slide, image_path, caption, part_no, sub_no, title, page_no,
                total, footer):
    from PIL import Image
    import tempfile, os
    _content_header(slide, part_no, sub_no, title, page_no, total, footer)
    box_l, box_t, box_w, box_h = _content_box()
    try:
        im = Image.open(image_path)
        iw, ih = im.size
        max_w = box_w - int(0.4 * IN)
        max_h = box_h - (int(0.5 * IN) if caption else 0)
        scale = min(max_w / iw, max_h / ih, 1.0)
        disp_w = int(iw * scale); disp_h = int(ih * scale)
        px_w = int(disp_w / 914400 * 96); px_h = int(disp_h / 914400 * 96)
        if px_h > 0 and px_w > 0:
            fd, tmp = tempfile.mkstemp(suffix=".png"); os.close(fd)
            im.resize((px_w, px_h), Image.LANCZOS).save(tmp)
            x = box_l + (box_w - disp_w) // 2
            y = box_t + (box_h - disp_h - (int(0.4 * IN) if caption else 0)) // 2
            slide.shapes.add_picture(tmp, x, y, disp_w, disp_h)
            try:
                os.remove(tmp)
            except OSError:
                pass
        if caption:
            _text(slide, caption, 10.5, MUTE, left=box_l, top=box_t + box_h - int(0.42 * IN),
                  width=box_w, align=PP_ALIGN.CENTER)
    except Exception:
        _draw_cards(slide, ["【图片加载失败，请在 PowerPoint 中补图】"],
                    part_no, sub_no, title, page_no, total, footer)


def _draw_thanks(slide, data):
    _bg(slide, DEEP)
    _text(slide, "汇报完毕", 40, WHITE, bold=True, left=MARGIN,
          top=int(2.3 * IN), width=SW - 2 * MARGIN, align=PP_ALIGN.CENTER)
    _text(slide, "恳请各位老师批评指正！", 16, RGBColor(0xBD, 0xD3, 0xEA),
          left=MARGIN, top=int(3.4 * IN), width=SW - 2 * MARGIN, align=PP_ALIGN.CENTER)
    _gold_line(slide, int(SW / 2 - 1.1 * IN), int(4.1 * IN), w=int(2.2 * IN), h=int(0.055 * IN))
    cover = data.get("cover", {})
    _text(slide, f"汇报人：{cover.get('作者姓名', '')} · {cover.get('日期', '')}",
          13, RGBColor(0xE6, 0xEE, 0xF6), left=MARGIN, top=int(4.6 * IN),
          width=SW - 2 * MARGIN, align=PP_ALIGN.CENTER)


# ── 主流程 ────────────────────────────────────────────────────
DIAGRAM_LAYOUTS = {"stats", "compare", "table", "flow", "gantt",
                   "pipeline", "bars"}


def build(content_path: Path, output: Path):
    data = json.loads(content_path.read_text(encoding="utf-8"))
    deck_title = data.get("title", "")
    cover = data.get("cover", {})
    school = data.get("school", "北京航空航天大学")
    dept = cover.get("培养学院", "公共管理学院")
    major = cover.get("专业名称", "公共管理")
    footer = f"{school} · {dept} · {major}"

    prs = Presentation()
    prs.slide_width = Emu(SW)
    prs.slide_height = Emu(SH)

    # 封面
    _draw_cover(_slide(prs), data, deck_title)

    # 目录（按 part 分组）
    parts = []
    for ch in data.get("chapters", []):
        key = (ch.get("part", "一、研究概述"), ch.get("part_en", ""))
        if key not in parts:
            parts.append(key)
    _draw_toc(_slide(prs), parts, deck_title)

    page_no = 3
    total = 3 + len(parts) + sum(len(ch.get("slides", [])) for ch in data["chapters"]) + 1
    for pi, (pname, pen) in enumerate(parts):
        pno = pi + 1
        # 章节分隔页
        date = cover.get("日期", "")
        _draw_section(_slide(prs), pno, pname, pen, deck_title, date)
        page_no += 1
        sub_no = 0
        for ch in data.get("chapters", []):
            if (ch.get("part", "一、研究概述") != pname):
                continue
            for spec in ch.get("slides", []):
                slide = _slide(prs)
                layout = spec.get("layout", "text_only")
                title = spec.get("title", "") or ch.get("name", "")
                sub_no += 1
                if layout in DIAGRAM_LAYOUTS:
                    _content_header(slide, pno, sub_no, title, page_no, total, footer)
                    box = _content_box()
                    if not draw_diagram(slide, layout, box, spec.get("extra") or {}):
                        _draw_cards(slide, spec.get("bullets", []), pno, sub_no,
                                    title, page_no, total, footer)
                elif layout == "image_center":
                    extra = spec.get("extra") or {}
                    _draw_image(slide, extra.get("image", ""), extra.get("caption", ""),
                                pno, sub_no, title, page_no, total, footer)
                else:
                    _draw_cards(slide, spec.get("bullets", []), pno, sub_no,
                                title, page_no, total, footer)
                if spec.get("notes"):
                    try:
                        slide.notes_slide.notes_text_frame.text = spec["notes"]
                    except Exception:
                        pass
                page_no += 1

    # 致谢
    _draw_thanks(_slide(prs), data)

    prs.save(str(output))
    print("saved:", output)
    print(f"  页数: {len(prs.slides)}")


def main():
    ap = argparse.ArgumentParser(description="学术答辩风开题汇报 PPT 生成器")
    ap.add_argument("--content", required=True, type=Path)
    ap.add_argument("--output", required=True, type=Path)
    a = ap.parse_args()
    build(a.content, a.output)


if __name__ == "__main__":
    main()
