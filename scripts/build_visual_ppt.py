#!/usr/bin/env python3
"""现代视觉版汇报 PPT 生成器（不依赖外部模板，自建版式）。

解决"汇报 PPT 全文字、丑"：用一套清爽的北航蓝设计系统重画整份 PPT——
封面/目录/章节分隔/内容卡/图表页/致谢每一页都有设计；文字要点做成卡片，
不再堆条目；研究框架/思路/方法/计划等关键页用 render_diagrams 画原生图形。

用法：
    python build_visual_ppt.py --content ppt_content.json --output 开题汇报-<题>.pptx

ppt_content.json 结构同 build_ppt_from_template.py（layout 支持：
cover/toc/section/stats/compare/table/flow/gantt/pipeline/bars/cards/
image_center/thanks；其中 cards 为文字卡片页，其余图表布局由
render_diagrams.draw 渲染）。
"""
from __future__ import annotations
import argparse, json
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from render_diagrams import (
    NAVY, BLUE, LIGHT, PALE, GRAY, LINEC, TEXT, WHITE, ACCENT,
    _apply_font, _shape, _fit_text, _wrap,
    draw as draw_diagram,
)

SW = 12192000          # 16:9 宽 (EMU)
SH = 6858000           # 16:9 高
IN = 914400            # 1 inch in EMU
_GRAY_TXT = RGBColor(0x8A, 0x8A, 0x8A)

HEADER_H = int(0.55 * IN)      # 顶部导航带
GOLD = int(0.09 * IN)          # 金色细线高
MARGIN = int(0.95 * IN)        # 侧边距
PAGE_TOP = HEADER_H + int(0.55 * IN)   # 标题区顶部


def _add_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    return s


def _band(slide):
    """顶部深蓝导航带 + 金色细线。"""
    _shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SW, HEADER_H, fill=NAVY, line=None)
    _shape(slide, MSO_SHAPE.RECTANGLE, 0, HEADER_H, SW, GOLD, fill=ACCENT, line=None)


def _footer(slide, deck_title, page_no):
    """底部细线 + 标题小字 + 页码。"""
    fy = SH - int(0.34 * IN)
    _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, fy - int(0.02 * IN),
           SW - 2 * MARGIN, Emu(2 * 13716), fill=LINEC, line=None)
    t = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, fy - int(0.2 * IN),
               int(6 * IN), int(0.28 * IN), fill=None, line=None)
    _fit_text(t, deck_title, 9, color=_GRAY_TXT, align=PP_ALIGN.LEFT, margin_pct=0.02)
    p = _shape(slide, MSO_SHAPE.RECTANGLE, SW - MARGIN - int(0.7 * IN),
               fy - int(0.2 * IN), int(0.7 * IN), int(0.28 * IN), fill=None, line=None)
    _fit_text(p, f"{page_no:02d}", 10, color=NAVY, bold=True, margin_pct=0.02)


def _header(slide, tag, title, page_no, deck_title):
    """内容页标题区：章节标签 + 观点标题 + 金色下划线。"""
    _band(slide)
    if tag:
        tg = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, PAGE_TOP - int(0.42 * IN),
                    int(4 * IN), int(0.26 * IN), fill=None, line=None)
        _fit_text(tg, tag.upper(), 10.5, color=BLUE, bold=True,
                  align=PP_ALIGN.LEFT, margin_pct=0.02)
    tt = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, PAGE_TOP,
                int(10.5 * IN), int(0.72 * IN), fill=None, line=None)
    _fit_text(tt, title, 25, color=NAVY, bold=True, align=PP_ALIGN.LEFT,
              margin_pct=0.02, anchor=MSO_ANCHOR.TOP)
    # 金色下划线
    _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, PAGE_TOP + int(0.68 * IN),
           int(1.4 * IN), Emu(4 * 13716), fill=ACCENT, line=None)
    _footer(slide, deck_title, page_no)


def _content_box():
    """标题下方的内容绘图区。"""
    top = PAGE_TOP + int(0.95 * IN)
    return (MARGIN, top, SW - 2 * MARGIN, SH - top - int(0.62 * IN))


# ── 各页类型 ──────────────────────────────────────────────────
def _draw_cover(slide, data, deck_title):
    _band(slide)
    # 顶部小字
    t0 = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, HEADER_H + int(0.4 * IN),
                int(8 * IN), int(0.3 * IN), fill=None, line=None)
    _fit_text(t0, "北京航空航天大学 · 硕士学位论文开题报告", 12, color=_GRAY_TXT,
              align=PP_ALIGN.LEFT, margin_pct=0.02)
    # 题目
    title = data.get("title", "")
    tt = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, int(2.1 * IN),
                int(10.6 * IN), int(1.7 * IN), fill=None, line=None)
    _fit_text(tt, title, 33, color=NAVY, bold=True, align=PP_ALIGN.LEFT,
              margin_pct=0.02, anchor=MSO_ANCHOR.TOP)
    # 副标题
    sub = data.get("subtitle", "")
    if sub:
        ts = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, int(3.75 * IN),
                    int(9 * IN), int(0.5 * IN), fill=None, line=None)
        _fit_text(ts, sub, 15, color=BLUE, align=PP_ALIGN.LEFT, margin_pct=0.02)
    # 金色线
    _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, int(4.35 * IN),
           int(2.2 * IN), Emu(5 * 13716), fill=ACCENT, line=None)
    # 底部信息格
    cover = data.get("cover", {})
    info = [
        ("汇报人", cover.get("作者姓名", "")),
        ("指导教师", cover.get("指导教师", "")),
        ("日期", cover.get("日期", "")),
        ("专业", cover.get("专业名称", "")),
    ]
    iw = (SW - 2 * MARGIN) / 4
    iy = SH - int(1.5 * IN)
    for i, (k, v) in enumerate(info):
        x = MARGIN + i * iw
        box = _shape(slide, MSO_SHAPE.RECTANGLE, x, iy, iw, int(0.9 * IN),
                     fill=None, line=None)
        text = f"{k}：{v}" if v else k
        _fit_text(box, text, 12, color=TEXT, align=PP_ALIGN.LEFT, margin_pct=0.06)
        # 信息格上方的分隔短竖线
        _shape(slide, MSO_SHAPE.RECTANGLE, x + int(0.12 * IN), iy - int(0.18 * IN),
               Emu(2 * 13716), int(0.18 * IN), fill=LINEC, line=None)
    _footer(slide, deck_title, 1)


def _draw_toc(slide, chapters, deck_title):
    _header(slide, "", "目录", 2, deck_title)
    items = chapters
    cols, rows = 2, (len(items) + 1) // 2
    box_l, box_t, box_w, box_h = _content_box()
    gap = int(0.28 * IN)
    cw = (box_w - gap) / cols
    ch = min((box_h - gap * (rows - 1)) / rows, int(1.15 * IN))
    for i, name in enumerate(items):
        r, c = divmod(i, cols)
        x = box_l + c * (cw + gap)
        y = box_t + r * (ch + gap)
        card = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, cw, ch,
                      fill=PALE, line=LINEC, line_w=1.0)
        try:
            card.adjustments[0] = 0.12
        except Exception:
            pass
        badge = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                       x + int(0.22 * IN), y + int(0.24 * IN),
                       int(0.62 * IN), int(0.62 * IN), fill=NAVY, line=None)
        try:
            badge.adjustments[0] = 0.25
        except Exception:
            pass
        _fit_text(badge, f"{i + 1:02d}", 16, color=WHITE, bold=True)
        nb = _shape(slide, MSO_SHAPE.RECTANGLE,
                    x + int(1.05 * IN), y, cw - int(1.2 * IN), ch,
                    fill=None, line=None)
        _fit_text(nb, name, 15, color=NAVY, bold=True, align=PP_ALIGN.LEFT,
                  margin_pct=0.05)


def _draw_section(slide, idx, name, deck_title, page_no):
    _band(slide)
    num = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, int(1.6 * IN),
                 int(4 * IN), int(2.2 * IN), fill=None, line=None)
    _fit_text(num, f"{idx:02d}", 100, color=LIGHT, bold=True,
              align=PP_ALIGN.LEFT, margin_pct=0.02, anchor=MSO_ANCHOR.TOP)
    nm = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, int(3.6 * IN),
                int(9 * IN), int(0.9 * IN), fill=None, line=None)
    _fit_text(nm, name, 30, color=NAVY, bold=True, align=PP_ALIGN.LEFT,
              margin_pct=0.02)
    _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, int(4.5 * IN),
           int(1.6 * IN), Emu(4 * 13716), fill=ACCENT, line=None)
    _footer(slide, deck_title, page_no)


def _draw_cards(slide, title, bullets, deck_title, page_no, tag=""):
    """文字要点卡片页：每条要点一张卡片，带序号徽章。"""
    _header(slide, tag, title, page_no, deck_title)
    box_l, box_t, box_w, box_h = _content_box()
    items = bullets[:5] or ["（本页要点待补充）"]
    gap = int(0.22 * IN)
    ch = min((box_h - gap * (len(items) - 1)) / len(items), int(1.05 * IN))
    for i, b in enumerate(items):
        y = box_t + i * (ch + gap)
        card = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                      box_l, y, box_w, ch, fill=PALE, line=LINEC, line_w=1.0)
        try:
            card.adjustments[0] = 0.10
        except Exception:
            pass
        badge = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                       box_l + int(0.25 * IN), y + int(0.26 * IN),
                       int(0.52 * IN), int(0.52 * IN), fill=NAVY, line=None)
        try:
            badge.adjustments[0] = 0.3
        except Exception:
            pass
        _fit_text(badge, f"{i + 1}", 14, color=WHITE, bold=True)
        tb = _shape(slide, MSO_SHAPE.RECTANGLE,
                    box_l + int(1.1 * IN), y, box_w - int(1.25 * IN), ch,
                    fill=None, line=None)
        _fit_text(tb, b, 13.5, color=TEXT, align=PP_ALIGN.LEFT,
                  margin_pct=0.06)


def _set_notes(slide, text):
    if text:
        try:
            slide.notes_slide.notes_text_frame.text = text
        except Exception:
            pass


def _draw_image(slide, title, image_path, caption, deck_title, page_no, tag=""):
    from PIL import Image
    import tempfile, os
    _header(slide, tag, title, page_no, deck_title)
    box_l, box_t, box_w, box_h = _content_box()
    try:
        im = Image.open(image_path)
        iw, ih = im.size
        max_w = box_w - int(0.4 * IN)
        max_h = box_h - (int(0.5 * IN) if caption else 0)
        scale = min(max_w / iw, max_h / ih, 1.0)
        disp_w = int(iw * scale)
        disp_h = int(ih * scale)
        px_w = int(disp_w / 914400 * 96)
        px_h = int(disp_h / 914400 * 96)
        if px_h > 0 and px_w > 0:
            im2 = im.resize((px_w, px_h), Image.LANCZOS)
            fd, tmp = tempfile.mkstemp(suffix=".png")
            os.close(fd)
            im2.save(tmp)
            x = box_l + (box_w - disp_w) // 2
            y = box_t + (box_h - disp_h - (int(0.4 * IN) if caption else 0)) // 2
            slide.shapes.add_picture(tmp, x, y, disp_w, disp_h)
            try:
                os.remove(tmp)
            except OSError:
                pass
        if caption:
            cb = _shape(slide, MSO_SHAPE.RECTANGLE, box_l, box_t + box_h - int(0.4 * IN),
                        box_w, int(0.35 * IN), fill=None, line=None)
            _fit_text(cb, caption, 10.5, color=_GRAY_TXT, align=PP_ALIGN.CENTER)
    except Exception:
        _draw_cards(slide, title, ["【图片加载失败，请在 PowerPoint 中补图】"],
                    deck_title, page_no, tag)


def _draw_thanks(slide, data, deck_title):
    _band(slide)
    t = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, int(2.3 * IN),
               SW - 2 * MARGIN, int(1.1 * IN), fill=None, line=None)
    _fit_text(t, "感谢聆听", 42, color=NAVY, bold=True)
    s = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, int(3.5 * IN),
               SW - 2 * MARGIN, int(0.5 * IN), fill=None, line=None)
    _fit_text(s, "恳请各位专家批评指正", 15, color=_GRAY_TXT)
    _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, int(4.15 * IN),
           int(2.2 * IN), Emu(4 * 13716), fill=ACCENT, line=None)
    cover = data.get("cover", {})
    who = _shape(slide, MSO_SHAPE.RECTANGLE, MARGIN, int(4.7 * IN),
                 SW - 2 * MARGIN, int(0.4 * IN), fill=None, line=None)
    _fit_text(who, f"汇报人：{cover.get('作者姓名', '')} · {cover.get('日期', '')}",
              13, color=TEXT)
    _footer(slide, deck_title, 1)


# ── 主流程 ────────────────────────────────────────────────────
DIAGRAM_LAYOUTS = {"stats", "compare", "table", "flow", "gantt",
                   "pipeline", "bars"}


def build(content_path: Path, output: Path):
    data = json.loads(content_path.read_text(encoding="utf-8"))
    deck_title = data.get("title", "")
    prs = Presentation()
    prs.slide_width = Emu(SW)
    prs.slide_height = Emu(SH)

    # 封面
    _draw_cover(_add_slide(prs), data, deck_title)

    # 目录（取各章规范名，去重）
    chapters = []
    for ch in data.get("chapters", []):
        n = ch.get("name", "")
        if n and n not in chapters:
            chapters.append(n)
    _draw_toc(_add_slide(prs), chapters, deck_title)

    page_no = 3
    for ci, ch in enumerate(data.get("chapters", [])):
        name = ch.get("name", "")
        slides = ch.get("slides", [])
        # 章节分隔页
        _draw_section(_add_slide(prs), ci + 1, name, deck_title, page_no)
        page_no += 1
        for spec in slides:
            slide = _add_slide(prs)
            layout = spec.get("layout", "text_only")
            title = spec.get("title", "") or name
            tag = name
            if layout in DIAGRAM_LAYOUTS:
                _header(slide, tag, title, page_no, deck_title)
                box = _content_box()
                if not draw_diagram(slide, layout, box, spec.get("extra") or {}):
                    _draw_cards(slide, title, spec.get("bullets", []),
                                deck_title, page_no, tag)
            elif layout == "image_center":
                extra = spec.get("extra") or {}
                _draw_image(slide, title, extra.get("image", ""),
                            extra.get("caption", ""), deck_title, page_no, tag)
            elif layout == "thanks":
                _draw_thanks(slide, data, deck_title)
            else:  # text / cards
                _draw_cards(slide, title, spec.get("bullets", []),
                            deck_title, page_no, tag)
            _set_notes(slide, spec.get("notes", ""))
            page_no += 1

    # 致谢
    _draw_thanks(_add_slide(prs), data, deck_title)

    prs.save(str(output))
    print("saved:", output)
    print(f"  页数: {len(prs.slides)}")


def main():
    ap = argparse.ArgumentParser(description="现代视觉版开题汇报 PPT 生成器")
    ap.add_argument("--content", required=True, type=Path)
    ap.add_argument("--output", required=True, type=Path)
    a = ap.parse_args()
    build(a.content, a.output)


if __name__ == "__main__":
    main()
