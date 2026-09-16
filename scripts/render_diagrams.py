#!/usr/bin/env python3
"""原生绘制 PPT 图表（python-pptx 形状，可编辑，零新增依赖）。

解决"PPT 全是文字"的问题：把研究框架/技术路线/实施计划/方法对比/数据页
从文字 bullets 变成真正的图形——流程图、阶段条、甘特时间轴、大数字卡片、
对比表、柱状图。全部用 pptx 原生形状/图表，在 PowerPoint 里可编辑；
配色固定为北航蓝体系。

对外接口（供 build_ppt_from_template.py 调用）：
    draw(slide, layout, box, data)  → 按 layout 分派
box = (left, top, width, height)，单位 EMU。data 结构见各 draw_* 函数 docstring。
"""
from __future__ import annotations
from dataclasses import dataclass
import re

from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LABEL_POSITION

# ── 北航蓝配色体系 ─────────────────────────────────────────────
NAVY   = RGBColor(0x00, 0x33, 0x66)   # 主色·深蓝
BLUE   = RGBColor(0x00, 0x5B, 0xAC)   # 次主色·亮蓝
LIGHT  = RGBColor(0xDC, 0xE9, 0xF7)   # 浅蓝底
PALE   = RGBColor(0xEE, 0xF4, 0xFB)   # 极浅底
GRAY   = RGBColor(0xF2, 0xF2, 0xF2)   # 中性灰底
LINEC  = RGBColor(0xB9, 0xC8, 0xD9)   # 分隔线
TEXT   = RGBColor(0x2B, 0x2B, 0x2B)   # 正文深灰
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT = RGBColor(0xD9, 0x7E, 0x00)   # 强调·暖橙（仅少量点缀）

EA_FONT = "微软雅黑"
LATIN_FONT = "Calibri"

EMU_PER_PT = 12700


# ── 基础工具 ───────────────────────────────────────────────────
def _apply_font(run, pt, bold=False, color=TEXT, ea=EA_FONT, latin=LATIN_FONT):
    from pptx.oxml.ns import qn
    from lxml import etree
    f = run.font
    f.size = Pt(pt)
    f.bold = bold
    f.color.rgb = color
    f.name = latin
    rpr = run._r.get_or_add_rPr()
    latin_el = rpr.find(qn("a:latin"))
    if latin_el is None:
        latin_el = etree.SubElement(rpr, qn("a:latin"))
    latin_el.set("typeface", latin)
    ea_el = rpr.find(qn("a:ea"))
    if ea_el is None:
        ea_el = etree.SubElement(rpr, qn("a:ea"))
    ea_el.set("typeface", ea)


def _shape(slide, kind, left, top, width, height, fill=LIGHT, line=NAVY,
           line_w=1.25):
    """建形状并统一去掉默认阴影、设置填充与边框。"""
    sp = slide.shapes.add_shape(kind, Emu(left), Emu(top), Emu(width), Emu(height))
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    return sp


def _char_w(char: str, pt: float) -> float:
    """单字符在给定字号下的近似宽度（EMU）。全角≈1em，半角≈0.55em。"""
    w = 1.0 if ord(char) > 0x2E80 or char in "（）：；，。！？、￥%《》" else 0.55
    return w * pt * EMU_PER_PT


def _wrap(text: str, pt: float, max_w: float) -> list[str]:
    """按可用宽度逐字换行（中文无空格，不能按词断）。保留显式换行。"""
    lines: list[str] = []
    for para in text.split("\n"):
        if not para:
            lines.append("")
            continue
        cur, cur_w = "", 0.0
        for ch in para:
            cw = _char_w(ch, pt)
            if cur and cur_w + cw > max_w:
                lines.append(cur)
                cur, cur_w = ch, cw
            else:
                cur += ch
                cur_w += cw
        if cur:
            lines.append(cur)
    return lines or [""]


def _fit_text(shape, text, pt, color=TEXT, bold=False, min_pt=7.5,
              align=PP_ALIGN.CENTER, line_spacing=1.14, margin_pct=0.08,
              anchor=MSO_ANCHOR.MIDDLE):
    """把文字写进形状：溢出则逐步缩字号；返回最终字号。"""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    ins = int(min(shape.width, shape.height) * margin_pct)
    tf.margin_left = tf.margin_right = ins
    tf.margin_top = tf.margin_bottom = max(Emu(2 * EMU_PER_PT), ins // 2)

    inner_w = shape.width - 2 * ins
    inner_h = shape.height - 2 * max(Emu(2 * EMU_PER_PT), ins // 2)

    def lines_for(p):
        return _wrap(text, p, inner_w)

    def height_for(lines, p):
        return len(lines) * p * line_spacing * EMU_PER_PT * 1.25

    lines = lines_for(pt)
    while (height_for(lines, pt) > inner_h and pt > min_pt):
        pt -= 0.5
        lines = lines_for(pt)

    tf.clear()
    for i, ln in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.line_spacing = line_spacing
        para.space_before = Pt(0)
        para.space_after = Pt(0)
        run = para.add_run()
        run.text = ln or " "
        _apply_font(run, pt, bold=bold, color=color)
    return pt


def _connector(slide, x1, y1, x2, y2, color=NAVY, width_pt=1.6,
               arrow=True):
    """直线连接线；arrow 时加实心三角箭头。"""
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                   Emu(x1), Emu(y1), Emu(x2), Emu(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width_pt)
    c.shadow.inherit = False
    if arrow:
        from pptx.oxml.ns import qn
        from lxml import etree
        ln = c.line._get_or_add_ln()
        for child in list(ln):
            if child.tag == qn("a:tailEnd"):
                ln.remove(child)
        tail = etree.SubElement(ln, qn("a:tailEnd"))
        tail.set("type", "triangle")
        tail.set("w", "med")
        tail.set("len", "med")
    return c


def _box_h_for_text(texts: list[str], pt: float, box_w: float,
                    line_spacing: float = 1.14) -> float:
    """估算一组文本放同一行所需的最小高度。"""
    inner_w = box_w * (1 - 2 * 0.08)
    max_lines = max(len(_wrap(t, pt, inner_w)) for t in texts)
    return max_lines * pt * line_spacing * EMU_PER_PT * 1.25 + 6 * EMU_PER_PT


# ── 分派入口 ───────────────────────────────────────────────────
def draw(slide, layout: str, box, data: dict | None):
    """按 layout 在 box 内绘图。data 为 None 或空则返回 False（调用方降级为文字页）。"""
    if not data:
        return False
    left, top, width, height = box
    if layout == "flow":
        return _draw_flow(slide, left, top, width, height, data)
    if layout == "pipeline":
        return _draw_pipeline(slide, left, top, width, height, data)
    if layout == "gantt":
        return _draw_gantt(slide, left, top, width, height, data)
    if layout == "stats":
        return _draw_stats(slide, left, top, width, height, data)
    if layout == "cards":
        return _draw_cards(slide, left, top, width, height, data)
    if layout == "panels":
        return _draw_panels(slide, left, top, width, height, data)
    if layout == "compare":
        return _draw_table(slide, left, top, width, height, data, first_col_label=True)
    if layout == "table":
        return _draw_table(slide, left, top, width, height, data, first_col_label=False)
    if layout == "bars":
        return _draw_bars(slide, left, top, width, height, data)
    return False


# ── 流程图（研究框架 / 技术路线）──────────────────────────────
def _normalize_flow_nodes(nodes) -> list:
    """nodes 元素：字符串=单框；数组=并列框。支持 "a→b→c" 文本拆分。"""
    out = []
    for n in nodes:
        if isinstance(n, str):
            if "→" in n:
                for part in n.split("→"):
                    part = part.strip()
                    if part:
                        out.append(part)
            else:
                out.append(n)
        else:
            out.append([str(x).strip() for x in n if str(x).strip()])
    return out


def _draw_flow(slide, left, top, width, height, data) -> bool:
    nodes = _normalize_flow_nodes(data.get("nodes") or data.get("steps") or [])
    if not nodes:
        return False
    direction = data.get("direction", "vertical")
    if direction == "horizontal" and all(isinstance(n, str) for n in nodes):
        return _draw_flow_h(slide, left, top, width, height, nodes)
    return _draw_flow_v(slide, left, top, width, height, nodes)


def _draw_flow_v(slide, left, top, width, height, nodes) -> bool:
    m = int(0.05 * height)
    gap = int(min(max(0.12 * height, 12 * EMU_PER_PT), 26 * EMU_PER_PT))
    n = len(nodes)
    usable_h = height - 2 * m
    layer_h = (usable_h - gap * (n - 1)) / n
    if layer_h < 30 * EMU_PER_PT:
        layer_h = 30 * EMU_PER_PT
    pt = 13.0

    y = top + m
    for li, node in enumerate(nodes):
        items = node if isinstance(node, list) else [node]
        nw = len(items)
        gap_x = int(0.03 * width)
        bw = (width - 2 * m - gap_x * (nw - 1)) / nw
        for xi, item in enumerate(items):
            bx = left + m + xi * (bw + gap_x)
            box = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                         bx, y, bw, layer_h, fill=LIGHT, line=NAVY, line_w=1.4)
            try:
                box.adjustments[0] = 0.12
            except Exception:
                pass
            _fit_text(box, item, pt, color=NAVY, bold=True)
        # 层间箭头
        if li < n - 1:
            cx = left + width // 2
            _connector(slide, cx, int(y + layer_h), cx, int(y + layer_h + gap),
                       color=NAVY, arrow=True)
        y += layer_h + gap
    return True


def _draw_flow_h(slide, left, top, width, height, nodes) -> bool:
    m = int(0.12 * height)
    gap = int(min(max(0.10 * width, 18 * EMU_PER_PT), 40 * EMU_PER_PT))
    n = len(nodes)
    usable_w = width - 2 * m
    bw = (usable_w - gap * (n - 1)) / n
    box_h = height - 2 * m
    pt = 13.0

    x = left + m
    for item in nodes:
        box = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                     x, top + m, bw, box_h, fill=LIGHT, line=NAVY, line_w=1.4)
        try:
            box.adjustments[0] = 0.14
        except Exception:
            pass
        _fit_text(box, item, pt, color=NAVY, bold=True)
        x += bw + gap
    for i in range(n - 1):
        x1 = left + m + (i + 1) * bw + i * gap
        cy = top + height // 2
        _connector(slide, x1, cy, x1 + gap, cy, color=NAVY, arrow=True)
    return True


# ── 阶段条（阶段一 → 阶段二 → 阶段三）────────────────────────
def _draw_pipeline(slide, left, top, width, height, data) -> bool:
    stages = data.get("stages") or data.get("items") or []
    if isinstance(stages, dict):
        stages = list(stages.values())
    stages = [s if isinstance(s, dict) else {"label": str(s)} for s in stages]
    if not stages:
        return False

    m = int(0.06 * height)
    n = len(stages)
    overlap = int(0.10 * width / n)          # 箭头槽重叠
    bw = (width - 2 * m + overlap * (n - 1)) / n
    box_h = height - 2 * m

    colors = _interp_colors(NAVY, BLUE, n)
    x = left + m
    for i, st in enumerate(stages):
        sp = _shape(slide, MSO_SHAPE.CHEVRON, x, top + m, bw, box_h,
                    fill=colors[i], line=None)
        try:
            sp.adjustments[0] = 0.28
        except Exception:
            pass
        tf = sp.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        ins = Emu(10 * EMU_PER_PT)
        tf.margin_left = tf.margin_right = ins
        tf.margin_top = tf.margin_bottom = Emu(2 * EMU_PER_PT)
        label = st.get("label", "")
        desc = st.get("desc", "")
        inner_w = sp.width - 2 * ins
        # 标签 14pt 白、说明 9.5pt 白（若放得下）
        lines = _wrap(label, 14, inner_w)
        if desc:
            desc_lines = _wrap(desc, 9.5, inner_w)
            if len(lines) + len(desc_lines) <= 3:
                lines = lines + desc_lines
                label_pt, desc_pt = 14.0, 9.5
            else:
                desc_pt = None
        else:
            desc_pt = None
        total_pt = 14.0
        inner_h = box_h - 4 * EMU_PER_PT
        while (len(lines) * total_pt * 1.14 * EMU_PER_PT * 1.25 > inner_h
               and total_pt > 9):
            total_pt -= 0.5
            lines = _wrap(label, total_pt, inner_w)
        tf.clear()
        for j, ln in enumerate(lines):
            para = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            para.alignment = PP_ALIGN.CENTER
            para.line_spacing = 1.05
            run = para.add_run()
            run.text = ln or " "
            if desc_pt and j == len(lines) - 1 and desc and ln in desc_lines:
                _apply_font(run, desc_pt, color=WHITE)
            else:
                _apply_font(run, total_pt, bold=True, color=WHITE)
        x += bw - overlap
    return True


def _interp_colors(c1, c2, n):
    if n <= 1:
        return [c1]
    out = []
    for i in range(n):
        t = i / (n - 1)
        out.append(RGBColor(
            round(c1[0] + (c2[0] - c1[0]) * t),
            round(c1[1] + (c2[1] - c1[1]) * t),
            round(c1[2] + (c2[2] - c1[2]) * t)))
    return out


# ── 甘特时间轴（实施计划）─────────────────────────────────────
_MONTH_RE = re.compile(r"(\d{4})[.\-/年](\d{1,2})")


def _parse_month(s) -> tuple[int, int] | None:
    m = _MONTH_RE.search(str(s))
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def _month_idx(s) -> int | None:
    p = _parse_month(s)
    return p[0] * 12 + p[1] if p else None


def _draw_gantt(slide, left, top, width, height, data) -> bool:
    tasks = data.get("tasks") or data.get("rows") or []
    if not tasks:
        return False
    # 兼容 plan_table 形态：{阶段/内容/时间} 或 [阶段,内容,时间]
    norm = []
    for t in tasks:
        if isinstance(t, (list, tuple)):
            t = {"phase": t[0], "content": t[1] if len(t) > 1 else "", "time": t[-1]}
        phase = t.get("phase") or t.get("阶段") or ""
        content = t.get("content") or t.get("内容") or ""
        time = t.get("time") or t.get("时间") or ""
        start = _month_idx(time.split("–")[0].split("-")[0].split("—")[0])
        end_part = time
        for sep in ("–", "-", "—", "~", "至"):
            if sep in time:
                end_part = time.split(sep)[-1]
                break
        end = _month_idx(end_part)
        if start is None or end is None:
            continue
        norm.append({"phase": phase, "content": content, "time": time,
                     "start": start, "end": max(start, end)})
    if not norm:
        return False

    label_w = int(0.24 * width)
    axis_h = int(max(0.09 * height, 18 * EMU_PER_PT))
    m = int(0.03 * height)
    row_h = (height - m - axis_h) / len(norm)
    if row_h < 24 * EMU_PER_PT:
        row_h = 24 * EMU_PER_PT

    # 时间轴范围：任务最值，两端各留 1 个月
    lo = min(t["start"] for t in norm) - 1
    hi = max(t["end"] for t in norm) + 1
    span = max(hi - lo, 1)
    plot_x0 = left + label_w
    plot_w = (left + width) - plot_x0

    def x_for(idx):
        return plot_x0 + (idx - lo) / span * plot_w

    # 月份刻度（步长≥1）
    step = max(1, span // max(6, int(plot_w / (30 * EMU_PER_PT))))
    # 网格竖线 + 顶部月份标签
    idx = lo
    while idx <= hi:
        if (idx - lo) % step == 0:
            gx = int(x_for(idx))
            _connector(slide, gx, top + axis_h, gx, top + height - m,
                       color=LINEC, width_pt=0.75, arrow=False)
            ym, mo = divmod(idx - 1, 12)
            lbl = f"{ym % 100:02d}.{mo + 1:02d}"
            tbox = _shape(slide, MSO_SHAPE.RECTANGLE, gx - int(22 * EMU_PER_PT),
                          top, int(44 * EMU_PER_PT), axis_h, fill=None, line=None)
            _fit_text(tbox, lbl, 9, color=TEXT, bold=False)
        idx += 1

    # 任务行：左侧阶段名 + 右侧色条
    for i, t in enumerate(norm):
        ry = top + axis_h + m + i * row_h
        # 阶段名（标签列）
        lbl_box = _shape(slide, MSO_SHAPE.RECTANGLE, left, ry,
                         label_w - int(4 * EMU_PER_PT), row_h, fill=None, line=None)
        _fit_text(lbl_box, t["phase"], 11, color=NAVY, bold=True,
                  align=PP_ALIGN.LEFT, margin_pct=0.05, anchor=MSO_ANCHOR.MIDDLE)
        # 内容小字（阶段名下）
        if t.get("content"):
            cbox = _shape(slide, MSO_SHAPE.RECTANGLE, left + int(0.30 * label_w),
                          ry, int(0.68 * label_w), row_h, fill=None, line=None)
            _fit_text(cbox, t["content"], 9, color=RGBColor(0x8A, 0x8A, 0x8A),
                      align=PP_ALIGN.LEFT, margin_pct=0.04,
                      anchor=MSO_ANCHOR.MIDDLE)
        # 色条
        bx1 = x_for(t["start"])
        bx2 = x_for(t["end"])
        bar = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                     int(bx1), int(ry + row_h * 0.18),
                     int(max(bx2 - bx1, 6 * EMU_PER_PT)), int(row_h * 0.64),
                     fill=BLUE if i % 2 == 0 else NAVY, line=None)
        try:
            bar.adjustments[0] = 0.35
        except Exception:
            pass
        if (bx2 - bx1) > int(36 * EMU_PER_PT):
            _fit_text(bar, t["phase"], 9.5, color=WHITE, bold=True)
    return True


# ── 大数字卡片（数据页）───────────────────────────────────────
def _draw_stats(slide, left, top, width, height, data) -> bool:
    stats = data.get("stats") or data.get("items") or []
    if isinstance(stats, dict):
        stats = list(stats.values())
    stats = [s if isinstance(s, dict) else {"number": str(s)} for s in stats]
    if not stats:
        return False
    n = len(stats)
    cols = 2 if n <= 4 else 3
    rows = (n + cols - 1) // cols
    m = int(0.03 * width)
    gap = int(0.03 * width)
    card_w = (width - 2 * m - gap * (cols - 1)) / cols
    card_h = (height - gap * (rows - 1)) / rows

    for i, st in enumerate(stats):
        r, c = divmod(i, cols)
        cx = left + m + c * (card_w + gap)
        cy = top + gap * r + r * card_h
        card = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                      cx, cy, card_w, card_h, fill=PALE, line=LINEC, line_w=1.0)
        try:
            card.adjustments[0] = 0.06
        except Exception:
            pass
        number = str(st.get("number", ""))
        label = str(st.get("label", ""))
        num_color = ACCENT if i == 0 and number else NAVY
        # 数字区占 55% 高，标签区占 45%
        num_h = int(card_h * 0.55)
        nbox = _shape(slide, MSO_SHAPE.RECTANGLE, cx, cy, card_w, num_h,
                      fill=None, line=None)
        _fit_text(nbox, number, 34, color=num_color, bold=True)
        if label:
            lbox = _shape(slide, MSO_SHAPE.RECTANGLE, cx, cy + num_h,
                          card_w, card_h - num_h, fill=None, line=None)
            _fit_text(lbox, label, 11, color=TEXT, bold=False,
                      align=PP_ALIGN.CENTER)
    return True


# ── 编号卡片（文字要点也有视觉焦点，参考高分答辩稿的容器化排版）────
_IN = 914400


def _draw_cards(slide, left, top, width, height, data) -> bool:
    """要点卡片页：每条要点一张圆角卡 + 序号徽章。条目多时自动压紧凑，一页装下。"""
    items = data.get("items") or data.get("cards") or data.get("bullets") or []
    items = [str(i) for i in items if str(i).strip()]
    if not items:
        return False
    n = len(items)
    # 条目多 → 更小的间距与卡片高度（自适应，不拆页）
    gap = int(0.18 * _IN) if n <= 5 else int(0.10 * _IN)
    ch = (height - gap * (n - 1)) / n
    if ch < int(0.30 * _IN):
        ch = int(0.30 * _IN)          # 兜底最小高度（极少触发）
    font_pt = 13.0 if n <= 5 else 11.0
    for i, b in enumerate(items):
        y = int(top + i * (ch + gap))
        card = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, y, width, int(ch),
                      fill=PALE, line=LINEC, line_w=1.0)
        try:
            card.adjustments[0] = 0.09
        except Exception:
            pass
        badge = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                       left + int(0.22 * _IN), y + int(0.16 * _IN),
                       int(0.42 * _IN), int(0.42 * _IN), fill=NAVY, line=None)
        try:
            badge.adjustments[0] = 0.3
        except Exception:
            pass
        _fit_text(badge, f"{i + 1}", font_pt, color=WHITE, bold=True)
        tb = _shape(slide, MSO_SHAPE.RECTANGLE,
                    left + int(0.92 * _IN), y, width - int(1.08 * _IN), int(ch),
                    fill=None, line=None)
        _fit_text(tb, b, font_pt, color=TEXT, align=PP_ALIGN.LEFT, margin_pct=0.05)
    return True


# ── 双栏/多栏面板（意义/成果类：理论意义 | 现实意义 等，与卡片页形成版式差异）──
def _draw_panels(slide, left, top, width, height, data) -> bool:
    """大面板页：2–4 块并排面板，每块=标题栏 + 正文，用于"理论意义/现实意义"类内容。"""
    panels = data.get("panels") or data.get("items") or []
    if isinstance(panels, dict):
        panels = list(panels.values())
    panels = [p if isinstance(p, dict) else {"title": str(p), "text": ""} for p in panels]
    panels = panels[:4]
    if not panels:
        return False
    n = len(panels)
    gap = int(0.035 * width)
    pw = (width - gap * (n - 1)) / n
    ph = height
    title_h = int(max(0.16 * ph, 30 * _IN * 0.07))
    for i, p in enumerate(panels):
        x = int(left + i * (pw + gap))
        panel = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, top, int(pw), int(ph),
                       fill=PALE, line=LINEC, line_w=1.0)
        try:
            panel.adjustments[0] = 0.05
        except Exception:
            pass
        # 标题栏（上色块）
        tb = _shape(slide, MSO_SHAPE.RECTANGLE, x, top, int(pw), int(title_h),
                    fill=NAVY, line=None)
        _fit_text(tb, str(p.get("title", "")), 14, color=WHITE, bold=True)
        # 正文
        bb = _shape(slide, MSO_SHAPE.RECTANGLE,
                    x + int(0.06 * pw), top + int(title_h + 0.04 * pw),
                    int(0.88 * pw), int(ph - title_h - 0.08 * pw), fill=None, line=None)
        _fit_text(bb, str(p.get("text", "")), 11.5, color=TEXT,
                  align=PP_ALIGN.LEFT, margin_pct=0.04)
    return True


# ── 对比表（方法对比 / 文献对比 / 通用表）─────────────────────
def _draw_table(slide, left, top, width, height, data, first_col_label=False) -> bool:
    headers = data.get("headers") or data.get("columns") or []
    rows = data.get("rows") or []
    if not headers or not rows:
        return False
    ncols = len(headers)
    # 校验行宽，补齐
    rows = [r + [""] * (ncols - len(r)) for r in rows if r][:7]

    header_h = int(max(0.13 * height, 26 * EMU_PER_PT))
    row_h = (height - header_h) / len(rows)
    if row_h < 24 * EMU_PER_PT:
        row_h = 24 * EMU_PER_PT

    col_w = width / ncols
    x = left
    # 表头
    for j, h in enumerate(headers):
        hb = _shape(slide, MSO_SHAPE.RECTANGLE, x, top, col_w, header_h,
                    fill=NAVY, line=None)
        _fit_text(hb, str(h), 12, color=WHITE, bold=True)
        x += col_w
    # 数据行
    y = top + header_h
    for i, row in enumerate(rows):
        fill = PALE if i % 2 == 0 else GRAY
        for j, cell in enumerate(row):
            bold = (first_col_label and j == 0)
            cfill = fill
            if bold:
                cfill = LIGHT
            cb = _shape(slide, MSO_SHAPE.RECTANGLE,
                        left + j * col_w, y, col_w, row_h,
                        fill=cfill, line=WHITE, line_w=1.0)
            _fit_text(cb, str(cell), 10.5, color=NAVY if bold else TEXT,
                      bold=bold)
        y += row_h
    # 外框
    from pptx.oxml.ns import qn
    from lxml import etree
    for bj in range(ncols):
        for ri in range(len(rows) + 1):
            pass
    return True


# ── 柱状图（数据趋势）─────────────────────────────────────────
def _draw_bars(slide, left, top, width, height, data) -> bool:
    categories = data.get("categories") or []
    values = data.get("values") or []
    if not categories or not values:
        return False
    # 兼容 {"x": [...], "y": [...]}
    if not categories and data.get("x"):
        categories = data["x"]
    if not values and data.get("y"):
        values = data["y"]

    cd = CategoryChartData()
    cd.categories = [str(c) for c in categories]
    cd.add_series(data.get("series") or "数据", [float(v) for v in values])

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, cd)
    chart = chart_frame.chart
    chart.has_legend = False
    try:
        plot = chart.plots[0]
        plot.gap_width = 120
        plot.overlap = 0
        if not plot.has_data_labels:
            plot.has_data_labels = True
        dls = plot.data_labels
        dls.show_value = True
        dls.number_format = "0.#"
        dls.number_format_is_linked = False
        dls.position = XL_LABEL_POSITION.OUTSIDE_END
        dls.font.size = Pt(10)
        dls.font.color.rgb = NAVY
        dls.font.bold = True
        ser = plot.series[0]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = BLUE
    except Exception:
        pass
    # 坐标轴
    for axis in (chart.category_axis, chart.value_axis):
        try:
            axis.tick_labels.font.size = Pt(10)
            axis.tick_labels.font.color.rgb = TEXT
            axis.format.line.color.rgb = LINEC
        except Exception:
            pass
    try:
        chart.value_axis.has_major_gridlines = True
        chart.value_axis.major_gridlines.format.line.color.rgb = GRAY
    except Exception:
        pass
    return True


if __name__ == "__main__":
    import sys
    print("render_diagrams.py 是库模块，请由 build_ppt_from_template.py 调用。")
    print("支持 layout: flow / pipeline / gantt / stats / compare / table / bars")
    sys.exit(0)
