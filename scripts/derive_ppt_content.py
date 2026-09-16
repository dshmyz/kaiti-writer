#!/usr/bin/env python3
"""从 content.json 自动派生 ppt_content.json，省去手工重复写 PPT 内容。

用法：
    python derive_ppt_content.py --content content.json --output ppt_content.json

布局自动识别（优先顺序）：
- 「实施计划」节（配合顶层 plan_table）→ gantt 甘特时间轴
- 「研究思路 / 研究框架 / 技术路线」节（含 → 链或顶层 route.nodes）→ flow 流程图
- 「研究方法」节 → compare 方法对比表
- 含多个百分比的数据页 → stats 大数字卡片
- 含「阶段一/第一阶段/前期中期」等 → pipeline 阶段条
- 其余 → text_only 文字要点页（超过 5 条自动拆两页）

图表页的 `extra` 字段即 render_diagrams.py 各 draw_* 的 data 结构；
`bullets` 保留原文要点，供演讲者备注与降级兜底用。
"""
import argparse
import json
import re
from pathlib import Path

# 章节规范顺序：每个元组 = (规范名, 匹配 token 列表)。键名写法不一（"文献综述"/
# "国内外研究现状"、"研究框架（内容）"），按含有的 token 归位。
_SECTION_TOKENS = [
    ("研究背景", ["研究背景", "问题的提出", "选题背景"]),
    ("选题意义", ["选题意义", "研究意义"]),
    ("文献综述", ["文献综述", "国内外研究现状", "国内外现状", "研究现状"]),
    ("研究框架", ["研究框架", "分析框架"]),
    ("研究思路", ["研究思路"]),
    ("研究方法", ["研究方法"]),
    ("创新之处", ["创新之处", "创新点"]),
    ("论文大纲", ["论文大纲", "内容安排", "写作安排"]),
    ("实施计划", ["实施计划", "时间安排", "进度安排", "研究计划"]),
    ("预期成果", ["预期目标", "预期成果", "预期目标和成果"]),
]


def _section_rank(key: str) -> int:
    for i, (canon, toks) in enumerate(_SECTION_TOKENS):
        if any(t in key for t in toks):
            return i
    return len(_SECTION_TOKENS)


def clean_section_name(key: str) -> str:
    """（一）研究背景 → 研究背景；命中规范 token 则用规范名。"""
    n = re.sub(r"^[（(一二三四五六七八九十]+[）)]\s*", "", key).strip()
    for canon, toks in _SECTION_TOKENS:
        if any(t in key or t in n for t in toks):
            return canon
    return n


def extract_bullets(items: list, max_bullets: int = 6) -> tuple[list, list]:
    """从 content_by_section 的值数组中提取文本 bullets 和特殊布局页。

    返回 (bullets, special_slides)：
    - bullets: 普通文字要点（达到上限后停止收集，但不影响后续 dict 块）
    - special_slides: 需要特殊布局的页（图片/图表/表格）
    """
    bullets = []
    special_slides = []
    bullets_done = False
    for item in items:
        if isinstance(item, str):
            if bullets_done:
                continue
            text = item.strip()
            if not text:
                continue
            if len(text) > 60:
                parts = re.split(r"[。；]", text)
                for p in parts:
                    p = p.strip()
                    if p and len(p) > 5:
                        bullets.append(p)
                        if len(bullets) >= max_bullets:
                            bullets_done = True
                            break
            else:
                bullets.append(text)
                if len(bullets) >= max_bullets:
                    bullets_done = True
        elif isinstance(item, dict):
            if "image" in item:
                caption = item.get("caption", "技术路线图")
                image_path = item.get("image", "")
                special_slides.append({
                    "title": caption, "layout": "image_center",
                    "bullets": [],
                    "extra": {"image": image_path, "caption": caption}
                })
            elif "table" in item:
                caption = item.get("caption", "")
                headers = item["table"].get("headers", [])
                rows = item["table"].get("rows", [])
                special_slides.append({
                    "title": caption or " ".join(headers[:3]),
                    "layout": "table",
                    "bullets": [],
                    "extra": {"table_data": {"headers": headers, "rows": rows}}
                })
            elif "list" in item and not bullets_done:
                for li in item["list"][:3]:
                    bullets.append(f"● {li}")
                    if len(bullets) >= max_bullets:
                        bullets_done = True
                        break
            elif "chart" in item:
                chart_type = item["chart"].get("type", "bar")
                chart_data = item["chart"].get("data", {})
                special_slides.append({
                    "title": item.get("caption", "数据图表"),
                    "layout": "chart",
                    "bullets": [],
                    "extra": {"chart_type": chart_type, "chart_data": chart_data}
                })
    return bullets, special_slides


# ── 布局识别 ───────────────────────────────────────────────────
def _is_plan(key: str) -> bool:
    return any(k in key for k in ("实施计划", "时间安排", "进度", "计划安排"))


def _is_route(key: str) -> bool:
    return any(k in key for k in ("研究思路", "研究框架", "技术路线", "研究方案"))


def _is_method(key: str) -> bool:
    return "方法" in key and "研究思路" not in key


def _route_nodes(content: dict, items: list) -> list | None:
    """提取流程图节点：顶层 route.nodes → 「→」链 → 框架段落里的链条描述。"""
    route = content.get("route") or {}
    nodes = route.get("nodes")
    if nodes:
        return nodes
    for item in items:
        if isinstance(item, str) and "→" in item:
            parts = [p.strip() for p in item.split("→") if p.strip()]
            if len(parts) >= 3:
                return parts
    # 框架段落含「沿……链条」描述：按「→」或关键连接词拆
    joined = "".join(i for i in items if isinstance(i, str))
    if "→" in joined:
        parts = [p.strip() for p in joined.split("→") if p.strip()]
        if len(parts) >= 3:
            return parts
    return None


def _plan_gantt(content: dict, items: list) -> dict | None:
    """顶层 plan_table 或节内「阶段/内容/时间」表块 → gantt 数据。"""
    pt = content.get("plan_table") or {}
    rows = pt.get("rows") or []
    if not rows:
        for it in items:
            if isinstance(it, dict) and "table" in it:
                hdrs = it["table"].get("headers") or []
                if any("时间" in h or "阶段" in h for h in hdrs):
                    rows = it["table"].get("rows") or []
                    break
    if not rows:
        return None
    tasks = []
    for r in rows:
        if isinstance(r, (list, tuple)) and len(r) >= 2:
            tasks.append({
                "phase": str(r[0]),
                "content": str(r[1]) if len(r) > 2 else "",
                "time": str(r[2] if len(r) > 2 else r[1]),
            })
        elif isinstance(r, dict):
            tasks.append({
                "phase": str(r.get("阶段") or r.get("phase") or ""),
                "content": str(r.get("内容") or r.get("content") or ""),
                "time": str(r.get("时间") or r.get("time") or ""),
            })
    return {"tasks": tasks}


def _method_rows(bullets: list) -> list | None:
    """方法节 → 对比表行：拆「N、方法名。说明」或「方法名：说明」。不足两行返回 None。"""
    rows = []
    for b in bullets:
        b = re.sub(r"^[●•]\s*", "", b).strip()
        b = re.sub(r"^\d+\s*[、.)]\s*", "", b)   # 去掉 "1、" 序号
        parts = re.split(r"[：:]", b, 1)
        if len(parts) == 2:
            name, desc = parts[0].strip(), parts[1].strip()
        else:
            parts = re.split(r"[。；]", b, 1)
            if len(parts) == 2:
                name, desc = parts[0].strip(), parts[1].strip()
            else:
                continue
        # 过滤非方法名（如"上述方法中，…"结尾句）
        if not name or len(name) > 14 or "上述" in name:
            continue
        desc = re.split(r"[。；]", desc)[0][:40]
        rows.append([name, desc])
    return rows if len(rows) >= 2 else None


def _find_table_items(items: list) -> list:
    """返回 items 里的 table 块字典列表。"""
    return [it for it in items if isinstance(it, dict) and "table" in it]


_PANEL_PREFIXES = ("理论意义", "现实意义", "前沿性", "必要性",
                   "预期目标", "预期成果", "风险预案", "预期")


def _panels_data(items: list) -> dict | None:
    """意义/成果类：拆「理论意义：…」式前缀 → 双栏面板。不足两栏返回 None。"""
    panels = []
    for it in items:
        if not isinstance(it, str):
            continue
        t = it.strip()
        m = re.match(r"^([一-龥A-Za-z]{2,8}[：:])\s*(.*)$", t)
        if not m:
            continue
        title = m.group(1)[:-1]
        if not any(title.startswith(p) or p.startswith(title) for p in _PANEL_PREFIXES):
            continue
        text = m.group(2).strip()
        if len(text) > 56:
            text = text[:56] + "…"
        if title and text:
            panels.append({"title": title, "text": text})
    return {"panels": panels} if len(panels) >= 2 else None


def _outline_table(items: list, bullets: list) -> dict | None:
    """论文大纲节：拆「第一章 绪论：…」→ 章节|内容 表。优先从原文 items 解析
    （bullets 被切碎且截断，会丢后面的章节），bullets 兜底。不足三章返回 None。"""

    def parse(source):
        rows = []
        for b in source:
            if not isinstance(b, str):
                continue
            m = re.match(r"^(第[一二三四五六七八九十]+章)\s*([^：:：]*?)[：:]\s*(.*)$", b.strip())
            if m:
                rows.append([m.group(1) + m.group(2), m.group(3)[:32]])
        return rows

    rows = parse(items) or parse(bullets)
    return {"headers": ["章节", "内容"], "rows": rows} if len(rows) >= 3 else None


def _method_rows_from_items(items: list) -> list | None:
    """从方法节原文 items 拆对比表行（bullets 已被切碎，须用原文）。"""
    rows = []
    for it in items:
        if not isinstance(it, str):
            continue
        b = re.sub(r"^\d+\s*[、.)]\s*", "", it.strip())
        parts = re.split(r"[。；]", b, 1)
        if len(parts) == 2:
            name, desc = parts[0].strip(), parts[1].strip()
        else:
            parts = re.split(r"[：:]", b, 1)
            if len(parts) == 2:
                name, desc = parts[0].strip(), parts[1].strip()
            else:
                continue
        if not name or len(name) > 14 or "上述" in name:
            continue
        desc = re.split(r"[。；]", desc)[0][:40]
        rows.append([name, desc])
    return rows if len(rows) >= 2 else None


def _stats_data(bullets: list) -> dict | None:
    """含 ≥2 个百分比的段落 → 大数字卡片数据。"""
    stats = []
    for b in bullets:
        for m in re.finditer(r"(\d{1,3}(?:\.\d)?)\s*%", b):
            label = re.sub(r"(\d{1,3}(?:\.\d)?)\s*%", "‖", b).strip("，。；、 ")
            label = label.split("‖")[0].strip()
            if len(label) > 18:
                label = label[:18] + "…"
            stats.append({"number": f"{m.group(1)}%", "label": label})
            if len(stats) >= 4:
                break
        if len(stats) >= 4:
            break
    return {"stats": stats} if len(stats) >= 2 else None


_STAGE_RE = re.compile(r"(第[一二三四五六]阶段|阶段[一二三四五六]|前期|中期|后期|步骤[一二三四五])")
_STAGE_NO = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6}


def _stages_data(bullets: list) -> dict | None:
    """含「阶段/步骤」标记 → 阶段条数据。"""
    stages = []
    seen = set()
    for b in bullets:
        m = _STAGE_RE.search(b)
        if not m:
            continue
        tag = m.group(0)
        key = _STAGE_NO.get(tag[-1], len(stages) + 1)
        if key in seen:
            continue
        seen.add(key)
        label = tag
        desc = b.replace(tag, "", 1).lstrip("：:： ").strip("，。；、 ")
        desc = re.sub(r"^[●•]\s*", "", desc)
        if len(desc) > 20:
            desc = desc[:20] + "…"
        stages.append({"label": label, "desc": desc})
    if not stages:
        return None
    stages.sort(key=lambda s: _STAGE_NO.get(s["label"][-1], 99))
    return {"stages": stages}


def _make_slide(title: str, bullets: list, layout: str, extra=None,
                notes_bullets=None) -> dict:
    return {"title": title, "bullets": bullets, "layout": layout,
            "extra": extra}


def _build_chapter_slides(key: str, name: str, bullets: list, special_slides: list,
                          content: dict, items: list) -> list:
    """按优先级决定本节的页面布局。返回 slides 列表。"""
    slides = []

    if _is_plan(key):
        g = _plan_gantt(content, items)
        if g:
            slides.append(_make_slide("实施计划与时间安排", bullets, "gantt", g))
            # 该阶段的表块已转成甘特，避免再重复一张表格页
            special_slides = [sp for sp in special_slides
                              if sp.get("layout") != "table"]
            slides.extend(special_slides)
            return slides

    if _is_route(key):
        # 研究框架自带表格（如三链×三层矩阵）→ 用表格页，避免与思路页重复 flowchart
        if "研究框架" in key:
            tables = _find_table_items(items)
            if tables:
                t = tables[0]["table"]
                slides.append(_make_slide(
                    name, bullets, "compare",
                    {"headers": t.get("headers"), "rows": t.get("rows")}))
                special_slides = [sp for sp in special_slides
                                  if sp.get("layout") != "table"]
                slides.extend(special_slides)
                return slides
        nodes = _route_nodes(content, items)
        if nodes:
            flat = all(isinstance(n, str) for n in nodes) and len(nodes) <= 6
            slides.append(_make_slide(
                name, bullets, "flow",
                {"nodes": nodes, "direction": "horizontal" if flat else "vertical"}))
            slides.extend(special_slides)
            return slides

    if _is_method(key):
        rows = _method_rows_from_items(items) or _method_rows(bullets)
        if rows:
            slides.append(_make_slide(
                "研究方法", bullets, "compare",
                {"headers": ["研究方法", "做法要点"], "rows": rows}))
            slides.extend(special_slides)
            return slides

    if bullets:
        stats = _stats_data(bullets)
        if stats:
            slides.append(_make_slide(name, bullets, "stats", stats))
            slides.extend(special_slides)
            return slides
        stages = _stages_data(bullets)
        if stages:
            slides.append(_make_slide(name, bullets, "pipeline", stages))
            slides.extend(special_slides)
            return slides

    # 兜底：按内容挑版式，不千篇一律——意义/成果带前缀 → 双栏面板；
    # 论文大纲 → 章节表格；其余 → 编号卡片（**一页装下全部**，渲染层自动压紧凑，
    # 不拆页——换大纲/稿子也不乱结构）
    panels = _panels_data(items)
    if panels:
        slides.append(_make_slide(name, bullets, "panels", panels))
        slides.extend(special_slides)
        return slides
    outline = _outline_table(items, bullets)
    if outline:
        slides.append(_make_slide(name, bullets, "compare", outline))
        slides.extend(special_slides)
        return slides
    if bullets:
        trimmed = [b if len(b) <= 26 else b[:26] + "…" for b in bullets]
        slides.append(_make_slide(name, trimmed, "cards", {"items": trimmed}))
    slides.extend(special_slides)
    return slides


# ── 演讲者备注 ─────────────────────────────────────────────────
def _generate_notes(chapter_name: str, slide: dict) -> str:
    title = slide.get("title", "")
    bullets = slide.get("bullets", [])
    layout = slide.get("layout", "text_only")
    notes_parts = []

    if "背景" in chapter_name or "意义" in chapter_name:
        notes_parts.append("开场：用具体案例/数据引出问题，不要从宏观政策开始。")
    elif "文献" in chapter_name:
        notes_parts.append("过渡语：前面讲了问题，现在看看别人怎么做的，有什么不足。")
    elif "框架" in chapter_name or "思路" in chapter_name:
        notes_parts.append("过渡语：基于文献不足，我的研究思路是…")
    elif "方法" in chapter_name:
        notes_parts.append("过渡语：具体怎么做？用这三种方法。")
    elif "创新" in chapter_name:
        notes_parts.append("重点：这是评审最关注的页，讲清楚新在哪里。")
    elif "计划" in chapter_name:
        notes_parts.append("收尾：时间节点清晰，让评审觉得可行。")

    if layout == "gantt":
        notes_parts.append("指向时间轴：先讲整体周期，再讲几个关键节点，不逐行念。")
    elif layout == "flow":
        notes_parts.append("指向流程图：沿箭头讲清研究推进的逻辑链，每框一句话。")
    elif layout == "pipeline":
        notes_parts.append("指向阶段条：讲阶段划分与衔接，每阶段一句话。")
    elif layout == "compare":
        notes_parts.append("指向表格：对比差异，不要逐格念。")
    elif layout == "stats":
        notes_parts.append("指向大数字：每个数字配一句说明，突出关键结论。")
    elif layout == "chart":
        notes_parts.append("指向图表：重点讲数据趋势，不要逐个读数字。")
    elif layout == "image_center":
        notes_parts.append("指向图片：解释图中关键要素，说明其与研究的关系。")
    elif layout == "table":
        notes_parts.append("指向表格：对比差异，不要逐行念。")
    elif len(bullets) > 3:
        notes_parts.append("要点较多，挑重点讲，其余让评审自己看。")

    if "创新" in title:
        notes_parts.append("强调：每个创新点用一句话说清楚。")
    elif "实施" in title or "计划" in title:
        notes_parts.append("结尾：时间节点明确，展示可行性。")

    return "\n".join(notes_parts) if notes_parts else ""


# ── 主流程 ─────────────────────────────────────────────────────
def derive(content: dict) -> dict:
    ppt = {
        "title": content.get("title", ""),
        "subtitle": "开题汇报",
        "cover": content.get("cover", {}),
        "chapters": [],
    }

    sections = content.get("content_by_section", {})
    ordered_keys = sorted(sections, key=_section_rank)

    for key in ordered_keys:
        items = sections[key]
        name = clean_section_name(key)
        bullets, special_slides = extract_bullets(items)
        slides = _build_chapter_slides(key, name, bullets, special_slides,
                                       content, items)
        if not slides:
            continue
        for slide in slides:
            slide["notes"] = _generate_notes(name, slide)
        ppt["chapters"].append({"name": name, "slides": slides})

    return ppt


def main():
    ap = argparse.ArgumentParser(description="从 content.json 派生 ppt_content.json")
    ap.add_argument("--content", required=True, type=Path)
    ap.add_argument("--output", required=True, type=Path)
    a = ap.parse_args()

    content = json.loads(a.content.read_text(encoding="utf-8"))
    # 同目录的 route.json 若存在则并入，供研究思路/框架识别 flow（思路文本常用"—"非"→"）
    route_path = a.content.parent / "route.json"
    if route_path.is_file() and "route" not in content:
        try:
            content["route"] = json.loads(route_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    ppt = derive(content)
    a.output.write_text(json.dumps(ppt, ensure_ascii=False, indent=2), encoding="utf-8")
    layouts = [s["layout"] for ch in ppt["chapters"] for s in ch["slides"]]
    print(f"saved: {a.output} ({len(ppt['chapters'])} chapters, "
          f"{len(layouts)} 页：{' / '.join(sorted(set(layouts)))})")


if __name__ == "__main__":
    main()
