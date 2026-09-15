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

# content_by_section 的键 → PPT chapter 名的映射（按常见顺序）
SECTION_ORDER = [
    "（一）研究背景",
    "（二）选题意义",
    "（三）国内外研究现状",
    "研究框架（内容）",
    "（一）研究思路",
    "（二）研究方法",
    "（三）创新之处",
    "四、学位论文实施计划",
    "五、预期目标和成果",
]

SECTION_NAME_ALIAS = {
    "国内外研究现状": "文献综述",
    "研究框架（内容）": "研究框架",
    "（一）研究思路": "研究思路",
    "（二）研究方法": "研究方法",
    "（三）创新之处": "创新之处",
    "四、学位论文实施计划": "实施计划",
    "五、预期目标和成果": "预期成果",
}


def clean_section_name(key: str) -> str:
    """（一）研究背景 → 研究背景"""
    n = re.sub(r"^[（(一二三四五六七八九十]+[）)]\s*", "", key).strip()
    return SECTION_NAME_ALIAS.get(n, n)


def extract_bullets(items: list, max_bullets: int = 6) -> tuple[list, list]:
    """从 content_by_section 的值数组中提取文本 bullets 和特殊布局页。

    返回 (bullets, special_slides)：
    - bullets: 普通文字要点
    - special_slides: 需要特殊布局的页（图片/图表/表格）
    """
    bullets = []
    special_slides = []
    for item in items:
        if isinstance(item, str):
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
                            break
            else:
                bullets.append(text)
                if len(bullets) >= max_bullets:
                    break
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
            elif "list" in item:
                for li in item["list"][:3]:
                    bullets.append(f"● {li}")
                    if len(bullets) >= max_bullets:
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
        if len(bullets) >= max_bullets:
            break
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


def _plan_gantt(content: dict) -> dict | None:
    """顶层 plan_table → gantt 数据。plan_table: {headers, rows:[阶段,内容,时间]}。"""
    pt = content.get("plan_table") or {}
    rows = pt.get("rows") or []
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
    """方法节 → 对比表行：拆「方法名：说明」。不足两行则返回 None。"""
    rows = []
    for b in bullets:
        b = re.sub(r"^[●•]\s*", "", b).strip()
        if "：" in b:
            name, desc = b.split("：", 1)
        elif ":" in b:
            name, desc = b.split(":", 1)
        else:
            continue
        name = name.strip()
        desc = re.split(r"[。；]", desc.strip())[0][:40]
        if name:
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
        g = _plan_gantt(content)
        if g:
            slides.append(_make_slide("实施计划与时间安排", bullets, "gantt", g))
            slides.extend(special_slides)
            return slides

    if _is_route(key):
        nodes = _route_nodes(content, items)
        if nodes:
            flat = all(isinstance(n, str) for n in nodes) and len(nodes) <= 6
            slides.append(_make_slide(
                name, bullets, "flow",
                {"nodes": nodes, "direction": "horizontal" if flat else "vertical"}))
            slides.extend(special_slides)
            return slides

    if _is_method(key):
        rows = _method_rows(bullets)
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

    # 兜底：文字要点页
    if bullets:
        if len(bullets) > 5:
            slides.append(_make_slide(name, bullets[:5], "text_only"))
            slides.append(_make_slide(f"{name}（续）", bullets[5:], "text_only"))
        else:
            slides.append(_make_slide(name, bullets, "text_only"))
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
    ordered_keys = [k for k in SECTION_ORDER if k in sections]
    for k in sections:
        if k not in ordered_keys:
            ordered_keys.append(k)

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
    ppt = derive(content)
    a.output.write_text(json.dumps(ppt, ensure_ascii=False, indent=2), encoding="utf-8")
    layouts = [s["layout"] for ch in ppt["chapters"] for s in ch["slides"]]
    print(f"saved: {a.output} ({len(ppt['chapters'])} chapters, "
          f"{len(layouts)} 页：{' / '.join(sorted(set(layouts)))})")


if __name__ == "__main__":
    main()
