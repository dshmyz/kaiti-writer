#!/usr/bin/env python3
"""从 content.json 自动派生 ppt_content.json，省去手工重复写 PPT 内容。

用法：
    python derive_ppt_content.py --content content.json --output ppt_content.json

映射规则：
- content_by_section 的每个节 → PPT 的一个 chapter
- 节标题 → chapter.name
- 节内的文本段落 → slides 的 bullets（前 5 条）
- 节内的 image 块 → 标记为【图：caption】
- 节内的 table 块 → 标记为【表：caption】
- 封面信息直接复用
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

# 节标题清理：去掉序号前缀，保留核心名
def clean_section_name(key: str) -> str:
    """（一）研究背景 → 研究背景"""
    return re.sub(r"^[（(一二三四五六七八九十]+[）)]\s*", "", key).strip()


def extract_bullets(items: list, max_bullets: int = 5) -> tuple[list, list]:
    """从 content_by_section 的值数组中提取文本 bullets 和特殊布局页。

    返回 (bullets, special_slides)：
    - bullets: 普通文字要点
    - special_slides: 需要特殊布局的页（图片/图表/表格）
    """
    bullets = []
    special_slides = []
    for item in items:
        if isinstance(item, str):
            # 长段落按句号拆分，取前几句
            text = item.strip()
            if not text:
                continue
            if len(text) > 60:
                # 按句号/分号拆，取前 max_bullets 段
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
                # 图片页单独一页，不混在 bullets 里
                special_slides.append({
                    "title": caption,
                    "layout": "image_center",
                    "bullets": [],
                    "extra": {"image": image_path, "caption": caption}
                })
            elif "table" in item:
                caption = item.get("caption", "")
                headers = item["table"].get("headers", [])
                rows = item["table"].get("rows", [])
                # 表格页单独一页
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
                # 图表数据
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


def derive(content: dict) -> dict:
    """从 content.json 派生 ppt_content.json。"""
    ppt = {
        "title": content.get("title", ""),
        "subtitle": "开题汇报",
        "cover": content.get("cover", {}),
        "chapters": [],
    }

    sections = content.get("content_by_section", {})

    # 按 SECTION_ORDER 的顺序排列，未在列表里的节追加到末尾
    ordered_keys = []
    for k in SECTION_ORDER:
        if k in sections:
            ordered_keys.append(k)
    for k in sections:
        if k not in ordered_keys:
            ordered_keys.append(k)

    for key in ordered_keys:
        items = sections[key]
        name = clean_section_name(key)
        bullets, special_slides = extract_bullets(items)

        slides = []

        # 普通文字页（智能检测布局）
        if bullets:
            # 检测是否应该用特殊布局
            smart_layout, smart_extra = _detect_smart_layout(name, bullets, items)
            if smart_layout != "text_only" and smart_extra:
                slides.append({
                    "title": name,
                    "bullets": bullets,
                    "layout": smart_layout,
                    "extra": smart_extra
                })
            elif len(bullets) > 5:
                # bullets 超过 5 条，拆成两页
                slides.append({"title": name, "bullets": bullets[:5]})
                slides.append({"title": f"{name}（续）", "bullets": bullets[5:]})
            else:
                slides.append({"title": name, "bullets": bullets})

        # 特殊布局页（图片/图表/表格）
        for special in special_slides:
            slides.append(special)

        if not slides:
            continue

        # 为每页生成 speaker notes
        for slide in slides:
            slide["notes"] = _generate_notes(name, slide)

        ppt["chapters"].append({"name": name, "slides": slides})

    return ppt


def _generate_notes(chapter_name: str, slide: dict) -> str:
    """根据页面内容自动生成演讲者备注。"""
    title = slide.get("title", "")
    bullets = slide.get("bullets", [])
    layout = slide.get("layout", "text_only")

    notes_parts = []

    # 开场提示
    if "背景" in chapter_name or "意义" in chapter_name:
        notes_parts.append("开场：用具体案例/数据引出问题，不要从宏观政策开始。")
    elif "文献" in chapter_name:
        notes_parts.append("过渡语：前面讲了问题，现在看看别人怎么做的，有什么不足。")
    elif "框架" in chapter_name or "思路" in chapter_name:
        notes_parts.append("过渡语：基于文献不足，我的研究思路是...")
    elif "方法" in chapter_name:
        notes_parts.append("过渡语：具体怎么做？用三种方法。")
    elif "创新" in chapter_name:
        notes_parts.append("重点：这是评审最关注的页，讲清楚新在哪里。")
    elif "计划" in chapter_name:
        notes_parts.append("收尾：时间节点清晰，让评审觉得可行。")

    # 内容提示
    if layout == "chart":
        notes_parts.append("指向图表：重点讲数据趋势，不要逐个读数字。")
    elif layout == "image_center":
        notes_parts.append("指向图片：解释图中关键要素，说明其与研究的关系。")
    elif layout == "table":
        notes_parts.append("指向表格：对比差异，不要逐行念。")
    elif len(bullets) > 3:
        notes_parts.append("要点较多，挑重点讲，其余让评审自己看。")

    # 结尾提示
    if "创新" in title:
        notes_parts.append("强调：每个创新点用一句话说清楚。")
    elif "实施" in title or "计划" in title:
        notes_parts.append("结尾：时间节点明确，展示可行性。")

    return "\n".join(notes_parts) if notes_parts else ""


def _detect_smart_layout(chapter_name: str, bullets: list[str], items: list) -> tuple[str, dict | None]:
    """智能检测页面布局：根据内容自动推荐最佳布局。

    返回 (layout, extra)：
    - layout: text_only / big_number / comparison / timeline / table
    - extra: 特殊布局的额外数据
    """
    # 检测百分比数据 → 大数字布局
    for b in bullets:
        pct_match = re.search(r"(\d{1,3})%", b)
        if pct_match:
            num = pct_match.group(1)
            # 提取上下文作为标签
            label = re.sub(r"\d{1,3}%", "", b).strip().strip("，。、：")
            if len(label) > 20:
                label = label[:20] + "..."
            return "big_number", {"number": f"{num}%", "label": label}

    # 检测对比结构（"A vs B"、"优于""高于""低于"）→ 对比布局
    comparison_keywords = ["优于", "高于", "低于", "多于", "少于", "相比", "对比", "差异"]
    for b in bullets:
        if any(kw in b for kw in comparison_keywords):
            # 提取对比双方
            parts = re.split(r"[，。；]", b)
            if len(parts) >= 2:
                return "comparison", {"items": [p.strip() for p in parts if p.strip()][:4]}

    # 检测时间线（"第一阶段""2024年""2025年"等）→ 时间线布局
    timeline_patterns = [r"第[一二三四]阶段", r"\d{4}年", r"前期|中期|后期", r"个月"]
    for b in bullets:
        if any(re.search(p, b) for p in timeline_patterns):
            return "timeline", {"items": [re.sub(r"^[●•]\s*", "", b) for b in bullets[:5]]}

    # 检测方法对比（含"方法""方法论"且有多条）→ 表格布局
    if "方法" in chapter_name and len(bullets) >= 3:
        return "table", {
            "headers": ["方法", "适用场景", "优势"],
            "rows": [re.split(r"[，：:]", b)[:3] for b in bullets[:4]]
        }

    # 默认：普通文字
    return "text_only", None


def main():
    ap = argparse.ArgumentParser(description="从 content.json 派生 ppt_content.json")
    ap.add_argument("--content", required=True, type=Path)
    ap.add_argument("--output", required=True, type=Path)
    a = ap.parse_args()

    content = json.loads(a.content.read_text(encoding="utf-8"))
    ppt = derive(content)
    a.output.write_text(json.dumps(ppt, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"saved: {a.output} ({len(ppt['chapters'])} chapters)")


if __name__ == "__main__":
    main()
