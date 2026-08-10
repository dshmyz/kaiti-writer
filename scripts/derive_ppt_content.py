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


def extract_bullets(items: list, max_bullets: int = 5) -> list[str]:
    """从 content_by_section 的值数组中提取文本 bullets。"""
    bullets = []
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
                bullets.append(f"【图：{caption}】")
            elif "table" in item:
                caption = item.get("caption", "")
                headers = item["table"].get("headers", [])
                bullets.append(f"【表：{caption or ' '.join(headers[:3])}】")
            elif "list" in item:
                for li in item["list"][:3]:
                    bullets.append(f"● {li}")
                    if len(bullets) >= max_bullets:
                        break
        if len(bullets) >= max_bullets:
            break
    return bullets


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
        bullets = extract_bullets(items)

        if not bullets:
            continue

        slides = [{"title": name, "bullets": bullets}]

        # 如果 bullets 超过 5 条，拆成两页
        if len(bullets) > 5:
            slides = [
                {"title": name, "bullets": bullets[:5]},
                {"title": f"{name}（续）", "bullets": bullets[5:]},
            ]

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
    print(f"saved: {a.output} ({len(ppt['chapters'])} chapters)")


if __name__ == "__main__":
    main()
