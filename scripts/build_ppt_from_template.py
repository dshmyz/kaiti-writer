#!/usr/bin/env python3
"""基于内置北航答辩 PPT 模板生成开题汇报 PPT。

策略：不新建幻灯片，只从模板里**挑选**需要的页、替换其占位文字、删除未用页，
保留模板全部母版/版式/配色/图形（同 build_from_template.py 的"改而不建"思路）。
"""
from __future__ import annotations
import argparse, json, re
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    raise SystemExit(
        "缺少 python-pptx，无法生成 PPT。\n"
        "  安装：pip install python-pptx（或 python3 -m pip install python-pptx）\n"
        "  不想安装：可改为只要「markdown 逐页大纲」，自己粘进 PPT。"
    )

# 各模板的页角色索引（1-based，来自实际探查）
TEMPLATE_MAP = {
    "模板1": {"cover": 1, "toc": 3, "section": [5, 10, 15, 20], "content": [7, 11, 12], "thanks": 38},
    "模板2": {"cover": 1, "toc": 3, "section": [4, 8, 12, 16, 20, 24], "content": [9, 10, 14], "thanks": 32},
    "模板3": {"cover": 1, "toc": 2, "section": [3, 9, 16, 22, 27], "content": [5, 7, 12], "thanks": 36},
}

# 模板作者水印 / 样例信息，一律清掉或替换
WATERMARKS = ["星空情报站", "微信公众号", "公众号", "微信"]

# 判定为"占位文字"的特征（含则视为可替换的占位）
PLACEHOLDER_HINTS = [
    "请输入", "请在此输入", "点击此处", "点击输入", "点击这里", "在这里输入", "在此录入",
    "单击此处", "ADD YOUR TITLE", "SAMPLE TITLE", "SIMPLE TITLE", "JUNE 12th",
    "你的标题", "你的模板你做主", "右键点击图片", "输入你的", "输入小标题", "输入标题",
    "输入文字说明", "输入你的观点", "你的模板", "Copy paste fonts", "Key words",
    "做的好的地方", "输入分点", "输入主题", "在这里输入", "输入你的内容",
]

# 纯装饰性文字：清空但不占内容位
DECOR_HINTS = ["JUNE 12th", "做的好的地方", "需要继续提升的地方", "/" * 6,
               "1997.11.12", "XXXX-XX", "Common PPT Template", "GRADUATION REPLY",
               "General template for teaching", "ADD YOUR TITLE", "请输入你的副标题",
               "请输入你的标题说明", "以简洁为主说明本章", "摘 要", "关键词0",
               "Welcoming", "Thank you for your", "图示页排版", "纯图片页排版",
               "小标题一", "小标题二", "小标题三", "小标题四", "段落标题"]

# 纯装饰的"占位符号"（整块就是这些字符时才清），如 XX、XX%、A/B/C 角标、示例日期
DECOR_EXACT = re.compile(
    r"^(X{1,4}%?|[A-F]|VS|\d{1,2}%|20\d{2}|\d{4}\.\d{2}\.\d{2}"
    r"|20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)$")





def is_placeholder(t: str) -> bool:
    t = t.strip()
    return bool(t) and any(h in t for h in PLACEHOLDER_HINTS)


def set_text_keep_style(shape, text: str):
    """只改第一个 run 的文字、删掉其余 run，保留字体/字号/颜色等全部样式。"""
    tf = shape.text_frame
    if not tf.paragraphs:
        return
    p0 = tf.paragraphs[0]
    if p0.runs:
        p0.runs[0].text = text
        for r in p0.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p0.text = text
    # 删掉后续段落，避免残留占位行
    for p in list(tf.paragraphs[1:]):
        p._p.getparent().remove(p._p)


def iter_text_shapes(container):
    """递归遍历（含 GROUP 组合形状内部）所有有文字的形状。

    模板1/模板2 的封面与多数页把文字放在组合形状里，只看顶层 shape 会全部漏掉。
    """
    for sh in container.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_text_shapes(sh)
        elif sh.has_text_frame and sh.text_frame.text.strip():
            yield sh


def text_shapes(slide):
    """按视觉顺序（上→下、左→右）返回有文字的形状（含组合内部）。"""
    return sorted(iter_text_shapes(slide), key=lambda s: (s.top or 0, s.left or 0))


def clean_watermarks(slide):
    """清掉模板作者水印（把含水印的整块文字替换成空）。含组合形状内部。"""
    for sh in iter_text_shapes(slide):
        t = sh.text_frame.text
        if any(w in t for w in WATERMARKS):
            cleaned = t
            for w in WATERMARKS:
                cleaned = cleaned.replace(w, "")
            cleaned = re.sub(r"[：:）)\s]+$", "", cleaned).strip()
            set_text_keep_style(sh, cleaned)


def clear_decor(slide):
    """清空纯装饰性的模板样例文字（英文通用模板说明、装饰串、示例日期、XX 类占位符号）。"""
    for sh in iter_text_shapes(slide):
        t = sh.text_frame.text
        if any(d in t for d in DECOR_HINTS) or DECOR_EXACT.match(t.strip()):
            set_text_keep_style(sh, "")


def fill_cover(slide, data):
    """封面：最长的那块文字当题目，含"汇报人/答辩人/指导老师/日期"的块按标签填。"""
    cover = data.get("cover", {})
    title = data.get("title", "")
    shapes = text_shapes(slide)
    # 题目 = 字数最多且非标签的那块
    cand = [sh for sh in shapes
            if not any(k in sh.text_frame.text for k in ["汇报人", "答辩人", "指导老师", "年", "BEIHANG"])]
    if cand and title:
        target = max(cand, key=lambda s: len(s.text_frame.text))
        set_text_keep_style(target, title)
    for sh in shapes:
        t = sh.text_frame.text
        if ("汇报人" in t or "答辩人" in t) and cover.get("作者姓名"):
            lbl = "答辩人" if "答辩人" in t else "汇报人"
            set_text_keep_style(sh, f"{lbl}：{cover['作者姓名']}")
        elif "指导老师" in t and cover.get("指导教师"):
            set_text_keep_style(sh, f"指导老师：{cover['指导教师']}")
        elif re.search(r"20\d{2}\s*年", t) and cover.get("日期"):
            set_text_keep_style(sh, cover["日期"])
    clear_decor(slide)
    clean_watermarks(slide)


def fill_toc(slide, chapters):
    """目录页：占位标题块按顺序换成章节名，多余的清空；纯序号块重排为 01/02…。"""
    num_slots, title_slots = [], []
    for sh in text_shapes(slide):
        t = sh.text_frame.text.strip()
        if re.fullmatch(r"0?\d{1,2}", t):
            num_slots.append(sh)
        elif is_placeholder(t) and "ADD YOUR" not in t:
            title_slots.append(sh)
    for i, sh in enumerate(title_slots):
        set_text_keep_style(sh, chapters[i] if i < len(chapters) else "")
    for i, sh in enumerate(num_slots):
        set_text_keep_style(sh, f"{i + 1:02d}" if i < len(chapters) else "")
    clear_decor(slide)
    clean_watermarks(slide)


def fill_section(slide, idx, name):
    """章节过渡页：序号写 01/02…，第一个标题占位写章节名，其余占位清空。"""
    first = True
    for sh in text_shapes(slide):
        t = sh.text_frame.text.strip()
        if re.fullmatch(r"0?\d{1,2}", t):
            set_text_keep_style(sh, f"{idx:02d}")
        elif is_placeholder(t):
            if first and "ADD YOUR" not in t:
                set_text_keep_style(sh, name)
                first = False
            else:
                set_text_keep_style(sh, "")
    clear_decor(slide)
    clean_watermarks(slide)


def fill_content(slide, title, bullets):
    """内容页：首块占位当标题，其余占位块依次填要点，多余清空。"""
    clear_decor(slide)
    slots = [sh for sh in text_shapes(slide) if is_placeholder(sh.text_frame.text)]
    if not slots:
        clean_watermarks(slide)
        return
    set_text_keep_style(slots[0], title)
    for i, sh in enumerate(slots[1:]):
        set_text_keep_style(sh, bullets[i] if i < len(bullets) else "")
    clean_watermarks(slide)


def drop_slides(prs, keep_indices):
    """删除未保留的页（1-based keep_indices），同时清理 sldIdLst 与关系。"""
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    for i in range(len(ids) - 1, -1, -1):
        if (i + 1) not in keep_indices:
            rId = ids[i].get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            prs.part.drop_rel(rId)
            sldIdLst.remove(ids[i])


def reorder_slides(prs, order):
    """按 order（slide 对象列表）重排 sldIdLst，使输出顺序为 封面→目录→各章→致谢。"""
    sldIdLst = prs.slides._sldIdLst
    id_by_slide = {}
    for sid in list(sldIdLst):
        rId = sid.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        id_by_slide[prs.part.related_part(rId).partname] = sid
    for slide in order:
        sid = id_by_slide.get(slide.part.partname)
        if sid is not None:
            sldIdLst.remove(sid)
            sldIdLst.append(sid)


def duplicate_slide(prs, src_slide, pristine_xml=None):
    """复制一页（含所有形状与样式），追加到末尾并返回新页。用于内容页复用。

    pristine_xml：源页"未被填写前"的 spTree 副本。必须传，否则复制到的是已填过
    上一章内容的页面（同一模板页被多章复用时会串内容）。
    """
    from copy import deepcopy
    new_slide = prs.slides.add_slide(src_slide.slide_layout)
    for sh in list(new_slide.shapes):      # 清掉版式带来的占位符
        sh._element.getparent().remove(sh._element)
    src_tree = pristine_xml if pristine_xml is not None else src_slide.shapes._spTree
    for child in src_tree:
        tag = child.tag.split('}')[-1]
        if tag in ("nvGrpSpPr", "grpSpPr"):   # 组属性新页已有，跳过
            continue
        new_slide.shapes._spTree.append(deepcopy(child))
    for rel in src_slide.part.rels.values():
        if rel.is_external:
            new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_slide.part.rels.get_or_add(rel.reltype, rel._target)
    return new_slide


def build(template: Path, content_path: Path, output: Path):
    data = json.loads(content_path.read_text(encoding="utf-8"))
    key = next((k for k in TEMPLATE_MAP if k in template.stem), None)
    if key is None:
        raise SystemExit(f"未识别的模板：{template.name}（应含 模板1/模板2/模板3）")
    roles = TEMPLATE_MAP[key]

    prs = Presentation(str(template))
    slides = list(prs.slides)
    chapters = data.get("chapters", [])
    if not chapters:
        raise SystemExit("content.json 缺少 chapters（章节列表）")

    cover_slide = slides[roles["cover"] - 1]
    toc_slide = slides[roles["toc"] - 1]
    thanks_slide = slides[roles["thanks"] - 1] if roles.get("thanks") else None

    fill_cover(cover_slide, data)
    fill_toc(toc_slide, [c["name"] for c in chapters])

    order = [cover_slide, toc_slide]
    sec_pool = [slides[i - 1] for i in roles["section"]]
    con_pool = [slides[i - 1] for i in roles["content"]]

    # 关键：在任何填写之前先留存"干净"的形状树副本，供复制页使用，
    # 否则复用同一模板页时会复制到上一章已填的内容。
    from copy import deepcopy
    pristine = {s.part.partname: deepcopy(s.shapes._spTree)
                for s in sec_pool + con_pool}

    for ci, ch in enumerate(chapters):
        # 章节过渡页：模板页够用就直接用，不够则复制最后一个章节页样式
        if ci < len(sec_pool):
            sec = sec_pool[ci]
        else:
            src = sec_pool[-1]
            sec = duplicate_slide(prs, src, pristine.get(src.part.partname))
        fill_section(sec, ci + 1, ch["name"])
        order.append(sec)

        # 内容页：轮转取模板内容页样式；首次直接用、其后按干净副本复制
        for si, spec in enumerate(ch.get("slides", [])):
            base = con_pool[si % len(con_pool)]
            if base in order:            # 该模板页已用过 → 复制一份干净的
                slide = duplicate_slide(prs, base, pristine.get(base.part.partname))
            else:
                slide = base
            fill_content(slide, spec.get("title", ""), spec.get("bullets", []))
            order.append(slide)

    if thanks_slide is not None:
        clear_decor(thanks_slide)
        clean_watermarks(thanks_slide)
        cover = data.get("cover", {})
        for sh in iter_text_shapes(thanks_slide):
            t = sh.text_frame.text.strip()
            if t in ("汇报人", "答辩人", "汇报人：", "答辩人：") and cover.get("作者姓名"):
                set_text_keep_style(sh, f"{t.rstrip('：')}：{cover['作者姓名']}")
            elif t in ("指导老师", "指导老师：", "指导教师", "指导教师：") and cover.get("指导教师"):
                set_text_keep_style(sh, f"{t.rstrip('：')}：{cover['指导教师']}")
        order.append(thanks_slide)

    # 删除所有未用到的模板页，再按逻辑顺序重排
    keep_partnames = {s.part.partname for s in order}
    keep_idx = [i + 1 for i, s in enumerate(prs.slides) if s.part.partname in keep_partnames]
    drop_slides(prs, keep_idx)
    reorder_slides(prs, order)

    # 清掉内置模板带来的第三方个人信息：模板是借用的往届答辩稿，
    # docProps/core.xml 里留着原作者姓名与邮箱。不清则每份汇报 PPT 都会在
    # 属性面板显示陌生人的邮箱，既泄露他人隐私，也像是别人做的稿子。
    cp = prs.core_properties
    scrubbed = [k for k, v in (("author", cp.author), ("last_modified_by", cp.last_modified_by))
                if v]
    au = str(data.get("author", "") or "")
    if any(m in au for m in ("〈", "〉", "XXX", "待补", "待定")):
        au = ""
    cp.author = au
    cp.last_modified_by = au
    cp.title = str(data.get("title", "") or "")

    prs.save(str(output))
    print("saved:", output)
    validate(output, data, scrubbed=scrubbed)
    validate_content_density(data)


def validate(path, data, scrubbed=None):
    prs = Presentation(str(path))
    all_text = []
    for s in prs.slides:
        for sh in iter_text_shapes(s):
            all_text.append(sh.text_frame.text)
    text = "\n".join(all_text)
    leftover_ph = [h for h in ("请输入", "点击此处", "在这里输入", "JUNE 12th", "ADD YOUR") if h in text]
    leftover_wm = [w for w in WATERMARKS if w in text]
    print("=" * 50)
    print("自检：")
    print(f"  页数: {len(prs.slides)}")
    print(f"  题目已写入: {data.get('title','')[:20] in text}")
    print(f"  章节数: {len(data.get('chapters', []))}")
    print(f"  水印已清: {not leftover_wm} {leftover_wm if leftover_wm else ''}")
    print(f"  残留占位: {leftover_ph if leftover_ph else '无'}")
    if leftover_ph:
        print("  ↑ 提示：复杂图表页可能仍有占位文字/图片，需在 PowerPoint 手动微调")
    if scrubbed:
        print(f"  文档属性已清理: {scrubbed}（模板原作者姓名/邮箱已清空）")


def validate_content_density(data):
    """检查 ppt_content.json 的内容密度是否达标，给出改进建议。"""
    chapters = data.get("chapters", [])
    total_slides = sum(len(ch.get("slides", [])) for ch in chapters)
    issues = []

    # 总页数检查（含封面/目录/致谢约 +3 页）
    effective = total_slides + 3
    if effective < 18:
        issues.append(f"总页数仅 {effective} 页，8 分钟汇报建议 20–25 页；"
                      "每章至少 2 页（1 章节页 + ≥1 内容页），重点章节 3–4 页")

    # 逐章检查
    thin_chapters = []
    for ch in chapters:
        n = len(ch.get("slides", []))
        if n < 2:
            thin_chapters.append(f"「{ch['name']}」仅 {n} 页")
    if thin_chapters:
        issues.append("以下章节内容过薄（<2 页），建议补充：" + "；".join(thin_chapters))

    # 检查是否有图表占位
    has_diagram = False
    for ch in chapters:
        for sl in ch.get("slides", []):
            for b in sl.get("bullets", []):
                if "【图" in b or "【表" in b or "路线图" in b:
                    has_diagram = True
    if not has_diagram:
        issues.append("未发现图表占位（【图：xxx】/【表：xxx】），研究框架/路线图/时间表建议嵌图")

    # 检查章节标题是否太泛
    generic_titles = {"研究背景", "文献综述", "研究方法", "创新之处", "实施计划",
                      "选题意义", "研究框架", "研究思路"}
    for ch in chapters:
        for sl in ch.get("slides", []):
            if sl.get("title", "") in generic_titles:
                issues.append(f"页面标题「{sl['title']}」太泛——应写具体观点，不是章节名")
                break

    # 检查占位残留
    for ch in chapters:
        for sl in ch.get("slides", []):
            for b in sl.get("bullets", []):
                if "见下页" in b or "添加" in b:
                    issues.append(f"「{ch['name']}」存在占位残留：「{b}」— 要么嵌入内容，要么写【图：xxx】占位")

    if issues:
        print("\n⚠ 内容密度自检：")
        for i, issue in enumerate(issues, 1):
            print(f"  {i}. {issue}")
        print("  ↑ 以上问题不影响生成，但会影响汇报质量。建议补充后重新生成。")
    else:
        print("\n✓ 内容密度自检通过")


def main():
    ap = argparse.ArgumentParser(description="用内置北航模板生成开题汇报 PPT")
    ap.add_argument("--template", required=True, type=Path)
    ap.add_argument("--content", required=True, type=Path)
    ap.add_argument("--output", required=True, type=Path)
    a = ap.parse_args()
    build(a.template, a.content, a.output)


if __name__ == "__main__":
    main()
