#!/usr/bin/env python3
"""读取已有开题报告 .docx，按节切分正文，供「审阅/修订已有稿」流程使用。

只读，不改原文件。跳过封面与目录（用正文「开题报告」标题作起点），
遇「参考文献」/「附录」/「书写规范」停止收集正文节。

用法：
    python read_docx.py --input 我的开题报告.docx                # 人读的概览
    python read_docx.py --input 我的开题报告.docx --json out.json  # 供改写用的结构化数据

老 .doc（OLE 格式，非 zip）读不了，先转换：
    textutil -convert docx 我的开题报告.doc      # macOS
"""
import argparse
import json
import re
import zipfile
from pathlib import Path

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 节标题形态：「一、xxx」「（一）xxx」，外加模板里没编号的「研究框架（内容）」
SEC_RE = re.compile(r"^([一二三四五六七八九十]、|（[一二三四五六七八九十]）)")
STOP_PREFIXES = ("参考文献", "附录", "书写规范", "目录按章")
EXTRA_SECTIONS = ("研究框架（内容）",)


def is_section_heading(text):
    if len(text) > 30:
        return False
    return bool(SEC_RE.match(text)) or text in EXTRA_SECTIONS


def zh_count(s):
    return sum(1 for c in s if "一" <= c <= "鿿")


def read_docx(path):
    from lxml import etree

    path = Path(path).expanduser()
    if not path.is_file():
        raise SystemExit(f"文件不存在: {path}")
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise SystemExit(
            f"{path.name} 不是 docx（可能是老 .doc OLE 格式）。先转换：\n"
            f"  textutil -convert docx '{path}'      # macOS\n"
            "  或用 Word 另存为 .docx"
        )
    if "word/document.xml" not in z.namelist():
        raise SystemExit(f"{path.name} 里没有 word/document.xml，不是有效的 Word 文档。")

    body = etree.fromstring(z.read("word/document.xml")).find(W + "body")
    kids = list(body)

    # 正文起点：精确匹配「开题报告」标题，避开封面 banner 与目录
    start = 0
    for i, el in enumerate(kids):
        if el.tag == W + "p" and "".join(el.itertext()).strip() == "开题报告":
            start = i + 1          # 取最后一次出现（目录之后即正文）

    title = ""
    for el in kids[:start]:
        if el.tag != W + "p":
            continue
        t = "".join(el.itertext()).strip()
        if len(t) > 6 and not t.startswith(("北京航空航天大学", "专业硕士", "目录", "开题报告")):
            title = t
            break

    sections, refs, appendix = {}, [], []
    cur, phase = None, "body"
    for el in kids[start:]:
        tag = etree.QName(el).localname
        if tag == "tbl":
            # 附录现在也能放表格（问卷量表题），一并收进来，否则回读会漏掉整张量表。
            bucket = sections.get(cur) if phase == "body" and cur else (
                appendix if phase == "appendix" else None)
            if bucket is not None:
                rows = []
                for tr in el.findall(W + "tr"):
                    rows.append([" ".join("".join(tc.itertext()).split())
                                 for tc in tr.findall(W + "tc")])
                bucket.append({"table": rows})
            continue
        if tag != "p":
            continue
        t = " ".join("".join(el.itertext()).split())
        if not t:
            continue
        if t.startswith("参考文献"):
            phase = "refs"; continue
        # 只在首次遇到「附录」标题时切相位。附录内的条目也常以「附录A/附录B」开头，
        # 若不加这个守卫，它们会被当成标题反复 continue 掉，附录永远读成空的。
        if t.startswith("附录") and phase != "appendix":
            phase = "appendix"; continue
        if t.startswith(("书写规范", "目录按章")):
            break
        if phase == "refs":
            refs.append(t); continue
        if phase == "appendix":
            appendix.append(t); continue
        if is_section_heading(t):
            cur = t; sections.setdefault(cur, [])
        elif cur:
            sections[cur].append(t)

    return {"file": str(path), "title": title, "sections": sections,
            "refs": refs, "appendix": appendix}


def main():
    ap = argparse.ArgumentParser(description="读取已有开题报告 docx 并按节切分")
    ap.add_argument("--input", required=True, help="已有开题报告 .docx")
    ap.add_argument("--json", help="把结构化结果写到这个 json（供改写用）")
    ap.add_argument("--full", action="store_true", help="打印每节全文而非首句")
    args = ap.parse_args()

    data = read_docx(args.input)
    secs = data["sections"]
    body_zh = sum(zh_count(b) for blocks in secs.values()
                  for b in blocks if isinstance(b, str))

    print(f"题目: {data['title'] or '（未识别到）'}")
    print(f"正文: {len(secs)} 节 / 约 {body_zh} 中文字")
    print(f"参考文献: {len(data['refs'])} 条    附录: {len(data['appendix'])} 条")
    print("=" * 50)
    for name, blocks in secs.items():
        paras = [b for b in blocks if isinstance(b, str)]
        tbls = [b for b in blocks if isinstance(b, dict)]
        # 章标题（「一、」）是容器，正文挂在其下的「（一）」里，本身没段落不算缺
        is_chapter = bool(re.match(r"^[一二三四五六七八九十]、", name))
        flag = "" if (paras or tbls or is_chapter) else "  ⚠️ 本节无正文"
        print(f"\n【{name}】{len(paras)} 段 / {zh_count(''.join(paras))} 字"
              f"{f' / {len(tbls)} 表' if tbls else ''}{flag}")
        for p in (paras if args.full else paras[:2]):
            print(f"    {p if args.full else p[:60] + ('…' if len(p) > 60 else '')}")
        if not args.full and len(paras) > 2:
            print(f"    …（另 {len(paras) - 2} 段）")

    missing = [s for s in ("（一）研究背景", "（二）选题意义", "（三）文献综述",
                           "（一）研究思路", "（一）研究方法", "（二）创新之处")
               if s not in secs]
    if missing:
        print(f"\n⚠️ 未识别到这些常规节（可能标题写法不同或确实缺失）: {missing}")
    if not any("预期" in s for s in secs):
        print("⚠️ 未见「预期目标和成果」——《管理办法》第三条 7 项必备之一。")
    # 模板附录硬要求：「若论文包含访谈或问卷，须附访谈提纲或问卷设计」。
    # 正文提了问卷/访谈却没附录，是审阅时的高发问题，这里直接点出来。
    body_blob = "".join(b for blocks in secs.values()
                        for b in blocks if isinstance(b, str))
    uses = [w for w in ("问卷", "访谈") if w in body_blob]
    # 只有「附录A」「附录B」这类光标题、底下一条题项都没有，等于没附——照样要报。
    ap_subst = [b for b in data["appendix"]
                if isinstance(b, dict) or not re.match(r"^附录[A-Za-z一二三四五六七八九十]?\s*$", b)]
    if uses and not ap_subst:
        print(f"⚠️ 正文提到{'/'.join(uses)}，但附录"
              f"{'为空' if not data['appendix'] else '只有标题没有实际题项'}"
              f"——模板要求「若论文包含访谈或问卷，须附访谈提纲或问卷设计」，"
              f"做法见 references/问卷与访谈提纲.md。")

    if args.json:
        out = Path(args.json).expanduser()
        out.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
        print(f"\n结构化数据已写入: {out}")


if __name__ == "__main__":
    main()
