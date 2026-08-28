#!/usr/bin/env python3
"""GB/T 7714 顺序编码制引用编号校验与重排工具。

顺序编码制要求：正文引用标注 [n] 按首次出现顺序递增（[8] 不得早于 [2] 出现），
参考文献表按编号升序排列。插入新文献后重编号全靠手改是踩过的坑，本工具自动完成。

用法：
  校验（只报不修）：  python renumber_refs.py --docx 报告.docx --content content.json --check
  自动重排（正文+文献表同步）：  python renumber_refs.py --docx 报告.docx --content content.json --fix
  插入新文献：content.json 的 refs 数组末尾加 "[99] 新文献条目"，正文引用处写 [99]，
  然后跑 --fix，99 及其后所有编号自动归位（99→实际位次，原 19-21 顺延为 20-22）。

原理：扫描正文引用标注的首次出现顺序 → 建立 旧编号→新编号 映射 → 同步重写
正文标注与文献表序号。--check 与 build_and_validate.py 第 4.5 步规则一致。
"""
import argparse
import json
import re
import sys


CITE_RE = re.compile(r"\[(\d+(?:\s*[,，]\s*\d+)*(?:\s*-\s*\d+)?)\]")


def extract_first_order(docx_path):
    """从 docx 正文（参考文献表之前）提取引用编号的首次出现顺序。"""
    try:
        from docx import Document
    except ImportError:
        sys.exit("缺少 python-docx：pip install python-docx")
    doc = Document(docx_path)
    refs_start = None
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        # 取最后一个"参考文献"标题：目录里的"参考文献\t页码"占位在前，
        # 正文插入的才是真内容（与 build_and_validate.py 第 4 节同规则）
        ref_title = t.split(chr(9))[0] if chr(9) in t else t
        if "参考文献" in ref_title and len(ref_title) < 10:
            refs_start = i
    order, seen = [], set()
    for i, p in enumerate(doc.paragraphs):
        if refs_start is not None and i == refs_start:
            break
        for m in CITE_RE.finditer(p.text):
            for num_str in re.split(r"[,，\-]", m.group(1)):
                n = int(num_str.strip())
                if n and n not in seen:
                    seen.add(n)
                    order.append(n)
    return order, doc


def build_mapping(order):
    """首次出现顺序 → 新编号映射（首次出现的排 1, 2, 3…）。"""
    return {old: new + 1 for new, old in enumerate(order)}


def remap_cite(match, mapping):
    """把一处标注 [3, 5-7] 的每个编号按映射替换，保持多编号/区间结构。"""
    inner = match.group(1)
    def sub_num(m):
        old = int(m.group(0))
        return str(mapping.get(old, old))
    return "[" + re.sub(r"\d+", sub_num, inner) + "]"


def sort_citation(text):
    """重排后标注内的编号可能乱序（如 [5, 2]），排序回 [2, 5]。"""
    def _sort(m):
        nums = [int(x) for x in re.split(r"[,，\s]+", m.group(1).strip()) if x]
        if len(nums) > 1 and nums != sorted(nums):
            return "[" + ", ".join(str(n) for n in sorted(nums)) + "]"
        return m.group(0)
    return re.sub(r"\[(\d+(?:\s*[,，]\s*\d+)+)\]", _sort, text)


def check_docx(order):
    """返回乱序编号列表（空 = 合规）。"""
    return [b for a, b in zip(order, order[1:]) if b < a]


def main():
    ap = argparse.ArgumentParser(description="GB/T 7714 顺序编码制引用编号校验与重排")
    ap.add_argument("--docx", required=True, help="开题报告 .docx 路径")
    ap.add_argument("--content", help="content.json 路径（--fix 时同步重排 refs 序号）")
    ap.add_argument("--check", action="store_true", help="只校验不修改")
    ap.add_argument("--fix", action="store_true", help="自动重排正文标注 + 文献表序号")
    args = ap.parse_args()

    if not args.check and not args.fix:
        ap.error("至少指定 --check 或 --fix 之一")

    order, doc = extract_first_order(args.docx)
    if not order:
        print("正文无引用标注，无需处理")
        return
    bad = check_docx(order)
    if not bad:
        print(f"✅ 引用编号按首次出现顺序合规（{min(order)}-{max(order)}，共 {len(order)} 个编号）")
        return
    if args.check:
        print(f"⚠️  引用编号未按首次出现顺序：乱序出现在 {bad[:8]}")
        print("    → 运行本工具 --fix 自动重排")
        sys.exit(1)

    # --- fix ---
    mapping = build_mapping(order)
    print(f"重排映射（旧→新）：{ {k: v for k, v in sorted(mapping.items()) if k != v} or '无需变动' }")

    # 1) 重写 docx 正文标注
    n_cites = 0
    refs_start = None
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        ref_title = t.split(chr(9))[0] if chr(9) in t else t
        if "参考文献" in ref_title and len(ref_title) < 10:
            refs_start = i  # 不 break：取最后一个（目录占位在前）
    for i, p in enumerate(doc.paragraphs):
        if refs_start is not None and i == refs_start:
            break
        for run in p.runs:
            if run.text and CITE_RE.search(run.text):
                run.text = sort_citation(CITE_RE.sub(lambda m: remap_cite(m, mapping), run.text))
                n_cites += 1

    # 2) 重写文献表序号（[旧] → [新]，按映射）
    if refs_start is not None:
        for p in doc.paragraphs[refs_start + 1:]:
            t = p.text.strip()
            if t.startswith(("附录", "附件")):
                break
            m = re.match(r"^\[(\d+)\]", t)
            if not m:
                continue
            old = int(m.group(1))
            new = mapping.get(old)
            if new and new != old:
                for run in p.runs:
                    if run.text and f"[{old}]" in run.text:
                        run.text = run.text.replace(f"[{old}]", f"[{new}]", 1)
                        break
                else:
                    # 序号可能在首个 run 与文本分离时——整体重写首 run
                    if p.runs:
                        p.runs[0].text = re.sub(r"^\[\d+\]", f"[{new}]", p.runs[0].text)
    doc.save(args.docx)
    print(f"✅ 已重排：正文 {n_cites} 处标注 + 文献表序号，保存至 {args.docx}")

    # 3) 同步重排 content.json：refs 数组（按新编号排序）+ 正文文本里的引用标注
    if args.content:
        with open(args.content, encoding="utf-8") as f:
            data = json.load(f)

        # 3a) 所有文本字段里的 [n] 标注按映射重写（正文引用在 section 文本里）
        def remap_text(v):
            if isinstance(v, str):
                if CITE_RE.search(v):
                    return sort_citation(CITE_RE.sub(lambda m: remap_cite(m, mapping), v))
                return v
            if isinstance(v, list):
                return [remap_text(x) for x in v]
            if isinstance(v, dict):
                return {k: remap_text(x) for k, x in v.items()}
            return v
        data = remap_text(data)

        refs = data.get("refs", [])
        def ref_num(s):
            m = re.match(r"^\[(\d+)\]", s)
            return mapping.get(int(m.group(1)), int(m.group(1))) if m else 10**9
        data["refs"] = sorted(refs, key=ref_num)
        # 重写序号本身
        fixed = []
        for s in data["refs"]:
            m = re.match(r"^\[(\d+)\]", s)
            if m:
                old = int(m.group(1))
                s = re.sub(r"^\[\d+\]", f"[{mapping.get(old, old)}]", s)
            fixed.append(s)
        data["refs"] = fixed
        with open(args.content, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"✅ content.json refs 已同步重排（{len(fixed)} 条）")


if __name__ == "__main__":
    main()
