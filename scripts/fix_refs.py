#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""参考文献表 GB/T 7714 收尾核查与修复。

生成开题报告参考文献（写进 content.json 的 refs、进而转成 docx）后，
在交付前对每条目做收尾核查：去末尾结束符、清"待核验/占位"括号（缺项连标点略去）、
统一半角标点。规则出处见 references/参考文献著录.md「写入前收尾核查」。

用法：
    # 修 content.json 的 refs（顺便列出待补真实数据的缺项清单）
    python fix_refs.py --content content.json \
        --list-missing --out missing.txt
    # 直接修已有开题报告 docx 的参考文献段（需 lxml；就地改，先自备备份）
    python fix_refs.py --docx 开题报告-xxx.docx --list-missing

不依赖第三方库（json 模式）；docx 模式需要 lxml（缺则报错并提示 pip install lxml）。
"""
from __future__ import annotations
import argparse, json, os, re, sys, tempfile


def fix_one(s: str) -> str:
    """对单条参考文献做 GB/T 7714 收尾：去占位括号、去末尾结束符、半角化孤立尾标点。"""
    # 1) 删含"待核验/待补/待验证/TODO"的成对括号（缺口径，避免误删真实期号/页码，如 38(2)）
    s = re.sub(r"[（(][^（）()]*?(?:待核验|待补|待验证|TODO)[^（）()]*?[）)]", "", s)
    # 2) 反复消化条目尾部：先切掉尾空白，再去末尾结束符/多余标点（GB/T 7714：每条最后不用加结束符）
    #    循环到稳定，因为删占位可能残留"标点+空格+标点"之类（如 "[Z]. （待核验）." 删占位后剩 "[Z]. ."）
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r"\s+$", "", s)               # 去尾空白
        s = re.sub(r"[.,;:、，。；：]+$", "", s)   # 去末尾结束符/标点
    # 3) 孤立的类型标识尾部句点，如 "[Z]." -> "[Z]"（缺项连同标点一并略去）
    s = re.sub(r"(\[[A-Z/]+\])\s*[.,;:]\s*$", r"\1", s)
    return s.strip()


def detect_missing(s: str) -> str | None:
    """返回条目里的缺项原文（含"待核验/待补"的括号），无则 None。"""
    m = re.search(r"[（(][^（）()]*?(?:待核验|待补|待验证|TODO)[^（）()]*?[）)]", s)
    return m.group(0) if m else None


def fix_json(path: str, list_missing: bool, missing_out: str) -> None:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    refs = data.get("refs")
    if not isinstance(refs, list):
        print("⚠️  json 无 refs 数组，跳过")
        return
    missing = []
    changed = 0
    for i, r in enumerate(refs):
        if not isinstance(r, str):
            continue
        m = detect_missing(r)
        if m:
            missing.append((i + 1, r, m))
        new = fix_one(r)
        if new != r:
            refs[i] = new
            changed += 1
            print(f"  ★ [{i+1}] {r}")
            print(f"     →  {new}")
    if changed:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"json refs 修复 {changed} 条 → 已写回 {path}")
    else:
        print("json refs 均符合 GB/T 7714，无需修复")
    if list_missing:
        _report_missing(missing, missing_out)


def _report_missing(missing, missing_out):
    if not missing:
        print("…无'待核验/待补'缺项")
        return
    lines = ["检测到待补真实数据的缺项（按 [序号] 原文 → 缺项 → 需补内容）："]
    for n, r, m in missing:
        lines.append(f"[{n}] {r}")
        lines.append(f"     缺项: {m} → 请用户核验补全真实值再定稿")
    print("\n".join(lines))
    if missing_out:
        with open(missing_out, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print(f"缺项清单已写: {missing_out}")
    print("⚠️ 这些是待补真值，不是格式问题——不要自行编造文献细节。")


def fix_docx(path: str, list_missing: bool, missing_out: str) -> None:
    try:
        from lxml import etree as ET
    except ImportError:
        sys.exit("缺少 lxml，请先安装：pip install lxml（处理 docx 必需）")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    q = lambda t: f"{{{W}}}{t}"
    import zipfile
    with zipfile.ZipFile(path) as z:
        members = {n: z.read(n) for n in z.namelist()}
    root = ET.fromstring(members["word/document.xml"])
    body = root.find(q("body"))
    paras = body.findall(q("p"))
    txt = lambda p: "".join(t.text or "" for t in p.iter(q("t")))
    hdr = next(i for i, p in enumerate(paras) if txt(p).strip() == "参考文献：")
    missing, changed = [], 0
    for p in paras[hdr + 1:]:
        t = txt(p).strip()
        if not re.match(r"^\[\d+\]", t):
            break
        m = detect_missing(t)
        if m:
            missing.append((len(missing) + 1, t, m))
        new = fix_one(t)
        if new == t:
            continue
        runs = p.findall(q("r"))
        wts = [w for w in runs[0].iter(q("t"))]
        if not wts:
            continue
        for r in runs[1:]:
            p.remove(r)
        for wt in wts[1:]:
            wt.getparent().remove(wt)
        wts[0].text = new
        changed += 1
        print(f"  ★ {t[:40]}…\n     →  {new}")
    if changed:
        members["word/document.xml"] = ET.tostring(
            root, encoding="utf-8", xml_declaration=True, standalone=True)
        fd, tmp = tempfile.mkstemp(dir=".", prefix=".tmpdocx_")
        with os.fdopen(fd, "wb") as f, zipfile.ZipFile(f, "w", zipfile.ZIP_DEFLATED) as zout:
            for n, d in members.items():
                zout.writestr(n, d)
        os.replace(tmp, path)
        print(f"docx 参考文献修复 {changed} 段 → 已就地写回 {path}")
    else:
        print("docx 参考文献均符合 GB/T 7714，无需修复")
    if list_missing:
        _report_missing(missing, missing_out)


def main():
    ap = argparse.ArgumentParser(description="参考文献 GB/T 7714 收尾核查")
    ap.add_argument("--content", help="content.json 路径（修其 refs 数组）")
    ap.add_argument("--docx", help="开题报告 .docx 路径（修参考文献段）")
    ap.add_argument("--list-missing", action="store_true",
                    help="列出待补真实数据的缺项（不编造），默认打印，可配合 --out 存档")
    ap.add_argument("--out", default=None, help="缺项清单写到此文件")
    args = ap.parse_args()

    if not args.content and not args.docx:
        ap.error("至少给 --content 或 --docx 之一")
    if args.content:
        fix_json(args.content, args.list_missing, args.out)
    if args.docx:
        fix_docx(args.docx, args.list_missing, args.out)


if __name__ == "__main__":
    main()
