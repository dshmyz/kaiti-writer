#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""参考文献表 GB/T 7714 收尾核查与修复。

生成开题报告参考文献（写进 content.json 的 refs、进而转成 docx）后，
在交付前对每条目做收尾核查：去末尾结束符、清"待核验/占位"括号（缺项连标点略去）、
统一半角标点。另有"缺项就地红色标注"，规则出处见 references/参考文献著录.md
「写入前收尾核查」与「缺项的就地标注」。

用法：
    # 修 content.json 的 refs（顺便列出待补真实数据的缺项清单）
    python fix_refs.py --content content.json --list-missing --out missing.txt

    # 修已有 docx 的参考文献段（需 lxml；就地改，先自备备份）
    python fix_refs.py --docx 开题报告-xxx.docx

    # 待核验阶段：给 docx 里缺项的条目就地加红色提示（content.json 保持干净）
    python fix_refs.py --docx 开题报告-xxx.docx \
        --mark-missing "6:发文字号待补;11:作者与卷期页待核验;12:出版社与年份待核验"

    # 定稿：清掉 docx 里所有红色待补提示 run，恢复完全合规
    python fix_refs.py --docx 开题报告-xxx.docx --unmark-red

不依赖第三方库（json 模式）；docx 模式需要 lxml（缺则报错并提示 pip install lxml）。
"""
from __future__ import annotations
import argparse, json, os, re, sys, tempfile, zipfile


def fix_one(s: str) -> str:
    """对单条参考文献做 GB/T 7714 收尾：去占位括号、去末尾结束符、半角化孤立尾标点。"""
    # 1) 删含"待核验/待补/待验证/TODO"的成对括号（缺口径，避免误删真实期号/页码，如 38(2)）
    s = re.sub(r"[（(][^（）()]*?(?:待核验|待补|待验证|TODO)[^（）()]*?[）)]", "", s)
    # 2) 反复消化条目尾部：先切掉尾空白，再去末尾结束符/多余标点（GB/T 7714：每条最后不用加结束符）
    #    循环到稳定，因为删占位可能残留"标点+空格+标点"之类（如 "[Z]. （待核验）." 删占位后剩 "[Z]. ."）
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r"\s+$", "", s)
        s = re.sub(r"[.,;:、，。；：]+$", "", s)
    # 3) 孤立的类型标识尾部句点，如 "[Z]." -> "[Z]"（缺项连同标点一并略去）
    s = re.sub(r"(\[[A-Z/]+\])\s*[.,;:]\s*$", r"\1", s)
    return s.strip()


def detect_missing(s: str) -> str | None:
    """返回条目里的缺项原文（含"待核验/待补"的括号），无则 None。"""
    m = re.search(r"[（(][^（）()]*?(?:待核验|待补|待验证|TODO)[^（）()]*?[）)]", s)
    return m.group(0) if m else None


# ---------- content.json ----------
def fix_json(path: str, list_missing: bool, missing_out: str) -> None:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    refs = data.get("refs")
    if not isinstance(refs, list):
        print("⚠️  json 无 refs 数组，跳过")
        return
    missing, changed = [], 0
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


# ---------- docx 公共工具 ----------
def _load_docx(path):
    try:
        from lxml import etree as ET
    except ImportError:
        sys.exit("缺少 lxml，请先安装：pip install lxml（处理 docx 必需）")
    with zipfile.ZipFile(path) as z:
        members = {n: z.read(n) for n in z.namelist()}
    root = ET.fromstring(members["word/document.xml"])
    return ET, root, members


def _write_docx(members, path, root):
    from lxml import etree as ET
    members["word/document.xml"] = ET.tostring(
        root, encoding="utf-8", xml_declaration=True, standalone=True)
    fd, tmp = tempfile.mkstemp(dir=".", prefix=".tmpdocx_")
    with os.fdopen(fd, "wb") as f, zipfile.ZipFile(f, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, d in members.items():
            zout.writestr(n, d)
    os.replace(tmp, path)


def _ref_paras(ET, root, q):
    body = root.find(q("body"))
    paras = body.findall(q("p"))
    txt = lambda p: "".join(t.text or "" for t in p.iter(q("t")))
    hdr = next((i for i, p in enumerate(paras) if txt(p).strip() == "参考文献："), None)
    if hdr is None:
        return []
    refs = []
    for p in paras[hdr + 1:]:
        t = txt(p).strip()
        if not re.match(r"^\[\d+\]", t):
            break
        refs.append(p)
    return refs


def fix_docx(path: str, list_missing: bool, missing_out: str) -> None:
    ET, root, members = _load_docx(path)
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    q = lambda t: f"{{{W}}}{t}"
    missing, changed = [], 0
    for p in _ref_paras(ET, root, q):
        t = "".join(x.text or "" for x in p.iter(q("t"))).strip()
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
        print(f"  ★ {t[:38]}…\n     →  {new}")
    if changed:
        _write_docx(members, path, root)
        print(f"docx 参考文献修复 {changed} 段 → 已就地写回 {path}")
    else:
        print("docx 参考文献均符合 GB/T 7714，无需修复")
    if list_missing:
        _report_missing(missing, missing_out)


def mark_missing_docx(path: str, mapping: str) -> None:
    """对给定 '序号:提示文案;…' 映射，在 docx 对应参考文献条目末尾就地加红色(FF0000) run。"""
    ET, root, members = _load_docx(path)
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    q = lambda t: f"{{{W}}}{t}"
    RED = "FF0000"
    marks = {}
    for seg in mapping.split(";"):
        seg = seg.strip()
        if not seg or ":" not in seg:
            continue
        n, tip = seg.split(":", 1)
        if n.strip().isdigit():
            marks[int(n.strip())] = tip.strip()
    added = 0
    for p in _ref_paras(ET, root, q):
        t = "".join(x.text or "" for x in p.iter(q("t"))).strip()
        m = re.match(r"^\[(\d+)\]", t)
        if not m or int(m.group(1)) not in marks:
            continue
        n = int(m.group(1))
        # 跳过已带红色提示的条目，避免重复
        already = any(
            (r.find(q("rPr")) is not None and r.find(q("rPr")).find(q("color")) is not None
             and r.find(q("rPr")).find(q("color")).get(q("val")) == RED)
            for r in p.findall(q("r")))
        if already:
            print(f"  [{n}] 已有红色提示，跳过")
            continue
        runs = p.findall(q("r"))
        if not runs:
            continue
        base_rpr = runs[0].find(q("rPr"))
        new_r = ET.SubElement(p, q("r"))
        if base_rpr is not None:
            from copy import deepcopy
            rpr = deepcopy(base_rpr)
            c = rpr.find(q("color"))
            if c is not None:
                rpr.remove(c)
            ET.SubElement(rpr, q("color")).set(q("val"), RED)
            new_r.append(rpr)
        else:
            rpr = ET.SubElement(new_r, q("rPr"))
            ET.SubElement(rpr, q("color")).set(q("val"), RED)
        t_el = ET.SubElement(new_r, q("t"))
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_el.text = " 〈" + marks[n] + "〉"
        added += 1
        print(f"  [{n}] 就地红色提示 →  〈{marks[n]}〉")
    if added:
        _write_docx(members, path, root)
        print(f"docx 已就地标红 {added} 条缺项（content.json 保持干净）。定稿前用 --unmark-red 清除。")
    else:
        print("未新增红色标注")


def unmark_red_docx(path: str) -> None:
    """定稿用：清除 docx 中由本脚本加的红色(FF0000)待补提示 run，恢复单色合规。

    只删颜色恰为 FF0000（或 FFFF0000）的 run，避免误删参考文献里本来就着色的文字
    （如黑色显式着色、脚注链接色等）。只作用于参考文献区（见 _ref_paras）。
    """
    ET, root, members = _load_docx(path)
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    q = lambda t: f"{{{W}}}{t}"
    is_red = lambda col: col is not None and col.get(q("val"), "").upper() in ("FF0000", "FFFF0000")
    removed = 0
    for p in _ref_paras(ET, root, q):
        for r in list(p.findall(q("r"))):
            rpr = r.find(q("rPr"))
            if is_red(rpr.find(q("color")) if rpr is not None else None):
                p.remove(r)
                removed += 1
    if removed:
        _write_docx(members, path, root)
        print(f"已清除 {removed} 个红色待补 run → 参考文献可定稿")
    else:
        print("没有可清除的红色待补 run")


def clear_noise_red_docx(path: str, keep_marker: str = "〈待确认"):
    """清全文档(正文/目录/附录)里的"噪声红"，只保留带保留标记的确认红。

    场景：正文里残留模板自带的目录缓存红字、误染红的文字等，统一清掉；而
    你**要留给用户确认**的红用统一的保留标记（默认 `〈待确认`）标出，这一类
    不删。与 `--unmark-red`（只清参考文献区待补红）互补、不冲突：
        --unmark-red          只清参考文献区，删 FF0000 红（含标记红）
        --clear-noise-red     清正文/目录/附录区，删 FF0000 但【不含】保留标记的红
    参考文献区（"参考文献："之后、以 [n] 开头的段）整体跳过，避免撞车。
    保留判定用"run 文本以 keep_marker 开头"（前缀，默认 `〈待确认`）：
    `〈待确认〉`、`〈待确认：数据口径〉`、`〈待确认xxx〉` 都命中保留；严格子串
    会把 `〈待确认：…〉` 当噪声误删，故用前缀而非整串。
    只删有文字的 run：空 run 与组织结构 run 不动，降低误删风险。
    """
    ET, root, members = _load_docx(path)
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    q = lambda t: f"{{{W}}}{t}"
    is_red = lambda col: col is not None and col.get(q("val"), "").upper() in ("FF0000", "FFFF0000")
    # 参考文献区段的集合（该区归 --unmark-red 管，这里跳过）
    ref_paras = set(id(p) for p in _ref_paras(ET, root, q))
    removed = 0
    for body_p in root.iter(q("p")):
        if id(body_p) in ref_paras:
            continue
        for r in list(body_p.findall(q("r"))):
            rpr = r.find(q("rPr"))
            col = rpr.find(q("color")) if rpr is not None else None
            if not is_red(col):
                continue
            rtext = "".join(t.text or "" for t in r.iter(q("t")))
            if not rtext.strip():
                continue            # 空 run / 结构 run 不动
            if rtext.lstrip().startswith(keep_marker):
                continue            # 带保留标记（前缀）的确认红，保留
            body_p.remove(r)
            removed += 1
    if removed:
        _write_docx(members, path, root)
        print(f"已清除 {removed} 个非确认红 run（保留含 {keep_marker} 前缀的红）→ 正文噪声红已清理")
    else:
        print("未发现需清除的非确认红，或正文已干净")


def main():
    ap = argparse.ArgumentParser(description="参考文献 GB/T 7714 收尾核查与缺项红色标注")
    ap.add_argument("--content", help="content.json 路径（修其 refs 数组）")
    ap.add_argument("--docx", help="开题报告 .docx 路径")
    ap.add_argument("--list-missing", action="store_true",
                    help="列出待补真实数据的缺项（不编造）")
    ap.add_argument("--out", default=None, help="缺项清单写到此文件")
    ap.add_argument("--mark-missing", default=None, metavar='"6:发文字号待补;11:…"',
                    help="给 docx 缺项条目就地加红色提示（序号:提示文案，分号分隔）")
    ap.add_argument("--unmark-red", action="store_true",
                    help="定稿：清除 docx 参考文献区里所有红色待补提示 run")
    ap.add_argument("--clear-noise-red", action="store_true",
                    help="清全文档(正文/目录/附录)非确认红，只保留含〈待确认〉标记的红（跳过参考文献区）")
    args = ap.parse_args()

    if args.mark_missing:
        mark_missing_docx(args.docx, args.mark_missing)
    elif args.unmark_red:
        unmark_red_docx(args.docx)
    elif args.clear_noise_red:
        if not args.docx:
            ap.error("--clear-noise-red 需要 --docx")
        clear_noise_red_docx(args.docx)
    else:
        if not args.content and not args.docx:
            ap.error("至少给 --content 或 --docx 之一")
        if args.content:
            fix_json(args.content, args.list_missing, args.out)
        if args.docx:
            fix_docx(args.docx, args.list_missing, args.out)


if __name__ == "__main__":
    main()
