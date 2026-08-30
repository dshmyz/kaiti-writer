#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
参考文献质量闸：类型分布 / 近5年占比 / 中英构成 / 核心期刊命中 / 待核验残留
==========================================================================
用法：
  python3 check_refs_quality.py --content content.json
  python3 check_refs_quality.py --content content.json \
      --current-year 2026 --core-journals "中国行政管理,公共管理学报,法学研究"

输出逐项 ✅/⚠️，供 agent 在文献检索步骤（SKILL.md 步骤 2 质量闸）与
构建流程（build_and_validate Step 4）自动调用。只报告不修改。
退出码：0 全过 / 1 有 ⚠️（供 CI 式把关）。
"""
import argparse
import json
import re
import sys
from collections import OrderedDict

# 类型标识 → 展示名（按条内首次出现位置判定，避免 [EB/OL] 里嵌 [J] 误判）
TYPE_PATTERNS = OrderedDict([
    ("期刊[J]", re.compile(r"\[J\]")),
    ("会议[C]", re.compile(r"\[C\]")),
    ("专著[M]", re.compile(r"\[M\]")),
    ("学位论文[D]", re.compile(r"\[D\]")),
    ("报告[R]", re.compile(r"\[R\]")),
    ("政策/电子[EB/OL]", re.compile(r"\[EB/OL\]")),
    ("标准[S]", re.compile(r"\[S\]")),
    ("报纸[N]", re.compile(r"\[N\]")),
    ("其他", re.compile(r"\[")),
])
PEER_REVIEW = {"期刊[J]", "会议[C]", "专著[M]", "学位论文[D]"}


def year_of(entry):
    m = re.search(r"(19|20)\d{2}", entry)
    return int(m.group(0)) if m else None


def main():
    ap = argparse.ArgumentParser(description="参考文献质量闸（只报告不修改）")
    ap.add_argument("--content", required=True, help="content.json 路径")
    ap.add_argument("--current-year", type=int, default=None, help="当前年份（默认取系统时间）")
    ap.add_argument("--core-journals", default="",
                    help="逗号分隔的本学科核心期刊名，命中计数（学科配比判断依据）")
    args = ap.parse_args()

    import datetime
    cy = args.current_year or datetime.date.today().year
    core_list = [j.strip() for j in args.core_journals.split(",") if j.strip()]

    with open(args.content, encoding="utf-8") as f:
        refs = json.load(f).get("refs", [])
    n = len(refs)
    if not n:
        print("⚠️ refs 为空")
        sys.exit(1)

    issues = []
    print(f"参考文献共 {n} 条，当前年份 {cy}")

    # 1) 类型分布
    type_count = OrderedDict((name, 0) for name in TYPE_PATTERNS)
    peer = 0
    for r in refs:
        for name, pat in TYPE_PATTERNS.items():
            if pat.search(r):
                type_count[name] += 1
                if name in PEER_REVIEW:
                    peer += 1
                break
    print("类型分布:", "，".join(f"{k}×{v}" for k, v in type_count.items() if v))
    peer_pct = round(peer / n * 100, 1)
    if peer_pct >= 60:
        print(f"✅ 同行评议文献占比 {peer_pct}%（{peer}/{n}，阈值 60%）")
    else:
        issues.append(f"同行评议文献占比仅 {peer_pct}%（{peer}/{n}，阈值 60%）——补充期刊/会议/专著文献")

    # 2) 近 5 年占比
    years = [year_of(r) for r in refs]
    recent = sum(1 for y in years if y and y >= cy - 4)
    recent_pct = round(recent / n * 100, 1)
    if recent_pct >= 50:
        print(f"✅ 近 5 年占比 {recent_pct}%（{recent}/{n}，阈值 50%）")
    else:
        issues.append(f"近 5 年占比仅 {recent_pct}%（{recent}/{n}，阈值 50%）——补充近年文献")

    # 3) 中英构成
    cjk = re.compile(r"[一-鿿]")
    cn = sum(1 for r in refs if cjk.search(r))
    print(f"中英构成: 中文 {cn} / 外文 {n - cn}")

    # 4) 核心期刊命中（学科配比的客观证据，最终判断由 agent 结合选题给出）
    if core_list:
        hits = []
        for r in refs:
            for j in core_list:
                if j in r:
                    hits.append(j)
                    break
        print(f"核心期刊命中 {len(hits)}/{n}：{'，'.join(hits) if hits else '无'}")
        if len(hits) < 3:
            issues.append(f"本学科核心期刊命中仅 {len(hits)} 条（建议 ≥3）——学科配比可能被评委质疑")

    # 5) 待核验残留
    pending = [i + 1 for i, r in enumerate(refs) if re.search(r"待核验|待补|待验证|TODO", r)]
    if pending:
        issues.append(f"第 {pending} 条仍含'待核验/待补'占位——定稿前必须清零，不留用户侧待办")
    else:
        print("✅ 无'待核验/待补'残留")

    print("\n结论:", "✅ 质量闸通过" if not issues else f"⚠️ {len(issues)} 项待改进")
    for it in issues:
        print(f"  - {it}")
    sys.exit(0 if not issues else 1)


if __name__ == "__main__":
    main()
