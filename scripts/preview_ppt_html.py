#!/usr/bin/env python3
"""把 ppt_content.json 渲染成 HTML 幻灯片预览，供用户过目后再生成正式 PPTX。

解决"生成 PPTX 用户才看到效果、不满意来回改"的痛点：先用 HTML 把每一页
按版式画出来（北航蓝、16:9），浏览器里就能审——封面/目录/章节分隔/
大数字/编号卡片/对比表/流程图/甘特/阶段条/图片，全部按 layout 渲染。
用户确认（或要求改大纲）后，再跑 build_ppt_from_template.py 出正式 PPTX。

用法：
    python preview_ppt_html.py --content ppt_content.json --output 开题汇报预览.html
    open 开题汇报预览.html

ppt_content.json 结构同 build_ppt_from_template.py / derive_ppt_content.py
（chapters[{name, slides[{title, layout, bullets, extra}]}]）。
"""
from __future__ import annotations
import argparse, html, json, re
from pathlib import Path

ESC = html.escape


# ── 各版式的 HTML 片段（北航蓝 + 金色点缀，与 render_diagrams 同源配色）──
NAVY = "#003366"
BLUE = "#005BAC"
LIGHT = "#DCE9F7"
PALE = "#EEF4FB"
LINE = "#B9C8D9"
INK = "#2B2B2B"
MUTE = "#6B7A8C"
GOLD = "#C89A4B"

BASE_CSS = f"""
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ background:#e9edf2; font-family:"Microsoft YaHei","PingFang SC",sans-serif; }}
  .deck {{ max-width:1080px; margin:0 auto; padding:28px 0 60px; }}
  .deck-title {{ color:{NAVY}; font-size:22px; font-weight:700; padding:6px 12px 14px; }}
  .slide {{ width:960px; height:540px; margin:14px auto; background:#fff; border-radius:10px;
           box-shadow:0 4px 18px rgba(0,0,0,.10); position:relative; overflow:hidden;
           display:flex; flex-direction:column; }}
  .band {{ height:14px; background:{NAVY}; }}
  .band::after {{ content:""; display:block; height:4px; background:{GOLD}; }}
  .hd {{ padding:10px 44px 4px; }}
  .hd .tag {{ color:{BLUE}; font-size:12px; font-weight:700; }}
  .hd .title {{ color:{INK}; font-size:24px; font-weight:700; margin-top:2px; }}
  .hd .gold {{ width:54px; height:4px; background:{GOLD}; margin-top:6px; }}
  .body {{ flex:1; padding:10px 44px 30px; display:flex; flex-direction:column; }}
  .foot {{ position:absolute; left:44px; right:44px; bottom:18px; font-size:11px; color:{MUTE};
          display:flex; justify-content:space-between; border-top:1px solid #e2e8ef; padding-top:6px; }}
  .pageno {{ color:{NAVY}; font-weight:700; }}
  /* 编号卡片：≤4 条单列；>4 条自动两列网格。卡片自然高度不拉伸，铺不满就留白 */
  .cards-wrap {{ display:flex; flex-direction:column; gap:10px; }}
  .cards-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:10px 14px; align-content:start; }}
  .card {{ display:flex; gap:14px; background:{PALE}; border:1px solid {LINE}; border-radius:8px;
          padding:10px 16px; align-items:center; }}
  .card .no {{ width:30px; height:30px; background:{NAVY}; color:#fff; border-radius:7px;
              font-weight:700; display:flex; align-items:center; justify-content:center; flex:none; font-size:14px; }}
  .card .txt {{ font-size:13.5px; color:{INK}; line-height:1.45; }}
  .cards-grid .card {{ padding:8px 14px; }}
  .cards-grid .card .no {{ width:26px; height:26px; font-size:12.5px; }}
  .cards-grid .card .txt {{ font-size:12.5px; }}
  /* 双栏/多栏面板 */
  .panels {{ display:flex; gap:18px; flex:1; }}
  .panel {{ flex:1; background:{PALE}; border:1px solid {LINE}; border-radius:10px; overflow:hidden;
           display:flex; flex-direction:column; }}
  .panel .pt {{ background:{NAVY}; color:#fff; font-size:15px; font-weight:700; padding:12px 14px; }}
  .panel .pb {{ font-size:13px; color:{INK}; line-height:1.6; padding:14px 16px; }}
  /* 大数字 */
  .stats {{ display:grid; grid-template-columns:1fr 1fr; gap:20px; flex:1; align-content:center; }}
  .stat {{ background:{PALE}; border:1px solid {LINE}; border-radius:10px; text-align:center;
          padding:22px 10px; }}
  .stat .num {{ font-size:44px; font-weight:800; color:{NAVY}; }}
  .stat .lbl {{ font-size:12.5px; color:{MUTE}; margin-top:6px; line-height:1.4; }}
  /* 表格 */
  table {{ width:100%; border-collapse:collapse; font-size:13px; }}
  th {{ background:{NAVY}; color:#fff; padding:9px 10px; text-align:left; font-weight:700; }}
  td {{ border:1px solid #fff; padding:8px 10px; color:{INK}; vertical-align:top; }}
  tr:nth-child(even) td {{ background:{PALE}; }}
  tr:nth-child(odd) td {{ background:#f7f9fc; }}
  /* 流程图 */
  .flow {{ display:flex; flex-direction:column; gap:6px; flex:1; justify-content:center; }}
  .flow .row {{ display:flex; justify-content:center; gap:16px; align-items:center; }}
  .fbox {{ background:{LIGHT}; border:2px solid {NAVY}; border-radius:8px; padding:10px 14px;
          font-size:14px; font-weight:700; color:{NAVY}; text-align:center; }}
  .arrow {{ text-align:center; color:{NAVY}; font-size:16px; line-height:1; }}
  /* 甘特 */
  .gantt {{ display:flex; flex-direction:column; gap:8px; flex:1; }}
  .grow {{ display:grid; grid-template-columns:200px 1fr; align-items:center; gap:10px; }}
  .grow .phase {{ font-size:12.5px; font-weight:700; color:{NAVY}; text-align:right; }}
  .track {{ position:relative; height:26px; background:#f0f4f9; border-radius:6px; }}
  .bar {{ position:absolute; top:3px; bottom:3px; border-radius:5px; background:{BLUE};
          display:flex; align-items:center; padding-left:8px; color:#fff; font-size:10.5px;
          overflow:hidden; white-space:nowrap; }}
  .axis {{ font-size:10px; color:{MUTE}; display:flex; justify-content:space-between;
          padding-left:210px; }}
  /* 阶段条 */
  .pipe {{ display:flex; gap:0; flex:1; align-items:center; }}
  .stage {{ flex:1; background:{BLUE}; color:#fff; text-align:center; padding:14px 6px;
           clip-path:polygon(0 0, calc(100% - 14px) 0, 100% 50%, calc(100% - 14px) 100%, 0 100%, 14px 50%);
           margin-right:-14px; }}
  .stage:first-child {{ clip-path:polygon(0 0, calc(100% - 14px) 0, 100% 50%, calc(100% - 14px) 100%, 0 100%); }}
  .stage:last-child {{ clip-path:polygon(0 0, 100% 0, 100% 100%, 0 100%, 14px 50%); margin-right:0; }}
  .stage .lb {{ font-size:13px; font-weight:700; }}
  .stage .ds {{ font-size:9.5px; opacity:.85; margin-top:3px; }}
  .imgwrap {{ flex:1; display:flex; align-items:center; justify-content:center; }}
  .imgwrap img {{ max-width:100%; max-height:100%; border:1px solid {LINE}; border-radius:6px; }}
  .cap {{ text-align:center; font-size:11px; color:{MUTE}; margin-top:6px; }}
  /* 深蓝页 */
  .dark {{ background:{NAVY}; color:#fff; justify-content:center; }}
  .dark .band {{ display:none; }}
  .cover .ct {{ color:#cfe0f0; font-size:13px; margin-bottom:22px; }}
  .cover h1 {{ font-size:34px; font-weight:800; line-height:1.35; }}
  .cover .sub {{ color:#bdd3ea; font-size:16px; margin-top:12px; }}
  .cover .gold {{ width:80px; height:5px; background:{GOLD}; margin:22px 0; }}
  .cover .info {{ display:flex; gap:38px; font-size:13px; color:#e6eef6; margin-top:26px; }}
  .section .num {{ font-size:120px; color:{LIGHT}; font-weight:800; line-height:1; }}
  .section h2 {{ font-size:32px; font-weight:800; margin-top:4px; }}
  .section .en {{ color:#bdd3ea; font-size:15px; margin-top:8px; }}
  .section .gold {{ width:80px; height:5px; background:{GOLD}; margin-top:20px; }}
  .toc-item {{ display:flex; align-items:center; gap:14px; padding:7px 0; }}
  .toc-item .no {{ color:{LIGHT}; font-size:15px; font-weight:800; width:30px; }}
  .toc-item .nm {{ font-size:15px; font-weight:700; }}
  .toc-item .en {{ margin-left:auto; color:{LIGHT}; font-size:11.5px; }}
  .toc-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:4px 36px; }}
"""


def _month_idx(s: str) -> int | None:
    m = re.search(r"(\d{4})[.\-/年](\d{1,2})", s or "")
    if not m:
        return None
    return int(m.group(1)) * 12 + int(m.group(2))


def _slide_cover(data) -> str:
    cover = data.get("cover", {})
    info = "　".join(
        f"{k}：{v}" for k, v in
        [("汇报人", cover.get("作者姓名", "")), ("指导教师", cover.get("指导教师", "")),
         ("日期", cover.get("日期", ""))] if v)
    return f"""
  <div class="slide dark cover">
    <div class="body" style="justify-content:center;">
      <div class="ct">北京航空航天大学 · 硕士学位论文开题报告</div>
      <h1>{ESC(data.get('title',''))}</h1>
      <div class="sub">{ESC(data.get('subtitle',''))}</div>
      <div class="gold"></div>
      <div class="info">{ESC(info)}</div>
    </div>
    <div class="foot" style="color:{LIGHT};">{ESC(data.get('school','北京航空航天大学'))}<span class="pageno" style="color:#fff;">1</span></div>
  </div>"""


def _slide_toc(chapters) -> str:
    items = "".join(
        f'<div class="toc-item"><span class="no">{i+1:02d}</span>'
        f'<span class="nm">{ESC(c)}</span></div>'
        for i, c in enumerate(chapters))
    grid = f'<div class="toc-grid">{items}</div>' if len(chapters) > 6 else items
    return f"""
  <div class="slide dark" style="background:{BLUE};">
    <div class="body" style="justify-content:center; padding-top:40px;">
      <div style="font-size:28px; font-weight:800;">目录</div>
      <div style="color:#9ebede; font-size:13px; margin-top:4px;">CONTENTS</div>
      <div class="gold" style="width:56px; height:4px; background:{GOLD}; margin:12px 0 16px;"></div>
      {grid}
    </div>
  </div>"""


def _slide_section(idx, name) -> str:
    return f"""
  <div class="slide dark section">
    <div class="body" style="justify-content:center;">
      <div class="num">{idx:02d}</div>
      <h2>{ESC(name)}</h2>
      <div class="gold"></div>
    </div>
  </div>"""


def _slide_content(title, layout, bullets, extra, page_no, total, footer, tag="") -> str:
    body = ""
    if layout in ("stats",):
        stats = (extra or {}).get("stats", [])
        cards = "".join(
            f'<div class="stat"><div class="num">{ESC(str(s.get("number","")))}</div>'
            f'<div class="lbl">{ESC(str(s.get("label","")))}</div></div>'
            for s in stats)
        body = f'<div class="stats">{cards}</div>'
    elif layout == "cards":
        items = (extra or {}).get("items") or bullets or []
        cls = "cards-grid" if len(items) > 4 else "cards-wrap"
        cards = "".join(
            f'<div class="card"><div class="no">{i+1}</div>'
            f'<div class="txt">{ESC(str(b))}</div></div>'
            for i, b in enumerate(items))
        body = f'<div class="{cls}">{cards}</div>'
    elif layout == "panels":
        panels = (extra or {}).get("panels") or []
        p = "".join(
            f'<div class="panel"><div class="pt">{ESC(str(x.get("title","")))}</div>'
            f'<div class="pb">{ESC(str(x.get("text","")))}</div></div>'
            for x in panels[:4])
        body = f'<div class="panels">{p}</div>'
    elif layout in ("compare", "table"):
        extra = extra or {}
        headers = extra.get("headers") or []
        rows = extra.get("rows") or []
        head = "".join(f"<th>{ESC(str(h))}</th>" for h in headers)
        trs = "".join(
            "<tr>" + "".join(f"<td>{ESC(str(c))}</td>" for c in r) + "</tr>"
            for r in rows)
        body = f'<table><thead><tr>{head}</tr></thead><tbody>{trs}</tbody></table>'
    elif layout == "flow":
        nodes = extra.get("nodes") if extra else None
        body = _flow_html(nodes)
    elif layout == "gantt":
        body = _gantt_html(extra)
    elif layout == "pipeline":
        stages = (extra or {}).get("stages") or []
        st = "".join(
            f'<div class="stage"><div class="lb">{ESC(str(s.get("label","")))}</div>'
            f'<div class="ds">{ESC(str(s.get("desc","")))}</div></div>'
            for s in stages)
        body = f'<div class="pipe">{st}</div>'
    elif layout == "image_center":
        extra = extra or {}
        src = extra.get("image", "")
        cap = extra.get("caption", "")
        imgtag = f'<img src="file://{src}" alt="">' if src else "【图片】"
        body = f'<div class="imgwrap">{imgtag}</div>' + (f'<div class="cap">{ESC(cap)}</div>' if cap else "")
    else:
        items = bullets or []
        cls = "cards-grid" if len(items) > 4 else "cards-wrap"
        cards = "".join(
            f'<div class="card"><div class="no">{i+1}</div>'
            f'<div class="txt">{ESC(str(b))}</div></div>'
            for i, b in enumerate(items))
        body = f'<div class="{cls}">{cards}</div>' if cards else \
            "<div class='card'><div class='txt'>（本页要点待补充）</div></div>"
    tagline = f'<div class="tag">{ESC(tag)} · {ESC(layout)}</div>' if tag else ""
    return f"""
  <div class="slide">
    <div class="band"></div>
    <div class="hd">
      {tagline}
      <div class="title">{ESC(title)}</div>
      <div class="gold"></div>
    </div>
    <div class="body">{body}</div>
    <div class="foot"><span>{ESC(footer)}</span><span class="pageno">{page_no} / {total}</span></div>
  </div>"""


def _flow_html(nodes) -> str:
    if not nodes:
        return ""
    rows = []
    for n in nodes:
        if isinstance(n, list):
            boxes = "".join(f'<div class="fbox">{ESC(str(x))}</div>' for x in n)
            rows.append(f'<div class="row">{boxes}</div>')
        else:
            rows.append(f'<div class="row"><div class="fbox">{ESC(str(n))}</div></div>')
        rows.append('<div class="arrow">▼</div>')
    rows.pop()
    return f'<div class="flow">{"".join(rows)}</div>'


def _gantt_html(extra) -> str:
    tasks = (extra or {}).get("tasks") or []
    spans = []
    for t in tasks:
        time = str(t.get("time", ""))
        lo = _month_idx(time.split("–")[0].split("-")[0].split("—")[0])
        hi = _month_idx(time.split("–")[-1].split("-")[-1].split("—")[-1]) if ("–" in time or "-" in time or "—" in time) else lo
        if lo is None or hi is None:
            continue
        spans.append((t, lo, max(lo, hi)))
    if not spans:
        return ""
    lo0 = min(s[1] for s in spans) - 1
    hi0 = max(s[2] for s in spans) + 1
    total = max(hi0 - lo0, 1)
    rows = []
    for t, lo, hi in spans:
        left = (lo - lo0) / total * 100
        width = max((hi - lo) / total * 100, 3)
        rows.append(
            f'<div class="grow"><div class="phase">{ESC(str(t.get("phase","")))}</div>'
            f'<div class="track"><div class="bar" style="left:{left:.1f}%;width:{width:.1f}%;">'
            f'{ESC(str(t.get("content","")))}</div></div></div>')
    return f'<div class="gantt">{"".join(rows)}</div>'


def build(content_path: Path, output: Path):
    data = json.loads(content_path.read_text(encoding="utf-8"))
    chapters = data.get("chapters", [])
    footer = data.get("school", "北京航空航天大学")
    cover = data.get("cover", {})
    dept = cover.get("培养学院", "")
    if dept:
        footer += f" · {dept}"

    # 目录：章节名去重
    toc = []
    for ch in chapters:
        n = ch.get("name", "")
        if n and n not in toc:
            toc.append(n)

    slides = [_slide_cover(data)]
    slides.append(_slide_toc(toc))

    page_no = 3
    total = len(slides) + len(chapters) \
        + sum(len(ch.get("slides", [])) for ch in chapters) + 1
    for ci, ch in enumerate(chapters):
        slides.append(_slide_section(ci + 1, ch.get("name", "")))
        page_no += 1
        for spec in ch.get("slides", []):
            slides.append(_slide_content(
                spec.get("title", "") or ch.get("name", ""),
                spec.get("layout", "text_only"),
                spec.get("bullets", []),
                spec.get("extra", {}),
                page_no, total, footer, tag=ch.get("name", "")))
            page_no += 1

    # 致谢
    who = f"汇报人：{cover.get('作者姓名','')}"
    slides.append(f"""
  <div class="slide dark cover">
    <div class="body" style="justify-content:center; text-align:center;">
      <h1 style="text-align:center;">汇报完毕</h1>
      <div class="sub" style="text-align:center;">恳请各位老师批评指正！</div>
      <div class="gold" style="margin:22px auto;"></div>
      <div class="info" style="justify-content:center;">{ESC(who)}</div>
    </div>
  </div>""")

    doc = f"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{ESC(data.get('title',''))} · 开题汇报预览</title>
<style>{BASE_CSS}</style></head>
<body><div class="deck">
<div class="deck-title">{ESC(data.get('title',''))} · 开题汇报预览（HTML 版，确认后生成 PPTX）</div>
{''.join(slides)}
</div></body></html>"""

    output.write_text(doc, encoding="utf-8")
    print("saved:", output)
    print(f"  预览页数: {len(slides)}")


def main():
    ap = argparse.ArgumentParser(description="生成开题汇报 PPT 的 HTML 预览")
    ap.add_argument("--content", required=True, type=Path)
    ap.add_argument("--output", required=True, type=Path)
    a = ap.parse_args()
    build(a.content, a.output)


if __name__ == "__main__":
    main()
