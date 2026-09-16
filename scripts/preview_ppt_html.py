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


# ── 各版式的 HTML 片段（"批注过的图纸"：纸白 + 双阶蓝 + 赭金批注）──
NAVY = "#003366"
BLUE = "#005BAC"
LIGHT = "#DCE9F7"
PALE = "#EEF4FB"
LINE = "#C9D6E6"
PAPER = "#FBFCFE"
INK = "#2B2B2B"
MUTE = "#6B7A8C"
GOLD = "#B98A2F"

BASE_CSS = f"""
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  :root {{ --navy:{NAVY}; --blue:{BLUE}; --hair:{LINE}; --paper:{PAPER};
          --gold:{GOLD}; --ink:{INK}; --mute:{MUTE}; }}
  body {{ background:#E8EDF4; font-family:"Microsoft YaHei","PingFang SC",sans-serif; }}
  .serif {{ font-family:"Noto Serif SC","Songti SC","STSong","SimSun",serif; }}
  .deck {{ max-width:1080px; margin:0 auto; padding:28px 0 60px; }}
  .deck-title {{ color:var(--navy); font-size:21px; font-weight:700;
                font-family:"Noto Serif SC","Songti SC","STSong","SimSun",serif;
                padding:6px 12px 14px; }}
  .slide {{ width:960px; height:540px; margin:14px auto; background:#fff; border-radius:10px;
           box-shadow:0 4px 18px rgba(0,0,0,.10); position:relative; overflow:hidden;
           display:flex; flex-direction:column; }}
  .band {{ height:14px; background:var(--navy); }}
  .band::after {{ content:""; display:block; height:3px; background:var(--gold); }}
  .hd {{ padding:10px 44px 4px; }}
  .hd .tag {{ color:var(--blue); font-size:12px; font-weight:700; letter-spacing:.06em; }}
  .hd .title {{ color:var(--ink); font-size:25px; font-weight:700; margin-top:2px;
               font-family:"Noto Serif SC","Songti SC","STSong","SimSun",serif; }}
  .hd .gold {{ width:54px; height:3px; background:var(--gold); margin-top:7px; }}
  .body {{ flex:1; padding:10px 44px 30px; display:flex; flex-direction:column; }}
  .foot {{ position:absolute; left:44px; right:44px; bottom:18px; font-size:11px; color:var(--mute);
          display:flex; justify-content:space-between; border-top:1px solid #e2e8ef; padding-top:6px; }}
  .pageno {{ color:var(--navy); font-weight:700; }}
  /* 编号卡片 = 索引卡：纸白 + 发丝边 + 左侧靛青细线 + 细环序号；三色轮换已废除 */
  .cards-wrap {{ display:flex; flex-direction:column; gap:12px; flex:1; }}
  .cards-wrap .card {{ flex:1; min-height:0; }}
  .cards-grid {{ display:grid; grid-template-columns:1fr 1fr; grid-auto-rows:minmax(0,1fr);
                gap:12px 16px; flex:1; }}
  .card {{ position:relative; display:flex; gap:14px; background:var(--paper);
          border:1px solid var(--hair); border-radius:8px; padding:11px 16px 11px 20px;
          align-items:center; box-shadow:0 2px 8px rgba(31,45,61,.07); }}
  .card::before {{ content:""; position:absolute; left:0; top:10px; bottom:10px; width:3px;
                  border-radius:2px; background:var(--blue); }}
  .card .no {{ width:26px; height:26px; background:#fff; border:1.5px solid var(--blue);
              color:var(--navy); border-radius:50%; font-weight:700; display:flex;
              align-items:center; justify-content:center; flex:none; font-size:12.5px; }}
  .card .txt {{ font-size:13.5px; color:var(--ink); line-height:1.5; }}
  .card .txt b {{ color:var(--navy); }}
  .cards-grid .card {{ padding:9px 14px 9px 18px; }}
  .cards-grid .card .no {{ width:24px; height:24px; font-size:11.5px; }}
  .cards-grid .card .txt {{ font-size:12.5px; }}
  /* 两块大磁贴（恰好 2 条要点时） */
  .tiles {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; flex:1; }}
  .tile {{ background:var(--paper); border:1px solid var(--hair); border-radius:10px;
          box-shadow:0 2px 10px rgba(31,45,61,.08); padding:18px 22px 16px;
          display:flex; flex-direction:column; }}
  .tile .ghost {{ font-size:56px; font-weight:800; color:var(--hair); line-height:1;
                 font-family:Georgia,"Times New Roman",serif; }}
  .tile .tt {{ margin-top:auto; font-size:14px; color:var(--ink); line-height:1.55; }}
  .tile .tt b {{ color:var(--navy); }}
  .tile .ubar {{ width:64px; height:3px; border-radius:2px; background:var(--blue); margin-top:10px; }}
  /* 结论条 = 图纸标题块：墨蓝底 + 赭金内衬细线 + 戳记 */
  .takeaway {{ position:relative; background:var(--navy); color:#fff; border-radius:8px;
              padding:11px 18px; display:flex; align-items:center; gap:12px;
              font-size:14px; font-weight:700; box-shadow:0 2px 10px rgba(11,40,80,.22); }}
  .takeaway::after {{ content:""; position:absolute; inset:4px; border:1px solid rgba(185,138,47,.85);
                     border-radius:5px; pointer-events:none; }}
  .takeaway .tk {{ background:var(--gold); color:#fff; font-size:11px; font-weight:700;
                  padding:3px 10px; border-radius:3px; flex:none; position:relative; }}
  /* 双栏/多栏面板 */
  .panels {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(0,1fr));
            gap:18px; align-content:center; align-items:stretch; flex:1; }}
  .panel {{ background:var(--paper); border:1px solid var(--hair); border-radius:10px;
           overflow:hidden; display:flex; flex-direction:column;
           box-shadow:0 2px 8px rgba(31,45,61,.06); }}
  .panel .pt {{ background:var(--navy); color:#fff; font-size:15px; font-weight:700;
               padding:12px 14px; }}
  .panel .pb {{ font-size:13px; color:var(--ink); line-height:1.6; padding:14px 16px; }}
  /* 三线表 = 铺在衬纸上的学术表 */
  .sheet {{ background:var(--paper); border:1px solid var(--hair); border-radius:8px;
           box-shadow:0 2px 10px rgba(31,45,61,.08); padding:8px 14px; }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; }}
  th {{ background:none; color:var(--navy); padding:10px 10px; text-align:left;
       font-weight:700; border-top:2.5px solid var(--navy); border-bottom:1.2px solid var(--navy); }}
  td {{ border:none; border-bottom:.75px solid #DCE4EE; padding:9px 10px; color:var(--ink);
       vertical-align:top; background:none; }}
  tr:last-child td {{ border-bottom:2.5px solid var(--navy); }}
  /* 大数字 + 尺寸标注线（签名元素） */
  .stats {{ display:grid; grid-template-columns:1fr 1fr; gap:20px; flex:1; align-content:center; }}
  .stat {{ background:var(--paper); border:1px solid var(--hair); border-radius:10px;
          text-align:center; padding:24px 10px 18px; box-shadow:0 2px 8px rgba(31,45,61,.06); }}
  .stat .num {{ font-size:46px; font-weight:800; color:var(--navy);
               font-family:Georgia,"Noto Serif SC","Songti SC",serif; letter-spacing:.01em; }}
  .stat .dim {{ width:56%; height:2px; background:var(--hair); margin:10px auto 0;
               position:relative; }}
  .stat .dim::before, .stat .dim::after {{ content:""; position:absolute; top:-3px;
      width:2px; height:8px; background:var(--gold); }}
  .stat .dim::before {{ left:0; }}
  .stat .dim::after {{ right:0; }}
  .stat .lbl {{ font-size:12.5px; color:var(--mute); margin-top:9px; line-height:1.4; }}
  /* 流程图：纸白框 + 发丝蓝边 */
  .flow {{ display:flex; flex-direction:column; gap:5px; flex:1; justify-content:center; }}
  .flow .row {{ display:flex; justify-content:center; gap:16px; align-items:center; }}
  .fbox {{ background:var(--paper); border:1.5px solid var(--navy); border-radius:7px;
          padding:10px 14px; font-size:14px; font-weight:700; color:var(--navy);
          text-align:center; }}
  .arrow {{ text-align:center; color:var(--navy); font-size:15px; line-height:1; }}
  /* 甘特 = 绘图坐标纸 */
  .gantt {{ display:flex; flex-direction:column; gap:8px; flex:1; }}
  .grow {{ display:grid; grid-template-columns:200px 1fr; align-items:center; gap:10px; }}
  .grow .phase {{ font-size:12.5px; font-weight:700; color:var(--navy); text-align:right; }}
  .track {{ position:relative; height:26px; background:#F1F5FA; border-radius:3px;
           border:1px solid #E4EBF3; }}
  .bar {{ position:absolute; top:3px; bottom:3px; border-radius:2px; background:var(--blue);
          display:flex; align-items:center; padding-left:8px; color:#fff; font-size:10.5px;
          overflow:hidden; white-space:nowrap; }}
  .axis {{ font-size:10px; color:var(--mute); display:flex; justify-content:space-between;
          padding-left:210px; }}
  /* 阶段条 */
  .pipe {{ display:flex; gap:0; flex:1; align-items:center; }}
  .stage {{ flex:1; background:var(--blue); color:#fff; text-align:center; padding:14px 6px;
           clip-path:polygon(0 0, calc(100% - 14px) 0, 100% 50%, calc(100% - 14px) 100%, 0 100%, 14px 50%);
           margin-right:-14px; }}
  .stage:first-child {{ clip-path:polygon(0 0, calc(100% - 14px) 0, 100% 50%, calc(100% - 14px) 100%, 0 100%); }}
  .stage:last-child {{ clip-path:polygon(0 0, 100% 0, 100% 100%, 0 100%, 14px 50%); margin-right:0; }}
  .stage .lb {{ font-size:13px; font-weight:700; }}
  .stage .ds {{ font-size:9.5px; opacity:.85; margin-top:3px; }}
  .imgwrap {{ flex:1; display:flex; align-items:center; justify-content:center; }}
  .imgwrap img {{ max-width:100%; max-height:100%; border:1px solid var(--hair);
                 border-radius:6px; background:#fff; }}
  .cap {{ text-align:center; font-size:11px; color:var(--mute); margin-top:6px; }}
  /* 深蓝页：封面/章节加图纸内衬框 */
  .dark {{ background:var(--navy); color:#fff; justify-content:center; }}
  .dark .band {{ display:none; }}
  .dark::after {{ content:""; position:absolute; inset:16px; border:1px solid rgba(255,255,255,.24);
                 border-radius:4px; pointer-events:none; }}
  .cover .ct {{ color:#cfe0f0; font-size:13px; margin-bottom:22px; letter-spacing:.08em; }}
  .cover h1 {{ font-size:35px; font-weight:700; line-height:1.4;
              font-family:"Noto Serif SC","Songti SC","STSong","SimSun",serif; }}
  .cover .sub {{ color:#bdd3ea; font-size:16px; margin-top:12px; }}
  .cover .gold {{ width:80px; height:4px; background:var(--gold); margin:22px 0; }}
  .cover .info {{ display:flex; gap:38px; font-size:13px; color:#e6eef6; margin-top:26px; }}
  .section .num {{ font-size:118px; color:var(--hair); font-weight:800; line-height:1;
                  font-family:Georgia,"Times New Roman",serif; }}
  .section h2 {{ font-size:33px; font-weight:700; margin-top:4px;
                font-family:"Noto Serif SC","Songti SC","STSong","SimSun",serif; }}
  .section .en {{ color:#bdd3ea; font-size:15px; margin-top:8px; }}
  .section .gold {{ width:80px; height:4px; background:var(--gold); margin-top:20px; }}
  .toc-item {{ display:flex; align-items:center; gap:14px; padding:7px 0; }}
  .toc-item .no {{ color:var(--hair); font-size:15px; font-weight:800; width:30px;
                  font-family:Georgia,"Times New Roman",serif; }}
  .toc-item .nm {{ font-size:15px; font-weight:700; }}
  .toc-item .en {{ margin-left:auto; color:var(--hair); font-size:11.5px; }}
  .toc-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:4px 36px; }}
"""


def _month_idx(s: str) -> int | None:
    m = re.search(r"(\d{4})[.\-/年](\d{1,2})", s or "")
    if not m:
        return None
    return int(m.group(1)) * 12 + int(m.group(2))


def _card_txt(b) -> str:
    """卡片正文：「前缀：正文」→ 前缀加粗（前缀 ≤10 字才算）。"""
    b = str(b)
    if "：" in b:
        pre, rest = b.split("：", 1)
        if 0 < len(pre) <= 10:
            return f"<b>{ESC(pre)}：</b>{ESC(rest)}"
    return ESC(b)


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
            f'<div class="dim"></div>'
            f'<div class="lbl">{ESC(str(s.get("label","")))}</div></div>'
            for s in stats)
        body = f'<div class="stats">{cards}</div>'
    elif layout == "cards":
        items = (extra or {}).get("items") or bullets or []
        parts = []
        if len(items) == 2:
            tiles = "".join(
                f'<div class="tile" style="--ac:{[BLUE, NAVY, GOLD][i % 3]}">'
                f'<div class="ghost">{i+1:02d}</div>'
                f'<div class="tt">{_card_txt(b)}</div><div class="ubar"></div></div>'
                for i, b in enumerate(items))
            parts.append(f'<div class="tiles">{tiles}</div>')
        else:
            grid_items = items[:-1] if len(items) >= 3 else items
            if grid_items:
                cls = "cards-grid" if len(grid_items) > 4 else "cards-wrap"
                cards = "".join(
                    f'<div class="card"><div class="no">{i+1}</div>'
                    f'<div class="txt">{_card_txt(b)}</div></div>'
                    for i, b in enumerate(grid_items))
                parts.append(f'<div class="{cls}">{cards}</div>')
            if len(items) >= 3:
                parts.append(f'<div class="takeaway"><span class="tk">核心要点</span>'
                             f'<span>{_card_txt(items[-1])}</span></div>')
        body = "".join(parts)
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
        body = ('<div class="sheet"><table><thead><tr>' + head +
                '</tr></thead><tbody>' + trs + '</tbody></table></div>')
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
            f'<div class="txt">{_card_txt(b)}</div></div>'
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
