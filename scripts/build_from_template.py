#!/usr/bin/env python3
"""基于官方 .docx 模板生成开题报告：深拷贝模板段落 pPr，只换 run 文本，保留全部样式。
仿 buaa-final-assignment 的 build_from_template 思路。"""
from __future__ import annotations
import argparse, json
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
try:
    from lxml import etree as ET
except ImportError:
    raise SystemExit("缺少 lxml，请先安装：pip install lxml（lxml 原生保留命名空间前缀，"
                     "stdlib ElementTree 会把 wp/r/mc/w14 等前缀改名成 ns0/ns1，破坏 Word 文件）")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def q(t): return f"{{{W}}}{t}"

# 字号（半磅）：小四=24，五号=21，四号=28
SZ_BODY, SZ_FIVE, SZ_FOUR = 24, 21, 28

# 允许自动新建的节 → 插在哪一段之前。
# 官方模板正文只有一~四章，缺《管理办法》第三条要求的「预期目标和成果」（模板里"预期"出现 0 次）。
# 白名单故意是闭集：键名打错时仍报「未找到节标题」，不会静默建出错节。
CREATABLE_SECTIONS = {"五、预期目标和成果": "参考文献："}

def text_of(p): return "".join(t.text or "" for t in p.iter(q("t")))

def is_empty_p(p):
    """空段：无文字、无续接图形（表格单元格/绘图/嵌入对象）——用于清理多余空行。"""
    if p.tag != q("p"): return False
    if text_of(p).strip(): return False
    # 带 sectPr 的段落是分节符（封面→目录→正文的换页靠它），绝不能当空段删
    ppr = p.find(q("pPr"))
    if ppr is not None and ppr.find(q("sectPr")) is not None: return False
    # 带 drawing/pict（图片、文本框）或 bookmarkStart 起定位作用的段落不算“空”，不能随手删
    if p.find(q("r") + "/" + q("drawing")) is not None: return False
    if p.find(q("r") + "/" + q("pict")) is not None: return False
    if p.find(q("bookmarkStart")) is not None: return False
    return True

def find_para(body, needle):
    for p in body.findall(q("p")):
        if needle in text_of(p): return p
    return None

def find_para_exact(body, needle):
    """精确匹配（去空白后相等），用于定位正文“开题报告”标题，避免命中封面“专业硕士开题报告”。"""
    for p in body.findall(q("p")):
        if text_of(p).strip() == needle: return p
    return None

def trim_toc_field_end(body):
    """清掉目录 TOC 字段末尾的红色模板说明行。

    模板目录是一个 TOC 域，其"结果"缓存里带一段样板条目；字段末尾（fldChar=end
    所在段）常缀一句红字帮助语，如「附录（若论文包含访谈或问卷，需提供访谈、问卷提纲）」。
    这句既不是用户自己的内容，也不该出现在目录里（它是给作者看的说明），而且 Word
    打开时 F9 会用真实标题重建整个目录。这里把该段的可见文字（保留 fldChar=end run）
    清成空，避免用户看到模板残留。
    """
    end_p = None
    for p in body.findall(q("p")):
        if p.find(f"{q('r')}/{q('fldChar')}") is not None:
            for fc in p.iter(q("fldChar")):
                if fc.get(q("fldCharType")) == "end":
                    end_p = p
    if end_p is None:
        return False
    # 保留仅含 fldChar=end 的 run，删掉其余文本/tab run
    for r in list(end_p.findall(q("r"))):
        if r.find(q("fldChar")) is not None:
            continue
        end_p.remove(r)
    return True

def find_para_after(body, needle, start_p):
    """在 start_p 之后的 body 子元素中查找含 needle 的 <w:p>（避免命中目录 TOC）。用 lxml getnext 链遍历兄弟。"""
    sib = start_p.getnext()
    while sib is not None:
        if sib.tag == q("p") and needle in text_of(sib):
            return sib
        sib = sib.getnext()
    return None

def first_run_rpr(p):
    r = p.find(q("r"))
    if r is not None:
        rpr = r.find(q("rPr"))
        if rpr is not None: return deepcopy(rpr)
    return None

def next_para_sibling(body, p):
    """返回 p 的下一个 <w:p> 兄弟，无则 None（用 lxml getnext）。"""
    sib = p.getnext()
    while sib is not None:
        if sib.tag == q("p"):
            return sib
        sib = sib.getnext()
    return None

def insert_after(body, ref, new_el):
    """把 new_el 插到 ref 之后（lxml 原生 addnext）。"""
    ref.addnext(new_el)

def set_run_text(p, text, *, eastasia="宋体", ascii_font="Times New Roman",
                 sz_halfpt=None, bold=None):
    """清空 p 的 runs，新建一个 run，复用 p 原 rPr 或新建，按参数覆盖字体字号。"""
    rpr = first_run_rpr(p)
    if rpr is None:
        rpr = ET.Element(q("rPr"))
    # eastAsia / ascii
    rf = rpr.find(q("rFonts"))
    if rf is None:
        rf = ET.Element(q("rFonts")); rpr.insert(0, rf)
    if eastasia: rf.set(q("eastAsia"), eastasia)
    if ascii_font:
        rf.set(q("ascii"), ascii_font); rf.set(q("hAnsi"), ascii_font)
    # 字号
    if sz_halfpt is not None:
        for stag in (q("sz"), q("szCs")):
            for old in list(rpr.findall(stag)): rpr.remove(old)
            s = ET.Element(stag); s.set(q("val"), str(sz_halfpt)); rpr.append(s)
    # 粗体
    if bold is not None:
        for btag in (q("b"), q("bCs")):
            for old in list(rpr.findall(btag)): rpr.remove(old)
            b = ET.Element(btag)
            if not bold: b.set(q("val"), "0")
            rpr.append(b)
    # 清旧 runs
    for r in list(p.findall(q("r"))): p.remove(r)
    r = ET.Element(q("r")); r.append(rpr)
    t = ET.Element(q("t"))
    if text[:1].isspace() or text[-1:].isspace():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text; r.append(t); p.append(r)

def set_outline(p, level):
    """给段落设大纲级别 w:outlineLvl（0 章 / 1 节 / 2 条），使 TOC 域 (TOC \\o "1-4") 收录。

    模板的节/条/参考文献/附录标题没有标题样式；TOC 域按大纲级别/标题样式收集条目，
    不设 outlineLvl 的话一旦刷新目录域(F9)，这些项就从目录里消失。
    outlineLvl 在 CT_PPr 序列里位于 jc 之后、rPr 之前。
    """
    pPr = p.find(q("pPr"))
    if pPr is None:
        pPr = ET.Element(q("pPr")); p.insert(0, pPr)
    ol = pPr.find(q("outlineLvl"))
    if ol is None:
        ol = ET.Element(q("outlineLvl"))
        anchor = next((e for e in pPr if e.tag in (q("rPr"), q("sectPr"))), None)
        if anchor is not None:
            anchor.addprevious(ol)
        else:
            pPr.append(ol)
    ol.set(q("val"), str(level))
    return p


def set_spacing(p, *, line=None, before_lines=None, after_lines=None):
    """按书写规范显式强制段落的行距与段前段后，覆盖模板继承的旧值。

    格式规范（北航 1.2 页面设置）：
      章/节/条三级标题 → 单倍行距（line=240），段前、段后各 0.5 行（beforeLines/afterLines=50）
      正文           → 1.5 倍行距（line=360），段前、段后 0 行
    line 为 None 则不写行距（保持单倍自动）；用 beforeLines/afterLines 表达"行"，
    与 Word 的"段前段后 X 行"一致，不写 before/after 的 twips 数值。
    """
    pPr = p.find(q("pPr"))
    if pPr is None:
        pPr = ET.Element(q("pPr")); p.insert(0, pPr)
    sp = pPr.find(q("spacing"))
    if sp is None:
        # CT_PPr 顺序：spacing → ind → jc → rPr，插到 ind/jc/rPr 三者中靠前者的前面
        sp = ET.Element(q("spacing"))
        anchor = next((e for e in pPr if e.tag in (q("ind"), q("jc"), q("rPr"))), None)
        if anchor is not None:
            anchor.addprevious(sp)
        else:
            pPr.append(sp)
    # 重建 spacing，避免残留模板/示例的 before/after/line 干扰（这些是 XML 属性，须按名清除）
    for k in ("before", "after", "line", "lineRule", "beforeLines", "afterLines",
              "beforeAutospacing", "afterAutospacing"):
        sp.attrib.pop(q(k), None)
    if line is not None:
        sp.set(q("line"), str(line)); sp.set(q("lineRule"), "auto")
    if before_lines is not None:
        sp.set(q("beforeLines"), str(before_lines))
    if after_lines is not None:
        sp.set(q("afterLines"), str(after_lines))
    return p


def clone_body_para(template_body_p, text):
    """深拷贝正文模板段（保留 pPr 缩进/段落样式），注入正文 run，并强制规范行距。

    正文：1.5 倍行距（line=360）、段前段后 0 行（格式规范 1.2）。deepcopy 继承的
    模板 before/after 与行距一律以这里显式设置的为准，避免把示例段旧值带进全文。
    """
    p = deepcopy(template_body_p)
    for r in list(p.findall(q("r"))): p.remove(r)
    set_spacing(p, line=360, before_lines=0, after_lines=0)
    set_run_text(p, text, eastasia="宋体", ascii_font="Times New Roman", sz_halfpt=SZ_BODY)
    return p

def make_heading_para(text, heading_tpl):
    """克隆章标题段（保留 pStyle=3 以便 F9 后进目录），只换文本。

    字体字号一律沿用样板段自己的 rPr：模板一~四章的 run 实测都是 sz=24、
    rFonts 只有 hint=eastAsia、不加粗。写死宋体/四号/加粗会让新建的「五、」
    和其他章长得不一样，所以这里先取样板 run 的 rPr 再清 run。
    """
    p = deepcopy(heading_tpl)
    rpr = first_run_rpr(p)
    rpr = deepcopy(rpr) if rpr is not None else None
    for r in list(p.findall(q("r"))): p.remove(r)
    # 章/节/条标题：单倍行距（line=240），段前、段后各 0.5 行（格式规范 1.2）
    set_spacing(p, line=240, before_lines=50, after_lines=50)
    run = ET.SubElement(p, q("r"))
    if rpr is not None:
        run.append(rpr)
    t = ET.SubElement(run, q("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return p


def ensure_section(body, title, before_needle, body_start, heading_tpl):
    """确保正文里存在 title 这一节标题；模板没有就新建（插在 before_needle 段之前）。

    幂等：已存在则直接返回它，不会建出第二个。
    模板原文写有「可按照导师建议进行调整」，故新增章节符合模板意图。
    """
    if body_start is None or heading_tpl is None:
        return None
    hp = find_para_after(body, title, body_start)
    if hp is not None:
        return hp
    ref = find_para_after(body, before_needle, body_start)
    if ref is None:
        return None
    hp = make_heading_para(title, heading_tpl)
    ref.addprevious(hp)
    return hp


def _border(tag, sz="8"):
    b = ET.Element(q(tag))
    b.set(q("val"), "single"); b.set(q("sz"), sz); b.set(q("space"), "0"); b.set(q("color"), "000000")
    return b


def make_caption(text, kind, num, *, body_tpl=None):
    """图题/表题：五号宋体加粗居中，编号与题间空半角 2 格（格式规范：表题在表上、图题在图下）。"""
    p = deepcopy(body_tpl) if body_tpl is not None else ET.Element(q("p"))
    for r in list(p.findall(q("r"))): p.remove(r)
    pPr = p.find(q("pPr"))
    if pPr is None:
        pPr = ET.Element(q("pPr")); p.insert(0, pPr)
    for tag in (q("ind"), q("jc")):
        for old in list(pPr.findall(tag)): pPr.remove(old)
    # 居中段（表题/图题/插图）不再继承正文模板的 before/after 行距：单倍、段前段后 0
    set_spacing(p, line=240, before_lines=0, after_lines=0)
    jc = ET.SubElement(pPr, q("jc")); jc.set(q("val"), "center")
    if (rpr := pPr.find(q("rPr"))) is not None:
        pPr.remove(jc); rpr.addprevious(jc)
    set_run_text(p, f"{kind}{num}  {text}", eastasia="宋体",
                 ascii_font="Times New Roman", sz_halfpt=SZ_FIVE, bold=True)
    return p


def _png_size(raw):
    if raw[:8] != b"\x89PNG\r\n\x1a\n": return None
    import struct
    w, h = struct.unpack(">II", raw[16:24])
    return w, h


def _jpeg_size(raw):
    import struct
    if raw[:2] != b"\xff\xd8": return None
    i = 2
    while i < len(raw) - 9:
        if raw[i] != 0xFF: i += 1; continue
        marker = raw[i + 1]
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB):
            h, w = struct.unpack(">HH", raw[i + 5:i + 9])
            return w, h
        if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
            i += 2; continue
        i += 2 + struct.unpack(">H", raw[i + 2:i + 4])[0]
    return None


EMU_PER_CM = 360000


def image_extent(raw, max_width_cm=14.0):
    """按图片像素算显示尺寸（等比缩放到不超过 max_width_cm），返回 EMU。"""
    size = _png_size(raw) or _jpeg_size(raw)
    max_w = int(max_width_cm * EMU_PER_CM)
    if not size:
        return max_w, int(max_w * 0.6)
    px_w, px_h = size
    cx = min(int(px_w / 96 * 2.54 * EMU_PER_CM), max_w)
    return cx, max(1, int(cx * px_h / px_w))


def make_image_para(rid, cx, cy, name="image", *, body_tpl=None):
    """居中的内嵌图片段落（w:drawing/wp:inline）。"""
    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    p = deepcopy(body_tpl) if body_tpl is not None else ET.Element(q("p"))
    for r in list(p.findall(q("r"))): p.remove(r)
    pPr = p.find(q("pPr"))
    if pPr is None:
        pPr = ET.Element(q("pPr")); p.insert(0, pPr)
    for tag in (q("ind"), q("jc")):
        for old in list(pPr.findall(tag)): pPr.remove(old)
    # 居中段（表题/图题/插图）不再继承正文模板的 before/after 行距：单倍、段前段后 0
    set_spacing(p, line=240, before_lines=0, after_lines=0)
    jc = ET.SubElement(pPr, q("jc")); jc.set(q("val"), "center")
    if (rpr := pPr.find(q("rPr"))) is not None:
        pPr.remove(jc); rpr.addprevious(jc)
    run = ET.SubElement(p, q("r"))
    drawing = ET.SubElement(run, q("drawing"))
    inline = ET.SubElement(drawing, f"{{{WP}}}inline")
    for k in ("distT", "distB", "distL", "distR"):
        inline.set(k, "0")
    ext = ET.SubElement(inline, f"{{{WP}}}extent"); ext.set("cx", str(cx)); ext.set("cy", str(cy))
    ET.SubElement(inline, f"{{{WP}}}docPr", id="1", name=name)
    graphic = ET.SubElement(inline, f"{{{A}}}graphic")
    gd = ET.SubElement(graphic, f"{{{A}}}graphicData")
    gd.set("uri", PIC)
    pic = ET.SubElement(gd, f"{{{PIC}}}pic")
    nv = ET.SubElement(pic, f"{{{PIC}}}nvPicPr")
    ET.SubElement(nv, f"{{{PIC}}}cNvPr", id="0", name=name)
    ET.SubElement(nv, f"{{{PIC}}}cNvPicPr")
    bf = ET.SubElement(pic, f"{{{PIC}}}blipFill")
    ET.SubElement(bf, f"{{{A}}}blip").set(f"{{{R}}}embed", rid)
    ET.SubElement(ET.SubElement(bf, f"{{{A}}}stretch"), f"{{{A}}}fillRect")
    sp = ET.SubElement(pic, f"{{{PIC}}}spPr")
    xfrm = ET.SubElement(sp, f"{{{A}}}xfrm")
    off = ET.SubElement(xfrm, f"{{{A}}}off"); off.set("x", "0"); off.set("y", "0")
    e2 = ET.SubElement(xfrm, f"{{{A}}}ext"); e2.set("cx", str(cx)); e2.set("cy", str(cy))
    geom = ET.SubElement(sp, f"{{{A}}}prstGeom"); geom.set("prst", "rect")
    ET.SubElement(geom, f"{{{A}}}avLst")
    return p


def make_list_para(text, num_id, *, body_tpl=None, ilvl=0):
    """带项目符号/编号的段落（引用 numbering.xml 里的 numId）。"""
    p = deepcopy(body_tpl) if body_tpl is not None else ET.Element(q("p"))
    for r in list(p.findall(q("r"))): p.remove(r)
    pPr = p.find(q("pPr"))
    if pPr is None:
        pPr = ET.Element(q("pPr")); p.insert(0, pPr)
    for tag in (q("numPr"), q("ind")):
        for old in list(pPr.findall(tag)): pPr.remove(old)
    # CT_PPr 顺序：numPr → spacing → ind → jc → rPr；这里按序重建以规避顺序违规
    set_spacing(p, line=360, before_lines=0, after_lines=0)
    rpr = pPr.find(q("rPr"))
    numPr = ET.Element(q("numPr"))
    ET.SubElement(numPr, q("ilvl")).set(q("val"), str(ilvl))
    ET.SubElement(numPr, q("numId")).set(q("val"), str(num_id))
    sp = pPr.find(q("spacing"))
    if rpr is not None:
        rpr.addprevious(numPr)
    elif sp is not None:
        sp.addprevious(numPr)
    else:
        pPr.append(numPr)
    ind = ET.SubElement(pPr, q("ind")) if rpr is None else ET.Element(q("ind"))
    ind.set(q("left"), str(420 + 420 * ilvl)); ind.set(q("firstLine"), "0")
    if rpr is not None:
        rpr.addprevious(ind)
    set_run_text(p, text, eastasia="宋体", ascii_font="Times New Roman", sz_halfpt=SZ_BODY)
    return p


def make_footnote_run(fid):
    """正文里的脚注引用（阿拉伯数字上标）。"""
    r = ET.Element(q("r"))
    rpr = ET.SubElement(r, q("rPr"))
    ET.SubElement(rpr, q("vertAlign")).set(q("val"), "superscript")
    ET.SubElement(rpr, q("sz")).set(q("val"), str(SZ_BODY))
    ref = ET.SubElement(r, q("footnoteReference"))
    ref.set(q("id"), str(fid))
    return r


FOOTNOTES_SKELETON = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<w:footnotes xmlns:w="{W}">'
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" '
    'w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr>'
    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r>'
    '<w:continuationSeparator/></w:r></w:p></w:footnote>'
    '</w:footnotes>'
)


def make_footnote(fid, text):
    """footnotes.xml 里的一条脚注：五号宋体，注号上标。"""
    fn = ET.Element(q("footnote")); fn.set(q("id"), str(fid))
    p = ET.SubElement(fn, q("p"))
    pPr = ET.SubElement(p, q("pPr"))
    ET.SubElement(pPr, q("ind")).set(q("firstLine"), "0")
    r0 = ET.SubElement(p, q("r"))
    rpr0 = ET.SubElement(r0, q("rPr"))
    ET.SubElement(rpr0, q("vertAlign")).set(q("val"), "superscript")
    ET.SubElement(rpr0, q("sz")).set(q("val"), str(SZ_FIVE))
    ET.SubElement(r0, q("footnoteRef"))
    r = ET.SubElement(p, q("r"))
    rpr = ET.SubElement(r, q("rPr"))
    rf = ET.SubElement(rpr, q("rFonts"))
    rf.set(q("eastAsia"), "宋体"); rf.set(q("ascii"), "Times New Roman")
    rf.set(q("hAnsi"), "Times New Roman")
    ET.SubElement(rpr, q("sz")).set(q("val"), str(SZ_FIVE))
    ET.SubElement(rpr, q("szCs")).set(q("val"), str(SZ_FIVE))
    t = ET.SubElement(r, q("t")); t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = " " + text
    return fn


LIST_NUMBERING = (
    # 编号列表（1. 2. 3.）与项目符号列表，用未占用的 abstractNumId/numId
    '<w:abstractNum w:abstractNumId="90"><w:multiLevelType w:val="singleLevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
    '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
    '<w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr></w:lvl></w:abstractNum>'
    '<w:abstractNum w:abstractNumId="91"><w:multiLevelType w:val="singleLevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
    '<w:lvlText w:val="●"/><w:lvlJc w:val="left"/>'
    '<w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr>'
    '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr>'
    '</w:lvl></w:abstractNum>'
)
NUM_ORDERED, NUM_BULLET = 90, 91

def make_three_line_table(headers, rows, content_width=8306):
    """学术三线表：顶线、表头下线、底线，无竖线；居中；宋体+TNR 五号；表头加粗。"""
    ncols = len(headers)
    col_w = content_width // ncols
    tbl = ET.Element(q("tbl"))
    tblPr = ET.SubElement(tbl, q("tblPr"))
    tblW = ET.SubElement(tblPr, q("tblW")); tblW.set(q("w"), str(content_width)); tblW.set(q("type"), "dxa")
    ET.SubElement(tblPr, q("jc")).set(q("val"), "center")
    borders = ET.SubElement(tblPr, q("tblBorders"))
    borders.append(_border("top", "12")); borders.append(_border("bottom", "12"))
    for edge in ("left", "right", "insideV"):
        ET.SubElement(borders, q(edge)).set(q("val"), "none")
    grid = ET.SubElement(tbl, q("tblGrid"))
    for _ in range(ncols):
        ET.SubElement(grid, q("gridCol")).set(q("w"), str(col_w))
    def cell(text, bold=False, bottom=False):
        tc = ET.Element(q("tc"))
        tcPr = ET.SubElement(tc, q("tcPr"))
        ET.SubElement(tcPr, q("tcW")).set(q("w"), str(col_w)); tcPr.find(q("tcW")).set(q("type"), "dxa")
        if bottom:
            cb = ET.SubElement(tcPr, q("tcBorders")); cb.append(_border("bottom", "8"))
        ET.SubElement(tcPr, q("vAlign")).set(q("val"), "center")
        p = ET.SubElement(tc, q("p"))
        ET.SubElement(ET.SubElement(p, q("pPr")), q("jc")).set(q("val"), "center")
        r = ET.SubElement(p, q("r"))
        rpr = ET.SubElement(r, q("rPr"))
        rf = ET.SubElement(rpr, q("rFonts"))
        rf.set(q("eastAsia"), "宋体"); rf.set(q("ascii"), "Times New Roman"); rf.set(q("hAnsi"), "Times New Roman")
        if bold: ET.SubElement(rpr, q("b"))
        ET.SubElement(rpr, q("sz")).set(q("val"), str(SZ_FIVE))
        ET.SubElement(rpr, q("szCs")).set(q("val"), str(SZ_FIVE))
        t = ET.SubElement(r, q("t"))
        if text[:1].isspace() or text[-1:].isspace():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        return tc
    hr = ET.SubElement(tbl, q("tr"))
    for h in headers: hr.append(cell(h, bold=True, bottom=True))
    for row in rows:
        tr = ET.SubElement(tbl, q("tr"))
        cells = list(row) + [""] * (ncols - len(row))
        for c in cells[:ncols]: tr.append(cell(str(c)))
    return tbl

def scrub_docprops(entries: dict, author: str = "") -> list:
    """清掉模板带来的第三方个人信息（作者名、邮箱、最后修改人、打印时间）。

    内置模板是借用的往届/官方文件，docProps/core.xml 里留着原作者的姓名和邮箱。
    不清的话每份生成稿都会把陌生人的身份写进属性面板，交上去即泄露他人隐私，
    也会让答辩材料显示成别人做的。author 给了就写用户自己的姓名，否则留空。
    """
    import re
    changed = []
    core = entries.get("docProps/core.xml")
    if core is None:
        return changed
    s = core.decode("utf-8")
    for tag in ("dc:creator", "cp:lastModifiedBy"):
        # 只在原本确有内容时才记入 changed——模板已清空的情况下再报"已清理"是假信息
        had = re.search(rf"<{tag}>(.+?)</{tag}>", s, flags=re.S)
        s, n = re.subn(rf"<{tag}>.*?</{tag}>",
                       f"<{tag}>{author}</{tag}>" if author else f"<{tag}></{tag}>",
                       s, flags=re.S)
        if n and had: changed.append(tag)
    # 打印时间是原作者的行为痕迹，与本文档无关
    s, n = re.subn(r"<cp:lastPrinted>.*?</cp:lastPrinted>", "", s, flags=re.S)
    if n: changed.append("cp:lastPrinted")
    entries["docProps/core.xml"] = s.encode("utf-8")
    app = entries.get("docProps/app.xml")
    if app is not None:
        a = app.decode("utf-8")
        for tag in ("Company", "Manager"):
            had = re.search(rf"<{tag}>(.+?)</{tag}>", a, flags=re.S)
            a, n = re.subn(rf"<{tag}>.*?</{tag}>", f"<{tag}></{tag}>", a, flags=re.S)
            if n and had: changed.append(tag)
        entries["docProps/app.xml"] = a.encode("utf-8")
    # 批注作者表也带姓名与账号（常是没有任何批注的孤儿记录），一并清掉
    for key in ("word/commentsExtended.xml", "ppt/commentAuthors.xml", "word/people.xml"):
        blob = entries.get(key)
        if blob is None:
            continue
        b = blob.decode("utf-8", "replace")
        b2 = re.sub(r'(name|initials|userId|author)="[^"]*"', r'\1=""', b)
        if b2 != b:
            entries[key] = b2.encode("utf-8"); changed.append(key.split("/")[-1])
    return changed

def build(template: Path, content_path: Path, output: Path):
    data = json.loads(content_path.read_text(encoding="utf-8"))
    with ZipFile(template) as zin:
        entries = {it.filename: zin.read(it.filename) for it in zin.infolist()}
        infolist = zin.infolist()
    xml = entries["word/document.xml"]
    root = ET.fromstring(xml)
    body = root.find(q("body"))

    # ---- 1. 封面题目 ----
    title_p = find_para(body, "宏大百货网上购物系统")
    if title_p is not None:
        set_run_text(title_p, data["title"], eastasia=None, ascii_font=None)
        sub_p = next_para_sibling(body, title_p)
        if sub_p is not None and "的设计与实现" in text_of(sub_p):
            sub = data.get("subtitle", "")
            if sub:
                set_run_text(sub_p, sub, eastasia=None, ascii_font=None)
            else:
                body.remove(sub_p)  # 无副标题则删掉空行，避免封面留空白居中行

    # 清掉“论文题目论文题目”这类残留标签段
    for p in body.findall(q("p")):
        if "论文题目论文题目" in text_of(p):
            for r in list(p.findall(q("r"))): p.remove(r)

    # ---- 2. 封面表格替换填写位（专业名称/培养学院默认保留模板值，可在 cover 里覆盖） ----
    cover = data.get("cover", {})
    tbl = None
    for el in body:
        if el.tag == q("tbl"):
            tbl = el; break  # 第一个表格 = 封面表
    cover_field_labels = ["作者姓名", "专业名称", "指导教师", "培养学院"]
    if tbl is not None:
        for row in tbl.findall(q("tr")):
            cells = row.findall(q("tc"))
            if len(cells) < 2: continue
            label = text_of(cells[0]).strip()
            if label in cover_field_labels and label in cover:
                set_run_text(cells[1].findall(q("p"))[0], cover[label],
                              eastasia="宋体", ascii_font="Times New Roman")

    # ---- 3. 删除目录说明段（楷体那条） ----
    for p in list(body.findall(q("p"))):
        if text_of(p).startswith("目录按章、节、条"):
            body.remove(p); break

    # ---- 4. 找正文模板段（“开题报告”标题之后的第一个空段，避开目录） ----
    body_start = find_para_exact(body, "开题报告")  # 正文起始标记（idx44），精确匹配避免命中封面 banner
    # body 模板 = （一）研究背景 后的空段，但用正文区那次
    sec_bg = find_para_after(body, "（一）研究背景", body_start) if body_start is not None else None
    body_tpl = next_para_sibling(body, sec_bg) if sec_bg is not None else None
    # 章标题样板：用「四、」——它是单 run 承载整个标题（一~三章把序号和标题拆成两个 run），
    # 克隆时取它的 rPr 最干净；四章 run 的字号实测一致（都是 sz=24）。
    heading_tpl = find_para_after(body, "四、学位论文实施计划", body_start) \
        if body_start is not None else None

    # ---- 5. 删除模板中所有”导师建议”占位提示段（”可按照””可结合”两种变体都删） ----
    for p in list(body.findall(q("p"))):
        t = text_of(p)
        if "导师建议" in t and ("可按照" in t or "可结合" in t):
            body.remove(p)
    # 目录 TOC 字段末的红字说明（如「附录（若论文包含访谈或问卷…）」）也清掉
    trim_toc_field_end(body)

    # ---- 6. 在各节标题后插入正文内容（只在正文区查找，避开目录 TOC） ----
    # 混排块：字符串=正文段；{"table":…} / {"image":…} / {"list":…} / {"footnote":…}
    ctx = {
        "tbl_no": 0, "fig_no": 0,          # 图表跨节连续编号（格式规范：连续到附录前）
        "images": [], "footnotes": [],     # 待写入包的图片与脚注
        "next_rid": 100, "next_fid": 1,
        "sections_ok": {}, "sections_new": [],
        # 附录里的表/图另起一套「附表1/附图1」编号，不占正文的连续号
        "in_appendix": False, "ap_tbl_no": 0, "ap_fig_no": 0,
        "demo": bool(data.get("_demo")),
    }

    def bump_caption(kind):
        """返回 (题头文字, 编号)。附录内用「附表/附图」并走独立计数器。"""
        if ctx["in_appendix"]:
            key = "ap_tbl_no" if kind == "表" else "ap_fig_no"
            ctx[key] += 1
            return f"附{kind}", ctx[key]
        key = "tbl_no" if kind == "表" else "fig_no"
        ctx[key] += 1
        return kind, ctx[key]

    def insert_block(block, anchor):
        """插入一个内容块，返回新的 anchor。"""
        if isinstance(block, str):
            if body_tpl is None: return anchor
            p = clone_body_para(body_tpl, block)
            insert_after(body, anchor, p); return p

        if not isinstance(block, dict): return anchor

        if "table" in block:
            t = block["table"]
            headers, rows = t.get("headers", []), t.get("rows", [])
            if not headers: return anchor
            cap = t.get("caption")
            if cap:                                    # 表题在表【上方】
                kind, num = bump_caption("表")
                cp = make_caption(cap, kind, num, body_tpl=body_tpl)
                insert_after(body, anchor, cp); anchor = cp
            tb = make_three_line_table(headers, rows)
            insert_after(body, anchor, tb); return tb

        if "image" in block:
            src = Path(block["image"]).expanduser()
            if not src.is_file():
                hint = "（示例文件预期如此，换成自己的图片绝对路径即可）" if ctx["demo"] else ""
                print(f"⚠️ 图片不存在，已跳过: {src}{hint}"); return anchor
            raw = src.read_bytes()
            seq = ctx["next_rid"]; ctx["next_rid"] += 1
            rid = f"rIdImg{seq}"
            ext = src.suffix.lower().lstrip(".") or "png"
            ctx["images"].append((rid, f"image_{seq}.{ext}", raw, ext))
            cx, cy = image_extent(raw, block.get("max_width_cm", 14.0))
            ip = make_image_para(rid, cx, cy, name=src.stem, body_tpl=body_tpl)
            insert_after(body, anchor, ip); anchor = ip
            if block.get("caption"):                   # 图题在图【下方】
                kind, num = bump_caption("图")
                cp = make_caption(block["caption"], kind, num, body_tpl=body_tpl)
                insert_after(body, anchor, cp); anchor = cp
            return anchor

        if "list" in block:
            nid = NUM_ORDERED if block.get("ordered") else NUM_BULLET
            for item in block["list"]:
                lp = make_list_para(str(item), nid, body_tpl=body_tpl)
                insert_after(body, anchor, lp); anchor = lp
            return anchor

        if "footnote" in block:
            # {"text": 正文段, "footnote": 注文} → 段末加上标注号
            if body_tpl is None: return anchor
            p = clone_body_para(body_tpl, block.get("text", ""))
            fid = ctx["next_fid"]; ctx["next_fid"] += 1
            ctx["footnotes"].append((fid, block["footnote"]))
            p.append(make_footnote_run(fid))
            insert_after(body, anchor, p); return p

        return anchor

    for sec_title, blocks in data.get("content_by_section", {}).items():
        hp = find_para_after(body, sec_title, body_start) if body_start is not None else None
        if hp is None and sec_title in CREATABLE_SECTIONS:
            hp = ensure_section(body, sec_title, CREATABLE_SECTIONS[sec_title],
                                body_start, heading_tpl)
            if hp is not None:
                ctx["sections_new"].append(sec_title)
        if hp is None:
            print(f"⚠️ 未找到节标题: {sec_title}")
            ctx["sections_ok"][sec_title] = False
            continue
        anchor = hp
        for block in blocks:
            anchor = insert_block(block, anchor)
        # 空数组视为"本节留给 plan_table / 用户手填"，不算失败
        ctx["sections_ok"][sec_title] = (anchor is not hp) or not blocks

    # plan_table 兼容旧格式：若"四、"节没在混排里给过表，仍按老写法补一张
    if data.get("plan_table"):
        sec4 = next((s for s in data.get("content_by_section", {}) if s.startswith("四、")), None)
        already = any(isinstance(b, dict) and "table" in b
                      for b in data.get("content_by_section", {}).get(sec4 or "", []))
        if not already:
            hp = find_para_after(body, sec4 or "四、学位论文实施计划", body_start) if body_start is not None else None
            if hp is not None:
                pt = data["plan_table"]
                anchor = hp
                # 插到该节已有正文之后，避免正文被挤到表下方
                sib = hp.getnext()
                while sib is not None and sib.tag == q("p") and text_of(sib).strip() \
                        and not text_of(sib).strip().startswith(("参考文献", "附录")):
                    anchor = sib; sib = sib.getnext()
                # 格式规范要求表必有题；旧格式未给 caption 时用默认题名
                cap = pt.get("caption") or "学位论文研究进度安排"
                ctx["tbl_no"] += 1
                cp = make_caption(cap, "表", ctx["tbl_no"], body_tpl=body_tpl)
                insert_after(body, anchor, cp); anchor = cp
                insert_after(body, anchor,
                             make_three_line_table(pt.get("headers", []), pt.get("rows", [])))

    # ---- 7. 参考文献条目替换 ----
    refs = data.get("refs", [])
    ref_placeholders = [p for p in body.findall(q("p"))
                        if text_of(p).strip().startswith("［")]
    for i, p in enumerate(ref_placeholders):
        if i < len(refs):
            set_run_text(p, refs[i], eastasia="宋体", ascii_font="Times New Roman")
        else:
            body.remove(p)
    # 若 refs 多于占位条数，在最后一个占位后追加
    if len(refs) > len(ref_placeholders) and ref_placeholders:
        anchor = ref_placeholders[-1]
        for txt in refs[len(ref_placeholders):]:
            new_p = clone_body_para(ref_placeholders[-1], txt)
            insert_after(body, anchor, new_p); anchor = new_p

    # ---- 8. 附录 ----
    # 与 content_by_section 用同一套混排块：问卷量表题本来就该排成表格，
    # 只能一条一段的纯字符串会逼用户把「题干 □1 □2…」硬挤成一行。
    ap = find_para(body, "附录：（若论文包含访谈或问卷")
    if ap is not None and data.get("appendix"):
        set_run_text(ap, "附录", eastasia=None, ascii_font=None)
        ctx["in_appendix"] = True
        anchor = ap
        for blk in data["appendix"]:
            anchor = insert_block(blk, anchor)
        ctx["in_appendix"] = False

    # ---- 8.5 规范正文区章/节/条标题的行距 + 大纲级别 ----
    # 模板自带标题段行距是 1.5 倍（line=360），而书写规范 1.2 要求三级标题单倍、
    # 段前段后各 0.5 行。逐段把正文区（body_start 之后）的标题段强制成规范值覆盖模板旧值；
    # 新建正文段已由构造函数落实，这里补模板自带的标题段与换行后的标题。
    # 同时给章/节/条标题设 outlineLvl（章0 / 节1 / 条2）：模板的节/条/参考文献/附录标题
    # 都没挂标题样式，TOC 域 (TOC \o "1-4") 只按大纲级别/标题样式收录——不加的话，
    # 用户一刷新目录域(F9)，节/条/附录/参考文献就会从目录里消失，只剩章标题。
    if body_start is not None:
        from re import compile as _re
        _ZHANG = _re(r"^(一|二|三|四|五|六|七|八|九|十)、")
        _JIE = _re(r"^[（(](一|二|三|四|五|六|七|八|九|十)[)）]")
        _TIAO = _re(r"^\d+[、.]")
        paras_ = body.findall(q("p"))
        start_idx = list(body).index(body_start)
        # 只在主文档区处理标题（到 参考文献/附录 之前为止）——附录里的访谈提纲、
        # 问卷子标题、题号不进目录。目录粒度对齐模板缓存：章(一级) + 节(二级)。
        end_idx = len(paras_)
        for a in paras_[start_idx:]:
            if text_of(a).strip().startswith(("参考文献", "附录")):
                end_idx = list(body).index(a); break
        for p in paras_[start_idx:end_idx]:
            t = text_of(p).strip()
            if _ZHANG.match(t):
                set_outline(p, 0)
                set_spacing(p, line=240, before_lines=50, after_lines=50)
            elif _JIE.match(t):
                set_outline(p, 1)
                set_spacing(p, line=240, before_lines=50, after_lines=50)
            elif _TIAO.match(t):
                # 条标题：只规范行距（单倍+0.5行），不进目录（粒度到节级）
                set_spacing(p, line=240, before_lines=50, after_lines=50)

    # 参考文献 / 附录 标题进目录（一级）；它们标题无 pStyle，若不设 outlineLvl，
    # 模板缓存里手写的目录项一次 F9 刷新就没了。
    # 用【精确文本】匹配标题段——find_para_after 是子串匹配，正文里"…（见附录）…"
    # 会误命中，须按整段文本等于"附录"/"参考文献："来找。
    for needle, level in (("参考文献：", 0), ("附录", 0)):
        hp = None
        if body_start is not None:
            sib = body_start.getnext()
            while sib is not None and hp is None:
                if sib.tag == q("p") and text_of(sib).strip() == needle:
                    hp = sib
                sib = sib.getnext()
        if hp is None:
            for _p in body.findall(q("p")):
                if text_of(_p).strip() == needle:
                    hp = _p; break
        if hp is not None:
            set_outline(hp, level)

    # ---- 9. 删除“书写规范”段及之后全部（含 body-level sectPr），让 idx90 的 sectPr 收尾 ----
    spec_p = find_para(body, "书写规范")
    if spec_p is not None:
        # 收集 spec_p 及其所有后续兄弟，逐个从 body 移除
        sib = spec_p
        to_remove = []
        while sib is not None:
            to_remove.append(sib); sib = sib.getnext()
        for el in to_remove:
            body.remove(el)

    # ---- 9.5 清理全文档多余空行，消除“一大堆空白页/很多空行” ----
    # 官方模板的封面区、目录区自带大量空白段（有的行距还特别大，如 line=460 exact），
    # 会在封面页底部、目录页末尾堆出大片空白，看起来像“夹了一页空白/很多空行”。
    # 这里对整个 body 统一处理（不再限于正文区）：
    #   1) 把连续≥2 个空段折叠成 1 个（保留一处作为节与节之间的自然隔行/封面推位）；
    #   2) 删掉文档最末尾的连续空段（避免文末拖出空白页）。
    # 分节符段（sectPr）已由 is_empty_p 排除，不会被折叠或误删——封面→目录→正文的
    # nextPage 换页由 SECT 自身保证，不依赖空段“撑着”。
    body_children = list(body)
    # (a) 从 body 末尾往前找最后一个非空段，作为末尾裁剪界；SECT 视为非空，天然止住
    end = len(body_children)
    while end - 1 > 0 and is_empty_p(body_children[end - 1]):
        end -= 1
    # (b) 中间段折叠：连续≥2 空段折成 1 个
    i = 0
    while i < end:
        if is_empty_p(body_children[i]):
            span = 1
            while i + span < end and is_empty_p(body_children[i + span]):
                span += 1
            if span > 1:
                for k in range(i + 1, i + span):
                    body.remove(body_children[k])
                del body_children[i + 1: i + span]
                end -= (span - 1)
            i += 1
        else:
            i += 1
    # (c) 末尾裁剪：删掉 end 之后到 body 末尾的空段
    for k in range(end, len(body_children)):
        if is_empty_p(body_children[k]):
            body.remove(body_children[k])
    del body_children[end:]

    # ---- 写出（lxml tostring 保留命名空间前缀与 standalone 声明）----
    out_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True, standalone=True)

    # 追加图片 / 脚注 / 列表编号所需的包内部件与关系
    extra = {}
    rels_add, ct_add = [], []
    REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    if ctx["images"]:
        for rid, fname, raw, ext in ctx["images"]:
            extra[f"word/media/{fname}"] = raw
            rels_add.append(f'<Relationship Id="{rid}" Type="{REL_NS}/image" '
                            f'Target="media/{fname}"/>')
            if ext in ("jpg", "jpeg"):
                ct_add.append('<Default Extension="jpeg" ContentType="image/jpeg"/>')
            elif ext == "gif":
                ct_add.append('<Default Extension="gif" ContentType="image/gif"/>')

    if ctx["footnotes"]:
        fns = ET.fromstring(FOOTNOTES_SKELETON.encode("utf-8"))
        for fid, text in ctx["footnotes"]:
            fns.append(make_footnote(fid, text))
        extra["word/footnotes.xml"] = ET.tostring(fns, encoding="utf-8",
                                                  xml_declaration=True, standalone=True)
        rels_add.append(f'<Relationship Id="rIdFootnotes" Type="{REL_NS}/footnotes" '
                        f'Target="footnotes.xml"/>')
        ct_add.append('<Override PartName="/word/footnotes.xml" ContentType="application/'
                      'vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>')

    if rels_add:
        rels = entries["word/_rels/document.xml.rels"].decode("utf-8")
        entries["word/_rels/document.xml.rels"] = rels.replace(
            "</Relationships>", "".join(rels_add) + "</Relationships>").encode("utf-8")

    if ct_add:
        ct = entries["[Content_Types].xml"].decode("utf-8")
        for frag in dict.fromkeys(ct_add):          # 去重，且不重复声明已有的
            key = frag.split('Extension="')[-1].split('"')[0] if frag.startswith("<Default") \
                else frag.split('PartName="')[-1].split('"')[0]
            if key not in ct:
                ct = ct.replace("</Types>", frag + "</Types>")
        entries["[Content_Types].xml"] = ct.encode("utf-8")

    # 列表编号：把自定义 abstractNum/num 追加进 numbering.xml
    if any(isinstance(b, dict) and "list" in b
           for blocks in data.get("content_by_section", {}).values() for b in blocks):
        numx = entries["word/numbering.xml"].decode("utf-8")
        if 'w:abstractNumId="90"' not in numx:
            nums = (f'<w:num w:numId="{NUM_ORDERED}"><w:abstractNumId w:val="90"/></w:num>'
                    f'<w:num w:numId="{NUM_BULLET}"><w:abstractNumId w:val="91"/></w:num>')
            # abstractNum 必须排在所有 num 之前
            idx = numx.find("<w:num ")
            if idx == -1:
                numx = numx.replace("</w:numbering>", LIST_NUMBERING + nums + "</w:numbering>")
            else:
                numx = numx[:idx] + LIST_NUMBERING + numx[idx:]
                numx = numx.replace("</w:numbering>", nums + "</w:numbering>")
            entries["word/numbering.xml"] = numx.encode("utf-8")

    # 目录 TOC 域缓存的还是模板范例页码。置 updateFields 让 Word 打开时更新目录，
    # 省掉用户手动 F9（不改域本身，只加这一条设置）。
    # CT_Settings 是有序 sequence，updateFields 必须紧排在 hdrShapeDefaults 之前，
    # 位置放错 Word 会报「需要修复」。
    setx = entries["word/settings.xml"].decode("utf-8")
    if "w:updateFields" not in setx:
        for anchor in ("<w:hdrShapeDefaults", "<w:footnotePr", "<w:compat"):
            if anchor in setx:
                setx = setx.replace(anchor, '<w:updateFields w:val="true"/>' + anchor, 1)
                entries["word/settings.xml"] = setx.encode("utf-8")
                break

    # 封面姓名若还是〈XXX〉占位，就别写进属性里，留空比写个占位干净
    _au = str((data.get("cover") or {}).get("作者姓名", "") or "")
    if any(m in _au for m in ("〈", "〉", "XXX", "待补", "待定")):
        _au = ""
    scrubbed = scrub_docprops(entries, _au)

    with ZipFile(output, "w", ZIP_DEFLATED) as zout:
        for item in infolist:
            payload = entries[item.filename]
            if item.filename == "word/document.xml":
                payload = out_xml
            zout.writestr(item, payload)
        for name, payload in extra.items():
            zout.writestr(name, payload)
    print("saved:", output)
    validate(output, data, ctx, scrubbed=scrubbed)

def flatten_block_text(blocks, *, include_captions=True, include_tables=True):
    """把混排块摊平成文本串，供字数统计与占位扫描共用。

    include_tables 只在占位扫描时开——图表单元格不是"写"出来的篇幅，
    不该计入字数，但里面留着〈待补〉必须能被查出来。
    """
    for b in blocks:
        if isinstance(b, str):
            yield b
        elif isinstance(b, dict):
            if b.get("text"): yield b["text"]
            if include_captions and b.get("caption"): yield b["caption"]
            if isinstance(b.get("list"), list):
                for x in b["list"]: yield str(x)
            if b.get("footnote"): yield str(b["footnote"])
            t = b.get("table")
            if include_tables and isinstance(t, dict):
                if t.get("caption"): yield t["caption"]
                for h in t.get("headers", []) or []: yield str(h)
                for row in t.get("rows", []) or []:
                    for cell in row or []: yield str(cell)


def section_char_counts(data):
    """按 content_by_section 统计每节中文字数，供对照篇幅档位查哪节欠字。

    只数正文与列表项、表格单元格与图表题不计入——图表题不是"写"出来的篇幅。
    """
    out = {}
    for sec, blocks in data.get("content_by_section", {}).items():
        buf = flatten_block_text(blocks, include_captions=False, include_tables=False)
        out[sec] = sum(1 for c in "".join(buf) if '一' <= c <= '鿿')
    return out


def check_refs_format(refs):
    """按 GB/T 7714 著录规则（references/参考文献著录.md）对 refs 逐条做自动化核查。

    只做机器能判定的硬性违规，返回 [(序号, 问题)…]；逐条说明性规范（作者名写法、
    页码起止等）仍靠人工按 自查清单，这里不误报。占位条目（〈〉/待核验/TODO）单列。
    """
    import re
    # GB/T 7714 文献类型标识：[J][M][A/C][D][P][S][N][R][G][Z]
    DOC_TYPES = "AJMNDPSRGZ"
    # 电子文献 [类型/载体]：类型∈{DB,CP,M,EB}，载体∈{MT,DK,CD,OL}
    ELEC_MAIN = "(?:DB|CP|M|EB)"
    ELEC_CARRIER = "(?:MT|DK|CD|OL)"
    # 合法标识头：[J]、[Z]、[EB/OL]、[M/CD]、[DB/MT]…
    TYPE_RE = rf"\[(?:[{DOC_TYPES}]|{ELEC_MAIN}/{ELEC_CARRIER})\]"
    problems, placeholders = [], []
    for i, raw in enumerate(refs, 1):
        s = str(raw)
        # 占位：待核验/待补/〈XXX〉/TODO——作为"未完成"单独报，不掺进格式问题
        if any(m in s for m in ("〈", "〉", "待核验", "待补", "待验证", "TODO", "XXX")):
            placeholders.append(i)
        # ① 缺文献类型标识 [J]/[M]/[Z]/[EB/OL]…（含电子文献）
        if not re.search(TYPE_RE, s):
            problems.append((i, "缺文献类型标识 [J]/[M]/[Z]/[R]/[S]/[EB/OL] 等"))
        # ② 条目末尾带了结束符（句点/逗号/分号/顿号）
        if re.search(r"[.,;:、，。；：]+$", s.strip()):
            problems.append((i, "条目末尾有结束符（GB/T 7714 每条末不加标点）"))
    return problems, sorted(set(placeholders))


def check_line_spacing(root):
    """按书写规范 1.2 自动核验正文区段落的行距/段前段后，返回不合规清单。

    规范（references/格式规范.md「页面」）：
      章/节/条三级标题 → 单倍行距（w:spacing line=240）、段前段后各 0.5 行（beforeLines/afterLines=50）
      正文           → 1.5 倍行距（line=360）、段前段后 0 行
    只扫正文区（"开题报告"标记之后、到"参考文献/附录"之前），逐段分类：命中章/节/条标题
    模式的要求单倍+0.5 行；其余带正文文本的段要求 1.5 倍行距。附录里的访谈提纲/问卷子标题、
    题号不是文档标题，不入检。表题/图题（表n/图n 开头）与参考文献（[n]）单列不误报。
    返回 [(起头文本, 问题描述), …]；空列表即全部合规。
    """
    import re as _re
    _ZHANG = _re.compile(r"^(一|二|三|四|五|六|七|八|九|十)、")
    _JIE = _re.compile(r"^[（(](一|二|三|四|五|六|七|八|九|十)[)）]")
    _TIAO = _re.compile(r"^\d+[、.]\s*\S")
    _CAP = _re.compile(r"^(表|图)\d+")
    _REF = _re.compile(r"^[\[［]\d+")
    body = root.find(q("body"))
    if body is None:
        return []
    paras = body.findall(q("p"))
    body_start = next((i for i, p in enumerate(paras)
                       if "".join(x.text or "" for x in p.iter(q("t"))).strip() == "开题报告"), -1)
    if body_start < 0:
        return []
    # 主文档区：到 参考文献/附录 之前为止（附录内的访谈提纲/题号非文档标题，不核）
    end_idx = len(paras)
    for i in range(body_start + 1, len(paras)):
        t0 = "".join(x.text or "" for x in paras[i].iter(q("t"))).strip()
        if t0.startswith(("参考文献", "附录")):
            end_idx = i
            break
    issues = []
    for p in paras[body_start + 1:end_idx]:
        t = "".join(x.text or "" for x in p.iter(q("t"))).strip()
        if not t or _REF.match(t) or _CAP.match(t):
            continue
        ppr = p.find(q("pPr"))
        sp = ppr.find(q("spacing")) if ppr is not None else None
        line = int(sp.get(q("line"), 0)) if sp is not None else 0
        bl = int(sp.get(q("beforeLines"), 0)) if sp is not None else 0
        al = int(sp.get(q("afterLines"), 0)) if sp is not None else 0
        head = bool(_ZHANG.match(t) or _JIE.match(t) or _TIAO.match(t))
        if head:
            if line != 240 or bl != 50 or al != 50:
                issues.append((t[:20], f"标题应单倍行距、段前段后各0.5行（实测 line={line or 240} beforeLines={bl} afterLines={al}）"))
        elif len(t) >= 10:
            if line != 360:
                issues.append((t[:20], f"正文应1.5倍行距（实测 line={line or 240}）"))
    return issues


def validate(path, data, ctx=None, scrubbed=None):
    from xml.etree import ElementTree as ET
    with ZipFile(path) as z:
        names = z.namelist()
        root = ET.fromstring(z.read("word/document.xml"))
        text = "".join(t.text or "" for t in root.iter(q("t")))
        media = [n for n in names if n.startswith("word/media/")]
    paras = root.findall(f"{q('body')}/{q('p')}")
    tables = root.findall(f"{q('body')}/{q('tbl')}")
    chinese = sum(1 for c in text if '一' <= c <= '鿿')
    leftover = any(x in text for x in ["宏大百货", "书写规范", "可按照导师建议",
                                       "附录（若论文包含访谈或问卷", "目录按章、节、条"])
    # 未填占位：〈〉是本 skill 大纲/范例的占位约定，连同 TODO/待补/待定/XXX 一起查。
    # 与 leftover 分开报——leftover 是"模板自带说明没删净"，这里是"内容自己没写完"，
    # 混在一条里会让用户以为改的是同一个地方。
    # 只查 content_by_section/refs/appendix 的实际内容，不查整篇 text：
    # 封面占位单独由下面 cover_ph 报，否则同一个〈XXX〉会被报两遍、看不出在哪。
    body_strings = list(flatten_block_text(
        [b for blocks in (data.get("content_by_section") or {}).values() for b in blocks]
        + list(data.get("appendix") or [])))
    body_blob = "".join(body_strings)
    unfilled = [m for m in ("〈", "〉", "TODO", "待补", "待定", "XXX") if m in body_blob]
    unfilled_refs = sum(1 for r in (data.get("refs") or [])
                        if any(m in str(r) for m in ("〈", "〉", "待补", "待定")))
    # templates/内容示例-*.json 自带占位封面与占位参考文献——那是故意的：
    # 换成像真的假名假文献，反而会诱导用户把编造的文献直接交上去。
    # 所以示例文件用 "_demo": true 自报身份，警告照打但标明"示例预期"，
    # 免得每次跑示例都像出了三个问题。用户自己的 content.json 不带这个键。
    demo = bool(data.get("_demo"))
    tag = "（示例文件预期如此，换成自己的内容时须填实）" if demo else ""
    print("=" * 50)
    print("自检：")
    print(f"  段落数: {len(paras)}")
    print(f"  中文字符数: {chinese}")
    per_sec = section_char_counts(data)
    if per_sec:
        print("  各节中文字数（对照 references/篇幅档位.md 看哪节欠字）:")
        for name, n in per_sec.items():
            print(f"    {name}: {n}")
    print(f"  题目已换: {'宏大百货' not in text}")
    cover = data.get("cover", {})
    cover_ok = [k for k, v in cover.items() if v and v in text]
    # 示例 json 的 cover 给的是〈XXX〉这类占位，直接报"已填"会让用户以为封面没问题。
    cover_ph = [k for k, v in cover.items()
                if v and any(m in str(v) for m in ("〈", "〉", "XXX", "待补", "待定"))]
    print(f"  封面填写: {cover_ok if cover_ok else '未给 cover'}"
          f"{'' if len(cover_ok) == len(cover) else f' ⚠️ 缺 {sorted(set(cover) - set(cover_ok))}'}"
          f"{f' ⚠️ {cover_ph} 仍是占位{tag}' if cover_ph else ''}")
    if ctx is not None:
        ok = ctx["sections_ok"]
        failed = [s for s, v in ok.items() if not v]
        print(f"  内容节插入: {len(ok) - len(failed)}/{len(ok)} 成功"
              f"{'' if not failed else f' ⚠️ 失败: {failed}'}")
        if ctx["sections_new"]:
            print(f"  新建节: {ctx['sections_new']}（模板原无，已按「可按导师建议调整」新增）")
        ap_t, ap_f = ctx.get("ap_tbl_no", 0), ctx.get("ap_fig_no", 0)
        ap_t_note = f"；附录另有 附表1–附表{ap_t}" if ap_t else ""
        ap_f_note = f"；附录另有 附图1–附图{ap_f}" if ap_f else ""
        print(f"  表格: {len(tables) - 1} 个（正文，不含封面表）；"
              f"表题编号至 表{ctx['tbl_no']}{ap_t_note}")
        print(f"  插图: {ctx['fig_no']} 个；图题编号至 图{ctx['fig_no']}{ap_f_note}")
        print(f"  脚注: {len(ctx['footnotes'])} 条"
              f"{'（word/footnotes.xml 已写入）' if ctx['footnotes'] else ''}")
    print(f"  参考文献: {bool(data.get('refs')) and str(data['refs'][0])[:20] in text}"
          f"{f' ⚠️ 其中 {unfilled_refs} 条仍是占位（未换成真实文献）{tag}' if unfilled_refs else ''}")
    # GB/T 7714 逐条自动核查（格式硬性违规 + 占位单列）
    ref_problems, _ph = check_refs_format(data.get("refs") or [])
    if ref_problems:
        detail = "；".join(f"[{n}]{msg}" for n, msg in ref_problems[:8])
        more = f"；另有 {len(ref_problems) - 8} 条类似" if len(ref_problems) > 8 else ""
        print(f"  ⚠️ 参考文献著录格式待改（对照 references/参考文献著录.md）: {detail}{more}")
    else:
        print("  参考文献著录格式: 通过（自动核查未见硬性违规，仍建议人工过一遍姓名/页码写法）")
    # 行距自动核验（书写规范 1.2：标题单倍+段前段后0.5行，正文1.5倍）
    _sp_issues = check_line_spacing(root)
    if _sp_issues:
        _d = "；".join(f"「{t}」{m}" for t, m in _sp_issues[:6])
        _m = f"；另有 {len(_sp_issues) - 6} 段类似" if len(_sp_issues) > 6 else ""
        print(f"  ⚠️ 行距不合规范（对照 references/格式规范.md「页面」）: {_d}{_m}")
    else:
        print("  行距: 通过（标题单倍+段前段后各0.5行，正文1.5倍，表图题单倍）")
    ap_items = data.get("appendix") or []
    ap_first = next(iter(flatten_block_text(ap_items[:1])), "") if ap_items else ""
    print(f"  附录: {bool(ap_items) and bool(ap_first) and ap_first[:15] in text}"
          f"{f'（{len(ap_items)} 块）' if ap_items else ''}")
    print(f"  无模板残留: {not leftover} {'' if not leftover else leftover}")
    if unfilled:
        print(f"  ⚠️ 检出未填占位标记: {unfilled}（正文里还有没写完的地方，定稿前须清掉）{tag}")
    print(f"  书写规范已删: {'书写规范' not in text}")
    print(f"  媒体资源: {media if media else '无'}")
    if scrubbed:
        print(f"  文档属性已清理: {scrubbed}（模板原作者姓名/邮箱/单位已清空，避免把他人信息带进你的稿件）")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True, type=Path)
    ap.add_argument("--content", required=True, type=Path)
    ap.add_argument("--output", required=True, type=Path)
    a = ap.parse_args()
    build(a.template, a.content, a.output)

if __name__ == "__main__":
    main()
