#!/usr/bin/env node
/* pptxgenjs 渲染层：把 ppt_content.json 生成汇报 PPT。
 *
 * 集成 pptx skill 的正确姿势（其 SKILL.md 规定：创建 deck 走 pptxgenjs，
 * 坑已列明——LAYOUT_WIDE 13.3×7.5、色值不带 #、阴影 offset≥0、选项对象
 * 不复用、三线表用 per-cell border、图标 react-icons→sharp 光栅化、
 * 出稿后跑 office/validate.py）。
 *
 * 用法：
 *   node gen_deck.js --content ppt_content.json --output 开题汇报.pptx
 *
 * 设计系统「批注过的图纸」：纸白 + 北航蓝双阶 + 发丝线，赭金只用于批注。
 * 版式随内容变化：cards(≥3 网格+核心条 / 2 磁贴)、三线表、flow、gantt、
 * stats(尺寸标注线)、panels；章节分隔页带 react-icons 图标（金色，克制）。
 */
"use strict";
const fs = require("fs");
const pptxgen = require("pptxgenjs");

// ── tokens ────────────────────────────────────────────────────
const NAVY = "003366", BLUE = "005BAC", PAPER = "FBFCFE", HAIR = "C9D6E6",
      CARDLINE = "D9E2EE", GOLD = "B98A2F", INK = "2B2B2B", MUTE = "6B7A8C",
      WHITE = "FFFFFF", GHOST = "E3EDF7", SKY = "BDD3EA", MIST = "9EBEDE";
const FONT = "Microsoft YaHei";
const W = 13.333, H = 7.5, MX = 0.95;
const IN = (n) => n; // 尺寸单位英寸

const SECTION_ICONS = {
  "研究背景": "FaCompass", "选题意义": "FaGem", "文献综述": "FaBookOpen",
  "研究框架": "FaDraftingCompass", "研究思路": "FaRoute", "研究方法": "FaTools",
  "创新之处": "FaLightbulb", "论文大纲": "FaLayerGroup",
  "实施计划": "FaCalendarAlt", "预期成果": "FaFlagCheckered",
};

async function iconData(name, colorHex) {
  try {
    const React = require("react");
    const { renderToStaticMarkup } = require("react-dom/server");
    const sharp = require("sharp");
    const fa = require("react-icons/fa");
    const Icon = fa[name];
    if (!Icon) return null;
    let svg = renderToStaticMarkup(React.createElement(Icon, { size: 256 }));
    svg = svg.replace(/currentColor/g, "#" + colorHex);
    const buf = await sharp(Buffer.from(svg)).png().toBuffer();
    return "image/png;base64," + buf.toString("base64");
  } catch (e) { return null; }
}

async function imageSize(p) {
  try {
    const sharp = require("sharp");
    const m = await sharp(p).metadata();
    return { w: m.width, h: m.height };
  } catch (e) { return null; }
}

function parseArgs() {
  const a = { _: [] };
  process.argv.slice(2).forEach((s, i, arr) => {
    if (s.startsWith("--")) a[s.slice(2)] = arr[i + 1];
  });
  return a;
}

// ── 通用元件 ──────────────────────────────────────────────────
function band(pres, slide) {
  slide.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.14, fill: { color: NAVY } });
  slide.addShape(pres.ShapeType.rect, { x: 0, y: 0.14, w: W, h: 0.03, fill: { color: GOLD } });
}

function footer(pres, slide, text, no, total) {
  const y = H - 0.42;
  slide.addShape(pres.ShapeType.rect, { x: MX, y: y - 0.10, w: W - 2 * MX, h: 0.008, fill: { color: "E2E8EF" } });
  slide.addText(text, { x: MX, y: y - 0.06, w: 8, h: 0.3, fontSize: 9.5, color: MUTE, fontFace: FONT, align: "left", margin: 0 });
  slide.addText(`${no} / ${total}`, { x: W - MX - 1.2, y: y - 0.06, w: 1.2, h: 0.3, fontSize: 10, color: NAVY, bold: true, fontFace: FONT, align: "right", margin: 0 });
}

function header(pres, slide, tag, title) {
  band(pres, slide);
  if (tag) slide.addText(tag.toUpperCase(), { x: MX, y: 0.36, w: 6, h: 0.26, fontSize: 10.5, color: BLUE, bold: true, fontFace: FONT, align: "left", margin: 0, charSpacing: 2 });
  slide.addText(title, { x: MX, y: 0.64, w: W - 2 * MX, h: 0.62, fontSize: 23, color: INK, bold: true, fontFace: FONT, align: "left", margin: 0, valign: "middle" });
  slide.addShape(pres.ShapeType.rect, { x: MX, y: 1.30, w: 1.15, h: 0.035, fill: { color: GOLD } });
}

function paperShadow(pres) { // 每次新建，不复用（pptxgenjs 会原地改写选项对象）
  return { type: "outer", angle: 90, offset: 2, blur: 5, opacity: 0.12 };
}

function splitLead(text) {
  const i = text.indexOf("：");
  if (i > 0 && i <= 10) return [{ text: text.slice(0, i + 1), options: { bold: true, color: NAVY } }, { text: text.slice(i + 1), options: {} }];
  return [{ text, options: {} }];
}

// ── 各页型 ────────────────────────────────────────────────────
function slideCover(pres, deck, total) {
  const s = pres.addSlide();
  const cover = deck.cover || {};
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.rect, { x: 0.22, y: 0.22, w: W - 0.44, h: H - 0.44, fill: { type: "none" }, line: { color: "3D5F8F", width: 0.75 } });
  s.addText(`${deck.school || "北京航空航天大学"} · ${cover.培养学院 || "公共管理学院"}`, { x: 0.85, y: 0.62, w: 8, h: 0.3, fontSize: 12, color: SKY, fontFace: FONT, align: "left", margin: 0 });
  s.addText("开题汇报 · 硕士学位论文开题", { x: W - 4.85, y: 0.62, w: 4, h: 0.3, fontSize: 12, color: SKY, fontFace: FONT, align: "right", margin: 0 });
  s.addText(deck.title || "", { x: 0.85, y: 2.15, w: 11.4, h: 1.55, fontSize: 32, color: WHITE, bold: true, fontFace: FONT, align: "left", margin: 0, valign: "top", lineSpacingMultiple: 1.25 });
  if (deck.subtitle) s.addText(deck.subtitle, { x: 0.85, y: 3.62, w: 9, h: 0.45, fontSize: 16, color: SKY, fontFace: FONT, align: "left", margin: 0 });
  s.addShape(pres.ShapeType.rect, { x: 0.85, y: 4.25, w: 2.2, h: 0.045, fill: { color: GOLD } });
  const info = [["汇报人", cover.作者姓名], ["指导教师", cover.指导教师], ["专业", cover.专业名称], ["日期", cover.日期]];
  const iw = (W - 1.7) / 4;
  info.forEach(([k, v], i) => {
    const x = 0.85 + i * iw;
    s.addText(`${k}：${v || ""}`, { x, y: H - 1.45, w: iw - 0.3, h: 0.35, fontSize: 12.5, color: "E6EEF6", fontFace: FONT, align: "left", margin: 0 });
    if (i > 0) s.addShape(pres.ShapeType.rect, { x: x - 0.18, y: H - 1.42, w: 0.008, h: 0.28, fill: { color: "2A4E7F" } });
  });
  return s;
}

function slideToc(pres, deck, chapters) {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.rect, { x: 0.22, y: 0.22, w: W - 0.44, h: H - 0.44, fill: { type: "none" }, line: { color: "3D5F8F", width: 0.75 } });
  s.addText("目录", { x: 1.1, y: 0.95, w: 3, h: 0.6, fontSize: 30, color: WHITE, bold: true, fontFace: FONT, align: "left", margin: 0 });
  s.addText("CONTENTS", { x: 1.12, y: 1.58, w: 3, h: 0.3, fontSize: 12, color: MIST, fontFace: "Calibri", align: "left", margin: 0, charSpacing: 3 });
  s.addShape(pres.ShapeType.rect, { x: 1.12, y: 1.98, w: 1.4, h: 0.04, fill: { color: GOLD } });
  const n = chapters.length;
  const cols = n > 5 ? 2 : 1;
  const rows = Math.ceil(n / cols);
  const y0 = 2.45, rowH = Math.min((H - 3.0) / rows, 0.62);
  chapters.forEach((name, i) => {
    const c = Math.floor(i / rows), r = i % rows;
    const x = 1.1 + c * 6.0, y = y0 + r * rowH;
    s.addText(String(i + 1).padStart(2, "0"), { x, y, w: 0.55, h: 0.4, fontSize: 16, color: MIST, bold: true, fontFace: "Calibri", align: "left", margin: 0, valign: "middle" });
    s.addText(name, { x: x + 0.62, y, w: 5.0, h: 0.4, fontSize: 15, color: WHITE, bold: true, fontFace: FONT, align: "left", margin: 0, valign: "middle" });
  });
  return s;
}

function slideSection(pres, deck, idx, name, ic, footerText, no, total) {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.rect, { x: 0.22, y: 0.22, w: W - 0.44, h: H - 0.44, fill: { type: "none" }, line: { color: "3D5F8F", width: 0.75 } });
  if (ic) s.addImage({ data: ic, x: W - 2.3, y: H - 2.45, w: 0.85, h: 0.85, transparency: 25 });
  s.addText(String(idx).padStart(2, "0"), { x: 1.1, y: 1.35, w: 3.5, h: 1.7, fontSize: 96, color: "1D4470", bold: true, fontFace: "Calibri", align: "left", margin: 0, valign: "top" });
  s.addText(name, { x: 1.15, y: 3.15, w: 9, h: 0.75, fontSize: 31, color: WHITE, bold: true, fontFace: FONT, align: "left", margin: 0 });
  s.addShape(pres.ShapeType.rect, { x: 1.15, y: 4.0, w: 1.8, h: 0.045, fill: { color: GOLD } });
  footer(pres, s, footerText, no, total);
  return s;
}

function contentArea() { return { x: MX, y: 1.55, w: W - 2 * MX, h: H - 1.55 - 0.55 }; }

function slideCards(pres, slide, spec) {
  const items = ((spec.extra || {}).items) || spec.bullets || [];
  if (!items.length) return;
  const a = contentArea();
  const n = items.length;
  const gap = 0.16;
  if (n === 2) { // 大磁贴
    const tw = (a.w - gap) / 2, th = Math.min(a.h, 2.9);
    const y0 = a.y + (a.h - th) / 2;
    items.forEach((b, i) => {
      const x = a.x + i * (tw + gap);
      slide.addShape(pres.ShapeType.roundRect, { x, y: y0, w: tw, h: th, rectRadius: 0.08, fill: { color: PAPER }, line: { color: CARDLINE, width: 0.75 }, shadow: paperShadow(pres) });
      slide.addText(String(i + 1).padStart(2, "0"), { x: x + 0.3, y: y0 + 0.25, w: 2.2, h: 1.0, fontSize: 44, color: GHOST, bold: true, fontFace: "Calibri", align: "left", margin: 0, valign: "top" });
      slide.addText(splitLead(b), { x: x + 0.3, y: y0 + 1.3, w: tw - 0.6, h: th - 1.55, fontSize: 13, color: INK, fontFace: FONT, align: "left", margin: 0, valign: "top", lineSpacingMultiple: 1.35 });
      slide.addShape(pres.ShapeType.rect, { x: x + 0.3, y: y0 + th - 0.2, w: 0.85, h: 0.04, fill: { color: BLUE } });
    });
    return;
  }
  const hasTakeaway = n >= 3;
  const grid = hasTakeaway ? items.slice(0, -1) : items;
  const barH = hasTakeaway ? 0.62 : 0;
  const gridH = a.h - barH - gap;
  const cols = grid.length > 4 ? 2 : 1;
  const rows = Math.ceil(grid.length / cols);
  const cw = (a.w - gap * (cols - 1)) / cols;
  const ch = Math.min((gridH - gap * (rows - 1)) / rows, 1.25);
  grid.forEach((b, i) => {
    const r = Math.floor(i / cols), c = i % cols;
    const x = a.x + c * (cw + gap), y = a.y + (gridH - (rows * ch + (rows - 1) * gap)) / 2 + r * (ch + gap);
    slide.addShape(pres.ShapeType.roundRect, { x, y, w: cw, h: ch, rectRadius: 0.07, fill: { color: PAPER }, line: { color: CARDLINE, width: 0.75 }, shadow: paperShadow(pres) });
    slide.addShape(pres.ShapeType.roundRect, { x: x, y: y + 0.1, w: 0.045, h: ch - 0.2, rectRadius: 0.02, fill: { color: BLUE } });
    slide.addText(String(i + 1), { shape: pres.ShapeType.ellipse, x: x + 0.17, y: y + (ch - 0.34) / 2, w: 0.34, h: 0.34, fill: { color: WHITE }, line: { color: BLUE, width: 1.1 }, fontSize: 11.5, color: NAVY, bold: true, fontFace: FONT, align: "center", margin: 0, valign: "middle" });
    slide.addText(splitLead(b), { x: x + 0.66, y, w: cw - 0.8, h: ch, fontSize: cols === 1 ? 12.5 : 11, color: INK, fontFace: FONT, align: "left", margin: 0, valign: "middle", lineSpacingMultiple: 1.25 });
  });
  if (hasTakeaway) {
    const by = a.y + a.h - barH;
    slide.addShape(pres.ShapeType.rect, { x: a.x, y: by, w: a.w, h: barH, fill: { color: NAVY }, shadow: { type: "outer", angle: 90, offset: 2, blur: 5, opacity: 0.18 } });
    slide.addShape(pres.ShapeType.rect, { x: a.x + 0.05, y: by + 0.05, w: a.w - 0.1, h: barH - 0.1, fill: { type: "none" }, line: { color: GOLD, width: 0.9 } });
    slide.addShape(pres.ShapeType.roundRect, { x: a.x + 0.2, y: by + (barH - 0.32) / 2, w: 0.84, h: 0.32, rectRadius: 0.04, fill: { color: GOLD } });
    slide.addText("核心要点", { x: a.x + 0.2, y: by + (barH - 0.32) / 2, w: 0.84, h: 0.32, fontSize: 9.5, color: WHITE, bold: true, fontFace: FONT, align: "center", margin: 0, valign: "middle" });
    slide.addText(items[items.length - 1], { x: a.x + 1.2, y: by, w: a.w - 1.4, h: barH, fontSize: 12.5, color: WHITE, bold: true, fontFace: FONT, align: "left", margin: 0, valign: "middle" });
  }
}

function slideStats(pres, slide, spec) {
  const stats = ((spec.extra || {}).stats) || [];
  if (!stats.length) return;
  const a = contentArea();
  const cols = stats.length <= 4 ? 2 : 3;
  const rows = Math.ceil(stats.length / cols);
  const gap = 0.22, m = 0.1;
  const cw = (a.w - 2 * m - gap * (cols - 1)) / cols;
  const chh = (a.h - gap * (rows - 1)) / rows;
  stats.forEach((st, i) => {
    const r = Math.floor(i / cols), c = i % cols;
    const x = a.x + m + c * (cw + gap), y = a.y + r * (chh + gap);
    slide.addShape(pres.ShapeType.roundRect, { x, y, w: cw, h: chh, rectRadius: 0.08, fill: { color: PAPER }, line: { color: CARDLINE, width: 0.75 }, shadow: paperShadow(pres) });
    slide.addText(String(st.number || ""), { x, y: y + 0.1, w: cw, h: chh * 0.5, fontSize: 38, color: NAVY, bold: true, fontFace: "Calibri", align: "center", margin: 0, valign: "middle" });
    const dlW = cw * 0.56, dlX = x + (cw - dlW) / 2, dlY = y + chh * 0.52 + 0.1;
    slide.addShape(pres.ShapeType.rect, { x: dlX, y: dlY, w: dlW, h: 0.014, fill: { color: HAIR } });
    slide.addShape(pres.ShapeType.rect, { x: dlX, y: dlY - 0.04, w: 0.016, h: 0.1, fill: { color: GOLD } });
    slide.addShape(pres.ShapeType.rect, { x: dlX + dlW - 0.016, y: dlY - 0.04, w: 0.016, h: 0.1, fill: { color: GOLD } });
    slide.addText(String(st.label || ""), { x: x + 0.2, y: dlY + 0.12, w: cw - 0.4, h: chh - dlY + y - 0.12, fontSize: 11, color: MUTE, fontFace: FONT, align: "center", margin: 0, valign: "top", lineSpacingMultiple: 1.2 });
  });
}

function slideTable(pres, slide, spec) {
  const ex = spec.extra || {};
  const headers = ex.headers || ex.columns || [];
  const rows = ex.rows || [];
  if (!headers.length || !rows.length) return;
  const a = contentArea();
  slide.addShape(pres.ShapeType.roundRect, { x: a.x, y: a.y, w: a.w, h: a.h, rectRadius: 0.05, fill: { color: PAPER }, line: { color: CARDLINE, width: 0.75 }, shadow: paperShadow(pres) });
  const nHead = headers.map((h) => ({
    text: String(h), options: {
      bold: true, color: NAVY, fontSize: 11.5, fontFace: FONT, align: "left", valign: "middle",
      fill: { color: PAPER },
      border: [{ pt: 2, color: NAVY }, { type: "none" }, { pt: 1, color: NAVY }, { type: "none" }],
    },
  }));
  const body = rows.map((r, i) => r.map((c, j) => ({
    text: String(c), options: {
      color: j === 0 ? NAVY : INK, bold: j === 0, fontSize: 10.5, fontFace: FONT, align: "left", valign: "middle",
      fill: { color: PAPER },
      border: [
        { type: "none" }, { type: "none" },
        i === rows.length - 1 ? { pt: 2, color: NAVY } : { pt: 0.5, color: "DCE4EE" },
        { type: "none" },
      ],
    },
  })));
  const pad = 0.12;
  const rowH = Math.min((a.h - pad * 2 - 0.5) / rows.length, 0.85);
  slide.addTable([nHead, ...body], {
    x: a.x + pad, y: a.y + pad, w: a.w - pad * 2,
    colW: Array(headers.length).fill((a.w - pad * 2) / headers.length),
    rowH: [0.5, ...Array(rows.length).fill(rowH)],
    margin: [0.04, 0.08, 0.04, 0.08],
  });
}

function slideFlow(pres, slide, spec) {
  const nodes = ((spec.extra || {}).nodes) || [];
  if (!nodes.length) return;
  const a = contentArea();
  const m = 0.15, gap = 0.28;
  const n = nodes.length;
  const layerH = Math.min((a.h - 2 * m - gap * (n - 1)) / n, 1.05);
  const cx = a.x + a.w / 2;
  nodes.forEach((node, li) => {
    const items = Array.isArray(node) ? node : [node];
    const nw = items.length;
    const gapX = 0.25;
    const bw = nw === 1 ? Math.min(a.w - 2 * m, 6.5) : (a.w - 2 * m - gapX * (nw - 1)) / nw;
    const y = a.y + m + li * (layerH + gap);
    items.forEach((item, xi) => {
      const x = nw === 1 ? cx - bw / 2 : a.x + m + xi * (bw + gapX);
      slide.addShape(pres.ShapeType.roundRect, { x, y, w: bw, h: layerH, rectRadius: 0.07, fill: { color: PAPER }, line: { color: NAVY, width: 1.1 }, shadow: paperShadow(pres) });
      slide.addText(String(item), { x: x + 0.12, y, w: bw - 0.24, h: layerH, fontSize: 12.5, color: NAVY, bold: true, fontFace: FONT, align: "center", margin: 0, valign: "middle", lineSpacingMultiple: 1.2 });
    });
    if (li < n - 1) {
      slide.addShape(pres.ShapeType.line, { x: cx, y: y + layerH + 0.02, w: 0, h: gap - 0.04, line: { color: NAVY, width: 1.2, endArrowType: "triangle" } });
    }
  });
}

function monthIdx(s) {
  const m = /(\d{4})[.\-/年](\d{1,2})/.exec(String(s || ""));
  return m ? parseInt(m[1]) * 12 + parseInt(m[2]) : null;
}

function slideGantt(pres, slide, spec) {
  const tasks = ((spec.extra || {}).tasks) || [];
  const spans = [];
  for (const t of tasks) {
    const time = String(t.time || "");
    const lo = monthIdx(time.split(/[–—~-]/)[0]);
    const hi = monthIdx(time.split(/[–—~-]/).pop()) || lo;
    if (lo != null) spans.push({ ...t, lo, hi: Math.max(lo, hi) });
  }
  if (!spans.length) return;
  const a = contentArea();
  const labelW = 2.1, axisH = 0.32, m = 0.1;
  const rowH = (a.h - axisH - m) / spans.length;
  const lo0 = Math.min(...spans.map(s => s.lo)) - 1;
  const hi0 = Math.max(...spans.map(s => s.hi)) + 1;
  const total = Math.max(hi0 - lo0, 1);
  const px0 = a.x + labelW, pw = a.x + a.w - px0;
  const X = (idx) => px0 + ((idx - lo0) / total) * pw;
  // 月份刻度
  const step = total > 14 ? 3 : total > 8 ? 2 : 1;
  for (let idx = lo0; idx <= hi0; idx++) {
    if ((idx - lo0) % step === 0) {
      const gx = X(idx);
      slide.addShape(pres.ShapeType.line, { x: gx, y: a.y + axisH, w: 0, h: a.h - axisH - m, line: { color: "E4EBF3", width: 0.6 } });
      const ym = Math.floor((idx - 1) / 12), mo = ((idx - 1) % 12) + 1;
      slide.addText(`${String(ym % 100).padStart(2, "0")}.${String(mo).padStart(2, "0")}`, { x: gx - 0.32, y: a.y, w: 0.64, h: 0.26, fontSize: 8.5, color: MUTE, fontFace: "Calibri", align: "center", margin: 0 });
    }
  }
  spans.forEach((t, i) => {
    const y = a.y + axisH + m + i * rowH;
    slide.addText(String(t.phase || ""), { x: a.x, y, w: labelW - 0.15, h: rowH, fontSize: 11.5, color: NAVY, bold: true, fontFace: FONT, align: "right", margin: 0, valign: "middle" });
    slide.addShape(pres.ShapeType.roundRect, { x: X(t.lo), y: y + rowH * 0.2, w: Math.max(X(t.hi) - X(t.lo), 0.12), h: rowH * 0.6, rectRadius: 0.03, fill: { color: i % 2 ? NAVY : BLUE }, shadow: paperShadow(pres) });
    if (X(t.hi) - X(t.lo) > 1.3) {
      slide.addText(String(t.phase || ""), { x: X(t.lo) + 0.1, y: y + rowH * 0.2, w: X(t.hi) - X(t.lo) - 0.2, h: rowH * 0.6, fontSize: 9.5, color: WHITE, bold: true, fontFace: FONT, align: "left", margin: 0, valign: "middle" });
    }
  });
}

function slidePanels(pres, slide, spec) {
  const panels = ((spec.extra || {}).panels) || [];
  if (!panels.length) return;
  const a = contentArea();
  const n = Math.min(panels.length, 4);
  const gap = 0.28;
  const pw = (a.w - gap * (n - 1)) / n;
  const titleH = 0.46;
  const ph = Math.min(a.h, 2.75);
  const y0 = a.y + (a.h - ph) / 2;
  panels.slice(0, n).forEach((p, i) => {
    const x = a.x + i * (pw + gap);
    slide.addShape(pres.ShapeType.roundRect, { x, y: y0, w: pw, h: ph, rectRadius: 0.06, fill: { color: PAPER }, line: { color: CARDLINE, width: 0.75 }, shadow: paperShadow(pres) });
    slide.addShape(pres.ShapeType.rect, { x, y: y0, w: pw, h: titleH, fill: { color: NAVY } });
    slide.addText(String(p.title || ""), { x: x + 0.18, y: y0, w: pw - 0.36, h: titleH, fontSize: 13.5, color: WHITE, bold: true, fontFace: FONT, align: "left", margin: 0, valign: "middle" });
    slide.addText(String(p.text || ""), { x: x + 0.2, y: y0 + titleH + 0.12, w: pw - 0.4, h: ph - titleH - 0.24, fontSize: 11, color: INK, fontFace: FONT, align: "left", margin: 0, valign: "top", lineSpacingMultiple: 1.4 });
  });
}

async function slideImage(pres, slide, spec) {
  const ex = spec.extra || {};
  const p = ex.image || "";
  const a = contentArea();
  if (!p) return;
  const dim = await imageSize(p);
  if (dim && dim.w && dim.h) {
    const capH = ex.caption ? 0.4 : 0.05;
    const scale = Math.min((a.w - 0.5) / dim.w, (a.h - capH - 0.1) / dim.h);
    const w = dim.w * scale, h = dim.h * scale;
    slide.addImage({ path: p, x: a.x + (a.w - w) / 2, y: a.y + (a.h - h - capH) / 2, w, h });
    if (ex.caption) slide.addText(ex.caption, { x: a.x, y: a.y + a.h - 0.4, w: a.w, h: 0.32, fontSize: 10.5, color: MUTE, fontFace: FONT, align: "center", margin: 0 });
  } else {
    slide.addText(`【图片：${p}】`, { x: a.x, y: a.y, w: a.w, h: a.h, fontSize: 14, color: MUTE, fontFace: FONT, align: "center", valign: "middle", margin: 0 });
  }
}

// ── 主流程 ────────────────────────────────────────────────────
async function main() {
  const argv = parseArgs();
  const deck = JSON.parse(fs.readFileSync(argv.content, "utf8"));
  const chapters = deck.chapters || [];
  const cover = deck.cover || {};
  const footerText = [deck.school || "北京航空航天大学", cover.培养学院, cover.专业名称].filter(Boolean).join(" · ");
  const total = 2 + chapters.length + chapters.reduce((a, c) => a + (c.slides || []).length, 0) + 1;

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = cover.作者姓名 || "";
  pres.title = deck.title || "";
  pres.subject = "开题汇报";

  slideCover(pres, deck, total);
  slideToc(pres, deck, chapters.map(c => c.name));

  let page = 3;
  for (let ci = 0; ci < chapters.length; ci++) {
    const ch = chapters[ci];
    const ic = await iconData(SECTION_ICONS[ch.name] || "FaCompass", GOLD);
    const sec = slideSection(pres, deck, ci + 1, ch.name, ic, footerText, page, total);
    if (ch.slides && ch.slides[0] && ch.slides[0].notes) sec.addNotes(ch.slides[0].notes);
    page++;
    for (const spec of (ch.slides || [])) {
      const slide = pres.addSlide();
      const layout = spec.layout || "text_only";
      header(pres, slide, ch.name, spec.title || ch.name);
      if (layout === "cards" || layout === "text_only") slideCards(pres, slide, spec);
      else if (layout === "stats") slideStats(pres, slide, spec);
      else if (layout === "compare" || layout === "table") slideTable(pres, slide, spec);
      else if (layout === "flow") slideFlow(pres, slide, spec);
      else if (layout === "gantt") slideGantt(pres, slide, spec);
      else if (layout === "panels") slidePanels(pres, slide, spec);
      else if (layout === "image_center") await slideImage(pres, slide, spec);
      if (spec.notes) slide.addNotes(spec.notes);
      footer(pres, slide, footerText, page, total);
      page++;
    }
  }

  const thx = pres.addSlide();
  thx.background = { color: NAVY };
  thx.addShape(pres.ShapeType.rect, { x: 0.22, y: 0.22, w: W - 0.44, h: H - 0.44, fill: { type: "none" }, line: { color: "3D5F8F", width: 0.75 } });
  thx.addText("汇报完毕", { x: 0, y: 2.3, w: W, h: 1.0, fontSize: 40, color: WHITE, bold: true, fontFace: FONT, align: "center", margin: 0 });
  thx.addText("恳请各位老师批评指正", { x: 0, y: 3.4, w: W, h: 0.5, fontSize: 15, color: SKY, fontFace: FONT, align: "center", margin: 0 });
  thx.addShape(pres.ShapeType.rect, { x: W / 2 - 1.1, y: 4.15, w: 2.2, h: 0.045, fill: { color: GOLD } });
  thx.addText(`汇报人：${cover.作者姓名 || ""} · ${cover.日期 || ""}`, { x: 0, y: 4.6, w: W, h: 0.4, fontSize: 13, color: "E6EEF6", fontFace: FONT, align: "center", margin: 0 });

  await pres.writeFile({ fileName: argv.output });
  console.log("saved:", argv.output, `(${total} 页)`);
}

main().catch((e) => { console.error(e); process.exit(1); });
