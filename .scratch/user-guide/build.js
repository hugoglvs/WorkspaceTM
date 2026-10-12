// WorkspacePM User Guide — v2 layout engine: fixed grid, measured text, no body placeholders.
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const { applyTheme } = require("/mnt/skills/public/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "WorkspacePM User Guide.pptx";
const THEME = {
  name: "WorkspacePM", headFontFace: "Arial", bodyFontFace: "Calibri",
  colors: { dk1: "1E1B2E", lt1: "FFFFFF", dk2: "3B1E66", lt2: "F5F1FA", accent1: "7500C0", accent2: "1967D2", accent3: "137333", accent4: "C5221F", accent5: "B06000", accent6: "5F6368", hlink: "1967D2", folHlink: "7500C0" },
};
const HX = {
  ink: "1E1B2E", deep: "3B1E66", lav: "F5F1FA", purple: "7500C0", blue: "1967D2", green: "137333", red: "C5221F", amber: "B06000", grey: "5F6368",
  tPurple: "F3E8FD", tBlue: "E8F0FE", tGreen: "E6F4EA", tRed: "FCE8E6", tGrey: "F1F3F4", tAmber: "FEF7E0", tOrange: "FDEBD0", orange: "B05C00",
  line: "DADCE0", soft: "80868B", body: "3C4043", white: "FFFFFF", lilac: "D9CCF0",
};

// ---------- grid ----------
const X0 = 0.6, X1 = 12.733, W = X1 - X0;
const Y0 = 1.7, Y1 = 6.7, H = Y1 - Y0;
const GAP = 0.3;
const cols = (n, gap = GAP, x = X0, w = W) => { const cw = (w - gap * (n - 1)) / n; return Array.from({ length: n }, (_, i) => ({ x: x + i * (cw + gap), w: cw })); };

// ---------- text measurement (conservative) ----------
function wrapLines(str, pt, w, face, bold) {
  const k = (face === "Arial" ? 0.5 : face === "Courier New" ? 0.6 : 0.47) * (bold ? 1.06 : 1);
  const cw = (pt / 72) * k;
  const maxc = Math.max(1, Math.floor(w / cw));
  let lines = 0;
  String(str).split("\n").forEach((para) => {
    let cur = 0; lines++;
    para.replace(/-/g, "- ").split(/\s+/).filter(Boolean).forEach((word) => {
      const L = word.length;
      if (cur === 0) { cur = L; while (cur > maxc) { lines++; cur -= maxc; } }
      else if (cur + 1 + L <= maxc) cur += 1 + L;
      else { lines++; cur = L; }
    });
  });
  return lines;
}
const LH = 1.2;
const lineH = (pt) => (pt / 72) * LH;
function textH(str, pt, w, face = "Calibri", bold = false) { return wrapLines(str, pt, w, face, bold) * lineH(pt); }
const itemText = (it) => (Array.isArray(it) ? it[0] + it[1] : it);
function listH(items, pt, w, space, bullet = true) {
  const tw = bullet ? w - 0.22 : w;
  return items.reduce((a, it) => a + textH(itemText(it), pt, tw, "Calibri", Array.isArray(it)), 0) + (space / 72) * (items.length - 1);
}

// ---------- pres + layouts ----------
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "WorkspacePM User Guide"; pres.author = "Hugo"; pres.company = "WorkspacePM";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
pres.defineSlideMaster({ title: "COVER", background: { color: HX.ink }, objects: [
  { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.2, w: 11.7, h: 1.0, fontSize: 54, bold: true, color: C.background1, fontFace: "Arial", valign: "top", align: "left", margin: 0 }, text: "" } },
] });
pres.defineSlideMaster({ title: "SECTION", background: { color: HX.ink }, objects: [
  { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.05, w: 6.9, h: 1.6, fontSize: 44, bold: true, color: C.background1, fontFace: "Arial", valign: "top", align: "left", margin: 0 }, text: "" } },
], slideNumber: { x: 12.13, y: 6.95, w: 0.6, h: 0.3, fontSize: 10, color: "9A8FB5", align: "right" } });
pres.defineSlideMaster({ title: "CONTENT", background: { color: HX.white }, margin: [0.6, 0.6, 0.7, 0.6], objects: [
  { placeholder: { options: { name: "title", type: "title", x: X0, y: 0.72, w: W, h: 0.7, fontSize: 32, bold: true, color: C.text1, fontFace: "Arial", valign: "top", align: "left", margin: 0 }, text: "" } },
  { text: { text: "WorkspacePM User Guide", options: { x: X0, y: 6.98, w: 5, h: 0.3, fontSize: 10, color: HX.soft, margin: 0 } } },
], slideNumber: { x: X1 - 0.6, y: 6.98, w: 0.6, h: 0.3, fontSize: 10, color: HX.soft, align: "right" } });

// ---------- icons ----------
const iconCache = {};
async function icon(name, color) {
  const key = name + color;
  if (!iconCache[key]) {
    const Comp = md[name]; if (!Comp) throw new Error("icon " + name);
    const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
    iconCache[key] = "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
  }
  return iconCache[key];
}

// ---------- primitives ----------
let oid = 0; const nid = (p) => `${p}-${++oid}`;
function T(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, objectName: nid("text"), fontFace: "Calibri", color: HX.ink, margin: 0, valign: "top", fit: "none" }, o)); }
function pillW(text, pt = 12) { return (text.length * (pt / 72) * 0.56) + 0.36; }
function pill(s, text, x, y, o = {}) {
  const pt = o.fontSize || 12; const w = o.w || pillW(text, pt); const h = o.h || 0.32;
  s.addText(text, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: h / 2, x, y, w, h, fill: { color: o.fill || HX.tPurple }, color: o.color || HX.purple, fontSize: pt, bold: true, align: "center", valign: "middle", margin: 0, fontFace: "Calibri", objectName: nid("pill") });
  return w;
}
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(o.rect ? pres.shapes.RECTANGLE : pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, rectRadius: o.r == null ? 0.08 : o.r, line: o.line ? { color: o.line, width: 1 } : { type: "none" }, objectName: nid("box") });
}
async function iconCircle(s, name, x, y, d, bg, fg) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" }, objectName: nid("iconbg") });
  const p = d * 0.22; s.addImage({ data: await icon(name, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: nid("icon") });
}
function bullets(s, items, x, y, w, o = {}) {
  const pt = o.fontSize || 15, space = o.space == null ? 6 : o.space, bul = o.bullet !== false;
  const runs = [];
  items.forEach((it, i) => {
    const last = i === items.length - 1;
    const base = { paraSpaceAfter: last ? 0 : space, fontSize: pt, color: o.color || HX.ink };
    if (bul) base.bullet = o.numbered ? { type: "number", indent: 16 } : { indent: 14 };
    if (Array.isArray(it)) {
      runs.push({ text: it[0], options: Object.assign({}, base, { bold: true, color: o.boldColor || o.color || HX.ink }) });
      runs.push({ text: it[1], options: Object.assign({}, base, { bullet: undefined, breakLine: !last }) });
    } else runs.push({ text: it, options: Object.assign({}, base, { breakLine: !last }) });
  });
  const h = listH(items, pt, w, space, bul);
  T(s, runs, { x, y, w, h: h + 0.1, fontSize: pt });
  return h;
}
function hLine(s, x1, y, x2, color, arrowEnd = true) { s.addShape(pres.shapes.LINE, { x: x1, y, w: x2 - x1, h: 0, line: { color: color || HX.soft, width: 1.5, endArrowType: arrowEnd ? "triangle" : undefined }, objectName: nid("arrow") }); }
function vLine(s, x, y1, y2, color, arrowEnd = true) {
  const o = { x, y: Math.min(y1, y2), w: 0, h: Math.abs(y2 - y1), line: { color: color || HX.soft, width: 1.5, endArrowType: arrowEnd ? "triangle" : undefined }, objectName: nid("arrow") };
  if (y2 < y1) o.flipV = true; s.addShape(pres.shapes.LINE, o);
}

// ---------- card: icon left (optional), title, body (string or list) ----------
const PAD = 0.25, ICON = 0.55;
function cardGeom(o) {
  const inner = o.w - 2 * PAD - (o.icon ? ICON + 0.2 : 0);
  const tpt = o.titleSize || 17, bpt = o.bodySize || 15;
  const th = o.title ? Math.max(o.titleH || 0, textH(o.title, tpt, inner, "Arial", true)) : 0;
  let bh = 0;
  if (o.body) bh = Array.isArray(o.body) ? listH(o.body, bpt, inner, o.space == null ? 5 : o.space, o.bulletList !== false) : textH(o.body, bpt, inner);
  const need = PAD + Math.max(o.icon ? ICON : 0, th + (o.body ? 0.12 + bh : 0)) + PAD;
  return { inner, tpt, bpt, th, bh, need };
}
async function card(s, o) {
  const g = cardGeom(o);
  const h = o.h || g.need;
  box(s, o.x, o.y, o.w, h, o.fill || HX.lav, { line: o.line });
  const tx = o.x + PAD + (o.icon ? ICON + 0.2 : 0);
  if (o.icon) await iconCircle(s, o.icon, o.x + PAD, o.y + PAD, ICON, o.iconBg || HX.white, o.iconColor || HX.purple);
  let ty = o.y + PAD + (o.icon ? Math.max(0, (ICON - lineH(g.tpt)) / 2) - 0.02 : 0);
  if (o.title) { T(s, o.title, { x: tx, y: ty, w: g.inner, h: g.th + 0.05, fontSize: g.tpt, bold: true, color: o.titleColor || HX.ink, fontFace: "Arial" }); ty += g.th + 0.12; }
  if (o.body) {
    if (Array.isArray(o.body)) bullets(s, o.body, tx, ty, g.inner, { fontSize: g.bpt, space: o.space == null ? 5 : o.space, bullet: o.bulletList !== false });
    else T(s, o.body, { x: tx, y: ty, w: g.inner, h: g.bh + 0.08, fontSize: g.bpt, color: o.bodyColor || HX.body });
  }
  return h;
}
// a row of cards with equal height; returns height
async function cardRow(s, items, y, n, opts = {}) {
  const cs = cols(n, opts.gap == null ? GAP : opts.gap, opts.x == null ? X0 : opts.x, opts.w == null ? W : opts.w);
  const g0 = items.map((it, i) => cardGeom(Object.assign({}, opts.base || {}, it, { w: cs[i].w })));
  const thMax = Math.max(...g0.map((g) => g.th));
  items = items.map((it) => Object.assign({}, it, { titleH: thMax }));
  const geo = items.map((it, i) => cardGeom(Object.assign({}, opts.base || {}, it, { w: cs[i].w })));
  const h = opts.fixedH || (Math.max(opts.minH || 0, ...geo.map((g) => g.need)) + (opts.extra || 0));
  for (let i = 0; i < items.length; i++) await card(s, Object.assign({}, opts.base || {}, items[i], { x: cs[i].x, y, w: cs[i].w, h }));
  return h;
}
function rowsNeed(items, n, opts = {}) {
  const cs = cols(n, opts.gap == null ? GAP : opts.gap, opts.x == null ? X0 : opts.x, opts.w == null ? W : opts.w);
  return Math.max(...items.map((it, i) => cardGeom(Object.assign({}, opts.base || {}, it, { w: cs[i % n].w })).need));
}

// ---------- table with measured rows ----------
function table(s, rows, o) {
  const pt = o.fontSize || 14;
  const colW = o.colW;
  const rowH = rows.map((r, ri) => {
    const m = Math.max(...r.map((c, j) => wrapLines(c, pt, colW[j] - 0.22, "Calibri", ri === 0 || j === 0) * lineH(pt)));
    return Math.max(o.minRowH || 0.4, m + 0.16);
  });
  const total = rowH.reduce((a, b) => a + b, 0);
  if (o.fillH && total < o.fillH) { const add = (o.fillH - total) / rows.length; for (let i = 0; i < rowH.length; i++) rowH[i] += add; }
  const head = rows[0].map((h) => ({ text: h, options: { bold: true, color: HX.white, fill: { color: HX.deep }, fontSize: pt, fontFace: "Calibri" } }));
  const body = rows.slice(1).map((r, i) => r.map((c, j) => ({ text: c, options: { fontSize: pt, color: HX.ink, fill: { color: i % 2 ? HX.white : HX.lav }, bold: j === 0 && o.boldFirst !== false, fontFace: "Calibri" } })));
  s.addTable([head].concat(body), { x: o.x, y: o.y, w: colW.reduce((a, b) => a + b, 0), colW, rowH, border: { type: "solid", color: HX.line, pt: 0.75 }, margin: [0.05, 0.11, 0.05, 0.11], valign: "middle", objectName: nid("table") });
  return rowH.reduce((a, b) => a + b, 0);
}

// ---------- slide factories ----------
let curSection = "", kicker = "";
function section(t) { curSection = t; pres.addSection({ title: t }); }
function content(title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: curSection });
  s.addText(title, { placeholder: "title" });
  pill(s, kicker, X0, 0.3, { fontSize: 11, h: 0.3 });
  return s;
}
function heading(s, text, x, y, w, color) { T(s, text, { x, y, w, h: 0.35, fontSize: 16, bold: true, color: color || HX.deep, fontFace: "Arial" }); return 0.45; }
function note(s, text, y) { T(s, text, { x: X0, y, w: W, h: 0.3, fontSize: 13, color: HX.soft, italic: true }); }
async function divider(num, title, body, list) {
  const s = pres.addSlide({ masterName: "SECTION", sectionTitle: curSection });
  pill(s, "Section " + num, 0.8, 1.45, { fill: HX.purple, color: HX.white, fontSize: 12, h: 0.36 });
  s.addText(title, { placeholder: "title" });
  const th = textH(title, 44, 6.9, "Arial", true);
  T(s, body, { x: 0.8, y: 2.05 + th + 0.35, w: 6.6, h: textH(body, 18, 6.6) + 0.1, fontSize: 18, color: HX.lilac });
  const lh = listH(list, 16, 3.6, 8, false);
  const bh = 0.35 + 0.5 + lh + 0.35;
  box(s, 8.4, 1.45, 4.33, bh, "2A2640", { r: 0.06 });
  T(s, "In this section", { x: 8.75, y: 1.8, w: 3.6, h: 0.3, fontSize: 12, bold: true, color: "B9A6E0" });
  bullets(s, list, 8.75, 2.3, 3.6, { fontSize: 16, color: HX.white, space: 8, bullet: false });
}
const CHIPS = [["Office", HX.tPurple, HX.purple], ["Client site", HX.tBlue, HX.blue], ["Remote", HX.tGreen, HX.green], ["Leave", HX.tRed, HX.red], ["Off", HX.tGrey, HX.grey], ["Not Logged", HX.tAmber, HX.amber]];

// =====================================================================
(async () => {
  section("Introduction");
  // ---- 1 cover
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: curSection });
    pill(s, "WorkspacePM · Google Sheets + Apps Script", 0.8, 1.5, { fill: HX.purple, color: HX.white, fontSize: 12, h: 0.38 });
    s.addText("WorkspacePM User Guide", { placeholder: "title" });
    T(s, "Team presence, contracts, projects and coverage in one Google Sheet: what it does, how to set it up, and how it works under the cover.", { x: 0.8, y: 3.45, w: 10.5, h: 1.0, fontSize: 20, color: HX.lilac });
    let x = 0.8; CHIPS.forEach(([t, f, c]) => { x += pill(s, t, x, 4.95, { fill: f, color: c, fontSize: 13, h: 0.38 }) + 0.15; });
    T(s, "Version of 12 October 2026 · Library build deploy-20261012-0813", { x: 0.8, y: 6.5, w: 8, h: 0.3, fontSize: 12, color: "9A8FB5" });
  }
  // ---- 2 audiences
  kicker = "Start here";
  {
    const s = content("Who this guide is for, and where to look");
    const aud = [["MdGroups", "Team members", "Sections 1–2", "What the sheet shows, how to declare your week and leaves, how projects stay covered, which emails you will get."],
      ["MdAdminPanelSettings", "Administrators", "Section 3", "Linking the Library, running Initialize / Upgrade, configuring locations, holidays and settings, maintaining triggers."],
      ["MdCode", "Maintainers", "Sections 4–6", "Architecture, data flows, the nightly job, the Sync Lock, the flexibility-to-integrity ratio and troubleshooting."]];
    const cs = cols(3);
    const bh = Math.max(...aud.map((a, i) => textH(a[3], 15, cs[i].w - 2 * PAD)));
    const ch = PAD + 0.7 + 0.25 + 0.4 + 0.1 + bh + PAD + 0.1;
    for (let i = 0; i < 3; i++) {
      const [ic, t, sec, b] = aud[i]; const { x, w } = cs[i];
      box(s, x, Y0, w, ch, HX.lav);
      await iconCircle(s, ic, x + PAD, Y0 + PAD, 0.7, HX.white, HX.purple);
      const pw = pillW(sec); pill(s, sec, x + w - PAD - pw, Y0 + PAD + 0.19);
      T(s, t, { x: x + PAD, y: Y0 + PAD + 0.95, w: w - 2 * PAD, h: 0.4, fontSize: 19, bold: true, fontFace: "Arial" });
      T(s, b, { x: x + PAD, y: Y0 + PAD + 1.45, w: w - 2 * PAD, h: bh + 0.1, fontSize: 15, color: HX.body });
    }
    let y = Y0 + ch + 0.5;
    y += heading(s, "Six sections", X0, y, W);
    const secs = ["What the system does", "Using it day to day", "Setting it up", "Under the cover", "Flexibility vs. data integrity", "Troubleshooting & reference"];
    secs.forEach((t, i) => {
      const { x } = cs[i % 3]; const yy = y + Math.floor(i / 3) * 0.6;
      s.addShape(pres.shapes.OVAL, { x, y: yy, w: 0.4, h: 0.4, fill: { color: HX.purple }, line: { type: "none" }, objectName: nid("num") });
      T(s, String(i + 1), { x, y: yy, w: 0.4, h: 0.4, fontSize: 14, bold: true, color: HX.white, align: "center", valign: "middle" });
      T(s, t, { x: x + 0.55, y: yy, w: cs[0].w - 0.6, h: 0.4, fontSize: 16, bold: true, color: HX.deep, valign: "middle" });
    });
  }

  // ================= SECTION 1
  section("1 · What the system does"); kicker = "1 · What the system does";
  await divider(1, "What the system does", "One shared Google Sheet becomes the team's operating picture: who works where, under which contract, who covers which project, and what is due.", ["Six jobs in one sheet", "The five visible tabs", "The menu", "Key terms", "Colour legend", "What runs automatically"]);
  {
    const s = content("One Google Sheet runs the team's day to day");
    const intro = "WorkspacePM is a Google Apps Script Library behind a normal Google Sheet. People type in grids or use guided sidebars; the Library keeps every table consistent and talks to Google Tasks, Calendar and Gmail.";
    const ih = textH(intro, 17, W); T(s, intro, { x: X0, y: Y0, w: W, h: ih + 0.1, fontSize: 17, color: HX.body });
    const caps = [["MdBadge", "Directory & contracts", "Members with role, department and status derived from their contracts."],
      ["MdEventAvailable", "Presence & leaves", "Where each person works, morning and afternoon: routines, one-off changes, leaves, public holidays."],
      ["MdShield", "Projects & coverage", "Lead, permanent backup and interim cover; who is in charge today; alerts before a gap happens."],
      ["MdTaskAlt", "Tasks & events", "Tasks sync with Google Tasks; events are created in Google Calendar."],
      ["MdNotificationsActive", "Reminders", "Monday presence reminder, birthday emails and a shared Team Birthdays calendar."],
      ["MdInsights", "KPIs", "A tidy KPI Data table, one row per person, day and half-day, for Looker Studio."]].map(([icon, title, body]) => ({ icon, title, body }));
    const top = Y0 + ih + 0.4;
    const rh = (Y1 - top - GAP) / 2;
    await cardRow(s, caps.slice(0, 3), top, 3, { fixedH: rh, base: { bodySize: 14, titleSize: 16 } });
    await cardRow(s, caps.slice(3), top + rh + GAP, 3, { fixedH: rh, base: { bodySize: 14, titleSize: 16 } });
  }
  {
    const s = content("Five tabs are visible; the rest work behind them");
    let y = Y0; y += heading(s, "Visible to everyone", X0, y, W, HX.grey);
    const vis = [["Members", "Directory: name, email, phone, role, department, status, today, birthday", HX.purple], ["Weekly Presence", "Days × people grid, AM and PM, colour-coded; type to change a day", HX.green], ["Projects", "Lead, backup and who is in charge today; red or orange when coverage fails", HX.blue], ["Tasks", "Work items with IDs, synced with Google Tasks", HX.amber], ["Events", "Team events, created in Google Calendar", HX.red]];
    const c5 = cols(5);
    const dh = Math.max(...vis.map((v, i) => textH(v[1], 15, c5[i].w - 0.4))) + 0.4;
    vis.forEach(([t, b, c], i) => {
      const { x, w } = c5[i];
      box(s, x, y, w, 0.5 + dh, HX.lav);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.5, fill: { color: c }, rectRadius: 0.08, line: { type: "none" }, objectName: nid("tab") });
      T(s, t, { x, y, w, h: 0.5, fontSize: 16, bold: true, color: HX.white, align: "center", valign: "middle" });
      T(s, b, { x: x + 0.2, y: y + 0.7, w: w - 0.4, h: dh - 0.3, fontSize: 15, color: HX.body });
    });
    y += 0.5 + dh + 0.45;
    y += heading(s, "Hidden by default — never locked", X0, y, W, HX.grey);
    const hid = [["Contracts", "contract history"], ["Routines", "weekly defaults"], ["Leaves", "absences"], ["Presence Log", "the daily record"], ["Backups", "permanent and interim cover"], ["Locations", "places and types"], ["Holidays", "public holidays"], ["KPI Data", "facts for Looker"]];
    const c4 = cols(4); const hh = Math.min(0.8, (Y1 - y - 0.2) / 2);
    hid.forEach(([t, b], i) => {
      const { x, w } = c4[i % 4]; const yy = y + Math.floor(i / 4) * (hh + 0.2);
      box(s, x, yy, w, hh, HX.tGrey);
      T(s, [{ text: t, options: { bold: true, color: HX.ink, breakLine: true } }, { text: b, options: { color: HX.grey, fontSize: 14 } }], { x: x + 0.2, y: yy, w: w - 0.4, h: hh, fontSize: 16, valign: "middle" });
    });
  }
  {
    const s = content("Everything starts from the Project Management menu");
    const mw = 4.2;
    box(s, X0, Y0, mw, H, HX.white, { line: HX.line });
    T(s, "Project Management ▾", { x: X0 + 0.3, y: Y0 + 0.25, w: mw - 0.6, h: 0.4, fontSize: 17, bold: true, color: HX.deep });
    const items = ["Initialize / Upgrade Spreadsheet", "—", "Declare Presence & Leaves", "Members & Contracts", "Interim Backups", "—", "Check Workspace", "Edit Locations", "Edit Holidays"];
    const step = (H - 1.0) / 9.0; let y = Y0 + 0.85;
    items.forEach((it) => {
      if (it === "—") { s.addShape(pres.shapes.LINE, { x: X0 + 0.3, y: y + step * 0.25, w: mw - 0.6, h: 0, line: { color: HX.line, width: 1 }, objectName: nid("sep") }); y += step * 0.5; return; }
      T(s, it, { x: X0 + 0.3, y, w: mw - 0.6, h: step, fontSize: 16, valign: "middle" }); y += step;
    });
    const tx = X0 + mw + 0.4, tw = X1 - tx;
    table(s, [["Item", "Who", "What it does"],
      ["Initialize / Upgrade", "Admin", "Builds or repairs every table, installs the triggers, runs a full sync, shows the five tabs"],
      ["Declare Presence & Leaves", "Everyone", "Sidebar: set your weekly routine or declare a leave"],
      ["Members & Contracts", "Everyone", "Sidebar: add a person, record a new contract, end one, edit contact details"],
      ["Interim Backups", "Everyone", "Sidebar: cover a project while its lead is away"],
      ["Check Workspace", "Everyone", "Report of everything that needs attention, each with a link to the fix"],
      ["Edit Locations / Holidays", "Admin", "Opens the hidden Locations or Holidays tab"]], { x: tx, y: Y0, colW: [2.35, 1.05, tw - 3.4], fontSize: 15, fillH: H });
  }
  {
    const s = content("Key terms used throughout the sheet");
    const terms = [["Member", "A person, identified by their Google people chip."], ["Contract", "Role, department, dates, type and FTE of one assignment."], ["Period", "A half-day: morning (AM) or afternoon (PM)."], ["Work Location", "A place of type Office, Client site or Remote."], ["Routine", "Default AM/PM locations per weekday, from a start date."], ["Leave", "Vacation, sick, parental or unpaid; Requested, Approved or Cancelled."], ["Presence Log", "Where each Member was, per day, up to today."], ["Public Holiday", "A day in Holidays, treated like a weekend everywhere."], ["Project Backup", "The permanent second person on a project."], ["Interim Backup", "A dated replacement covering a specific absence."], ["In charge today", "Lead, else interim, else backup, else “Nobody available”."], ["Undeclared Day", "A working day this week with no entry, routine or leave."]];
    const c2 = cols(2); const rh = (H - 5 * 0.15) / 6;
    terms.forEach(([t, d], i) => {
      const { x, w } = c2[i % 2]; const row = Math.floor(i / 2); const y = Y0 + row * (rh + 0.15);
      box(s, x, y, w, rh, row % 2 ? HX.white : HX.lav, { line: row % 2 ? HX.line : null });
      T(s, t, { x: x + 0.25, y, w: 1.9, h: rh, fontSize: 16, bold: true, color: HX.purple, valign: "middle" });
      T(s, d, { x: x + 2.25, y, w: w - 2.45, h: rh, fontSize: 15, color: HX.body, valign: "middle" });
    });
  }
  {
    const s = content("Colours mean the same thing everywhere");
    const ch = [["Office", "OneRoof", HX.tPurple, HX.purple], ["Client site", "JLL, C04, C60", HX.tBlue, HX.blue], ["Remote", "Remote", HX.tGreen, HX.green], ["Leave", "Vacation, Sick Leave and other leave types", HX.tRed, HX.red], ["Off", "Weekend, public holiday", HX.tGrey, HX.grey], ["Not Logged", "No entry for today yet", HX.tAmber, HX.amber]];
    const c6 = cols(6, 0.2); const chH = 1.55;
    ch.forEach(([t, ex, f, c], i) => {
      const { x, w } = c6[i]; box(s, x, Y0, w, chH, f);
      T(s, t, { x: x + 0.18, y: Y0 + 0.18, w: w - 0.36, h: 0.4, fontSize: 18, bold: true, color: c, fontFace: "Arial" });
      T(s, ex, { x: x + 0.18, y: Y0 + 0.65, w: w - 0.36, h: 0.8, fontSize: 14, color: c });
    });
    const cap = "Presence colours come from each location's Type in the Locations tab, so a new client site turns blue as soon as it is listed. They appear in Weekly Presence and in the Members “Today” column.";
    const capY = Y0 + chH + 0.3; const capH = textH(cap, 16, W);
    T(s, cap, { x: X0, y: capY, w: W, h: capH + 0.1, fontSize: 16, color: HX.body });
    const other = [[HX.tRed, HX.red, "Project row red", "Coverage conflict: approved absences leave a working day uncovered."], [HX.tOrange, HX.orange, "Project row orange", "Coverage warning: a Requested leave would leave a day uncovered."], ["F8D7DA", "721C24", "Cell red, bold text", "A required field is empty in a row you started."]];
    const c3 = cols(3); const y = capY + capH + 0.35; const hh = Y1 - y;
    other.forEach(([f, c, t, d], i) => {
      const { x, w } = c3[i]; box(s, x, y, w, hh, HX.white, { line: HX.line });
      box(s, x + PAD, y + PAD, w - 2 * PAD, 0.48, f, { r: 0.04 });
      T(s, t, { x: x + PAD + 0.15, y: y + PAD, w: w - 2 * PAD - 0.3, h: 0.48, fontSize: 15, bold: true, color: c, valign: "middle" });
      T(s, d, { x: x + PAD, y: y + PAD + 0.65, w: w - 2 * PAD, h: hh - 1.0, fontSize: 15, color: HX.body });
    });
  }
  {
    const s = content("What runs automatically, and when");
    const cl = [["MdBolt", "On every edit", "Within seconds", ["IDs created for new tasks, events and projects", "Tasks pushed to Google Tasks, events to Calendar", "Members row refreshed from contracts", "Coverage re-checked on projects, leaves, backups", "Missing required fields named in a toast"]],
      ["MdNightsStay", "Every night", "2:00", ["Directory sync; presence compiled up to today", "Weekly Presence and KPI Data rebuilt", "Coverage check; emails and tasks for gaps", "Google Tasks and Calendar reconciled", "Holidays seeded; birthday calendar updated"]],
      ["MdMarkEmailUnread", "Monday morning", "8:00", ["Email to each active Member with undeclared days this week", "Lists the days and links to Weekly Presence"]],
      ["MdCake", "Every working day", "9:00", ["Email announcing tomorrow's birthdays; Friday covers the weekend and Monday", "Nothing is sent on weekends or public holidays"]]];
    const c4 = cols(4);
    const lh = Math.max(...cl.map((c, i) => listH(c[3], 14, c4[i].w - 2 * PAD, 5)));
    const hh = Math.min(H, PAD + 0.65 + 0.2 + 0.45 + 0.1 + lh + PAD + 0.1);
    for (let i = 0; i < 4; i++) {
      const [ic, t, when, items] = cl[i]; const { x, w } = c4[i];
      box(s, x, Y0, w, hh, HX.lav);
      await iconCircle(s, ic, x + PAD, Y0 + PAD, 0.65, HX.purple, HX.white);
      const pw = pillW(when); pill(s, when, x + w - PAD - pw, Y0 + PAD + 0.16, { fill: HX.white });
      T(s, t, { x: x + PAD, y: Y0 + PAD + 0.85, w: w - 2 * PAD, h: 0.4, fontSize: 18, bold: true, fontFace: "Arial" });
      bullets(s, items, x + PAD, Y0 + PAD + 1.4, w - 2 * PAD, { fontSize: 14, space: 5 });
    }
  }

  // ================= SECTION 2
  section("2 · Using it day to day"); kicker = "2 · Using it day to day";
  await divider(2, "Using it day to day", "For every team member. Declare your week once, change single days when plans move, and let the sheet keep projects covered.", ["Declaring where you work", "The Presence & Leaves sidebar", "Weekly Presence", "Members & Contracts", "Projects and coverage", "Interim backups", "Tasks and events", "Check Workspace and emails", "Good habits"]);
  {
    const s = content("Three ways to say where you work");
    const st = [["MdRepeat", "1 · Routine", "Your usual week", "Set once in Declare Presence & Leaves: AM and PM location for each weekday, from a start date. It fills every day automatically.", HX.purple, HX.tPurple],
      ["MdEditCalendar", "2 · One-off change", "A single half-day", "Type in your AM or PM cell in Weekly Presence. It is saved immediately as that day's record.", HX.green, HX.tGreen],
      ["MdBeachAccess", "3 · Leave", "An absence", "Declare it in the sidebar, as Requested then Approved. Half-days are supported with AM/PM on the first and last day.", HX.red, HX.tRed]];
    const c3 = cols(3);
    const bh = Math.max(...st.map((v, i) => textH(v[3], 16, c3[i].w - 2 * PAD)));
    const ch = PAD + 0.7 + 0.3 + bh + PAD + 0.05;
    for (let i = 0; i < 3; i++) {
      const [ic, t, sub, b, c, f] = st[i]; const { x, w } = c3[i];
      box(s, x, Y0, w, ch, f);
      await iconCircle(s, ic, x + PAD, Y0 + PAD, 0.7, HX.white, c);
      T(s, t, { x: x + PAD + 0.9, y: Y0 + PAD + 0.02, w: w - 1.4, h: 0.38, fontSize: 18, bold: true, color: c, fontFace: "Arial" });
      T(s, sub, { x: x + PAD + 0.9, y: Y0 + PAD + 0.4, w: w - 1.4, h: 0.3, fontSize: 14, color: HX.grey });
      T(s, b, { x: x + PAD, y: Y0 + PAD + 1.0, w: w - 2 * PAD, h: bh + 0.1, fontSize: 16, color: HX.body });
    }
    let y = Y0 + ch + 0.5;
    y += heading(s, "Which one wins for a given day?", X0, y, W);
    const lad = [["Hand entry", HX.green], ["Approved leave", HX.red], ["Routine", HX.purple], ["Default office (OneRoof)", HX.grey]];
    const lw = (W - 3 * 0.5) / 4;
    lad.forEach(([t, c], i) => {
      const x = X0 + i * (lw + 0.5);
      box(s, x, y, lw, 0.55, c, { r: 0.25 });
      T(s, t, { x, y, w: lw, h: 0.55, fontSize: 16, bold: true, color: HX.white, align: "center", valign: "middle" });
      if (i < 3) T(s, "›", { x: x + lw, y, w: 0.5, h: 0.55, fontSize: 26, bold: true, color: HX.grey, align: "center", valign: "middle" });
    });
    note(s, "A day typed by hand is never overwritten. If you typed a day before your leave was approved, change that day in Weekly Presence.", y + 0.8);
  }
  {
    const s = content("The Presence & Leaves sidebar");
    const panels = [["Routine Baseline", ["Team Member: you, pre-selected", "Quick Presets: All Office, All Remote, Hybrid (Mon/Fri Remote)", "Weekly Schedule: AM and PM per weekday; + Weekend Days if you work them", "Effective Start Date; End Date optional", "Save Baseline"]],
      ["Submit Leave", ["Leave Type: Vacation, Sick Leave, Maternity/Paternity, Unpaid Leave", "Start and End Date, each with AM / PM for half-days", "Quick ranges: Today Only, Current Week, Next Week", "Status: Requested or Approved", "Submit Leave"]]];
    const pw = 3.55; const lhs = panels.map((p) => listH(p[1], 15, pw - 0.4, 7));
    const ph = Math.min(H, 0.6 + 0.25 + Math.max(...lhs) + 0.3);
    panels.forEach(([t, items], i) => {
      const x = X0 + i * (pw + 0.3);
      box(s, x, Y0, pw, ph, HX.white, { line: HX.line });
      box(s, x, Y0, pw, 0.6, HX.deep, { r: 0.08 });
      T(s, t, { x: x + 0.2, y: Y0, w: pw - 0.4, h: 0.6, fontSize: 15, bold: true, color: HX.white, valign: "middle" });
      bullets(s, items, x + 0.2, Y0 + 0.85, pw - 0.4, { fontSize: 15, space: 7, numbered: true });
    });
    const rx = X0 + 2 * pw + 0.7, rw = X1 - rx;
    heading(s, "What happens when you save", rx, Y0, rw);
    bullets(s, [["Routines keep history. ", "A new routine never erases the old one. If dates overlap, the sidebar asks to end the previous routine when the new one starts."],
      ["Approved leave ", "is written to Leaves and compiled into the Presence Log at once; your projects' coverage is re-checked."],
      ["Requested leave ", "does not change your presence, but raises an orange coverage warning if it would leave a project uncovered."],
      ["Only active locations ", "are offered; retired ones stay valid in history."],
      ["Dates ", "use your computer's local date, so “today” is right after midnight."]], rx, Y0 + 0.5, rw, { fontSize: 14, space: 8 });
  }
  {
    const s = content("Weekly Presence: the team's week at a glance");
    const P = [HX.tPurple, HX.purple], G = [HX.tGreen, HX.green], B = [HX.tBlue, HX.blue], R = [HX.tRed, HX.red], WE = [HX.grey, HX.white];
    const cell = (t, st) => ({ text: t, options: { fill: { color: st[0] }, color: st[1], bold: true, align: "center" } });
    const hdr = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true, color: HX.white, fill: { color: HX.deep }, align: "center" }, o) });
    const day = (t, we) => ({ text: t, options: { color: we ? HX.white : HX.ink, align: "center", fill: { color: we ? HX.grey : HX.white } } });
    const E = (n) => Array.from({ length: n }, () => cell("", WE));
    const rows = [
      [hdr("Week", { rowspan: 2, valign: "middle" }), hdr("Date", { rowspan: 2, valign: "middle" }), hdr("Alice", { colspan: 2 }), hdr("Ben", { colspan: 2 }), hdr("Chloé", { colspan: 2 })],
      [hdr("AM"), hdr("PM"), hdr("AM"), hdr("PM"), hdr("AM"), hdr("PM")],
      [{ text: "W42", options: { rowspan: 7, bold: true, valign: "middle", align: "center", fill: { color: HX.white }, color: HX.ink } }, day("Mon 12"), cell("OneRoof", P), cell("OneRoof", P), cell("Remote", G), cell("Remote", G), cell("C04", B), cell("C04", B)],
      [day("Tue 13"), cell("OneRoof", P), cell("Remote", G), cell("OneRoof", P), cell("OneRoof", P), cell("C04", B), cell("Remote", G)],
      [day("Wed 14"), cell("Vacation", R), cell("Vacation", R), cell("OneRoof", P), cell("OneRoof", P), cell("C04", B), cell("C04", B)],
      [day("Thu 15"), cell("Vacation", R), cell("Vacation", R), cell("Remote", G), cell("JLL", B), cell("OneRoof", P), cell("OneRoof", P)],
      [day("Fri 16"), cell("Remote", G), cell("Remote", G), cell("Remote", G), cell("Remote", G), cell("Remote", G), cell("Sick Leave", R)],
      [day("Sat 17", true)].concat(E(6)), [day("Sun 18", true)].concat(E(6)),
    ];
    const tw = 7.1; const rh = (H - 0.4) / 9;
    s.addTable(rows, { x: X0, y: Y0, w: tw, colW: [0.75, 0.95, 0.9, 0.9, 0.9, 0.9, 0.9, 0.9], rowH: Array(9).fill(rh), fontSize: 13, fontFace: "Calibri", border: { type: "solid", color: HX.white, pt: 1.5 }, valign: "middle", margin: 0.03, objectName: "weekly-presence-mock" });
    T(s, "Illustration with sample names", { x: X0, y: Y1 - 0.3, w: tw, h: 0.3, fontSize: 12, color: HX.soft, italic: true });
    const rx = X0 + tw + 0.45;
    bullets(s, [["One row per day, ", "grouped by ISO week; each Member has an AM and a PM column."], ["Type to change a half-day. ", "It is saved to the Presence Log straight away; a toast confirms it."], ["Dropdowns ", "appear from the current week on, with active locations only. Past weeks are history."], ["Weekends and public holidays ", "are greyed; a holiday's name shows as a note on its date."], ["Departed Members' ", "columns are hidden, not deleted; older weeks are folded away."], ["Future days ", "stay empty until the nightly job fills them on the day, unless someone types them in."]], rx, Y0, X1 - rx, { fontSize: 15, space: 8 });
  }
  {
    const s = content("Adding people and changing contracts");
    const flow = [["MdPersonAdd", "Start from a name", "Menu Members & Contracts, or (admins) type it in an empty Members row"], ["MdViewSidebar", "Sidebar opens", "Pre-filled when the name was typed in Members"], ["MdAssignment", "First contract", "Role, department, start, type, FTE %; optional email, phone, birthday"], ["MdAutorenew", "Row completes", "Role, Department and Status are filled from the contract"]];
    const c4 = cols(4, 0.5);
    const bh = Math.max(...flow.map((f, i) => textH(f[2], 14, c4[i].w - 2 * PAD)));
    const fh = PAD + 0.6 + 0.2 + 0.4 + bh + 0.1;
    for (let i = 0; i < 4; i++) {
      const [ic, t, b] = flow[i]; const { x, w } = c4[i];
      box(s, x, Y0, w, fh, HX.lav);
      await iconCircle(s, ic, x + PAD, Y0 + PAD, 0.6, HX.purple, HX.white);
      T(s, t, { x: x + PAD, y: Y0 + PAD + 0.8, w: w - 2 * PAD, h: 0.38, fontSize: 17, bold: true, fontFace: "Arial" });
      T(s, b, { x: x + PAD, y: Y0 + PAD + 1.2, w: w - 2 * PAD, h: bh + 0.1, fontSize: 14, color: HX.body });
      if (i < 3) hLine(s, x + w + 0.07, Y0 + fh / 2, x + w + 0.43, HX.purple);
    }
    let y = Y0 + fh + 0.35;
    y += heading(s, "For someone already in the team, the sidebar asks for one action at a time", X0, y, W);
    const acts = [{ title: "Contact details", body: "Accenture email, phone and birthday. Birthday is day and month only: type 14/03, March 14 or a date; it is stored as “14 Mar”." }, { title: "New contract", body: "Asks when the current contract ends (default: the day before the new start). Overlaps are refused; a gap is allowed and reported." }, { title: "End a contract", body: "Sets the end date of one contract. After the last day the person is Departed: hidden in Weekly Presence, excluded from emails." }];
    await cardRow(s, acts, y, 3, { base: { fill: HX.white, line: HX.line, titleColor: HX.purple, titleSize: 16, bodySize: 14 }, fixedH: Y1 - y });
  }
  {
    const s = content("Projects stay covered, day by day");
    const lw = 7.0;
    const th = table(s, [["Column", "Filled by", "Notes"], ["ID", "System", "PRJ-XXXXX, created on first edit"], ["Project, Status, Start, Target date", "You", "Status: Planning, Active, On Hold, Completed"], ["Lead", "You", "People chip; changing it asks for confirmation"], ["Backup", "You", "Permanent backup; a toast asks for one if empty"], ["In charge today", "System", "Lead, Name (interim), Name (backup) or Nobody available"]], { x: X0, y: Y0, colW: [2.55, 1.05, lw - 3.6], fontSize: 14 });
    const rx = X0 + lw + 0.35, rw = X1 - rx;
    const c1 = { icon: "MdError", iconColor: HX.red, title: "Red row: conflict", titleColor: HX.red, fill: HX.tRed, body: "Approved absences leave a working day uncovered. Email to the lead and a manager, plus an “Organize interim backup” task." };
    const c2 = { icon: "MdWarning", iconColor: HX.orange, title: "Orange row: warning", titleColor: HX.orange, fill: HX.tOrange, body: "Same, but a leave is only Requested. Warning email, no task." };
    const h1 = cardGeom(Object.assign({ w: rw }, c1)).need, h2 = cardGeom(Object.assign({ w: rw }, c2)).need;
    const hh = Math.max(th, h1 + h2 + 0.2);
    await card(s, Object.assign({ x: rx, y: Y0, w: rw, h: hh - h2 - 0.2 }, c1));
    await card(s, Object.assign({ x: rx, y: Y0 + hh - h2, w: rw, h: h2 }, c2));
    let y = Y0 + hh + 0.35;
    await card(s, { x: X0, y, w: W, icon: "MdSupervisorAccount", title: "Who is “a manager”?", body: "The first Member whose role contains Manager, Head or Director; if there is none, the account running the check.", h: Y1 - y });
  }
  {
    const s = content("Fixing a gap with Interim Backups");
    const intro = "Menu → Project Management → Interim Backups. The sidebar works on one project at a time and re-checks coverage when you save, without sending emails.";
    const ih = textH(intro, 16, W); T(s, intro, { x: X0, y: Y0, w: W, h: ih + 0.1, fontSize: 16, color: HX.body });
    const y = Y0 + ih + 0.4;
    await cardRow(s, [
      { icon: "MdFolderOpen", title: "1 · Pick the project", body: ["See its lead and permanent backup.", "See the uncovered periods, in red.", "See current and planned interim backups."] },
      { icon: "MdPersonAddAlt1", title: "2 · Add a backup", body: ["Click a gap: its dates fill in.", "Choose a person and add notes.", "Save. The lead is refused as their own interim; a leave of the chosen person is flagged."] },
      { icon: "MdEventBusy", title: "3 · Cancel, don't delete", body: ["Cancelling keeps the row with status Cancelled.", "History stays readable in the hidden Backups tab.", "Several consecutive interim backups can cover a long absence."] },
    ], y, 3, { base: { space: 8 } });
  }
  {
    const s = content("Tasks and Events connect to Google Tasks and Calendar");
    const t = { icon: "MdTaskAlt", title: "Tasks", body: [["Type a title: ", "the row gets an ID (TSK-YYMMDD-XXXX) and status Todo."], ["Optional: ", "Project ID, Assigned to, Description, Due date."], ["Status: ", "Todo, In Progress, Done, Blocked."], ["Google Tasks: ", "pushed to “Project Management Tasks” of the account running the sync, for tasks assigned to that account. Completing it there marks the row Done."], ["Done is final: ", "changing it offers to hide the row and create a follow-up task with a new ID."], ["System tasks: ", "onboarding for new Members; “Organize interim backup” for coverage conflicts."]] };
    const e = { icon: "MdEvent", title: "Events", body: [["Type a title and a date ", "(a date-time is fine): the row gets an ID (EVT-YYMMDD-XXXX) and a Google Calendar event."], ["Duration ", "in minutes, “45m”, “1h” or “1.5 hours” makes a timed slot; empty makes an all-day event."], ["Organizer: ", "the event goes to their calendar when the script can reach it, otherwise to the owner's; a team calendar can also be set."], ["Edits ", "to title, date, duration, location or organizer update the event."], ["No invitations ", "are sent automatically: the organizer invites guests in Calendar."]] };
    await cardRow(s, [t, e], Y0, 2, { base: { space: 6, bodySize: 14 } });
  }
  {
    const s = content("Check Workspace lists everything that needs attention");
    const items = [["MdPersonOff", "Members without a contract", "Members & Contracts"], ["MdShield", "Open projects without a backup", "Projects tab"], ["MdReportProblem", "Coverage conflicts and warnings", "Interim Backups"], ["MdLogout", "People leaving within 30 days who lead, back up or own tasks", "Members & Contracts"], ["MdHelpOutline", "Members not logged today", "Declare Presence"], ["MdEventBusy", "Undeclared days this week", "Declare Presence"], ["MdContentCopy", "Duplicate Presence Log rows", "Remove duplicates"]];
    const lw = 8.4, rh = (H - 6 * 0.1) / 7;
    for (let i = 0; i < items.length; i++) {
      const [ic, t, fix] = items[i]; const y = Y0 + i * (rh + 0.1);
      box(s, X0, y, lw, rh, i % 2 ? HX.white : HX.lav, { line: i % 2 ? HX.line : null });
      s.addImage({ data: await icon(ic, HX.purple), x: X0 + 0.2, y: y + (rh - 0.36) / 2, w: 0.36, h: 0.36, objectName: nid("icon") });
      T(s, t, { x: X0 + 0.75, y, w: 5.0, h: rh, fontSize: 15, bold: true, valign: "middle" });
      T(s, "→ " + fix, { x: X0 + 5.85, y, w: lw - 6.0, h: rh, fontSize: 14, color: HX.purple, valign: "middle" });
    }
    const rx = X0 + lw + 0.3, rw = X1 - rx;
    box(s, rx, Y0, rw, H, HX.deep);
    T(s, "How to use it", { x: rx + PAD, y: Y0 + PAD, w: rw - 2 * PAD, h: 0.4, fontSize: 18, bold: true, color: HX.white, fontFace: "Arial" });
    bullets(s, ["Open it from the menu: the dialog appears at once with “Checking…”, the report follows in about 25 s.", "Each section links to the place that fixes it.", "It only reports: no email or task is created.", "Click “Check again” after fixing."], rx + PAD, Y0 + PAD + 0.6, rw - 2 * PAD, { fontSize: 15, color: HX.white, space: 10 });
  }
  {
    const s = content("Emails the system sends");
    const th = table(s, [["Email", "When", "To", "Content"], ["Presence reminder", "Monday 8:00", "Each active Member with an undeclared day this week", "The days to declare and a link to Weekly Presence"], ["Birthday reminder", "9:00 on the working day before", "Every active Member with an email, except the person celebrated", "Coming birthdays, an “Add to my Google Calendar” link and a link to subscribe to Team Birthdays"], ["Coverage conflict", "When a check (edit, leave declaration, nightly) finds a new uncovered period", "Project lead and a manager", "Uncovered periods and the reason; an “Organize interim backup” task is also created"], ["Coverage warning", "Same, when a Requested leave is involved", "Project lead and a manager", "The periods at risk; no task"]], { x: X0, y: Y0, colW: [2.2, 2.9, 3.3, W - 8.4], fontSize: 14 });
    const y = Y1 - 1.35;
    await card(s, { x: X0, y, w: W, h: 1.35, icon: "MdInfoOutline", title: "Who sends them", bodySize: 14, body: "All emails and calendar changes come from the Google account that owns the triggers (the sheet owner). Members without an email address are logged, not emailed. A Member is “active” while a contract covers today." });
  }
  {
    const s = content("Good habits");
    const d = { icon: "MdCheckCircle", iconColor: HX.green, title: "Do", titleColor: HX.green, fill: HX.tGreen, body: ["Set your routine once; then only type the exceptions.", "Declare a leave as Requested as soon as you know, then switch it to Approved.", "Give every project a permanent backup.", "End or cancel things instead of deleting rows, so history stays readable.", "Use a sidebar when unsure: it checks dates and overlaps before writing.", "Run Check Workspace before your weekly meeting."] };
    const a = { icon: "MdDoNotDisturbOn", iconColor: HX.red, title: "Avoid", titleColor: HX.red, fill: HX.tRed, body: ["Typing over system columns (IDs, Role, Department, Status, Today, In charge today): the next sync rewrites them.", "Typing a day in Weekly Presence for a period you will be on leave: hand entries win over leaves.", "Adding, renaming or deleting columns: the schema is fixed; changes go through the maintainer.", "Deleting rows in hidden tabs: end, cancel or mark them instead.", "Using the mobile app: menus and sidebars are desktop only."] };
    await cardRow(s, [d, a], Y0, 2, { base: { space: 9, bodySize: 16 } });
  }

  // ================= SECTION 3
  section("3 · Setting it up"); kicker = "3 · Setting it up";
  await divider(3, "Setting it up", "For administrators. One Library holds all the code; each team sheet carries a thin wrapper that points to it.", ["How it is distributed", "Before you start", "Setup in seven steps", "What Initialize / Upgrade does", "Configuring for your team", "Changing settings", "Triggers and the owner account", "Upgrading and rolling back"]);
  {
    const s = content("One Library, many team sheets");
    const lw = 4.0, lh = H;
    box(s, X0, Y0, lw, lh, HX.deep);
    await iconCircle(s, "MdLibraryBooks", X0 + 0.3, Y0 + 0.3, 0.7, HX.purple, HX.white);
    T(s, "WorkspacePM Library", { x: X0 + 0.3, y: Y0 + 1.2, w: lw - 0.6, h: 0.45, fontSize: 20, bold: true, color: HX.white, fontFace: "Arial" });
    T(s, "One Apps Script project with all the logic: schema, sync, presence, coverage, Google Tasks, Calendar, emails and sidebars.", { x: X0 + 0.3, y: Y0 + 1.75, w: lw - 0.6, h: 1.8, fontSize: 16, color: HX.lilac });
    T(s, "Source: src/ in the repository, pushed with clasp", { x: X0 + 0.3, y: Y0 + lh - 0.7, w: lw - 0.6, h: 0.5, fontSize: 13, color: "B9A6E0" });
    const tx = X0 + lw + 1.0, tw = 3.0, th = (H - 2 * 0.3) / 3;
    for (let i = 0; i < 3; i++) {
      const y = Y0 + i * (th + 0.3);
      box(s, tx, y, tw, th, HX.lav);
      s.addImage({ data: await icon("MdGridOn", HX.green), x: tx + 0.25, y: y + 0.3, w: 0.45, h: 0.45, objectName: nid("icon") });
      T(s, "Team sheet " + "ABC"[i], { x: tx + 0.85, y: y + 0.3, w: tw - 1.0, h: 0.45, fontSize: 17, bold: true, valign: "middle" });
      T(s, "Tabs + Code.gs wrapper", { x: tx + 0.25, y: y + 0.95, w: tw - 0.5, h: 0.35, fontSize: 15, color: HX.grey });
      hLine(s, tx - 0.05, y + th / 2, X0 + lw + 0.08, HX.purple);
    }
    const rx = tx + tw + 0.45, rw = X1 - rx;
    heading(s, "Why it is built this way", rx, Y0, rw);
    bullets(s, [["Fix once: ", "a Library change reaches every sheet."], ["Small wrapper: ", "menu, triggers and sidebar forwarders only."], ["Version control: ", "a sheet uses a numbered version, or development mode (HEAD) while testing."], ["Additive upgrades: ", "rows are never deleted by an upgrade."]], rx, Y0 + 0.5, rw, { fontSize: 15, space: 10 });
  }
  {
    const s = content("Before you start");
    const r = [["MdAccountCircle", "An owner account", "The account that owns the sheet also owns the triggers. Emails, calendar updates and nightly runs happen as this account."], ["MdVpnKey", "Library access", "The Library's Script ID and at least view access to the Library project."], ["MdExtension", "Three advanced services", "Google Sheets API v4 (native tables), Google Tasks API v1, Google Calendar API v3."], ["MdPeople", "Google accounts", "People chips identify Members. Addresses without a Google account can be stored but not chipped."], ["MdComputer", "A desktop browser", "Menus, sidebars and dialogs do not exist in the Sheets mobile apps."], ["MdSchedule", "Time zone", "The project runs in Europe/Paris: triggers fire at 2:00, 8:00 and 9:00 Paris time."]].map(([icon, title, body]) => ({ icon, title, body }));
    const rh = Math.max(rowsNeed(r, 3), (H - GAP) / 2);
    await cardRow(s, r.slice(0, 3), Y0, 3, { minH: rh });
    await cardRow(s, r.slice(3), Y0 + rh + GAP, 3, { minH: rh });
  }
  {
    const s = content("Setup in seven steps");
    const steps = [["Open the team's Google Sheet", "Create a new sheet or use an existing one, signed in as the owner account."], ["Open Extensions → Apps Script", "Rename the project, e.g. “Team Alpha – WorkspacePM”."], ["Add the three services", "Services + → Google Sheets API, Tasks API, Calendar API, with their default identifiers."], ["Add the Library", "Libraries + → paste the Script ID → Look up → Identifier WorkspacePM (exact) → pick a version."], ["Paste the wrapper", "Replace Code.gs with consumer/Code.js from the repository, then save."], ["Initialize", "Reload the sheet → Project Management → Initialize / Upgrade Spreadsheet → authorise when Google asks."], ["Configure", "Check Locations and Holidays, then add the team in Members & Contracts."]];
    const c2 = cols(2); const rh = (H - 3 * 0.2) / 4;
    steps.forEach(([t, b], i) => {
      const col = i < 4 ? 0 : 1, row = i < 4 ? i : i - 4; const { x, w } = c2[col]; const y = Y0 + row * (rh + 0.2);
      box(s, x, y, w, rh, row % 2 ? HX.white : HX.lav, { line: row % 2 ? HX.line : null });
      s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + (rh - 0.6) / 2, w: 0.6, h: 0.6, fill: { color: HX.purple }, line: { type: "none" }, objectName: nid("num") });
      T(s, String(i + 1), { x: x + 0.25, y: y + (rh - 0.6) / 2, w: 0.6, h: 0.6, fontSize: 20, bold: true, color: HX.white, align: "center", valign: "middle" });
      const bh = textH(b, 14, w - 1.35); const blockH = 0.35 + 0.05 + bh; const ty = y + (rh - blockH) / 2;
      T(s, t, { x: x + 1.1, y: ty, w: w - 1.3, h: 0.35, fontSize: 16, bold: true });
      T(s, b, { x: x + 1.1, y: ty + 0.4, w: w - 1.3, h: bh + 0.05, fontSize: 14, color: HX.body });
    });
    const { x, w } = c2[1]; const y = Y0 + 3 * (rh + 0.2);
    box(s, x, y, w, rh, HX.deep);
    T(s, "The wrapper's menu, onEdit/onOpen forwarders and one forwarder per sidebar and trigger function are required: triggers and sidebars call the sheet's own script, not the Library.", { x: x + 0.3, y, w: w - 0.6, h: rh, fontSize: 14, color: HX.white, valign: "middle" });
  }
  {
    const s = content("What Initialize / Upgrade does");
    const st = [["Rename tabs", "Old technical tab names become readable ones"], ["Build tables", "Native tables, column types, dropdowns, fonts, protections, required-cell rules"], ["Restore presence", "One-off repair of an old merged table, then archive it"], ["Install triggers", "Edit, nightly, Monday and birthday triggers"], ["Full sync", "Runs the nightly job once"], ["Register tab IDs", "So tabs can be renamed safely"], ["Apply layout", "Show five tabs, hide the rest, delete an empty Sheet1"]];
    const c7 = cols(7, 0.18);
    const bh = Math.max(...st.map((v, i) => textH(v[1], 14, c7[i].w - 0.3)));
    const tH = Math.max(...st.map((v, i) => textH(v[0], 15, c7[i].w - 0.3, "Calibri", true)));
    const hh = 0.2 + 0.5 + 0.1 + tH + 0.12 + bh + 0.25;
    st.forEach(([t, b], i) => {
      const { x, w } = c7[i];
      box(s, x, Y0, w, hh, i % 2 ? HX.lav : HX.tPurple);
      T(s, String(i + 1), { x: x + 0.15, y: Y0 + 0.2, w: 0.5, h: 0.5, fontSize: 26, bold: true, color: HX.purple, fontFace: "Arial" });
      T(s, t, { x: x + 0.15, y: Y0 + 0.8, w: w - 0.3, h: tH + 0.05, fontSize: 15, bold: true });
      T(s, b, { x: x + 0.15, y: Y0 + 0.92 + tH, w: w - 0.3, h: bh + 0.05, fontSize: 14, color: HX.body });
    });
    const sy = Y0 + hh + 0.4; const c3 = cols(3);
    [["~2 min", "on the test sheet"], ["Re-runnable", "after a new version, a schema change, or when a tab looks broken"], ["0 rows", "deleted by an upgrade"]].forEach(([n, l], i) => {
      T(s, n, { x: c3[i].x, y: sy, w: c3[i].w, h: 0.75, fontSize: 40, bold: true, color: HX.purple, fontFace: "Arial" });
      T(s, l, { x: c3[i].x, y: sy + 0.8, w: c3[i].w, h: 0.35, fontSize: 15, color: HX.grey });
    });
  }
  {
    const s = content("Configuring the sheet for your team");
    const cs = [{ icon: "MdPlace", title: "Locations", body: ["Menu → Edit Locations. Columns: Location, Type, Active.", "Types: Office, Client site, Remote. Defaults: OneRoof, Remote, JLL, C04, C60.", "The first active Office is the default for days with no routine.", "Untick Active to retire a place: past days stay valid; a toast lists who still uses it."] },
      { icon: "MdFlag", title: "Holidays", body: ["Menu → Edit Holidays. Columns: Date, Holiday.", "The 11 French public holidays of this year and next are added once a year.", "Add bridges and office closures by hand.", "A deleted line stays deleted, e.g. remove Whit Monday if the team works it."] },
      { icon: "MdTune", title: "Settings", body: [["GOOGLE_CALENDAR_ID: ", "a team calendar for events (default: primary)."], ["GOOGLE_TASKS_LIST_ID: ", "set automatically to “Project Management Tasks”."], ["ADMIN_EMAILS: ", "who may edit locked columns, besides the owner."], ["Where: ", "document properties, set from the Apps Script editor (next slide)."]] }];
    await cardRow(s, cs, Y0, 3, { base: { space: 8 } });
  }
  {
    const s = content("Changing settings from the Apps Script editor");
    const intro = "Document properties cannot be edited on the Apps Script settings page. Add this function to Code.gs, choose it in the toolbar, click Run once, then delete it.";
    const ih = textH(intro, 16, W); T(s, intro, { x: X0, y: Y0, w: W, h: ih + 0.1, fontSize: 16, color: HX.body });
    const cy = Y0 + ih + 0.35, cw = 7.6, chh = 3.1;
    box(s, X0, cy, cw, chh, HX.ink);
    T(s, [
      { text: "function configureWorkspace() {", options: { breakLine: true } },
      { text: "  // Team calendar for Events (optional)", options: { color: "9A8FB5", breakLine: true } },
      { text: "  WorkspacePM.setCalendarId(\"team@group.calendar.google.com\");", options: { breakLine: true } },
      { text: "  // Who besides the owner may edit locked columns", options: { color: "9A8FB5", breakLine: true } },
      { text: "  WorkspacePM.setAuthorizedAdmins([\"lead@example.com\"]);", options: { breakLine: true } },
      { text: "  // Reuse an existing Google Tasks list (optional)", options: { color: "9A8FB5", breakLine: true } },
      { text: "  WorkspacePM.setTasksListId(\"<task list id>\");", options: { breakLine: true } },
      { text: "}", options: {} }], { x: X0 + 0.3, y: cy + 0.3, w: cw - 0.6, h: chh - 0.6, fontSize: 13, color: "E8E3F5", fontFace: "Courier New", lineSpacingMultiple: 1.15 });
    const rx = X0 + cw + 0.4, rw = X1 - rx;
    heading(s, "What each setting changes", rx, cy, rw);
    bullets(s, [["Calendar: ", "events are also written to this calendar; the owner must be able to edit it."], ["Admins: ", "applied when protections are next rebuilt (nightly, or Initialize / Upgrade)."], ["Tasks list: ", "only needed to reuse a list; otherwise the Library creates one."]], rx, cy + 0.5, rw, { fontSize: 15, space: 10 });
    note(s, "Called this way, the Library runs in the context of this sheet, so the values are stored for this sheet only.", cy + chh + 0.3);
  }
  {
    const s = content("Triggers and the owner account");
    const th = table(s, [["Trigger", "Function", "Schedule", "Does"], ["Edit", "installedTaskEditHandler", "Every edit", "Handles the edit as the owner: Tasks, Calendar, sync, coverage"], ["Nightly", "nightlySyncJob", "Daily 2:00", "Full reconciliation: ten steps, five-minute budget"], ["Presence reminder", "mondayPresenceReminder", "Monday 8:00", "Emails Members with undeclared days"], ["Birthdays", "birthdayReminderJob", "Daily 9:00", "Birthday email and calendar link"]], { x: X0, y: Y0, colW: [2.2, 3.1, 1.9, W - 7.2], fontSize: 15, minRowH: 0.48 });
    const y = Y0 + th + 0.4;
    await cardRow(s, [{ icon: "MdAccountCircle", title: "Runs as the owner", body: "Initialize / Upgrade installs all four for the account that runs it. Here that is the owner's personal Gmail: emails come from it and the Team Birthdays calendar belongs to it." }, { icon: "MdSync", title: "One run per edit", body: "Once the edit trigger exists, the simple onEdit returns at once (flag EDIT_TRIGGER_ACTIVE), so each edit is processed once. Re-run Initialize / Upgrade if triggers were deleted." }], y, 2, { minH: Y1 - y });
  }
  {
    const s = content("Upgrading, deploying and rolling back");
    await cardRow(s, [{ icon: "MdSystemUpdateAlt", title: "Upgrade a team sheet", body: ["Apps Script → Libraries → WorkspacePM → choose the new version → Save.", "Back in the sheet: Initialize / Upgrade.", "Development mode (HEAD) gets every push at once: fine for testing, not for production."] },
      { icon: "MdCloudUpload", title: "Deploy (maintainer)", body: ["Run the tests: node --test test/*.js", "Commit and tag deploy-YYYYMMDD-HHMM.", "clasp push in src/ (Library), and in consumer/ when the wrapper changed.", "Verify by pulling into .scratch/live/ and diffing."] },
      { icon: "MdHistory", title: "Roll back", body: ["Check out the previous deploy-* tag.", "Push src/ (and consumer/) again with clasp.", "Pinned sheets: select the previous numbered version instead.", "Data is not rolled back: upgrades are additive, so older code still reads it."] }], Y0, 3, { base: { space: 9 } });
  }

  // ================= SECTION 4
  section("4 · Under the cover"); kicker = "4 · Under the cover";
  await divider(4, "Under the cover", "For maintainers and curious admins: how the pieces fit, how data flows from what people type to what the sheet derives, and what keeps it consistent.", ["Architecture", "Library modules", "Data model", "The edit pipeline", "Presence compilation", "Coverage algorithm", "The nightly job", "Table registry", "Sync Lock", "KPI Data for Looker"]);
  {
    const s = content("Architecture: three layers");
    const lw = 8.3;
    const layers = [["Team sheet", HX.tGreen, HX.green, ["Visible tabs", "Hidden tabs", "Sidebars & dialogs", "Code.gs wrapper"]], ["WorkspacePM Library", HX.tPurple, HX.purple, ["Edit router", "Sync engines", "Presence & coverage", "Integrations"]], ["Google services", HX.tBlue, HX.blue, ["Sheets API v4", "Tasks API", "Calendar", "Gmail, Lock, Properties"]]];
    const lh = (H - 2 * 0.45) / 3;
    layers.forEach(([t, f, c, bx], i) => {
      const y = Y0 + i * (lh + 0.45);
      box(s, X0, y, lw, lh, f);
      T(s, t, { x: X0 + 0.25, y, w: 1.75, h: lh, fontSize: 17, bold: true, color: c, valign: "middle", fontFace: "Arial" });
      const bc = cols(4, 0.15, X0 + 2.1, lw - 2.3);
      bx.forEach((b, j) => { box(s, bc[j].x, y + 0.25, bc[j].w, lh - 0.5, HX.white, { line: HX.line }); T(s, b, { x: bc[j].x + 0.05, y: y + 0.25, w: bc[j].w - 0.1, h: lh - 0.5, fontSize: 14, bold: true, align: "center", valign: "middle" }); });
      if (i < 2) { vLine(s, X0 + 3.6, y + lh + 0.03, y + lh + 0.42, HX.grey); vLine(s, X0 + 6.3, y + lh + 0.42, y + lh + 0.03, HX.grey); }
    });
    const rx = X0 + lw + 0.4, rw = X1 - rx;
    heading(s, "How calls travel", rx, Y0, rw);
    bullets(s, [["Menus and triggers ", "call functions in Code.gs, which forward to WorkspacePM."], ["Sidebars ", "use google.script.run, which reaches the sheet's script; a forwarder passes the call on."], ["The Library ", "reads and writes the tabs, using the Sheets API for native tables."], ["Integrations ", "use Google Tasks, Calendar and Gmail as the trigger owner."], ["State ", "lives in tabs and document properties."]], rx, Y0 + 0.5, rw, { fontSize: 14, space: 8 });
  }
  {
    const s = content("Library modules and what each owns");
    table(s, [["Area", "Files", "Responsibility"], ["Entry & orchestration", "Code.js", "Menu actions, setupOrUpgrade, triggers, nightly job, edit router, Sync Lock"], ["Structure", "Schema.js, Tables.js, Init.js, Typography.js, Layout.js", "Column spec, table registry, table creation and migration, fonts, visible and hidden tabs"], ["Integrity", "ValidationManager.js, ProtectionManager.js", "Required fields and highlights; protections, admin list, warning snooze"], ["People & presence", "SyncEngine.js, MembersService.js, Sidebar.js, PresenceEngine.js, ViewBuilder.js, Locations.js, Holidays.js", "Members from contracts, sidebars, Presence Log compile, Weekly Presence, locations, holidays"], ["Projects", "ProjectBackupEngine.js, BackupService.js", "Coverage per day, In charge today, alerts; interim backup sidebar"], ["Integrations & reports", "TaskSyncService.js, CalendarSyncService.js, Birthdays.js, KpiEngine.js, WorkspaceCheck.js", "Google Tasks, Calendar, birthdays, KPI Data, Check Workspace"], ["User interface", "BaselineSidebar, MembersSidebar, BackupSidebar, WorkspaceCheckDialog (.html)", "The four HTML panels"]], { x: X0, y: Y0, colW: [2.3, 5.2, W - 7.5], fontSize: 13 });
  }
  {
    const s = content("Data model: what people write, what the system derives");
    const src = [["Contracts", "member, role, department, dates, type, FTE"], ["Routines", "member, weekday, AM/PM, dates"], ["Leaves", "member, dates, half-days, type, status"], ["Projects, Backups", "lead, backup, interim cover"], ["Locations, Holidays", "places and non-working days"], ["Tasks, Events", "work items and gatherings"]];
    const der = [["Members", "Role, Department, Status, Today"], ["Presence Log", "one row per Member per working day"], ["Weekly Presence", "grid view of the Presence Log"], ["In charge today", "coverage result per project"], ["KPI Data", "Member × day × half-day facts"], ["Google Tasks, Calendar", "external copies with stored IDs"]];
    const sw = 4.75, mx = X0 + sw + 0.45, mw = W - 2 * sw - 0.9, dx = mx + mw + 0.45;
    heading(s, "Written by people", X0, Y0, sw, HX.green); heading(s, "Derived by the Library", dx, Y0, sw, HX.purple);
    const top = Y0 + 0.5, rh = (Y1 - top - 5 * 0.15) / 6;
    const draw = (arr, x, f, c) => arr.forEach(([t, b], i) => { const y = top + i * (rh + 0.15); box(s, x, y, sw, rh, f); T(s, [{ text: t, options: { bold: true, color: c, breakLine: true } }, { text: b, options: { color: HX.body, fontSize: 13 } }], { x: x + 0.2, y, w: sw - 0.4, h: rh, fontSize: 15, valign: "middle" }); });
    draw(src, X0, HX.tGreen, HX.green); draw(der, dx, HX.tPurple, HX.purple);
    box(s, mx, top, mw, Y1 - top, HX.deep);
    T(s, "Library", { x: mx, y: top + 0.25, w: mw, h: 0.4, fontSize: 17, bold: true, color: HX.white, align: "center" });
    T(s, "joins rows on the Member's people chip and the project ID, then syncs, compiles, checks coverage and renders", { x: mx + 0.15, y: top + 0.8, w: mw - 0.3, h: 2.5, fontSize: 14, color: HX.lilac, align: "center" });
    for (let i = 0; i < 6; i++) { const y = top + i * (rh + 0.15) + rh / 2; hLine(s, X0 + sw + 0.04, y, mx - 0.04, HX.green); hLine(s, mx + mw + 0.04, y, dx - 0.04, HX.purple); }
  }
  {
    const s = content("The edit pipeline: what happens after you type");
    const st = [["Edit", "A cell changes in any tab"], ["Single run", "Simple onEdit steps aside when the installable trigger is active"], ["Sync Lock", "Waits up to 10 s; during the nightly job a toast says “Sync in progress”"], ["Route by tab", "The tab is found by stored ID, never by name"], ["Table handlers", "Weekly Presence → Presence Log; Tasks and Events → IDs, Google sync"], ["Validate", "Missing required fields named in a toast"], ["Protection check", "Warning dialog on protected cells, snoozable for 5 min"], ["Follow-ups", "Light Members sync; Locations → dropdowns; Projects → backup prompt"], ["Coverage", "Edits to Projects, Leaves or Backups re-check coverage"]];
    const c3 = cols(3); const rh = (H - 0.5 - 2 * 0.25) / 3;
    st.forEach(([t, b], i) => {
      const { x, w } = c3[i % 3]; const y = Y0 + Math.floor(i / 3) * (rh + 0.25);
      box(s, x, y, w, rh, Math.floor(i / 3) === 1 ? HX.tPurple : HX.lav);
      T(s, String(i + 1), { x: x + 0.2, y: y + 0.2, w: 0.5, h: 0.6, fontSize: 30, bold: true, color: HX.purple, fontFace: "Arial" });
      T(s, t, { x: x + 0.85, y: y + 0.22, w: w - 1.05, h: 0.35, fontSize: 17, bold: true });
      T(s, b, { x: x + 0.85, y: y + 0.6, w: w - 1.05, h: rh - 0.7, fontSize: 13, color: HX.body });
    });
    note(s, "Each step is isolated: an error is logged and the next step still runs. Edits slower than 3 s write an [edit timing] line to the execution log.", Y1 - 0.3);
  }
  {
    const s = content("Presence compilation: from rules to a daily record");
    const lw = 6.6;
    heading(s, "For each active Member and working day up to today", X0, Y0, lw);
    const lad = [["1", "A row already exists", "Kept as is: hand entries and earlier compiles are never overwritten.", HX.green, HX.tGreen], ["2", "An approved leave covers the day", "Leave type written for AM, PM or both (half-day leaves split).", HX.red, HX.tRed], ["3", "A routine matches the weekday", "Its AM and PM locations are written.", HX.purple, HX.tPurple], ["4", "Nothing applies", "The first active Office location (OneRoof).", HX.grey, HX.tGrey]];
    const top = Y0 + 0.55, rh = (Y1 - top - 3 * 0.18) / 4;
    lad.forEach(([n, t, b, c, f], i) => {
      const y = top + i * (rh + 0.18), ind = i * 0.3;
      box(s, X0 + ind, y, lw - ind, rh, f);
      T(s, n, { x: X0 + ind + 0.25, y, w: 0.45, h: rh, fontSize: 28, bold: true, color: c, valign: "middle", fontFace: "Arial" });
      const bh = textH(b, 14, lw - ind - 1.05); const blk = 0.36 + bh; const ty = y + (rh - blk) / 2;
      T(s, t, { x: X0 + ind + 0.85, y: ty, w: lw - ind - 1.05, h: 0.34, fontSize: 16, bold: true, color: c });
      T(s, b, { x: X0 + ind + 0.85, y: ty + 0.38, w: lw - ind - 1.05, h: bh + 0.05, fontSize: 14, color: HX.body });
    });
    const rx = X0 + lw + 0.45, rw = X1 - rx;
    heading(s, "Boundaries", rx, Y0, rw);
    bullets(s, [["Up to today only. ", "The Presence Log is history; the future is planned through routines and leaves."], ["Weekends ", "are written only when a routine says the Member works or an approved leave covers them."], ["Public holidays ", "are skipped unless someone types them in."], ["Departed Members ", "are skipped; days outside every contract are not compiled."], ["Today column ", "shows the day's value, or Weekend, Public holiday or Not Logged."], ["When it runs: ", "nightly, after a routine or approved leave is saved, and from the migration command."]], rx, Y0 + 0.55, rw, { fontSize: 14, space: 8 });
  }
  {
    const s = content("Coverage: who is in charge, day by day");
    const qw = 3.7, aw = 2.3, ax = X0 + qw + 0.8, qh = 0.75, step = (H - qh) / 3;
    const chain = [["Is the lead present?", "Lead"], ["Is an interim backup active and present?", "Name (interim)"], ["Is the Project Backup present?", "Name (backup)"]];
    chain.forEach(([q, a], i) => {
      const y = Y0 + i * step;
      box(s, X0, y, qw, qh, HX.lav);
      T(s, q, { x: X0 + 0.2, y, w: qw - 0.4, h: qh, fontSize: 15, bold: true, valign: "middle" });
      hLine(s, X0 + qw + 0.04, y + qh / 2, ax - 0.04, HX.green);
      T(s, "yes", { x: X0 + qw + 0.1, y: y + qh / 2 - 0.3, w: 0.6, h: 0.25, fontSize: 12, bold: true, color: HX.green, align: "center" });
      box(s, ax, y, aw, qh, HX.tGreen, { r: 0.2 });
      T(s, a, { x: ax, y, w: aw, h: qh, fontSize: 15, bold: true, color: HX.green, align: "center", valign: "middle" });
      vLine(s, X0 + qw / 2, y + qh + 0.03, y + step - 0.03, HX.red);
      T(s, "no", { x: X0 + qw / 2 + 0.12, y: y + qh + (step - qh) / 2 - 0.13, w: 0.4, h: 0.25, fontSize: 12, bold: true, color: HX.red });
    });
    const ly = Y0 + 3 * step; box(s, X0, ly, qw, qh, HX.tRed);
    s.addImage({ data: await icon("MdWarning", HX.red), x: X0 + 0.2, y: ly + (qh - 0.36) / 2, w: 0.36, h: 0.36, objectName: nid("icon") });
    T(s, "Nobody available", { x: X0 + 0.7, y: ly, w: qw - 0.9, h: qh, fontSize: 16, bold: true, color: HX.red, valign: "middle" });
    const rx = ax + aw + 0.5, rw = X1 - rx;
    heading(s, "Rules", rx, Y0, rw);
    bullets(s, [["Days checked: ", "each working day from today to the last relevant leave, at most 366 days; weekends and public holidays skipped."], ["Absent means: ", "an approved or requested leave, outside every contract, or (today only) absent in the Today column."], ["Conflict ", "if only approved absences are involved; warning if a requested leave is."], ["Consecutive days ", "are grouped into periods across weekends."], ["Backups: ", "the Backup column is mirrored as the open row in Backups; changing it ends the old row yesterday."], ["Runs ", "on edits to Projects, Leaves and Backups, on leave declarations, compile and nightly."]], rx, Y0 + 0.5, rw, { fontSize: 14, space: 6 });
  }
  {
    const s = content("The nightly job, step by step");
    const st = [["Presence restore", "one-off repair", ""], ["Directory sync", "Members from Contracts, protections", "29 s"], ["Presence compile", "up to today", "19 s"], ["Weekly Presence", "rebuild the grid", ""], ["KPI Data", "rebuild the fact table", ""], ["Coverage", "emails and tasks for new gaps", ""], ["Google Tasks", "push new, pull completions", ""], ["Calendar", "create and update events", ""], ["Holidays", "add next year's once", ""], ["Birthday calendar", "events and sharing", ""]];
    const c5 = cols(5, 0.2); const th = 1.5;
    st.forEach(([t, b, tm], i) => {
      const { x, w } = c5[i % 5]; const y = Y0 + Math.floor(i / 5) * (th + 0.2);
      box(s, x, y, w, th, i < 5 ? HX.tPurple : HX.lav);
      T(s, String(i + 1).padStart(2, "0"), { x: x + 0.2, y: y + 0.18, w: 0.7, h: 0.4, fontSize: 20, bold: true, color: HX.purple, fontFace: "Arial" });
      if (tm) { const pw = pillW(tm); pill(s, tm, x + w - 0.2 - pw, y + 0.2, { fill: HX.white }); }
      T(s, t, { x: x + 0.2, y: y + 0.65, w: w - 0.4, h: 0.35, fontSize: 16, bold: true });
      T(s, b, { x: x + 0.2, y: y + 1.0, w: w - 0.4, h: 0.45, fontSize: 13, color: HX.body });
    });
    const sy = Y0 + 2 * th + 0.2 + 0.55; const c3 = cols(3);
    [["~70 s", "total on the test sheet"], ["30 s", "wait for the Sync Lock, else skip that night"], ["5 min", "budget; remaining steps are skipped and logged"]].forEach(([n, l], i) => {
      T(s, n, { x: c3[i].x, y: sy, w: c3[i].w, h: 0.75, fontSize: 40, bold: true, color: HX.purple, fontFace: "Arial" });
      T(s, l, { x: c3[i].x, y: sy + 0.8, w: c3[i].w, h: 0.35, fontSize: 15, color: HX.grey });
    });
  }
  {
    const s = content("Table registry and safe schema migrations");
    await cardRow(s, [{ icon: "MdBookmarks", title: "Tabs are found by ID, not by name", body: ["src/Tables.js maps a stable key to each tab: members → Members, tenures → Contracts, baselines → Routines, weeklyMatrix → Weekly Presence …", "Lookup order: stored sheet ID (document property SHEET_ID_<key>), then readable name, then former technical name.", "So a team can rename a tab without breaking anything.", "Header lookups accept the internal key or the readable label."] },
      { icon: "MdSchema", title: "Upgrades only add", body: ["SCHEMA_SPEC lists every column with its type, required or optional status, protection and dropdown values.", "A new column widens the existing native table; data is never moved or deleted.", "Rules are replaced, not stacked: required-cell highlights, colour rules, dropdowns.", "A native table name clash with an archived tab is renamed archived_<name>_<id>.", "The schema is fixed for users: column changes go through the maintainer and an ADR."] }], Y0, 2, { base: { space: 9 } });
  }
  {
    const s = content("The Sync Lock keeps scripts from colliding");
    const intro = "People can always type; the lock only stops two scripts writing from stale snapshots at the same time.";
    T(s, intro, { x: X0, y: Y0, w: W, h: 0.4, fontSize: 16, color: HX.body });
    const ty = Y0 + 0.75;
    box(s, X0, ty + 0.17, W, 0.06, HX.line, { rect: true });
    const seg = [[X0 + 0.3, 2.2, HX.green, "Edit: runs"], [X0 + 2.8, 5.9, HX.purple, "Nightly job holds the lock (≈ 70 s)"], [X0 + 9.0, 2.8, HX.green, "Edits run normally"]];
    seg.forEach(([x, w, c, t]) => { box(s, x, ty, w, 0.4, c, { r: 0.2 }); T(s, t, { x, y: ty, w, h: 0.4, fontSize: 14, bold: true, color: HX.white, align: "center", valign: "middle" }); });
    const y = ty + 0.8;
    await cardRow(s, [{ icon: "MdLock", title: "What it is", body: "One LockService document lock per team sheet. Not a script lock: the Library is shared, and a script lock would make every team's 2:00 job queue behind the others." }, { icon: "MdHourglassTop", title: "Who waits", body: "The nightly job waits up to 30 s, otherwise skips that night. An edit waits up to 10 s; if it times out, the typed value stays, follow-up writes are skipped and a toast says “Sync in progress”." }, { icon: "MdReplay", title: "How it catches up", body: "Tasks and events without an external ID are pushed on the next run. A Weekly Presence entry typed during the sync may be redrawn away and must be typed again." }], y, 3, { minH: Y1 - y });
  }
  {
    const s = content("KPI Data feeds Looker Studio");
    T(s, "One row per Member × date × half-day. No totals in the sheet: Looker aggregates, so nothing is counted twice.", { x: X0, y: Y0, w: W, h: 0.4, fontSize: 16, color: HX.body });
    const cl = ["date", "month", "iso_week", "weekday", "is_weekend", "member", "department", "role", "fte_rate", "period", "location", "category", "worked_days", "leave_days", "capacity_days", "is_public_holiday"];
    const c8 = cols(8, 0.12);
    cl.forEach((c, i) => {
      const { x, w } = c8[i % 8]; const y = Y0 + 0.6 + Math.floor(i / 8) * 0.52;
      const hl = ["category", "worked_days", "leave_days", "capacity_days"].includes(c);
      box(s, x, y, w, 0.42, hl ? HX.tPurple : HX.tGrey, { r: 0.21 });
      T(s, c, { x, y, w, h: 0.42, fontSize: c.length > 13 ? 9.5 : 11, bold: hl, color: hl ? HX.purple : HX.ink, align: "center", valign: "middle", fontFace: "Courier New" });
    });
    const ty = Y0 + 1.75, tw = 8.4;
    const th = table(s, [["Measure", "Rule"], ["category", "office, client_site or remote from the location's Type; or leave, off, not_logged"], ["worked_days", "0.5 per half-day at an office, client site or remote"], ["leave_days", "0.5 per half-day on leave"], ["capacity_days", "0.5 × FTE per half-day on a working day inside a contract; 0 on weekends and public holidays"], ["fte_rate", "0.8 and 80 both mean 80 %"]], { x: X0, y: ty, colW: [2.0, tw - 2.0], fontSize: 14, fillH: Y1 - ty });
    const rx = X0 + tw + 0.3, rw = X1 - rx;
    box(s, rx, ty, rw, Y1 - ty, HX.deep);
    T(s, "Example measures in Looker", { x: rx + PAD, y: ty + PAD, w: rw - 2 * PAD, h: 0.4, fontSize: 16, bold: true, color: HX.white });
    bullets(s, ["Remote rate = remote worked ÷ worked", "Office presence by department and month", "Leave days vs. capacity", "Not-logged half-days per week"], rx + PAD, ty + PAD + 0.55, rw - 2 * PAD, { fontSize: 15, color: HX.white, space: 8 });
  }

  // ================= SECTION 5
  section("5 · Flexibility vs. data integrity"); kicker = "5 · Flexibility vs. data integrity";
  await divider(5, "The flexibility-to-integrity ratio", "The sheet stays a spreadsheet: people can type almost anywhere. Integrity is enforced only where a wrong value would silently corrupt something derived from it.", ["The design principle", "The ratio, in numbers", "Table by table", "Seven layers of guards", "Grid or sidebar", "What flexibility costs"]);
  {
    const s = content("Principle: maximum flexibility, guards where it counts");
    const pw = (W - 1.2) / 2;
    const L = ["Every visible grid stays editable, Weekly Presence included.", "Hidden tabs are hidden for tidiness, never locked.", "Sidebars are an extra path, not a gate.", "Full transparency: no privacy model; everyone sees every profile.", "History is ended or cancelled, not erased, so mistakes can be traced."];
    const R = ["Machine-owned values (IDs, timestamps, derived columns, Google links) are locked.", "Keys other rows depend on ask for confirmation before a change.", "Types, dropdowns and chips shape input; required fields are flagged.", "Scripts never write over each other (Sync Lock).", "The nightly job reconciles; Check Workspace reports what is off."];
    [[X0, HX.lav, HX.green, "Flexibility", L], [X0 + pw + 1.2, HX.tPurple, HX.purple, "Integrity", R]].forEach(([x, f, c, t, items]) => {
      box(s, x, Y0, pw, H, f);
      T(s, t, { x: x + 0.35, y: Y0 + 0.3, w: pw - 0.7, h: 0.5, fontSize: 22, bold: true, color: c, fontFace: "Arial" });
      bullets(s, items, x + 0.35, Y0 + 1.0, pw - 0.7, { fontSize: 16, space: 12 });
    });
    const cx = X0 + pw + 0.6 - 0.5, cy = Y0 + H / 2 - 0.5;
    s.addShape(pres.shapes.OVAL, { x: cx, y: cy, w: 1.0, h: 1.0, fill: { color: HX.ink }, line: { color: HX.white, width: 3 }, objectName: "balance" });
    s.addImage({ data: await icon("MdBalance", HX.white), x: cx + 0.22, y: cy + 0.22, w: 0.56, h: 0.56, objectName: "balance-icon" });
  }
  {
    const s = content("3.5 editable columns for every locked one");
    const st = [["72", "columns across the 11 managed tables", HX.ink, HX.tGrey], ["44", "free to edit · 61 %", HX.green, HX.lav], ["12", "need a confirmation · 17 %", HX.amber, HX.lav], ["16", "locked · 22 %", HX.purple, HX.lav]];
    const c4 = cols(4);
    st.forEach(([n, l, c, f], i) => {
      const { x, w } = c4[i]; box(s, x, Y0, w, 1.85, f);
      T(s, n, { x: x + PAD, y: Y0 + 0.15, w: w - 2 * PAD, h: 1.0, fontSize: 54, bold: true, color: c, fontFace: "Arial" });
      T(s, l, { x: x + PAD, y: Y0 + 1.2, w: w - 2 * PAD, h: 0.55, fontSize: 15, color: HX.body });
    });
    const by = Y0 + 2.2; let x = X0;
    [[44, HX.green, "Free 44"], [12, HX.amber, "Confirm 12"], [16, HX.purple, "Locked 16"]].forEach(([n, c, t]) => { const w = W * n / 72; box(s, x, by, w, 0.6, c, { rect: true }); T(s, t, { x, y: by, w, h: 0.6, fontSize: 15, bold: true, color: HX.white, align: "center", valign: "middle" }); x += w; });
    const y = by + 0.95; heading(s, "How to read it", X0, y, W);
    bullets(s, [["Editable : locked = 56 : 16 ≈ 3.5 : 1. ", "Free : guarded = 44 : 28 ≈ 1.6 : 1."], ["15 of the 16 locked columns are written by the system ", "(IDs, timestamps, derived values, Google links). The 16th is the Member chip in Members, the identity everything hangs on."], ["Locks cover rows already created, ", "so new entries never need an admin, except in Members, where the five derived columns are locked on every row: people join through a contract."]], X0, y + 0.45, W, { fontSize: 14, space: 6 });
  }
  {
    const s = content("Table by table: where the balance sits");
    const tbls = ["Locations", "Holidays", "Tasks", "Contracts", "Routines", "Leaves", "Backups", "Events", "Projects", "Presence Log", "Members"];
    const data = { Locations: [3, 0, 0], Holidays: [2, 0, 0], Tasks: [6, 0, 2], Contracts: [5, 2, 0], Routines: [5, 1, 1], Leaves: [5, 2, 1], Backups: [5, 2, 1], Events: [4, 1, 2], Projects: [4, 2, 2], "Presence Log": [2, 2, 2], Members: [3, 0, 5] };
    const labels = tbls.slice().reverse(); const cw = 7.7;
    s.addChart(pres.charts.BAR, [{ name: "Free", labels, values: labels.map((t) => data[t][0]) }, { name: "Confirm first", labels, values: labels.map((t) => data[t][1]) }, { name: "Locked", labels, values: labels.map((t) => data[t][2]) }], {
      x: X0, y: Y0, w: cw, h: H, barDir: "bar", barGrouping: "stacked", barGapWidthPct: 45, chartColors: [HX.green, HX.amber, HX.purple],
      showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFontSize: 12, dataLabelFontBold: true, dataLabelFormatCode: "#;;", dataLabelFontFace: "+mn-lt",
      catAxisLabelColor: HX.body, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt", valAxisLabelColor: HX.soft, valAxisLabelFontSize: 12, valAxisLabelFontFace: "+mn-lt", valAxisMaxVal: 8, valAxisMajorUnit: 2,
      valGridLine: { color: "E8EAED", size: 0.75 }, catGridLine: { style: "none" }, showLegend: true, legendPos: "t", legendFontSize: 13, legendFontFace: "+mn-lt", legendColor: HX.body,
      showTitle: true, title: "Columns per table by edit freedom", titleFontSize: 15, titleColor: HX.ink, titleFontFace: "+mn-lt", objectName: "ratio-chart" });
    const rx = X0 + cw + 0.4, rw = X1 - rx;
    heading(s, "Reading the chart", rx, Y0, rw);
    bullets(s, [["Configuration tables ", "(Locations, Holidays) are fully free: they only feed dropdowns and calendars."], ["Tasks ", "is the freest data table: only its ID and Google link are locked, because people also edit tasks in Google Tasks."], ["Presence Log ", "is the most guarded: KPIs and Weekly Presence read it."], ["Members ", "is mostly derived: change the contract, not the row."]], rx, Y0 + 0.5, rw, { fontSize: 15, space: 10 });
  }
  {
    const s = content("Seven layers of guards, from nudge to lock");
    const L = [["Shape", "Column types, dropdowns, people chips and placeholders prevent most typos.", HX.green], ["Signal", "Empty required cells turn red; a toast names the missing fields.", HX.green], ["Guide", "Sidebars refuse overlapping contracts, report gaps, offer to end an old routine, refuse a lead as their own interim.", HX.blue], ["Confirm", "Keys other rows rely on (contract start, leave type, project lead …) ask for confirmation.", HX.amber], ["Lock", "IDs and derived columns: only the owner and ADMIN_EMAILS can edit, with a warning (snoozable for 5 min).", HX.purple], ["Serialise", "The Sync Lock stops scripts writing from stale snapshots.", HX.purple], ["Reconcile", "The nightly sync rewrites derived values; compile is idempotent; Check Workspace reports drift.", HX.deep]];
    const rh = (H - 0.45 - 6 * 0.12) / 7;
    L.forEach(([t, b, c], i) => {
      const y = Y0 + i * (rh + 0.12); const w = 2.4 + i * 0.25;
      box(s, X0, y, w, rh, c, { r: 0.12 });
      T(s, (i + 1) + "   " + t, { x: X0 + 0.25, y, w: w - 0.3, h: rh, fontSize: 16, bold: true, color: HX.white, valign: "middle" });
      T(s, b, { x: X0 + 4.4, y, w: W - 4.4, h: rh, fontSize: 15, color: HX.body, valign: "middle" });
    });
    note(s, "Softer layers act on every edit; harder ones only on the few columns where an error would spread into derived data.", Y1 - 0.3);
  }
  {
    const s = content("Grid or sidebar: speed versus guidance");
    const th = table(s, [["Task", "In the grid (fast, free)", "In a sidebar (checked)"], ["Add a person", "Type the name in Members (admins only: the column is locked)", "Members & Contracts: first contract form, open to everyone"], ["Change a contract", "Unhide Contracts and type a row", "New contract: ends the current one, refuses overlaps, reports gaps"], ["Usual week", "Unhide Routines and type rows", "Routine Baseline: presets, overlap prompt, history kept"], ["One day", "A Weekly Presence cell", "—"], ["Leave", "Unhide Leaves and type a row", "Submit Leave: half-days, quick ranges, compiled at once"], ["Cover an absence", "Unhide Backups and type a row", "Interim Backups: shows gaps, flags the person's leaves"], ["Tasks, Events, settings", "The grid is the only path", "—"]], { x: X0, y: Y0, colW: [2.6, 4.5, W - 7.1], fontSize: 14 });
    note(s, "Both paths write the same tables and pass the same validation, protections and nightly reconciliation. The sidebar adds checks across rows (overlaps, gaps) that a single cell edit cannot see.", Y0 + th + 0.3);
  }
  {
    const s = content("What flexibility costs, and how it is caught");
    const th = table(s, [["Trade-off", "What can go wrong", "Safety net"], ["Hand entries win in the Presence Log", "A day typed before a leave was approved keeps the typed location, also in KPIs", "Coverage reads Leaves directly; fix the day in Weekly Presence"], ["Hidden tabs stay editable", "A row typed there skips the sidebar's cross-row checks", "Red required cells, protection at the next nightly run, Check Workspace"], ["Admins can edit locked columns", "A derived value is overwritten by hand", "Warning toast; the next sync rewrites it"], ["Edits during the 2:00 sync", "Follow-up writes are skipped", "“Sync in progress” toast; the next run catches up"], ["No privacy model", "Everyone sees every contract and contact", "Birthdays store day and month only; data stays in the shared sheet"], ["Requested leaves are not compiled", "Presence shows a routine for a day that may become a leave", "Orange coverage warning before approval"]], { x: X0, y: Y0, colW: [3.1, 4.6, W - 7.7], fontSize: 13 });
    const t = "Rule of thumb for maintainers: add a guard only when a wrong value would silently spread into derived data; otherwise prefer a signal (toast, highlight, Check Workspace) over a lock.";
    const y = Y0 + th + 0.35;
    box(s, X0, y, W, textH(t, 15, W - 0.6, "Calibri", true) + 0.4, HX.tPurple);
    T(s, t, { x: X0 + 0.3, y: y + 0.2, w: W - 0.6, h: textH(t, 15, W - 0.6, "Calibri", true) + 0.05, fontSize: 15, bold: true, color: HX.deep });
  }

  // ================= SECTION 6
  section("6 · Troubleshooting & reference"); kicker = "6 · Troubleshooting & reference";
  await divider(6, "Troubleshooting & reference", "Common symptoms and their fixes, performance figures and limits, and where to find the detail.", ["Symptoms and fixes", "Performance and limits", "Reference"]);
  {
    const s = content("Symptoms and fixes");
    table(s, [["Symptom", "Likely cause", "Fix"], ["No Project Management menu", "Sheet opened before the script loaded, or Library not linked", "Reload; check Libraries → WorkspacePM in Apps Script"], ["“Sync in progress” toast", "The nightly job holds the Sync Lock", "Wait a minute; retype a Weekly Presence entry"], ["Status says No Contract", "The Member has no contract", "Members & Contracts → first contract"], ["Today says Not Logged", "No Presence Log row for today yet", "Type today in Weekly Presence, or wait for the next compile"], ["Role, Status or Today changed back", "Derived columns are rewritten from contracts and presence", "Change the contract or the presence instead"], ["A location is refused", "It is retired (Active unticked) or not listed", "Edit Locations: add it or tick Active"], ["Project row red or orange", "A coverage gap", "Interim Backups, or change the backup"], ["Task not in my Google Tasks", "The sync uses the account that runs it, for its own tasks", "Check Assigned to; run the sync as that account"], ["Nightly steps skipped", "Five-minute budget reached", "See Apps Script → Executions; they run the next night"]], { x: X0, y: Y0, colW: [3.3, 4.5, W - 7.8], fontSize: 14, fillH: H });
  }
  {
    const s = content("Performance and platform limits");
    const tw = 6.3;
    table(s, [["Operation", "Measured on the test sheet"], ["Members edit (sync + protections)", "~15 s"], ["Initialize / Upgrade", "~110 s"], ["Nightly job", "~70 s (directory 29 s, compile 19 s)"], ["Check Workspace", "~25 s"], ["Interim Backups sidebar load", "~15 s"], ["Other edits", "Under 1 s for the simple trigger, a few seconds for the handler"]], { x: X0, y: Y0, colW: [3.0, tw - 3.0], fontSize: 15, fillH: H });
    const rx = X0 + tw + 0.45, rw = X1 - rx;
    heading(s, "Google limits the design respects", rx, Y0, rw);
    bullets(s, [["6 minutes ", "per execution: the nightly job stops at 5 and logs what it skipped."], ["30 seconds ", "for simple triggers: heavy work runs in the installable trigger."], ["Reads go to the last row with data, ", "never a fixed cap, so the Presence Log can grow without duplicates."], ["Typed table columns ", "refuse number formats; dates use the table's DATE type."], ["Looker's cache ", "is about 15 minutes; KPI Data refreshes nightly."]], rx, Y0 + 0.5, rw, { fontSize: 15, space: 10 });
  }
  {
    const s = content("Reference");
    const r = [["MdMenuBook", "GLOSSARY.md", "Every domain term and the words to avoid."], ["MdGavel", "docs/adr/0001–0044", "Architecture decisions, one per change; the latest are 0036–0044."], ["MdDescription", "CLAUDE.md", "Project facts and rules for maintainers."], ["MdIntegrationInstructions", "consumer/Code.js", "The wrapper to paste into each team sheet."], ["MdTableChart", "src/Schema.js", "Every column: type, required, protection, dropdown values."], ["MdScience", "test/", "214 Node tests; run node --test test/*.js."]].map(([icon, title, body]) => ({ icon, title, body }));
    const rh = (H - GAP) / 2;
    await cardRow(s, r.slice(0, 3), Y0, 3, { minH: rh });
    await cardRow(s, r.slice(3), Y0 + rh + GAP, 3, { minH: rh });
  }
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: curSection });
    s.addText("Questions or changes?", { placeholder: "title" });
    T(s, "Run Check Workspace first: most issues come with a link to their fix. Schema changes and new features go through the maintainer and land as an ADR.", { x: 0.8, y: 3.45, w: 10.5, h: 1.0, fontSize: 20, color: HX.lilac });
    let x = 0.8; CHIPS.forEach(([t, f, c]) => { x += pill(s, t, x, 4.95, { fill: f, color: c, fontSize: 13, h: 0.38 }) + 0.15; });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
