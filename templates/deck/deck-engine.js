// ============================================================================
// deck-engine.js — 柔光蓝建 deck 引擎(可复用)
// ----------------------------------------------------------------------------
// 把"柔光蓝风格 + 矢量图标 + 圆角版式 + 边距纪律"沉淀成一套 primitives。
// 任何大纲都可以:require 本引擎 → createEngine(pres) → 用 primitives 拼版式。
//
// 用法:
//   const pptxgen = require("pptxgenjs");
//   const { createEngine } = require("<path>/deck-engine");
//   const pres = new pptxgen(); pres.layout = "LAYOUT_WIDE";
//   const E = createEngine(pres);
//   const { C, IK, FA, card, panel, tabCard, statCard, newSlide, addPageTitle } = E;
//   // ... 用 primitives 画每一页 ...
//   await pres.writeFile({ fileName: "out.pptx" });
//
// 依赖(需全局或本地装好,运行时带 NODE_PATH=$(npm root -g)):
//   pptxgenjs  react-icons  react  react-dom  sharp
//
// 设计纪律(改这里之前先读 styles/soft-blue-unionpay.md 和 templates/deck/PLAYBOOK.md):
//   1) 单色系柔光蓝:一种主蓝主导 + 一个柔青辅色,负向用低饱和陶土且极少;不用金、不用纯红纯绿。
//   2) 全圆角、浅填充、留白多;三级层次都在一个蓝里(中蓝实底 / 淡蓝底 / 白卡)。
//   3) 边距纪律:每页内容一律收在 x=0.5 ~ 12.8(左右各 0.5" 留白),多卡行右缘必须对齐 12.8。
//   4) 图标克制:只在"有意义处"用(标题1枚 / 主卡系列 / 关键面板徽标),不做填充式水印。
//      —— iconAt 带 transparency 参数的"水印式"调用会被直接跳过(见函数内注释)。
// ============================================================================

const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

// ---------- 柔光蓝配色 v2「提纯版」(pptx 用,无 #)——蓝更净、青更清爽、浅色更冷,去灰去脏 ----------
const C = {
  // 文字色:标题用"深蓝"(混主题色、比纯黑轻)、正文深炭(配宋体不发淡)、次级 #666——用户三轮微调定稿
  ink: "2D5BA6", text: "262626", textMid: "666666", textMute: "999999",
  navy: "1E3A5C", // 深蓝:仅用于个别"深色表头/矩阵角"填充(配白字),不再当正文标题色
  blue: "2C66C2", blueMid: "5A8AD8", blueSoft: "A9CBEE",
  bluePale: "E7F1FC", blueWash: "F0F6FD", teal: "2C9CAB", tealPale: "D6EEF1",
  bg: "F6F9FD", card: "FFFFFF", border: "DCE7F3", borderDash: "B6CEEA",
  good: "33A07E", goodPale: "D7EFE6", warn: "D07C57", warnPale: "F6E4DA",
  white: "FFFFFF", coverPale: "D5E6F8",
};
// ---------- 图标用 CSS 色(带 #) ----------
const IK = { white: "#FFFFFF", blue: "#2C66C2", teal: "#2C9CAB", ink: "#1E3A5C", warn: "#D07C57", good: "#33A07E", mute: "#97A6B8", soft: "#A9CBEE" };

// 字体规范:标题=微软雅黑,正文=宋体(笔画较实、不发淡;仿宋偏细在小字下显"淡",已弃用为默认),数字/代码=Arial/Consolas。
// 均 Windows 自带;macOS 预览会替换为相近字体,要两端一致就 PowerPoint「嵌入字体」。
const FONT_TITLE = "微软雅黑", FONT_BODY = "宋体", FONT_NUM = "Arial", MONO = "Consolas";
const FONT = FONT_TITLE; // 向后兼容别名(默认=标题字体)
// 行高规范:多行正文 / 分点列表统一 lineSpacingMultiple: LH_BODY(=1.5),读着透气、不挤。
// 单行标签 / 大数字 / 表格密集数据不强求(给 1.5 会撑高、易溢出);容器高度按 1.5 预留。
const LH_BODY = 1.5;

// ---------- 正文造重点:**关键词** → 加粗(每行/每点给一个视觉落点,别整句粗) ----------
// rt(str, opt)    → 整段正文的 runs(传给 addText 第一参数)
// rtBul(arr, opt) → 一组分点的 runs(字面"•  "前导 + 段末 breakLine;不依赖 pptxgenjs bullet 段落属性——多 run 混排时它不稳)
function _segBold(str) {
  const p = String(str).split(/\*\*(.+?)\*\*/g), o = [];
  for (let i = 0; i < p.length; i++) if (p[i] !== "") o.push([p[i], i % 2 === 1]);
  return o;
}
function rt(str, opt = {}) {
  return _segBold(str).map(([t, b]) => ({ text: t, options: { fontFace: FONT_BODY, ...opt, bold: b || !!opt.bold } }));
}
function rtBul(arr, opt = {}) {
  const out = [], base = { fontFace: FONT_BODY, color: C.text, fontSize: 12, lineSpacingMultiple: LH_BODY, ...opt };
  (arr || []).forEach((str) => {
    const segs = _segBold(str);
    out.push({ text: "•  ", options: { ...base, bold: false, color: C.textMid } });
    segs.forEach(([t, b], ri) => out.push({
      text: t,
      options: { ...base, bold: b || !!opt.bold, ...(ri === segs.length - 1 ? { breakLine: true, paraSpaceAfter: opt.paraSpaceAfter != null ? opt.paraSpaceAfter : 4 } : {}) },
    }));
  });
  return out;
}
const PW = 13.3, PH = 7.5;            // LAYOUT_WIDE 画布
const MARGIN = 0.5;                    // 左右统一留白
const RIGHT = PW - MARGIN;             // = 12.8,所有内容右缘
const CONTENT_W = PW - MARGIN * 2;     // = 12.3,满宽内容宽度
// 纵向间距标准:内容带 BTOP→BBOT,块间距 GAP;最底的块底边落到 BBOT 附近(别让底部空一条、内容"吊在上面")。
const BTOP = 2.15, BBOT = 6.62, GAP = 0.3;
const botY = (h) => BBOT - h;          // 底栏锚定:高 h 的底栏从这个 y 起,底边正好落 BBOT
const RAD = 0.1, RADS = 0.07;          // 圆角

// ---------- 图标栅格化(memo,跨页缓存) ----------
const _ic = new Map();
async function rasterize(Comp, hex, size = 240) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: hex, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}
async function ICON(Comp, hex) {
  let m = _ic.get(Comp); if (!m) { m = new Map(); _ic.set(Comp, m); }
  if (m.has(hex)) return m.get(hex);
  const d = await rasterize(Comp, hex); m.set(hex, d); return d;
}

// ============================================================================
// createEngine(pres):返回一组绑定到该 pres 的版式 primitives。
// ============================================================================
function createEngine(pres) {
  const RR = pres.shapes.ROUNDED_RECTANGLE, RECT = pres.shapes.RECTANGLE,
    OVAL = pres.shapes.OVAL, LINE = pres.shapes.LINE;
  const lift = () => ({ type: "outer", color: "9FB3CC", blur: 6, offset: 2, angle: 90, opacity: 0.16 });

  // ---- 图标 ----
  // 图标 + 柔色圆圈(母题:章节标题左 / 卡头 / 要素卡)
  async function iconCircle(s, Comp, x, y, d, circleFill, iconHex, pad = 0.27) {
    const data = await ICON(Comp, iconHex);
    s.addShape(OVAL, { x, y, w: d, h: d, fill: { color: circleFill }, line: { color: circleFill, width: 0 } });
    const ip = d * pad;
    s.addImage({ data, x: x + ip, y: y + ip, w: d - 2 * ip, h: d - 2 * ip });
  }
  // 裸图标。⚠ 带 transparency 的"水印式"调用一律跳过 —— 图标只在有意义处出现,不做填充装饰。
  async function iconAt(s, Comp, iconHex, x, y, w, h, transparency) {
    if (transparency) return;
    const data = await ICON(Comp, iconHex);
    s.addImage({ data, x, y, w, h });
  }
  // 彩色面板上的图标徽标:白圈 + 彩色图标(用于中蓝/陶土实底面板)
  async function chipOnPanel(s, Comp, x, y, d, iconHex) {
    s.addShape(OVAL, { x, y, w: d, h: d, fill: { color: "FFFFFF" }, line: { width: 0, color: "FFFFFF" } });
    const ip = d * 0.27; await iconAt(s, Comp, iconHex, x + ip, y + ip, d - 2 * ip, d - 2 * ip);
  }

  // ---- 容器 ----
  function card(s, x, y, w, h, opt = {}) {
    s.addShape(RR, {
      x, y, w, h, rectRadius: opt.rad != null ? opt.rad : RAD,
      fill: { color: opt.fill || C.card }, line: { color: opt.border || C.border, width: opt.bw != null ? opt.bw : 0.75 },
      ...(opt.shadow ? { shadow: lift() } : {})
    });
  }
  function panel(s, x, y, w, h, fill = C.blue, rad = RAD) {
    s.addShape(RR, { x, y, w, h, rectRadius: rad, fill: { color: fill }, line: { color: fill, width: 0 } });
  }
  function groupBox(s, x, y, w, h, fill = C.blueWash) {
    s.addShape(RR, { x, y, w, h, rectRadius: RAD, fill: { color: fill }, line: { color: C.borderDash, width: 1, dashType: "dash" } });
  }
  // 连线箭头(画架构图 / 数据流 / 闭环):从 (x1,y1) 连到 (x2,y2),箭头指向终点。
  // opt: { color, width, both(双向), dash:'dash' }。水平/垂直最稳;靠 flip 让箭头落在 (x2,y2)。
  function connector(s, x1, y1, x2, y2, opt = {}) {
    s.addShape(LINE, {
      x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
      line: { color: opt.color || C.blueSoft, width: opt.width || 1.5, endArrowType: "triangle", beginArrowType: opt.both ? "triangle" : "none", ...(opt.dash ? { dashType: opt.dash } : {}) },
      flipH: x2 < x1, flipV: y2 < y1
    });
  }
  // 卡头(彩色实底)+ 卡身(白/浅),小间隙;hdrIcon 给个白色图标贴卡头左侧
  async function tabCard(s, x, y, w, totalH, headerH, headerFill, headerRuns, bodyFill, hdrIcon) {
    panel(s, x, y, w, headerH, headerFill, RADS);
    let tx = x + 0.18;
    if (hdrIcon) {
      const d = Math.min(0.32, headerH - 0.16);
      await iconAt(s, hdrIcon, IK.white, x + 0.2, y + (headerH - d) / 2, d, d);
      tx = x + 0.2 + d + 0.12;
    }
    s.addText(headerRuns, { x: tx, y, w: w - (tx - x) - 0.14, h: headerH, valign: "middle", margin: 0 });
    const by = y + headerH + 0.06, bh = totalH - headerH - 0.06;
    if (bh > 0.05) card(s, x, by, w, bh, { fill: bodyFill || C.card });
    return { by, bh };
  }

  // ---- 小元件 ----
  // ⚠ 参数顺序:numDot(s, x, y, num, size, color) —— 先 num(显示的数字)后 size(直径)。别和别处的 numBadge(…size,num…) 记混。
  function numDot(s, x, y, num, size = 0.36, color = C.blue) {
    s.addShape(OVAL, { x, y, w: size, h: size, fill: { color }, line: { color, width: 0 } });
    s.addText(String(num), { x, y, w: size, h: size, color: C.white, bold: true, fontSize: size > 0.34 ? 14 : 12, fontFace: FONT_NUM, align: "center", valign: "middle", margin: 0 });
  }
  // 大数字卡(数字用主蓝/柔青,不用金)
  function statCard(s, x, y, w, h, num, unit, label, accent = C.blue) {
    card(s, x, y, w, h, { fill: C.card });
    s.addText([
      { text: num, options: { color: accent, bold: true, fontSize: 34, fontFace: FONT_NUM } },
      { text: unit ? " " + unit : "", options: { color: C.textMid, fontSize: 13, fontFace: FONT } }
    ], { x: x + 0.22, y: y + 0.16, w: w - 0.34, h: 0.62, margin: 0, valign: "middle" });
    s.addText(label, { x: x + 0.22, y: y + 0.82, w: w - 0.34, h: h - 0.94, color: C.text, fontSize: 11, fontFace: FONT_BODY, margin: 0, valign: "top" });
  }
  // 小节标题(纯文字,克制,不挂图标;颜色区分组别)
  function sectionLabel(s, x, y, w, text, color = C.blue) {
    s.addText(text, { x, y, w, h: 0.32, color, bold: true, fontSize: 13, fontFace: FONT, charSpacing: 2, margin: 0, valign: "middle" });
  }

  // ---- 页面骨架 ----
  // 页眉:章节序号徽标 + 章节名(不再画"idx/total"页码——页码统一交母板自动字段,见 defineSlideMaster)
  function addHeader(s, partNo, sectionTitle, idx, total) {
    s.addShape(RR, { x: MARGIN, y: 0.34, w: 0.5, h: 0.34, rectRadius: RADS, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
    s.addText(partNo, { x: MARGIN, y: 0.34, w: 0.5, h: 0.34, color: C.white, bold: true, fontSize: 13, fontFace: FONT_NUM, align: "center", valign: "middle", margin: 0 });
    s.addText(sectionTitle, { x: 1.12, y: 0.34, w: 9, h: 0.34, color: C.ink, bold: true, fontSize: 13, fontFace: FONT, valign: "middle", margin: 0 });
  }
  // 页脚:细分隔线 + 左侧脚注(页码不在此画,交母板自动字段——重排幻灯片自动更新,不用手改)
  function addFooter(s, idx, footText = "") {
    s.addShape(LINE, { x: MARGIN, y: PH - 0.42, w: CONTENT_W, h: 0, line: { color: C.border, width: 0.75 } });
    s.addText(footText, { x: MARGIN, y: PH - 0.36, w: 6, h: 0.3, color: C.textMute, fontSize: 9, fontFace: FONT_BODY, margin: 0 });
  }
  // 页大标题 + 左侧圆圈图标(每页 1 枚,母题)
  async function addPageTitle(s, Comp, title, subtitle) {
    if (Comp) await iconCircle(s, Comp, MARGIN, 0.95, 0.5, C.bluePale, IK.blue, 0.26);
    const tx = Comp ? 1.16 : MARGIN;
    s.addText(title, { x: tx, y: 0.92, w: RIGHT - tx, h: 0.58, color: C.ink, fontSize: 25, bold: true, fontFace: FONT, margin: 0, valign: "middle" });
    if (subtitle) s.addText(subtitle, { x: MARGIN, y: 1.56, w: CONTENT_W, h: 0.36, color: C.textMid, fontSize: 13, fontFace: FONT_BODY, margin: 0, valign: "top" });
  }
  // 母板:统一承载"自动页码"字段(右下角)。重排/增删幻灯片时 PowerPoint 自动更新,无需手改。
  // 封面用普通 addSlide(不挂此母板)→ 不显示页码;内容页 newSlide 走此母板 → 自动出页码。
  const MASTER = "DECK_MASTER";
  pres.defineSlideMaster({
    title: MASTER,
    background: { color: C.bg },
    objects: [],
    slideNumber: { x: RIGHT - 0.8, y: PH - 0.4, w: 0.8, h: 0.3, align: "right", color: C.blue, bold: true, fontSize: 11, fontFace: FONT_NUM },
  });
  // 新建内容页(浅底 + 页眉 + 页脚;页码由母板自动字段提供)
  function newSlide(partNo, sectionTitle, idx, opt = {}) {
    const s = pres.addSlide({ masterName: MASTER }); s.background = { color: C.bg };
    addHeader(s, partNo, sectionTitle, idx, opt.total || 15);
    addFooter(s, idx, opt.foot || "");
    return s;
  }

  // 横向均分布局:返回每个卡片的 x 和宽度,保证整行收在 [MARGIN, RIGHT] 内
  // count 个卡 + (count-1) 个间隙 gap,左右贴合 MARGIN/RIGHT。
  function row(count, gap = 0.2, left = MARGIN, right = RIGHT) {
    const w = (right - left - gap * (count - 1)) / count;
    return { w, x: (i) => left + i * (w + gap) };
  }

  return {
    pres, C, IK, FA, FONT, FONT_TITLE, FONT_BODY, FONT_NUM, MONO, LH_BODY, rt, rtBul, PW, PH, MARGIN, RIGHT, CONTENT_W, BTOP, BBOT, GAP, botY, RAD, RADS,
    RR, RECT, OVAL, LINE,
    ICON, iconCircle, iconAt, chipOnPanel,
    card, panel, groupBox, connector, tabCard, numDot, statCard, sectionLabel,
    addHeader, addFooter, addPageTitle, newSlide, row,
  };
}

module.exports = { createEngine, C, IK, FA, LH_BODY, rt, rtBul, PW, PH, MARGIN, RIGHT, CONTENT_W, BTOP, BBOT, GAP, botY };
