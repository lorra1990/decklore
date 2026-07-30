// ============================================================================
// example-deck.js — 用 deck-engine 从"大纲"拼一份 deck 的最小起手式
// ----------------------------------------------------------------------------
// 复制本文件,把 OUTLINE 换成你的大纲,按"版式模式"往下填即可。
// 运行:  NODE_PATH=$(npm root -g) node example-deck.js
//        (需先全局装好 pptxgenjs react-icons react react-dom sharp)
// 看图:  python3 <pptx-skill>/scripts/office/soffice.py --headless --convert-to pdf example-deck.pptx
//        pdftoppm -jpeg -r 120 example-deck.pdf slide
// ============================================================================

const pptxgen = require("pptxgenjs");
const { createEngine } = require("./deck-engine");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

const E = createEngine(pres);
const { C, IK, FA, FONT, FONT_NUM, PW, RIGHT, MARGIN,
  card, panel, tabCard, statCard, numDot, sectionLabel, softChart,
  addPageTitle, newSlide, chipOnPanel, row } = E;

const FOOT = "示例 deck · 用 deck-engine 生成";
const slide = (p, t, i, total) => newSlide(p, t, i, { foot: FOOT, total });

// ---- 你的大纲(示意) ----
const TOTAL = 5;

(async () => {

  // ===== 模式 A:封面 =====
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.2, h: 4.2, fill: { color: C.bluePale }, line: { color: C.bluePale, width: 0 } });
    s.addText("DEMO · 起手式模板", { x: MARGIN, y: 0.9, w: 6, h: 0.4, color: C.blue, bold: true, fontSize: 12, fontFace: FONT_NUM, charSpacing: 1, margin: 0 });
    s.addText("一句话主标题", { x: MARGIN, y: 1.7, w: 12, h: 1.0, color: C.ink, fontSize: 52, bold: true, fontFace: FONT, margin: 0 });
    s.addText("副标题:这份模板演示四种最常用的版式模式。", { x: MARGIN, y: 2.9, w: 12, h: 0.5, color: C.textMid, fontSize: 16, fontFace: FONT, margin: 0 });
    // 关键数字 hero 带(满宽 0.5~12.8)
    const by = 4.4;
    panel(s, MARGIN, by, PW - 1.0, 1.7, C.blue);
    const nums = [["98%", "覆盖率"], ["3×", "提速"], ["0", "事故"], ["5", "接口"]];
    const r = row(4, 0.18, 0.8, RIGHT - 0.3);
    for (let i = 0; i < nums.length; i++) {
      const x = r.x(i);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: by + 0.55, w: r.w, h: 0.95, rectRadius: 0.07, fill: { color: "FFFFFF", transparency: 85 }, line: { color: "FFFFFF", width: 0.75, transparency: 45 } });
      s.addText([{ text: nums[i][0], options: { color: C.white, bold: true, fontSize: 32, fontFace: FONT_NUM } }, { text: "  " + nums[i][1], options: { color: C.white, fontSize: 13, fontFace: FONT } }], { x: x + 0.2, y: by + 0.65, w: r.w - 0.3, h: 0.75, margin: 0, valign: "middle" });
    }
  }

  // ===== 模式 B:能力栈 / 卡片网格(tabCard + row) =====
  {
    const s = slide("01", "总体框架", 2, TOTAL);
    await addPageTitle(s, FA.FaThLarge, "示例:N 个能力做成一行卡片", "用 row(n) 均分,整行自动收在 0.5~12.8。每张卡 = 彩色卡头 + 白卡身。");
    const caps = [
      { num: "①", ic: FA.FaFilter, t: "能力一", d: "一句话说清这格做什么" },
      { num: "②", ic: FA.FaLayerGroup, t: "能力二", d: "一句话说清这格做什么", hl: true },
      { num: "③", ic: FA.FaShieldAlt, t: "能力三", d: "一句话说清这格做什么" },
      { num: "④", ic: FA.FaPlug, t: "能力四", d: "一句话说清这格做什么" },
    ];
    const r = row(caps.length, 0.2);
    for (let i = 0; i < caps.length; i++) {
      const c = caps[i], x = r.x(i), hf = c.hl ? C.teal : C.blue;
      await tabCard(s, x, 2.3, r.w, 2.6, 0.5, hf,
        [{ text: c.num + "  ", options: { color: C.white, bold: true, fontSize: 16, fontFace: FONT_NUM } }, { text: c.t, options: { color: C.white, bold: true, fontSize: 14, fontFace: FONT } }],
        null, c.ic);
      s.addText(c.d, { x: x + 0.18, y: 3.1, w: r.w - 0.3, h: 1.6, color: C.text, fontSize: 12, fontFace: FONT, margin: 0, valign: "top" });
    }
    panel(s, MARGIN, 5.2, PW - 1.0, 0.7, C.blue);
    await chipOnPanel(s, FA.FaStream, 0.72, 5.34, 0.42, IK.blue);
    s.addText("关键结论面板:中蓝实底 + 白字,一页一个,放在底部收口。", { x: 1.3, y: 5.2, w: 11.4, h: 0.7, color: C.white, fontSize: 14, fontFace: FONT, valign: "middle", margin: 0 });
  }

  // ===== 模式 C:两栏对比(tabCard 左右) =====
  {
    const s = slide("02", "对比", 3, TOTAL);
    await addPageTitle(s, FA.FaBalanceScale, "示例:两栏对比", "两栏统一用 row(2) 或固定 0.5 / 6.8 两个 x,右缘收 12.8。");
    const r = row(2, 0.3);
    await tabCard(s, r.x(0), 2.2, r.w, 2.6, 0.5, C.textMute, [{ text: "✕  做法 A(劣)", options: { color: C.white, bold: true, fontSize: 15, fontFace: FONT } }]);
    s.addText([{ text: "缺点一", options: { color: C.textMid, fontSize: 13, bullet: true, breakLine: true, fontFace: FONT } }, { text: "缺点二", options: { color: C.textMid, fontSize: 13, bullet: true, fontFace: FONT } }], { x: r.x(0) + 0.22, y: 2.9, w: r.w - 0.5, h: 1.7, margin: 0, valign: "top", paraSpaceAfter: 6 });
    await tabCard(s, r.x(1), 2.2, r.w, 2.6, 0.5, C.blue, [{ text: "✓  做法 B(优)", options: { color: C.white, bold: true, fontSize: 15, fontFace: FONT } }]);
    s.addText([{ text: "优点一", options: { color: C.text, fontSize: 13, bullet: true, breakLine: true, fontFace: FONT } }, { text: "优点二", options: { color: C.text, fontSize: 13, bullet: true, fontFace: FONT } }], { x: r.x(1) + 0.22, y: 2.9, w: r.w - 0.5, h: 1.7, margin: 0, valign: "top", paraSpaceAfter: 6 });
  }

  // ===== 模式 D:大数字 + 小节标 =====
  {
    const s = slide("03", "成效", 4, TOTAL);
    await addPageTitle(s, FA.FaChartLine, "示例:量化成效用大数字卡", "statCard 自带大数字 + 单位 + 说明;一行用 row(n) 均分。");
    sectionLabel(s, MARGIN, 2.0, 6, "核心指标", C.blue);
    const r = row(4, 0.2);
    const stats = [["60", "实体", "结构化对象", C.blue], ["67", "关系", "知识网的边", C.teal], ["0/0", "lint", "零缺陷", C.blue], ["95%", "回溯", "带来源", C.teal]];
    for (let i = 0; i < stats.length; i++) statCard(s, r.x(i), 2.45, r.w, 1.5, stats[i][0], stats[i][1], stats[i][2], stats[i][3]);
    card(s, MARGIN, 4.3, PW - 1.0, 1.0, { fill: C.bluePale, border: C.bluePale });
    s.addText("淡蓝底面板:放补充说明 / 边界声明 / 一句话提醒。", { x: 0.8, y: 4.3, w: 11.7, h: 1.0, color: C.ink, fontSize: 13, fontFace: FONT, valign: "middle", margin: 0 });
  }

  // ===== 模式 E:原生可编辑图表(softChart)=====
  // 领导端可直接改数据;定制标注多的图(高亮带/参考线/平台期)仍手搭。
  {
    const s = slide("04", "数据图表", 5, TOTAL);
    await addPageTitle(s, FA.FaChartBar, "示例:原生可编辑图表用 softChart", "柔光蓝预设:主蓝/柔青系列、圆角图区、弱轴标、细网格、数据标签 12pt。");
    const r = row(2, 0.4);
    sectionLabel(s, r.x(0), 2.1, 5, "柱状:季度产值(亿元)", C.blue);
    softChart(s, "bar", [
      { name: "产值", labels: ["Q1", "Q2", "Q3", "Q4"], values: [4.2, 5.1, 6.3, 7.4] },
    ], r.x(0), 2.5, r.w, 3.8, { dataLabelFormatCode: "0.0" });
    sectionLabel(s, r.x(1), 2.1, 5, "折线:负荷率走势(%)", C.teal);
    softChart(s, "line", [
      { name: "负荷率", labels: ["7月", "8月", "9月", "10月", "11月", "12月"], values: [97, 100, 144, 144, 156, 129] },
    ], r.x(1), 2.5, r.w, 3.8, { chartColors: [C.teal] });
  }

  await pres.writeFile({ fileName: "example-deck.pptx" });
  console.log("✓ example-deck.pptx 生成完成");
})();
