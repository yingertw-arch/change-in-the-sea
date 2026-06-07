const docx = require('C:/Users/User/AppData/Roaming/npm/node_modules/docx');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, ShadingType, VerticalAlign
} = docx;
const fs = require('fs');

const FONT = "微軟正黑體";
const CW = 9026;                 // A4 content width (DXA)
const INK = "1B3A4B", BLUE = "1E6091", CORAL = "C0612E", GREEN = "2D8A5F";
const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: "9DB8C7" };
const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };
const cellMargins = { top: 70, bottom: 70, left: 110, right: 110 };

function run(t, o = {}) { return new TextRun({ text: t, font: FONT, size: o.size || 22, bold: o.bold || false, color: o.color || INK }); }
function p(children, o = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    spacing: { before: o.before ?? 40, after: o.after ?? 40, line: o.line ?? 300 },
    alignment: o.align, ...(o.indent ? { indent: o.indent } : {})
  });
}
// section heading bar
function heading(num, title) {
  return new Paragraph({
    spacing: { before: 220, after: 90 },
    shading: { type: ShadingType.CLEAR, fill: "D5E8F0" },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: BLUE, space: 6 } },
    children: [ new TextRun({ text: `  ${num}  ${title}`, font: FONT, size: 24, bold: true, color: BLUE }) ]
  });
}
// answer line (underline-style blank)
function aline(n = 1) {
  const arr = [];
  for (let i = 0; i < n; i++) arr.push(new Paragraph({
    spacing: { before: 60, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "B0BEC5", space: 2 } },
    children: [ new TextRun({ text: "", font: FONT }) ]
  }));
  return arr;
}
function cell(content, o = {}) {
  const kids = Array.isArray(content) ? content : [ p(run(content, { bold: o.bold, color: o.color }), { align: o.align, after: 20, before: 20 }) ];
  return new TableCell({
    width: { size: o.w, type: WidthType.DXA }, borders, margins: cellMargins,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill } : undefined,
    verticalAlign: VerticalAlign.CENTER, children: kids
  });
}
function table(widths, rows) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map(r => new TableRow({ children: r }))
  });
}
function blankRows(widths, n, h = 1) {
  const rows = [];
  for (let i = 0; i < n; i++) rows.push(widths.map(w => cell([ new Paragraph({ spacing: { before: 90 * h, after: 90 * h }, children: [run("", {})] }) ], { w })));
  return rows;
}

const kids = [];

// ===== 標題 =====
kids.push(new Paragraph({
  spacing: { after: 20 }, alignment: AlignmentType.CENTER,
  children: [ new TextRun({ text: "🌊 改變中的海", font: FONT, size: 40, bold: true, color: BLUE }) ]
}));
kids.push(new Paragraph({
  spacing: { after: 140 }, alignment: AlignmentType.CENTER,
  children: [ new TextRun({ text: "大自然的雕刻　×　人類的足跡　｜　永續發展學習單", font: FONT, size: 22, color: CORAL, bold: true }) ]
}));

// ===== 基本資料 =====
const w4 = [1400, 3113, 1400, 3113];
kids.push(table(w4, [[
  cell("班級", { w: w4[0], bold: true, fill: "EAF3F8", align: AlignmentType.CENTER }),
  cell("", { w: w4[1] }),
  cell("座號", { w: w4[2], bold: true, fill: "EAF3F8", align: AlignmentType.CENTER }),
  cell("", { w: w4[3] }),
], [
  cell("姓名", { w: w4[0], bold: true, fill: "EAF3F8", align: AlignmentType.CENTER }),
  cell("", { w: w4[1] }),
  cell("日期", { w: w4[2], bold: true, fill: "EAF3F8", align: AlignmentType.CENTER }),
  cell("", { w: w4[3] }),
]]));

// ===== 看片任務 =====
kids.push(heading("📺", "看片任務（搭配愛學網影片）"));
kids.push(p([
  run("影片① ", { bold: true, color: BLUE }), run("《海灘偵察記》（海洋的危機）　"),
  run("影片② ", { bold: true, color: BLUE }), run("《岩石的風化作用》（女王頭怎麼形成）"),
]));
kids.push(p(run("看影片時，找出下面的答案：", { bold: true })));
kids.push(p(run("1. 寄居蟹找不到貝殼，只好拿什麼當家？ ________________"), { indent: { left: 200 } }));
kids.push(p(run("2. 海灘的垃圾是被什麼帶來的？ ________________"), { indent: { left: 200 } }));
kids.push(p(run("3. 什麼是「風化作用」？它怎麼雕出女王頭？ ________________"), { indent: { left: 200 } }));

// ===== 任務一：抽抽樂的祕密 =====
kids.push(heading("任務一", "🎲 抽抽樂的祕密"));
kids.push(p(run("回想兩次抽抽樂遊戲，把它填完整：")));
{
  const w = [2426, 3300, 3300];
  kids.push(table(w, [
    [ cell("", { w: w[0], fill: "EAF3F8" }), cell("🪨 野柳抽抽樂", { w: w[1], bold: true, fill: "F1E8D5", align: AlignmentType.CENTER }), cell("🐠 海洋抽抽樂", { w: w[2], bold: true, fill: "D9EEF4", align: AlignmentType.CENTER }) ],
    [ cell("抽掉一塊，代表發生什麼事？", { w: w[0], bold: true, fill: "FAFAFA" }), cell("", { w: w[1] }), cell("", { w: w[2] }) ],
    [ cell("是「大自然」還是「人類」造成的？", { w: w[0], bold: true, fill: "FAFAFA" }), cell("", { w: w[1] }), cell("", { w: w[2] }) ],
    [ cell("結果怎麼了？（女王頭／生態系）", { w: w[0], bold: true, fill: "FAFAFA" }), cell("", { w: w[1] }), cell("", { w: w[2] }) ],
  ]));
}
kids.push(p([ run("💡 我的發現：", { bold: true, color: GREEN }), run("同樣是東西被一塊塊拿走，野柳是 ________ 造成的，海洋是 ________ 造成的。") ]));

// ===== 任務二：兩種改變比一比 =====
kids.push(heading("任務二", "🔍 兩種改變，比一比"));
kids.push(p(run("把你兩次活動親眼看到、親耳聽到的，填進這張表：")));
{
  const w = [2426, 3300, 3300];
  const rowLabels = ["① 是什麼力量造成改變？", "② 改變的速度（快／慢）？", "③ 大約花多久時間？", "④ 這個改變能不能挽回？", "⑤ 我能為它做什麼？"];
  const rows = [[ cell("比較項目", { w: w[0], bold: true, fill: "EAF3F8", align: AlignmentType.CENTER }), cell("🪨 野柳女王頭", { w: w[1], bold: true, fill: "F1E8D5", align: AlignmentType.CENTER }), cell("🐠 我浮潛的海洋", { w: w[2], bold: true, fill: "D9EEF4", align: AlignmentType.CENTER }) ]];
  rowLabels.forEach(l => rows.push([ cell(l, { w: w[0], bold: true, fill: "FAFAFA" }), cell("", { w: w[1] }), cell("", { w: w[2] }) ]));
  kids.push(table(w, rows));
}

// ===== 任務三：關鍵提問 =====
kids.push(heading("任務三", "✍️ 關鍵提問"));
kids.push(p([ run("「大自然用兩千萬年慢慢改變海岸，人類卻用幾十年快速改變海洋。」", { bold: true, color: BLUE }) ]));
kids.push(p(run("這兩種改變，你覺得哪一種是我們該負責、也能改變的？為什麼？")));
aline(2).forEach(l => kids.push(l));

// ===== 任務四：海與岩的循環 =====
kids.push(heading("任務四", "🔄 海與岩的循環"));
kids.push(p(run("1. 野柳的岩石裡找得到海洋生物的化石！這說明野柳這塊陸地以前其實是 ________________。")));
kids.push(p(run("2. 我浮潛看到的活珊瑚、貝殼，幾百萬年後有可能變成什麼？")));
kids.push(p(run("　 □ 沙子　　□ 岩石／化石　　□ 永遠消失　　我的理由：________________"), { indent: { left: 200 } }));
kids.push(p(run("3. 把循環補完整：")));
kids.push(p([ run("　 海裡的生命 🐠 →（死亡沉積）→ ________ →（變硬成岩）→ 🪨 →（海水侵蝕）→ 回到 ________", { }) ], { indent: { left: 200 } }));
kids.push(p([ run("💡 原來海和陸、生命和岩石，是同一個地球在循環。", { bold: true, color: GREEN }) ]));

// ===== 任務五：我是改變的力量 =====
kids.push(heading("任務五", "🌱 我是改變的力量"));
kids.push(p(run("有些改變我們只能珍惜（女王頭終會風化）；有些改變我們可以阻止（海洋污染與暖化）。")));
kids.push(p([ run("永續，就是分辨清楚，然後為「能改變的」採取行動。", { bold: true, color: CORAL }) ]));
kids.push(p(run("我為海洋做得到的三個具體行動：", { bold: true })));
kids.push(p(run("🛒 ____________________________________________")));
kids.push(p(run("♻️ ____________________________________________")));
kids.push(p(run("📣 ____________________________________________")));
kids.push(p(run("✍️ 我的永續宣言（用一句話總結今天的學習）：", { bold: true })));
aline(2).forEach(l => kids.push(l));

// ===== 自我評量 =====
kids.push(heading("⭐", "自我評量"));
{
  const w = [5126, 1300, 1300, 1300];
  const items = [
    "我能說出「自然改變」和「人為改變」的不同",
    "我懂得海與岩石其實是一個循環",
    "我能分辨哪些改變我能行動",
    "我願意付出具體的永續行動",
  ];
  const rows = [[ cell("我做到了…", { w: w[0], bold: true, fill: "EAF3F8" }), cell("😀 很棒", { w: w[1], bold: true, fill: "EAF3F8", align: AlignmentType.CENTER }), cell("🙂 還可以", { w: w[2], bold: true, fill: "EAF3F8", align: AlignmentType.CENTER }), cell("😐 加油", { w: w[3], bold: true, fill: "EAF3F8", align: AlignmentType.CENTER }) ]];
  items.forEach(it => rows.push([ cell(it, { w: w[0] }), cell("", { w: w[1] }), cell("", { w: w[2] }), cell("", { w: w[3] }) ]));
  kids.push(table(w, rows));
}

// ===== 教師對應參考 =====
kids.push(new Paragraph({
  spacing: { before: 200, after: 40 },
  border: { top: { style: BorderStyle.DASHED, size: 4, color: "B0BEC5", space: 6 } },
  children: [ new TextRun({ text: "📋 教師對應參考（不印給學生）", font: FONT, size: 18, bold: true, color: "78909C" }) ]
}));
[
  "議題融入：海洋教育＋環境教育（氣候變遷、永續發展）",
  "核心素養：自-E-A2 探究與推理、自-E-C1 珍視自然與永續實踐、自-E-A3 規劃與行動",
  "SDGs：12 責任消費、13 氣候行動、14 海洋生態、15 陸域生態",
  "評量重點：任務二「可逆性」對比＝永續素養核心；任務四循環概念為加分亮點",
  "搭配數位教材：互動闖關網頁《改變中的海》（看影片→闖關→永續宣言）",
].forEach(t => kids.push(new Paragraph({ spacing: { before: 20, after: 20 }, bullet: { level: 0 }, children: [ new TextRun({ text: t, font: FONT, size: 18, color: "607D8B" }) ] })));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, right: 1440, bottom: 1300, left: 1440 } } },
    children: kids
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("改變中的海_橋樑學習單.docx", buf);
  console.log("OK 已輸出 改變中的海_橋樑學習單.docx (" + buf.length + " bytes)");
});
