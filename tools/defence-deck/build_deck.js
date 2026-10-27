// Builds docs/Nisha_MS_Thesis_Defence.pptx
// Usage: node build_deck.js <out.pptx> [--qa]
//   --qa fills every picture placeholder with a grey box so the layout can be checked in a render.
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip");
const fs = require("fs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const OUT = process.argv[2] || "Nisha_MS_Thesis_Defence.pptx";
const QA = process.argv.includes("--qa");

// ---------- Design tokens ----------
const C = {
  green: "1E4D3A", bar: "2E8B57", gold: "8A6A1F", ink: "1F2A2E", muted: "4A5550",
  tint: "EFF4F1", footer: "DDEFE3", line: "D3DED7", white: "FFFFFF", onDark: "C9D8CF", stripe: "F7FAF8",
};
const HF = "Cambria", BF = "Calibri";
const X0 = 0.6, CW = 12.13; // left margin, content width (LAYOUT_WIDE is 13.333 x 7.5)

// Scientific names are set in italics; "sp." / "spp." / "L." stay roman.
const TAXA = /(Cercospora malayensis|C\. malayensis|Hibiscus sabdariffa|Trichoderma asperellum|Trichoderma harzianum|Trichoderma viride|T\. asperellum|T\. harzianum|T\. viride|Cercospora|Corynespora|Coniella|Trichoderma|Hibiscus)/;
function R(text, o = {}) {
  return String(text).split(new RegExp(TAXA.source, "g")).filter((s) => s !== "")
    .map((s) => ({ text: s, options: { ...o, ...(TAXA.test(s) && new RegExp(`^${TAXA.source}$`).test(s) ? { italic: !o.italic } : {}) } }));
}
// Paragraph list -> run array. Paragraph options go on the first run, breakLine on the last.
function paras(list, runOpts = {}, paraOpts = {}) {
  const out = [];
  list.forEach((item, i) => {
    const runs = typeof item === "string" ? R(item, runOpts) : item;
    runs[0].options = { ...paraOpts, ...runs[0].options };
    if (i < list.length - 1) runs[runs.length - 1].options = { ...runs[runs.length - 1].options, breakLine: true };
    out.push(...runs);
  });
  return out;
}
const shadow = () => ({ type: "outer", color: "000000", blur: 5, offset: 1.5, angle: 90, opacity: 0.1 });

async function iconPng(name, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: "256" }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}

// ---------- Masters (one per picture layout) ----------
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Nisha Mazumder";
pres.company = "Department of Plant Pathology, Bangladesh Agricultural University";
pres.title = "MS Thesis Defence: Cercospora leaf spot of roselle";

const BAND = { rect: { x: 0, y: 7.05, w: 13.333, h: 0.45, fill: { color: C.footer }, line: { color: C.footer } } };
const FOOTER = [
  { text: { text: "Nisha Mazumder  ·  MS Thesis Defence  ·  Department of Plant Pathology, BAU", options: { x: X0, y: 7.12, w: 9, h: 0.3, fontFace: BF, fontSize: 10.5, color: C.green, valign: "middle", margin: 0 } } },
];
const SLIDE_NO = { x: 12.13, y: 7.12, w: 0.6, h: 0.3, fontFace: BF, fontSize: 10.5, color: C.green, align: "right", valign: "middle" };
const masters = {};
function master(name, pics = [], { bg = C.white, footer = true } = {}) {
  masters[name] = pics.map((p, i) => ({ ...p, name: `${name}_pic${i + 1}` }));
  pres.defineSlideMaster({
    title: name,
    background: { color: C.white },
    objects: [
      BAND,
      ...(footer ? FOOTER : []),
      ...masters[name].map((p) => ({ placeholder: { options: { name: p.name, type: "pic", x: p.x, y: p.y, w: p.w, h: p.h, fontFace: BF, fontSize: 11, color: C.muted, align: "center", valign: "middle" }, text: p.prompt || "Click the icon to insert a picture" } })),
    ],
    ...(footer ? { slideNumber: SLIDE_NO } : {}),
  });
}
function newSlide(name, under) {
  const s = pres.addSlide({ masterName: name });
  if (under) under(s);
  for (const p of masters[name] || []) {
    s.addText("", { placeholder: p.name });
    if (QA) s.addText("photo", { x: p.x, y: p.y, w: p.w, h: p.h, fill: { color: "D9DFDB" }, fontFace: BF, fontSize: 11, color: C.muted, align: "center", valign: "middle", isTextBox: true });
  }
  return s;
}

// Picture geometry shared by layouts
const ROW5 = [0, 1, 2, 3, 4].map((i) => ({ x: X0 + i * 2.47, y: 1.62, w: 2.25, h: 2.3 }));
const GRID4 = [[0, 0], [1, 0], [0, 1], [1, 1]].map(([c, r]) => ({ x: X0 + c * 3.17, y: 1.62 + r * 2.42, w: 3.0, h: 1.95 }));
const ROW4 = [0, 1, 2, 3].map((i) => ({ x: X0 + i * 3.08, y: 1.62, w: 2.87, h: 2.87 }));
const BIG1 = [{ x: X0, y: 1.62, w: 7.1, h: 4.55 }];
const PAIR = [{ x: X0, y: 1.62, w: 3.45, h: 4.1 }, { x: X0 + 3.65, y: 1.62, w: 3.45, h: 4.1 }];
const PAIR_R = [{ x: 6.15, y: 1.62, w: 3.2, h: 4.1 }, { x: 9.53, y: 1.62, w: 3.2, h: 4.1 }];

master("TITLE", [{ x: 6.12, y: 0.45, w: 1.1, h: 1.1, prompt: "BAU logo" }, { x: 0.8, y: 4.72, w: 1.25, h: 1.62, prompt: "Photo" }], { footer: false });
master("DIVIDER", [], { footer: false });
master("CONTENT");
master("ROW5", ROW5);
master("GRID4", GRID4);
master("ROW4", ROW4);
master("BIG1", BIG1);
master("PAIR", PAIR);
master("PAIR_R", PAIR_R);

// ---------- Building blocks ----------
function header(s, kicker, title) {
  s.addText(kicker.toUpperCase(), { x: X0, y: 0.36, w: CW, h: 0.28, fontFace: BF, fontSize: 11, bold: true, color: C.gold, charSpacing: 2, margin: 0, isTextBox: true });
  s.addText(R(title), { x: X0, y: 0.64, w: CW, h: 0.72, fontFace: HF, fontSize: 28, bold: true, color: C.green, margin: 0, valign: "middle", isTextBox: true });
}
function caption(s, label, text, x, y, w, align = "left") {
  s.addText([{ text: label + "  ", options: { bold: true, color: C.ink } }, ...R(text, { color: C.muted })], { x, y, w, h: 0.32, fontFace: BF, fontSize: 11, margin: 0, align, valign: "top", isTextBox: true });
}
function picLabels(s, rects, labels, dy = 0.06) {
  rects.forEach((r, i) => s.addText(R(labels[i]), { x: r.x, y: r.y + r.h + dy, w: r.w, h: 0.3, fontFace: BF, fontSize: 12, color: C.ink, align: "center", margin: 0, isTextBox: true }));
}
function panel(s, x, y, w, h) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.tint }, line: { color: C.tint } });
}
function bullets(s, list, x, y, w, h, size = 15, extra = {}) {
  s.addText(paras(list, { fontFace: BF, fontSize: size, color: C.ink }, { bullet: { indent: 16 }, paraSpaceAfter: 7 }), { x, y, w, h, valign: "top", margin: [4, 6, 4, 4], isTextBox: true, ...extra });
}
function numBadge(s, n, x, y, d = 0.42, fill = C.green, color = C.white) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
  s.addText(String(n), { x, y, w: d, h: d, fontFace: BF, fontSize: d > 0.5 ? 18 : 13, bold: true, color, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
async function iconCircle(s, icon, x, y, d = 0.62, fill = C.green) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
  const p = d * 0.25;
  s.addImage({ data: await iconPng(icon, C.white), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}
function keyFinding(s, text, y = 6.12) {
  panel(s, X0, y, CW, 0.58);
  s.addText([{ text: "Key finding:  ", options: { bold: true, color: C.green } }, ...R(text, { color: C.ink })], { x: X0 + 0.2, y, w: CW - 0.4, h: 0.58, fontFace: BF, fontSize: 13.5, valign: "middle", margin: 0, isTextBox: true });
}
async function divider(num, title, sub) {
  const s = newSlide("DIVIDER");
  s.addImage({ data: await iconPng("FaLeaf", C.green), x: 8.7, y: 1.65, w: 4.2, h: 4.2, transparency: 88 });
  s.addText(num, { x: 0.9, y: 2.0, w: 3, h: 1.0, fontFace: HF, fontSize: 66, bold: true, color: C.gold, margin: 0, isTextBox: true });
  s.addText(title, { x: 0.9, y: 3.1, w: 8, h: 0.9, fontFace: HF, fontSize: 40, bold: true, color: C.green, margin: 0, isTextBox: true });
  s.addText(sub, { x: 0.9, y: 4.05, w: 8, h: 0.5, fontFace: BF, fontSize: 17, color: C.muted, margin: 0, isTextBox: true });
}

// Results tables (thesis Tables 5-14)
const TR = ["T0 = Control", "T1 = Nativo 500 ppm", "T2 = Tilt 500 ppm", "T3 = Folicur 500 ppm", "T4 = T. harzianum", "T5 = T. viride", "T6 = T. asperellum"];
function dataTable(s, cols, rows, se, cv, y = 1.98) {
  const n = cols.length;
  const first = 3.0, rest = (CW - first) / (n - 1);
  const head = cols.map((c, i) => ({ text: c, options: { bold: true, color: C.white, fill: { color: C.green }, align: i ? "center" : "left" } }));
  const body = rows.map((r, ri) => [
    { text: R(TR[ri]), options: { fill: { color: ri % 2 ? C.stripe : C.white }, align: "left" } },
    ...r.map((v) => ({ text: v, options: { fill: { color: ri % 2 ? C.stripe : C.white }, align: "center" } })),
  ]);
  const stat = (label, vals) => [{ text: label, options: { bold: true, color: C.muted, fill: { color: C.tint } } }, ...vals.map((v) => ({ text: String(v), options: { color: C.muted, fill: { color: C.tint }, align: "center" } }))];
  s.addTable([head, ...body, stat("SE", se), stat("CV", cv)], {
    x: X0, y, w: CW, colW: [first, ...Array(n - 1).fill(rest)], rowH: 0.33,
    fontFace: BF, fontSize: 12.5, color: C.ink, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1],
    border: { type: "solid", pt: 0.5, color: C.line },
  });
}
const FOOTNOTE = "Means followed by the same letter(s) in a column did not differ significantly at the 5% level (Tukey’s HSD).";
function tableSlide(kicker, title, tableNo, tableCap, cols, rows, se, cv, finding) {
  const s = newSlide("CONTENT");
  header(s, kicker, title);
  caption(s, `Table ${tableNo}.`, tableCap, X0, 1.52, CW);
  dataTable(s, cols, rows, se, cv);
  s.addText(FOOTNOTE, { x: X0, y: 5.68, w: CW, h: 0.3, fontFace: BF, fontSize: 10.5, italic: true, color: C.muted, margin: 0, isTextBox: true });
  keyFinding(s, finding);
  return s;
}

(async () => {
  // ===== 1. Title =====
  let s = newSlide("TITLE", (t) => panel(t, X0, 4.55, 3.95, 1.95));
  s.addText("MS THESIS DEFENCE", { x: X0, y: 1.78, w: CW, h: 0.3, fontFace: BF, fontSize: 12, bold: true, color: C.gold, charSpacing: 3, align: "center", margin: 0, isTextBox: true });
  s.addText("Department of Plant Pathology  ·  Bangladesh Agricultural University, Mymensingh", { x: X0, y: 2.08, w: CW, h: 0.3, fontFace: BF, fontSize: 13, color: C.muted, align: "center", margin: 0, isTextBox: true });
  s.addText(R("IDENTIFICATION AND CHARACTERIZATION OF Cercospora sp. CAUSING LEAF SPOT ON ROSELLE (Hibiscus sabdariffa L.) AND ITS MANAGEMENT"), { x: 1.2, y: 2.5, w: 10.93, h: 1.6, fontFace: HF, fontSize: 27, bold: true, color: C.green, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText([
    { text: "PRESENTED BY", options: { fontSize: 10, bold: true, color: C.gold, charSpacing: 2, breakLine: true } },
    { text: "Nisha Mazumder", options: { fontFace: HF, fontSize: 18, bold: true, color: C.green, breakLine: true } },
    { text: "ID No.: 24120508", options: { fontSize: 12, color: C.ink, breakLine: true } },
    { text: "Reg. No.: 49427", options: { fontSize: 12, color: C.ink, breakLine: true } },
    { text: "Session: 2018–19", options: { fontSize: 12, color: C.ink } },
  ], { x: 2.22, y: 4.72, w: 2.25, h: 1.62, fontFace: BF, valign: "middle", margin: 0, paraSpaceAfter: 2, isTextBox: true });
  const sup = [
    ["SUPERVISOR", "Prof. Dr. Mohammad Delwar Hossain", ["Professor, Department of Plant Pathology", "Bangladesh Agricultural University", "Mymensingh-2202"], 4.75, 3.9],
    ["CO-SUPERVISOR", "Prof. Dr. Islam Hamim", ["Professor and Head, Department of Plant Pathology", "Bangladesh Agricultural University", "Mymensingh-2202"], 8.83, 3.9],
  ];
  for (const [k, n, lines, x, w] of sup) {
    panel(s, x, 4.55, w, 1.95);
    s.addText([
      { text: k, options: { fontSize: 10, bold: true, color: C.gold, charSpacing: 2, breakLine: true } },
      { text: n, options: { fontFace: HF, fontSize: 15, bold: true, color: C.green, breakLine: true } },
      ...lines.map((l, i) => ({ text: l, options: { fontSize: 11.5, color: C.ink, breakLine: i < lines.length - 1 } })),
    ], { x: x + 0.22, y: 4.72, w: w - 0.4, h: 1.62, fontFace: BF, valign: "middle", margin: 0, paraSpaceAfter: 2, isTextBox: true });
  }

  // ===== 2. Outline =====
  s = newSlide("CONTENT");
  header(s, "Presentation outline", "Outline");
  const outline = [
    ["01", "Introduction", "Background, significance of the study and objectives"],
    ["02", "Materials and Methods", "Laboratory, pot and field experiments"],
    ["03", "Results", "Identification, in vitro evaluation, pot and field evaluation"],
    ["04", "Discussion and Conclusions", "Discussion, conclusions and recommendations"],
  ];
  outline.forEach(([n, t, d], i) => {
    const y = 1.75 + i * 1.18;
    s.addText(n, { x: X0 + 0.2, y, w: 1.1, h: 0.9, fontFace: HF, fontSize: 40, bold: true, color: C.gold, margin: 0, valign: "middle", isTextBox: true });
    s.addText([{ text: t, options: { fontFace: HF, fontSize: 22, bold: true, color: C.green, breakLine: true } }, { text: d, options: { fontFace: BF, fontSize: 14, color: C.muted } }], { x: X0 + 1.45, y, w: 9, h: 0.9, valign: "middle", margin: 0, isTextBox: true });
    if (i < 3) s.addShape(pres.shapes.LINE, { x: X0 + 1.45, y: y + 1.04, w: 9.5, h: 0, line: { color: C.line, width: 0.75 } });
  });

  // ===== Introduction =====
  await divider("01", "Introduction", "Background  ·  Significance  ·  Objectives");

  s = newSlide("CONTENT");
  header(s, "Introduction", "Roselle: a multipurpose crop");
  const intro = [
    ["FaSeedling", "The crop", "Roselle (Hibiscus sabdariffa L.) is a multipurpose annual crop of the family Malvaceae, locally known as ‘Chukur’ or ‘Mesta’ (Islam, 2019)."],
    ["FaUtensils", "Its uses", "Grown for bast fibre, calyces (tea, jam and syrup), edible leaves and seed oil (Chew et al., 2024)."],
    ["FaHeartbeat", "Its value", "Calyces are rich in anthocyanins and phenolics with antihypertensive, antioxidant and antimicrobial activity (Da-Costa-Rocha et al., 2014)."],
  ];
  for (let i = 0; i < 3; i++) {
    const x = X0 + i * 4.13;
    panel(s, x, 1.75, 3.87, 4.6);
    await iconCircle(s, intro[i][0], x + 0.35, 2.1, 0.8);
    s.addText(intro[i][1], { x: x + 0.35, y: 3.1, w: 3.2, h: 0.45, fontFace: HF, fontSize: 20, bold: true, color: C.green, margin: 0, isTextBox: true });
    s.addText(R(intro[i][2]), { x: x + 0.35, y: 3.65, w: 3.2, h: 2.5, fontFace: BF, fontSize: 15.5, color: C.ink, margin: 0, valign: "top", paraSpaceAfter: 4, isTextBox: true });
  }

  s = newSlide("CONTENT");
  header(s, "Introduction", "Significance of the study");
  panel(s, X0, 1.75, 7.35, 4.6);
  bullets(s, [
    "Foliar diseases reduce photosynthetic leaf area, calyx number and fibre quality.",
    "Leaf spot fungi of different genera (Cercospora, Corynespora, Coniella) produce near-identical lesions, so symptoms alone cannot confirm the cause.",
    "Cercospora leaf spots occur on 17 Hibiscus species and can cause complete crop loss (Park et al., 2017).",
    "A leaf spot of uncertain cause appeared in roselle fields of Mymensingh and Tangail.",
  ], X0 + 0.25, 2.05, 6.85, 4.2, 18, { paraSpaceAfter: 18 });
  s.addShape(pres.shapes.RECTANGLE, { x: 8.25, y: 1.75, w: 4.48, h: 2.2, fill: { color: C.green }, line: { color: C.green } });
  s.addText("~100%", { x: 8.55, y: 1.95, w: 3.9, h: 0.8, fontFace: HF, fontSize: 40, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText(R("incidence of Corynespora leaf and calyx spot of roselle in Mexico (Ortega-Acosta et al., 2016)", { color: C.onDark }), { x: 8.55, y: 2.8, w: 3.9, h: 0.95, fontFace: BF, fontSize: 13.5, margin: 0, valign: "top", isTextBox: true });
  s.addShape(pres.shapes.RECTANGLE, { x: 8.25, y: 4.15, w: 4.48, h: 2.2, fill: { color: C.white }, line: { color: C.green, width: 1.5 } });
  s.addText("No prior report", { x: 8.55, y: 4.35, w: 3.9, h: 0.7, fontFace: HF, fontSize: 26, bold: true, color: C.green, margin: 0, isTextBox: true });
  s.addText(R("of Cercospora leaf spot of roselle, its identity or its management in Bangladesh"), { x: 8.55, y: 5.1, w: 3.9, h: 1.0, fontFace: BF, fontSize: 13.5, color: C.ink, margin: 0, valign: "top", isTextBox: true });

  s = newSlide("CONTENT");
  header(s, "Introduction", "Objectives");
  s.addText("The present study was undertaken with the following objectives:", { x: X0, y: 1.65, w: CW, h: 0.4, fontFace: BF, fontSize: 16, italic: true, color: C.muted, margin: 0, isTextBox: true });
  const obj = [
    "To isolate and identify the causal fungus of leaf spot disease of roselle on the basis of its cultural and morphological characteristics.",
    "To evaluate the efficacy of Trichoderma spp. (T. asperellum, T. harzianum and T. viride) and selected fungicides in the management of Cercospora leaf spot of roselle.",
  ];
  obj.forEach((t, i) => {
    const y = 2.35 + i * 1.95;
    panel(s, X0, y, CW, 1.65);
    numBadge(s, i + 1, X0 + 0.35, y + 0.47, 0.72, C.green);
    s.addText(R(t), { x: X0 + 1.4, y, w: CW - 1.7, h: 1.65, fontFace: BF, fontSize: 19, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
  });

  // ===== Materials and Methods =====
  await divider("02", "Materials and Methods", "Laboratory  ·  Pot experiment  ·  Field experiment");

  s = newSlide("CONTENT");
  header(s, "Materials and methods", "Experimental approach");
  const appr = [
    ["FaFlask", "Laboratory", "CRD  ·  3 replications", ["20 diseased leaves from Mymensingh and Tangail", "Isolation on PDA and single-spore purification", "Cultural and morphological identification", "Growth at 10, 15, 20, 25 and 30 °C", "Pathogenicity test (Koch’s postulates)", "Poisoned food: 4 fungicides × 5 doses", "Dual culture: 3 Trichoderma spp."]],
    ["FaWarehouse", "Pot experiment", "Net house, SPC, BAU  ·  CRD", ["Sterilized clay loam soil", "Deep red roselle variety", "7 treatments (T0–T6)", "Disease and growth data at 90 and 180 DAS"]],
    ["FaTractor", "Field experiment", "Plant Pathology field, BAU  ·  RCBD", ["Sown on 30 March 2025, line sowing", "Recommended manures and fertilizers", "Same 7 treatments", "Disease, growth and yield data at 90 and 180 DAS"]],
  ];
  for (let i = 0; i < 3; i++) {
    const x = X0 + i * 4.13;
    panel(s, x, 1.65, 3.87, 4.95);
    await iconCircle(s, appr[i][0], x + 0.3, 1.88, 0.62);
    s.addText([{ text: appr[i][1], options: { fontFace: HF, fontSize: 19, bold: true, color: C.green, breakLine: true } }, { text: appr[i][2], options: { fontFace: BF, fontSize: 12, color: C.gold, bold: true } }], { x: x + 1.08, y: 1.83, w: 2.7, h: 0.75, margin: 0, valign: "middle", isTextBox: true });
    bullets(s, appr[i][3], x + 0.2, 2.8, 3.5, 3.7, 13.5, { paraSpaceAfter: 6 });
  }

  s = newSlide("ROW5");
  header(s, "Materials and methods", "Isolation and identification of the pathogen");
  picLabels(s, ROW5, ["1 · Infected leaves collected", "2 · Surface sterilization", "3 · Plating on PDA", "4 · Single-spore pure culture", "5 · Microscopic observation"]);
  bullets(s, [
    "Infected leaves collected from Mymensingh and Tangail; surface sterilized with 70% ethanol and 1% NaOCl.",
    "Plated on PDA at 30 ± 2 °C for 7–9 days; pure culture obtained by single-spore isolation.",
    "Identified on the basis of colony characters, conidiophores and conidia (Chupp, 1954; Crous and Braun, 2003).",
  ], X0, 4.4, CW, 1.65, 14.5, { paraSpaceAfter: 5 });
  caption(s, "Illustration.", "Isolation and identification of the pathogen (Park et al., 2017)", X0, 6.25, CW);

  s = newSlide("ROW5");
  header(s, "Materials and methods", "Pathogenicity test and effect of temperature");
  picLabels(s, ROW5, ["1 · Culture on PDA", "2 · Mycelial disc cut", "3 · Disc on attached leaf", "4 · Control: sterile PDA disc", "5 · Re-isolation"]);
  bullets(s, [
    "Culture grown on PDA at 25 ± 2 °C for 7 days; mycelial disc cut with a cork borer and placed on a non-wounded attached leaf, secured with moist cotton and cellophane tape.",
    "Sterile PDA disc used as the control; pathogen re-isolated from lesions to fulfil Koch’s postulates.",
    "Effect of temperature: mycelial discs on PDA incubated at 10, 15, 20, 25 and 30 °C; radial growth measured at 10 DAI.",
  ], X0, 4.4, CW, 1.75, 14.5, { paraSpaceAfter: 5 });
  caption(s, "Illustration.", "Pathogenicity test on 2-month-old pot-grown roselle plants (3 replications)", X0, 6.25, CW);

  s = newSlide("CONTENT");
  header(s, "Materials and methods", "In vitro evaluation of fungicides and bioagents");
  const iv = [
    ["Fungicides", "Poisoned food technique", ["Nativo 75 WG (Tebuconazole 50% + Trifloxystrobin 25%)", "Tilt 250 EC (Propiconazole 25%)", "Bavistin 50 WP (Carbendazim 50%)", "Folicur 250 EC (Tebuconazole 25%)"], "@ 50, 100, 250, 500 and 1000 ppm in PDA; 5 mm pathogen disc; 30 ± 2 °C"],
    ["Bioagents", "Dual culture technique", ["Trichoderma asperellum (AB935955)", "Trichoderma harzianum (AB935952)", "Trichoderma viride (TV6)"], "5 mm discs of pathogen and antagonist on opposite sides of a PDA plate; 25 ± 2 °C until control plates were full"],
  ];
  iv.forEach(([t, m, list, note], i) => {
    const x = X0 + i * 6.17;
    panel(s, x, 1.65, 5.96, 3.75);
    s.addText([{ text: t, options: { fontFace: HF, fontSize: 20, bold: true, color: C.green, breakLine: true } }, { text: m, options: { fontFace: BF, fontSize: 13, bold: true, color: C.gold } }], { x: x + 0.3, y: 1.8, w: 5.4, h: 0.8, margin: 0, valign: "middle", isTextBox: true });
    bullets(s, list, x + 0.2, 2.65, 5.55, 1.75, 14.5, { paraSpaceAfter: 4 });
    s.addText(note, { x: x + 0.3, y: 4.45, w: 5.4, h: 0.8, fontFace: BF, fontSize: 13, italic: true, color: C.muted, margin: 0, valign: "top", isTextBox: true });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: X0, y: 5.6, w: CW, h: 1.05, fill: { color: C.white }, line: { color: C.green, width: 1.25 } });
  s.addText([
    { text: "Percent inhibition,  L = [(C − T) / C] × 100", options: { fontFace: HF, fontSize: 19, bold: true, color: C.green, breakLine: true } },
    { text: "where C = radial growth in the control and T = radial growth in the treatment (Vincent, 1947)", options: { fontFace: BF, fontSize: 13, color: C.muted } },
  ], { x: X0 + 0.3, y: 5.6, w: CW - 0.6, h: 1.05, align: "center", valign: "middle", margin: 0, isTextBox: true });

  s = newSlide("PAIR_R");
  header(s, "Materials and methods", "Pot and field experiments");
  s.addText("Experimental sites", { x: X0, y: 1.65, w: 5.3, h: 0.4, fontFace: HF, fontSize: 18, bold: true, color: C.green, margin: 0, isTextBox: true });
  bullets(s, [
    "Pot: net house, Prof. Golam Ali Fakir Seed Pathology Center, BAU (CRD).",
    "Field: Plant Pathology field laboratory, BAU (RCBD); clay loam soil, pH 6.5.",
    "Variety: deep red plant with deep red calyx; sown on 30 March 2025.",
  ], X0, 2.1, 5.3, 2.3, 15, { paraSpaceAfter: 10 });
  panel(s, X0, 4.55, 5.3, 1.17);
  s.addText([
    { text: "Doses", options: { bold: true, color: C.green, breakLine: true } },
    { text: "Fungicides @ 500 ppm;  bioagents @ 10⁸ CFU/ml", options: { color: C.ink } },
  ], { x: X0 + 0.25, y: 4.55, w: 4.9, h: 1.17, fontFace: BF, fontSize: 15, valign: "middle", margin: 0, isTextBox: true });
  picLabels(s, PAIR_R, ["Figure 1. Pot experiment", "Figure 2. Field experiment"]);

  s = newSlide("CONTENT");
  header(s, "Materials and methods", "Treatments and spray schedule");
  const trt = [["T0", "Untreated control", "—"], ["T1", "Nativo 75 WG (tebuconazole + trifloxystrobin)", "500 ppm"], ["T2", "Tilt 250 EC (propiconazole)", "500 ppm"], ["T3", "Folicur 250 EC (tebuconazole)", "500 ppm"], ["T4", "Trichoderma harzianum (AB935952)", "10⁸ CFU/ml"], ["T5", "Trichoderma viride (TV6)", "10⁸ CFU/ml"], ["T6", "Trichoderma asperellum (AB935955)", "10⁸ CFU/ml"]];
  s.addTable([
    ["Code", "Treatment", "Dose"].map((t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.green } } })),
    ...trt.map((r, i) => r.map((v, j) => ({ text: j === 1 ? R(v) : v, options: { fill: { color: i % 2 ? C.stripe : C.white }, bold: j === 0 } }))),
  ], { x: X0, y: 1.7, w: 7.3, colW: [0.8, 4.9, 1.6], rowH: 0.47, fontFace: BF, fontSize: 13.5, color: C.ink, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1], border: { type: "solid", pt: 0.5, color: C.line } });
  s.addText("Application sequence", { x: 8.3, y: 1.65, w: 4.4, h: 0.4, fontFace: HF, fontSize: 18, bold: true, color: C.green, margin: 0, isTextBox: true });
  const seq = [["60 DAS", "First spray of treatments"], ["+ 48 h", "Cercospora spore suspension applied"], ["+ 48 h", "Second spray of treatments"], ["Every 25 days", "Repeat sprays until harvest"]];
  seq.forEach(([t, d], i) => {
    const y = 2.18 + i * 0.86;
    if (i < 3) s.addShape(pres.shapes.LINE, { x: 8.51, y: y + 0.42, w: 0, h: 0.44, line: { color: C.line, width: 1.5 } });
    numBadge(s, i + 1, 8.3, y, 0.42);
    s.addText([{ text: t, options: { bold: true, color: C.green, breakLine: true } }, ...R(d, { color: C.ink })], { x: 8.9, y: y - 0.06, w: 3.8, h: 0.62, fontFace: BF, fontSize: 13.5, margin: 0, valign: "top", isTextBox: true });
  });
  caption(s, "Schedule after", "Abo-Elyousr et al. (2017)", 8.3, 5.65, 4.4);
  s.addText("Recorded: disease incidence and severity (0–5 scale) on leaves and calyces; growth and yield traits.", { x: X0, y: 5.95, w: CW, h: 0.4, fontFace: BF, fontSize: 13.5, italic: true, color: C.muted, margin: 0, isTextBox: true });

  s = newSlide("CONTENT");
  header(s, "Materials and methods", "Disease assessment and data collection");
  caption(s, "Disease rating scale (0–5)", "(Apeyuan et al., 2017)", X0, 1.62, 5.6);
  const scale = [["0", "0", "No infection"], ["1", ">0–4", "Very mild"], ["2", ">4–12", "Mild"], ["3", ">12–29", "Moderate"], ["4", ">29–57", "Severe"], ["5", ">57–100", "Very severe"]];
  s.addTable([
    ["Score", "Leaf area infected (%)", "Remarks"].map((t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.green }, align: "center" } })),
    ...scale.map((r, i) => r.map((v) => ({ text: v, options: { fill: { color: i % 2 ? C.stripe : C.white }, align: "center" } }))),
  ], { x: X0, y: 2.0, w: 5.6, colW: [1.1, 2.4, 2.1], rowH: 0.52, fontFace: BF, fontSize: 14, color: C.ink, valign: "middle", border: { type: "solid", pt: 0.5, color: C.line } });
  const fm = [
    ["Disease incidence (%)", "= (Number of infected leaves or calyces ÷ Total number examined) × 100"],
    ["Disease severity (%)", "= Σ(Rating × Number in that rating) ÷ (Maximum rating × Total number) × 100"],
    ["Growth and yield (90 and 180 DAS)", "Plant height, base diameter, branches, leaves, leaf weight, fruit number, capsule and calyx weight"],
  ];
  fm.forEach(([t, d], i) => {
    const y = 2.0 + i * 1.42;
    panel(s, 6.6, y, 6.13, 1.25);
    s.addText([{ text: t, options: { fontFace: HF, fontSize: 16, bold: true, color: C.green, breakLine: true } }, { text: d, options: { fontFace: BF, fontSize: 14, color: C.ink } }], { x: 6.85, y, w: 5.7, h: 1.25, valign: "middle", margin: 0, isTextBox: true });
  });

  // ===== Results =====
  await divider("03", "Results", "Identification  ·  In vitro evaluation  ·  Pot and field evaluation");

  s = newSlide("GRID4");
  header(s, "Results  ·  Identification", "4.1 Symptoms of the disease");
  picLabels(s, GRID4, ["(a) Early symptom", "(b) Developed spot", "(c) Coalesced spots", "(d) Defoliation"], 0.04);
  panel(s, 7.05, 1.62, 5.68, 4.65);
  bullets(s, [
    "Small, circular to irregular, brown to reddish-brown necrotic spots, often restricted by leaf veins.",
    "Spots enlarged to about 10 mm, with pale brown to greyish-white centres and a dark reddish-purple margin.",
    "Under humid conditions, centres turned grey from abundant conidiophores and conidia.",
    "In the advanced stage, spots coalesced; leaves became chlorotic, dried and dropped prematurely.",
  ], 7.25, 1.85, 5.3, 4.3, 16.5, { paraSpaceAfter: 14 });
  caption(s, "Figure 3.", "Cercospora leaf spot symptoms on roselle", X0, 6.5, CW);

  s = newSlide("GRID4");
  header(s, "Results  ·  Identification", "4.2 Isolation and morphological identification");
  picLabels(s, GRID4, ["(a) Diseased leaves", "(b) Upper view of colony", "(c) Mycelium and conidiophores", "(d) Conidia"], 0.04);
  const morph = [["Colony on PDA", "Slow to moderate growth; greyish-white, olive-grey to dark grey mycelium of branched, septate hyphae."], ["Conidiophores", "Pigmented, septate, arising in fascicles."], ["Conidia", "Hyaline to pale olivaceous, cylindrical to obclavate, multiseptate."]];
  morph.forEach(([t, d], i) => {
    const y = 1.62 + i * 1.12;
    panel(s, 7.05, y, 5.68, 1.0);
    s.addText([{ text: t, options: { fontFace: HF, fontSize: 15, bold: true, color: C.green, breakLine: true } }, ...R(d, { fontFace: BF, fontSize: 13.5, color: C.ink })], { x: 7.25, y, w: 5.3, h: 1.0, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.05, y: 5.0, w: 5.68, h: 1.27, fill: { color: C.white }, line: { color: C.green, width: 1.25 } });
  s.addText(R("Provisionally identified as Cercospora sp. (Crous and Braun, 2003). Conidial size suggests C. malayensis; multilocus sequencing is needed for species-level identification.", { color: C.ink }), { x: 7.25, y: 5.0, w: 5.3, h: 1.27, fontFace: BF, fontSize: 13.5, valign: "middle", margin: 0, isTextBox: true });
  caption(s, "Figure 4.", "Different structures of Cercospora sp.", X0, 6.5, CW);

  s = newSlide("BIG1");
  header(s, "Results  ·  Identification", "4.3 Effect of temperature on mycelial growth");
  s.addShape(pres.shapes.RECTANGLE, { x: 8.0, y: 1.62, w: 4.73, h: 1.75, fill: { color: C.green }, line: { color: C.green } });
  s.addText([{ text: "90.00 mm", options: { fontFace: HF, fontSize: 36, bold: true, color: C.white, breakLine: true } }, { text: "maximum radial growth at 20 °C (10 DAI)", options: { fontFace: BF, fontSize: 14, color: C.onDark } }], { x: 8.3, y: 1.62, w: 4.2, h: 1.75, valign: "middle", margin: 0, isTextBox: true });
  bullets(s, [
    "Temperatures tested: 10, 15, 20, 25 and 30 °C.",
    "Growth at 20 °C was significantly higher than at the other temperatures at 10 DAI.",
    "Optimum range for radial growth: 20–22 °C; lowest growth at 10 °C.",
  ], 8.0, 3.6, 4.73, 2.6, 16.5, { paraSpaceAfter: 14 });
  caption(s, "Figure 5.", "Effect of temperature on the radial mycelial growth of Cercospora sp. on PDA at 10 DAI", X0, 6.3, CW);

  s = newSlide("PAIR");
  header(s, "Results  ·  Identification", "4.4 Pathogenicity test");
  picLabels(s, PAIR, ["(a) Healthy leaves (control)", "(b) Infected leaves, 4 DAI"]);
  panel(s, 8.0, 1.62, 4.73, 4.45);
  bullets(s, [
    "Conducted on non-detached leaves of pot-grown plants, without wounding the tissue.",
    "Irregular brown necrotic lesions appeared 4 days after inoculation.",
    "Control leaves remained healthy.",
    "Repeated twice with similar results; the pathogen was re-isolated, fulfilling Koch’s postulates.",
  ], 8.2, 1.85, 4.35, 4.1, 16.5, { paraSpaceAfter: 14 });
  caption(s, "Figure 6.", "Pathogenicity test of Cercospora sp. on roselle: (a) healthy leaves; (b) infected leaves", X0, 6.3, CW);

  s = newSlide("BIG1");
  header(s, "Results  ·  In vitro evaluation", "4.5 Effect of fungicides on mycelial growth");
  panel(s, 8.0, 1.62, 4.73, 4.55);
  bullets(s, [
    "Control showed the highest growth (90.00 mm) at 10 DAI.",
    "Complete inhibition (0.00 mm growth): Nativo 75 WG @ 250, 500 and 1000 ppm; Tilt 250 EC @ 500 and 1000 ppm; Bavistin @ 500 and 1000 ppm; Folicur @ 1000 ppm.",
    "Lowest inhibition: Folicur @ 50 ppm (25.66%, 43.15% and 49.08% at 6, 8 and 10 DAI).",
  ], 8.2, 1.85, 4.35, 4.2, 16.5, { paraSpaceAfter: 14 });
  caption(s, "Figure 7.", "Effect of fungicides and doses on radial mycelial growth inhibition of Cercospora sp. at (a) 6, (b) 8 and (c) 10 DAI", X0, 6.3, CW);

  s = newSlide("ROW5");
  header(s, "Results  ·  In vitro evaluation", "4.5 Fungicides at 500 ppm on PDA (10 DAI)");
  picLabels(s, ROW5, ["(a) Control", "(b) Nativo 75 WG", "(c) Tilt 250 EC", "(d) Bavistin 50 WP", "(e) Folicur 250 EC"]);
  panel(s, X0, 4.4, CW, 1.45);
  bullets(s, [
    "The untreated control plate was fully covered (90.00 mm) at 10 DAI.",
    "Nativo 75 WG, Tilt 250 EC and Bavistin at 500 ppm completely inhibited mycelial growth.",
  ], X0 + 0.2, 4.55, CW - 0.4, 1.2, 15, { paraSpaceAfter: 8 });
  caption(s, "Figure 8.", "Effect of fungicides at 500 ppm on radial mycelial growth of Cercospora sp. at 10 DAI", X0, 6.25, CW);

  s = newSlide("CONTENT");
  header(s, "Results  ·  In vitro evaluation", "4.6 Antagonistic effect of bioagents");
  s.addChart(pres.charts.BAR, [{ name: "Inhibition (%)", labels: ["T. asperellum", "T. harzianum", "T. viride"], values: [90.74, 88.15, 65.93] }], {
    x: X0, y: 1.62, w: 7.1, h: 4.55, barDir: "col", barGapWidthPct: 90, chartColors: [C.bar],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.00"%"', dataLabelColor: C.ink, dataLabelFontSize: 13, dataLabelFontFace: BF,
    catAxisLabelColor: C.ink, catAxisLabelFontSize: 14, catAxisLabelFontFace: BF, catAxisLineShow: true, catAxisLineColor: C.line,
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 20, valAxisLabelColor: C.muted, valAxisLabelFontSize: 11, valAxisLineShow: false,
    showValAxisTitle: true, valAxisTitle: "Mycelial growth inhibition (%)", valAxisTitleFontSize: 12, valAxisTitleColor: C.muted, valAxisTitleFontFace: BF,
    valGridLine: { color: "E3E9E5", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false,
  });
  panel(s, 8.0, 1.62, 4.73, 4.55);
  bullets(s, [
    "Mycelial growth at 10 DAI ranged from 8.33 mm to 90.00 mm (control).",
    "Maximum inhibition by T. asperellum (90.74%), statistically similar to T. harzianum (88.15%).",
    "T. viride gave the lowest inhibition (65.93%) and was moderately effective.",
  ], 8.2, 1.85, 4.35, 4.2, 16.5, { paraSpaceAfter: 14 });
  caption(s, "Figure 9.", "Antagonistic effect of fungal bioagents on radial mycelial growth inhibition (%) of Cercospora sp. at 10 DAI", X0, 6.3, CW);

  s = newSlide("ROW4");
  header(s, "Results  ·  In vitro evaluation", "4.6 Dual culture plates");
  picLabels(s, ROW4, ["T. harzianum (AB935952) vs Cercospora sp.", "T. viride (TV6) vs Cercospora sp.", "T. asperellum (AB935955) vs Cercospora sp.", "Control: Cercospora sp."]);
  panel(s, X0, 5.0, CW, 0.75);
  s.addText(R("Mycelial growth inhibition at 10 DAI:   T. asperellum 90.74%   ·   T. harzianum 88.15%   ·   T. viride 65.93%", { color: C.ink }), { x: X0, y: 5.0, w: CW, h: 0.75, fontFace: BF, fontSize: 15, align: "center", valign: "middle", margin: 0, isTextBox: true });
  caption(s, "Figure 10.", "Antagonistic effect of fungal bioagents on radial mycelial growth of Cercospora sp. causing leaf spot of roselle", X0, 6.05, CW);

  // ----- Pot and field evaluation (Tables 5-14) -----
  const DIDS = ["Treatments", "DI (%) 90 DAS", "DS (%) 90 DAS", "DI (%) 180 DAS", "DS (%) 180 DAS"];
  tableSlide("Results  ·  Pot experiment", "4.7 Leaf disease incidence and severity (pot)", 5,
    "Effect of different treatments on disease incidence (DI) and disease severity (DS) percentage of roselle plant under pot experiment", DIDS,
    [["53.91 a", "31.33 a", "73.41 a", "39.33 a"], ["14.04 de", "8.00 d", "21.93 e", "12.00 c"], ["12.10 e", "5.33 d", "15.98 f", "8.67 d"], ["16.70 d", "8.67 cd", "25.65 d", "13.33 c"], ["22.67 c", "12.67 b", "31.36 c", "18.00 b"], ["26.48 b", "14.67 b", "38.62 b", "18.67 b"], ["20.89 c", "12.00 bc", "29.69 c", "16.67 b"]],
    ["0.97", "1.01", "0.91", "0.80"], ["3.30", "3.44", "3.12", "2.72"],
    "T2 (Tilt 250 EC) gave the lowest leaf disease incidence and severity at both 90 and 180 DAS; among the bioagents, T. asperellum (T6) was the most effective.");
  tableSlide("Results  ·  Pot experiment", "4.8 Plant growth at 90 DAS (pot)", 6,
    "Effect of different treatments on various agronomic parameters of roselle at 90 DAS under pot experiment",
    ["Treatments", "Plant height (cm)", "Base diameter (mm)", "Leaf no.", "Fresh leaf weight (g)"],
    [["55.20 d", "12.67 e", "16.00 e", "17.67 f"], ["86.40 a", "25.33 a", "35.33 ab", "40.61 ab"], ["93.10 a", "27.33 a", "37.67 a", "44.15 a"], ["76.81 b", "21.67 b", "31.33 b", "35.19 bc"], ["67.33 c", "17.00 cd", "23.00 cd", "28.14 de"], ["66.30 c", "16.00 d", "20.67 d", "24.44 e"], ["71.17 bc", "19.33 bc", "26.33 c", "31.37 cd"]],
    ["2.75", "0.69", "1.18", "1.86"], ["9.38", "2.36", "4.04", "6.35"],
    "T2 (Tilt 250 EC) recorded the highest values for all growth parameters, statistically at par with T1 (Nativo 75 WG); T. asperellum was the best bioagent.");
  tableSlide("Results  ·  Pot experiment", "4.8 Plant growth at 180 DAS (pot)", 7,
    "Effect of different treatments on various agronomic parameters of roselle at 180 DAS under pot experiment",
    ["Treatments", "Plant height (cm)", "Base diameter (mm)", "Effective branch no.", "Leaf no.", "Fresh leaf weight (g)"],
    [["105.77 f", "28.17 e", "3.67 d", "43.33 f", "73.77 e"], ["152.77 b", "43.33 ab", "7.67 ab", "90.00 b", "146.41 b"], ["165.33 a", "47.33 a", "9.00 a", "101.67 a", "167.86 a"], ["146.33 bc", "39.67 bc", "6.33 bc", "79.33 c", "138.13 b"], ["130.13 de", "35.33 cd", "6.00 bc", "58.00 e", "103.91 c"], ["121.09 e", "33.50 d", "5.33 cd", "52.67 e", "87.61 d"], ["138.38 cd", "36.67 cd", "6.33 bc", "71.67 d", "114.82 c"]],
    ["3.61", "1.33", "0.59", "2.13", "4.03"], ["12.35", "4.56", "2.02", "7.28", "13.75"],
    "T2 (Tilt 250 EC) recorded the highest values for all growth parameters at 180 DAS; T. asperellum performed best among the bioagents.");
  tableSlide("Results  ·  Field experiment", "4.9 Leaf disease incidence and severity (field)", 8,
    "Effect of different treatments on disease incidence (DI) and disease severity (DS) percentage of roselle plant under field condition", DIDS,
    [["56.33 a", "33.33 a", "77.30 a", "57.33 a"], ["18.24 d", "8.67 de", "24.79 e", "14.67 e"], ["14.59 e", "6.67 e", "17.46 f", "10.00 f"], ["19.64 d", "11.33 cd", "29.67 d", "15.33 de"], ["26.63 bc", "15.33 b", "35.34 c", "19.33 bc"], ["23.70 c", "14.00 bc", "32.71 cd", "18.00 cd"], ["30.14 b", "16.67 b", "39.51 b", "22.00 b"]],
    ["1.04", "0.87", "1.09", "0.94"], ["3.56", "2.98", "3.72", "3.22"],
    "T2 (Tilt 250 EC) again gave the lowest leaf disease incidence and severity; among the bioagents, T. viride (T5) performed best under field conditions.");
  tableSlide("Results  ·  Field experiment", "4.10 Plant growth at 90 DAS (field)", 9,
    "Effect of different treatments on various agronomic parameters of roselle at 90 DAS under field experiment",
    ["Treatments", "Plant height (cm)", "Base diameter (mm)", "Effective branch no.", "No. of leaves", "Leaf fresh weight (g)"],
    [["70.55 d", "26.67 e", "5.33 c", "46.67 g", "61.50 c"], ["97.15 a", "35.33 ab", "8.67 a", "65.67 b", "72.60 ab"], ["103.05 a", "38.33 a", "9.00 a", "69.33 a", "75.80 a"], ["93.49 ab", "33.33 bc", "7.33 b", "59.33 c", "70.60 ab"], ["80.17 cd", "28.67 de", "7.00 b", "54.33 e", "66.70 bc"], ["76.44 cd", "28.67 de", "6.67 b", "52.00 f", "64.50 bc"], ["86.23 bc", "31.00 cd", "7.00 b", "57.33 d", "68.60 abc"]],
    ["3.18", "1.20", "0.36", "0.53", "2.47"], ["10.86", "4.09", "1.22", "1.83", "8.44"],
    "T2 (Tilt 250 EC) recorded the highest values for all growth parameters at 90 DAS; T. asperellum was generally the best bioagent.");
  tableSlide("Results  ·  Field experiment", "4.10 Plant growth at 180 DAS (field)", 10,
    "Effect of different treatments on various agronomic parameters of roselle at 180 DAS under field experiment",
    ["Treatments", "Plant height (cm)", "Base diameter (mm)", "Effective branch no.", "No. of leaves", "Leaf fresh weight (g)"],
    [["118.64 d", "44.33 d", "8.33 d", "124.00 f", "154.47 e"], ["151.33 b", "54.43 b", "13.67 b", "154.67 b", "184.23 b"], ["172.88 a", "63.33 a", "16.67 a", "163.00 a", "196.28 a"], ["148.30 b", "53.67 bc", "12.67 bc", "147.33 c", "179.72 bc"], ["129.74 cd", "51.33 bc", "12.00 c", "138.33 d", "171.13 cd"], ["124.73 d", "49.00 cd", "11.33 c", "133.67 e", "169.33 d"], ["140.44 bc", "52.67 bc", "12.33 bc", "144.67 c", "175.91 bcd"]],
    ["3.53", "1.45", "0.44", "1.33", "2.96"], ["12.06", "4.97", "1.49", "4.56", "10.11"],
    "T2 (Tilt 250 EC) recorded the highest values for all growth parameters at 180 DAS; T. asperellum was the best bioagent.");
  const CAL = ["Treatments", "DI (%) 90 DAS", "DI (%) 180 DAS", "DS (%) 90 DAS", "DS (%) 180 DAS"];
  tableSlide("Results  ·  Pot experiment", "4.11 Calyx disease incidence and severity (pot)", 11,
    "Effect of different treatments on disease incidence (DI) and disease severity (DS) percentage of roselle calyx under pot experiment", CAL,
    [["16.00 a", "27.67 a", "16.00 a", "26.00 a"], ["5.33 bc", "9.00 bc", "7.20 bc", "8.50 bc"], ["4.00 c", "6.00 c", "6.00 c", "7.00 c"], ["6.00 bc", "9.50 bc", "7.80 bc", "10.20 bc"], ["7.67 bc", "12.00 b", "9.00 bc", "12.80 b"], ["10.00 b", "14.00 b", "10.00 bc", "14.00 b"], ["9.00 bc", "13.00 b", "11.00 b", "13.00 b"]],
    ["1.56", "1.56", "1.41", "1.65"], ["5.34", "5.33", "4.83", "5.63"],
    "T2 (Tilt 250 EC) recorded the lowest calyx disease incidence and severity at both stages; T. harzianum (T4) was the most effective bioagent.");
  tableSlide("Results  ·  Field experiment", "4.12 Calyx disease incidence and severity (field)", 12,
    "Effect of different treatments on disease incidence (DI) and disease severity (DS) percentage of roselle calyx under field condition", CAL,
    [["20.00 a", "52.00 a", "18.00 a", "34.00 a"], ["7.20 c", "11.50 bc", "12.00 b", "9.20 d"], ["6.00 c", "7.50 c", "11.40 b", "8.80 d"], ["8.00 bc", "12.00 bc", "11.00 b", "11.00 cd"], ["12.00 b", "16.00 b", "7.50 b", "17.00 b"], ["12.20 b", "17.00 b", "8.00 b", "16.00 bc"], ["9.20 bc", "14.00 b", "9.00 b", "15.00 bc"]],
    ["1.36", "1.77", "1.50", "1.56"], ["4.66", "6.06", "5.13", "5.33"],
    "T2 (Tilt 250 EC) gave the lowest calyx disease incidence at both stages and the lowest severity at 180 DAS; at 90 DAS, severity did not differ among the treatments.");
  const YLD = ["Treatments", "Fruit no./plant", "Capsule fresh weight (g/plant)", "Calyx fresh weight (g/plant)"];
  tableSlide("Results  ·  Field experiment", "4.13 Yield attributes at 90 DAS (field)", 13,
    "Effect of different treatments on various agronomic parameters of roselle at 90 DAS under field condition", YLD,
    [["17.67 d", "21.61 d", "15.56 d"], ["30.33 a", "51.13 a", "41.25 a"], ["32.33 a", "55.45 a", "45.14 a"], ["25.33 b", "40.26 b", "29.55 b"], ["24.67 bc", "31.28 c", "26.17 bc"], ["22.33 c", "28.66 c", "23.09 c"], ["25.00 bc", "37.14 b", "28.04 b"]],
    ["0.82", "1.48", "1.17"], ["2.79", "5.06", "4.00"],
    "T2 (Tilt 250 EC) gave the highest yield attributes, at par with T1 (Nativo 75 WG); its calyx weight was about 2.9 times that of the control.");
  tableSlide("Results  ·  Field experiment", "4.13 Yield attributes at 180 DAS (field)", 14,
    "Effect of different treatments on various agronomic parameters of roselle at 180 DAS under field condition", YLD,
    [["183.67 d", "605.20 f", "377.31 e"], ["251.67 b", "926.70 b", "677.82 b"], ["289.33 a", "1071.30 a", "820.14 a"], ["249.67 b", "894.30 bc", "628.64 b"], ["225.33 bc", "786.20 de", "530.98 c"], ["215.33 c", "710.44 e", "451.23 d"], ["242.00 bc", "831.60 cd", "549.51 c"]],
    ["8.40", "27.57", "21.59"], ["4.35", "4.06", "4.59"],
    "T2 (Tilt 250 EC) gave the highest yield; calyx weight (820.14 g/plant) was more than double the control (377.31 g/plant). T. asperellum was the best bioagent.");

  // ===== Discussion and conclusions =====
  await divider("04", "Discussion and Conclusions", "Discussion  ·  Conclusions  ·  Recommendations");

  s = newSlide("CONTENT");
  header(s, "Discussion", "Key discussion points");
  const disc = [
    ["In vitro versus in planta efficacy", "Tilt 250 EC needed a higher dose than Nativo 75 WG for complete inhibition in vitro, but gave the best control on plants, consistent with Teramoto et al. (2011)."],
    ["Variation in bioagent performance", "T. asperellum was the most effective bioagent in dual culture and in pots, whereas T. viride and T. harzianum performed better in the field."],
    ["Growth promotion by Trichoderma", "T. asperellum gave the best growth and yield among the bioagents, possibly through root colonization and induced resistance."],
    ["Species identity unresolved", "Conidial size points to C. malayensis; sequencing of ITS, tef1 and actin or calmodulin genes is required for confirmation."],
  ];
  disc.forEach(([t, d], i) => {
    const x = X0 + (i % 2) * 6.17, y = 1.65 + Math.floor(i / 2) * 2.5;
    panel(s, x, y, 5.96, 2.3);
    numBadge(s, i + 1, x + 0.3, y + 0.3, 0.5);
    s.addText(R(t), { x: x + 1.0, y: y + 0.25, w: 4.75, h: 0.6, fontFace: HF, fontSize: 17, bold: true, color: C.green, valign: "middle", margin: 0, isTextBox: true });
    s.addText(R(d), { x: x + 1.0, y: y + 0.9, w: 4.75, h: 1.3, fontFace: BF, fontSize: 14.5, color: C.ink, valign: "top", margin: 0, isTextBox: true });
  });

  s = newSlide("CONTENT");
  header(s, "Conclusions", "Conclusions");
  const conc = [
    "Leaf spot of roselle in Mymensingh and Tangail is caused by Cercospora sp.; this is the first such study in Bangladesh.",
    "The pathogen grows best at 20–22 °C and infects intact leaves within 4 days.",
    "In vitro, Nativo 75 WG, Tilt 250 EC and Bavistin 50 WP gave 100% inhibition at ≤ 500 ppm; T. asperellum and T. harzianum gave about 90% inhibition.",
    "Tilt 250 EC @ 500 ppm, sprayed from 60 DAS at 25-day intervals, gave the best control and more than doubled calyx yield.",
    "Trichoderma spp. are promising, farmer-friendly options for integrated management.",
  ];
  conc.forEach((t, i) => {
    const y = 1.62 + i * 0.97;
    panel(s, X0, y, CW, 0.83);
    numBadge(s, i + 1, X0 + 0.22, y + 0.19, 0.45);
    s.addText(R(t), { x: X0 + 0.95, y, w: CW - 1.2, h: 0.83, fontFace: BF, fontSize: 16, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
  });

  s = newSlide("CONTENT");
  header(s, "Recommendations", "Recommendations and future work");
  const rec = [
    ["FaDna", "Molecular identification", "Multilocus sequencing (ITS, tef1, actin or calmodulin) to identify the species."],
    ["FaShieldAlt", "Integrated management", "Combine Trichoderma with reduced fungicide programmes to limit fungicide resistance."],
    ["FaSeedling", "Host resistance", "Screen roselle varieties for resistance to Cercospora leaf spot."],
    ["FaClipboardCheck", "Improved assessment", "Validated diagrammatic severity scale; multi-season, multi-location trials."],
  ];
  for (let i = 0; i < 4; i++) {
    const x = X0 + (i % 2) * 6.17, y = 1.65 + Math.floor(i / 2) * 2.5;
    panel(s, x, y, 5.96, 2.3);
    await iconCircle(s, rec[i][0], x + 0.3, y + 0.3, 0.62);
    s.addText(rec[i][1], { x: x + 1.15, y: y + 0.3, w: 4.6, h: 0.62, fontFace: HF, fontSize: 17, bold: true, color: C.green, valign: "middle", margin: 0, isTextBox: true });
    s.addText(R(rec[i][2]), { x: x + 1.15, y: y + 1.0, w: 4.6, h: 1.15, fontFace: BF, fontSize: 15, color: C.ink, valign: "top", margin: 0, isTextBox: true });
  }

  s = newSlide("DIVIDER");
  s.addImage({ data: await iconPng("FaLeaf", C.green), x: 8.7, y: 1.65, w: 4.2, h: 4.2, transparency: 88 });
  s.addText("Thank you", { x: 0.9, y: 2.3, w: 8, h: 1.1, fontFace: HF, fontSize: 54, bold: true, color: C.green, margin: 0, isTextBox: true });
  s.addText("Questions and suggestions are welcome", { x: 0.9, y: 3.45, w: 8, h: 0.5, fontFace: BF, fontSize: 20, color: C.gold, margin: 0, isTextBox: true });
  s.addText("This research was supported by BAURES Project No. 111/BAU/2024 and the NST Fellowship, Ministry of Science and Technology.", { x: 0.9, y: 5.9, w: 9, h: 0.6, fontFace: BF, fontSize: 13, color: C.muted, margin: 0, valign: "top", isTextBox: true });

  // ---------- Write, then mark custom placeholders as picture placeholders ----------
  const buf = await pres.write({ outputType: "nodebuffer" });
  const zip = await JSZip.loadAsync(buf);
  for (const f of Object.keys(zip.files).filter((f) => /^ppt\/(slides|slideLayouts)\/[^/]+\.xml$/.test(f))) {
    const xml = await zip.file(f).async("string");
    const fixed = xml
      .replace(/<p:ph\s+idx="(\d+)"\s+hasCustomPrompt="1"\s*\/>/g, '<p:ph type="pic" idx="$1" hasCustomPrompt="1"/>')
      .replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, inner) => {
        // a pPr is only valid as the paragraph's first child
        return "<a:p>" + inner.replace(/<a:pPr\b[^>]*?\/>|<a:pPr\b[^>]*?>[\s\S]*?<\/a:pPr>/g, (pp, off) => (off === 0 ? pp : "")) + "</a:p>";
      });
    if (fixed !== xml) zip.file(f, fixed);
  }
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT, QA ? "(QA preview)" : "");
})();
