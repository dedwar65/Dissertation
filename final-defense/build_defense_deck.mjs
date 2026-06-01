import fs from "node:fs/promises";
import path from "node:path";
import {
  Presentation,
  PresentationFile,
} from "file:///Users/dc/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const ROOT =
  "/Users/dc/Library/Mobile Documents/com~apple~CloudDocs/Github/Forks/JHU-Dissertation";
const WORKSPACE = path.join(
  ROOT,
  "outputs",
  "manual-20260601-final-defense",
  "presentations",
  "final-defense",
);
const ASSET_DIR = path.join(WORKSPACE, "assets");
const OUTPUT_PPTX = path.join(ROOT, "Dissertation_Defense_June_22_2026.pptx");

const colors = {
  paper: "#F7F2E9",
  white: "#FFFDFC",
  navy: "#0B2D4D",
  blue: "#355C7D",
  teal: "#4E7E87",
  gold: "#C79A3B",
  clay: "#B86854",
  ink: "#18212B",
  muted: "#5B6470",
  line: "#D8CEBF",
  paleBlue: "#E7EEF5",
  paleGold: "#F3E7C9",
};

const fonts = {
  sans: "Avenir Next",
  serif: "Baskerville",
};

const imageCache = new Map();

async function imagePayload(fileName) {
  const fullPath = path.join(ASSET_DIR, fileName);
  if (!imageCache.has(fullPath)) {
    const data = await fs.readFile(fullPath);
    imageCache.set(fullPath, {
      data,
      contentType: fileName.endsWith(".png") ? "image/png" : "image/jpeg",
    });
  }
  return imageCache.get(fullPath);
}

function setTextStyle(text, options = {}) {
  if (options.fontSize !== undefined) text.fontSize = options.fontSize;
  if (options.color) text.color = options.color;
  if (options.typeface) text.typeface = options.typeface;
  if (options.alignment) text.alignment = options.alignment;
  if (options.verticalAlignment) text.verticalAlignment = options.verticalAlignment;
  if (options.bold !== undefined) text.bold = options.bold;
  if (options.italic !== undefined) text.italic = options.italic;
  if (options.insets) text.insets = options.insets;
  if (options.lineSpacing !== undefined) text.lineSpacing = options.lineSpacing;
}

function normalizeFill(fill) {
  if (fill === undefined) return undefined;
  if (fill === "none") return { type: "none" };
  return fill;
}

function normalizeLine(line) {
  if (line === undefined) return undefined;
  if (line === "none") {
    return { width: 0, fill: { type: "none" } };
  }
  return line;
}

function addBox(slide, opts) {
  const shape = slide.shapes.add({
    geometry: "rect",
    position: {
      left: opts.left,
      top: opts.top,
      width: opts.width,
      height: opts.height,
    },
    fill: normalizeFill(opts.fill ?? "none"),
    line: normalizeLine(opts.line ?? "none"),
  });
  if (opts.radius !== undefined) {
    shape.borderRadius = opts.radius;
  }
  if (opts.text) {
    shape.text.set(opts.text);
    setTextStyle(shape.text, {
      fontSize: opts.fontSize ?? 22,
      color: opts.color ?? colors.ink,
      typeface: opts.typeface ?? fonts.sans,
      alignment: opts.alignment ?? "left",
      verticalAlignment: opts.verticalAlignment ?? "top",
      bold: opts.bold,
      italic: opts.italic,
      insets:
        opts.insets ?? { top: 18, right: 18, bottom: 18, left: 18 },
      lineSpacing: opts.lineSpacing,
    });
  }
  return shape;
}

function addTitle(slide, kicker, title, slideNum) {
  addBox(slide, {
    left: 64,
    top: 32,
    width: 14,
    height: 70,
    fill: colors.gold,
    line: "none",
    radius: 8,
  });
  addBox(slide, {
    left: 92,
    top: 28,
    width: 420,
    height: 28,
    text: kicker,
    fontSize: 14,
    color: colors.blue,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBox(slide, {
    left: 92,
    top: 54,
    width: 1080,
    height: 64,
    text: title,
    fontSize: 30,
    color: colors.navy,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBox(slide, {
    left: 92,
    top: 124,
    width: 1088,
    height: 2,
    fill: colors.line,
    line: "none",
  });
  addBox(slide, {
    left: 1115,
    top: 660,
    width: 80,
    height: 24,
    text: String(slideNum),
    fontSize: 14,
    color: colors.muted,
    alignment: "right",
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
}

function addFooter(slide, text = "Essays on Heterogeneous Returns, Trust, and Household Wealth") {
  addBox(slide, {
    left: 92,
    top: 660,
    width: 760,
    height: 22,
    text,
    fontSize: 12,
    color: colors.muted,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
}

function addBulletList(slide, left, top, width, bullets, options = {}) {
  const bulletText = bullets.map((item) => `• ${item}`).join("\n");
  return addBox(slide, {
    left,
    top,
    width,
    height: options.height ?? bullets.length * 52,
    text: bulletText,
    fontSize: options.fontSize ?? 22,
    color: options.color ?? colors.ink,
    line: "none",
    fill: "none",
    insets: options.insets ?? { top: 4, right: 0, bottom: 0, left: 0 },
    lineSpacing: options.lineSpacing ?? 1.18,
  });
}

function addStatCard(slide, { left, top, width, height, label, value, tone = "blue" }) {
  const fillMap = {
    blue: colors.paleBlue,
    gold: colors.paleGold,
    clay: "#F6E1DB",
  };
  addBox(slide, {
    left,
    top,
    width,
    height,
    fill: fillMap[tone] ?? colors.paleBlue,
    line: { width: 1, fill: colors.line },
    radius: 18,
  });
  addBox(slide, {
    left: left + 18,
    top: top + 16,
    width: width - 36,
    height: 28,
    text: label,
    fontSize: 14,
    color: colors.muted,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBox(slide, {
    left: left + 18,
    top: top + 42,
    width: width - 36,
    height: height - 50,
    text: value,
    fontSize: 24,
    color: colors.navy,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
}

async function addImage(slide, fileName, position) {
  slide.images.add({
    ...(await imagePayload(fileName)),
    fit: "contain",
    position,
  });
}

async function buildDeck() {
  const deck = Presentation.create();

  const titleSlide = deck.slides.add();
  titleSlide.background.fill = colors.paper;
  addBox(titleSlide, {
    left: 0,
    top: 0,
    width: 360,
    height: 720,
    fill: colors.navy,
    line: "none",
  });
  addBox(titleSlide, {
    left: 58,
    top: 82,
    width: 170,
    height: 8,
    fill: colors.gold,
    line: "none",
    radius: 4,
  });
  addBox(titleSlide, {
    left: 58,
    top: 108,
    width: 230,
    height: 42,
    text: "Final Dissertation Defense",
    fontSize: 17,
    color: "#E7EDF3",
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBox(titleSlide, {
    left: 58,
    top: 520,
    width: 230,
    height: 88,
    text: "Decory Edwards\nJune 22, 2026",
    fontSize: 20,
    color: "#F3F6FA",
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.15,
  });
  addBox(titleSlide, {
    left: 420,
    top: 102,
    width: 780,
    height: 210,
    text: "Essays on Heterogeneous Returns, Trust, and Household Wealth",
    fontSize: 42,
    color: colors.navy,
    typeface: fonts.serif,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.02,
  });
  addBox(titleSlide, {
    left: 422,
    top: 352,
    width: 660,
    height: 108,
    text:
      "Core theme: wealth inequality is shaped not only by earnings and savings, but by heterogeneity in rates of return, trust-related performance, and how we measure unequal wealth distributions.",
    fontSize: 24,
    color: colors.ink,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.18,
  });
  addBox(titleSlide, {
    left: 424,
    top: 512,
    width: 260,
    height: 116,
    fill: colors.white,
    line: { width: 1, fill: colors.line },
    radius: 18,
  });
  addBox(titleSlide, {
    left: 444,
    top: 534,
    width: 220,
    height: 22,
    text: "Defense structure",
    fontSize: 14,
    color: colors.blue,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBox(titleSlide, {
    left: 444,
    top: 565,
    width: 220,
    height: 56,
    text: "Introduction\nThree chapters\nSynthesis and conclusion",
    fontSize: 19,
    color: colors.ink,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.12,
  });

  const slide2 = deck.slides.add();
  slide2.background.fill = colors.paper;
  addTitle(slide2, "Roadmap", "The dissertation asks one broad question", 2);
  addBox(slide2, {
    left: 96,
    top: 168,
    width: 1030,
    height: 60,
    text:
      "How do heterogeneous rates of return reshape the explanation, performance, and measurement of wealth inequality?",
    fontSize: 31,
    color: colors.navy,
    typeface: fonts.serif,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  const roadmapCards = [
    {
      left: 96,
      title: "Chapter 1\nStructural macro",
      body:
        "Embed return heterogeneity in a standard heterogeneous-agent model and estimate the return distribution needed to match wealth moments.",
      tone: colors.paleBlue,
    },
    {
      left: 435,
      title: "Chapter 2\nTrust and performance",
      body:
        "Measure returns to net wealth in the HRS and test whether trust has a nonlinear relationship with long-run household investment performance.",
      tone: colors.paleGold,
    },
    {
      left: 774,
      title: "Chapter 3\nWealth inequality under ambiguity",
      body:
        "Motivated by Rawls, build an ambiguity-based wealth inequality measure and compare it with the standard objective-uncertainty benchmark.",
      tone: "#F6E1DB",
    },
  ];
  for (const card of roadmapCards) {
    addBox(slide2, {
      left: card.left,
      top: 280,
      width: 300,
      height: 248,
      fill: card.tone,
      line: { width: 1, fill: colors.line },
      radius: 22,
    });
    addBox(slide2, {
      left: card.left + 22,
      top: 304,
      width: 256,
      height: 70,
      text: card.title,
      fontSize: 24,
      color: colors.navy,
      bold: true,
      line: "none",
      fill: "none",
      insets: { top: 0, right: 0, bottom: 0, left: 0 },
      lineSpacing: 1.05,
    });
    addBox(slide2, {
      left: card.left + 22,
      top: 392,
      width: 256,
      height: 116,
      text: card.body,
      fontSize: 19,
      color: colors.ink,
      line: "none",
      fill: "none",
      insets: { top: 0, right: 0, bottom: 0, left: 0 },
      lineSpacing: 1.16,
    });
  }
  addFooter(slide2);

  const slide3 = deck.slides.add();
  slide3.background.fill = colors.paper;
  addTitle(slide3, "Chapter 1", "Matching wealth moments with heterogeneous returns", 3);
  addBox(slide3, {
    left: 96,
    top: 164,
    width: 440,
    height: 52,
    text: "Claim",
    fontSize: 18,
    color: colors.blue,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBox(slide3, {
    left: 96,
    top: 196,
    width: 470,
    height: 126,
    text:
      "Return heterogeneity is a quantitatively important missing ingredient in structural macro models of wealth inequality.",
    fontSize: 31,
    color: colors.navy,
    typeface: fonts.serif,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.03,
  });
  addBulletList(slide3, 96, 354, 470, [
    "Start from a standard heterogeneous-agent model with labor income risk.",
    "Estimate the return distribution needed to match SCF wealth moments.",
    "Interpret the implied dispersion as consistent with empirical heterogeneity in realized rates of return.",
  ]);
  addBox(slide3, {
    left: 620,
    top: 176,
    width: 520,
    height: 332,
    fill: colors.white,
    line: { width: 1, fill: colors.line },
    radius: 22,
  });
  addBox(slide3, {
    left: 650,
    top: 208,
    width: 460,
    height: 32,
    text: "Model logic",
    fontSize: 18,
    color: colors.blue,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBulletList(slide3, 650, 252, 430, [
    "Income risk and life-cycle structure explain part of wealth dispersion, but not enough.",
    "A return distribution lets the model fit the lower and upper tails much more closely.",
    "That better fit matters for aggregate MPCs, tax comparisons, and the interpretation of wealth concentration.",
  ], { fontSize: 21, height: 210 });
  addFooter(slide3);

  const slide4 = deck.slides.add();
  slide4.background.fill = colors.paper;
  addTitle(slide4, "Chapter 1 Result", "Heterogeneous returns sharply improve the fit to the wealth distribution", 4);
  await addImage(slide4, "ch1_fig1_page31_cropped.png", {
    left: 80,
    top: 180,
    width: 510,
    height: 315,
  });
  addStatCard(slide4, {
    left: 650,
    top: 182,
    width: 235,
    height: 110,
    label: "Estimated return distribution",
    value: "R = 1.0204\n∇ = 0.0683",
    tone: "blue",
  });
  addStatCard(slide4, {
    left: 905,
    top: 182,
    width: 235,
    height: 110,
    label: "Top 5% wealth share",
    value: "Data 57.4%\nNo het. 23.3%\nHet. 58.8%",
    tone: "gold",
  });
  addStatCard(slide4, {
    left: 650,
    top: 312,
    width: 490,
    height: 110,
    label: "Life-cycle implication",
    value: "Aggregate MPC rises from about 12.6% to 27.8% once return heterogeneity is added.",
    tone: "clay",
  });
  addBulletList(slide4, 650, 460, 470, [
    "The Lorenz curve with return heterogeneity tracks the SCF much better than the no-heterogeneity benchmark.",
    "The estimated return distribution is modest in scale, not an implausibly extreme calibration.",
    "Takeaway: heterogeneous returns can be a disciplined explanation for observed wealth concentration.",
  ], { fontSize: 20, height: 170 });
  addFooter(slide4);

  const slide5 = deck.slides.add();
  slide5.background.fill = colors.paper;
  addTitle(slide5, "Chapter 2", "Moderate trust and maximum performance on financial investments", 5);
  addBox(slide5, {
    left: 96,
    top: 174,
    width: 450,
    height: 124,
    text:
      "Question: if trust shapes how households interact with financial institutions and markets, does a moderate amount of trust maximize returns to net wealth?",
    fontSize: 30,
    color: colors.navy,
    typeface: fonts.serif,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.05,
  });
  addBox(slide5, {
    left: 96,
    top: 350,
    width: 480,
    height: 190,
    fill: colors.white,
    line: { width: 1, fill: colors.line },
    radius: 22,
  });
  addBulletList(slide5, 122, 376, 430, [
    "Use the RAND HRS panel to construct realized returns to net wealth.",
    "Trust in people is observed in 2020 on a 1 to 10 scale.",
    "Estimate pooled, FE, RE, and CRE specifications to separate time-varying and persistent return components.",
  ], { fontSize: 20, height: 138 });
  addStatCard(slide5, {
    left: 640,
    top: 190,
    width: 220,
    height: 100,
    label: "Net wealth return distribution",
    value: "Mean 0.201\nMedian 0.084\nP99 2.697",
    tone: "blue",
  });
  addStatCard(slide5, {
    left: 882,
    top: 190,
    width: 258,
    height: 100,
    label: "Why CRE matters",
    value: "Trust is time-invariant in the HRS, so panel structure matters for identification.",
    tone: "gold",
  });
  addBox(slide5, {
    left: 640,
    top: 334,
    width: 500,
    height: 206,
    fill: "#F6E1DB",
    line: { width: 1, fill: colors.line },
    radius: 22,
  });
  addBox(slide5, {
    left: 668,
    top: 362,
    width: 444,
    height: 34,
    text: "Empirical target",
    fontSize: 18,
    color: colors.blue,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBox(slide5, {
    left: 668,
    top: 404,
    width: 430,
    height: 118,
    text:
      "Estimate a hump-shaped relationship in trust while retaining an interpretable persistent household component in returns.",
    fontSize: 24,
    color: colors.ink,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.13,
  });
  addFooter(slide5);

  const slide6 = deck.slides.add();
  slide6.background.fill = colors.paper;
  addTitle(slide6, "Chapter 2 Result", "Returns peak at moderate trust, and persistent heterogeneity survives richer controls", 6);
  await addImage(slide6, "ch2_trustdist_page122_cropped.png", {
    left: 60,
    top: 186,
    width: 430,
    height: 360,
  });
  addStatCard(slide6, {
    left: 540,
    top: 188,
    width: 184,
    height: 100,
    label: "Avg. returns peak",
    value: "Trust* = 7.07",
    tone: "gold",
  });
  addStatCard(slide6, {
    left: 742,
    top: 188,
    width: 184,
    height: 100,
    label: "Pooled OLS peak",
    value: "Trust* = 7.61",
    tone: "blue",
  });
  addStatCard(slide6, {
    left: 944,
    top: 188,
    width: 184,
    height: 100,
    label: "CRE peak",
    value: "Trust* = 6.59",
    tone: "clay",
  });
  addBox(slide6, {
    left: 540,
    top: 322,
    width: 588,
    height: 142,
    fill: colors.white,
    line: { width: 1, fill: colors.line },
    radius: 22,
  });
  addBulletList(slide6, 566, 346, 540, [
    "CRE estimates keep trust and trust squared jointly significant while allowing a persistent return component.",
    "The estimated persistent component is comparable in shape to empirical return heterogeneity found in other datasets.",
    "The RE restriction is rejected; CRE is the preferred panel specification.",
  ], { fontSize: 19, height: 108 });
  addBox(slide6, {
    left: 540,
    top: 490,
    width: 588,
    height: 98,
    fill: colors.paleBlue,
    line: { width: 1, fill: colors.line },
    radius: 22,
  });
  addBox(slide6, {
    left: 564,
    top: 512,
    width: 540,
    height: 56,
    text:
      "Robustness: once portfolio-share controls are added, the fixed-effect distribution changes little and the trust result survives.",
    fontSize: 23,
    color: colors.navy,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.1,
  });
  addFooter(slide6);

  const slide7 = deck.slides.add();
  slide7.background.fill = colors.paper;
  addTitle(slide7, "Chapter 3", "Wealth inequality under ambiguity", 7);
  addBox(slide7, {
    left: 96,
    top: 178,
    width: 460,
    height: 122,
    text:
      "Motivation: if wealth is tied to ownership rights, institutions, and social history, then ranking wealth distributions should allow for ambiguity, not only objective risk.",
    fontSize: 29,
    color: colors.navy,
    typeface: fonts.serif,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.05,
  });
  addBox(slide7, {
    left: 96,
    top: 352,
    width: 474,
    height: 172,
    fill: colors.white,
    line: { width: 1, fill: colors.line },
    radius: 22,
  });
  addBox(slide7, {
    left: 124,
    top: 380,
    width: 418,
    height: 28,
    text: "Rawls-inspired move",
    fontSize: 18,
    color: colors.blue,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBulletList(slide7, 124, 420, 410, [
    "Treat socially relevant states, such as race, as part of the ranking problem.",
    "Use choice under ambiguity to build a welfare-based wealth inequality measure.",
    "Compare it directly with the objective-uncertainty benchmark used in the standard Atkinson-style approach.",
  ], { fontSize: 19, height: 104 });
  addBox(slide7, {
    left: 640,
    top: 192,
    width: 500,
    height: 330,
    fill: colors.paleGold,
    line: { width: 1, fill: colors.line },
    radius: 28,
  });
  addBox(slide7, {
    left: 676,
    top: 228,
    width: 432,
    height: 30,
    text: "Proposed measure",
    fontSize: 18,
    color: colors.blue,
    bold: true,
    alignment: "center",
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBox(slide7, {
    left: 700,
    top: 278,
    width: 384,
    height: 72,
    text: "IA = 1 - wEDE / μ",
    fontSize: 40,
    color: colors.navy,
    typeface: fonts.serif,
    bold: true,
    alignment: "center",
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBox(slide7, {
    left: 690,
    top: 380,
    width: 404,
    height: 94,
    text:
      "Interpretation: the equally distributed equivalent wealth level falls when ambiguity about socially relevant states matters, pushing measured inequality up.",
    fontSize: 22,
    color: colors.ink,
    alignment: "center",
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.14,
  });
  addFooter(slide7);

  const slide8 = deck.slides.add();
  slide8.background.fill = colors.paper;
  addTitle(slide8, "Chapter 3 Result", "Ambiguity-based wealth inequality is systematically higher and tracks racial wealth-gap dynamics", 8);
  await addImage(slide8, "ch3_gap_page164_cropped.png", {
    left: 60,
    top: 180,
    width: 525,
    height: 400,
  });
  await addImage(slide8, "ch3_measures_page166_cropped.png", {
    left: 640,
    top: 180,
    width: 560,
    height: 400,
  });
  addBox(slide8, {
    left: 74,
    top: 592,
    width: 514,
    height: 50,
    text: "Figure 3.2: the median Black-White wealth gap is positive and persistent over time.",
    fontSize: 17,
    color: colors.muted,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addBox(slide8, {
    left: 648,
    top: 592,
    width: 520,
    height: 50,
    text: "Figure 3.3: even at low curvature, IA stays above the objective benchmark IO.",
    fontSize: 17,
    color: colors.muted,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addFooter(slide8);

  const slide9 = deck.slides.add();
  slide9.background.fill = colors.paper;
  addTitle(slide9, "Synthesis", "The three chapters connect through one unifying lesson", 9);
  addBox(slide9, {
    left: 96,
    top: 170,
    width: 1080,
    height: 72,
    text:
      "Rates of return matter for wealth inequality at three margins: distributional explanation, realized performance, and welfare-based measurement.",
    fontSize: 32,
    color: colors.navy,
    typeface: fonts.serif,
    bold: true,
    alignment: "center",
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  const synthCards = [
    {
      left: 104,
      title: "Explaining concentration",
      body:
        "Structural macro models need heterogeneity in returns to reproduce empirical wealth moments and realistic MPC implications.",
      fill: colors.paleBlue,
    },
    {
      left: 430,
      title: "Explaining performance",
      body:
        "Moderate trust is associated with higher returns to net wealth, while persistent household-specific return heterogeneity remains economically important.",
      fill: colors.paleGold,
    },
    {
      left: 756,
      title: "Explaining measurement",
      body:
        "Once socially relevant ambiguity is admitted, standard objective inequality measures understate wealth inequality in the data.",
      fill: "#F6E1DB",
    },
  ];
  for (const card of synthCards) {
    addBox(slide9, {
      left: card.left,
      top: 294,
      width: 276,
      height: 210,
      fill: card.fill,
      line: { width: 1, fill: colors.line },
      radius: 22,
    });
    addBox(slide9, {
      left: card.left + 22,
      top: 320,
      width: 232,
      height: 56,
      text: card.title,
      fontSize: 24,
      color: colors.navy,
      bold: true,
      line: "none",
      fill: "none",
      insets: { top: 0, right: 0, bottom: 0, left: 0 },
    });
    addBox(slide9, {
      left: card.left + 22,
      top: 394,
      width: 232,
      height: 92,
      text: card.body,
      fontSize: 18,
      color: colors.ink,
      line: "none",
      fill: "none",
      insets: { top: 0, right: 0, bottom: 0, left: 0 },
      lineSpacing: 1.15,
    });
  }
  addBox(slide9, {
    left: 180,
    top: 560,
    width: 920,
    height: 46,
    text:
      "Big picture: wealth inequality is not only about who saves more. It is also about who earns different returns, why, and how we judge the inequality those returns create.",
    fontSize: 24,
    color: colors.navy,
    bold: true,
    alignment: "center",
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addFooter(slide9);

  const slide10 = deck.slides.add();
  slide10.background.fill = colors.paper;
  addTitle(slide10, "Conclusion", "Main takeaways for the defense", 10);
  addBox(slide10, {
    left: 92,
    top: 172,
    width: 1040,
    height: 84,
    text:
      "Across the dissertation, heterogeneous returns are central to understanding wealth dynamics in the data and to evaluating wealth inequality in theory.",
    fontSize: 33,
    color: colors.navy,
    typeface: fonts.serif,
    bold: true,
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
    lineSpacing: 1.05,
  });
  addBulletList(slide10, 102, 300, 960, [
    "Chapter 1: return heterogeneity can match key SCF wealth moments in a structural macro model, with an implied return distribution close to empirical evidence.",
    "Chapter 2: returns to net wealth display a hump-shaped relationship with trust, peaking in a moderate range around 6 to 7, while persistent heterogeneity remains after richer controls.",
    "Chapter 3: an ambiguity-based wealth inequality measure implies more inequality than the objective benchmark for the same wealth distribution and co-moves with the Black-White median wealth gap.",
  ], { fontSize: 23, height: 224, lineSpacing: 1.18 });
  addBox(slide10, {
    left: 310,
    top: 580,
    width: 660,
    height: 52,
    text: "Thank you. I look forward to your questions.",
    fontSize: 30,
    color: colors.navy,
    bold: true,
    alignment: "center",
    line: "none",
    fill: "none",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addFooter(slide10);

  const pptxBlob = await PresentationFile.exportPptx(deck);
  await fs.writeFile(OUTPUT_PPTX, Buffer.from(pptxBlob.data));
  console.log(`Wrote ${OUTPUT_PPTX}`);
}

await buildDeck();
