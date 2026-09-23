// Builds the finance CV as a Word document.
// Usage: node build_cv.js  ->  output/Noah_Sargeant_CV_Finance.docx
//
// Phone number and address are read from contact.json (gitignored, because this
// repo is public). Without it, the CV is built with highlighted placeholders.

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, TabStopType,
  BorderStyle, LevelFormat, ExternalHyperlink,
} = require("docx");

// ---------- contact details (private) ----------
let contact = { phone: "[PHONE]", email: "noahsargeant858@gmail.com", location: "Sheffield / Liverpool", linkedin: "" };
const contactFile = path.join(__dirname, "contact.json");
if (fs.existsSync(contactFile)) contact = { ...contact, ...JSON.parse(fs.readFileSync(contactFile, "utf8")) };

// ---------- layout constants ----------
const FONT = "Calibri";
const NAVY = "1F3A5F";
const PAGE_W = 11906;          // A4
const MARGIN = 794;            // ~1.4 cm
const TEXT_W = PAGE_W - 2 * MARGIN;
const BODY = 20;               // 10pt (half-points)

// Text with [square-bracket] placeholders highlighted yellow so they can't be missed.
function runs(text, opts = {}) {
  return text.split(/(\[[^\]]+\])/).filter(Boolean).map((part) =>
    new TextRun({ text: part, font: FONT, size: BODY, ...opts, ...(part.startsWith("[") ? { highlight: "yellow" } : {}) }));
}

function section(title) {
  return new Paragraph({
    spacing: { before: 170, after: 70 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 2 } },
    children: [new TextRun({ text: title.toUpperCase(), font: FONT, size: 22, bold: true, color: NAVY, characterSpacing: 20 })],
  });
}

// Bold left heading with right-aligned dates.
function entry(left, right, { before = 90 } = {}) {
  return new Paragraph({
    spacing: { before, after: 0 },
    keepNext: true,
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W }],
    children: [...runs(left, { bold: true }), ...(right ? [new TextRun({ text: "\t", font: FONT }), ...runs(right, { bold: true })] : [])],
  });
}

function sub(text) {
  return new Paragraph({ spacing: { after: 30 }, keepNext: true, children: runs(text, { italics: true, color: "404040" }) });
}

function bullet(text, lead) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 25, line: 250 },
    children: [...(lead ? runs(lead, { bold: true }) : []), ...runs(text)],
  });
}

function para(text, opts = {}) {
  return new Paragraph({ spacing: { after: 40, line: 260 }, alignment: AlignmentType.JUSTIFIED, children: runs(text, opts) });
}

// ---------- header ----------
const contactLine = [contact.location, contact.phone, contact.email].filter(Boolean);
const header = [
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 30 },
    children: [new TextRun({ text: "NOAH SARGEANT", font: FONT, size: 40, bold: true, color: NAVY, characterSpacing: 40 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 60 },
    children: [
      ...contactLine.flatMap((c, i) => [...(i ? [new TextRun({ text: "  |  ", font: FONT, size: BODY, color: "808080" })] : []), ...runs(c)]),
      // LinkedIn only appears once contact.json has a "linkedin" URL.
      ...(contact.linkedin
        ? [new TextRun({ text: "  |  ", font: FONT, size: BODY, color: "808080" }),
          new ExternalHyperlink({ link: contact.linkedin, children: [new TextRun({ text: contact.linkedin.replace(/^https?:\/\/(www\.)?/, ""), font: FONT, size: BODY, style: "Hyperlink" })] })]
        : []),
    ],
  }),
];

// ---------- content ----------
const body = [
  section("Profile"),
  para("Final-year Business with Finance student at Liverpool John Moores University, seeking a graduate role in finance. " +
    "Passed the Alpha Futures evaluation to trade a $50,000 funded account, returning 8% in three months within strict 1% daily and " +
    "2% weekly drawdown limits. Grounded in investment appraisal (NPV, IRR, WACC, CAPM) with a First in Business Analytics, and " +
    "experienced in directing AI agents to build research and tracking tools. Brings the composure and client focus developed " +
    "serving guests at a Michelin-starred restaurant."),

  section("Education"),
  entry("Liverpool John Moores University — BSc (Hons) Business with Finance", "2024 – 2027", { before: 40 }),
  sub("On track for a 2:1"),
  bullet(" full investment appraisal of a £1m, five-year project. Calculated WACC across seven gearing levels to identify the optimal " +
    "capital structure, built inflation-adjusted cash flows with 14% writing-down allowances and lagged tax, and advised the board " +
    "using NPV, IRR and payback.", "Investment & Financial Analysis:"),
  bullet(" derived a share's beta from its returns against the FTSE (covariance / variance), used CAPM to set a project discount " +
    "rate and critically evaluated CAPM's limits for appraisal. Also appraised trade-credit policy changes and presented a " +
    "£100,000 client portfolio recommendation covering risk, return, tax and the Efficient Markets Hypothesis.", "Risk & portfolio analysis:"),
  bullet(" First Class. Excel-intensive module in data analysis and business decision-making.", "Business Analytics —"),
  bullet(" evaluating a £350m North Sea investment for a listed oil producer, covering cost of capital, regression-based cost " +
    "estimation, oil-price and FX sensitivity, and a financing recommendation.", "International Corporate Finance (in progress):"),

  entry("Silverdale School Sixth Form, Sheffield — A Levels", "2021 – 2023"),
  sub("Business & Economics, History, Media"),
  entry("Silverdale School, Sheffield — 9 GCSEs", "2016 – 2021"),
  sub("Including History (7), Business (6), Maths (5) and English Language (4)"),

  section("Trading & Markets Experience"),
  entry("Independent Forex Trader — Alpha Futures funded account", "Apr 2025 – Present", { before: 40 }),
  bullet("Passed Alpha Futures' evaluation to qualify for a $50,000 funded account, hitting the profit target within the firm's risk rules."),
  bullet("Returned 8% in three months, trading only GBP/USD and EUR/GBP — a deliberately narrow focus on two sterling markets."),
  bullet("Worked to a 1% maximum daily and 2% maximum weekly drawdown, treating capital preservation as the first constraint on every trade."),
  bullet("Set a weekly directional bias from higher-timeframe points of interest, then timed entries with EMAs and trend-line structure."),
  bullet("Planned exposure around high-impact economic releases using the Forex Factory calendar, combining macro news with technical analysis."),

  section("AI & Automation"),
  bullet("Use Claude Code, Anthropic's agentic AI tool, across multiple GitHub repositories — directing an AI agent through structured, " +
    "multi-step tasks rather than one-off chat prompts."),
  bullet("Built a version-controlled study system for International Corporate Finance that turns lecture material into notes, formula keys " +
    "and worked appraisals; checked AI output against the source material, which caught two conflicting assessment deadlines."),
  bullet("Directed the build of interactive web tools, including a graduate job tracker covering 40+ employers with live deadline " +
    "countdowns and eligibility filters matched to my qualifications."),
  bullet("Connected AI to Gmail, Google Calendar and Google Drive, and choose model, effort and permission settings to suit each task."),
  bullet("Use AI within university academic-integrity rules: generated practice questions to self-teach WACC, writing-down allowances and " +
    "IRR, disclosed all AI use, and kept every submitted calculation and judgement my own."),

  section("Client Service Experience"),
  // Reverse chronological. Dates are approximate (worked back from Noah's age at each job).
  entry("Front of House / Barista — Devonshire Arms, Beeley", "2023 – 2024", { before: 40 }),
  bullet("Rotated across bar, floor and barista roles in a busy gastro-pub, balancing competing priorities under pressure."),
  entry("Front of House / Bar Staff — Brocco on the Park, Sheffield", "2022 – 2023"),
  bullet("Led the main bar in a boutique hotel and restaurant; fully trained in cocktail preparation and high-standard drinks service at pace."),
  entry("Café Assistant — Chatsworth Carriage House Café", "2022"),
  bullet("Handled cash and card payments at the till and worked across front- and back-of-house in a high-footfall visitor attraction."),
  entry("Waiter / Host — Fischers at Baslow Hall (Michelin-starred)", "Summer 2021"),
  bullet("Guided guests through a Michelin-starred tasting menu, tailoring wine-pairing recommendations to each guest from detailed product knowledge."),
  bullet("Ran a 20-cover restaurant floor independently during quieter shifts, owning service standards from greeting to close."),

  section("Skills"),
  bullet(" DCF appraisal (NPV, IRR, payback), WACC and optimal capital structure, CAPM and beta, portfolio theory, working-capital analysis.", "Financial analysis:"),
  bullet(" technical analysis (EMAs, trend lines, points of interest), macroeconomic news analysis, drawdown and position risk management.", "Markets:"),
  bullet(" Microsoft Excel, Microsoft Office, Claude Code, GitHub.", "Tools:"),
  bullet(" full UK driving licence with own vehicle.", "Other:"),

  section("Interests"),
  para("Financial markets and economic news; health and fitness, keeping a disciplined gym routine and structured diet; music and fashion."),
];

const doc = new Document({
  creator: "Noah Sargeant",
  title: "Noah Sargeant — CV",
  styles: { default: { document: { run: { font: FONT, size: BODY } } } },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 300, hanging: 220 } } } }],
    }],
  },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 680, bottom: 680, left: MARGIN, right: MARGIN } } },
    children: [...header, ...body],
  }],
});

const out = path.join(__dirname, "output", "Noah_Sargeant_CV_Finance.docx");
fs.mkdirSync(path.dirname(out), { recursive: true });
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("Wrote " + out); });
