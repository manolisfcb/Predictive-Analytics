const pptxgen = require("pptxgenjs");
const {
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require("./pptxgenjs_helpers/layout");

const TEAM_NAMES = "[Add team names]";
const PRESENTATION_DATE = "March 16, 2026";
const OUTPUT_FILE = "MercadoLibre_Final_Project_Presentation.pptx";

const COLORS = {
  bg: "F6F1E8",
  ink: "1B2A41",
  forest: "1F5C4D",
  coral: "E76F51",
  gold: "D6A64F",
  teal: "2A9D8F",
  sky: "DCEAF2",
  mint: "DCEFE9",
  sand: "EFE4D0",
  rose: "F7E1D9",
  smoke: "5C6770",
  white: "FFFFFF",
  line: "D8D1C4",
};

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Codex";
pptx.company = "UNF";
pptx.subject = "Mercado Libre final project presentation";
pptx.title = "Marketplace Dynamics: Price Prediction and Condition Classification";
pptx.lang = "en-US";
pptx.theme = {
  headFontFace: "Trebuchet MS",
  bodyFontFace: "Aptos",
  lang: "en-US",
};

function addBackdrop(slide) {
  slide.background = { color: COLORS.bg };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.33,
    h: 0.38,
    line: { color: COLORS.forest, transparency: 100 },
    fill: { color: COLORS.forest },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0.38,
    w: 13.33,
    h: 0.06,
    line: { color: COLORS.coral, transparency: 100 },
    fill: { color: COLORS.coral },
  });
}

function addHeader(slide, number, kicker, title, subtitle) {
  addBackdrop(slide);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.65,
    y: 0.62,
    w: 1.2,
    h: 0.34,
    rectRadius: 0.06,
    line: { color: COLORS.sand, transparency: 100 },
    fill: { color: COLORS.sand },
  });
  slide.addText(number, {
    x: 0.65,
    y: 0.665,
    w: 1.2,
    h: 0.18,
    fontFace: "Trebuchet MS",
    fontSize: 11,
    bold: true,
    align: "center",
    color: COLORS.forest,
    margin: 0,
  });

  if (kicker) {
    slide.addText(kicker.toUpperCase(), {
      x: 2.05,
      y: 0.64,
      w: 3.2,
      h: 0.2,
      fontFace: "Aptos",
      fontSize: 10,
      bold: true,
      color: COLORS.coral,
      charSpace: 0.6,
      margin: 0,
    });
  }

  slide.addText(title, {
    x: 0.65,
    y: 1.0,
    w: 8.8,
    h: 0.72,
    fontFace: "Trebuchet MS",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.67,
      y: 1.78,
      w: 8.9,
      h: 0.38,
      fontFace: "Aptos",
      fontSize: 11,
      color: COLORS.smoke,
      margin: 0,
    });
  }
}

function addSpeakerBadge(slide, speaker) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 11.55,
    y: 0.62,
    w: 1.15,
    h: 0.34,
    rectRadius: 0.06,
    line: { color: COLORS.mint, transparency: 100 },
    fill: { color: COLORS.mint },
  });
  slide.addText(speaker, {
    x: 11.55,
    y: 0.67,
    w: 1.15,
    h: 0.18,
    fontFace: "Aptos",
    fontSize: 10,
    bold: true,
    align: "center",
    color: COLORS.forest,
    margin: 0,
  });
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.7,
    y: 7.06,
    w: 11.2,
    h: 0.16,
    fontFace: "Aptos",
    fontSize: 8,
    color: COLORS.smoke,
    margin: 0,
  });
}

function addBulletList(slide, items, x, y, w, opts = {}) {
  const gap = opts.gap || 0.52;
  const fontSize = opts.fontSize || 15;
  items.forEach((item, idx) => {
    const top = y + idx * gap;
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: top + 0.11,
      w: 0.11,
      h: 0.11,
      line: { color: opts.bulletColor || COLORS.coral, transparency: 100 },
      fill: { color: opts.bulletColor || COLORS.coral },
    });
    slide.addText(item, {
      x: x + 0.22,
      y: top,
      w,
      h: 0.32,
      fontFace: "Aptos",
      fontSize,
      color: opts.color || COLORS.ink,
      bold: !!opts.bold,
      margin: 0,
    });
  });
}

function addQuoteBand(slide, text, y) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.85,
    y,
    w: 11.7,
    h: 0.72,
    rectRadius: 0.08,
    line: { color: COLORS.rose, transparency: 100 },
    fill: { color: COLORS.rose },
  });
  slide.addText(text, {
    x: 1.1,
    y: y + 0.15,
    w: 11.2,
    h: 0.35,
    fontFace: "Georgia",
    italic: true,
    fontSize: 15,
    color: COLORS.ink,
    align: "center",
    margin: 0,
  });
}

function addMetricCard(slide, x, y, w, h, label, value, subtext, fill) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    line: { color: fill, transparency: 100 },
    fill: { color: fill },
  });
  slide.addText(label.toUpperCase(), {
    x: x + 0.18,
    y: y + 0.14,
    w: w - 0.36,
    h: 0.12,
    fontFace: "Aptos",
    fontSize: 8.5,
    bold: true,
    color: COLORS.smoke,
    charSpace: 0.4,
    margin: 0,
  });
  slide.addText(value, {
    x: x + 0.18,
    y: y + 0.34,
    w: w - 0.36,
    h: 0.28,
    fontFace: "Trebuchet MS",
    fontSize: 20,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  if (subtext) {
    slide.addText(subtext, {
      x: x + 0.18,
      y: y + h - 0.24,
      w: w - 0.36,
      h: 0.14,
      fontFace: "Aptos",
      fontSize: 8.5,
      color: COLORS.smoke,
      margin: 0,
    });
  }
}

function addTwoColumnComparison(slide, rows, leftTitle, rightTitle) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.88,
    y: 2.25,
    w: 11.6,
    h: 3.85,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1.1 },
    fill: { color: COLORS.white },
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1.1,
    y: 2.48,
    w: 5.35,
    h: 0.46,
    rectRadius: 0.05,
    line: { color: COLORS.sand, transparency: 100 },
    fill: { color: COLORS.sand },
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.86,
    y: 2.48,
    w: 4.95,
    h: 0.46,
    rectRadius: 0.05,
    line: { color: COLORS.mint, transparency: 100 },
    fill: { color: COLORS.mint },
  });
  slide.addText(leftTitle, {
    x: 1.2,
    y: 2.61,
    w: 5.1,
    h: 0.18,
    fontFace: "Trebuchet MS",
    fontSize: 16,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText(rightTitle, {
    x: 6.98,
    y: 2.61,
    w: 4.7,
    h: 0.18,
    fontFace: "Trebuchet MS",
    fontSize: 16,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  rows.forEach((row, idx) => {
    const top = 3.12 + idx * 0.58;
    slide.addText(row.label.toUpperCase(), {
      x: 1.22,
      y: top,
      w: 0.9,
      h: 0.16,
      fontFace: "Aptos",
      fontSize: 8.5,
      bold: true,
      color: COLORS.smoke,
      charSpace: 0.5,
      margin: 0,
    });
    slide.addText(row.left, {
      x: 2.3,
      y: top - 0.02,
      w: 3.55,
      h: 0.24,
      fontFace: "Aptos",
      fontSize: 14,
      color: COLORS.ink,
      margin: 0,
    });
    slide.addText(row.right, {
      x: 7.16,
      y: top - 0.02,
      w: 3.55,
      h: 0.24,
      fontFace: "Aptos",
      fontSize: 14,
      color: COLORS.ink,
      margin: 0,
    });
    slide.addShape(pptx.ShapeType.line, {
      x: 1.2,
      y: top + 0.34,
      w: 10.5,
      h: 0,
      line: { color: COLORS.line, pt: 0.8 },
    });
  });
}

function addRankedBars(slide, items, x, y, w, maxValue) {
  items.forEach((item, idx) => {
    const top = y + idx * 0.52;
    slide.addText(item.label, {
      x,
      y: top,
      w: 2.15,
      h: 0.2,
      fontFace: "Aptos",
      fontSize: 12,
      color: COLORS.ink,
      margin: 0,
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 2.25,
      y: top + 0.02,
      w: w,
      h: 0.18,
      rectRadius: 0.03,
      line: { color: COLORS.sand, transparency: 100 },
      fill: { color: COLORS.sand },
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 2.25,
      y: top + 0.02,
      w: Math.max(0.2, w * (item.value / maxValue)),
      h: 0.18,
      rectRadius: 0.03,
      line: { color: item.color || COLORS.forest, transparency: 100 },
      fill: { color: item.color || COLORS.forest },
    });
    slide.addText(item.valueLabel, {
      x: x + 2.25 + w + 0.15,
      y: top - 0.02,
      w: 0.7,
      h: 0.22,
      fontFace: "Aptos",
      fontSize: 12,
      bold: true,
      color: COLORS.ink,
      margin: 0,
    });
  });
}

function finalizeSlide(slide) {
  warnIfSlideHasOverlaps(slide, pptx, { ignoreDecorativeShapes: true });
  warnIfSlideElementsOutOfBounds(slide, pptx);
}

function slide1() {
  const slide = pptx.addSlide();
  addBackdrop(slide);

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.75,
    y: 0.8,
    w: 0.16,
    h: 2.55,
    line: { color: COLORS.coral, transparency: 100 },
    fill: { color: COLORS.coral },
  });

  slide.addText("Marketplace Dynamics", {
    x: 1.15,
    y: 0.92,
    w: 6.7,
    h: 0.5,
    fontFace: "Trebuchet MS",
    fontSize: 27,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText("Price Prediction and Product Condition Classification", {
    x: 1.15,
    y: 1.48,
    w: 7.15,
    h: 0.78,
    fontFace: "Trebuchet MS",
    fontSize: 21,
    bold: true,
    color: COLORS.forest,
    margin: 0,
  });
  slide.addText("Mercado Libre final project presentation", {
    x: 1.15,
    y: 2.34,
    w: 6.2,
    h: 0.24,
    fontFace: "Georgia",
    italic: true,
    fontSize: 14,
    color: COLORS.smoke,
    margin: 0,
  });

  addMetricCard(slide, 8.8, 0.95, 1.65, 1.1, "Listings", "100k", "Real marketplace data", COLORS.sand);
  addMetricCard(slide, 10.62, 0.95, 1.65, 1.1, "Top model", "0.82", "RF accuracy", COLORS.mint);
  addMetricCard(slide, 8.8, 2.23, 1.65, 1.1, "Regression", "0.36", "ElasticNet R^2", COLORS.sky);
  addMetricCard(slide, 10.62, 2.23, 1.65, 1.1, "Insight", "Behavior", "Seller signals dominate", COLORS.rose);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1.12,
    y: 4.25,
    w: 11.15,
    h: 1.65,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.white },
  });

  const info = [
    { label: "Project", value: "Mercado Libre marketplace ML project", x: 1.35, y: 4.55, w: 3.2 },
    { label: "Course", value: "DAMO510", x: 4.75, y: 4.55, w: 1.4 },
    { label: "Team", value: TEAM_NAMES, x: 6.3, y: 4.55, w: 2.6 },
    { label: "Date", value: PRESENTATION_DATE, x: 9.15, y: 4.55, w: 2.35 },
  ];

  info.forEach((item) => {
    slide.addText(item.label.toUpperCase(), {
      x: item.x,
      y: item.y,
      w: item.w,
      h: 0.16,
      fontFace: "Aptos",
      fontSize: 9,
      bold: true,
      color: COLORS.smoke,
      charSpace: 0.5,
      margin: 0,
    });
    slide.addText(item.value, {
      x: item.x,
      y: item.y + 0.26,
      w: item.w,
      h: 0.34,
      fontFace: "Trebuchet MS",
      fontSize: 16,
      bold: true,
      color: COLORS.ink,
      margin: 0,
    });
  });

  addFooter(slide, "Speaker notes included in PowerPoint presenter view.");
  slide.addNotes(
    "Person 1\n" +
      "Today we present a machine learning project focused on understanding marketplace dynamics through price prediction and product condition classification."
  );
  finalizeSlide(slide);
}

function slide2() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "02",
    "Problem & Motivation",
    "Why this marketplace problem matters",
    "Prediction is useful only if it also reveals the marketplace logic behind each listing."
  );
  addSpeakerBadge(slide, "Person 1");

  slide.addText("Core business problems", {
    x: 0.9,
    y: 2.35,
    w: 3.2,
    h: 0.22,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  addBulletList(
    slide,
    [
      "Predict listing price in a noisy, heterogeneous marketplace.",
      "Classify products as new or used from imperfect operational signals.",
      "Move beyond prediction toward marketplace understanding.",
    ],
    0.92,
    2.78,
    4.6,
    { fontSize: 15 }
  );

  const cards = [
    {
      x: 6.1,
      title: "Trust",
      body: "Better condition signals reduce buyer uncertainty and misleading listings.",
      fill: COLORS.sand,
    },
    {
      x: 8.4,
      title: "Search Ranking",
      body: "Condition and price quality directly affect relevance, ranking, and conversion.",
      fill: COLORS.mint,
    },
    {
      x: 10.7,
      title: "Seller Behavior",
      body: "Patterns in listings reveal who is acting casually versus professionally.",
      fill: COLORS.sky,
    },
  ];
  cards.forEach((card) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: card.x,
      y: 2.45,
      w: 1.95,
      h: 2.35,
      rectRadius: 0.08,
      line: { color: card.fill, transparency: 100 },
      fill: { color: card.fill },
    });
    slide.addText(card.title, {
      x: card.x + 0.15,
      y: 2.7,
      w: 1.65,
      h: 0.22,
      fontFace: "Trebuchet MS",
      fontSize: 15,
      bold: true,
      color: COLORS.ink,
      align: "center",
      margin: 0,
    });
    slide.addText(card.body, {
      x: card.x + 0.16,
      y: 3.18,
      w: 1.62,
      h: 1.25,
      fontFace: "Aptos",
      fontSize: 11.5,
      color: COLORS.ink,
      align: "center",
      valign: "mid",
      margin: 0.02,
    });
  });

  addQuoteBand(
    slide,
    "We wanted not only to predict, but to understand how marketplaces actually behave.",
    5.95
  );
  slide.addNotes(
    "Person 1\n" +
      "The marketplace is complex because price and condition are shaped by many operational choices.\n" +
      "This matters for trust, search ranking, and understanding seller behavior.\n" +
      "We wanted not only to predict, but to understand how marketplaces actually behave."
  );
  finalizeSlide(slide);
}

function slide3() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "03",
    "Dataset Overview",
    "A real production-style marketplace dataset",
    "The dataset combines conventional tabular features with nested marketplace metadata."
  );
  addSpeakerBadge(slide, "Person 1");

  addMetricCard(slide, 0.95, 2.25, 2.35, 1.18, "Rows", "100,000", "Mercado Libre listings", COLORS.sand);
  addMetricCard(slide, 3.5, 2.25, 2.35, 1.18, "Raw columns", "45", "Before initial pruning", COLORS.mint);
  addMetricCard(slide, 6.05, 2.25, 2.35, 1.18, "After drop", "27", "Core working dataset", COLORS.sky);
  addMetricCard(slide, 8.6, 2.25, 2.35, 1.18, "Data type", "Mixed", "Structured + JSON fields", COLORS.rose);

  slide.addText("Feature families", {
    x: 0.98,
    y: 4.0,
    w: 2.5,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  const families = [
    { text: "Pricing", fill: COLORS.sand, x: 1.0, y: 4.45 },
    { text: "Seller", fill: COLORS.mint, x: 2.65, y: 4.45 },
    { text: "Inventory", fill: COLORS.sky, x: 4.2, y: 4.45 },
    { text: "Shipping", fill: COLORS.rose, x: 6.0, y: 4.45 },
    { text: "Listing type", fill: COLORS.sand, x: 7.75, y: 4.45 },
    { text: "Nested JSON", fill: COLORS.mint, x: 9.55, y: 4.45 },
  ];
  families.forEach((chip) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: chip.x,
      y: chip.y,
      w: 1.35,
      h: 0.42,
      rectRadius: 0.06,
      line: { color: chip.fill, transparency: 100 },
      fill: { color: chip.fill },
    });
    slide.addText(chip.text, {
      x: chip.x,
      y: chip.y + 0.12,
      w: 1.35,
      h: 0.15,
      fontFace: "Aptos",
      fontSize: 10.5,
      bold: true,
      align: "center",
      color: COLORS.ink,
      margin: 0,
    });
  });

  slide.addText("Condition split", {
    x: 0.98,
    y: 5.45,
    w: 2.3,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1.0,
    y: 5.9,
    w: 9.8,
    h: 0.34,
    rectRadius: 0.06,
    line: { color: COLORS.line, transparency: 100 },
    fill: { color: COLORS.line },
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1.0,
    y: 5.9,
    w: 5.27,
    h: 0.34,
    rectRadius: 0.06,
    line: { color: COLORS.forest, transparency: 100 },
    fill: { color: COLORS.forest },
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.27,
    y: 5.9,
    w: 4.53,
    h: 0.34,
    rectRadius: 0.06,
    line: { color: COLORS.coral, transparency: 100 },
    fill: { color: COLORS.coral },
  });
  slide.addText("New: 53,758 (53.8%)", {
    x: 1.15,
    y: 6.05,
    w: 2.6,
    h: 0.18,
    fontFace: "Aptos",
    fontSize: 11,
    bold: true,
    color: COLORS.forest,
    margin: 0,
  });
  slide.addText("Used: 46,242 (46.2%)", {
    x: 8.4,
    y: 6.05,
    w: 2.3,
    h: 0.18,
    fontFace: "Aptos",
    fontSize: 11,
    bold: true,
    color: COLORS.coral,
    margin: 0,
  });

  slide.addNotes(
    "Person 1\n" +
      "We worked with a real Mercado Libre dataset of about one hundred thousand listings.\n" +
      "It mixes structured fields with nested JSON, which makes it much closer to production data than a clean classroom dataset."
  );
  finalizeSlide(slide);
}

function slide4() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "04",
    "Key Challenges",
    "Real data comes with real production messiness",
    "The main challenge was not model choice alone, but making the data usable without cheating."
  );
  addSpeakerBadge(slide, "Person 2");

  const challenges = [
    {
      x: 0.95,
      y: 2.35,
      title: "Skewed price",
      body: "Median price is 250, but the maximum reaches 2.22B. The tail is extreme.",
      fill: COLORS.sand,
    },
    {
      x: 6.65,
      y: 2.35,
      title: "High cardinality",
      body: "category_id explodes to 10,907 unique values, making direct encoding difficult.",
      fill: COLORS.mint,
    },
    {
      x: 0.95,
      y: 4.2,
      title: "Data leakage",
      body: "base_price is almost a direct copy of price with correlation equal to 1.00.",
      fill: COLORS.sky,
    },
    {
      x: 6.65,
      y: 4.2,
      title: "Nested JSON",
      body: "seller_address, shipping, tags, and attributes had to be unpacked into usable signals.",
      fill: COLORS.rose,
    },
  ];

  challenges.forEach((card) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: card.x,
      y: card.y,
      w: 5.5,
      h: 1.45,
      rectRadius: 0.08,
      line: { color: card.fill, transparency: 100 },
      fill: { color: card.fill },
    });
    slide.addText(card.title, {
      x: card.x + 0.18,
      y: card.y + 0.2,
      w: 2.3,
      h: 0.22,
      fontFace: "Trebuchet MS",
      fontSize: 16,
      bold: true,
      color: COLORS.ink,
      margin: 0,
    });
    slide.addText(card.body, {
      x: card.x + 0.18,
      y: card.y + 0.55,
      w: 5.05,
      h: 0.6,
      fontFace: "Aptos",
      fontSize: 13,
      color: COLORS.ink,
      margin: 0,
    });
  });

  addQuoteBand(
    slide,
    "This dataset is not clean - it reflects real production complexity.",
    6.0
  );
  slide.addNotes(
    "Person 2\n" +
      "This dataset is not clean. It reflects real production complexity.\n" +
      "We had to handle skewed prices, thousands of categories, leakage risk, and multiple JSON-like fields."
  );
  finalizeSlide(slide);
}

function slide5() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "05",
    "Data Preparation",
    "Feature engineering turned raw marketplace traces into model-ready signals",
    "Most of the lift came from designing meaningful features, not from using the fanciest algorithm."
  );
  addSpeakerBadge(slide, "Person 2");

  const steps = [
    "Drop sparse and irrelevant fields.",
    "Parse seller and shipping JSON fields.",
    "Engineer listing_age_days and count-based features.",
    "Convert warranty into a binary signal.",
    "Remove leakage-heavy variables before training.",
  ];
  slide.addText("Pipeline", {
    x: 0.95,
    y: 2.2,
    w: 2.0,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  steps.forEach((step, idx) => {
    const top = 2.65 + idx * 0.66;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 1.0,
      y: top - 0.02,
      w: 0.34,
      h: 0.34,
      line: { color: COLORS.forest, transparency: 100 },
      fill: { color: COLORS.forest },
    });
    slide.addText(String(idx + 1), {
      x: 1.0,
      y: top + 0.05,
      w: 0.34,
      h: 0.12,
      fontFace: "Aptos",
      fontSize: 9,
      bold: true,
      align: "center",
      color: COLORS.white,
      margin: 0,
    });
    slide.addText(step, {
      x: 1.48,
      y: top,
      w: 4.25,
      h: 0.22,
      fontFace: "Aptos",
      fontSize: 14,
      color: COLORS.ink,
      margin: 0,
    });
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.55,
    y: 2.25,
    w: 5.6,
    h: 3.65,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText("Engineered features", {
    x: 6.8,
    y: 2.5,
    w: 2.5,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  const feats = [
    "listing_age_days",
    "shipping_free",
    "shipping_pickup",
    "shipping_mode",
    "seller_state",
    "tag_count",
    "attribute_count",
    "has_warranty",
  ];
  feats.forEach((feat, idx) => {
    const chipX = 6.82 + (idx % 2) * 2.45;
    const chipY = 3.0 + Math.floor(idx / 2) * 0.56;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: chipX,
      y: chipY,
      w: 2.1,
      h: 0.38,
      rectRadius: 0.05,
      line: { color: idx % 2 === 0 ? COLORS.mint : COLORS.sky, transparency: 100 },
      fill: { color: idx % 2 === 0 ? COLORS.mint : COLORS.sky },
    });
    slide.addText(feat, {
      x: chipX,
      y: chipY + 0.11,
      w: 2.1,
      h: 0.15,
      fontFace: "Aptos",
      fontSize: 10.3,
      bold: true,
      align: "center",
      color: COLORS.ink,
      margin: 0,
    });
  });

  slide.addText("Removed or controlled: base_price, category_id, sparse fields, raw JSON containers.", {
    x: 6.82,
    y: 5.28,
    w: 5.0,
    h: 0.42,
    fontFace: "Aptos",
    fontSize: 12.5,
    color: COLORS.smoke,
    margin: 0,
  });

  addQuoteBand(
    slide,
    "Most of the model performance came from feature engineering, not the raw data.",
    6.05
  );
  slide.addNotes(
    "Person 2\n" +
      "Most of the model performance came from feature engineering, not the raw data.\n" +
      "We created features such as listing_age_days, shipping signals, seller location, and warranty indicators, while removing leakage-prone variables."
  );
  finalizeSlide(slide);
}

function slide6() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "06",
    "Regression (Price)",
    "ElasticNet captured part of the pricing signal, but the task remained hard",
    "Price prediction is fundamentally limited when product-level semantics are missing."
  );
  addSpeakerBadge(slide, "Person 2");

  addMetricCard(slide, 0.95, 2.3, 2.25, 1.35, "Model", "ElasticNet", "On log-transformed price", COLORS.sand);
  addMetricCard(slide, 3.45, 2.3, 2.1, 1.35, "MAE", "1.07", "Lower is better", COLORS.mint);
  addMetricCard(slide, 5.72, 2.3, 2.1, 1.35, "RMSE", "1.39", "Log-price scale", COLORS.sky);
  addMetricCard(slide, 7.99, 2.3, 2.1, 1.35, "R^2", "0.36", "Moderate explanatory power", COLORS.rose);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.95,
    y: 4.15,
    w: 5.2,
    h: 1.78,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText("Why it is difficult", {
    x: 1.18,
    y: 4.4,
    w: 2.3,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  addBulletList(
    slide,
    [
      "Listings span many product categories and seller strategies.",
      "Important product-level details were not modeled from text or images.",
      "The price distribution is heavily skewed even after log transformation.",
    ],
    1.18,
    4.8,
    3.9,
    { fontSize: 12.7, bulletColor: COLORS.gold, gap: 0.35 }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.55,
    y: 4.15,
    w: 5.6,
    h: 1.78,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText("Interpretation", {
    x: 6.8,
    y: 4.4,
    w: 2.4,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText(
    "The model explains broad marketplace trends, but exact price remains hard to predict because much of the product-specific context is still missing.",
    {
      x: 6.8,
      y: 4.8,
      w: 4.95,
      h: 0.65,
      fontFace: "Aptos",
      fontSize: 13.3,
      color: COLORS.ink,
      margin: 0,
    }
  );

  slide.addNotes(
    "Person 2\n" +
      "For regression we used ElasticNet.\n" +
      "We obtained an MAE around 1.07, RMSE around 1.39, and R squared around 0.36.\n" +
      "Price is difficult to predict because the dataset lacks richer product-level information."
  );
  finalizeSlide(slide);
}

function slide7() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "07",
    "Classification Models",
    "Comparing linear and non-linear classifiers",
    "Non-linear models better captured the structure hidden in seller and listing behavior."
  );
  addSpeakerBadge(slide, "Person 3");

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.95,
    y: 2.2,
    w: 7.55,
    h: 3.7,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText("Accuracy comparison", {
    x: 1.2,
    y: 2.45,
    w: 2.8,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  const modelRows = [
    { label: "Random Forest", value: 0.8199, valueLabel: "0.82", color: COLORS.forest },
    { label: "Gradient Boosting", value: 0.8115, valueLabel: "0.81", color: COLORS.teal },
    { label: "KNN", value: 0.80395, valueLabel: "0.80", color: COLORS.gold },
    { label: "Logistic Regression", value: 0.77575, valueLabel: "0.78", color: COLORS.coral },
  ];
  addRankedBars(slide, modelRows, 1.2, 3.0, 3.5, 0.85);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.82,
    y: 2.2,
    w: 3.35,
    h: 3.7,
    rectRadius: 0.08,
    line: { color: COLORS.sky, transparency: 100 },
    fill: { color: COLORS.sky },
  });
  slide.addText("Main message", {
    x: 9.08,
    y: 2.5,
    w: 2.4,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  addBulletList(
    slide,
    [
      "Tree ensembles outperformed the linear baseline.",
      "The gap suggests the problem is non-linear.",
      "Feature interactions matter for condition classification.",
    ],
    9.08,
    3.0,
    2.45,
    { fontSize: 13.1, bulletColor: COLORS.forest, gap: 0.54 }
  );

  slide.addNotes(
    "Person 3\n" +
      "We compared linear and non-linear models to understand the complexity of the task.\n" +
      "Random Forest performed best with about 82 percent accuracy, followed closely by Gradient Boosting."
  );
  finalizeSlide(slide);
}

function slide8() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "08",
    "ROC & PR Results",
    "Random Forest was the strongest model across threshold-based metrics",
    "The ranking stayed consistent even when we looked beyond plain accuracy."
  );
  addSpeakerBadge(slide, "Person 3");

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1.0,
    y: 2.2,
    w: 5.25,
    h: 3.45,
    rectRadius: 0.08,
    line: { color: COLORS.sand, transparency: 100 },
    fill: { color: COLORS.sand },
  });
  slide.addText("ROC-AUC", {
    x: 1.28,
    y: 2.55,
    w: 1.5,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 18,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText("0.91", {
    x: 1.28,
    y: 3.0,
    w: 1.7,
    h: 0.48,
    fontFace: "Trebuchet MS",
    fontSize: 30,
    bold: true,
    color: COLORS.forest,
    margin: 0,
  });
  slide.addText("Random Forest", {
    x: 1.3,
    y: 3.6,
    w: 2.0,
    h: 0.2,
    fontFace: "Aptos",
    fontSize: 14,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText("Notebook interpretation: all models stayed above 0.85 AUC, but Random Forest achieved the strongest separation between new and used listings.", {
    x: 1.28,
    y: 4.05,
    w: 4.45,
    h: 0.95,
    fontFace: "Aptos",
    fontSize: 12.6,
    color: COLORS.ink,
    margin: 0,
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.75,
    y: 2.2,
    w: 5.25,
    h: 3.45,
    rectRadius: 0.08,
    line: { color: COLORS.mint, transparency: 100 },
    fill: { color: COLORS.mint },
  });
  slide.addText("PR-AUC", {
    x: 7.02,
    y: 2.55,
    w: 1.5,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 18,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText("Highest", {
    x: 7.02,
    y: 3.0,
    w: 2.0,
    h: 0.42,
    fontFace: "Trebuchet MS",
    fontSize: 28,
    bold: true,
    color: COLORS.forest,
    margin: 0,
  });
  slide.addText("Random Forest", {
    x: 7.04,
    y: 3.58,
    w: 2.0,
    h: 0.2,
    fontFace: "Aptos",
    fontSize: 14,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText("Precision-recall is especially important here because we care about correctly finding used products, and Random Forest had the best balance across thresholds.", {
    x: 7.02,
    y: 4.05,
    w: 4.45,
    h: 0.95,
    fontFace: "Aptos",
    fontSize: 12.6,
    color: COLORS.ink,
    margin: 0,
  });

  addQuoteBand(
    slide,
    "Random Forest performs best across all metrics, especially in detecting used products.",
    6.0
  );
  slide.addNotes(
    "Person 3\n" +
      "Random Forest performs best across all metrics, especially in detecting used products.\n" +
      "Its ROC-AUC reached 0.91, and it also achieved the strongest precision-recall behavior."
  );
  finalizeSlide(slide);
}

function slide9() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "09",
    "Key Insights",
    "The model learns seller behavior more than direct product condition",
    "This is the most important slide because it changes the interpretation of the whole project."
  );
  addSpeakerBadge(slide, "Person 4");

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.95,
    y: 2.2,
    w: 5.15,
    h: 3.9,
    rectRadius: 0.08,
    line: { color: COLORS.rose, transparency: 100 },
    fill: { color: COLORS.rose },
  });
  slide.addText("Boom moment", {
    x: 1.25,
    y: 2.55,
    w: 2.0,
    h: 0.2,
    fontFace: "Aptos",
    fontSize: 10,
    bold: true,
    color: COLORS.coral,
    charSpace: 0.6,
    margin: 0,
  });
  slide.addText("The model is not detecting 'used' directly - it is learning seller behavior.", {
    x: 1.25,
    y: 2.92,
    w: 4.45,
    h: 1.2,
    fontFace: "Trebuchet MS",
    fontSize: 21,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  addBulletList(
    slide,
    [
      "Low inventory strongly points to used listings.",
      "No warranty is another strong used-product signal.",
      "Free listing types are heavily associated with casual sellers.",
    ],
    1.25,
    4.5,
    3.8,
    { fontSize: 13.1, bulletColor: COLORS.forest, gap: 0.43 }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.45,
    y: 2.2,
    w: 5.75,
    h: 3.9,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText("Top Random Forest drivers", {
    x: 6.72,
    y: 2.52,
    w: 3.1,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  addRankedBars(
    slide,
    [
      { label: "listing_type_id_free", value: 0.242377, valueLabel: "0.242", color: COLORS.coral },
      { label: "available_quantity", value: 0.225296, valueLabel: "0.225", color: COLORS.forest },
      { label: "initial_quantity", value: 0.201578, valueLabel: "0.202", color: COLORS.teal },
      { label: "sold_quantity", value: 0.090722, valueLabel: "0.091", color: COLORS.gold },
      { label: "listing_age_days", value: 0.049315, valueLabel: "0.049", color: COLORS.smoke },
    ],
    6.72,
    3.0,
    1.85,
    0.25
  );

  slide.addNotes(
    "Person 4\n" +
      "The model is not detecting used directly - it is learning seller behavior.\n" +
      "Signals like low inventory, no warranty, and free listings tell us more about the seller profile than about the item alone."
  );
  finalizeSlide(slide);
}

function slide10() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "10",
    "Marketplace Insight",
    "From classification to behavioral segmentation",
    "The classifier can be reinterpreted as a segmentation model for seller types."
  );
  addSpeakerBadge(slide, "Person 4");

  addTwoColumnComparison(
    slide,
    [
      { label: "Condition", left: "Used", right: "New" },
      { label: "Inventory", left: "Low inventory", right: "High inventory" },
      { label: "Listing type", left: "Free listing", right: "Paid listing" },
      { label: "Warranty", left: "Often none", right: "Often present" },
      { label: "Seller profile", left: "Casual seller", right: "Professional seller" },
    ],
    "Casual sellers",
    "Professional sellers"
  );

  addQuoteBand(
    slide,
    "This is essentially a behavioral segmentation model.",
    6.15
  );
  slide.addNotes(
    "Person 4\n" +
      "This is essentially a behavioral segmentation model.\n" +
      "On one side we see casual sellers with used items, low inventory, and free listings. On the other side we see professional sellers with new items and stronger operational signals."
  );
  finalizeSlide(slide);
}

function slide11() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "11",
    "Conclusion",
    "What we learned from the project",
    "The project delivered both predictive models and a clearer view of how marketplace behavior is encoded in data."
  );
  addSpeakerBadge(slide, "Person 4");

  const takeaways = [
    {
      x: 0.98,
      title: "Tree models > linear",
      body: "Random Forest and Gradient Boosting captured the non-linear marketplace structure better than Logistic Regression.",
      fill: COLORS.sand,
    },
    {
      x: 4.5,
      title: "Feature engineering = key",
      body: "Performance improved because raw JSON and operational fields were transformed into meaningful behavioral signals.",
      fill: COLORS.mint,
    },
    {
      x: 8.02,
      title: "Behavior > product features",
      body: "The strongest signals described seller activity and listing strategy more than product condition itself.",
      fill: COLORS.sky,
    },
  ];
  takeaways.forEach((card) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: card.x,
      y: 2.4,
      w: 3.1,
      h: 3.2,
      rectRadius: 0.08,
      line: { color: card.fill, transparency: 100 },
      fill: { color: card.fill },
    });
    slide.addText(card.title, {
      x: card.x + 0.18,
      y: 2.72,
      w: 2.72,
      h: 0.5,
      fontFace: "Trebuchet MS",
      fontSize: 18,
      bold: true,
      color: COLORS.ink,
      align: "center",
      margin: 0,
    });
    slide.addText(card.body, {
      x: card.x + 0.2,
      y: 3.58,
      w: 2.68,
      h: 1.35,
      fontFace: "Aptos",
      fontSize: 12.5,
      color: COLORS.ink,
      align: "center",
      valign: "mid",
      margin: 0.02,
    });
  });

  slide.addNotes(
    "Person 4\n" +
      "Our main conclusion is that tree models outperformed linear ones, feature engineering was the key lever, and behavioral patterns carried more information than product-specific attributes."
  );
  finalizeSlide(slide);
}

function slide12() {
  const slide = pptx.addSlide();
  addHeader(
    slide,
    "12",
    "Reflection",
    "Limitations and future work",
    "The next step is to combine behavioral signals with richer product representations."
  );
  addSpeakerBadge(slide, "Person 4");

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1.0,
    y: 2.25,
    w: 5.2,
    h: 3.25,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText("Limitations", {
    x: 1.25,
    y: 2.55,
    w: 2.0,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  addBulletList(
    slide,
    [
      "No NLP over titles or descriptions.",
      "No image or embedding-based product representation.",
      "Single marketplace snapshot with limited seller history.",
    ],
    1.25,
    3.0,
    3.95,
    { fontSize: 13.2, bulletColor: COLORS.coral, gap: 0.55 }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.55,
    y: 2.25,
    w: 5.2,
    h: 3.25,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText("Future directions", {
    x: 6.8,
    y: 2.55,
    w: 2.4,
    h: 0.2,
    fontFace: "Trebuchet MS",
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  addBulletList(
    slide,
    [
      "Add NLP on titles and descriptions.",
      "Use embeddings to capture product similarity.",
      "Integrate richer seller behavior and temporal history.",
    ],
    6.8,
    3.0,
    3.95,
    { fontSize: 13.2, bulletColor: COLORS.forest, gap: 0.55 }
  );

  addQuoteBand(
    slide,
    "Questions?",
    6.05
  );
  slide.addNotes(
    "Person 4\n" +
      "Our main limitations were the lack of NLP and richer product representations.\n" +
      "Future work could add text models, embeddings, and deeper seller history."
  );
  finalizeSlide(slide);
}

async function main() {
  slide1();
  slide2();
  slide3();
  slide4();
  slide5();
  slide6();
  slide7();
  slide8();
  slide9();
  slide10();
  slide11();
  slide12();

  await pptx.writeFile({ fileName: OUTPUT_FILE });
  console.log(`Wrote ${OUTPUT_FILE}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
