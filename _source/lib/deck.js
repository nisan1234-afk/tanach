// Shared deck engine for תנ"ך לבגרות — Hebrew RTL PowerPoint decks
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const PAL = {
  ink: "1F2A44",     // deep indigo — dominant
  ink2: "2E3D60",
  gold: "D4A24C",    // accent
  goldSoft: "F6EBD5",
  tint: "EEF1F7",    // card tint
  white: "FFFFFF",
  text: "1F2A44",
  muted: "5B6478",
  five: "8E3B46",    // 5-units marker
};
const FONT = "Arial";
const W = 10, H = 5.625;

async function icon(name, color = PAL.white, size = 256) {
  const C = fa[name];
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(C, { color: "#" + color, size }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Hebrew typography: geresh/gershayim inside words so bidi never flips them
function heb(t) {
  if (typeof t !== "string") return t;
  return t.replace(/([\u05D0-\u05EA])"(?=[\u05D0-\u05EA])/g, "$1\u05F4").replace(/([\u05D0-\u05EA])'/g, "$1\u05F3");
}
function patchSlide(slide) {
  const orig = slide.addText.bind(slide);
  slide.addText = (text, opts) => orig(Array.isArray(text) ? text.map(r => Object.assign({}, r, { text: heb(r.text) })) : heb(text), opts);
  const origT = slide.addTable.bind(slide);
  slide.addTable = (rows, opts) => origT(rows.map(r => r.map(c => (c && typeof c === "object") ? Object.assign({}, c, { text: heb(c.text) }) : heb(c))), opts);
  return slide;
}

function T(extra = {}) {
  return Object.assign({ fontFace: FONT, rtlMode: true, align: "right", isTextBox: true, color: PAL.text, valign: "top", margin: 0.05 }, extra);
}

class Deck {
  constructor({ title, unitLabel, fileName }) {
    this.pres = new pptxgen();
    this.pres.layout = "LAYOUT_16x9";
    this.pres.rtlMode = true;
    this.pres.title = title;
    this.pres.author = "תנ\"ך לבגרות";
    this.unitLabel = unitLabel;
    this.fileName = fileName;
    this.n = 0;
    this.icons = {};
    const add = this.pres.addSlide.bind(this.pres);
    this.pres.addSlide = (...a) => patchSlide(add(...a));
  }
  async ic(name, color) {
    const k = name + color;
    if (!this.icons[k]) this.icons[k] = await icon(name, color);
    return this.icons[k];
  }
  footer(s, dark = false) {
    this.n++;
    s.addText(this.unitLabel, T({ x: 0.5, y: H - 0.42, w: 6, h: 0.3, fontSize: 10, color: dark ? "AAB3C8" : PAL.muted }));
    s.addText(String(this.n), T({ x: 0.4, y: H - 0.42, w: 0.6, h: 0.3, fontSize: 10, align: "left", rtlMode: false, color: dark ? "AAB3C8" : PAL.muted }));
  }
  title(s, text, dark = false) {
    s.addText(text, T({ x: 0.5, y: 0.35, w: 9, h: 0.75, fontSize: 28, bold: true, color: dark ? PAL.white : PAL.ink, valign: "middle" }));
  }
  async iconCircle(s, name, x, y, d = 0.6, bg = PAL.ink, fg = PAL.white) {
    s.addShape(this.pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
    s.addImage({ data: await this.ic(name, fg), x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56 });
  }
  fiveTag(s, x, y, label = "5 יח\"ל") {
    s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.0, h: 0.32, fill: { color: PAL.five }, line: { color: PAL.five }, rectRadius: 0.08 });
    s.addText(label, T({ x, y, w: 1.0, h: 0.32, fontSize: 12, bold: true, color: PAL.white, align: "center", valign: "middle" }));
  }

  // ---------- slide types ----------
  async cover({ title, sub, meta }) {
    const s = this.pres.addSlide();
    s.background = { color: PAL.ink };
    await this.iconCircle(s, "FaScroll", 8.3, 0.7, 1.1, PAL.gold, PAL.ink);
    s.addText(title, T({ x: 0.6, y: 1.9, w: 8.8, h: 1.1, fontSize: 40, bold: true, color: PAL.white }));
    s.addText(sub, T({ x: 0.6, y: 3.0, w: 8.8, h: 0.6, fontSize: 22, color: PAL.gold }));
    s.addText(meta, T({ x: 0.6, y: 3.7, w: 8.8, h: 0.5, fontSize: 14, color: "C9D0E0" }));
    this.footer(s, true);
    return s;
  }
  async roadmap({ title, steps }) {
    // steps: [{n, label}] — up to 10, two rows of five, RTL order
    const s = this.pres.addSlide();
    s.background = { color: PAL.white };
    this.title(s, title);
    const per = 5, cw = 1.66, ch = 1.45, gap = 0.12, x0 = W - 0.5 - cw;
    steps.forEach((st, i) => {
      const r = Math.floor(i / per), c = i % per;
      const x = x0 - c * (cw + gap), y = 1.3 + r * (ch + 0.2);
      s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: st.five ? "F4E6E8" : PAL.tint }, line: { color: st.five ? "F4E6E8" : PAL.tint }, rectRadius: 0.1 });
      s.addShape(this.pres.shapes.OVAL, { x: x + cw - 0.55, y: y + 0.12, w: 0.42, h: 0.42, fill: { color: st.five ? PAL.five : PAL.ink }, line: { color: st.five ? PAL.five : PAL.ink } });
      s.addText(String(st.n), T({ x: x + cw - 0.55, y: y + 0.12, w: 0.42, h: 0.42, fontSize: 14, bold: true, color: PAL.white, align: "center", valign: "middle", rtlMode: false }));
      s.addText(st.label, T({ x: x + 0.1, y: y + 0.6, w: cw - 0.2, h: ch - 0.7, fontSize: 13, bold: true }));
    });
    this.footer(s);
    return s;
  }
  async segment({ n, total, title, ref, five }) {
    const s = this.pres.addSlide();
    s.background = { color: PAL.ink };
    s.addText(`מקטע ${n} מתוך ${total}`, T({ x: 0.6, y: 1.3, w: 8.8, h: 0.5, fontSize: 18, color: PAL.gold, bold: true }));
    s.addText(title, T({ x: 0.6, y: 1.85, w: 8.8, h: 1.2, fontSize: 36, bold: true, color: PAL.white }));
    if (ref) s.addText(ref, T({ x: 0.6, y: 3.1, w: 8.8, h: 0.5, fontSize: 18, color: "C9D0E0" }));
    if (five) this.fiveTag(s, 0.6, 1.4);
    this.footer(s, true);
    return s;
  }
  // verse card + points beside it
  async verse({ title, verse, ref, points = [], icon = "FaBookOpen", five, note }) {
    const s = this.pres.addSlide();
    s.background = { color: PAL.white };
    this.title(s, title);
    if (five) this.fiveTag(s, 0.5, 0.55);
    // verse card on the right
    s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, { x: 5.0, y: 1.3, w: 4.5, h: 3.5, fill: { color: PAL.ink }, line: { color: PAL.ink }, rectRadius: 0.12 });
    s.addImage({ data: await this.ic("FaQuoteRight", PAL.gold), x: 8.85, y: 1.45, w: 0.45, h: 0.45 });
    s.addText(verse, T({ x: 5.25, y: 1.95, w: 4.0, h: 2.3, fontSize: 19, bold: true, color: PAL.white, valign: "middle", lineSpacingMultiple: 1.15 }));
    s.addText(ref, T({ x: 5.25, y: 4.25, w: 4.0, h: 0.4, fontSize: 13, color: PAL.gold }));
    // points on the left
    let y = 1.35;
    for (const p of points) {
      await this.iconCircle(s, p.icon || icon, 4.0, y, 0.5, PAL.gold, PAL.ink);
      s.addText(p.head, T({ x: 0.5, y: y - 0.02, w: 3.35, h: 0.38, fontSize: 16, bold: true }));
      s.addText(p.body, T({ x: 0.5, y: y + 0.36, w: 3.35, h: 0.72, fontSize: 14, color: PAL.muted }));
      y += 1.15;
    }
    if (note) s.addText(note, T({ x: 0.5, y: 4.65, w: 4.0, h: 0.35, fontSize: 12, italic: true, color: PAL.muted }));
    this.footer(s);
    return s;
  }
  // two (or three) commentator cards side by side, RTL: first card is rightmost
  async cards({ title, cards, five, lead, bottom }) {
    const s = this.pres.addSlide();
    s.background = { color: PAL.white };
    this.title(s, title);
    if (five) this.fiveTag(s, 0.5, 0.55);
    let top = 1.25;
    if (lead) { s.addText(lead, T({ x: 0.5, y: 1.15, w: 9, h: 0.5, fontSize: 15, color: PAL.muted })); top = 1.7; }
    const n = cards.length, gap = 0.3, cw = (9 - gap * (n - 1)) / n;
    const ch = n === 1 ? 2.6 : (bottom ? 4.2 : 4.85) - top;
    for (let i = 0; i < n; i++) {
      const c = cards[i];
      const x = W - 0.5 - cw - i * (cw + gap);
      const tagged = c.five;
      s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, { x, y: top, w: cw, h: ch, fill: { color: tagged ? "F7EDEF" : PAL.tint }, line: { color: tagged ? "F7EDEF" : PAL.tint }, rectRadius: 0.1 });
      await this.iconCircle(s, c.icon || "FaUserGraduate", x + cw - 0.75, top + 0.18, 0.55, tagged ? PAL.five : PAL.ink, PAL.white);
      s.addText(c.head, T({ x: x + 0.2, y: top + 0.2, w: cw - 1.05, h: 0.5, fontSize: 18, bold: true, color: tagged ? PAL.five : PAL.ink, valign: "middle" }));
      if (c.tag) s.addText(c.tag, T({ x: x + 0.2, y: top + 0.72, w: cw - 0.4, h: 0.32, fontSize: 12, italic: true, color: PAL.muted }));
      const by = top + (c.tag ? 1.05 : 0.85);
      const items = (Array.isArray(c.body) ? c.body : [c.body]).map((b, j, a) => ({ text: b, options: { bullet: a.length > 1 ? { indent: 12 } : false, breakLine: j < a.length - 1, paraSpaceAfter: 6, rtlMode: true, align: "right" } }));
      s.addText(items, T({ x: x + 0.2, y: by, w: cw - 0.4, h: top + ch - by - 0.15, fontSize: n > 2 ? 14 : 15 }));
    }
    if (bottom) {
      s.addShape(this.pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.3, w: 9, h: 0.6, fill: { color: PAL.goldSoft }, line: { color: PAL.goldSoft }, rectRadius: 0.08 });
      s.addText(bottom, T({ x: 0.7, y: 4.3, w: 8.6, h: 0.6, fontSize: 14, bold: true, valign: "middle" }));
    }
    this.footer(s);
    return s;
  }
  async table({ title, head, rows, five, lead, colW }) {
    const s = this.pres.addSlide();
    s.background = { color: PAL.white };
    this.title(s, title);
    if (five) this.fiveTag(s, 0.5, 0.55);
    let y = 1.25;
    if (lead) { s.addText(lead, T({ x: 0.5, y: 1.15, w: 9, h: 0.45, fontSize: 15, color: PAL.muted })); y = 1.65; }
    const hdr = head.map(h => ({ text: h, options: { bold: true, color: PAL.white, fill: { color: PAL.ink }, align: "right", rtlMode: true, fontFace: FONT, fontSize: 15 } }));
    const body = rows.map((r, i) => r.map(c => ({ text: c, options: { color: PAL.text, fill: { color: i % 2 ? PAL.white : PAL.tint }, align: "right", rtlMode: true, fontFace: FONT, fontSize: 14 } })));
    const rev = a => a.slice().reverse();
    s.addTable([rev(hdr), ...body.map(rev)], { x: 0.5, y, w: 9, colW: rev(colW || head.map(() => 9 / head.length)), border: { type: "solid", pt: 0.75, color: "D5DAE5" }, margin: 0.08, rowH: 0.42 });
    this.footer(s);
    return s;
  }
  async practice({ q, topic, hint }) {
    const s = this.pres.addSlide();
    s.background = { color: PAL.gold };
    await this.iconCircle(s, "FaPencilAlt", 4.45, 0.8, 1.1, PAL.ink, PAL.gold);
    s.addText("תרגול בדף העבודה", T({ x: 0.5, y: 2.05, w: 9, h: 0.7, fontSize: 30, bold: true, align: "center", color: PAL.ink }));
    s.addText(q, T({ x: 0.5, y: 2.75, w: 9, h: 0.8, fontSize: 40, bold: true, align: "center", color: PAL.white }));
    s.addText(topic, T({ x: 0.5, y: 3.6, w: 9, h: 0.5, fontSize: 18, align: "center", color: PAL.ink }));
    if (hint) s.addText(hint, T({ x: 0.5, y: 4.1, w: 9, h: 0.45, fontSize: 14, italic: true, align: "center", color: PAL.ink }));
    this.footer(s);
    return s;
  }
  async write() {
    await this.pres.writeFile({ fileName: this.fileName });
    return this.fileName;
  }
}

module.exports = { Deck, PAL, T, FONT, W, H };
