// Shared Word (docx) engine for printed worksheets + answer keys — Hebrew RTL
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell, WidthType,
  BorderStyle, ShadingType, Footer, PageNumber, LevelFormat, TableLayoutType,
} = require("docx");

const C = { ink: "1F2A44", gold: "B8862F", goldSoft: "F6EBD5", tint: "EEF1F7", muted: "5B6478", five: "8E3B46", line: "B9C0CF" };
const FONT = "Arial";
const PAGE_W = 11906, MARGIN = 1000, CONTENT_W = PAGE_W - 2 * MARGIN; // A4

function heb(t) {
  return "\u200F" + String(t).replace(/([א-ת])"(?=[א-ת])/g, "$1״").replace(/([א-ת])'/g, "$1׳");
}
function run(text, o = {}) {
  return new TextRun({ text: heb(text), font: FONT, size: (o.size || 12) * 2, sizeComplexScript: (o.size || 12) * 2, bold: o.bold, boldComplexScript: o.bold, italics: o.italic, italicsComplexScript: o.italic, color: o.color || C.ink, rightToLeft: true });
}
function P(children, o = {}) {
  if (typeof children === "string") children = [run(children, o)];
  return new Paragraph({
    children, bidirectional: true, alignment: o.center ? AlignmentType.CENTER : AlignmentType.START,
    spacing: { before: o.before ?? 0, after: o.after ?? 80, line: o.line ?? 300 },
    keepNext: o.keepNext, keepLines: o.keepLines, numbering: o.numbering,
    border: o.border, indent: o.indent,
  });
}
function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    borders: o.borders,
  });
}
const noB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NOBORDERS = { top: noB, bottom: noB, left: noB, right: noB };
// one-cell shaded box (used for "בקצרה" summaries and source tags)
function box(paras, fill) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W], visuallyRightToLeft: true, layout: TableLayoutType.FIXED,
    rows: [new TableRow({ children: [cell(paras, { w: CONTENT_W, fill, borders: NOBORDERS })] })],
  });
}
function answerLines(n) {
  // writing lines = one-column table, bottom border only on each row (paragraph borders merge in Word)
  const lineB = { style: BorderStyle.SINGLE, size: 6, color: C.line };
  const rows = [];
  for (let i = 0; i < n; i++) rows.push(new TableRow({ height: { value: 480, rule: "exact" }, children: [new TableCell({ children: [P(" ", { after: 0 })], width: { size: CONTENT_W, type: WidthType.DXA }, borders: { top: noB, left: noB, right: noB, bottom: lineB } })] }));
  return [new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W], visuallyRightToLeft: true, layout: TableLayoutType.FIXED, borders: { top: noB, bottom: noB, left: noB, right: noB, insideHorizontal: noB, insideVertical: noB }, rows }), P("", { after: 120 })];
}
function gridTable(head, rows, widths) {
  // RTL table: first column rightmost (visuallyRightToLeft)
  const total = widths.reduce((a, b) => a + b, 0);
  const mk = (txt, w, fill, bold) => cell(P(txt || " ", { bold, after: 0 }), { w, fill });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths, visuallyRightToLeft: true, layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({ tableHeader: true, children: head.map((h, i) => mk(h, widths[i], C.tint, true)) }),
      ...rows.map(r => new TableRow({ height: { value: 600, rule: "atLeast" }, children: r.map((c, i) => mk(c, widths[i])) })),
    ],
  });
}

/*
  spec = {
    title, subtitle, unitLabel, fileName, keyFileName,
    intro: [strings],
    segments: [{ n, title, ref, five, summary:[strings], question: {
        source, five, intro:[strings], text (passage, optional), parts: [{ label, text:[strings], points, lines, table:{head,rows,widths}, five }],
        key: [{ label, points:[strings] }] } }]
  }
*/
function buildWorksheet(spec) {
  const body = [];
  body.push(P(spec.title, { size: 22, bold: true, after: 40 }));
  body.push(P(spec.subtitle, { size: 13, color: C.gold, bold: true, after: 60 }));
  body.push(P([run("שם התלמיד/ה: ", { size: 12, bold: true }), run("__________________________   ", { color: C.line }), run("כיתה: ", { size: 12, bold: true }), run("________   ", { color: C.line }), run("תאריך: ", { size: 12, bold: true }), run("____________", { color: C.line })], { after: 160 }));
  for (const t of spec.intro || []) body.push(P(t, { size: 11, color: C.muted, after: 60 }));

  for (const seg of spec.segments) {
    body.push(P([run(`מקטע ${seg.n}`, { size: 11, bold: true, color: C.gold }), run(`   ${seg.ref || ""}`, { size: 11, color: C.muted }), ...(seg.five ? [run("   · לתלמידי 5 יח\"ל", { size: 11, bold: true, color: C.five })] : [])], { before: 280, after: 0, keepNext: true }));
    body.push(P(seg.title, { size: 17, bold: true, after: 100, keepNext: true }));
    body.push(box([P("בקצרה", { size: 11, bold: true, color: C.gold, after: 40, keepNext: true }), ...seg.summary.map(s => P(s, { size: 12, after: 60, keepNext: true, keepLines: true }))], C.tint));
    const q = seg.question;
    body.push(P([run(`שאלה ${seg.n}`, { size: 15, bold: true }), run(`   ${q.source}`, { size: 10, italic: true, color: C.muted }), ...(q.five ? [run("   (5 יח\"ל)", { size: 11, bold: true, color: C.five })] : [])], { before: 200, after: 80, keepNext: true, keepLines: true }));
    for (const t of q.intro || []) body.push(P(t, { size: 12, after: 60, keepNext: true }));
    if (q.text) body.push(box(q.text.map(t => P(t, { size: 12, after: 40 })), C.goldSoft));
    for (const part of q.parts) {
      const lead = [run(`${part.label}. `, { bold: true }), ...(part.five ? [run("(5 יח\"ל) ", { bold: true, color: C.five })] : []), ...(part.extra ? [run("(הרחבה) ", { bold: true, color: C.gold })] : [])];
      part.text.forEach((t, i) => body.push(P(i === 0 ? [...lead, run(t)] : [run(t)], { after: 40, keepNext: true, indent: i ? { right: 280 } : undefined })));
      if (part.points) body.push(P(`(${part.points} נקודות)`, { size: 10, color: C.muted, after: 60, keepNext: true }));
      if (part.table) body.push(gridTable(part.table.head, part.table.rows, part.table.widths), P("", { after: 80 }));
      if (part.lines) body.push(...answerLines(part.lines));
    }
  }

  const doc = new Document({
    creator: "תנ\"ך לבגרות", title: spec.title,
    styles: { default: { document: { run: { font: FONT, size: 24 }, paragraph: { alignment: AlignmentType.START } } } },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 900, bottom: 900, left: MARGIN, right: MARGIN } }, bidi: true },
      footers: { default: new Footer({ children: [new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, children: [run(spec.unitLabel + "   ·   עמ' ", { size: 9, color: C.muted }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: C.muted })] })] }) },
      children: body,
    }],
  });
  return Packer.toBuffer(doc).then(b => { fs.writeFileSync(spec.fileName, b); return spec.fileName; });
}

function buildKey(spec) {
  const body = [];
  body.push(P("מחוון למורה", { size: 13, bold: true, color: C.gold, after: 60 }));
  body.push(P(spec.title.replace(/ — דף עבודה$/, ""), { size: 22, bold: true, after: 40 }));
  body.push(P(spec.subtitle, { size: 12, color: C.gold, bold: true, after: 80 }));
  body.push(P("נקודות התשובה המצופות לכל סעיף. בשאלות מבגרות — על בסיס המחוון הרשמי כשיש כזה; אחרת — לפי חומר ההוראה של משרד החינוך.", { size: 11, color: C.muted, after: 200 }));
  for (const seg of spec.segments) {
    const q = seg.question;
    body.push(P([run(`שאלה ${seg.n} — ${seg.title}`, { size: 15, bold: true }), run(`   ${q.source}`, { size: 10, italic: true, color: C.muted })], { before: 240, after: 80, keepNext: true }));
    for (const k of q.key) {
      body.push(P([run(`${k.label}. `, { bold: true, color: C.gold })], { after: 20, keepNext: true }));
      for (const pt of k.points) body.push(P([run("•  ", { color: C.gold }), run(pt)], { after: 40, indent: { right: 280 } }));
    }
  }
  const doc = new Document({
    creator: "תנ\"ך לבגרות", title: "מחוון — " + spec.title,
    styles: { default: { document: { run: { font: FONT, size: 24 }, paragraph: { alignment: AlignmentType.START } } } },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 900, bottom: 900, left: MARGIN, right: MARGIN } }, bidi: true },
      footers: { default: new Footer({ children: [new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, children: [run("מחוון למורה · " + spec.unitLabel + "   ·   עמ' ", { size: 9, color: C.muted }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: C.muted })] })] }) },
      children: body,
    }],
  });
  return Packer.toBuffer(doc).then(b => { fs.writeFileSync(spec.keyFileName, b); return spec.keyFileName; });
}

module.exports = { buildWorksheet, buildKey };
