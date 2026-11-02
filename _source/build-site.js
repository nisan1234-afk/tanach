// Builds the static GitHub Pages site into ./site from site.config.js and the unit folders.
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");
const cfg = require("./site.config");

const OUT = path.join(__dirname, "site");
const SOFFICE = "/root/.claude/skills/synced/36d484e1-ee5a-4c87-8d66-0260e587552e_6e81539c-d0eb-4321-860f-f0380d298029/pptx/scripts/office/soffice.py";
const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const he = s => esc(String(s).replace(/([א-ת])"(?=[א-ת])/g, "$1״").replace(/([א-ת])'/g, "$1׳"));
const mkdir = d => fs.mkdirSync(d, { recursive: true });

function page({ title, desc, canonical, body, depth }) {
  const up = depth ? "../".repeat(depth) : "";
  return `<!doctype html>
<html lang="he" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${he(title)}</title>
<meta name="description" content="${he(desc)}">
<link rel="canonical" href="${canonical}">
<meta property="og:title" content="${he(title)}">
<meta property="og:description" content="${he(desc)}">
<meta property="og:type" content="website">
<meta property="og:locale" content="he_IL">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Frank+Ruhl+Libre:wght@500;700;900&family=Heebo:wght@400;500;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="${up}assets/style.css">
</head>
<body>
<header class="top"><a class="brand" href="${up}index.html">תנ״ך לבגרות</a><span class="code">שאלון 2381</span></header>
<main>
${body}
</main>
<footer class="foot">
<p>חומרי הוראה לתנ״ך ממלכתי־דתי (חמ״ד), כיתה י״א — מצגות, דפי עבודה ומחוונים, על בסיס חומרי משרד החינוך ושאלות מבחינות בגרות.</p>
</footer>
</body>
</html>
`;
}

function renderUnitAssets(book, u, dstDir) {
  const src = path.join(__dirname, book.slug, u.slug);
  const files = { pptx: `${u.file}.pptx`, ws: `${u.file}-worksheet.docx`, key: `${u.file}-key.docx` };
  for (const f of Object.values(files)) fs.copyFileSync(path.join(src, f), path.join(dstDir, f));
  // PDFs for printing + slide previews
  for (const f of [files.ws, files.key, files.pptx]) {
    execSync(`python3 ${SOFFICE} --headless --convert-to pdf --outdir "${dstDir}" "${path.join(dstDir, f)}"`, { stdio: "ignore" });
  }
  const prev = path.join(dstDir, "slides"); mkdir(prev);
  execSync(`rm -f "${prev}"/*.jpg && pdftoppm -jpeg -jpegopt quality=82 -scale-to-x 960 -scale-to-y -1 "${path.join(dstDir, u.file + ".pdf")}" "${prev}/s"`);
  const slides = fs.readdirSync(prev).filter(f => f.endsWith(".jpg")).sort();
  // the deck PDF itself is also useful as a fallback viewer
  return { files, slides };
}

function unitCard(book, u, depth) {
  const up = depth ? "../".repeat(depth) : "";
  if (u.ready) {
    return `<li class="unit ready"><a href="${up}${book.slug}/${u.slug}/index.html"><span class="num">${u.n}</span><span class="t">${he(u.title)}</span><span class="r">${he(u.ref)}</span></a></li>`;
  }
  return `<li class="unit soon"><span class="num">${u.n}</span><span class="t">${he(u.title)}</span><span class="r">${he(u.ref)}, בהכנה</span></li>`;
}

function build() {
  fs.rmSync(OUT, { recursive: true, force: true });
  mkdir(path.join(OUT, "assets"));
  fs.copyFileSync(path.join(__dirname, "assets", "style.css"), path.join(OUT, "assets", "style.css"));
  const urls = [cfg.baseUrl];

  // ---------- home ----------
  const bookBlocks = cfg.books.map(b => {
    const ready = b.units.filter(u => u.ready).length;
    return `<section class="book" id="${b.slug}">
  <div class="book-head"><h2>${he(b.name)}</h2><p>${he(b.blurb)}</p><p class="count">${ready ? `${ready} מתוך ${b.units.length} יחידות מוכנות` : `${b.units.length} יחידות — בהכנה`}</p></div>
  <ol class="units">${b.units.map(u => unitCard(b, u, 0)).join("")}</ol>
</section>`;
  }).join("\n");
  const home = `
<section class="hero">
  <h1>תנ״ך לבגרות</h1>
  <p class="lead">לכל יחידה: מצגת להקרנה בכיתה, דף עבודה להדפסה עם שאלות מבחינות בגרות אמיתיות, ומחוון למורה.</p>
  <ol class="how">
    <li><strong>בוחרים יחידה</strong> מהרשימה למטה.</li>
    <li><strong>מקרינים את המצגת</strong> — היא מחולקת לקטעים קצרים.</li>
    <li><strong>כשמופיע שקף צהוב "תרגול"</strong> — התלמידים עונים על השאלה עם אותו מספר בדף העבודה.</li>
  </ol>
</section>
${bookBlocks}`;
  fs.writeFileSync(path.join(OUT, "index.html"), page({ title: "תנ״ך לבגרות — שאלון 2381 | מצגות ודפי עבודה לכיתה י״א", desc: "מצגות, דפי עבודה ומחוונים לתנ״ך חמ״ד כיתה י״א (שאלון 2381): בראשית, ישעיהו ומלכים ב, יונה ותהילים, מגילת רות — עם שאלות מבחינות בגרות.", canonical: cfg.baseUrl, body: home, depth: 0 }));

  // ---------- unit pages ----------
  for (const b of cfg.books) {
    for (const u of b.units.filter(x => x.ready)) {
      const dir = path.join(OUT, b.slug, u.slug); mkdir(dir);
      const { files, slides } = renderUnitAssets(b, u, dir);
      const base = files.pptx.replace(/\.pptx$/, "");
      const segs = u.segments.map((s, i) => `<li><span class="q">שאלה ${i + 1}</span><span class="st">${he(s[0])}</span><span class="src">${he(s[1])}</span></li>`).join("");
      const thumbs = slides.map((f, i) => `<a href="slides/${f}" target="_blank" rel="noopener"><img src="slides/${f}" alt="שקף ${i + 1}" loading="lazy" width="960" height="540"></a>`).join("");
      const prevUnit = b.units.find(x => x.n === u.n - 1 && x.ready), nextUnit = b.units.find(x => x.n === u.n + 1 && x.ready);
      const body = `
<nav class="crumbs"><a href="../../index.html">כל היחידות</a> <span aria-hidden="true">/</span> <a href="../../index.html#${b.slug}">${he(b.name)}</a></nav>
<section class="unit-hero">
  <p class="kicker">${he(b.name)}, יחידה ${u.n}</p>
  <h1>${he(u.title)}</h1>
  <p class="ref">${he(b.name.replace("חומש ", ""))} ${he(u.ref)}</p>
</section>
<section class="downloads" aria-label="הורדות">
  <a class="dl main" href="${files.pptx}" download><span class="dl-t">הורדת המצגת</span><span class="dl-s">PowerPoint, ${slides.length} שקפים</span></a>
  <a class="dl" href="${base}-worksheet.pdf" target="_blank" rel="noopener"><span class="dl-t">דף העבודה להדפסה</span><span class="dl-s">PDF</span></a>
  <a class="dl" href="${base}-key.pdf" target="_blank" rel="noopener"><span class="dl-t">המחוון למורה</span><span class="dl-s">PDF</span></a>
  <p class="alt">גרסאות לעריכה ב־Word: <a href="${files.ws}" download>דף העבודה</a>, <a href="${files.key}" download>המחוון</a>, וגם <a href="${base}.pdf" target="_blank" rel="noopener">המצגת כ־PDF</a></p>
</section>
<section class="plan">
  <h2>מבנה היחידה</h2>
  <p>המצגת מחולקת ל־${u.segments.length} קטעים קצרים. בסוף כל קטע — שאלה בדף העבודה.</p>
  <ol class="segs">${segs}</ol>
</section>
<section class="preview">
  <h2>השקפים</h2>
  <div class="grid">${thumbs}</div>
</section>
<nav class="pager">${prevUnit ? `<a href="../${prevUnit.slug}/index.html">יחידה ${prevUnit.n}: ${he(prevUnit.title)}</a>` : "<span></span>"}${nextUnit ? `<a href="../${nextUnit.slug}/index.html">יחידה ${nextUnit.n}: ${he(nextUnit.title)}</a>` : ""}</nav>`;
      const url = `${cfg.baseUrl}${b.slug}/${u.slug}/`;
      urls.push(url);
      fs.writeFileSync(path.join(dir, "index.html"), page({
        title: `${u.title} — ${b.name} ${u.ref} | תנ״ך לבגרות 2381`,
        desc: `מצגת, דף עבודה ומחוון ל${b.name} ${u.ref} (${u.title}) — תנ״ך חמ״ד כיתה י״א, שאלון 2381, עם שאלות מבחינות בגרות.`,
        canonical: url, body, depth: 2,
      }));
    }
  }

  // ---------- SEO files ----------
  const today = new Date().toISOString().slice(0, 10);
  fs.writeFileSync(path.join(OUT, "sitemap.xml"), `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${urls.map(u => `  <url><loc>${u}</loc><lastmod>${today}</lastmod></url>`).join("\n")}\n</urlset>\n`);
  fs.writeFileSync(path.join(OUT, "robots.txt"), `User-agent: *\nAllow: /\nSitemap: ${cfg.baseUrl}sitemap.xml\n`);
  fs.writeFileSync(path.join(OUT, ".nojekyll"), "");
  console.log("built", urls.length, "pages");
}
build();
