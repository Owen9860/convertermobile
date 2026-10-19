/* PDF Toolkit - PDF <-> Word converters. Everything runs offline on the phone.
   Needs: JSZip, PDFLib (pdf-lib) and a pdf.js module passed in by the caller. */
(function (root, factory) {
  root.Convert = factory();
})(typeof self !== "undefined" ? self : this, function () {
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const clean = (s) => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, "");
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const median = (arr) => {
    if (!arr.length) return 0;
    const s = arr.slice().sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
  };
  const pct = (arr, p) => {
    if (!arr.length) return 0;
    const s = arr.slice().sort((a, b) => a - b);
    return s[clamp(Math.floor(s.length * p), 0, s.length - 1)];
  };

  /* =====================================================================
     PDF  ->  WORD (.docx)
     ===================================================================== */

  const NS =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

  function fontFamilyOf(name) {
    const f = (name || "").toLowerCase();
    if (/courier|consolas|mono|menlo|typewriter/.test(f)) return "Courier New";
    if (/times|serif|georgia|cambria|garamond|palatino|minion|book|century|constantia/.test(f) && !/sans/.test(f)) return "Times New Roman";
    if (/calibri/.test(f)) return "Calibri";
    return "Arial";
  }

  /** Pull positioned text lines out of one pdf.js page. */
  async function readPageLines(page) {
    const tc = await page.getTextContent();
    try { await page.getOperatorList(); } catch (_) {} // makes real font names available
    const fontCache = {};
    const fontOf = (id) => {
      if (fontCache[id]) return fontCache[id];
      let name = "", bold = false, italic = false;
      try {
        if (page.commonObjs.has(id)) {
          const f = page.commonObjs.get(id);
          name = f.name || f.loadedName || "";
          bold = !!f.bold; italic = !!f.italic;
        }
      } catch (_) {}
      const n = name.toLowerCase();
      if (/bold|black|heavy|semibold|demi/.test(n)) bold = true;
      if (/italic|oblique/.test(n)) italic = true;
      return (fontCache[id] = { family: fontFamilyOf(name), bold, italic });
    };

    const items = [];
    for (const it of tc.items) {
      if (typeof it.str !== "string" || !it.str || !it.transform) continue;
      const [a, b, , , e, f] = it.transform;
      if (Math.abs(b) > Math.abs(a) * 0.5) continue; // rotated text (watermarks, side labels)
      const size = Math.hypot(a, b) || it.height || 10;
      if (size < 2) continue;
      const fo = fontOf(it.fontName);
      items.push({ s: it.str, x: e, y: f, w: it.width, size, fam: fo.family, bold: fo.bold, italic: fo.italic });
    }
    items.sort((p, q) => q.y - p.y || p.x - q.x);

    const rows = [];
    for (const it of items) {
      let row = null;
      for (let k = rows.length - 1; k >= 0 && k >= rows.length - 4; k--) {
        if (Math.abs(rows[k].y - it.y) <= 0.4 * Math.max(rows[k].size, it.size)) { row = rows[k]; break; }
      }
      if (!row) rows.push((row = { y: it.y, size: it.size, items: [] }));
      row.items.push(it);
    }

    const lines = [];
    for (const row of rows) {
      row.items.sort((p, q) => p.x - q.x);
      const runs = [];
      let prevEnd = null, chars = 0;
      const sizeVotes = {};
      for (const it of row.items) {
        let text = it.s;
        if (prevEnd !== null) {
          const gap = it.x - prevEnd;
          const prevSize = runs.length ? runs[runs.length - 1].size : it.size;
          if (gap > 4 * it.size) text = "\t" + text;
          else if (gap > 0.2 * it.size && !/\s$/.test(runs.length ? runs[runs.length - 1].text : "") && !/^\s/.test(text)) text = " " + text;
        }
        prevEnd = it.x + it.w;
        const last = runs[runs.length - 1];
        const sz = Math.round(it.size * 2) / 2;
        if (last && last.bold === it.bold && last.italic === it.italic && last.size === sz && last.fam === it.fam) last.text += text;
        else runs.push({ text, bold: it.bold, italic: it.italic, size: sz, fam: it.fam });
        const n = it.s.trim().length;
        chars += n;
        sizeVotes[sz] = (sizeVotes[sz] || 0) + n;
      }
      if (!chars) continue;
      let dom = row.size, best = -1;
      for (const k in sizeVotes) if (sizeVotes[k] > best) { best = sizeVotes[k]; dom = parseFloat(k); }
      const first = row.items[0], lastIt = row.items[row.items.length - 1];
      lines.push({
        y: row.y, size: dom, x0: first.x, x1: lastIt.x + lastIt.w, runs, chars,
        text: runs.map((r) => r.text).join(""),
        allBold: runs.every((r) => r.bold || !r.text.trim()),
      });
    }
    return lines;
  }

  const LIST_START = /^\s*([•●▪◦■□○\-–—*·]|\(?\d{1,3}[.)]|\(?[a-zA-Z][.)])\s/;

  /** Turn lines of one page into paragraphs. */
  function buildParagraphs(lines, ctx) {
    const { pageW, left, right, bodySize } = ctx;
    const textW = Math.max(50, right - left);
    // typical distance between two lines of the same paragraph
    const pitches = [];
    for (let i = 1; i < lines.length; i++) {
      const p = lines[i - 1], l = lines[i];
      const d = p.y - l.y;
      if (Math.abs(p.size - l.size) < 0.6 && Math.abs(p.x0 - l.x0) < l.size * 2 && d > l.size * 0.9 && d < l.size * 2.6) pitches.push(d);
    }
    const basePitch = median(pitches) || bodySize * 1.2;

    const paras = [];
    let cur = null;
    for (let i = 0; i < lines.length; i++) {
      const L = lines[i], P = i ? lines[i - 1] : null;
      let brk = !cur;
      if (!brk) {
        const d = P.y - L.y;
        const ownPitch = cur.lines.length > 1 ? cur.pitch : basePitch * (P.size / bodySize > 1.1 ? P.size / bodySize : 1);
        if (d > Math.max(ownPitch, L.size * 1.15) * 1.35) brk = true;
        else if (Math.abs(L.size - P.size) > 0.8) brk = true;
        else if (LIST_START.test(L.text) && !(cur.lines.length === 1 && /^\s*[•●▪◦■□○\-–—*·]\s*$/.test(P.text))) brk = true;
        else if (P.allBold && !L.allBold) brk = true;
        else if (P.x1 < right - 0.25 * textW) brk = true; // previous line ended early
        else if (L.x0 - P.x0 > L.size * 0.8 && !LIST_START.test(P.text) && cur.lines.length === 1 && false) brk = true;
      }
      if (brk) {
        cur = { lines: [], pitch: 0 };
        paras.push(cur);
      } else cur.pitch = P.y - L.y;
      cur.lines.push(L);
    }
    return paras;
  }

  function paraToXml(para, ctx, nextGap, firstOnPage, pageIndex) {
    const { left, pageW, bodySize } = ctx;
    const lines = para.lines;
    const L0 = lines[0], Ln = lines[lines.length - 1];
    const size = L0.size;
    const x0 = L0.x0;
    const contX = lines.length > 1 ? lines[1].x0 : x0;
    const centerOfPage = pageW / 2;
    let jc = "";
    const lineCenter = (L0.x0 + L0.x1) / 2;
    if (lines.every((l) => Math.abs((l.x0 + l.x1) / 2 - centerOfPage) < pageW * 0.035) && x0 - left > pageW * 0.06) jc = "center";
    else if (lines.length === 1 && Math.abs(L0.x1 - ctx.right) < pageW * 0.015 && x0 - left > (ctx.right - left) * 0.35) jc = "right";

    let ind = "";
    if (!jc) {
      const eps = size * 0.8;
      if (contX > x0 + eps) ind = `<w:ind w:left="${Math.round((contX - left) * 20)}" w:hanging="${Math.round((contX - x0) * 20)}"/>`;
      else if (x0 > contX + eps) ind = `<w:ind w:left="${Math.round(Math.max(0, contX - left) * 20)}" w:firstLine="${Math.round((x0 - contX) * 20)}"/>`;
      else if (x0 - left > eps) ind = `<w:ind w:left="${Math.round((x0 - left) * 20)}"/>`;
    }

    const ratio = size / bodySize;
    let style = "";
    if (lines.length <= 3 && ratio >= 1.15) style = ratio >= 1.6 ? "Heading1" : ratio >= 1.3 ? "Heading2" : "Heading3";

    let spacing;
    const ownPitch = lines.length > 1 ? para.pitch : 0;
    let lineH = size * 1.15;
    if (ownPitch && ownPitch > size * 0.95 && ownPitch < size * 3) {
      lineH = ownPitch;
      spacing = { line: Math.round(ownPitch * 20), rule: "exact" };
    } else spacing = { line: 240, rule: "auto" };
    const after = nextGap == null ? 0 : Math.max(0, nextGap - lineH);
    const pPr =
      (style ? `<w:pStyle w:val="${style}"/>` : "") +
      (firstOnPage && pageIndex > 0 ? "<w:pageBreakBefore/>" : "") +
      `<w:spacing w:before="0" w:after="${Math.round(after * 20)}" w:line="${spacing.line}" w:lineRule="${spacing.rule}"/>` +
      ind +
      (jc ? `<w:jc w:val="${jc}"/>` : "");

    // runs, with a space between lines
    let xml = "";
    lines.forEach((l, li) => {
      l.runs.forEach((r, ri) => {
        let text = r.text;
        if (ri === 0 && li > 0) {
          const prev = lines[li - 1].text;
          if (!/-$/.test(prev) || /^\s/.test(text)) text = (/\s$/.test(prev) ? "" : " ") + text.replace(/^\s+/, "");
        }
        if (ri === 0) text = text.replace(/^\t+/, ""); // leading gap = indent, already handled
        const parts = text.split("\t");
        parts.forEach((part, pi) => {
          if (pi > 0) xml += `<w:r><w:tab/></w:r>`;
          if (!part) return;
          const hp = Math.round(r.size * 2);
          xml +=
            `<w:r><w:rPr><w:rFonts w:ascii="${r.fam}" w:hAnsi="${r.fam}" w:cs="${r.fam}"/>` +
            (r.bold ? "<w:b/><w:bCs/>" : "") + (r.italic ? "<w:i/><w:iCs/>" : "") +
            `<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/></w:rPr><w:t xml:space="preserve">${esc(clean(part))}</w:t></w:r>`;
        });
      });
    });
    return `<w:p><w:pPr>${pPr}</w:pPr>${xml}</w:p>`;
  }

  const STYLES_XML =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
    [1, 2, 3].map((n) =>
      `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="${n - 1}"/></w:pPr></w:style>`
    ).join("") +
    `</w:styles>`;

  const CONTENT_TYPES = (hasImages) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
    (hasImages ? `<Default Extension="jpg" ContentType="image/jpeg"/>` : "") +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`;

  const ROOT_RELS =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

  function docRels(images) {
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      images.map((n, i) => `<Relationship Id="rIdImg${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${n}"/>`).join("") +
      `</Relationships>`
    );
  }

  async function zipDocx(JSZip, documentXml, images /* [{name, bytes}] */) {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", CONTENT_TYPES(images.length > 0));
    zip.file("_rels/.rels", ROOT_RELS);
    zip.file("word/document.xml", documentXml);
    zip.file("word/styles.xml", STYLES_XML);
    zip.file("word/_rels/document.xml.rels", docRels(images.map((i) => i.name)));
    images.forEach((i) => zip.file("word/media/" + i.name, i.bytes));
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }

  /**
   * pdfjs: the pdf.js module, data: ArrayBuffer/Uint8Array, mode: "text" | "image"
   * onProgress(done, total)
   */
  async function pdfToDocx(pdfjs, JSZip, data, mode, onProgress) {
    const task = pdfjs.getDocument({ data: new Uint8Array(data), isEvalSupported: false });
    const pdf = await task.promise;
    const n = pdf.numPages;
    try {
      if (mode === "image") return await pdfToDocxImages(pdf, JSZip, onProgress);
      const pages = [];
      let first = null;
      for (let p = 1; p <= n; p++) {
        const page = await pdf.getPage(p);
        const vp = page.getViewport({ scale: 1 });
        if (!first) first = { w: vp.width, h: vp.height };
        pages.push({ w: vp.width, h: vp.height, lines: await readPageLines(page) });
        page.cleanup();
        onProgress && onProgress(p, n);
      }
      const all = pages.flatMap((p) => p.lines);
      const totalChars = all.reduce((s, l) => s + l.chars, 0);
      if (totalChars < 3) {
        const err = new Error("This PDF has no selectable text (it looks scanned). Use the “Exact look” option instead.");
        err.code = "NO_TEXT";
        throw err;
      }
      const votes = {};
      all.forEach((l) => (votes[l.size] = (votes[l.size] || 0) + l.chars));
      let bodySize = 11, best = -1;
      for (const k in votes) if (votes[k] > best) { best = votes[k]; bodySize = parseFloat(k); }

      const left = clamp(pct(all.map((l) => l.x0), 0.05), 18, first.w * 0.4);
      const right = clamp(pct(all.map((l) => l.x1), 0.95), first.w * 0.5, first.w - 18);
      const topM = clamp(Math.min(...pages.filter((p) => p.lines.length).map((p) => p.h - p.lines[0].y - p.lines[0].size * 0.85)), 24, 100);
      const botM = clamp(Math.min(...pages.filter((p) => p.lines.length).map((p) => p.lines[p.lines.length - 1].y - 0.25 * p.lines[p.lines.length - 1].size)), 24, 100);

      let body = "";
      pages.forEach((pg, pi) => {
        if (!pg.lines.length) {
          body += `<w:p><w:pPr>${pi > 0 ? "<w:pageBreakBefore/>" : ""}</w:pPr></w:p>`;
          return;
        }
        const ctx = { pageW: pg.w, left, right, bodySize };
        const paras = buildParagraphs(pg.lines, ctx);
        paras.forEach((para, i) => {
          const nxt = paras[i + 1];
          const gap = nxt ? para.lines[para.lines.length - 1].y - nxt.lines[0].y : null;
          body += paraToXml(para, ctx, gap, i === 0, pi);
        });
      });
      const sect =
        `<w:sectPr><w:pgSz w:w="${Math.round(first.w * 20)}" w:h="${Math.round(first.h * 20)}"/>` +
        `<w:pgMar w:top="${Math.round(topM * 20)}" w:right="${Math.round((first.w - right) * 20)}" w:bottom="${Math.round(botM * 20)}" w:left="${Math.round(left * 20)}" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>`;
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${sect}</w:body></w:document>`;
      return await zipDocx(JSZip, xml, []);
    } finally {
      try { await pdf.destroy(); } catch (_) {}
    }
  }

  async function pdfToDocxImages(pdf, JSZip, onProgress) {
    const n = pdf.numPages;
    const images = [];
    let body = "", first = null;
    for (let p = 1; p <= n; p++) {
      const page = await pdf.getPage(p);
      const vp1 = page.getViewport({ scale: 1 });
      if (!first) first = { w: vp1.width, h: vp1.height };
      const scale = Math.min(3, 1800 / Math.max(vp1.width, vp1.height));
      const vp = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(vp.width);
      canvas.height = Math.ceil(vp.height);
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, canvas, viewport: vp }).promise;
      const blob = await new Promise((r) => canvas.toBlob(r, "image/jpeg", 0.85));
      canvas.width = canvas.height = 0;
      const name = `page${p}.jpg`;
      images.push({ name, bytes: new Uint8Array(await blob.arrayBuffer()) });
      const k = 0.99; // keep a hair inside the page so Word doesn't add blank pages
      const cx = Math.round(vp1.width * 12700 * k), cy = Math.round(vp1.height * 12700 * k);
      body +=
        `<w:p><w:pPr>${p > 1 ? "<w:pageBreakBefore/>" : ""}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:drawing>` +
        `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${p}" name="Page ${p}"/>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${p}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="rIdImg${p}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
      page.cleanup();
      onProgress && onProgress(p, n);
    }
    const sect =
      `<w:sectPr><w:pgSz w:w="${Math.round(first.w * 20)}" w:h="${Math.round(first.h * 20)}"/>` +
      `<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${sect}</w:body></w:document>`;
    return zipDocx(JSZip, xml, images);
  }

  /* =====================================================================
     WORD (.docx)  ->  PDF
     ===================================================================== */

  const WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const kids = (el, name) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
  const kid = (el, name) => {
    if (!el) return null;
    for (const c of el.children) if (c.localName === name) return c;
    return null;
  };
  const wattr = (el, name) => (el ? el.getAttributeNS(WNS, name) ?? el.getAttribute("w:" + name) : null);
  const onoff = (el) => {
    if (!el) return undefined;
    const v = wattr(el, "val");
    return !(v === "0" || v === "false" || v === "off");
  };
  const intAttr = (el, name) => {
    const v = wattr(el, name);
    return v === null || v === "" || isNaN(parseInt(v, 10)) ? undefined : parseInt(v, 10);
  };
  const isEmptyObj = (o) => !o || !Object.keys(o).length;

  function readRPr(el) {
    const r = {};
    if (!el) return r;
    let v;
    if ((v = onoff(kid(el, "b"))) !== undefined) r.b = v;
    if ((v = onoff(kid(el, "i"))) !== undefined) r.i = v;
    if ((v = onoff(kid(el, "strike"))) !== undefined) r.strike = v;
    if ((v = onoff(kid(el, "caps"))) !== undefined) r.caps = v;
    if ((v = onoff(kid(el, "smallCaps"))) !== undefined) r.caps = v;
    const u = kid(el, "u");
    if (u) r.u = wattr(u, "val") !== "none";
    const sz = intAttr(kid(el, "sz"), "val");
    if (sz) r.sz = sz;
    const col = wattr(kid(el, "color"), "val");
    if (col && col !== "auto" && /^[0-9a-fA-F]{6}$/.test(col)) r.color = col;
    const rf = kid(el, "rFonts");
    if (rf) {
      const f = wattr(rf, "ascii") || wattr(rf, "hAnsi") || wattr(rf, "cs");
      if (f) r.font = f;
    }
    const va = wattr(kid(el, "vertAlign"), "val");
    if (va === "superscript" || va === "subscript") r.vert = va;
    const rs = wattr(kid(el, "rStyle"), "val");
    if (rs) r.rStyle = rs;
    return r;
  }
  function readPPr(el) {
    const p = {};
    if (!el) return p;
    const jc = wattr(kid(el, "jc"), "val");
    if (jc) p.jc = jc === "start" ? "left" : jc === "end" ? "right" : jc;
    const ind = kid(el, "ind");
    if (ind) {
      p.ind = {};
      const l = intAttr(ind, "left") ?? intAttr(ind, "start");
      const r = intAttr(ind, "right") ?? intAttr(ind, "end");
      if (l !== undefined) p.ind.left = l;
      if (r !== undefined) p.ind.right = r;
      const fl = intAttr(ind, "firstLine"), hg = intAttr(ind, "hanging");
      if (hg !== undefined) { p.ind.hanging = hg; p.ind.firstLine = undefined; }
      else if (fl !== undefined) { p.ind.firstLine = fl; p.ind.hanging = undefined; }
    }
    const sp = kid(el, "spacing");
    if (sp) {
      p.spacing = {};
      for (const k of ["before", "after", "line"]) { const v = intAttr(sp, k); if (v !== undefined) p.spacing[k] = v; }
      const rule = wattr(sp, "lineRule");
      if (rule) p.spacing.rule = rule;
    }
    const np = kid(el, "numPr");
    if (np) {
      const id = intAttr(kid(np, "numId"), "val"), lvl = intAttr(kid(np, "ilvl"), "val");
      if (id !== undefined) p.numId = id;
      if (lvl !== undefined) p.ilvl = lvl;
    }
    const ps = wattr(kid(el, "pStyle"), "val");
    if (ps) p.pStyle = ps;
    if (kid(el, "pageBreakBefore")) p.pageBreakBefore = onoff(kid(el, "pageBreakBefore"));
    const shd = kid(el, "shd");
    if (shd) { const f = wattr(shd, "fill"); if (f && f !== "auto") p.shd = f; }
    const mark = kid(el, "rPr");
    if (mark) p.markR = readRPr(mark);
    return p;
  }
  const mergeP = (...ps) => {
    const out = { ind: {}, spacing: {} };
    for (const p of ps) {
      if (!p) continue;
      for (const k in p) {
        if (k === "ind" || k === "spacing") Object.assign(out[k], p[k]);
        else if (p[k] !== undefined) out[k] = p[k];
      }
    }
    return out;
  };
  const mergeR = (...rs) => Object.assign({}, ...rs.filter(Boolean));

  function parseStyles(doc) {
    const S = { byId: {}, defP: {}, defR: {}, defaultPara: null, defaultTable: null, memo: {} };
    if (!doc) return S;
    const root = doc.documentElement;
    const dd = kid(root, "docDefaults");
    if (dd) {
      S.defR = readRPr(kid(kid(dd, "rPrDefault"), "rPr"));
      S.defP = readPPr(kid(kid(dd, "pPrDefault"), "pPr"));
    }
    for (const el of kids(root, "style")) {
      const id = wattr(el, "styleId");
      const type = wattr(el, "type");
      const st = {
        id, type,
        basedOn: wattr(kid(el, "basedOn"), "val"),
        p: readPPr(kid(el, "pPr")),
        r: readRPr(kid(el, "rPr")),
        borders: false,
      };
      const tb = kid(kid(el, "tblPr"), "tblBorders");
      if (tb) st.borders = Array.from(tb.children).some((b) => !["none", "nil"].includes(wattr(b, "val")));
      S.byId[id] = st;
      if (wattr(el, "default") === "1") {
        if (type === "paragraph") S.defaultPara = id;
        if (type === "table") S.defaultTable = id;
      }
    }
    return S;
  }
  function resolveStyle(S, id) {
    if (!id || !S.byId[id]) return { p: {}, r: {}, borders: false };
    if (S.memo[id]) return S.memo[id];
    S.memo[id] = { p: {}, r: {}, borders: false }; // cycle guard
    const st = S.byId[id];
    const base = st.basedOn ? resolveStyle(S, st.basedOn) : { p: {}, r: {}, borders: false };
    return (S.memo[id] = { p: mergeP(base.p, st.p), r: mergeR(base.r, st.r), borders: st.borders || base.borders });
  }

  function parseNumbering(doc) {
    const N = { abstracts: {}, nums: {} };
    if (!doc) return N;
    const root = doc.documentElement;
    for (const an of kids(root, "abstractNum")) {
      const id = intAttr(an, "abstractNumId");
      const levels = {};
      for (const lv of kids(an, "lvl")) {
        const il = intAttr(lv, "ilvl");
        levels[il] = {
          fmt: wattr(kid(lv, "numFmt"), "val") || "decimal",
          text: wattr(kid(lv, "lvlText"), "val") ?? "",
          start: intAttr(kid(lv, "start"), "val") ?? 1,
          p: readPPr(kid(lv, "pPr")),
          r: readRPr(kid(lv, "rPr")),
        };
      }
      N.abstracts[id] = levels;
    }
    for (const nm of kids(root, "num")) {
      const id = intAttr(nm, "numId");
      const overrides = {};
      for (const ov of kids(nm, "lvlOverride")) {
        const il = intAttr(ov, "ilvl");
        const so = intAttr(kid(ov, "startOverride"), "val");
        if (so !== undefined) overrides[il] = so;
      }
      N.nums[id] = { abs: intAttr(kid(nm, "abstractNumId"), "val"), overrides };
    }
    return N;
  }
  const roman = (n, lower) => {
    const m = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
    let s = "";
    for (const [v, r] of m) while (n >= v) { s += r; n -= v; }
    return lower ? s.toLowerCase() : s;
  };
  function fmtNum(n, fmt) {
    switch (fmt) {
      case "lowerLetter": return String.fromCharCode(96 + ((n - 1) % 26) + 1);
      case "upperLetter": return String.fromCharCode(64 + ((n - 1) % 26) + 1);
      case "lowerRoman": return roman(n, true);
      case "upperRoman": return roman(n, false);
      case "none": return "";
      default: return String(n);
    }
  }

  /** Parse document.xml into a simple block model. */
  function parseDocument(docXml, S, N, rels) {
    const body = kid(docXml.documentElement, "body");
    const counters = {}; // numId -> array
    const imageRefs = new Set();

    const resolveNumber = (numId, ilvl) => {
      const num = N.nums[numId];
      if (!num) return null;
      const levels = N.abstracts[num.abs];
      if (!levels || !levels[ilvl]) return null;
      const c = counters[numId] || (counters[numId] = []);
      const lv = levels[ilvl];
      const start = num.overrides[ilvl] ?? lv.start;
      c[ilvl] = c[ilvl] === undefined ? start : c[ilvl] + 1;
      for (let k = ilvl + 1; k < 10; k++) c[k] = undefined;
      let text;
      if (lv.fmt === "bullet") {
        text = lv.text;
        if (!text || /[-]/.test(text)) text = "•";
      } else {
        text = lv.text.replace(/%(\d)/g, (_, d) => {
          const li = parseInt(d, 10) - 1;
          const l2 = levels[li];
          const val = c[li] === undefined ? (l2 ? l2.start : 1) : c[li];
          return fmtNum(val, l2 ? l2.fmt : "decimal");
        });
      }
      return { text, lv };
    };

    function parseRunContent(r, props, out, link) {
      for (const c of r.children) {
        switch (c.localName) {
          case "t": out.push({ text: c.textContent, props, link }); break;
          case "tab": out.push({ text: "\t", props, link }); break;
          case "br": case "cr": {
            const t = wattr(c, "type");
            out.push(t === "page" ? { pageBreak: true } : { br: true });
            break;
          }
          case "noBreakHyphen": out.push({ text: "-", props, link }); break;
          case "drawing": {
            const ext = c.getElementsByTagNameNS("*", "extent")[0];
            const blip = c.getElementsByTagNameNS("*", "blip")[0];
            if (blip) {
              const rid = blip.getAttributeNS(RNS, "embed");
              const cx = ext ? parseInt(ext.getAttribute("cx"), 10) : 0, cy = ext ? parseInt(ext.getAttribute("cy"), 10) : 0;
              if (rid && rels[rid]) { imageRefs.add(rid); out.push({ img: rid, w: cx / 12700, h: cy / 12700 }); }
            }
            break;
          }
          case "pict": {
            const im = c.getElementsByTagNameNS("*", "imagedata")[0];
            if (im) {
              const rid = im.getAttributeNS(RNS, "id");
              if (rid && rels[rid]) { imageRefs.add(rid); out.push({ img: rid, w: 0, h: 0 }); }
            }
            break;
          }
          case "AlternateContent": {
            const ch = kid(c, "Choice") || kid(c, "Fallback");
            if (ch) parseRunContent(ch, props, out, link);
            break;
          }
          default: break;
        }
      }
    }

    function walkInline(parent, pctx, out, link) {
      for (const c of parent.children) {
        switch (c.localName) {
          case "r": {
            const rp = readRPr(kid(c, "rPr"));
            const cs = rp.rStyle ? resolveStyle(S, rp.rStyle).r : null;
            const props = mergeR(S.defR, pctx.styleR, cs, link ? { u: true, color: "0563C1" } : null, rp);
            parseRunContent(c, props, out, link);
            break;
          }
          case "hyperlink": {
            const rid = c.getAttributeNS(RNS, "id");
            walkInline(c, pctx, out, (rid && rels[rid] && rels[rid].url) || wattr(c, "anchor") || true);
            break;
          }
          case "fldSimple": case "ins": case "smartTag": case "sdt": case "moveTo": case "customXml":
            walkInline(c.localName === "sdt" ? kid(c, "sdtContent") || c : c, pctx, out, link);
            break;
          case "AlternateContent": {
            const ch = kid(c, "Choice") || kid(c, "Fallback");
            if (ch) walkInline(ch, pctx, out, link);
            break;
          }
          default: break; // del, moveFrom, bookmarks, proofErr, pPr ...
        }
      }
    }

    function parsePara(p, tblStyle) {
      const direct = readPPr(kid(p, "pPr"));
      const styleId = direct.pStyle || S.defaultPara;
      const st = resolveStyle(S, styleId);
      let num = null;
      const numId = direct.numId !== undefined ? direct.numId : st.p.numId;
      const ilvl = direct.ilvl !== undefined ? direct.ilvl : st.p.ilvl ?? 0;
      let numP = null;
      if (numId) {
        num = resolveNumber(numId, ilvl);
        if (num) numP = num.lv.p;
      }
      const props = mergeP(S.defP, tblStyle && tblStyle.p, st.p, numP, direct);
      const pctx = { styleR: mergeR(tblStyle && tblStyle.r, st.r) };
      const runs = [];
      walkInline(p, pctx, runs, null);
      const markR = mergeR(S.defR, pctx.styleR, direct.markR);
      let marker = null;
      if (num && num.text) marker = { text: num.text, props: mergeR(markR, num.lv.r) };
      return { type: "p", runs, props, marker, markR };
    }

    function parseTable(t, outerStyle) {
      const tblPr = kid(t, "tblPr");
      const styleId = wattr(kid(tblPr, "tblStyle"), "val") || S.defaultTable;
      const st = resolveStyle(S, styleId);
      let borders = st.borders;
      const tb = kid(tblPr, "tblBorders");
      if (tb) borders = Array.from(tb.children).some((b) => !["none", "nil"].includes(wattr(b, "val")));
      const grid = kids(kid(t, "tblGrid"), "gridCol").map((g) => (intAttr(g, "w") || 0) / 20);
      const tw = kid(tblPr, "tblW");
      const mar = kid(tblPr, "tblCellMar");
      const padL = mar && kid(mar, "left") ? (intAttr(kid(mar, "left"), "w") || 0) / 20 : 5.4;
      const padR = mar && kid(mar, "right") ? (intAttr(kid(mar, "right"), "w") || 0) / 20 : 5.4;
      const padT = mar && kid(mar, "top") ? (intAttr(kid(mar, "top"), "w") || 0) / 20 : 0;
      const padB = mar && kid(mar, "bottom") ? (intAttr(kid(mar, "bottom"), "w") || 0) / 20 : 0;
      const rows = [];
      for (const tr of kids(t, "tr")) {
        const trPr = kid(tr, "trPr");
        const minH = (intAttr(kid(trPr, "trHeight"), "val") || 0) / 20;
        const cells = [];
        for (const tc of kids(tr, "tc")) {
          const tcPr = kid(tc, "tcPr");
          const span = intAttr(kid(tcPr, "gridSpan"), "val") || 1;
          const vm = kid(tcPr, "vMerge");
          const shd = kid(tcPr, "shd");
          const fill = shd ? wattr(shd, "fill") : null;
          const tcb = kid(tcPr, "tcBorders");
          let cellBorders = null;
          if (tcb) {
            cellBorders = {};
            for (const b of Array.from(tcb.children)) cellBorders[b.localName] = !["none", "nil"].includes(wattr(b, "val"));
          }
          const w = (intAttr(kid(tcPr, "tcW"), "w") || 0) / 20;
          cells.push({
            span, w, fill: fill && fill !== "auto" && /^[0-9a-fA-F]{6}$/.test(fill) ? fill : null,
            vMergeCont: !!vm && wattr(vm, "val") !== "restart",
            borders: cellBorders,
            blocks: parseBlocks(tc, st),
          });
        }
        rows.push({ cells, minH });
      }
      let g = grid;
      if (!g.length && rows.length) {
        g = [];
        rows[0].cells.forEach((c) => { for (let i = 0; i < c.span; i++) g.push((c.w || 72) / c.span); });
      }
      return { type: "tbl", grid: g, rows, borders, padL, padR, padT, padB, widthPct: tw && wattr(tw, "type") === "pct" ? (intAttr(tw, "w") || 5000) / 5000 : null };
    }

    function parseBlocks(parent, tblStyle) {
      const out = [];
      for (const c of parent.children) {
        if (c.localName === "p") out.push(parsePara(c, tblStyle));
        else if (c.localName === "tbl") out.push(parseTable(c, tblStyle));
        else if (c.localName === "sdt") out.push(...parseBlocks(kid(c, "sdtContent") || c, tblStyle));
        else if (c.localName === "AlternateContent") { const ch = kid(c, "Choice"); if (ch) out.push(...parseBlocks(ch, tblStyle)); }
      }
      return out;
    }

    const blocks = parseBlocks(body, null);
    // page setup from the last section properties
    let page = { w: 595.28, h: 841.89, t: 72, r: 72, b: 72, l: 72 };
    const sp = kid(body, "sectPr");
    if (sp) {
      const sz = kid(sp, "pgSz"), mg = kid(sp, "pgMar");
      if (sz) {
        const w = intAttr(sz, "w"), h = intAttr(sz, "h");
        if (w && h) { page.w = w / 20; page.h = h / 20; }
      }
      if (mg) {
        const v = (k, d) => (intAttr(mg, k) !== undefined ? intAttr(mg, k) / 20 : d);
        page.t = Math.max(0, v("top", 72)); page.b = Math.max(0, v("bottom", 72));
        page.l = v("left", 72); page.r = v("right", 72);
      }
    }
    return { blocks, page, imageRefs };
  }

  /* ---------- rendering with pdf-lib ---------- */
  const CHAR_MAP = {
    " ": " ", " ": " ", " ": " ", " ": " ", " ": " ", " ": " ", " ": " ",
    "‐": "-", "‑": "-", "‒": "-", "−": "-", "­": "", "​": "", "‌": "", "‍": "", "﻿": "",
    "→": "->", "←": "<-", "≤": "<=", "≥": ">=", "≠": "!=", "✓": "v", "✔": "v", "✗": "x",
    "○": "o", "□": "-", "●": "•", "▪": "•", "■": "•", "◦": "o", "‣": "•", "⁃": "-", "·": "·", "∙": "•",
    "‘": "‘", "’": "’", "“": "“", "”": "”", "…": "…",
  };

  const FONT_NAMES = {
    sans: ["Helvetica", "HelveticaBold", "HelveticaOblique", "HelveticaBoldOblique"],
    serif: ["TimesRoman", "TimesRomanBold", "TimesRomanItalic", "TimesRomanBoldItalic"],
    mono: ["Courier", "CourierBold", "CourierOblique", "CourierBoldOblique"],
  };
  function familyKey(name) {
    const f = (name || "").toLowerCase();
    if (!f) return "serif";
    if (/courier|consolas|mono|menlo|lucida console/.test(f)) return "mono";
    if (/times|serif|georgia|cambria|garamond|palatino|minion|book|century|constantia|didot/.test(f) && !/sans/.test(f)) return "serif";
    return "sans";
  }
  const hex = (h) => (h ? [parseInt(h.slice(0, 2), 16) / 255, parseInt(h.slice(2, 4), 16) / 255, parseInt(h.slice(4, 6), 16) / 255] : [0, 0, 0]);

  async function docxToPdf(JSZip, PDFLib, data, opts) {
    opts = opts || {};
    const Parser = opts.DOMParser || DOMParser;
    let zip;
    try { zip = await JSZip.loadAsync(data); } catch (_) { throw new Error("This doesn't look like a valid .docx file."); }
    const readXml = async (path) => {
      const f = zip.file(path);
      return f ? new Parser().parseFromString(await f.async("string"), "application/xml") : null;
    };
    const docXml = await readXml("word/document.xml");
    if (!docXml || !docXml.documentElement || docXml.getElementsByTagName("parsererror").length) {
      throw new Error("This doesn't look like a valid .docx file.");
    }
    const S = parseStyles(await readXml("word/styles.xml"));
    const N = parseNumbering(await readXml("word/numbering.xml"));
    const relDoc = await readXml("word/_rels/document.xml.rels");
    const rels = {};
    if (relDoc) {
      for (const r of Array.from(relDoc.documentElement.children)) {
        const type = r.getAttribute("Type") || "";
        rels[r.getAttribute("Id")] = {
          target: r.getAttribute("Target"),
          mode: r.getAttribute("TargetMode"),
          isImage: /\/image$/.test(type),
          url: r.getAttribute("TargetMode") === "External" ? r.getAttribute("Target") : null,
        };
      }
    }
    const model = parseDocument(docXml, S, N, rels);

    const pdf = await PDFLib.PDFDocument.create();
    pdf.setProducer("PDF Toolkit");
    pdf.setCreator("PDF Toolkit");

    // fonts
    const fontCache = {};
    const charsets = {};
    const getFont = async (fam, bold, italic) => {
      const key = fam + (bold ? "B" : "") + (italic ? "I" : "");
      if (!fontCache[key]) {
        const idx = (bold ? 1 : 0) + (italic ? (bold ? 2 : 2) : 0);
        const nm = FONT_NAMES[fam][bold && italic ? 3 : bold ? 1 : italic ? 2 : 0];
        const f = await pdf.embedFont(PDFLib.StandardFonts[nm]);
        f.__set = new Set(f.getCharacterSet());
        f.__w = new Map();
        fontCache[key] = f;
      }
      return fontCache[key];
    };
    for (const fam of ["sans", "serif", "mono"]) for (const b of [false, true]) for (const i of [false, true]) await getFont(fam, b, i);

    const encode = (font, text) => {
      let out = "";
      for (const ch of text) {
        let c = CHAR_MAP[ch] !== undefined ? CHAR_MAP[ch] : ch;
        for (const cc of c) out += font.__set.has(cc.codePointAt(0)) ? cc : cc.charCodeAt(0) < 32 ? "" : "?";
      }
      return out;
    };
    const widthOf = (font, size, text) => {
      if (!text) return 0;
      const k = text + "|" + size;
      let w = font.__w.get(k);
      if (w === undefined) {
        w = 0; // measure glyph by glyph: pdf-lib's whole-string width applies kerning, but drawText does not
        for (const ch of text) {
          const ck = "c" + ch + "|" + size;
          let cw = font.__w.get(ck);
          if (cw === undefined) { cw = font.widthOfTextAtSize(ch, size); font.__w.set(ck, cw); }
          w += cw;
        }
        font.__w.set(k, w);
      }
      return w;
    };

    // images
    const images = {};
    for (const rid of model.imageRefs) {
      const target = rels[rid].target;
      const path = target.startsWith("/") ? target.slice(1) : "word/" + target.replace(/^\.\//, "");
      const norm = path.split("/").reduce((acc, seg) => (seg === ".." ? (acc.pop(), acc) : (acc.push(seg), acc)), []).join("/");
      const f = zip.file(norm);
      if (!f) continue;
      try {
        const bytes = await f.async("uint8array");
        const ext = (norm.split(".").pop() || "").toLowerCase();
        let img;
        if (ext === "png") img = await pdf.embedPng(bytes);
        else if (ext === "jpg" || ext === "jpeg") img = await pdf.embedJpg(bytes);
        else if (opts.convertImage) img = await pdf.embedPng(await opts.convertImage(bytes, ext));
        if (img) images[rid] = img;
      } catch (_) { /* skip pictures that can't be read */ }
    }

    const PG = model.page;
    const contentW = Math.max(72, PG.w - PG.l - PG.r);
    const bottomLimit = PG.h - PG.b;

    /* ---- paragraph layout: returns {pieces} ---- */
    function layoutParagraph(P, width) {
      const props = P.props;
      const ind = props.ind || {};
      let left = (ind.left || 0) / 20, right = (ind.right || 0) / 20;
      const hanging = (ind.hanging || 0) / 20, firstLine = (ind.firstLine || 0) / 20;
      const sp = props.spacing || {};
      const before = (sp.before || 0) / 20, after = (sp.after || 0) / 20;
      const lineSpec = sp.line || 240, rule = sp.rule || "auto";
      const jc = props.jc || "left";
      const avail = Math.max(20, width - left - right);

      // tokens
      const toks = [];
      for (const r of P.runs) {
        if (r.pageBreak) { toks.push({ k: "pb" }); continue; }
        if (r.br) { toks.push({ k: "br" }); continue; }
        if (r.img) {
          const im = images[r.img];
          if (im) {
            let w = r.w || im.width * 0.75, h = r.h || im.height * 0.75;
            const maxW = avail;
            if (w > maxW) { h = (h * maxW) / w; w = maxW; }
            toks.push({ k: "img", w, h, im });
          }
          continue;
        }
        const pr = r.props || {};
        const size0 = (pr.sz || 20) / 2;
        const size = pr.vert ? size0 * 0.65 : size0;
        const rise = pr.vert === "superscript" ? size0 * 0.35 : pr.vert === "subscript" ? -size0 * 0.12 : 0;
        const style = {
          fam: familyKey(pr.font), bold: !!pr.b, italic: !!pr.i, size, rise,
          color: pr.color || null, u: !!pr.u, strike: !!pr.strike,
        };
        let text = pr.caps ? r.text.toUpperCase() : r.text;
        const parts = text.split(/(\t|[     ]+)/);
        for (const part of parts) {
          if (!part) continue;
          if (part === "\t") toks.push({ k: "tab", style });
          else if (/^[     ]+$/.test(part)) toks.push({ k: "sp", text: " ".repeat(part.length), style });
          else toks.push({ k: "w", text: part, style });
        }
      }

      const fonts = new Map();
      const fontFor = (st) => {
        const key = st.fam + st.bold + st.italic;
        let f = fonts.get(key);
        if (!f) { f = fontCache[st.fam + (st.bold ? "B" : "") + (st.italic ? "I" : "")]; fonts.set(key, f); }
        return f;
      };
      const markStyle = (() => {
        const m = P.markR || {};
        return { fam: familyKey(m.font), bold: !!m.b, italic: !!m.i, size: (m.sz || 20) / 2, rise: 0, color: m.color || null, u: false, strike: false };
      })();

      // line breaking
      const lines = [];
      let cur = null;
      let markerW = 0;
      if (P.marker) {
        const mp = P.marker.props;
        const mf = fontCache[familyKey(mp.font) + (mp.b ? "B" : "") + (mp.i ? "I" : "")];
        markerW = widthOf(mf, (mp.sz || 20) / 2, encode(mf, P.marker.text));
      }
      const startLine = () => {
        const isFirst = lines.length === 0;
        let offset = isFirst ? left + (hanging ? -hanging : firstLine) : left;
        if (isFirst && P.marker) {
          if (hanging) offset = left;
          else { const mEnd = left + firstLine + markerW; offset = Math.ceil((mEnd + 1) / 36) * 36; } // number, then a tab
        }
        const off = Math.max(0, offset);
        cur = { items: [], w: 0, off, avail: width - off - right, forced: false, pageBreakAfter: false, trailing: 0 };
        return cur;
      };
      const endLine = (forced, pageBreakAfter) => {
        // drop trailing spaces
        while (cur.items.length && cur.items[cur.items.length - 1].k === "sp") cur.w -= cur.items.pop().w;
        cur.forced = !!forced;
        cur.pageBreakAfter = !!pageBreakAfter;
        lines.push(cur);
        cur = null;
      };
      startLine();
      let pendingSp = 0;
      for (const t of toks) {
        if (!cur) startLine();
        if (t.k === "br") { endLine(true); continue; }
        if (t.k === "pb") { endLine(true, true); continue; }
        if (t.k === "tab") {
          const pos = cur.off + cur.w;
          const stop = Math.ceil((pos + 0.5) / 36) * 36;
          const w = Math.min(stop - pos, Math.max(0, cur.avail - cur.w));
          cur.items.push({ k: "sp", w, style: t.style, text: "" });
          cur.w += w;
          continue;
        }
        if (t.k === "sp") {
          if (!cur.items.length) continue; // no spaces at line start
          const f = fontFor(t.style);
          const w = widthOf(f, t.style.size, t.text);
          cur.items.push({ k: "sp", w, style: t.style, text: t.text, font: f });
          cur.w += w;
          continue;
        }
        if (t.k === "img") {
          if (cur.items.length && cur.w + t.w > cur.avail + 0.1) endLine(false), startLine();
          cur.items.push({ k: "img", w: t.w, h: t.h, im: t.im });
          cur.w += t.w;
          continue;
        }
        // word
        const f = fontFor(t.style);
        let text = encode(f, t.text);
        if (!text) continue;
        let w = widthOf(f, t.style.size, text);
        if (cur.items.length && cur.w + w > cur.avail + 0.1) { endLine(false); startLine(); }
        while (w > cur.avail && text.length > 1) { // single word longer than the line: split it
          let n = text.length - 1;
          while (n > 1 && widthOf(f, t.style.size, text.slice(0, n)) > cur.avail - cur.w) n--;
          const head = text.slice(0, n);
          cur.items.push({ k: "w", text: head, w: widthOf(f, t.style.size, head), style: t.style, font: f });
          cur.w += widthOf(f, t.style.size, head);
          endLine(false); startLine();
          text = text.slice(n);
          w = widthOf(f, t.style.size, text);
        }
        cur.items.push({ k: "w", text, w, style: t.style, font: f });
        cur.w += w;
      }
      if (cur) endLine(true);

      // heights
      const out = [];
      const lastIdx = lines.length - 1;
      lines.forEach((ln, i) => {
        let maxSize = 0, imgH = 0;
        ln.items.forEach((it) => {
          if (it.k === "img") imgH = Math.max(imgH, it.h);
          else if (it.style) maxSize = Math.max(maxSize, it.style.size);
        });
        if (!maxSize && !imgH) maxSize = markStyle.size;
        if (i === 0 && P.marker) maxSize = Math.max(maxSize, P.marker.props.sz ? P.marker.props.sz / 2 : 0);
        const natural = Math.max(maxSize * 1.17, imgH ? imgH + maxSize * 0.22 : 0);
        let h;
        if (rule === "exact") h = lineSpec / 20;
        else if (rule === "atLeast") h = Math.max(natural, lineSpec / 20);
        else h = natural * (lineSpec / 240);
        ln.h = Math.max(h, 1);
        ln.base = ln.h - maxSize * 0.24 - (rule === "exact" ? 0 : 0);
        if (imgH && !maxSize) ln.base = ln.h;
        ln.last = i === lastIdx || ln.forced;
        ln.lastOfPara = i === lastIdx;
      });

      return { lines, left, hanging, firstLine, jc, before, after, props, width, marker: P.marker };
    }

    function drawLineItems(pg, ln, L, x, yTop /* from page top, top of line */, jc) {
      const baseY = PG.h - (yTop + ln.base);
      let gaps = 0;
      ln.items.forEach((it) => { if (it.k === "sp" && it.text) gaps++; });
      let off = 0, extra = 0;
      const avail = ln.avail;
      if (jc === "right") off = Math.max(0, avail - ln.w);
      else if (jc === "center") off = Math.max(0, (avail - ln.w) / 2);
      else if ((jc === "both" || jc === "distribute") && !ln.last && gaps) extra = Math.max(0, (avail - ln.w) / gaps);
      let cx = x + ln.off + off;
      for (const it of ln.items) {
        if (it.k === "img") {
          pg.drawImage(it.im, { x: cx, y: baseY, width: it.w, height: it.h });
          cx += it.w;
          continue;
        }
        const w = it.w + (it.k === "sp" && it.text ? extra : 0);
        if (it.k === "w") {
          const rgb = hex(it.style.color);
          pg.drawText(it.text, { x: cx, y: baseY + it.style.rise, size: it.style.size, font: it.font, color: PDFLib.rgb(rgb[0], rgb[1], rgb[2]) });
          if (it.style.u) pg.drawRectangle({ x: cx, y: baseY + it.style.rise - it.style.size * 0.12, width: w, height: Math.max(0.5, it.style.size * 0.05), color: PDFLib.rgb(rgb[0], rgb[1], rgb[2]) });
          if (it.style.strike) pg.drawRectangle({ x: cx, y: baseY + it.style.rise + it.style.size * 0.28, width: w, height: Math.max(0.5, it.style.size * 0.05), color: PDFLib.rgb(rgb[0], rgb[1], rgb[2]) });
        } else if (it.k === "sp" && it.text && (it.style.u || it.style.strike) && it.style.u) {
          // underline carried across spaces between two underlined words
          const rgb = hex(it.style.color);
          pg.drawRectangle({ x: cx, y: baseY - it.style.size * 0.12, width: w, height: Math.max(0.5, it.style.size * 0.05), color: PDFLib.rgb(rgb[0], rgb[1], rgb[2]) });
        }
        cx += w;
      }
    }

    function paragraphPieces(P, width) {
      const L = layoutParagraph(P, width);
      const pieces = [];
      const pbBefore = !!L.props.pageBreakBefore;
      L.lines.forEach((ln, i) => {
        pieces.push({
          pre: i === 0 ? L.before : 0,
          h: ln.h,
          post: i === L.lines.length - 1 ? L.after : 0,
          pageBreakBefore: i === 0 && pbBefore,
          pageBreakAfter: ln.pageBreakAfter && i < L.lines.length - 1,
          pageBreakAfterLast: ln.pageBreakAfter && i === L.lines.length - 1,
          draw: (pg, x, yTop) => {
            if (i === 0 && L.marker) {
              const m = L.marker;
              const st = { fam: familyKey(m.props.font), bold: !!m.props.b, italic: !!m.props.i, size: (m.props.sz || 20) / 2 };
              const f = fontCache[st.fam + (st.bold ? "B" : "") + (st.italic ? "I" : "")];
              const txt = encode(f, m.text);
              const mw = widthOf(f, st.size, txt);
              const mx = x + Math.max(0, L.hanging ? L.left - L.hanging : L.left + L.firstLine);
              const rgb = hex(m.props.color);
              pg.drawText(txt, { x: mx, y: PG.h - (yTop + ln.base), size: st.size, font: f, color: PDFLib.rgb(rgb[0], rgb[1], rgb[2]) });
            }
            drawLineItems(pg, ln, L, x, yTop, L.jc);
          },
        });
      });
      return pieces;
    }

    function tablePieces(T, width) {
      let grid = T.grid.slice();
      let total = grid.reduce((a, b) => a + b, 0);
      if (!grid.length || !total) { const n = Math.max(1, ...T.rows.map((r) => r.cells.reduce((s, c) => s + c.span, 0))); grid = Array(n).fill(width / n); total = width; }
      const target = T.widthPct ? width * T.widthPct : total;
      const sc = total > width + 0.5 || T.widthPct ? Math.min(width, target) / total : 1;
      grid = grid.map((g) => g * sc);
      const colX = [0];
      grid.forEach((g, i) => colX.push(colX[i] + g));
      const pieces = [];
      for (const row of T.rows) {
        let ci = 0;
        const cells = [];
        let rowH = row.minH;
        for (const cell of row.cells) {
          const x = colX[Math.min(ci, grid.length)];
          const w = colX[Math.min(ci + cell.span, grid.length)] - x;
          ci += cell.span;
          const inner = Math.max(10, w - T.padL - T.padR);
          const sub = cell.vMergeCont ? [] : blocksToPieces(cell.blocks, inner);
          const h = sub.reduce((s, p) => s + p.pre + p.h + p.post, 0) + T.padT + T.padB;
          rowH = Math.max(rowH, h);
          cells.push({ x, w, sub, cell });
        }
        rowH = Math.max(rowH, 4);
        pieces.push({
          pre: 0, h: rowH, post: 0,
          draw: (pg, x, yTop) => {
            for (const c of cells) {
              if (c.cell.fill) {
                const rgb = hex(c.cell.fill);
                pg.drawRectangle({ x: x + c.x, y: PG.h - (yTop + rowH), width: c.w, height: rowH, color: PDFLib.rgb(rgb[0], rgb[1], rgb[2]) });
              }
              let yy = yTop + T.padT;
              for (const sp of c.sub) { yy += sp.pre; sp.draw(pg, x + c.x + T.padL, yy); yy += sp.h + sp.post; }
              const b = c.cell.borders;
              const sides = b ? { top: !!b.top, bottom: !!b.bottom, left: !!(b.left || b.start), right: !!(b.right || b.end) } : T.borders ? { top: true, bottom: true, left: true, right: true } : null;
              if (sides) {
                const x0 = x + c.x, x1 = x0 + c.w, yA = PG.h - yTop, yB = PG.h - (yTop + rowH);
                const line = (ax, ay, bx, by) => pg.drawLine({ start: { x: ax, y: ay }, end: { x: bx, y: by }, thickness: 0.5, color: PDFLib.rgb(0, 0, 0) });
                if (sides.top) line(x0, yA, x1, yA);
                if (sides.bottom) line(x0, yB, x1, yB);
                if (sides.left) line(x0, yA, x0, yB);
                if (sides.right) line(x1, yA, x1, yB);
              }
            }
          },
        });
      }
      return pieces;
    }

    function blocksToPieces(blocks, width) {
      const out = [];
      for (const b of blocks) {
        if (b.type === "p") out.push(...paragraphPieces(b, width));
        else if (b.type === "tbl") out.push(...tablePieces(b, width));
      }
      return out;
    }

    const pieces = blocksToPieces(model.blocks, contentW);

    // flow onto pages
    let pg = null, y = 0, pendingBreak = false;
    const newPage = () => { pg = pdf.addPage([PG.w, PG.h]); y = PG.t; pendingBreak = false; };
    newPage();
    let first = true;
    const topOfPage = () => y <= PG.t + 0.01;
    for (const p of pieces) {
      if (pendingBreak || (p.pageBreakBefore && !topOfPage())) newPage();
      const pre = topOfPage() ? 0 : p.pre;
      if (y + pre + p.h > bottomLimit + 0.5 && !topOfPage()) { newPage(); }
      y += topOfPage() ? 0 : p.pre;
      p.draw(pg, PG.l, y);
      y += p.h + p.post;
      if (p.pageBreakAfter || p.pageBreakAfterLast) pendingBreak = true;
      first = false;
    }
    return pdf.save();
  }

  return { pdfToDocx, docxToPdf };
});
