/* PDF Toolkit core - pure logic (works in the browser and in Node for testing). */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Core = factory();
})(typeof self !== "undefined" ? self : this, function () {
  const A4 = { w: 595.28, h: 841.89 };

  /** "2, 5-7" -> Set{2,5,6,7}; throws readable errors. */
  function parsePages(text, total) {
    const pages = new Set();
    const parts = String(text).replace(/\s+/g, "").split(",").filter(Boolean);
    for (const part of parts) {
      const m = /^(\d+)(?:-(\d+))?$/.exec(part);
      if (!m) throw new Error("Enter page numbers like:  2, 5-7");
      let a = parseInt(m[1], 10), b = m[2] ? parseInt(m[2], 10) : a;
      if (a > b) [a, b] = [b, a];
      for (let i = a; i <= b; i++) pages.add(i);
    }
    if (!pages.size) throw new Error("Enter the pages to delete, e.g.  2, 5-7");
    const min = Math.min(...pages), max = Math.max(...pages);
    if (min < 1 || max > total) throw new Error("Page out of range. The PDF has " + total + " page(s).");
    if (pages.size >= total) throw new Error("You can't delete every page.");
    return pages;
  }

  async function mergePdfs(PDFLib, buffers) {
    const out = await PDFLib.PDFDocument.create();
    for (const buf of buffers) {
      const src = await PDFLib.PDFDocument.load(buf, { ignoreEncryption: false });
      const pages = await out.copyPages(src, src.getPageIndices());
      pages.forEach((p) => out.addPage(p));
    }
    return out.save();
  }

  async function pageCount(PDFLib, buffer) {
    const doc = await PDFLib.PDFDocument.load(buffer);
    return doc.getPageCount();
  }

  async function removePages(PDFLib, buffer, pagesText) {
    const doc = await PDFLib.PDFDocument.load(buffer);
    const remove = parsePages(pagesText, doc.getPageCount());
    [...remove].sort((a, b) => b - a).forEach((p) => doc.removePage(p - 1));
    return doc.save();
  }

  /** images: [{bytes: Uint8Array (JPEG), }], mode: "a4" | "original" */
  async function imagesToPdf(PDFLib, images, mode) {
    const doc = await PDFLib.PDFDocument.create();
    for (const im of images) {
      const img = await doc.embedJpg(im.bytes);
      const iw = img.width, ih = img.height;
      if (mode === "original") {
        const page = doc.addPage([iw * 0.75, ih * 0.75]);
        page.drawImage(img, { x: 0, y: 0, width: iw * 0.75, height: ih * 0.75 });
      } else {
        const landscape = iw > ih;
        const pw = landscape ? A4.h : A4.w, ph = landscape ? A4.w : A4.h;
        const margin = 24;
        const scale = Math.min((pw - margin * 2) / iw, (ph - margin * 2) / ih);
        const w = iw * scale, h = ih * scale;
        const page = doc.addPage([pw, ph]);
        page.drawImage(img, { x: (pw - w) / 2, y: (ph - h) / 2, width: w, height: h });
      }
    }
    return doc.save();
  }

  return { parsePages, mergePdfs, pageCount, removePages, imagesToPdf };
});
