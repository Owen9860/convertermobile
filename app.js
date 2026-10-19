/* PDF Toolkit UI - By Owen */
const $ = (s) => document.querySelector(s);
const isNative = () => !!(window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform());

const state = { images: [], merge: [], del: null, delPages: 0, p2w: null, w2p: null };
const MIME_PDF = "application/pdf";
const MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/* ---------- navigation ---------- */
document.querySelectorAll(".tab").forEach((t) =>
  t.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("active", x === t));
    document.querySelectorAll(".view").forEach((v) => v.classList.toggle("active", v.id === t.dataset.view));
    setStatus("");
  })
);

function setStatus(msg, kind) {
  const el = $("#status");
  el.textContent = msg;
  el.className = "status " + (kind || "");
}
function busy(on) {
  document.body.classList.toggle("busy", on);
  document.querySelectorAll(".primary").forEach((b) => (b.disabled = on));
}
function fmtSize(n) {
  return n > 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB";
}
function stamp() {
  const d = new Date(), p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
}
function move(arr, i, d) {
  const j = i + d;
  if (j < 0 || j >= arr.length) return;
  [arr[i], arr[j]] = [arr[j], arr[i]];
}

/* ---------- delivering the result ---------- */
function toBase64(bytes) {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  return btoa(bin);
}
async function deliver(bytes, name, mime) {
  mime = mime || MIME_PDF;
  if (isNative()) {
    const call = (p, m, o) => window.Capacitor.nativePromise(p, m, o);
    const res = await call("Filesystem", "writeFile", { path: name, data: toBase64(bytes), directory: "CACHE", recursive: true });
    await call("Share", "share", { title: name, url: res.uri, dialogTitle: mime === MIME_DOCX ? "Save or share your Word file" : "Save or share your PDF" });
  } else {
    const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
    const a = Object.assign(document.createElement("a"), { href: url, download: name });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }
}

/* ---------- list rendering (images + merge) ---------- */
function renderList(listEl, items, kind) {
  listEl.innerHTML = "";
  if (!items.length) {
    listEl.innerHTML = '<div class="empty">Nothing added yet</div>';
    return;
  }
  items.forEach((it, i) => {
    const row = document.createElement("div");
    row.className = "row";
    const thumb = kind === "images" ? `<img src="${it.url}" alt="">` : `<div class="pdfico">PDF</div>`;
    row.innerHTML = `${thumb}<div class="meta"><b></b><span>${i + 1} · ${fmtSize(it.file.size)}</span></div>
      <div class="acts"><button data-a="up" aria-label="Move up">▲</button><button data-a="down" aria-label="Move down">▼</button><button data-a="del" aria-label="Remove">✕</button></div>`;
    row.querySelector("b").textContent = it.file.name;
    row.querySelector(".acts").addEventListener("click", (e) => {
      const a = e.target.dataset.a;
      if (!a) return;
      if (a === "up") move(items, i, -1);
      if (a === "down") move(items, i, 1);
      if (a === "del") { if (it.url) URL.revokeObjectURL(it.url); items.splice(i, 1); }
      renderList(listEl, items, kind);
    });
    listEl.appendChild(row);
  });
}

/* ---------- Images -> PDF ---------- */
$("#imgInput").addEventListener("change", (e) => {
  for (const f of e.target.files) state.images.push({ file: f, url: URL.createObjectURL(f) });
  e.target.value = "";
  renderList($("#imgList"), state.images, "images");
  setStatus(state.images.length + " image(s) added.");
});

async function fileToJpeg(file) {
  let bmp;
  try {
    bmp = await createImageBitmap(file, { imageOrientation: "from-image" });
  } catch (_) {
    bmp = await new Promise((res, rej) => {
      const img = new Image();
      img.onload = () => res(img);
      img.onerror = () => rej(new Error("Can't read " + file.name));
      img.src = URL.createObjectURL(file);
    });
  }
  const w0 = bmp.width, h0 = bmp.height;
  const k = Math.min(1, 3000 / Math.max(w0, h0));
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(w0 * k);
  canvas.height = Math.round(h0 * k);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(bmp, 0, 0, canvas.width, canvas.height);
  const blob = await new Promise((r) => canvas.toBlob(r, "image/jpeg", 0.92));
  if (!blob) throw new Error("Can't convert " + file.name);
  return { bytes: new Uint8Array(await blob.arrayBuffer()) };
}

$("#imgGo").addEventListener("click", async () => {
  if (!state.images.length) return setStatus("Add at least one image first.", "err");
  busy(true);
  try {
    const jpgs = [];
    for (let i = 0; i < state.images.length; i++) {
      setStatus(`Preparing image ${i + 1} of ${state.images.length}…`);
      jpgs.push(await fileToJpeg(state.images[i].file));
    }
    setStatus("Creating PDF…");
    const bytes = await Core.imagesToPdf(PDFLib, jpgs, $("#pageMode").value);
    await deliver(bytes, `images-${stamp()}.pdf`);
    setStatus("✔ PDF created. Choose where to save it.", "ok");
  } catch (e) {
    setStatus("Failed: " + e.message, "err");
  }
  busy(false);
});

/* ---------- Merge ---------- */
$("#mergeInput").addEventListener("change", (e) => {
  for (const f of e.target.files) state.merge.push({ file: f });
  e.target.value = "";
  renderList($("#mergeList"), state.merge, "pdf");
  setStatus(state.merge.length + " PDF(s) in the list.");
});

$("#mergeGo").addEventListener("click", async () => {
  if (state.merge.length < 2) return setStatus("Add at least 2 PDFs first.", "err");
  busy(true);
  try {
    setStatus("Merging…");
    const bufs = [];
    for (const it of state.merge) bufs.push(await it.file.arrayBuffer());
    const bytes = await Core.mergePdfs(PDFLib, bufs);
    await deliver(bytes, `merged-${stamp()}.pdf`);
    setStatus("✔ Merged. Choose where to save it.", "ok");
  } catch (e) {
    setStatus("Failed: " + e.message, "err");
  }
  busy(false);
});

/* ---------- Delete pages ---------- */
$("#delInput").addEventListener("change", async (e) => {
  const f = e.target.files[0];
  e.target.value = "";
  if (!f) return;
  try {
    state.delPages = await Core.pageCount(PDFLib, await f.arrayBuffer());
    state.del = f;
    $("#delInfo").textContent = `${f.name} · ${state.delPages} page(s)`;
    setStatus("");
  } catch (err) {
    state.del = null;
    $("#delInfo").textContent = "No PDF chosen";
    setStatus("Can't open that PDF (it may be password protected).", "err");
  }
});

$("#delGo").addEventListener("click", async () => {
  if (!state.del) return setStatus("Choose a PDF first.", "err");
  busy(true);
  try {
    const bytes = await Core.removePages(PDFLib, await state.del.arrayBuffer(), $("#pagesText").value);
    const base = state.del.name.replace(/\.pdf$/i, "");
    await deliver(bytes, `${base} (edited).pdf`);
    setStatus("✔ Pages deleted. Choose where to save it.", "ok");
  } catch (e) {
    setStatus(e.message, "err");
  }
  busy(false);
});

/* ---------- PDF -> Word ---------- */
let pdfjsPromise = null;
function loadPdfjs() {
  if (!pdfjsPromise) {
    pdfjsPromise = import("./lib/pdf.min.js").then((m) => {
      m.GlobalWorkerOptions.workerSrc = new URL("./lib/pdf.worker.min.js", document.baseURI).href;
      return m;
    });
  }
  return pdfjsPromise;
}

$("#p2wInput").addEventListener("change", (e) => {
  const f = e.target.files[0];
  e.target.value = "";
  if (!f) return;
  state.p2w = f;
  $("#p2wInfo").textContent = `${f.name} · ${fmtSize(f.size)}`;
  setStatus("");
});

$("#p2wGo").addEventListener("click", async () => {
  if (!state.p2w) return setStatus("Choose a PDF first.", "err");
  busy(true);
  try {
    setStatus("Opening PDF…");
    const pdfjs = await loadPdfjs();
    const data = await state.p2w.arrayBuffer();
    const bytes = await Convert.pdfToDocx(pdfjs, JSZip, data, $("#p2wMode").value, (d, n) => setStatus(`Converting page ${d} of ${n}…`));
    const base = state.p2w.name.replace(/\.pdf$/i, "");
    await deliver(bytes, `${base}.docx`, MIME_DOCX);
    setStatus("✔ Word file created. Choose where to save it.", "ok");
  } catch (e) {
    if (e && e.code === "NO_TEXT") $("#p2wMode").value = "image";
    setStatus(e && e.name === "PasswordException" ? "This PDF is password protected." : "Failed: " + (e && e.message ? e.message : e), "err");
  }
  busy(false);
});

/* ---------- Word -> PDF ---------- */
$("#w2pInput").addEventListener("change", (e) => {
  const f = e.target.files[0];
  e.target.value = "";
  if (!f) return;
  state.w2p = f;
  $("#w2pInfo").textContent = `${f.name} · ${fmtSize(f.size)}`;
  setStatus(/\.doc$/i.test(f.name) ? "Old .doc files aren't supported. Open it in Word and Save As .docx first." : "", "err");
});

async function imageToPng(bytes, ext) {
  const bmp = await createImageBitmap(new Blob([bytes], { type: "image/" + (ext === "jpg" ? "jpeg" : ext) }));
  const c = document.createElement("canvas");
  c.width = bmp.width; c.height = bmp.height;
  c.getContext("2d").drawImage(bmp, 0, 0);
  const blob = await new Promise((r) => c.toBlob(r, "image/png"));
  return new Uint8Array(await blob.arrayBuffer());
}

$("#w2pGo").addEventListener("click", async () => {
  if (!state.w2p) return setStatus("Choose a Word file first.", "err");
  if (/\.doc$/i.test(state.w2p.name)) return setStatus("Old .doc files aren't supported. Open it in Word and Save As .docx first.", "err");
  busy(true);
  try {
    setStatus("Converting…");
    const bytes = await Convert.docxToPdf(JSZip, PDFLib, await state.w2p.arrayBuffer(), { convertImage: imageToPng });
    const base = state.w2p.name.replace(/\.docx$/i, "");
    await deliver(bytes, `${base}.pdf`, MIME_PDF);
    setStatus("✔ PDF created. Choose where to save it.", "ok");
  } catch (e) {
    setStatus("Failed: " + (e && e.message ? e.message : e), "err");
  }
  busy(false);
});

renderList($("#imgList"), state.images, "images");
renderList($("#mergeList"), state.merge, "pdf");
