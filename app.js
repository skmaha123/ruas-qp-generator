/* RUAS AI Question Paper Generator — browser version.
 * Runs entirely client-side (works on GitHub Pages).
 * Libraries: JSZip (DOCX read/write), pdf.js (PDF text), qrcode-generator (QR).
 */
"use strict";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

const state = {
  sources: [],          // loaded files
  questions: [],        // extracted question groups
  usedIds: new Set(),   // questions used in this session
  result: null,         // last generated paper
};

/* =========================================================
 * QUESTION ANALYSIS (ported from the Python version)
 * ======================================================= */
function words(t) { return t.split(/\s+/).filter(Boolean); }

function extractCO(text) {
  const m = text.match(/CO[:\-]?\s*(\d+)/);
  return m ? m[1] : "NA";
}

function scoreQuestion(text) {
  let s = 1;
  if (words(text).length > 6) s += 1;
  const t = text.toLowerCase();
  if (["analyze", "evaluate", "design"].some(w => t.includes(w))) s += 2;
  return Math.min(s, 5);
}

const SCORE_LABELS = ["", "Very Poor", "Poor", "Average", "Good", "Excellent"];
function interpretScore(s) { return SCORE_LABELS[s]; }

function detectBloom(text) {
  const t = text.toLowerCase();
  const has = list => list.some(w => t.includes(w));
  if (has(["design", "develop", "create", "formulate"])) return "L6";
  if (has(["evaluate", "justify", "critique", "recommend"])) return "L5";
  if (has(["analyze", "compare", "differentiate", "examine"])) return "L4";
  if (has(["apply", "solve", "determine", "calculate", "compute", "classify", "find"])) return "L3";
  if (has(["explain", "describe", "discuss", "summarize"])) return "L2";
  if (has(["state", "define", "list", "identify"])) return "L1";
  return "L2";
}

function suggestImprovement(text, bloom, score) {
  const out = [];
  const upgrade = {
    L1: "Upgrade to L2/L3 (Explain/Apply)",
    L2: "Upgrade to L3/L4 (Apply/Analyze)",
    L3: "Upgrade to L4/L5 (Analyze/Evaluate)",
    L4: "Upgrade to L5/L6 (Evaluate/Design)",
    L5: "Upgrade to L6 (Create)",
    L6: "Maintain high-order | Add real-world case",
  };
  out.push(upgrade[bloom] || "Improve cognitive level");
  if (score <= 2) out.push("Increase complexity with multi-step solving");
  if (words(text).length < 8) out.push("Add technical depth/domain details");
  const verbs = {
    L1: ["explain", "describe"], L2: ["apply", "solve"], L3: ["analyze", "compare"],
    L4: ["evaluate", "justify"], L5: ["design", "develop"],
  };
  if (verbs[bloom] && !verbs[bloom].some(v => text.toLowerCase().includes(v))) {
    out.push(`Use verbs like: ${verbs[bloom].join(", ")}`);
  }
  return out.join(" | ");
}

const BLOOM_LEGEND = [
  ["L1 - Remember", "#ffcccc"], ["L2 - Understand", "#ff9999"], ["L3 - Apply", "#ffcc66"],
  ["L4 - Analyze", "#ffff66"], ["L5 - Evaluate", "#99ff99"], ["L6 - Create", "#66cc66"],
];

/* =========================================================
 * XML HELPERS
 * ======================================================= */
function parseXml(str) {
  return new DOMParser().parseFromString(str, "application/xml");
}
function serializeXml(doc) {
  let s = new XMLSerializer().serializeToString(doc);
  if (!s.startsWith("<?xml")) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return s;
}
function toArray(list) { const a = []; for (let i = 0; i < list.length; i++) a.push(list.item ? list.item(i) : list[i]); return a; }
function childEls(el, local) {
  const out = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && n.namespaceURI === W_NS && (!local || n.localName === local)) out.push(n);
  }
  return out;
}
function descEls(el, local) { return toArray(el.getElementsByTagNameNS(W_NS, local)); }

function paragraphText(p) {
  let s = "";
  (function walk(n) {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      if ((c.namespaceURI === W_NS || c.namespaceURI === M_NS) && c.localName === "t") s += c.textContent;
      else if (c.namespaceURI === W_NS && c.localName === "tab") s += "\t";
      else if (c.namespaceURI === W_NS && (c.localName === "br" || c.localName === "cr")) s += "\n";
      else walk(c);
    }
  })(p);
  return s;
}
function cellText(tc) { return childEls(tc, "p").map(paragraphText).join("\n"); }

function setCellText(tc, text) {
  const doc = tc.ownerDocument;
  const ps = childEls(tc, "p");
  let p = ps[0];
  if (!p) { p = doc.createElementNS(W_NS, "w:p"); tc.appendChild(p); }
  ps.slice(1).forEach(x => tc.removeChild(x));
  let rPr = null;
  const firstR = descEls(p, "r")[0];
  if (firstR) { const rp = childEls(firstR, "rPr")[0]; if (rp) rPr = rp.cloneNode(true); }
  childEls(p).filter(c => c.localName !== "pPr").forEach(c => p.removeChild(c));
  // remove non-W children (e.g. math) too
  for (let c = p.firstChild; c;) { const nx = c.nextSibling; if (c.nodeType === 1 && c.namespaceURI !== W_NS) p.removeChild(c); c = nx; }
  const r = doc.createElementNS(W_NS, "w:r");
  if (rPr) r.appendChild(rPr);
  String(text).split("\n").forEach((line, i) => {
    if (i > 0) r.appendChild(doc.createElementNS(W_NS, "w:br"));
    const t = doc.createElementNS(W_NS, "w:t");
    t.setAttributeNS(XML_NS, "xml:space", "preserve");
    t.appendChild(doc.createTextNode(line));
    r.appendChild(t);
  });
  p.appendChild(r);
}

/* =========================================================
 * FILE LOADING & EXTRACTION
 * ======================================================= */
let uidCounter = 0;

async function loadDocx(name, buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error(`${name} is not a valid Word document`);
  const xml = parseXml(await docFile.async("string"));
  const relsFile = zip.file("word/_rels/document.xml.rels");
  const rels = relsFile ? parseXml(await relsFile.async("string")) : null;
  return { uid: ++uidCounter, name, kind: "docx", buffer, zip, xml, rels };
}

function extractFromDocx(src) {
  const body = childEls(src.xml.documentElement, "body")[0];
  const groups = new Map();
  if (!body) return [];
  for (const tbl of childEls(body, "tbl")) {
    for (const tr of childEls(tbl, "tr")) {
      const cells = childEls(tr, "tc").map(tc => cellText(tc).trim());
      if (cells.length < 2) continue;
      const qno = cells[0], qtext = cells[1];
      const marks = cells.length > 2 ? cells[2] : "NA";
      const co = cells.length > 3 ? cells[3] : extractCO(qtext);
      if (/^\d+/.test(qno) && qtext.length > 5) {
        const base = qno.match(/^\d+/)[0];
        if (!groups.has(base)) groups.set(base, []);
        groups.get(base).push({ tr, text: qtext, marks, co });
      }
    }
  }
  return [...groups].map(([qno, rows]) => ({
    id: `${src.uid}#${qno}`, qno, rows,
    text: rows[0].text, marks: rows[0].marks, co: rows[0].co,
    module: "M" + qno, source: src.name, src,
  }));
}

async function extractFromPdf(name, buffer) {
  const src = { uid: ++uidCounter, name, kind: "pdf" };
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;
  const items = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const txt = content.items.map(i => i.str).join(" ").replace(/\s+/g, " ");
    for (const m of txt.matchAll(/(\d+\s+.*?)(?=\d+\s+|$)/g)) {
      const text = m[1].trim();
      const qno = text.match(/^\d+/)[0];
      items.push({
        id: `${src.uid}#p${p}#${items.length}`, qno, rows: null,
        text, marks: "NA", co: extractCO(text), module: "M" + qno, source: name, src,
      });
    }
  }
  return { src, items };
}

/* =========================================================
 * GENERATION
 * ======================================================= */
function randomInt(n) {
  const a = new Uint32Array(1);
  crypto.getRandomValues(a);
  return a[0] % n;
}

function selectQuestions(questions, n, usedIds) {
  const pool = new Map();
  for (const q of questions) {
    if (!pool.has(q.module)) pool.set(q.module, []);
    pool.get(q.module).push(q);
  }
  const keys = [...pool.keys()].sort((a, b) => {
    const na = parseInt(a.slice(1), 10), nb = parseInt(b.slice(1), 10);
    return (isNaN(na) ? 1e9 : na) - (isNaN(nb) ? 1e9 : nb) || a.localeCompare(b);
  });
  const selected = [];
  for (const k of keys.slice(0, n)) {
    let available = pool.get(k).filter(q => !usedIds.has(q.id));
    if (!available.length) available = pool.get(k);
    const q = available[randomInt(available.length)];
    selected.push(q);
    usedIds.add(q.id);
  }
  return selected;
}

const CONTENT_TYPES = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp",
  tif: "image/tiff", tiff: "image/tiff", emf: "image/x-emf", wmf: "image/x-wmf", svg: "image/svg+xml",
  bin: "application/vnd.openxmlformats-officedocument.oleObject",
};

function resolvePartPath(target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = ("word/" + target).split("/");
  const out = [];
  for (const p of parts) { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); }
  return out.join("/");
}

/* Copies images / embedded objects referenced by a row from another .docx into the output */
function makeRelRemapper(outZip, relsDoc, ctDoc) {
  const relRoot = relsDoc.documentElement;
  const existing = new Set(toArray(relRoot.getElementsByTagNameNS(PKG_REL_NS, "Relationship")).map(r => r.getAttribute("Id")));
  const cache = new Map();
  let counter = 0;

  function nextId() {
    let id;
    do { id = "rIdQP" + (++counter); } while (existing.has(id));
    existing.add(id);
    return id;
  }
  function ensureContentType(path) {
    const ext = path.split(".").pop().toLowerCase();
    const root = ctDoc.documentElement;
    const defaults = toArray(root.getElementsByTagNameNS(CT_NS, "Default"));
    if (defaults.some(d => (d.getAttribute("Extension") || "").toLowerCase() === ext)) return;
    const d = ctDoc.createElementNS(CT_NS, "Default");
    d.setAttribute("Extension", ext);
    d.setAttribute("ContentType", CONTENT_TYPES[ext] || "application/octet-stream");
    root.insertBefore(d, root.firstChild);
  }

  return async function remap(node, src) {
    if (!src.rels) return;
    const srcRels = new Map(toArray(src.rels.documentElement.getElementsByTagNameNS(PKG_REL_NS, "Relationship"))
      .map(r => [r.getAttribute("Id"), r]));
    const els = [node, ...toArray(node.getElementsByTagName("*"))];
    for (const el of els) {
      const attrs = toArray(el.attributes).filter(a => a.namespaceURI === R_NS);
      for (const a of attrs) {
        const oldId = a.value;
        const key = src.uid + "|" + oldId;
        if (!cache.has(key)) {
          const rel = srcRels.get(oldId);
          if (!rel) continue;
          const newId = nextId();
          const mode = rel.getAttribute("TargetMode");
          let target = rel.getAttribute("Target");
          if (mode !== "External") {
            const srcPath = resolvePartPath(target);
            const f = src.zip.file(srcPath);
            if (f) {
              const dir = srcPath.substring(0, srcPath.lastIndexOf("/"));
              const newPath = `${dir}/qp${counter}_${srcPath.split("/").pop()}`;
              outZip.file(newPath, await f.async("uint8array"));
              ensureContentType(newPath);
              target = newPath.startsWith("word/") ? newPath.slice(5) : "/" + newPath;
            }
          }
          const r = relsDoc.createElementNS(PKG_REL_NS, "Relationship");
          r.setAttribute("Id", newId);
          r.setAttribute("Type", rel.getAttribute("Type"));
          r.setAttribute("Target", target);
          if (mode) r.setAttribute("TargetMode", mode);
          relRoot.appendChild(r);
          cache.set(key, newId);
        }
        el.setAttributeNS(R_NS, a.name, cache.get(key));
      }
    }
  };
}

async function buildPaper(selected, template) {
  const zip = await JSZip.loadAsync(template.buffer);
  const xml = parseXml(await zip.file("word/document.xml").async("string"));
  const relsPath = "word/_rels/document.xml.rels";
  const rels = zip.file(relsPath)
    ? parseXml(await zip.file(relsPath).async("string"))
    : parseXml(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PKG_REL_NS}"/>`);
  const ct = parseXml(await zip.file("[Content_Types].xml").async("string"));

  const body = childEls(xml.documentElement, "body")[0];
  const tables = childEls(body, "tbl");
  if (!tables.length) throw new Error(`The template "${template.name}" has no table to put questions in.`);
  const table = tables[tables.length - 1];
  const rows = childEls(table, "tr");
  if (!rows.length) throw new Error("The question table in the template is empty.");

  // keep header only
  rows.slice(1).forEach(r => table.removeChild(r));
  const header = rows[0];
  const hcells = childEls(header, "tc");
  ["Q No", "Questions", "Marks", "CO"].forEach((h, i) => { if (hcells[i]) setCellText(hcells[i], h); });

  const remap = makeRelRemapper(zip, rels, ct);

  for (let i = 0; i < selected.length; i++) {
    const q = selected[i];
    const num = String(i + 1);
    if (q.rows) {
      // copy original rows, including equations / images
      for (const part of q.rows) {
        const clone = xml.importNode(part.tr, true);
        if (q.src !== template) await remap(clone, q.src);
        const firstTc = childEls(clone, "tc")[0];
        if (firstTc) setCellText(firstTc, cellText(firstTc).trim().replace(/^\d+/, num));
        table.appendChild(clone);
      }
    } else {
      // PDF question: build a fresh row based on the header's formatting
      const clone = header.cloneNode(true);
      descEls(clone, "tblHeader").forEach(e => e.parentNode.removeChild(e));
      descEls(clone, "b").forEach(e => e.parentNode.removeChild(e));
      const cells = childEls(clone, "tc");
      [num, q.text, q.marks, q.co].forEach((v, k) => { if (cells[k]) setCellText(cells[k], v); });
      table.appendChild(clone);
    }
  }

  zip.file("word/document.xml", serializeXml(xml));
  zip.file(relsPath, serializeXml(rels));
  zip.file("[Content_Types].xml", serializeXml(ct));
  return zip.generateAsync({
    type: "uint8array",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
  });
}

function analyse(selected) {
  return selected.map((q, i) => {
    const score = scoreQuestion(q.text);
    const bloom = detectBloom(q.text);
    return {
      no: i + 1,
      preview: q.text ? q.text.slice(0, 40) : "[Equation/Image Question]",
      source: q.source, co: q.co, bloom,
      score, scoreText: `${score} (${interpretScore(score)})`,
      suggestion: suggestImprovement(q.text, bloom, score),
    };
  });
}

/* =========================================================
 * AUTHORIZATION KEY, QR & ENCRYPTION
 * ======================================================= */
function generateAccessKey(length = 20) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let out = "";
  const buf = new Uint8Array(1);
  while (out.length < length) {
    crypto.getRandomValues(buf);
    if (buf[0] < 252) out += chars[buf[0] % 36]; // unbiased
  }
  return out;
}

const MAGIC = new TextEncoder().encode("RUASQP01");

async function deriveKey(pass, salt) {
  const base = await crypto.subtle.importKey("raw", new TextEncoder().encode(pass), "PBKDF2", false, ["deriveKey"]);
  return crypto.subtle.deriveKey(
    { name: "PBKDF2", salt, iterations: 250000, hash: "SHA-256" },
    base, { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
}

async function encryptPaper(bytes, pass) {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const key = await deriveKey(pass, salt);
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, bytes));
  const out = new Uint8Array(MAGIC.length + 16 + 12 + ct.length);
  out.set(MAGIC, 0); out.set(salt, 8); out.set(iv, 24); out.set(ct, 36);
  return out;
}

async function decryptPaper(data, pass) {
  for (let i = 0; i < MAGIC.length; i++) if (data[i] !== MAGIC[i]) throw new Error("Not a protected question paper (.qp) file.");
  const salt = data.slice(8, 24), iv = data.slice(24, 36), ct = data.slice(36);
  const key = await deriveKey(pass, salt);
  try {
    return new Uint8Array(await crypto.subtle.decrypt({ name: "AES-GCM", iv }, key, ct));
  } catch {
    throw new Error("Invalid key — access denied.");
  }
}

function qrDataUrl(text) {
  const qr = qrcode(0, "M");
  qr.addData(text);
  qr.make();
  const count = qr.getModuleCount(), box = 10, border = 4;
  const size = (count + border * 2) * box;
  const canvas = document.createElement("canvas");
  canvas.width = canvas.height = size;
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, size, size);
  ctx.fillStyle = "#000";
  for (let r = 0; r < count; r++)
    for (let c = 0; c < count; c++)
      if (qr.isDark(r, c)) ctx.fillRect((c + border) * box, (r + border) * box, box, box);
  return canvas.toDataURL("image/png");
}

/* =========================================================
 * UI
 * ======================================================= */
function download(data, filename, type) {
  const blob = data instanceof Blob ? data : new Blob([data], { type: type || "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function setStatus(el, msg, kind) {
  el.textContent = msg;
  el.className = "status" + (kind ? " " + kind : "");
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function initUI() {
  const $ = id => document.getElementById(id);
  if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = "lib/pdf.worker.min.js";
  }

  // legend
  $("legend").innerHTML = BLOOM_LEGEND.map(([t, c]) => `<span style="background:${c}">${t}</span>`).join("");

  const drop = $("dropZone"), input = $("fileInput");
  ["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", e => handleFiles([...e.dataTransfer.files]));
  input.addEventListener("change", () => { handleFiles([...input.files]); input.value = ""; });

  async function handleFiles(files) {
    files = files.filter(f => /\.(docx|pdf)$/i.test(f.name));
    if (!files.length) return;
    const status = $("uploadStatus");
    setStatus(status, "Reading files…");
    state.sources = []; state.questions = []; state.usedIds.clear();
    const errors = [];
    for (const f of files) {
      try {
        const buf = await f.arrayBuffer();
        if (/\.docx$/i.test(f.name)) {
          const src = await loadDocx(f.name, buf);
          const items = extractFromDocx(src);
          src.count = items.length;
          state.sources.push(src); state.questions.push(...items);
        } else {
          if (!window.pdfjsLib) throw new Error("PDF library failed to load");
          const { src, items } = await extractFromPdf(f.name, buf);
          src.count = items.length;
          state.sources.push(src); state.questions.push(...items);
        }
      } catch (err) {
        console.error(err);
        errors.push(`${f.name}: ${err.message}`);
      }
    }
    const template = state.sources.find(s => s.kind === "docx");
    $("fileList").innerHTML = state.sources.map(s =>
      `<li class="${s === template ? "template" : ""}" title="${s === template ? "Template" : ""}">${escapeHtml(s.name)}<span class="count">(${s.count} Q${s === template ? ", template" : ""})</span></li>`).join("");
    const slots = new Set(state.questions.map(q => q.module)).size;
    if (slots) $("numQuestions").value = slots;
    $("btnGenerate").disabled = !state.questions.length || !template;
    let msg = `${state.sources.length} file(s) loaded · ${state.questions.length} questions extracted · ${slots} question slots`;
    if (!template && state.sources.length) msg += " — upload at least one .docx to use as the template";
    if (errors.length) msg += "\n" + errors.join("\n");
    setStatus(status, msg, errors.length || !template ? "err" : "ok");
  }

  $("btnGenerate").addEventListener("click", async () => {
    const template = state.sources.find(s => s.kind === "docx");
    if (!state.questions.length || !template) return alert("Upload files first");
    const n = parseInt($("numQuestions").value, 10);
    if (!n || n < 1) return alert("Enter the number of questions");
    const btn = $("btnGenerate");
    btn.disabled = true; btn.textContent = "Generating…";
    try {
      const selected = selectQuestions(state.questions, n, state.usedIds);
      const docx = await buildPaper(selected, template);
      const key = generateAccessKey();
      const protectedBytes = await encryptPaper(docx, key);
      const base = `RUAS_QP_${1000 + randomInt(9000)}`;
      const qr = qrDataUrl(key);
      state.result = { base, key, docx, protectedBytes, qr };
      showResult(analyse(selected));
    } catch (err) {
      console.error(err);
      alert("Could not generate the paper:\n" + err.message);
    } finally {
      btn.disabled = false; btn.textContent = "Generate / Regenerate";
    }
  });

  function showResult(rows) {
    const r = state.result;
    $("result").hidden = false;
    $("keyText").textContent = r.key;
    $("qrImg").src = r.qr;
    $("savedName").textContent = `Paper: ${r.base}.qp — share the key separately from the file.`;
    $("analysis").querySelector("tbody").innerHTML = rows.map(x => `
      <tr class="b${x.bloom.slice(1)}">
        <td class="c">${x.no}</td><td>${escapeHtml(x.preview)}</td><td>${escapeHtml(x.source)}</td>
        <td class="c">${escapeHtml(x.co)}</td><td class="c">${x.bloom}</td><td class="c">${x.scoreText}</td>
        <td>${escapeHtml(x.suggestion)}</td>
      </tr>`).join("");
    const count = arr => arr.reduce((m, v) => (m[v] = (m[v] || 0) + 1, m), {});
    const fmt = obj => Object.entries(obj).map(([k, v]) => `${escapeHtml(k)}: ${v}`).join(" · ") || "—";
    const avg = rows.length ? (rows.reduce((s, x) => s + x.score, 0) / rows.length).toFixed(2) : "0";
    $("dashboard").innerHTML = `
      <div class="tile"><b>Average score</b><span>${avg} / 5</span></div>
      <div class="tile"><b>Bloom levels</b><span>${fmt(count(rows.map(x => x.bloom)))}</span></div>
      <div class="tile"><b>Course outcomes</b><span>${fmt(count(rows.map(x => /^\d+$/.test(x.co) ? "CO" + x.co : x.co)))}</span></div>`;
    $("result").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  $("dlProtected").addEventListener("click", () => download(state.result.protectedBytes, `${state.result.base}.qp`));
  $("dlKey").addEventListener("click", () => download(state.result.key, `${state.result.base}_KEY.txt`, "text/plain"));
  $("dlQr").addEventListener("click", async () => download(await (await fetch(state.result.qr)).blob(), `${state.result.base}_QR.png`));
  $("dlDocx").addEventListener("click", () => download(state.result.docx, `${state.result.base}.docx`,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));

  $("btnClear").addEventListener("click", () => {
    state.sources = []; state.questions = []; state.usedIds.clear(); state.result = null;
    $("fileList").innerHTML = "";
    $("result").hidden = true;
    $("btnGenerate").disabled = true;
    setStatus($("uploadStatus"), "");
  });

  $("btnOpen").addEventListener("click", async () => {
    const status = $("openStatus");
    const file = $("qpFile").files[0];
    const key = $("qpKey").value.trim().toUpperCase();
    if (!file) return setStatus(status, "Choose a .qp file first.", "err");
    if (!key) return setStatus(status, "Enter the authorization key.", "err");
    setStatus(status, "Verifying key…");
    try {
      const bytes = await decryptPaper(new Uint8Array(await file.arrayBuffer()), key);
      download(bytes, file.name.replace(/\.qp$/i, "") + ".docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      setStatus(status, "Valid key — document access granted. The .docx has been downloaded.", "ok");
    } catch (err) {
      setStatus(status, err.message, "err");
    }
  });
}

if (typeof document !== "undefined") {
  document.addEventListener("DOMContentLoaded", initUI);
}
if (typeof module !== "undefined") {
  module.exports = { loadDocx, extractFromDocx, selectQuestions, buildPaper, analyse, encryptPaper, decryptPaper, generateAccessKey, detectBloom };
}
