import * as pdfjs from "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs";
import { PDFDocument, degrees } from "https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/+esm";
import { createWorker } from "https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/+esm";
import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";
import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

pdfjs.GlobalWorkerOptions.workerSrc =
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs";

const SIGN_X = 426;
const SIGN_Y = 129;
const SIGN_WIDTH = 154;

const NAME_HEADERS = ["deligate name", "delegate name", "candidate name", "name"];
const COMPANY_HEADERS = ["company name", "company", "employer"];
const DATE_HEADERS = ["training date", "date", "course date"];
const TITLE_HEADERS = ["traninig title", "training title", "course", "title"];
const MONTHS = {
  jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2, apr: 3, april: 3,
  may: 4, jun: 5, june: 5, jul: 6, july: 6, aug: 7, august: 7, sep: 8, sept: 8,
  september: 8, oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11,
};

function normHeader(value) {
  return String(value || "").toLowerCase().replace(/[:_]/g, " ").replace(/\s+/g, " ").trim();
}
function findColumn(headers, aliases) {
  for (const alias of aliases) {
    const idx = headers.findIndex((h) => h === alias || h.includes(alias));
    if (idx !== -1) return idx;
  }
  return -1;
}
function dateKey(date) {
  if (!date || Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}
function parseTrainingEndDate(raw) {
  const compact = String(raw || "").replace(/,/g, " ").replace(/\s+/g, " ").trim().replace(/(\d{1,2})(?:st|nd|rd|th)/gi, "$1");
  const ranged = compact.match(/(\d{1,2})\s*[-–]\s*(\d{1,2})\s*([A-Za-z]+)\s*(\d{4})/i);
  if (ranged && MONTHS[ranged[3].toLowerCase()] != null) return new Date(Number(ranged[4]), MONTHS[ranged[3].toLowerCase()], Number(ranged[2]));
  const cross = compact.match(/(\d{1,2})\s*([A-Za-z]+)\s*[-–]\s*(\d{1,2})\s*([A-Za-z]+)\s*(\d{4})/i);
  if (cross && MONTHS[cross[4].toLowerCase()] != null) return new Date(Number(cross[5]), MONTHS[cross[4].toLowerCase()], Number(cross[3]));
  const single = compact.match(/(\d{1,2})\s*([A-Za-z]+)\s*(\d{4})/i);
  if (single && MONTHS[single[2].toLowerCase()] != null) return new Date(Number(single[3]), MONTHS[single[2].toLowerCase()], Number(single[1]));
  return null;
}
function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: "", raw: false });
  const headers = (rows[0] || []).map(normHeader);
  const nameIdx = findColumn(headers, NAME_HEADERS);
  const companyIdx = findColumn(headers, COMPANY_HEADERS);
  const dateIdx = findColumn(headers, DATE_HEADERS);
  const titleIdx = findColumn(headers, TITLE_HEADERS);
  if (nameIdx === -1 || companyIdx === -1) throw new Error("Excel must have Company Name and Delegate Name columns.");
  const records = [];
  for (let i = 1; i < rows.length; i += 1) {
    const name = String(rows[i][nameIdx] || "").trim();
    const company = String(rows[i][companyIdx] || "").trim();
    if (!name && !company) continue;
    const dateRaw = dateIdx === -1 ? "" : String(rows[i][dateIdx] || "").trim();
    records.push({
      name, company, dateRaw, dateKey: dateKey(parseTrainingEndDate(dateRaw)),
      title: titleIdx === -1 ? "" : String(rows[i][titleIdx] || "").trim(),
    });
  }
  return records;
}

function normalizeName(value) {
  return String(value || "").toUpperCase().replace(/[,.`']/g, "").replace(/[^A-Z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
}
function tokens(value) {
  return normalizeName(value).split(" ").filter((t) => t.length > 1);
}
function levenshtein(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const prev = Array.from({ length: b.length + 1 }, (_, j) => j);
  const curr = new Array(b.length + 1);
  for (let i = 1; i <= a.length; i += 1) {
    curr[0] = i;
    for (let j = 1; j <= b.length; j += 1) {
      curr[j] = Math.min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    for (let j = 0; j <= b.length; j += 1) prev[j] = curr[j];
  }
  return prev[b.length];
}
function similarToken(a, b) {
  if (a === b) return true;
  return Math.max(a.length, b.length) >= 5 && levenshtein(a, b) <= 1;
}
function scoreNames(ocrName, excelName) {
  const a = normalizeName(ocrName);
  const b = normalizeName(excelName);
  if (!a || !b) return 0;
  if (a === b) return 100;
  if (a.includes(b) || b.includes(a)) return Math.round(88 + (Math.min(a.length, b.length) / Math.max(a.length, b.length)) * 10);
  const ta = tokens(a);
  const tb = tokens(b);
  if (!ta.length || !tb.length) return 0;
  const used = new Set();
  let overlap = 0;
  for (const token of ta) {
    const idx = tb.findIndex((other, i) => !used.has(i) && similarToken(token, other));
    if (idx !== -1) { used.add(idx); overlap += 1; }
  }
  const jaccard = overlap / new Set([...ta, ...tb]).size;
  const prefixBonus = ta.slice(0, 2).join(" ") === tb.slice(0, 2).join(" ") ? 12 : 0;
  return Math.round(jaccard * 70 + (1 - levenshtein(a, b) / Math.max(a.length, b.length)) * 20 + prefixBonus);
}
function matchCertificate(ocr, records) {
  if (!ocr.name) return { record: null, score: 0 };
  let best = null;
  let bestScore = 0;
  for (const record of records) {
    let score = scoreNames(ocr.name, record.name);
    if (ocr.dateKey && record.dateKey && ocr.dateKey === record.dateKey) score += 8;
    if (ocr.course && record.title) {
      const course = normalizeName(ocr.course);
      const title = normalizeName(record.title);
      if (course.includes("INSPECTOR") && title.includes("INSPECTOR")) score += 4;
      if (course.includes("ERECTOR") && title.includes("ERECTOR")) score += 4;
      if (course.includes("SUPERVISOR") && title.includes("SUPERVISOR")) score += 4;
    }
    if (score > bestScore) { bestScore = score; best = record; }
  }
  if (bestScore < 68) return { record: null, score: bestScore };
  return { record: best, score: Math.min(bestScore, 100) };
}

function parseFilenameDate(filename) {
  const match = filename.match(/(\d{2})(\d{2})(\d{2})(?!\d)/);
  if (!match) return "";
  const month = Number(match[1]);
  const day = Number(match[2]);
  const year = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return "";
  return dateKey(new Date(year < 50 ? 2000 + year : 1900 + year, month - 1, day));
}
function parseCertText(text) {
  const lines = String(text || "").split(/\n/).map((line) => line.replace(/\s+/g, " ").trim()).filter(Boolean);
  let name = "";
  const certifyIdx = lines.findIndex((line) => /certify that/i.test(line));
  if (certifyIdx !== -1) {
    for (let i = certifyIdx + 1; i < Math.min(certifyIdx + 5, lines.length); i += 1) {
      const cleaned = lines[i].replace(/[^A-Za-z .'-]/g, " ").replace(/\s+/g, " ").trim();
      if (cleaned.length >= 5 && !/certificate|completed|course|scaffold|frame|tube|coupler|inspector|erector|supervisor|training|institute/i.test(cleaned)) {
        name = cleaned.toUpperCase();
        break;
      }
    }
  }
  const certMatch = text.match(/#\s*STI\s*(\d{4,})/i) || text.match(/STI\s*#?\s*(\d{4,})/i);
  const courseMatch = text.match(/Scaffold[^\n"]{0,40}(Inspector|Erector|Supervisor)/i);
  const dateMatch = text.match(/(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})/i);
  const parsedDate = dateMatch ? new Date(Number(dateMatch[3]), MONTHS[dateMatch[1].toLowerCase()], Number(dateMatch[2])) : null;
  return { name: name.trim(), certNo: certMatch ? `STI${certMatch[1]}` : "", course: courseMatch ? courseMatch[0].replace(/\s+/g, " ").trim() : "", dateKey: dateKey(parsedDate), text };
}
function displayCompany(name) {
  const text = String(name || "").trim();
  if (!text || /^(n\/a|na|nil|none|-)$/i.test(text)) return "_No Company";
  return text;
}
function safeName(name) {
  return String(name || "Unknown").replace(/[<>:"/\\|?*]/g, " ").replace(/\s+/g, " ").replace(/[. ]+$/g, "").trim().slice(0, 120) || "Unknown";
}

function rotateCanvas(source, deg) {
  const canvas = document.createElement("canvas");
  const rad = (deg * Math.PI) / 180;
  if (deg % 180 === 0) {
    canvas.width = source.width;
    canvas.height = source.height;
  } else {
    canvas.width = source.height;
    canvas.height = source.width;
  }
  const ctx = canvas.getContext("2d");
  ctx.translate(canvas.width / 2, canvas.height / 2);
  ctx.rotate(rad);
  ctx.drawImage(source, -source.width / 2, -source.height / 2);
  return canvas;
}
function cropCenter(source) {
  const canvas = document.createElement("canvas");
  const left = Math.floor(source.width * 0.18);
  const top = Math.floor(source.height * 0.34);
  canvas.width = Math.floor(source.width * 0.64);
  canvas.height = Math.floor(source.height * 0.34);
  canvas.getContext("2d").drawImage(source, left, top, canvas.width, canvas.height, 0, 0, canvas.width, canvas.height);
  return canvas;
}
async function ocrCanvas(worker, canvas) {
  const result = await worker.recognize(canvas);
  return parseCertText(result.data.text || "");
}
async function extractFromCanvas(worker, canvas) {
  let best = { name: "", certNo: "", course: "", dateKey: "", text: "" };
  for (const deg of [90, 0, 270, 180]) {
    const rotated = deg === 0 ? canvas : rotateCanvas(canvas, deg);
    const parsed = await ocrCanvas(worker, cropCenter(rotated));
    if (parsed.name && /certify that/i.test(parsed.text)) return parsed;
    if (parsed.name.length > best.name.length) best = parsed;
  }
  return best;
}

async function renderPdfPage(pdf, pageNumber) {
  const page = await pdf.getPage(pageNumber);
  const viewport = page.getViewport({ scale: 2 });
  const canvas = document.createElement("canvas");
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
  return canvas;
}

async function extractSignedPage(sourceDoc, pageIndex, signaturePng) {
  const srcPage = sourceDoc.getPage(pageIndex);
  const { width, height } = srcPage.getSize();
  const output = await PDFDocument.create();
  const embedded = await output.embedPage(srcPage);
  const page = height > width ? output.addPage([height, width]) : output.addPage([width, height]);
  if (height > width) page.drawPage(embedded, { x: 0, y: width, rotate: degrees(-90) });
  else page.drawPage(embedded, { x: 0, y: 0, width, height });
  if (signaturePng) {
    const image = await output.embedPng(signaturePng);
    const sigH = SIGN_WIDTH * (image.height / image.width);
    page.drawImage(image, { x: SIGN_X, y: SIGN_Y, width: SIGN_WIDTH, height: sigH });
  }
  return output.save();
}

async function loadTransparentSignature() {
  const img = new Image();
  img.crossOrigin = "anonymous";
  img.src = "./sti-sign.png";
  await img.decode();
  const canvas = document.createElement("canvas");
  canvas.width = img.width;
  canvas.height = img.height;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(img, 0, 0);
  const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height);
  for (let i = 0; i < pixels.data.length; i += 4) {
    if (pixels.data[i] > 240 && pixels.data[i + 1] > 240 && pixels.data[i + 2] > 240) pixels.data[i + 3] = 0;
  }
  ctx.putImageData(pixels, 0, 0);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  return new Uint8Array(await blob.arrayBuffer());
}

const excelInput = document.getElementById("excelInput");
const pdfInput = document.getElementById("pdfInput");
const runBtn = document.getElementById("runBtn");
const downloadBtn = document.getElementById("downloadBtn");
const progress = document.getElementById("progress");
const status = document.getElementById("status");
const bar = document.getElementById("bar");
const stats = document.getElementById("stats");
const errorBox = document.getElementById("error");
const tableWrap = document.getElementById("tableWrap");
let zipBlob = null;

function bindDrop(el, input) {
  el.addEventListener("click", () => input.click());
  el.addEventListener("dragover", (e) => { e.preventDefault(); el.classList.add("drag"); });
  el.addEventListener("dragleave", () => el.classList.remove("drag"));
  el.addEventListener("drop", (e) => {
    e.preventDefault();
    el.classList.remove("drag");
    input.files = e.dataTransfer.files;
    input.dispatchEvent(new Event("change"));
  });
}
bindDrop(document.getElementById("excelDrop"), excelInput);
bindDrop(document.getElementById("pdfDrop"), pdfInput);
excelInput.addEventListener("change", () => {
  document.getElementById("excelList").textContent = excelInput.files[0]?.name || "";
});
pdfInput.addEventListener("change", () => {
  document.getElementById("pdfList").textContent = [...pdfInput.files].map((f) => f.name).join("\n");
});

runBtn.addEventListener("click", async () => {
  errorBox.textContent = "";
  tableWrap.innerHTML = "";
  zipBlob = null;
  downloadBtn.disabled = true;
  if (!excelInput.files[0] || !pdfInput.files.length) {
    progress.style.display = "block";
    errorBox.textContent = "Please choose the Excel file and at least one PDF.";
    return;
  }
  runBtn.disabled = true;
  progress.style.display = "block";
  status.textContent = "Reading Excel...";
  bar.style.width = "4%";
  try {
    const records = parseExcel(await excelInput.files[0].arrayBuffer());
    const signaturePng = await loadTransparentSignature();
    const worker = await createWorker("eng");
    await worker.setParameters({ tessedit_pageseg_mode: "6" });
    const zip = new JSZip();
    const report = [];
    let matched = 0;
    let unmatched = 0;
    const used = new Map();
    const pdfFiles = [...pdfInput.files];
    let pageTotal = 0;
    const loaded = [];
    for (const file of pdfFiles) {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const pdf = await pdfjs.getDocument({ data: bytes }).promise;
      const sourceDoc = await PDFDocument.load(bytes);
      pageTotal += pdf.numPages;
      loaded.push({ file, pdf, sourceDoc, dateKey: parseFilenameDate(file.name) });
    }
    let pageDone = 0;
    for (const item of loaded) {
      for (let pageIndex = 0; pageIndex < item.pdf.numPages; pageIndex += 1) {
        pageDone += 1;
        status.textContent = `Reading ${item.file.name} page ${pageIndex + 1} of ${item.pdf.numPages}`;
        bar.style.width = `${Math.max(6, Math.round((pageDone / pageTotal) * 100))}%`;
        const canvas = await renderPdfPage(item.pdf, pageIndex + 1);
        const ocr = await extractFromCanvas(worker, canvas);
        if (item.dateKey && !ocr.dateKey) ocr.dateKey = item.dateKey;
        const { record, score } = matchCertificate(ocr, records);
        const pdfBytes = await extractSignedPage(item.sourceDoc, pageIndex, signaturePng);
        let company = "";
        let matchedName = "";
        let statusLabel = "unmatched";
        let outPath = "";
        if (record) {
          statusLabel = "matched";
          matched += 1;
          company = displayCompany(record.company);
          matchedName = record.name;
          const folder = `by-company/${safeName(company)}`;
          const certPart = ocr.certNo ? `_${ocr.certNo}` : "";
          let filename = `${safeName(record.name || ocr.name)}${certPart}.pdf`;
          const key = `${folder}/${filename}`.toLowerCase();
          const count = used.get(key) || 0;
          used.set(key, count + 1);
          if (count > 0) filename = `${safeName(record.name)}${certPart}_${count + 1}.pdf`;
          outPath = `${folder}/${filename}`;
          zip.file(outPath, pdfBytes);
        } else {
          unmatched += 1;
          outPath = `_Unmatched/${safeName(ocr.name || `${item.file.name}_p${pageIndex + 1}`)}.pdf`;
          zip.file(outPath, pdfBytes);
        }
        report.push({ sourcePdf: item.file.name, page: pageIndex + 1, ocrName: ocr.name, matchedName, company, score, status: statusLabel });
      }
    }
    await worker.terminate();
    const csv = ["sourcePdf,page,ocrName,matchedName,company,score,status", ...report.map((row) => [row.sourcePdf, row.page, row.ocrName, row.matchedName, row.company, row.score, row.status].map((v) => `"${String(v ?? "").replace(/"/g, '""')}"`).join(","))].join("\n");
    zip.file("matching-report.csv", csv);
    zipBlob = await zip.generateAsync({ type: "blob" });
    status.textContent = `Done. ${matched} matched, ${unmatched} unmatched.`;
    bar.style.width = "100%";
    const companies = new Set(report.filter((r) => r.status === "matched").map((r) => r.company));
    stats.innerHTML = `<span><b>${pageTotal}</b> certificates</span><span><b>${matched}</b> matched</span><span><b>${unmatched}</b> unmatched</span><span><b>${companies.size}</b> company folders</span>`;
    tableWrap.innerHTML = `<table><thead><tr><th>PDF</th><th>Page</th><th>Read name</th><th>Matched name</th><th>Company</th><th>Score</th><th>Status</th></tr></thead><tbody>${report.map((row) => `<tr><td>${row.sourcePdf}</td><td>${row.page}</td><td>${row.ocrName || ""}</td><td>${row.matchedName || ""}</td><td>${row.company || ""}</td><td>${row.score || 0}</td><td class="${row.status === "matched" ? "ok" : "bad"}">${row.status}</td></tr>`).join("")}</tbody></table>`;
    downloadBtn.disabled = false;
  } catch (error) {
    errorBox.textContent = error.message || String(error);
  } finally {
    runBtn.disabled = false;
  }
});

downloadBtn.addEventListener("click", () => {
  if (!zipBlob) return;
  const url = URL.createObjectURL(zipBlob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "company-certificates.zip";
  a.click();
  URL.revokeObjectURL(url);
});
