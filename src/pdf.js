import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import sharp from "sharp";
import { PDFDocument, degrees } from "pdf-lib";
import * as mupdf from "mupdf";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const SIGN_X = 426;
const SIGN_Y = 129;
const SIGN_WIDTH = 154;

export function parseFilenameDate(filename) {
  const base = path.basename(filename);
  const match = base.match(/(\d{2})(\d{2})(\d{2})(?!\d)/);
  if (!match) return { date: null, dateKey: "" };
  const month = Number(match[1]);
  const day = Number(match[2]);
  const year = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) {
    return { date: null, dateKey: "" };
  }
  const fullYear = year < 50 ? 2000 + year : 1900 + year;
  const date = new Date(fullYear, month - 1, day);
  const dateKey = `${fullYear}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  return { date, dateKey };
}

export function openPdf(fileBytes) {
  return mupdf.Document.openDocument(fileBytes, "application/pdf");
}

export function renderPagePngFromDoc(doc, pageIndex, scale = 2) {
  const page = doc.loadPage(pageIndex);
  const pixmap = page.toPixmap(
    mupdf.Matrix.scale(scale, scale),
    mupdf.ColorSpace.DeviceRGB,
    false,
    true
  );
  return Buffer.from(pixmap.asPNG());
}

export function renderPagePng(fileBytes, pageIndex, scale = 2) {
  const doc = openPdf(fileBytes);
  try {
    return renderPagePngFromDoc(doc, pageIndex, scale);
  } finally {
    doc.destroy();
  }
}

export function countPages(fileBytes) {
  const doc = openPdf(fileBytes);
  try {
    return doc.countPages();
  } finally {
    doc.destroy();
  }
}

export function rotatePageToLandscape(page) {
  const { width, height } = page.getSize();
  const current = ((page.getRotation()?.angle || 0) % 360 + 360) % 360;
  if (width > height && current === 0) return;
  if (current === 90) return;
  page.setRotation(degrees(90));
}

export async function rotatePdfBufferToLandscape(fileBytes) {
  const doc = await PDFDocument.load(fileBytes);
  for (const page of doc.getPages()) rotatePageToLandscape(page);
  return Buffer.from(await doc.save());
}

export async function loadInstructorSignature() {
  const file = fs
    .readdirSync(ROOT)
    .find((name) => /^sti\s*sign\.png$/i.test(name));
  if (!file) return null;
  const { data, info } = await sharp(path.join(ROOT, file))
    .ensureAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });
  for (let i = 0; i < data.length; i += 4) {
    if (data[i] > 240 && data[i + 1] > 240 && data[i + 2] > 240) data[i + 3] = 0;
  }
  return sharp(data, {
    raw: { width: info.width, height: info.height, channels: 4 },
  })
    .png()
    .toBuffer();
}

async function drawInstructorSignature(pdfDoc, page, signaturePng) {
  if (!signaturePng) return;
  const image = await pdfDoc.embedPng(signaturePng);
  const width = SIGN_WIDTH;
  const height = width * (image.height / image.width);
  page.drawImage(image, { x: SIGN_X, y: SIGN_Y, width, height });
}

export async function extractPagePdf(sourceDoc, pageIndex, signaturePng) {
  const srcPage = sourceDoc.getPage(pageIndex);
  const { width, height } = srcPage.getSize();
  const output = await PDFDocument.create();
  const embedded = await output.embedPage(srcPage);
  const page =
    height > width ? output.addPage([height, width]) : output.addPage([width, height]);

  if (height > width) {
    page.drawPage(embedded, {
      x: 0,
      y: width,
      rotate: degrees(-90),
    });
  } else {
    page.drawPage(embedded, { x: 0, y: 0, width, height });
  }

  await drawInstructorSignature(output, page, signaturePng);
  return Buffer.from(await output.save());
}

export function safeFolderName(name) {
  const cleaned = String(name || "Unknown")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, " ")
    .replace(/\s+/g, " ")
    .replace(/[. ]+$/g, "")
    .trim();
  return cleaned || "Unknown";
}

export function safeFileName(name) {
  return safeFolderName(name).slice(0, 120);
}
