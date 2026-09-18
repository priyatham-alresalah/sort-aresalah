import fs from "fs";
import path from "path";
import { PDFDocument } from "pdf-lib";
import { parseExcel } from "./excel.js";
import { matchCertificate } from "./match.js";
import { createOcrWorker, extractFromImage } from "./ocr.js";
import {
  parseFilenameDate,
  openPdf,
  renderPagePngFromDoc,
  extractPagePdf,
  loadInstructorSignature,
  safeFolderName,
  safeFileName,
} from "./pdf.js";

function displayCompany(name) {
  const text = String(name || "").trim();
  if (!text || /^(n\/a|na|nil|none|-)$/i.test(text)) return "_No Company";
  return text;
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

export async function processJob({ excelBuffer, pdfs, outputDir, onProgress }) {
  const emit = onProgress || (() => {});
  emit({ stage: "excel", message: "Reading Excel..." });
  const records = parseExcel(excelBuffer);
  if (!records.length) throw new Error("No candidate rows found in the Excel file.");

  const companies = new Set(records.map((r) => r.company).filter(Boolean));
  emit({
    stage: "excel",
    message: `Excel loaded: ${records.length} candidates, ${companies.size} companies.`,
    records: records.length,
    companies: companies.size,
  });

  const byCompanyDir = path.join(outputDir, "by-company");
  const unmatchedDir = path.join(outputDir, "_Unmatched");
  ensureDir(byCompanyDir);
  ensureDir(unmatchedDir);

  const usedNames = new Map();
  const report = [];
  let matched = 0;
  let unmatched = 0;
  let pageTotal = 0;

  const pdfInfos = [];
  for (const pdf of pdfs) {
    const doc = openPdf(pdf.buffer);
    const pages = doc.countPages();
    pageTotal += pages;
    pdfInfos.push({
      ...pdf,
      pages,
      mupdfDoc: doc,
      ...parseFilenameDate(pdf.originalName),
    });
  }

  emit({ stage: "pdfs", message: `Found ${pageTotal} certificate pages.`, pageTotal });

  const signaturePng = await loadInstructorSignature();
  if (!signaturePng) {
    emit({
      stage: "pdfs",
      message: "Instructor signature not found. Certificates will be saved without it.",
      pageTotal,
    });
  }

  const worker = await createOcrWorker();
  let pageDone = 0;

  try {
    for (const pdf of pdfInfos) {
      const sourceDoc = await PDFDocument.load(pdf.buffer);
      for (let pageIndex = 0; pageIndex < pdf.pages; pageIndex += 1) {
        pageDone += 1;
        emit({
          stage: "ocr",
          message: `Reading ${pdf.originalName} page ${pageIndex + 1} of ${pdf.pages}`,
          pageDone,
          pageTotal,
        });

        const image = renderPagePngFromDoc(pdf.mupdfDoc, pageIndex, 2);
        const ocr = await extractFromImage(worker, image);
        if (pdf.dateKey && !ocr.dateKey) ocr.dateKey = pdf.dateKey;

        const { record, score } = matchCertificate(ocr, records);
        const pagePdf = await extractPagePdf(sourceDoc, pageIndex, signaturePng);

        let status = "unmatched";
        let company = "";
        let matchedName = "";
        let outPath = "";

        if (record) {
          status = "matched";
          matched += 1;
          company = displayCompany(record.company);
          matchedName = record.name;
          const folder = path.join(byCompanyDir, safeFolderName(company));
          ensureDir(folder);
          const base = safeFileName(record.name || ocr.name || `page-${pageIndex + 1}`);
          const certPart = ocr.certNo ? `_${ocr.certNo}` : "";
          let filename = `${base}${certPart}.pdf`;
          const key = `${folder}/${filename}`.toLowerCase();
          const count = usedNames.get(key) || 0;
          usedNames.set(key, count + 1);
          if (count > 0) {
            filename = `${base}${certPart}_${count + 1}.pdf`;
          }
          outPath = path.join(folder, filename);
          fs.writeFileSync(outPath, pagePdf);
        } else {
          unmatched += 1;
          const fallback = safeFileName(
            ocr.name || `${path.parse(pdf.originalName).name}_p${pageIndex + 1}`
          );
          outPath = path.join(unmatchedDir, `${fallback}.pdf`);
          fs.writeFileSync(outPath, pagePdf);
        }

        report.push({
          sourcePdf: pdf.originalName,
          page: pageIndex + 1,
          ocrName: ocr.name,
          certNo: ocr.certNo,
          course: ocr.course,
          date: ocr.dateRaw || pdf.dateKey,
          matchedName,
          company,
          score,
          status,
          output: path.relative(outputDir, outPath),
        });
      }
      pdf.mupdfDoc.destroy();
    }
  } finally {
    await worker.terminate();
  }

  const reportPath = path.join(outputDir, "matching-report.csv");
  const headers = [
    "sourcePdf",
    "page",
    "ocrName",
    "certNo",
    "course",
    "date",
    "matchedName",
    "company",
    "score",
    "status",
    "output",
  ];
  const csv = [
    headers.join(","),
    ...report.map((row) => headers.map((key) => csvEscape(row[key])).join(",")),
  ].join("\n");
  fs.writeFileSync(reportPath, csv);

  const companyCount = fs
    .readdirSync(byCompanyDir, { withFileTypes: true })
    .filter((entry) => entry.isDirectory()).length;

  emit({
    stage: "done",
    message: `Done. ${matched} matched, ${unmatched} unmatched, ${companyCount} company folders.`,
    matched,
    unmatched,
    companyCount,
    pageTotal,
    report,
  });

  return {
    matched,
    unmatched,
    companyCount,
    pageTotal,
    records: records.length,
    report,
    byCompanyDir,
    unmatchedDir,
    reportPath,
  };
}
