import XLSX from "xlsx";

const NAME_HEADERS = ["deligate name", "delegate name", "candidate name", "name"];
const COMPANY_HEADERS = ["company name", "company", "employer"];
const DATE_HEADERS = ["training date", "date", "course date"];
const TITLE_HEADERS = ["traninig title", "training title", "course", "title"];
const CERT_HEADERS = ["certificate number", "cert number", "certificate no"];

const MONTHS = {
  jan: 0,
  january: 0,
  feb: 1,
  february: 1,
  mar: 2,
  march: 2,
  apr: 3,
  april: 3,
  may: 4,
  jun: 5,
  june: 5,
  jul: 6,
  july: 6,
  aug: 7,
  august: 7,
  sep: 8,
  sept: 8,
  september: 8,
  oct: 9,
  october: 9,
  nov: 10,
  november: 10,
  dec: 11,
  december: 11,
};

function normHeader(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[:_]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function findColumn(headers, aliases) {
  for (const alias of aliases) {
    const idx = headers.findIndex((h) => h === alias || h.includes(alias));
    if (idx !== -1) return idx;
  }
  return -1;
}

export function parseTrainingEndDate(raw) {
  const text = String(raw || "")
    .replace(/,/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return null;

  const compact = text.replace(/(\d{1,2})(?:st|nd|rd|th)/gi, "$1");
  const ranged = compact.match(
    /(\d{1,2})\s*[-–]\s*(\d{1,2})\s*([A-Za-z]+)\s*(\d{4})/i
  );
  if (ranged) {
    const month = MONTHS[ranged[3].toLowerCase()];
    if (month != null) return new Date(Number(ranged[4]), month, Number(ranged[2]));
  }

  const crossMonth = compact.match(
    /(\d{1,2})\s*([A-Za-z]+)\s*[-–]\s*(\d{1,2})\s*([A-Za-z]+)\s*(\d{4})/i
  );
  if (crossMonth) {
    const month = MONTHS[crossMonth[4].toLowerCase()];
    if (month != null) {
      return new Date(Number(crossMonth[5]), month, Number(crossMonth[3]));
    }
  }

  const single = compact.match(/(\d{1,2})\s*([A-Za-z]+)\s*(\d{4})/i);
  if (single) {
    const month = MONTHS[single[2].toLowerCase()];
    if (month != null) return new Date(Number(single[3]), month, Number(single[1]));
  }

  return null;
}

export function dateKey(date) {
  if (!date || Number.isNaN(date.getTime())) return "";
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

export function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: false,
  });

  if (!rows.length) return [];

  const headers = rows[0].map(normHeader);
  const nameIdx = findColumn(headers, NAME_HEADERS);
  const companyIdx = findColumn(headers, COMPANY_HEADERS);
  const dateIdx = findColumn(headers, DATE_HEADERS);
  const titleIdx = findColumn(headers, TITLE_HEADERS);
  const certIdx = findColumn(headers, CERT_HEADERS);

  if (nameIdx === -1 || companyIdx === -1) {
    throw new Error(
      "Excel must have 'Company Name' and 'Delegate Name' columns."
    );
  }

  const records = [];
  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const name = String(row[nameIdx] || "").trim();
    const company = String(row[companyIdx] || "").trim();
    if (!name && !company) continue;

    const dateRaw = dateIdx === -1 ? "" : String(row[dateIdx] || "").trim();
    const endDate = parseTrainingEndDate(dateRaw);
    records.push({
      rowNumber: i + 1,
      name,
      company,
      dateRaw,
      dateKey: dateKey(endDate),
      title: titleIdx === -1 ? "" : String(row[titleIdx] || "").trim(),
      certNo: certIdx === -1 ? "" : String(row[certIdx] || "").trim(),
    });
  }

  return records;
}
