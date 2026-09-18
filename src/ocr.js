import sharp from "sharp";
import { createWorker } from "tesseract.js";
import { dateKey } from "./excel.js";

const MONTHS = {
  january: 0,
  february: 1,
  march: 2,
  april: 3,
  may: 4,
  june: 5,
  july: 6,
  august: 7,
  september: 8,
  october: 9,
  november: 10,
  december: 11,
};

export function parseCertText(text) {
  const lines = String(text || "")
    .split(/\n/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);

  let name = "";
  const certifyIdx = lines.findIndex((line) => /certify that/i.test(line));
  if (certifyIdx !== -1) {
    for (let i = certifyIdx + 1; i < Math.min(certifyIdx + 5, lines.length); i += 1) {
      const cleaned = lines[i]
        .replace(/[^A-Za-z .'-]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
      if (
        cleaned.length >= 5 &&
        !/certificate|completed|course|scaffold|frame|tube|coupler|inspector|erector|supervisor|training|institute/i.test(
          cleaned
        )
      ) {
        name = cleaned.toUpperCase();
        break;
      }
    }
  }

  if (!name) {
    const caps = lines
      .map((line) => line.replace(/[^A-Za-z .'-]/g, " ").replace(/\s+/g, " ").trim())
      .filter(
        (line) =>
          /^[A-Z][A-Z .'-]{4,}$/.test(line) &&
          !/SCAFFOLD|TRAINING|INSTITUTE|CERTIFICATE|INSTRUCTOR|RETRAINING/.test(line)
      )
      .sort((a, b) => b.length - a.length);
    name = caps[0] || "";
  }

  const certMatch =
    text.match(/#\s*STI\s*(\d{4,})/i) ||
    text.match(/STI\s*#?\s*(\d{4,})/i) ||
    text.match(/Certificate[^\d]{0,20}(\d{5,})/i);
  const certNo = certMatch ? `STI${certMatch[1]}` : "";

  const courseMatch = text.match(
    /Scaffold[^\n"]{0,40}(Inspector|Erector|Supervisor)/i
  );
  const course = courseMatch ? courseMatch[0].replace(/\s+/g, " ").trim() : "";

  const dateMatch = text.match(
    /(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})/i
  );
  let parsedDate = null;
  if (dateMatch) {
    parsedDate = new Date(
      Number(dateMatch[3]),
      MONTHS[dateMatch[1].toLowerCase()],
      Number(dateMatch[2])
    );
  }

  return {
    name: name.trim(),
    certNo,
    course,
    dateRaw: dateMatch ? dateMatch[0] : "",
    dateKey: dateKey(parsedDate),
    text,
  };
}

async function cropCenter(imageBuffer) {
  const image = sharp(imageBuffer);
  const meta = await image.metadata();
  const left = Math.floor(meta.width * 0.18);
  const top = Math.floor(meta.height * 0.34);
  const width = Math.floor(meta.width * 0.64);
  const height = Math.floor(meta.height * 0.34);
  return image
    .extract({ left, top, width, height })
    .grayscale()
    .normalize()
    .png()
    .toBuffer();
}

async function recognizeBuffer(worker, imageBuffer) {
  const cropped = await cropCenter(imageBuffer);
  const result = await worker.recognize(cropped);
  return parseCertText(result.data.text || "");
}

export async function extractFromImage(worker, imageBuffer) {
  const rotations = [90, 0, 270, 180];
  let best = { name: "", certNo: "", course: "", dateRaw: "", dateKey: "", text: "" };

  for (const deg of rotations) {
    const rotated =
      deg === 0
        ? imageBuffer
        : await sharp(imageBuffer).rotate(deg).png().toBuffer();
    const parsed = await recognizeBuffer(worker, rotated);
    if (parsed.name && /certify that/i.test(parsed.text)) return parsed;
    if (parsed.name.length > best.name.length) best = parsed;
  }

  return best;
}

export async function createOcrWorker() {
  const worker = await createWorker("eng");
  await worker.setParameters({
    tessedit_pageseg_mode: "6",
  });
  return worker;
}
