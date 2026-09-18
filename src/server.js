import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import express from "express";
import multer from "multer";
import { ZipArchive } from "archiver";
import { processJob } from "./process.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const PUBLIC_DIR = path.join(ROOT, "public");
const JOBS_DIR = path.join(ROOT, "jobs");

fs.mkdirSync(JOBS_DIR, { recursive: true });

const jobs = new Map();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024, files: 80 },
});

const app = express();
app.use(express.static(PUBLIC_DIR));

function send(res, event, data) {
  res.write(`event: ${event}\n`);
  res.write(`data: ${JSON.stringify(data)}\n\n`);
}

app.post(
  "/api/jobs",
  upload.fields([
    { name: "excel", maxCount: 1 },
    { name: "pdfs", maxCount: 60 },
  ]),
  async (req, res) => {
    try {
      const excelFile = (req.files?.excel || [])[0];
      const pdfFiles = req.files?.pdfs || [];
      if (!excelFile) {
        return res.status(400).json({ error: "Please upload the Excel file." });
      }
      if (!pdfFiles.length) {
        return res.status(400).json({ error: "Please upload at least one PDF." });
      }

      const jobId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      const outputDir = path.join(JOBS_DIR, jobId);
      fs.mkdirSync(outputDir, { recursive: true });

      const job = {
        id: jobId,
        status: "queued",
        events: [],
        clients: new Set(),
        outputDir,
        result: null,
        error: null,
      };
      jobs.set(jobId, job);

      const emit = (payload) => {
        job.events.push(payload);
        job.status = payload.stage === "done" ? "done" : payload.stage;
        for (const client of job.clients) send(client, "progress", payload);
      };

      res.json({ jobId });

      setImmediate(async () => {
        try {
          const result = await processJob({
            excelBuffer: excelFile.buffer,
            pdfs: pdfFiles.map((file) => ({
              buffer: file.buffer,
              originalName: file.originalname,
            })),
            outputDir,
            onProgress: emit,
          });
          job.result = result;
          job.status = "done";
        } catch (error) {
          job.status = "error";
          job.error = error.message;
          emit({ stage: "error", message: error.message });
        }
      });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  }
);

app.get("/api/jobs/:id/events", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).end();

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders();

  for (const event of job.events) send(res, "progress", event);
  if (job.status === "done" && job.result) {
    send(res, "progress", {
      stage: "done",
      message: job.events.at(-1)?.message || "Done.",
      matched: job.result.matched,
      unmatched: job.result.unmatched,
      companyCount: job.result.companyCount,
      pageTotal: job.result.pageTotal,
      report: job.result.report,
    });
  }
  if (job.status === "error") {
    send(res, "progress", { stage: "error", message: job.error });
  }

  job.clients.add(res);
  req.on("close", () => job.clients.delete(res));
});

app.get("/api/jobs/:id/download", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job || job.status !== "done") {
    return res.status(404).json({ error: "Job is not ready." });
  }

  res.setHeader("Content-Type", "application/zip");
  res.setHeader(
    "Content-Disposition",
    'attachment; filename="company-certificates.zip"'
  );

  const archive = new ZipArchive({ zlib: { level: 9 } });
  archive.on("error", (error) => {
    if (!res.headersSent) res.status(500).json({ error: error.message });
  });
  archive.pipe(res);
  archive.directory(path.join(job.outputDir, "by-company"), "by-company");
  if (fs.existsSync(path.join(job.outputDir, "_Unmatched"))) {
    archive.directory(path.join(job.outputDir, "_Unmatched"), "_Unmatched");
  }
  if (fs.existsSync(path.join(job.outputDir, "matching-report.csv"))) {
    archive.file(path.join(job.outputDir, "matching-report.csv"), {
      name: "matching-report.csv",
    });
  }
  archive.finalize();
});

const port = Number(process.env.PORT) || 3000;
const server = app.listen(port, "127.0.0.1", () => {
  console.log(`STI Certificate Sorter running at http://127.0.0.1:${port}`);
});
server.on("error", (error) => {
  console.error("Server error:", error);
  process.exit(1);
});
