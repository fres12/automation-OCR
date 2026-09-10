require("dotenv").config();
const { GoogleGenerativeAI } = require("@google/generative-ai");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// ─── Konfigurasi ───────────────────────────────────────────────────────────────
const GEMINI_API_KEY = (process.env.GEMINI_API_KEY || "").trim();
const IMAGE_FOLDER = process.env.IMAGE_FOLDER || "C:\\Users\\fresn\\OneDrive\\Documents\\Bali Nusra";
const IMAGE_EXTENSIONS = (process.env.IMAGE_EXTENSIONS || ".jpg,.jpeg,.png,.webp,.gif,.bmp")
  .split(",")
  .map((e) => e.trim().toLowerCase());

if (!GEMINI_API_KEY) {
  console.error("ERROR: GEMINI_API_KEY belum diisi di file .env");
  process.exit(1);
}

// ─── Inisialisasi Gemini AI ─────────────────────────────────────────────────
const MODEL_NAME = process.env.GEMINI_MODEL || "gemini-3.5-flash-lite";
const genAI = new GoogleGenerativeAI(GEMINI_API_KEY);
const model = genAI.getGenerativeModel({ model: MODEL_NAME });

// Delay antar request (ms) — sesuaikan jika masih rate-limit
const DELAY_MS = parseInt(process.env.DELAY_MS || "3000", 10);
// Timeout request Gemini API (ms)
const TIMEOUT_MS = parseInt(process.env.TIMEOUT_MS || "60000", 10);
// Jumlah maksimum retry saat kena rate limit
const MAX_RETRIES = 3;

// ─── Whitelist: muat ID dari whitelist.csv ─────────────────────────────────
const WHITELIST_PATH = process.env.WHITELIST_PATH || path.join(__dirname, "whitelist.csv");
function loadWhitelist() {
  if (!fs.existsSync(WHITELIST_PATH)) {
    console.warn(`[WARN] whitelist.csv tidak ditemukan di: ${WHITELIST_PATH}`);
    console.warn("[WARN] Semua file akan diproses (tanpa filter whitelist).");
    return null; // null = proses semua
  }
  const raw = fs.readFileSync(WHITELIST_PATH, "utf8");
  const ids = raw
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  return new Set(ids);
}
const WHITELIST = loadWhitelist();

// ─── Helper: Ambil ID sebelum underscore pertama ────────────────────────────
function extractIdFromFilename(filename) {
  const basename = path.basename(filename, path.extname(filename));
  const underscoreIndex = basename.indexOf("_");
  if (underscoreIndex === -1) return basename;
  return basename.substring(underscoreIndex + 1);
}

// ─── Helper: Konversi gambar ke base64 ─────────────────────────────────────
function imageToBase64(filePath) {
  const buffer = fs.readFileSync(filePath);
  return buffer.toString("base64");
}

// ─── Helper: Deteksi MIME type dari ekstensi ───────────────────────────────
function getMimeType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const mimeMap = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
  };
  return mimeMap[ext] || "image/jpeg";
}

// ─── Helper: Parse nilai angka dari string (misal "320", "79,5", "79.5") ────
function parseNumeric(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value).trim().replace(/[^0-9,.-]/g, "").replace(/,/g, ".");
  if (!cleaned) return 0;
  const parsed = parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

// ─── Helper: Sleep ──────────────────────────────────────────────────────────
function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ─── Fungsi: OCR satu gambar dengan retry ──────────────────────────────────
async function ocrImage(filePath) {
  const base64Data = imageToBase64(filePath);
  const mimeType = getMimeType(filePath);

  // Prompt khusus untuk ekstrak NILAI ANGKA kecepatan internet (bukan menghitung kata)
  const prompt = `Ini adalah gambar hasil speedtest jaringan internet.
Tugasmu: temukan dan ekstrak NILAI ANGKA kecepatan download (unduh) dan upload (unggah) dari gambar ini.

Petunjuk:
- Cari label "Unduh", "Download", atau ikon panah ke bawah → ambil angkanya
- Cari label "Unggah", "Upload", atau ikon panah ke atas → ambil angkanya
- Angka bisa dalam format seperti: 320, 79.5, 79,5 — abaikan satuan (Mbps, Kbps, dll)
- Jika tidak ditemukan, gunakan 0

Kembalikan HANYA JSON valid ini, tanpa penjelasan, tanpa markdown:
{"unduh": <angka_download>, "unggah": <angka_upload>}`;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const result = await Promise.race([
        model.generateContent({
          contents: [
            {
              role: "user",
              parts: [
                { inlineData: { mimeType: mimeType, data: base64Data } },
                { text: prompt },
              ],
            },
          ],
          generationConfig: {
            responseMimeType: "application/json",
          },
        }),
        new Promise((_, reject) =>
          setTimeout(() => reject(new Error(`Timeout ${TIMEOUT_MS / 1000}s ke Gemini API`)), TIMEOUT_MS)
        ),
      ]);

      const responseText = result.response.text().trim();
      const cleaned = responseText.replace(/```json\s*/gi, "").replace(/```\s*/g, "").trim();

      let parsed;
      try {
        parsed = JSON.parse(cleaned);
      } catch (e) {
        console.warn(`\n  [WARN] Gagal parse JSON: "${cleaned}" → gunakan 0`);
        parsed = { unduh: 0, unggah: 0 };
      }

      return {
        unduh: parseNumeric(parsed.unduh),
        unggah: parseNumeric(parsed.unggah),
      };
    } catch (err) {
      const isRateLimit = err.message && (
        err.message.includes("429") ||
        err.message.includes("quota") ||
        err.message.includes("RESOURCE_EXHAUSTED")
      );

      if (isRateLimit && attempt < MAX_RETRIES) {
        const waitSec = 60 * attempt; // tunggu 60s, 120s, dst
        process.stdout.write(`\n  [Rate Limit] Tunggu ${waitSec}s lalu retry (${attempt}/${MAX_RETRIES})... `);
        await sleep(waitSec * 1000);
        continue;
      }
      throw err;
    }
  }
}

// ─── Fungsi utama: Proses semua gambar di folder ───────────────────────────
async function processFolder() {
  console.log("OCR Gemini AI - Ekstraksi nilai unduh & unggah (Mbps)");
  console.log("------------------------------------------------------------");
  console.log(`Folder  : ${IMAGE_FOLDER}`);
  console.log(`Model   : ${MODEL_NAME}`);
  console.log(`API Key : ${GEMINI_API_KEY.substring(0, 8)}...`);
  console.log(`Delay   : ${DELAY_MS}ms antar request`);
  console.log(`Timeout : ${TIMEOUT_MS / 1000}s per request`);
  console.log("------------------------------------------------------------");

  if (!fs.existsSync(IMAGE_FOLDER)) {
    console.error(`ERROR: Folder tidak ditemukan: ${IMAGE_FOLDER}`);
    process.exit(1);
  }

  const allFiles = fs.readdirSync(IMAGE_FOLDER);
  const imageFiles = allFiles.filter((file) =>
    IMAGE_EXTENSIONS.includes(path.extname(file).toLowerCase())
  );

  if (imageFiles.length === 0) {
    console.warn("WARN: Tidak ada file gambar ditemukan di folder tersebut.");
    process.exit(0);
  }

  if (WHITELIST) {
    console.log(`Whitelist : ${WHITELIST.size} Site ID dimuat dari whitelist.csv`);
  } else {
    console.log(`Whitelist : tidak ada — semua file diproses`);
  }
  console.log(`Ditemukan ${imageFiles.length} file gambar\n`);

  const excelProject = path.join(__dirname, "hasil_ocr.xlsx");
  const excelFolder = path.join(IMAGE_FOLDER, "hasil_ocr.xlsx");
  const jsonProject = path.join(__dirname, "hasil_ocr.json");

  console.log(`📁 File Excel akan otomatis disimpan ke:`);
  console.log(`   1. ${excelProject}`);
  console.log(`   2. ${excelFolder}\n`);

  const results = [];
  await saveToExcel(results, excelProject, excelFolder);

  let processedCount = 0;
  let skippedCount = 0;

  for (let i = 0; i < imageFiles.length; i++) {
    const filename = imageFiles[i];
    const filePath = path.join(IMAGE_FOLDER, filename);
    const fileId = extractIdFromFilename(filename);
    const inWhitelist = !WHITELIST || WHITELIST.has(fileId);

    process.stdout.write(`[${i + 1}/${imageFiles.length}] ${filename} (ID: ${fileId}) ... `);

    if (!inWhitelist) {
      // ID tidak ada di whitelist → skip Gemini, isi null
      console.log(`SKIP (tidak di whitelist)`);
      results.push({ id: fileId, filename, unduh: null, unggah: null, filePath, skipped: true });
      skippedCount++;
    } else {
      // ID ada di whitelist → proses dengan Gemini
      const startTime = Date.now();
      try {
        const { unduh, unggah } = await ocrImage(filePath);
        const duration = ((Date.now() - startTime) / 1000).toFixed(1);
        console.log(`OK | unduh=${unduh} Mbps, unggah=${unggah} Mbps (${duration}s)`);
        results.push({ id: fileId, filename, unduh, unggah, filePath });
      } catch (err) {
        const duration = ((Date.now() - startTime) / 1000).toFixed(1);
        console.log(`ERROR (${duration}s): ${err.message}`);
        results.push({ id: fileId, filename, unduh: null, unggah: null, filePath, error: err.message });
      }
      processedCount++;

      // Delay antar request agar tidak kena rate limit (hanya setelah request ke Gemini)
      const isLastWhitelistedFile = imageFiles.slice(i + 1).every((f) => {
        const id = extractIdFromFilename(f);
        return WHITELIST && !WHITELIST.has(id);
      });
      if (!isLastWhitelistedFile) {
        await sleep(DELAY_MS);
      }
    }

    await saveToExcel(results, excelProject, excelFolder);
    fs.writeFileSync(jsonProject, JSON.stringify(results, null, 2), "utf8");
  }

  console.log("\n============================================================");
  console.log("SELESAI! HASIL OCR TERSIMPAN");
  console.log("============================================================");
  console.log(`Excel : ${excelProject}`);
  console.log(`JSON  : ${jsonProject}`);

  const processed = results.filter((r) => !r.skipped && !r.error);
  const totalUnduh = processed.reduce((sum, r) => sum + (r.unduh || 0), 0);
  const totalUnggah = processed.reduce((sum, r) => sum + (r.unggah || 0), 0);
  const errors = results.filter((r) => r.error).length;
  const skipped = results.filter((r) => r.skipped).length;
  console.log("\nRingkasan:");
  console.log(`  Total semua file        : ${results.length}`);
  console.log(`  Diproses (whitelist)    : ${processed.length + errors}`);
  console.log(`  Skip (bukan whitelist)  : ${skipped}`);
  console.log(`  Total unduh diproses    : ${totalUnduh.toFixed(1)} Mbps`);
  console.log(`  Total unggah diproses   : ${totalUnggah.toFixed(1)} Mbps`);
  if (errors > 0) console.log(`  File error              : ${errors} file`);

  return results;
}

// ─── Fungsi: Simpan ke Excel (Autosave real-time) ──────────────────────────
async function saveToExcel(results, path1, path2) {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Hasil OCR");

    worksheet.columns = [
      { header: "Site ID", key: "id", width: 18 },
      { header: "Unduh (Mbps)", key: "unduh", width: 14 },
      { header: "Unggah (Mbps)", key: "unggah", width: 14 },
      { header: "Path File", key: "filePath", width: 75 },
    ];

    worksheet.getRow(1).font = { bold: true };

    results.forEach((r) => {
      worksheet.addRow({
        id: r.id,
        unduh: r.unduh ?? null,
        unggah: r.unggah ?? null,
        filePath: r.filePath,
      });
    });

    // Simpan ke path 1 (folder project saat ini)
    if (path1) {
      await workbook.xlsx.writeFile(path1);
    }
    // Simpan ke path 2 (folder gambar) jika bisa
    if (path2 && path2 !== path1) {
      try {
        await workbook.xlsx.writeFile(path2);
      } catch (e) {
        /* ignore jika locked / tanpa izin */
      }
    }
  } catch (err) {
    console.error(`  [Gagal save Excel: ${err.message}]`);
  }
}

// ─── Jalankan ─────────────────────────────────────────────────────────────
processFolder().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
