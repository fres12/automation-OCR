require('dotenv').config();
const { chromium } = require('playwright');
const xlsx = require('xlsx');
const fs = require('fs');
const fse = require('fs');
const path = require('path');
const os = require('os');

// ─── CONFIG ───────────────────────────────────────────────────────────────────
const FORM_URL = 'https://forms.gle/9oteZBmCz87DBRrt9';
const FORM_FILE = path.join(__dirname, 'form_format.xlsx');
const IMAGE_FOLDER = process.env.IMAGE_FOLDER || 'C:\\Users\\fresn\\OneDrive\\Documents\\Bali Nusra';

// Gunakan profile sementara agar tidak bentrok dengan Chrome yang sudah terbuka.
const CHROME_APP_PATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const TEMP_PROFILE_DIR = path.join(os.tmpdir(), 'playwright_form_profile');

// Kolom upload di Excel → diisi nama file saja, kita resolve ke path lokal
const UPLOAD_COLUMNS = [
  'Upload hasil snapshot size maksimum 1Mb dan aktifkan location tagging',
  'Upload aktivitas outlet engagement ( size maksimum 1Mb ) dan aktifkan location tagging',
  'Upload foto aktivitas outlet branding ( size maksimum 1Mb dan aktifkan location tagging )',
  'Upload Aktivitas Advocacy ( size maksimum 1Mb dan aktifkan location tagging )',
  'Upload Aktivitas Street Branding ( size maksimum 1Mb dan aktifkan location tagging )',
  'Upload Aktivitas POI Attack ( size maksimum 1Mb dan aktifkan location tagging )',
  'Upload Aktivitas Digital Campaign ( size maksimum 1Mb dan aktifkan location tagging )',
];

// ─── FILE SEARCH ──────────────────────────────────────────────────────────────
/**
 * Cari file secara rekursif di dalam folder berdasarkan nama file saja (case-insensitive).
 * Return path absolut jika ditemukan, null jika tidak.
 */
function findFileByName(folder, fileName) {
  if (!fileName || !fs.existsSync(folder)) return null;
  const target = fileName.trim().toLowerCase();

  function walk(dir) {
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return null;
    }
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        const found = walk(fullPath);
        if (found) return found;
      } else if (entry.name.toLowerCase() === target) {
        return fullPath;
      }
    }
    return null;
  }

  return walk(folder);
}

// ─── READ EXCEL ───────────────────────────────────────────────────────────────
function readFormData() {
  if (!fs.existsSync(FORM_FILE)) {
    throw new Error(`File tidak ditemukan: ${FORM_FILE}`);
  }
  const workbook = xlsx.readFile(FORM_FILE);
  const sheet = workbook.Sheets['Sheet1'] || workbook.Sheets[workbook.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: '', raw: false });
  if (!rows.length) throw new Error('Sheet1 kosong.');
  return rows;
}

// ─── PLAYWRIGHT HELPERS ───────────────────────────────────────────────────────
/**
 * Isi text field / short-answer.
 * Mencoba berbagai locator agar robust terhadap variasi label Google Form.
 */
async function fillTextField(page, labelRegex, value) {
  if (!value || String(value).trim() === '') return false;
  const v = String(value).trim();

  // Google Form: label ada di div[data-params], input ada di input[type=text] atau textarea
  // Strategi: cari container yang mengandung teks label, lalu isi input di dalamnya
  try {
    const containers = page.locator('[data-params], [jsname]').filter({
      hasText: labelRegex,
    });
    const count = await containers.count();
    if (count > 0) {
      const container = containers.first();
      const input = container.locator('input[type="text"], input:not([type]), textarea').first();
      if ((await input.count()) > 0) {
        await input.click();
        await input.fill(v);
        return true;
      }
    }
  } catch {}

  // Fallback: getByLabel
  try {
    const lbl = page.getByLabel(labelRegex);
    if ((await lbl.count()) > 0) {
      await lbl.first().fill(v);
      return true;
    }
  } catch {}

  // Fallback: getByRole textbox
  try {
    const tb = page.getByRole('textbox', { name: labelRegex });
    if ((await tb.count()) > 0) {
      await tb.first().fill(v);
      return true;
    }
  } catch {}

  return false;
}

/**
 * Pilih opsi pada dropdown / radio / select.
 * Google Form biasanya render select sebagai custom div, bukan <select>.
 */
async function selectOption(page, labelRegex, value) {
  if (!value || String(value).trim() === '') return false;
  const v = String(value).trim();

  // Coba native <select>
  try {
    const sel = page.locator('select').filter({ hasText: labelRegex });
    if ((await sel.count()) > 0) {
      await sel.first().selectOption({ label: v });
      return true;
    }
  } catch {}

  // Coba Google Form custom dropdown: klik container lalu klik opsi
  try {
    // Cari label, klik dropdown, lalu pilih option
    const containers = page.locator('[data-params]').filter({ hasText: labelRegex });
    if ((await containers.count()) > 0) {
      // Klik arrow/dropdown
      const dropBtn = containers.first().locator('[role="listbox"], [role="option"], select, [jsname="VncBx"]').first();
      if ((await dropBtn.count()) > 0) {
        await dropBtn.click();
        await page.waitForTimeout(300);
      } else {
        await containers.first().click();
        await page.waitForTimeout(300);
      }
      // Klik opsi yang matching
      const option = page.locator('[role="option"]').filter({ hasText: new RegExp(`^${v}$`, 'i') });
      if ((await option.count()) > 0) {
        await option.first().click();
        return true;
      }
    }
  } catch {}

  // Fallback: isi sebagai text
  return fillTextField(page, labelRegex, value);
}

/**
 * Upload file ke file input.
 * Google Form: ada tombol "Tambahkan file" yang membuka dialog,
 * tapi kita pakai metode langsung setInputFiles pada input[type=file].
 */
async function uploadFile(page, labelRegex, filePath) {
  if (!filePath) {
    console.log(`  [SKIP UPLOAD] Tidak ada file untuk: ${labelRegex}`);
    return false;
  }
  if (!fs.existsSync(filePath)) {
    console.log(`  [FILE NOT FOUND] ${filePath}`);
    return false;
  }

  console.log(`  [UPLOAD] ${path.basename(filePath)}`);

  try {
    // Cari tombol upload di dalam container label
    const containers = page.locator('[data-params]').filter({ hasText: labelRegex });
    if ((await containers.count()) > 0) {
      const fileInput = containers.first().locator('input[type="file"]');
      if ((await fileInput.count()) > 0) {
        await fileInput.setInputFiles(filePath);
        await page.waitForTimeout(1500);
        return true;
      }

      // Klik tombol "Tambahkan file" lalu intercept file chooser
      const addBtn = containers
        .first()
        .locator('button, [role="button"]')
        .filter({ hasText: /tambah|add|pilih|choose|upload/i })
        .first();
      if ((await addBtn.count()) > 0) {
        const [fileChooser] = await Promise.all([
          page.waitForEvent('filechooser', { timeout: 5000 }),
          addBtn.click(),
        ]);
        await fileChooser.setFiles(filePath);
        await page.waitForTimeout(1500);
        return true;
      }
    }

    // Fallback: file input global
    const allInputs = page.locator('input[type="file"]');
    if ((await allInputs.count()) > 0) {
      await allInputs.first().setInputFiles(filePath);
      await page.waitForTimeout(1500);
      return true;
    }
  } catch (err) {
    console.log(`  [UPLOAD ERROR] ${err.message}`);
  }

  return false;
}



const { exec, spawn } = require('child_process');
const http = require('http');

// ─── HELPER CDP & CHROME LAUNCHER ─────────────────────────────────────────────
async function isCdpReady() {
  return new Promise((resolve) => {
    const req = http.get('http://127.0.0.1:9222/json/version', (res) => {
      resolve(res.statusCode === 200);
    });
    req.on('error', () => resolve(false));
    req.setTimeout(1000, () => {
      req.destroy();
      resolve(false);
    });
  });
}

async function ensureChromeWithCDP() {
  if (await isCdpReady()) {
    console.log('  [BROWSER] Chrome dengan debugging port 9222 sudah aktif.');
    return;
  }

  console.log('  [BROWSER] Membuka Chrome secara otomatis (Profile: siregarfresnel@gmail.com)...');

  const chromeProcess = spawn(
    CHROME_APP_PATH,
    [
      '--remote-debugging-port=9222',
      `--user-data-dir=${CHROME_USER_DATA}`,
      `--profile-directory=${CHROME_PROFILE}`,
      '--restore-last-session=false',
      '--no-first-run',
      '--no-default-browser-check',
    ],
    {
      detached: true,
      stdio: 'ignore',
    }
  );
  chromeProcess.unref();

  for (let i = 0; i < 30; i++) {
    await new Promise((r) => setTimeout(r, 500));
    if (await isCdpReady()) {
      console.log('  [BROWSER] Chrome berhasil dibuka dan terhubung!');
      return;
    }
  }

  throw new Error(
    'Gagal membuka Chrome dengan mode debugging. Pastikan Chrome tidak terkunci, lalu jalankan ulang script.'
  );
}

// ─── MAIN FILL FUNCTION ───────────────────────────────────────────────────────
async function fillAndSubmitRow(row) {
  let browser;
  let context;
  let page;

  try {
    console.log('  [BROWSER] Membuka Chrome dengan profile sementara...');
    context = await chromium.launchPersistentContext(TEMP_PROFILE_DIR, {
      channel: 'chrome',
      headless: false,
      args: [
        '--no-first-run',
        '--no-default-browser-check',
      ],
    });
    page = await context.newPage();
  } catch (err) {
    console.error('\n❌ GAGAL MEMBUKA CHROME!', err.message);
    process.exit(1);
  }

  try {
    console.log('\n========================================');
    console.log(`Membuka form: ${FORM_URL}`);
    await page.goto(FORM_URL, { waitUntil: 'networkidle', timeout: 30000 });
    await page.waitForTimeout(2000);

    // ── TEXT FIELDS ────────────────────────────────────────────────────────
    const textFields = [
      { label: /^TITLE$/i,    value: row['TITLE'] },
      { label: /^CIRCLE$/i,   value: row['CIRCLE'] },
      { label: /^REGION$/i,   value: row['REGION'] },
      { label: /^BRAND$/i,    value: row['BRAND'] },
      { label: /^AREA$/i,     value: row['AREA'] },
      { label: /Site ID/i,    value: row['Site ID'] },
      {
        label: /Site Testing Download Speed 5G dalam Mbps/i,
        value: row['Site Testing Download Speed 5G dalam Mbps'],
      },
      {
        label: /Site Testing Upload Speed 5G dalam Mbps/i,
        value: row['Site Testing Upload Speed 5G dalam Mbps'],
      },
      {
        label: /Step 1.*Site Testing/i,
        value: row['Step 1 : Site Testing\r\n( Kecepatan minimum 50 Mbps )'],
      },
      { label: /Outlet ID/i,    value: row['Outlet ID'] },
      { label: /Nama Outlet/i,  value: row['Nama Outlet'] },
      {
        label: /Step 2.*Outlet Engagement/i,
        value: row['Step 2 : Outlet Engagement\r\n( Memastikan outlet memahami kegiatan/program yang ada & menjaga hubungan baik dengan outlet )'],
      },
      {
        label: /Step 3.*Outlet Branding/i,
        value: row['Step 3 : Outlet Branding\r\n(  Pemasangan branding di outlet yang meliputi sisi depan outlet, etalase, maupun didalam outlet )'],
      },
      {
        label: /Step 4.*Network.*Product Advocacy/i,
        value: row['Step 4 : Network & Product Advocacy\r\n( Memastikan awareness jaringan & produk di sekitar lokasi )'],
      },
      {
        label: /Step 5.*Street Branding/i,
        value: row['Step 5 : Street Branding\r\n(  Melakukan pemasangan branding material dijalan-jalan utama disekitar site )'],
      },
      {
        label: /Step 6.*POI Attack/i,
        value: row['Step 6 : POI Attack'],
      },
      {
        label: /Step 7.*Digital Campaign/i,
        value: row['Step 7 : Digital Campaign'],
      },
      { label: /EMPLOYEE ID/i, value: row['EMPLOYEE ID'] },
    ];

    for (const { label, value } of textFields) {
      if (!value || String(value).trim() === '') {
        console.log(`  [SKIP] ${label} (kosong)`);
        continue;
      }
      console.log(`  [FILL] ${label} = "${value}"`);
      const ok = await selectOption(page, label, value);
      if (!ok) console.log(`  [WARN] Tidak berhasil isi: ${label}`);
    }

    // ── UPLOAD FIELDS ──────────────────────────────────────────────────────
    const uploadDefs = [
      {
        label: /Upload hasil snapshot/i,
        column: 'Upload hasil snapshot size maksimum 1Mb dan aktifkan location tagging',
      },
      {
        label: /Upload aktivitas outlet engagement/i,
        column: 'Upload aktivitas outlet engagement ( size maksimum 1Mb ) dan aktifkan location tagging',
      },
      {
        label: /Upload foto aktivitas outlet branding/i,
        column: 'Upload foto aktivitas outlet branding ( size maksimum 1Mb dan aktifkan location tagging )',
      },
      {
        label: /Upload Aktivitas Advocacy/i,
        column: 'Upload Aktivitas Advocacy ( size maksimum 1Mb dan aktifkan location tagging )',
      },
      {
        label: /Upload Aktivitas Street Branding/i,
        column: 'Upload Aktivitas Street Branding ( size maksimum 1Mb dan aktifkan location tagging )',
      },
      {
        label: /Upload Aktivitas POI Attack/i,
        column: 'Upload Aktivitas POI Attack ( size maksimum 1Mb dan aktifkan location tagging )',
      },
      {
        label: /Upload Aktivitas Digital Campaign/i,
        column: 'Upload Aktivitas Digital Campaign ( size maksimum 1Mb dan aktifkan location tagging )',
      },
    ];

    for (const { label, column } of uploadDefs) {
      const fileName = row[column];
      if (!fileName || String(fileName).trim() === '') {
        console.log(`  [SKIP UPLOAD] ${column} (kosong)`);
        continue;
      }
      console.log(`  [CARI FILE] "${fileName}" di ${IMAGE_FOLDER}`);
      const filePath = findFileByName(IMAGE_FOLDER, fileName.trim());
      if (filePath) {
        console.log(`  [DITEMUKAN] ${filePath}`);
        await uploadFile(page, label, filePath);
      } else {
        console.log(`  [FILE TIDAK DITEMUKAN] ${fileName}`);
      }
    }

    // ── SUBMIT ─────────────────────────────────────────────────────────────
    console.log('\n  [SUBMIT] Mencari tombol submit...');
    await page.waitForTimeout(1000);

    const submitBtn = page
      .getByRole('button', { name: /kirim|submit/i })
      .first();

    if ((await submitBtn.count()) > 0) {
      console.log('  [SUBMIT] Klik tombol submit...');
      await submitBtn.click();
      await page.waitForTimeout(3000);
      console.log('  [DONE] Form berhasil disubmit!');
    } else {
      console.log('  [ERROR] Tombol submit tidak ditemukan!');
    }

    // Tunggu sebentar sebelum tutup
    await page.waitForTimeout(3000);
  } catch (err) {
    console.error('  [ERROR]', err.message);
  } finally {
    // Kita biarkan Chrome tetap terbuka (jangan tutup context)
    // Tapi kita bisa tutup tab/page yang baru kita buat
    if (page) await page.close();
  }
}

// ─── ENTRY POINT ──────────────────────────────────────────────────────────────
async function run() {
  console.log('=== GOOGLE FORM AUTO-FILL ===');
  console.log(`Excel: ${FORM_FILE}`);
  console.log(`Image Folder: ${IMAGE_FOLDER}`);

  const rows = readFormData();
  console.log(`\nDitemukan ${rows.length} baris data di Sheet1.`);

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    console.log(`\n[${i + 1}/${rows.length}] Site ID: ${row['Site ID'] || '(kosong)'}`);
    // Debug: tampilkan semua nilai
    for (const [k, v] of Object.entries(row)) {
      if (v && String(v).trim() !== '') {
        console.log(`   ${k}: ${v}`);
      }
    }
    await fillAndSubmitRow(row);
  }

  console.log('\n=== SELESAI ===');
}

run().catch((err) => {
  console.error('Fatal error:', err);
  process.exit(1);
});