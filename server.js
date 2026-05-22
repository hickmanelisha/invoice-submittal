'use strict';

const express    = require('express');
const multer     = require('multer');
const { execFile, execFileSync } = require('child_process');
const path       = require('path');
const fs         = require('fs');
const os         = require('os');
const { PDFDocument } = require('pdf-lib');
const PizZip     = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app  = express();
const PORT = process.env.PORT || 3002;

// ─── Paths ────────────────────────────────────────────────────────────────────
const TEMPLATES_DIR  = path.join(__dirname, 'templates');
const TEMPLATES_META = path.join(TEMPLATES_DIR, 'meta.json');

// ─── Configuration ─────────────────────────────────────────────────────────────
const CONTRACTORS = [
  'Looks Great Services',
  'AshBritt Environmental',
  'Ceres Environmental',
  'DRC Emergency Services',
  'Crowder Gulf',
  'TFR Enterprises',
  'SDR',
  'CTC Disaster Response',
];

const SENDERS = {
  'Angelia Cruthirds': 'signatures/angelia_cruthirds.png',
  'Elisha Hickman':    'signatures/elisha_hickman.png',
  'Buck Dickinson':    'signatures/buck_dickinson.png',
};

// ─── Date helpers ──────────────────────────────────────────────────────────────
const MONTHS = [
  'January','February','March','April','May','June',
  'July','August','September','October','November','December',
];

function isoToLong(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  return `${MONTHS[m - 1]} ${d}, ${y}`;
}

function isoToMDY(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${m}-${d}-${y}`;
}

// ─── Template bug fixes ────────────────────────────────────────────────────────
// Must be called AFTER PizZip has decompressed the file — raw ZIP bytes are
// DEFLATE-compressed so string replacement on the buffer does nothing.

// Word splits placeholder text across multiple <w:r> runs and inserts
// <w:proofErr> spell-check markers between them.  After stripping proofErr,
// docxtemplater still fails because it counts every isolated {{ as a new tag
// opener, producing "duplicate open tag" errors.
//
// The fix: collapse any sequence of adjacent runs whose combined <w:t> text
// forms a complete {{…}} tag into a single run, then hand the clean XML to
// docxtemplater.
function mergeTagRuns(xml) {
  const rpr = '(?:<w:rPr>(?:(?!<\\/w:rPr>)[\\s\\S])*<\\/w:rPr>)?';
  const run = `<w:r\\b[^>]*>${rpr}<w:t[^>]*>`;

  // Pass 1 — {{ alone in its own run (most common Word split)
  const pat1 = new RegExp(
    `${run}\\{\\{<\\/w:t><\\/w:r>((?:${run}[^<]*<\\/w:t><\\/w:r>)+?)${run}\\}\\}\\}?<\\/w:t><\\/w:r>`,
    'g'
  );
  xml = xml.replace(pat1, (_m, middle) => {
    const tag = [...middle.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
    return `<w:r><w:t>{{${tag}}}</w:t></w:r>`;
  });

  // Pass 2 — text + {{ in same run (e.g. " {{"), closing run may have }}} not just }}
  const pat2 = new RegExp(
    `(${run})([^<]+)\\{\\{<\\/w:t><\\/w:r>((?:${run}[^<]*<\\/w:t><\\/w:r>)+?)${run}\\}\\}\\}?<\\/w:t><\\/w:r>`,
    'g'
  );
  xml = xml.replace(pat2, (_m, openRun, leadText, middle) => {
    const tag = [...middle.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
    return `<w:r><w:t xml:space="preserve">${leadText}</w:t></w:r><w:r><w:t>{{${tag}}}</w:t></w:r>`;
  });

  return xml;
}

// ─── Font enforcement: Times New Roman 12 pt ──────────────────────────────────
// Replaces all explicit font and size declarations in a piece of OOXML so that
// the rendered document uses Times New Roman at 12 pt throughout.
// 24 half-points = 12 pt (Word stores font sizes in half-points).
function applyTimesNewRoman12pt(xml) {
  // Replace every <w:rFonts .../> with a clean TNR declaration
  xml = xml.replace(
    /<w:rFonts\b[^>]*\/?>/g,
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
  );
  // Replace every <w:sz .../> (word boundary avoids matching <w:szCs>)
  xml = xml.replace(/<w:sz\b[^>]*\/?>/g,   '<w:sz w:val="24"/>');
  xml = xml.replace(/<w:szCs\b[^>]*\/?>/g, '<w:szCs w:val="24"/>');
  return xml;
}

function fixTemplateBugsInZip(zip) {
  zip.file(/\.xml/).forEach(file => {
    let xml = file.asText();

    // 1. Remove spell-check markers (they separate the {{ and }} from the tag name)
    xml = xml.replace(/<w:proofErr\b[^>]*\/?>/g, '');

    // 2. Merge split runs so {{ tag_name }} becomes one run docxtemplater can see
    xml = mergeTagRuns(xml);

    // 3. Fix known placeholder typos in this template
    xml = xml.replace(/\{\{invoice amount\}\}/g,    '{{invoice_amount}}');
    xml = xml.replace(/\{\{sender_signature\}\}/g,  '{{%sender_signature}}');

    // 4. Remove stray single-} runs left by template authoring bugs
    const rpr2 = '(?:<w:rPr>(?:(?!<\\/w:rPr>)[\\s\\S])*<\\/w:rPr>)?';
    xml = xml.replace(
      new RegExp('<w:r\\b[^>]*>' + rpr2 + '<w:t[^>]*>}<\\/w:t><\\/w:r>', 'g'),
      ''
    );

    // 5. Enforce Times New Roman 12 pt
    xml = applyTimesNewRoman12pt(xml);

    zip.file(file.name, xml);
  });
}

// ─── Multi-template helpers ────────────────────────────────────────────────────
function getTemplatesMeta() {
  if (!fs.existsSync(TEMPLATES_META)) return {};
  try { return JSON.parse(fs.readFileSync(TEMPLATES_META, 'utf8')); }
  catch (_) { return {}; }
}

function saveTemplatesMeta(meta) {
  fs.writeFileSync(TEMPLATES_META, JSON.stringify(meta, null, 2));
}

function listTemplates() {
  if (!fs.existsSync(TEMPLATES_DIR)) return [];
  const meta = getTemplatesMeta();
  return Object.entries(meta)
    .map(([id, info]) => {
      const filePath = path.join(TEMPLATES_DIR, id);
      if (!fs.existsSync(filePath)) return null;
      const stat = fs.statSync(filePath);
      return { id, name: info.originalName, savedAt: info.savedAt, sizeBytes: stat.size };
    })
    .filter(Boolean)
    .sort((a, b) => a.savedAt - b.savedAt);
}

function getTemplatePath(id) {
  if (!id) return null;
  // Prevent directory traversal
  const safe = path.basename(id);
  const p = path.join(TEMPLATES_DIR, safe);
  return fs.existsSync(p) ? p : null;
}

function sanitizeFilename(name) {
  return name.replace(/[^a-zA-Z0-9._-]/g, '_');
}

// ─── Sync meta.json to match files on disk at startup ────────────────────────
// Removes meta entries whose file is gone, removes orphan files not in meta.
// This prevents stale Railway volume entries from showing in the dropdown.
(function syncTemplatesOnStartup() {
  if (!fs.existsSync(TEMPLATES_DIR)) return;
  const meta = getTemplatesMeta();
  let changed = false;

  // Remove meta entries with no matching file
  for (const id of Object.keys(meta)) {
    if (!fs.existsSync(path.join(TEMPLATES_DIR, id))) {
      console.log(`Removing stale meta entry: ${id}`);
      delete meta[id];
      changed = true;
    }
  }

  // Remove files not in meta (orphans from old deployments)
  for (const file of fs.readdirSync(TEMPLATES_DIR)) {
    if (file === 'meta.json') continue;
    if (!meta[file]) {
      console.log(`Removing orphan template file: ${file}`);
      fs.rmSync(path.join(TEMPLATES_DIR, file), { force: true });
    }
  }

  if (changed) saveTemplatesMeta(meta);
})();

// ─── Migrate legacy single-template dir on first run ──────────────────────────
(function migrateLegacyTemplate() {
  const OLD_DIR  = path.join(__dirname, 'template');
  const OLD_META = path.join(OLD_DIR, 'meta.json');
  if (!fs.existsSync(OLD_DIR)) return;
  const oldFiles = fs.readdirSync(OLD_DIR).filter(f => f !== 'meta.json');
  if (!oldFiles.length) return;
  let originalName = oldFiles[0];
  try {
    const m = JSON.parse(fs.readFileSync(OLD_META, 'utf8'));
    if (m.originalName) originalName = m.originalName;
  } catch (_) {}
  if (!fs.existsSync(TEMPLATES_DIR)) fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
  const id = sanitizeFilename(originalName);
  const dest = path.join(TEMPLATES_DIR, id);
  if (!fs.existsSync(dest)) {
    fs.copyFileSync(path.join(OLD_DIR, oldFiles[0]), dest);
    const meta = getTemplatesMeta();
    meta[id] = { originalName, savedAt: Date.now() };
    saveTemplatesMeta(meta);
    console.log(`Migrated legacy template: ${originalName}`);
  }
})();

// ─── PDF converter detection (LibreOffice preferred, Word fallback) ────────────

// LibreOffice — preferred converter (headless, no dialogs, no AppleScript)
const LO_CANDIDATES = [
  '/Applications/LibreOffice.app/Contents/MacOS/soffice',
  '/usr/bin/soffice',
  '/usr/bin/libreoffice',
  '/snap/bin/libreoffice',
  'soffice',
  'libreoffice',
];
let LIBRE_OFFICE_PATH = null;
for (const c of LO_CANDIDATES) {
  try {
    execFileSync(c, ['--version'], { stdio: 'pipe', timeout: 8000 });
    LIBRE_OFFICE_PATH = c;
    break;
  } catch (_) {}
}

// Microsoft Word — fallback if LibreOffice is not installed
const WORD_APP_PATHS = [
  '/Applications/Microsoft Word.app',
  `${os.homedir()}/Applications/Microsoft Word.app`,
];
const HAS_WORD = WORD_APP_PATHS.some(p => fs.existsSync(p));

if (LIBRE_OFFICE_PATH) {
  console.log(`LibreOffice found at ${LIBRE_OFFICE_PATH} — will use for DOCX→PDF conversion`);
} else if (HAS_WORD) {
  console.log('Microsoft Word found — will use for DOCX→PDF conversion (install LibreOffice for better reliability)');
} else {
  console.warn('No PDF converter found. Install LibreOffice (libreoffice.org) to enable PDF generation.');
}

// ─── DOCX → PDF conversion ────────────────────────────────────────────────────
// Writes the PDF next to the input file in the same directory.
// Returns the path to the generated PDF.
async function convertDocxToPdf(docxPath, outputDir) {
  const pdfPath = path.join(outputDir, 'cover_letter.pdf');

  if (LIBRE_OFFICE_PATH) {
    await new Promise((resolve, reject) => {
      execFile(
        LIBRE_OFFICE_PATH,
        ['--headless', '--convert-to', 'pdf', '--outdir', outputDir, docxPath],
        { timeout: 60_000 },
        (err, _out, stderr) => {
          if (err) reject(new Error(`LibreOffice error: ${stderr || err.message}`));
          else resolve();
        }
      );
    });

  } else if (HAS_WORD) {
    // Word fallback via AppleScript — less reliable than LibreOffice
    const escaped = (s) => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    const script = [
      'tell application "Microsoft Word"',
      '  activate',
      '  delay 1',
      '  set docsBefore to count of documents',
      `  open POSIX file "${escaped(docxPath)}"`,
      '  set waited to 0',
      '  repeat while (count of documents) <= docsBefore and waited < 30',
      '    delay 0.5',
      '    set waited to waited + 0.5',
      '  end repeat',
      '  if (count of documents) <= docsBefore then error "Document failed to open within 30 seconds"',
      '  set theDoc to document 1',
      `  save as theDoc file name "${escaped(pdfPath)}" file format format PDF`,
      '  close theDoc saving no',
      'end tell',
    ].join('\n');
    const scriptPath = path.join(outputDir, 'convert.scpt');
    fs.writeFileSync(scriptPath, script);
    await new Promise((resolve, reject) => {
      execFile('osascript', [scriptPath], { timeout: 60_000 }, (err, _out, stderr) => {
        if (err) reject(new Error(`Microsoft Word failed: ${stderr || err.message}`));
        else resolve();
      });
    });

  } else {
    throw new Error(
      'No PDF converter is available. Install LibreOffice (libreoffice.org) to enable PDF generation.'
    );
  }

  return pdfPath;
}

// ─── Invoice PDF text parser ──────────────────────────────────────────────────
// Converts M/D/YY, M/D/YYYY, or M.D.YYYY → YYYY-MM-DD for <input type="date">
function mdyToIso(mdy) {
  if (!mdy) return '';
  const [m, d, y] = mdy.split(/[\/\.]/);
  const year = y.length === 2 ? '20' + y : y;
  return `${year}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
}

// ─── SDR invoice parser ───────────────────────────────────────────────────────
// SDR invoices use "BILL TO" (no colon), "INVOICE #", "DATE", "BALANCE DUE"
function parseSDRInvoice(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const result = {};

  result.contractor_name = 'SDR';

  // Invoice number — "INVOICE # 24130103111"
  const invMatch = text.match(/INVOICE\s*#\s+([A-Z0-9-]+)/i);
  if (invMatch) result.invoice_number = invMatch[1];

  // PE date — "DATE 04/30/2026" (single date, not a range)
  const dateMatch = text.match(/\bDATE\s+(\d{1,2}\/\d{1,2}\/\d{2,4})/i);
  if (dateMatch) result.pe_date = mdyToIso(dateMatch[1]);
  // No start_date on SDR invoices — user enters manually

  // Amount — "BALANCE DUE $75,930.15"
  const balMatch = text.match(/BALANCE DUE\s+\$?([\d,]+(?:\.\d{2})?)/i);
  if (balMatch) result.invoice_amount = '$' + balMatch[1];

  // Bill-To block — lines after "BILL TO" until an invoice-field keyword
  const billToIdx = lines.findIndex(l => /^BILL\s+TO$/i.test(l));
  if (billToIdx >= 0) {
    const addrLines = [];
    for (let i = billToIdx + 1; i < Math.min(billToIdx + 8, lines.length); i++) {
      const l = lines[i];
      if (/^(INVOICE|DATE|DUE DATE|TERMS)/i.test(l)) break;
      if (/@/.test(l)) continue;
      addrLines.push(l);
    }
    if (addrLines[0]) result.project_name   = addrLines[0];
    if (addrLines[1]) result.street_address = addrLines[1];
    const cityLine = addrLines.find(l => /\d{5}/.test(l));
    if (cityLine) result.city_state_zip = cityLine;
  }

  return result;
}

// ─── MS invoice parser (unchanged) ───────────────────────────────────────────
// MS invoices use "Bill To:" (with colon + Attn: block), "Invoice Number:", date range in parens
function parseMSInvoice(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const result = {};

  // Contractor name — first line before "Bill To:"
  const billToIdx = lines.findIndex(l => /bill\s+to/i.test(l));
  if (billToIdx > 0) result.contractor_name = lines[0];

  // Invoice number
  const invNumMatch = text.match(/Invoice\s+Number[:\s]+([A-Z0-9-]+)/i);
  if (invNumMatch) result.invoice_number = invNumMatch[1];

  // Invoice total — last "Total $X" on the page
  const totalMatches = [...text.matchAll(/\bTotal\s+\$?([\d,]+(?:\.\d{2})?)/gi)];
  if (totalMatches.length) {
    const last = totalMatches[totalMatches.length - 1];
    result.invoice_amount = '$' + last[1];
  }

  // Date range — "Invoice (3/29/2026 - 04/11/2026)"
  const dateRangeMatch = text.match(
    /\((\d{1,2}\/\d{1,2}\/\d{2,4})\s*[-–]\s*(\d{1,2}\/\d{1,2}\/\d{2,4})\)/
  );
  if (dateRangeMatch) {
    result.start_date = mdyToIso(dateRangeMatch[1]);
    result.pe_date    = mdyToIso(dateRangeMatch[2]);
  }

  // Bill-To block: project name, contact, street, city/state/zip
  if (billToIdx >= 0) {
    const clientLine = lines[billToIdx + 1] || '';
    const clientName = clientLine.replace(/Invoice Number.*$/i, '').trim();
    if (clientName) result.project_name = clientName;

    for (let i = billToIdx; i < Math.min(billToIdx + 10, lines.length); i++) {
      const attnMatch = lines[i].match(/^Attn:\s*(.+)/i);
      if (!attnMatch) continue;

      result.contact_name = attnMatch[1].replace(/Job Number.*$/i, '').trim();

      const addrCandidates = [];
      for (let j = i + 1; j < Math.min(i + 7, lines.length); j++) {
        const l = lines[j];
        if (/@/.test(l)) continue;
        if (/Invoice|Bill To/i.test(l)) break;
        addrCandidates.push(l);
      }
      if (addrCandidates[0]) result.street_address = addrCandidates[0];
      const cityLine = addrCandidates.find(l => /\d{5}/.test(l));
      if (cityLine) result.city_state_zip = cityLine;
      break;
    }
  }

  return result;
}

// ─── CTC invoice parser ───────────────────────────────────────────────────────
// CTC invoices: "Bill to" (lowercase), "Invoice no.:", "Period Ending" date, "Ship to" boundary
function parseCTCInvoice(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const result = {};

  result.contractor_name = 'CTC Disaster Response';

  // Invoice number — "Invoice no.: CTC10030"
  const invMatch = text.match(/Invoice\s+no\.\s*:\s*([A-Z0-9-]+)/i);
  if (invMatch) result.invoice_number = invMatch[1];

  // PE date — "Period Ending 05/16/2026" or "Period Ending 04.30.2026"
  const peMatch = text.match(/Period\s+Ending\s+(\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4})/i);
  if (peMatch) result.pe_date = mdyToIso(peMatch[1]);
  // No start_date — user enters manually

  // Amount — "Total $2,258,543.91"
  const totalMatch = text.match(/\bTotal\s+\$?([\d,]+(?:\.\d{2})?)/i);
  if (totalMatch) result.invoice_amount = '$' + totalMatch[1];

  // Bill-to block — lines after "Bill to", stop at "Ship to"
  const billToIdx = lines.findIndex(l => /^Bill\s+to$/i.test(l));
  if (billToIdx >= 0) {
    const addrLines = [];
    for (let i = billToIdx + 1; i < Math.min(billToIdx + 8, lines.length); i++) {
      const l = lines[i];
      if (/^Ship\s+to$/i.test(l)) break;
      if (/@/.test(l)) continue;
      addrLines.push(l);
    }
    if (addrLines[0]) result.project_name   = addrLines[0];
    if (addrLines[1]) result.street_address = addrLines[1];
    const cityLine = addrLines.find(l => /\d{5}/.test(l));
    if (cityLine) result.city_state_zip = cityLine;
  }

  return result;
}

// ─── Invoice parser dispatcher ────────────────────────────────────────────────
function parseInvoiceText(text) {
  // SDR invoices: uppercase "INVOICE #" with space, and "SDR" in text
  const isSDR = /INVOICE\s*#\s+\d/i.test(text) && /SDR/i.test(text);
  // CTC invoices: "Invoice no.:" pattern and "CTC" in text
  const isCTC = /Invoice\s+no\.\s*:/i.test(text) && /CTC/i.test(text);
  if (isSDR) return parseSDRInvoice(text);
  if (isCTC) return parseCTCInvoice(text);
  return parseMSInvoice(text);
}

// ─── Express / Multer ─────────────────────────────────────────────────────────
app.use(express.static(path.join(__dirname, 'public')));

const memUpload  = multer({ storage: multer.memoryStorage(), limits: { fileSize: 100 * 1024 * 1024 } });
const diskUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

// ─── GET /api/options ─────────────────────────────────────────────────────────
app.get('/api/options', (_req, res) => {
  res.json({ contractors: CONTRACTORS, senders: Object.keys(SENDERS) });
});

// ─── GET /api/templates ───────────────────────────────────────────────────────
// Returns list of all stored templates.
app.get('/api/templates', (_req, res) => {
  res.json(listTemplates());
});

// ─── POST /api/templates ──────────────────────────────────────────────────────
// Uploads and saves a new template (does not replace existing ones).
app.post('/api/templates', diskUpload.single('template'), (req, res) => {
  const file = req.file;
  if (!file) return res.status(400).json({ error: 'No file uploaded.' });

  const lower = file.originalname.toLowerCase();
  if (!lower.endsWith('.docx') && !lower.endsWith('.dotx')) {
    return res.status(400).json({ error: 'Template must be a .docx or .dotx file.' });
  }

  if (!fs.existsSync(TEMPLATES_DIR)) fs.mkdirSync(TEMPLATES_DIR, { recursive: true });

  const id       = sanitizeFilename(file.originalname);
  const savePath = path.join(TEMPLATES_DIR, id);
  fs.writeFileSync(savePath, file.buffer);

  const meta = getTemplatesMeta();
  meta[id] = { originalName: file.originalname, savedAt: Date.now() };
  saveTemplatesMeta(meta);

  console.log(`Template saved: ${file.originalname} (id: ${id})`);
  res.json({ success: true, id, name: file.originalname });
});

// ─── DELETE /api/templates/:id ────────────────────────────────────────────────
app.delete('/api/templates/:id', (req, res) => {
  const id       = path.basename(req.params.id);   // prevent traversal
  const filePath = path.join(TEMPLATES_DIR, id);
  if (fs.existsSync(filePath)) fs.rmSync(filePath, { force: true });
  const meta = getTemplatesMeta();
  delete meta[id];
  saveTemplatesMeta(meta);
  res.json({ success: true });
});

// ─── POST /api/parse-invoice ──────────────────────────────────────────────────
// Accepts a PDF upload and returns extracted field values for auto-filling the form.
// Uses LibreOffice (already required for DOCX→PDF) to extract the text.
app.post('/api/parse-invoice', memUpload.single('invoice'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'inv-parse-'));
  const tmpPdf = path.join(tmpDir, 'invoice.pdf');

  try {
    fs.writeFileSync(tmpPdf, req.file.buffer);

    // Extract plain text from PDF using Python's pypdf
    const pyScript = [
      'from pypdf import PdfReader, errors',
      'import sys',
      'try:',
      '  reader = PdfReader(sys.argv[1])',
      '  print("\\n".join(p.extract_text() or "" for p in reader.pages))',
      'except Exception as e:',
      '  sys.exit(0)',  // exit cleanly — form will stay blank
    ].join('\n');

    const text = await new Promise((resolve) => {
      execFile('python3', ['-c', pyScript, tmpPdf], { timeout: 30_000 }, (err, stdout) => {
        resolve(err ? '' : stdout);
      });
    });

    res.json(parseInvoiceText(text));
  } catch (e) {
    res.status(500).json({ error: 'Could not parse PDF: ' + e.message });
  } finally {
    // Clean up temp files
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch (_) {}
  }
});

// ─── POST /generate ───────────────────────────────────────────────────────────
app.post(
  '/generate',
  memUpload.single('invoice'),
  async (req, res) => {
    let tempDir = null;

    try {
      // ── 1. Resolve template ──────────────────────────────────────────────
      const { template_id } = req.body;
      let templatePath = null;
      if (template_id) {
        templatePath = getTemplatePath(template_id);
      } else {
        // Fall back: use first available template
        const all = listTemplates();
        if (all.length) templatePath = getTemplatePath(all[0].id);
      }
      if (!templatePath) {
        return res.status(400).json({
          error: 'No cover letter template selected. Please upload and select a template first.',
        });
      }

      // ── 2. Validate inputs ───────────────────────────────────────────────
      const {
        today_date,
        project_name,
        contact_name,
        street_address,
        city_state_zip,
        contractor_name,
        pe_date,
        start_date,
        invoice_number,
        invoice_amount,
        sender_name,
      } = req.body;

      const missing = [];
      if (!today_date?.trim())       missing.push('Letter Date');
      if (!project_name?.trim())     missing.push('Client / County');
      if (!street_address?.trim())   missing.push('Street Address');
      if (!city_state_zip?.trim())   missing.push('City, State, ZIP');
      if (!contractor_name?.trim())  missing.push('Contractor Name');
      if (!pe_date?.trim())          missing.push('Period Ending Date');
      if (!start_date?.trim())       missing.push('Work Start Date');
      if (!invoice_number?.trim())   missing.push('Invoice Number');
      if (!invoice_amount?.trim())   missing.push('Invoice Amount');
      if (!sender_name?.trim())      missing.push('Sender Name');
      if (!req.file)                 missing.push('Invoice PDF');

      if (missing.length) {
        return res.status(400).json({ error: `Missing: ${missing.join(', ')}` });
      }

      if (!req.file.originalname.toLowerCase().endsWith('.pdf')) {
        return res.status(400).json({ error: 'Invoice file must be a .pdf' });
      }

      if (!HAS_WORD && !LIBRE_OFFICE_PATH) {
        return res.status(500).json({
          error: 'Microsoft Word is not installed. Please install Word and restart the server.',
        });
      }

      // ── 3. Derive formatted dates ────────────────────────────────────────
      const date_range         = isoToMDY(pe_date);
      const invoice_end_date   = isoToLong(pe_date);
      const invoice_start_date = isoToLong(start_date);
      const formatted_today    = isoToLong(today_date);

      // ── 4. Load template & fix known placeholder bugs in the decompressed XML
      let zip;
      try { zip = new PizZip(fs.readFileSync(templatePath)); }
      catch (_) { return res.status(500).json({ error: 'Could not parse the stored template.' }); }

      fixTemplateBugsInZip(zip);

      // ── 5. Attach image module if signature exists ────────────────────────
      const sigRelPath  = SENDERS[sender_name];
      const sigAbsPath  = sigRelPath ? path.resolve(__dirname, sigRelPath) : null;
      const hasSignature = sigAbsPath && fs.existsSync(sigAbsPath);

      const modules = [];
      if (hasSignature) {
        try {
          const ImageModule = require('docxtemplater-image-module-free');
          modules.push(new ImageModule({
            centered: false,
            fileType: 'docx',
            getImage(v) { return fs.readFileSync(v); },
            getSize()   { return [150, 50]; },
          }));
        } catch (e) {
          console.warn('Image module unavailable:', e.message);
        }
      }

      // ── 6. Fill template ─────────────────────────────────────────────────
      const doc = new Docxtemplater(zip, {
        modules,
        paragraphLoop: true,
        linebreaks: true,
        delimiters: { start: '{{', end: '}}' },
      });

      try {
        doc.render({
          today_date: formatted_today,
          project_name,
          contact_name,
          street_address,
          city_state_zip,
          contractor_name,
          date_range,
          invoice_start_date,
          invoice_end_date,
          invoice_number,
          invoice_amount,
          sender_name,
          ...(hasSignature && modules.length ? { sender_signature: sigAbsPath } : {}),
        });
      } catch (e) {
        const errs = e?.properties?.errors;
        const details = errs?.length
          ? errs.map(err => err?.properties?.explanation || err?.message || JSON.stringify(err)).join('; ')
          : e.message;
        return res.status(400).json({ error: `Template error: ${details}` });
      }

      const processedDocx = doc.getZip().generate({ type: 'nodebuffer' });

      // ── 7. Convert DOCX → PDF via Word (or LibreOffice fallback) ────────────
      tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'inv-'));
      const tempDocxPath = path.join(tempDir, 'cover_letter.docx');
      fs.writeFileSync(tempDocxPath, processedDocx);

      const tempPdfPath = await convertDocxToPdf(tempDocxPath, tempDir);

      if (!fs.existsSync(tempPdfPath)) {
        throw new Error('Conversion ran but produced no PDF. Check that Word is not showing any dialog boxes.');
      }

      // ── 8. Merge PDFs ────────────────────────────────────────────────────
      const merged = await PDFDocument.create();

      let clPdf;
      try { clPdf = await PDFDocument.load(fs.readFileSync(tempPdfPath)); }
      catch (_) { throw new Error('Could not parse the generated cover letter PDF.'); }

      let invPdf;
      try { invPdf = await PDFDocument.load(req.file.buffer); }
      catch (_) { throw new Error('Could not parse the invoice PDF — it may be encrypted.'); }

      for (const p of await merged.copyPages(clPdf,  clPdf.getPageIndices()))  merged.addPage(p);
      for (const p of await merged.copyPages(invPdf, invPdf.getPageIndices())) merged.addPage(p);

      const finalBuffer = Buffer.from(await merged.save());

      // ── 9. Build filename & send ─────────────────────────────────────────
      const san = s => s.replace(/[\\/:*?"<>|]/g, '-').trim();
      const filename =
        `${san(contractor_name)} Invoice Submittal Package` +
        ` - PE ${san(date_range)} ${san(invoice_number)}.pdf`;

      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.setHeader('X-Download-Filename', encodeURIComponent(filename));
      res.send(finalBuffer);

    } catch (err) {
      console.error('Generation error:', err);
      if (!res.headersSent) res.status(500).json({ error: err.message || 'Unexpected error.' });
    } finally {
      if (tempDir) {
        try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch (_) {}
      }
    }
  }
);

// ─── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  const templates = listTemplates();
  if (templates.length) {
    console.log(`Stored templates (${templates.length}): ${templates.map(t => t.name).join(', ')}`);
  } else {
    console.log('No templates stored yet — upload one via the UI before generating.');
  }
  console.log(`\nInvoice Submittal Generator → http://localhost:${PORT}\n`);
});
