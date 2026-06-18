/**
 * Sticker generator — Express server
 *
 * Endpoints:
 *   POST /api/upload    — upload halindeling .xlsx, retourneert beschikbare prefixen
 *   POST /api/generate  — genereer PDF op basis van geselecteerde prefixen
 *
 * Start:  npm install && npm start  → http://localhost:3000
 */
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const PDFDocument = require('pdfkit');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 },
});

const sessions = new Map();
const SESSION_TTL_MS = 30 * 60 * 1000;

function cleanupSessions() {
  const now = Date.now();
  for (const [k, v] of sessions) {
    if (now - v.ts > SESSION_TTL_MS) sessions.delete(k);
  }
}

app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const PT_PER_CM = 72 / 2.54;

function locPrefix(loc) {
  if (!loc) return '';
  return loc.slice(0, 2);
}

function customerPrefix(code) {
  if (!code) return '';
  if (/^\d/.test(code)) return code.slice(0, 3);
  return code.slice(0, 2);
}

function stripLeadingG(loc) {
  if (loc && loc[0] && loc[0].toLowerCase() === 'g') {
    return loc.slice(1).trimStart();
  }
  return loc;
}

function parseHalindeling(buffer) {
  const wb = xlsx.read(buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames.includes('ERP_PASTE') ? 'ERP_PASTE'
                   : wb.SheetNames.includes('Blad1')    ? 'Blad1'
                   : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: null });

  const data = [];
  let current = null;
  for (const row of rows) {
    const loc = row[0];
    const klant = row[1];
    let isHeader = false;
    if (typeof loc === 'string') {
      const s = loc.trim();
      if (!s || s.startsWith('Hal:') || s.startsWith('---') || s.startsWith('#') || s === 'Locatie') {
        isHeader = true;
      } else {
        current = s;
      }
    }
    if (isHeader) continue;
    if (klant && typeof klant === 'string' && klant.trim()) {
      if (current) data.push({ location: current, customer: klant.trim() });
    }
  }
  return data;
}

app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Geen bestand ontvangen' });
    const data = parseHalindeling(req.file.buffer);
    if (data.length === 0) {
      return res.status(400).json({ error: 'Geen geldige halindeling-data gevonden' });
    }

    const locSet = new Set(data.map(d => locPrefix(d.location)).filter(Boolean));
    const custSet = new Set(data.map(d => customerPrefix(d.customer)).filter(Boolean));

    const custByLoc = {};
    for (const { location, customer } of data) {
      const lp = locPrefix(location);
      const cp = customerPrefix(customer);
      if (!custByLoc[lp]) custByLoc[lp] = new Set();
      custByLoc[lp].add(cp);
    }
    const custByLocObj = {};
    for (const lp of Object.keys(custByLoc)) {
      custByLocObj[lp] = Array.from(custByLoc[lp]).sort();
    }

    const id = crypto.randomBytes(8).toString('hex');
    sessions.set(id, { data, ts: Date.now() });
    cleanupSessions();

    res.json({
      id,
      locPrefixes: Array.from(locSet).sort(),
      custPrefixes: Array.from(custSet).sort(),
      custByLoc: custByLocObj,
      totalRows: data.length,
    });
  } catch (e) {
    console.error('upload error', e);
    res.status(500).json({ error: e.message || 'Upload mislukt' });
  }
});

app.post('/api/generate', (req, res) => {
  try {
    const { id, locPrefixes: chosenLoc = [], custPrefixes: chosenCust = [] } = req.body || {};
    if (!id) return res.status(400).json({ error: 'Geen sessie-id meegegeven' });
    const session = sessions.get(id);
    if (!session) return res.status(404).json({ error: 'Sessie verlopen — upload het bestand opnieuw' });

    const { data } = session;

    const filtered = data.filter(d => {
      const lp = locPrefix(d.location);
      const cp = customerPrefix(d.customer);
      const locOk = chosenLoc.length === 0 || chosenLoc.includes(lp);
      const custOk = chosenCust.length === 0 || chosenCust.includes(cp);
      return locOk && custOk;
    });

    const seen = new Set();
    const unique = [];
    for (const d of filtered) {
      if (seen.has(d.customer)) continue;
      seen.add(d.customer);
      unique.push(d);
    }

    if (unique.length === 0) {
      return res.status(400).json({ error: 'Geen klanten gevonden voor deze filters' });
    }

    const pageW = 10 * PT_PER_CM;
    const pageH = 15 * PT_PER_CM;
    const margin = 0.4 * PT_PER_CM;
    const gap = 0.4 * PT_PER_CM;
    const locRatio = 4;

    const doc = new PDFDocument({ size: [pageW, pageH], margin: 0, autoFirstPage: false });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="stickers.pdf"');
    doc.pipe(res);

    doc.font('Helvetica-Bold');

    // Belangrijk: gebruik geen pdfkit.heightOfString() omdat die wikkelt op
    // basis van beschikbare breedte (die verandert na text()-aanroepen).
    // Helvetica em-hoogte ~ fontSize, dus h = size benadering is veilig.
    const fitFontSize = (text, maxW, maxH) => {
      let size = 1;
      while (size < 600) {
        doc.fontSize(size + 1);
        const w = doc.widthOfString(text);
        const h = size + 1;
        if (w > maxW * 0.97 || h > maxH * 0.97) break;
        size++;
      }
      return size;
    };

    for (const { location, customer } of unique) {
      const loc = stripLeadingG(location);
      doc.addPage({ size: [pageW, pageH], margin: 0 });
      doc.save();
      doc.translate(0, pageH);
      doc.rotate(-90);

      const W = pageH;  // lange as
      const H = pageW;  // korte as
      const innerW = W - 2 * margin;
      const avail = H - 2 * margin - gap;
      const cliH = avail / (locRatio + 1);
      const locH = (avail * locRatio) / (locRatio + 1);

      // Klantcode (onder in gedraaid systeem)
      const cliSize = fitFontSize(customer, innerW, cliH);
      doc.fontSize(cliSize);
      const cliW = doc.widthOfString(customer);
      const cliTextH = cliSize * 0.72; // cap-hoogte voor centreren
      const cliY = H - margin - cliH + (cliH - cliTextH) / 2 - cliSize * 0.1;
      doc.text(customer, (W - cliW) / 2, cliY, { lineBreak: false });

      // Locatie (boven, groot — target = klant * 4)
      const targetLocSize = cliSize * locRatio;
      const maxLocSize = fitFontSize(loc, innerW, locH);
      const locSize = Math.min(targetLocSize, maxLocSize);
      doc.fontSize(locSize);
      const lW = doc.widthOfString(loc);
      const locTextH = locSize * 0.72;
      const locY = margin + (locH - locTextH) / 2 - locSize * 0.1;
      doc.text(loc, (W - lW) / 2, locY, { lineBreak: false });

      doc.restore();
    }

    doc.end();
  } catch (e) {
    console.error('generate error', e);
    if (!res.headersSent) {
      res.status(500).json({ error: e.message || 'Generatie mislukt' });
    } else {
      res.end();
    }
  }
});

app.listen(PORT, () => {
  console.log('Sticker app draait op http://localhost:' + PORT);
});
