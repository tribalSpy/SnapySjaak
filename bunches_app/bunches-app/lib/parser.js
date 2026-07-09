// parser.js — parse plakte data. Twee modi:
//   1. Ruw formaat (direct uit systeem, quoted CSV met separator lines) → auto-verwerkt
//   2. Excel.bradley formaat (tab-separated, na Excel-macro) → direct parsed
// Auto-detect: rijen die starten met "===" of "Screen" duiden ruw formaat aan.

const REQUIRED_HEADERS = [
  'Klantcode', 'Naam', 'Lengte', 'Cr2', 'Cr3', 'Land',
  'Aantal', 'Som', 'Vertrek', 'Bestel', 'Broncode', 'Tak', 'Hoes', 'Bun',
];

function trim(v) {
  if (v == null) return null;
  if (typeof v === 'string') return v.trim() || null;
  return v;
}

function parseNumber(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;
  const s = String(v).trim().replace(/\./g, '').replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? null : n;
}

function parseInteger(v) {
  const n = parseNumber(v);
  return n == null ? null : Math.trunc(n);
}

// Detect Tak uit naam (10T eerst — voorkomt match op '0T')
function extractTak(name) {
  if (!name) return null;
  const n = String(name);
  for (const p of ['10T', 'x10', '5T', 'x5', '4T', '3T']) {
    if (n.includes(p)) return p;
  }
  return null;
}

// Detect Hoes uit naam
function extractHoes(name) {
  if (!name) return null;
  return /plast/i.test(name) ? 'Plast' : 'Kraft';
}

// Bun filter: alleen rijen met "bun " (spatie erna) in naam
function isBunRow(name) {
  if (!name) return false;
  return /\bbun\s/i.test(String(name));
}

/**
 * Parse ruwe CSV (quoted, semicolon-separated) direct uit het systeem.
 * Voert de hele Excel-macro logica intern uit: filter Bun-rijen, splits Aantal/Som,
 * leidt Tak/Hoes af.
 */
function parseRawCsv(text) {
  const warnings = [];
  const rows = [];
  const lines = text.split(/\r?\n/);
  let rowIdx = 0;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    // Skip separator/decoration
    if (trimmed.startsWith('==') || trimmed.startsWith('"Screen') || trimmed.startsWith('"       "')) continue;

    // Parse quoted CSV met ; separator
    const cells = parseQuotedCsvLine(line);
    if (cells.length < 20) continue;  // te kort → geen data rij

    rowIdx++;
    const naam = (cells[2] || '').trim();
    if (!isBunRow(naam)) continue;  // filter: alleen Bun-rijen

    const broncode = parseInteger(cells[19]);
    if (broncode == null) {
      warnings.push(`Rij ${rowIdx}: geen broncode, overgeslagen`);
      continue;
    }

    // Splits kolom 10 op positie 12: aantal_text + som
    const aantalRaw = String(cells[10] || '');
    const aantalText = aantalRaw.substring(0, 12).trim();
    const som = parseInteger(aantalRaw.substring(12).trim()) ?? 0;

    // Vertrek datum normaliseren dd-mm-yyyy → yyyy-mm-dd voor consistentie
    const vertrekRaw = (cells[14] || '').trim();
    const vertrek = normalizeDate(vertrekRaw);

    rows.push({
      _row: rowIdx,
      klantcode: trim(cells[0]),
      naam,
      lengte: parseInteger(cells[6]),
      cr2: trim(cells[7]),
      cr3: trim(cells[8]),
      land: trim(cells[9]),
      aantal_text: aantalText,
      som,
      vertrek,
      bestel: trim(cells[15]),
      broncode,
      tak: extractTak(naam),
      hoes: extractHoes(naam),
      bun: 'Bun',
    });
  }

  if (rows.length === 0) {
    warnings.push('Geen Bun-rijen gevonden in ruwe data (filter: naam moet "Bun " bevatten)');
  }
  return { rows, warnings, format: 'raw' };
}

// Parse één regel als quoted CSV met ; separator
function parseQuotedCsvLine(line) {
  const cells = [];
  let cur = '';
  let inQuote = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuote && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuote = !inQuote;
      }
    } else if (ch === ';' && !inQuote) {
      cells.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  if (cur !== '' || cells.length > 0) cells.push(cur);
  // Trim trailing empty cells (van de trailing ; in raw format)
  while (cells.length > 0 && cells[cells.length - 1] === '') cells.pop();
  return cells;
}

// Normaliseer datum-formaten naar yyyy-mm-dd (voor consistentie met UI)
function normalizeDate(s) {
  if (!s) return null;
  s = s.trim();
  // dd-mm-yyyy
  let m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  // yyyy-mm-dd (al goed)
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  return s;
}

/**
 * Parse excel.bradley formaat (tab-separated na Excel-macro).
 * Verwacht header rij + data rijen.
 */
function parseBradleyPaste(text) {
  const warnings = [];
  const lines = text.split(/\r?\n/).filter(l => l.trim() !== '');
  if (lines.length < 2) {
    return { rows: [], warnings: ['Geen data rijen gevonden'], format: 'bradley' };
  }

  const sep = lines[0].includes('\t') ? '\t' : (lines[0].includes(';') ? ';' : '\t');
  const headerCells = lines[0].split(sep).map(c => c.trim());
  const colIdx = {};
  headerCells.forEach((h, i) => { colIdx[h.toLowerCase()] = i; });

  const missing = REQUIRED_HEADERS.filter(h => colIdx[h.toLowerCase()] === undefined);
  if (missing.length > 0) {
    warnings.push(`Ontbrekende kolommen: ${missing.join(', ')}. Verwachte volgorde: ${REQUIRED_HEADERS.join(' | ')}`);
    if (headerCells.length < 14) {
      return { rows: [], warnings, format: 'bradley' };
    }
  }

  const get = (cells, name, fallbackIdx) => {
    const idx = colIdx[name.toLowerCase()] ?? fallbackIdx;
    return idx != null ? cells[idx] : undefined;
  };

  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cells = lines[i].split(sep);
    if (cells.every(c => !c || c.trim() === '')) continue;

    const broncode = parseInteger(get(cells, 'Broncode', 10));
    if (broncode == null) {
      warnings.push(`Rij ${i + 1}: ongeldige Broncode, overgeslagen`);
      continue;
    }

    rows.push({
      _row: i + 1,
      klantcode: trim(get(cells, 'Klantcode', 0)),
      naam: trim(get(cells, 'Naam', 1)),
      lengte: parseInteger(get(cells, 'Lengte', 2)),
      cr2: trim(get(cells, 'Cr2', 3)),
      cr3: trim(get(cells, 'Cr3', 4)),
      land: trim(get(cells, 'Land', 5)),
      aantal_text: trim(get(cells, 'Aantal', 6)),
      som: parseInteger(get(cells, 'Som', 7)) ?? 0,
      vertrek: trim(get(cells, 'Vertrek', 8)),
      bestel: trim(get(cells, 'Bestel', 9)),
      broncode,
      tak: trim(get(cells, 'Tak', 11)),
      hoes: trim(get(cells, 'Hoes', 12)),
      bun: trim(get(cells, 'Bun', 13)),
    });
  }

  return { rows, warnings, format: 'bradley' };
}

/**
 * Hoofdfunctie: auto-detect formaat en parse.
 */
function parsePaste(text) {
  if (!text || !text.trim()) {
    return { rows: [], warnings: ['Lege invoer'], format: null };
  }
  // Ruw formaat heeft "====" separator lines of "Screen" headers
  const firstLines = text.split(/\r?\n/).slice(0, 5).join('\n');
  const isRaw = /^={5,}/m.test(firstLines) || /"Screen\s+"/.test(firstLines);
  if (isRaw) {
    return parseRawCsv(text);
  }
  return parseBradleyPaste(text);
}

module.exports = { parsePaste, parseRawCsv, parseBradleyPaste, REQUIRED_HEADERS };
