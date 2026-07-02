import { existsSync, promises as fs } from "node:fs";
import path from "node:path";

const REQUIRED_HEADERS = [
  "Klantcode",
  "Naam",
  "Lengte",
  "Cr2",
  "Cr3",
  "Land",
  "Aantal",
  "Som",
  "Vertrek",
  "Bestel",
  "Broncode",
  "Tak",
  "Hoes",
  "Bun",
];

const YYBU_SHEETS = {
  "Plast|3T": "YYBU3P",
  "Plast|4T": "YYBU4P",
  "Plast|5T": "YYBU5P",
  "Plast|10T": "YYB10P",
  "Kraft|3T": "YYBU3K",
  "Kraft|4T": "YYBU4K",
  "Kraft|5T": "YYBU5K",
  "Kraft|10T": "YYB10K",
};

function trimValue(value) {
  if (value == null) {
    return null;
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    return trimmed || null;
  }
  return value;
}

function parseNumber(value) {
  if (value == null || value === "") {
    return null;
  }
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  const normalized = String(value).trim().replace(/\./g, "").replace(",", ".");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseInteger(value) {
  const parsed = parseNumber(value);
  return parsed == null ? null : Math.trunc(parsed);
}

function normalizeDate(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return null;
  }
  let match = raw.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (match) {
    return `${match[3]}-${match[2]}-${match[1]}`;
  }
  match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    return `${match[1]}-${match[2]}-${match[3]}`;
  }
  return raw;
}

function extractTak(name) {
  if (!name) {
    return null;
  }
  const normalized = String(name);
  for (const pattern of ["10T", "x10", "5T", "x5", "4T", "3T"]) {
    if (normalized.includes(pattern)) {
      return pattern;
    }
  }
  return null;
}

function extractHoes(name) {
  if (!name) {
    return null;
  }
  return /plast/i.test(String(name)) ? "Plast" : "Kraft";
}

function isBunRow(name) {
  return /\bbun\s/i.test(String(name || ""));
}

function parseQuotedCsvLine(line) {
  const cells = [];
  let current = "";
  let inQuote = false;
  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];
    if (character === '"') {
      if (inQuote && line[index + 1] === '"') {
        current += '"';
        index += 1;
      } else {
        inQuote = !inQuote;
      }
    } else if (character === ";" && !inQuote) {
      cells.push(current);
      current = "";
    } else {
      current += character;
    }
  }
  if (current !== "" || cells.length > 0) {
    cells.push(current);
  }
  while (cells.length > 0 && cells[cells.length - 1] === "") {
    cells.pop();
  }
  return cells;
}

function parseRawCsv(text) {
  const warnings = [];
  const rows = [];
  const lines = String(text || "").split(/\r?\n/);
  let rowIndex = 0;
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("==") || trimmed.startsWith('"Screen') || trimmed.startsWith('"       "')) {
      continue;
    }
    const cells = parseQuotedCsvLine(line);
    if (cells.length < 20) {
      continue;
    }
    rowIndex += 1;
    const name = String(cells[2] || "").trim();
    if (!isBunRow(name)) {
      continue;
    }
    const broncode = parseInteger(cells[19]);
    if (broncode == null) {
      warnings.push(`Rij ${rowIndex}: geen broncode, overgeslagen`);
      continue;
    }
    const amountRaw = String(cells[10] || "");
    rows.push({
      _row: rowIndex,
      klantcode: trimValue(cells[0]),
      naam: name,
      lengte: parseInteger(cells[6]),
      cr2: trimValue(cells[7]),
      cr3: trimValue(cells[8]),
      land: trimValue(cells[9]),
      aantal_text: amountRaw.substring(0, 12).trim(),
      som: parseInteger(amountRaw.substring(12).trim()) ?? 0,
      vertrek: normalizeDate(cells[14]),
      bestel: trimValue(cells[15]),
      broncode,
      tak: extractTak(name),
      hoes: extractHoes(name),
      bun: "Bun",
    });
  }
  if (!rows.length) {
    warnings.push('Geen Bun-rijen gevonden in ruwe data (filter: naam moet "Bun " bevatten)');
  }
  return { rows, warnings, format: "raw" };
}

function parseBradleyPaste(text) {
  const warnings = [];
  const lines = String(text || "").split(/\r?\n/).filter((line) => line.trim() !== "");
  if (lines.length < 2) {
    return { rows: [], warnings: ["Geen data rijen gevonden"], format: "bradley" };
  }
  const separator = lines[0].includes("\t") ? "\t" : ";";
  const headers = lines[0].split(separator).map((cell) => cell.trim());
  const headerIndexes = {};
  headers.forEach((header, index) => {
    headerIndexes[header.toLowerCase()] = index;
  });
  const missing = REQUIRED_HEADERS.filter((header) => headerIndexes[header.toLowerCase()] == null);
  if (missing.length) {
    warnings.push(`Ontbrekende kolommen: ${missing.join(", ")}. Verwachte volgorde: ${REQUIRED_HEADERS.join(" | ")}`);
    if (headers.length < 14) {
      return { rows: [], warnings, format: "bradley" };
    }
  }
  const getCell = (cells, name, fallbackIndex) => {
    const index = headerIndexes[name.toLowerCase()] ?? fallbackIndex;
    return index == null ? undefined : cells[index];
  };
  const rows = [];
  for (let index = 1; index < lines.length; index += 1) {
    const cells = lines[index].split(separator);
    if (cells.every((cell) => !String(cell || "").trim())) {
      continue;
    }
    const broncode = parseInteger(getCell(cells, "Broncode", 10));
    if (broncode == null) {
      warnings.push(`Rij ${index + 1}: ongeldige Broncode, overgeslagen`);
      continue;
    }
    rows.push({
      _row: index + 1,
      klantcode: trimValue(getCell(cells, "Klantcode", 0)),
      naam: trimValue(getCell(cells, "Naam", 1)),
      lengte: parseInteger(getCell(cells, "Lengte", 2)),
      cr2: trimValue(getCell(cells, "Cr2", 3)),
      cr3: trimValue(getCell(cells, "Cr3", 4)),
      land: trimValue(getCell(cells, "Land", 5)),
      aantal_text: trimValue(getCell(cells, "Aantal", 6)),
      som: parseInteger(getCell(cells, "Som", 7)) ?? 0,
      vertrek: normalizeDate(getCell(cells, "Vertrek", 8)),
      bestel: trimValue(getCell(cells, "Bestel", 9)),
      broncode,
      tak: trimValue(getCell(cells, "Tak", 11)),
      hoes: trimValue(getCell(cells, "Hoes", 12)),
      bun: trimValue(getCell(cells, "Bun", 13)),
    });
  }
  return { rows, warnings, format: "bradley" };
}

function parsePaste(text) {
  if (!String(text || "").trim()) {
    return { rows: [], warnings: ["Lege invoer"], format: null };
  }
  const firstLines = String(text).split(/\r?\n/).slice(0, 5).join("\n");
  const isRaw = /^={5,}/m.test(firstLines) || /"Screen\s+"/.test(firstLines);
  return isRaw ? parseRawCsv(text) : parseBradleyPaste(text);
}

function normalizeArticleEntry(entry) {
  return {
    broncode: parseInteger(entry?.broncode),
    omschrijving: String(entry?.omschrijving || "").trim(),
    variant_code: String(entry?.variant_code || "").trim(),
    broncode_inlezen: parseInteger(entry?.broncode_inlezen),
    afronden: parseInteger(entry?.afronden),
    zonder_tak: Boolean(entry?.zonder_tak),
    active: entry?.active !== false,
    updated_at: String(entry?.updated_at || new Date().toISOString()),
  };
}

function normalizeApeEntry(entry) {
  return {
    omschrijving: String(entry?.omschrijving || "").trim(),
    ape: parseInteger(entry?.ape) ?? 0,
    updated_at: String(entry?.updated_at || new Date().toISOString()),
  };
}

function normalizeRunEntry(entry) {
  return {
    id: parseInteger(entry?.id) ?? 0,
    created_at: String(entry?.created_at || new Date().toISOString()),
    user: String(entry?.user || ""),
    row_count: parseInteger(entry?.row_count) ?? 0,
    total_som: parseInteger(entry?.total_som) ?? 0,
    status: String(entry?.status || "ok"),
    vertrek_datum: String(entry?.vertrek_datum || ""),
    label: entry?.label == null ? "" : String(entry.label),
    warnings: Array.isArray(entry?.warnings) ? entry.warnings : [],
    result: entry?.result && typeof entry.result === "object" ? entry.result : null,
  };
}

function csvCell(value) {
  if (value == null) {
    return "";
  }
  const stringValue = String(value);
  if (/[",;\n\r]/.test(stringValue)) {
    return `"${stringValue.replace(/"/g, '""')}"`;
  }
  return stringValue;
}

function formatDateNL(isoDate) {
  const match = String(isoDate || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  return match ? `${match[3]}-${match[2]}-${match[1]}` : String(isoDate || "");
}

function formatAmount(value) {
  if (value == null) {
    return "";
  }
  return Math.abs(value - Math.round(value)) < 0.001 ? String(Math.round(value)) : Number(value).toFixed(1);
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[character]));
}

function generateInlezenCsv(inlezen) {
  const sorted = [...inlezen].sort((left, right) => {
    const leftVariant = String(left.variant || "");
    const rightVariant = String(right.variant || "");
    if (leftVariant !== rightVariant) {
      return leftVariant.localeCompare(rightVariant);
    }
    return (left.broncode_inlezen || 0) - (right.broncode_inlezen || 0);
  });
  return `${sorted.map((row) => [`${row.variant} `, row.total, row.broncode_inlezen, 0, 0, 0, 0].join(";")).join("\r\n")}\r\n`;
}

function generateUniFile(sheetName, lines, dateString) {
  const rows = [
    "UNI_VERSION:3.6.34",
    `UNI_CUST_ID:${sheetName}`,
    `UNI_DATE:${formatDateNL(dateString) || "DATUM VERTREK"}`,
    "UNI_FTERM:;",
    "UNI_STANDING:J",
    "UNI_HEADER:",
    ["int_item_number", "group", "description", "remark", "amount"].join(";"),
  ];
  for (const line of lines) {
    rows.push([
      line.broncode,
      line.broncode_inlezen ?? "",
      String(line.naam || "").trim(),
      "",
      line.total_stems,
    ].join(";"));
  }
  return `${rows.join("\n")}\n`;
}

function generatePrintlijstHtml(title, items, dateString) {
  const displayDate = formatDateNL(dateString) || dateString || "";
  const rows = items.map((item) => `
    <tr>
      <td class="check"><input type="checkbox"></td>
      <td>${escapeHtml(item.naam)}</td>
      <td class="num">${escapeHtml(item.tak || "")}</td>
      <td class="num">${item.lengte ?? ""}</td>
      <td class="num">${item.aantal_eenheden != null ? formatAmount(item.aantal_eenheden) : "?"}</td>
      <td class="num">${item.ape || "?"}</td>
      <td class="num">${item.totaal_bossen ?? ""}</td>
    </tr>
  `).join("");
  return `<!doctype html>
<html lang="nl">
<head>
  <meta charset="utf-8">
  <title>${escapeHtml(title)}</title>
  <style>
    body { font-family: Arial, sans-serif; font-size: 11pt; margin: 20px; color: #132238; }
    h1 { font-size: 14pt; margin: 0 0 10px; }
    .meta { color: #5c708a; font-size: 10pt; margin-bottom: 16px; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border: 1px solid #9cb0c7; padding: 6px 8px; text-align: left; }
    th { background: #eef4fb; }
    .num { text-align: right; }
    .check { width: 34px; text-align: center; }
    .check input { width: 18px; height: 18px; }
    @media print {
      body { margin: 10mm; }
    }
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  <div class="meta">${escapeHtml(displayDate)} - ${items.length} regels</div>
  <table>
    <thead>
      <tr>
        <th class="check">OK</th>
        <th>Naam</th>
        <th class="num">Tak</th>
        <th class="num">Lengte</th>
        <th class="num">Aantal</th>
        <th class="num">APE</th>
        <th class="num">Totaal bossen</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>
</body>
</html>`;
}

function processData(rows, state) {
  const warnings = [];
  const articlesByBroncode = new Map(state.articles.filter((article) => article.active).map((article) => [article.broncode, article]));
  const apeByName = new Map(state.ape.map((entry) => [entry.omschrijving, entry.ape]));
  const enriched = [];
  const missingArticles = new Set();
  for (const row of rows) {
    const article = articlesByBroncode.get(row.broncode);
    if (!article) {
      missingArticles.add(row.broncode);
      continue;
    }
    enriched.push({ ...row, article });
  }
  if (missingArticles.size) {
    warnings.push({
      type: "missing_article",
      message: `${missingArticles.size} onbekende broncode(s) - voeg toe in Artikel-beheer`,
      details: [...missingArticles].sort((left, right) => left - right),
    });
  }

  const somPerBroncode = new Map();
  for (const row of enriched) {
    somPerBroncode.set(row.broncode, (somPerBroncode.get(row.broncode) || 0) + (row.som || 0));
  }

  const inlezenGroups = new Map();
  for (const [broncode, totalSom] of somPerBroncode.entries()) {
    const article = articlesByBroncode.get(broncode);
    if (!article) {
      continue;
    }
    const match = String(article.variant_code || "").match(/(\d+)/);
    const takInt = match ? parseInteger(match[1]) || 0 : 0;
    const multiplier = article.zonder_tak ? 1 : takInt;
    let totalStems = totalSom * multiplier;
    if (article.afronden) {
      totalStems = Math.ceil(totalStems / article.afronden) * article.afronden;
    }
    const key = `Y${article.variant_code}|${article.broncode_inlezen ?? ""}`;
    inlezenGroups.set(key, (inlezenGroups.get(key) || 0) + totalStems);
  }

  const inlezen = [...inlezenGroups.entries()].map(([key, total]) => {
    const [variant, broncodeInlezen] = key.split("|");
    return {
      variant,
      total,
      broncode_inlezen: broncodeInlezen ? parseInteger(broncodeInlezen) : null,
    };
  }).filter((row) => row.total > 0);

  const missingApe = new Set();
  function buildPrintlijst(hoesFilter) {
    const groups = new Map();
    for (const row of enriched) {
      if (row.hoes !== hoesFilter) {
        continue;
      }
      const name = String(row.naam || "").trim();
      if (!name) {
        continue;
      }
      const current = groups.get(name) || { totaal_bossen: 0, lengte: null };
      current.totaal_bossen += row.som || 0;
      if (current.lengte == null && row.lengte != null) {
        current.lengte = row.lengte;
      }
      groups.set(name, current);
    }
    const result = [];
    for (const [name, current] of groups.entries()) {
      if (current.totaal_bossen <= 0) {
        continue;
      }
      const ape = apeByName.get(name);
      if (ape == null) {
        missingApe.add(name);
      }
      result.push({
        naam: name,
        tak: extractTak(name),
        lengte: current.lengte,
        aantal_eenheden: ape ? current.totaal_bossen / ape : null,
        ape: ape || 0,
        totaal_bossen: current.totaal_bossen,
      });
    }
    return result.sort((left, right) => left.naam.localeCompare(right.naam));
  }

  const printlijst_plast = buildPrintlijst("Plast");
  const printlijst_kraft = buildPrintlijst("Kraft");
  if (missingApe.size) {
    warnings.push({
      type: "missing_ape",
      message: `${missingApe.size} artikel(en) zonder APE - voeg toe in APE-beheer`,
      details: [...missingApe].sort((left, right) => left.localeCompare(right)),
    });
  }

  const yybuGroups = new Map();
  for (const row of enriched) {
    if (!row.hoes || !row.tak) {
      continue;
    }
    const sheet = YYBU_SHEETS[`${row.hoes}|${row.tak}`];
    if (!sheet) {
      continue;
    }
    if (!yybuGroups.has(sheet)) {
      yybuGroups.set(sheet, new Map());
    }
    const byBroncode = yybuGroups.get(sheet);
    const current = byBroncode.get(row.broncode) || {
      naam: row.naam,
      broncode_inlezen: row.article.broncode_inlezen,
      total_som: 0,
      total_stems: 0,
    };
    const takInt = parseInteger(String(row.tak).match(/(\d+)/)?.[1]) || 0;
    const multiplier = row.article.zonder_tak ? 1 : takInt;
    current.total_som += row.som || 0;
    current.total_stems += (row.som || 0) * multiplier;
    byBroncode.set(row.broncode, current);
  }

  const yybu_files = {};
  for (const [sheet, byBroncode] of yybuGroups.entries()) {
    yybu_files[sheet] = [...byBroncode.entries()].map(([broncode, item]) => ({
      broncode,
      broncode_inlezen: item.broncode_inlezen,
      naam: item.naam,
      total_som: item.total_som,
      total_stems: item.total_stems,
    }));
  }

  const takOrder = ["3T", "4T", "5T", "10T"];
  const plastTaks = [...new Set(printlijst_plast.map((item) => item.tak).filter(Boolean))].sort((left, right) => takOrder.indexOf(left) - takOrder.indexOf(right));
  const kraftTaks = [...new Set(printlijst_kraft.map((item) => item.tak).filter(Boolean))].sort((left, right) => takOrder.indexOf(left) - takOrder.indexOf(right));

  return {
    summary: {
      input_rows: rows.length,
      enriched_rows: enriched.length,
      total_som: rows.reduce((sum, row) => sum + (row.som || 0), 0),
      inlezen_count: inlezen.length,
      inlezen_total_stems: inlezen.reduce((sum, row) => sum + (row.total || 0), 0),
      printlijst_plast_count: printlijst_plast.length,
      printlijst_kraft_count: printlijst_kraft.length,
      yybu_sheet_count: Object.keys(yybu_files).length,
    },
    inlezen,
    printlijst_plast,
    printlijst_kraft,
    yybu_files,
    warnings,
    availableTaks: {
      Plast: takOrder.filter((tak) => plastTaks.includes(tak)),
      Kraft: takOrder.filter((tak) => kraftTaks.includes(tak)),
    },
  };
}

export function createBunchesService({ statePath, seedDir }) {
  async function readJson(filePath, fallback) {
    try {
      return JSON.parse(await fs.readFile(filePath, "utf8"));
    } catch {
      return fallback;
    }
  }

  async function writeJson(filePath, payload) {
    await fs.mkdir(path.dirname(filePath), { recursive: true });
    await fs.writeFile(filePath, JSON.stringify(payload, null, 2), "utf8");
  }

  async function readSeedFile(fileName) {
    const filePath = path.join(seedDir, fileName);
    if (!existsSync(filePath)) {
      return [];
    }
    return readJson(filePath, []);
  }

  async function createSeededState() {
    const rawArticles = await readSeedFile("articles.json");
    const rawApe = await readSeedFile("ape.json");
    const articleMap = new Map();
    for (const entry of rawArticles) {
      const normalized = normalizeArticleEntry(entry);
      if (normalized.broncode == null || !normalized.variant_code) {
        continue;
      }
      articleMap.set(normalized.broncode, normalized);
    }
    const apeMap = new Map();
    for (const entry of rawApe) {
      const normalized = normalizeApeEntry(entry);
      if (!normalized.omschrijving) {
        continue;
      }
      apeMap.set(normalized.omschrijving, normalized);
    }
    return {
      articles: [...articleMap.values()].sort((left, right) => left.omschrijving.localeCompare(right.omschrijving)),
      ape: [...apeMap.values()].sort((left, right) => left.omschrijving.localeCompare(right.omschrijving)),
      runs: [],
      next_run_id: 1,
      seeded_at: new Date().toISOString(),
    };
  }

  async function readState() {
    const payload = await readJson(statePath, null);
    if (!payload) {
      const seeded = await createSeededState();
      await writeJson(statePath, seeded);
      return seeded;
    }
    return {
      articles: Array.isArray(payload.articles) ? payload.articles.map(normalizeArticleEntry).filter((entry) => entry.broncode != null && entry.omschrijving && entry.variant_code) : [],
      ape: Array.isArray(payload.ape) ? payload.ape.map(normalizeApeEntry).filter((entry) => entry.omschrijving) : [],
      runs: Array.isArray(payload.runs) ? payload.runs.map(normalizeRunEntry).filter((entry) => entry.id > 0) : [],
      next_run_id: parseInteger(payload.next_run_id) ?? 1,
      seeded_at: String(payload.seeded_at || ""),
    };
  }

  async function writeState(state) {
    await writeJson(statePath, {
      articles: state.articles,
      ape: state.ape,
      runs: state.runs,
      next_run_id: state.next_run_id,
      seeded_at: state.seeded_at,
    });
  }

  function summarizeRun(run) {
    return {
      id: run.id,
      created_at: run.created_at,
      user: run.user,
      row_count: run.row_count,
      total_som: run.total_som,
      status: run.status,
      vertrek_datum: run.vertrek_datum,
      label: run.label,
      warnings: run.warnings,
      result: run.result,
    };
  }

  async function getAppState() {
    const state = await readState();
    return {
      counts: {
        articles: state.articles.filter((entry) => entry.active).length,
        ape: state.ape.length,
        runs: state.runs.length,
      },
      articles: state.articles.sort((left, right) => left.omschrijving.localeCompare(right.omschrijving)),
      ape: state.ape.sort((left, right) => left.omschrijving.localeCompare(right.omschrijving)),
      runs: [...state.runs]
        .sort((left, right) => right.id - left.id)
        .slice(0, 100)
        .map(summarizeRun),
    };
  }

  async function processImport({ pasteText, vertrekDatum, label, user }) {
    const state = await readState();
    const parsed = parsePaste(pasteText);
    if (!parsed.rows.length) {
      throw new Error(parsed.warnings.join("; ") || "Geen bruikbare rijen gevonden");
    }
    const result = processData(parsed.rows, state);
    const warnings = [
      ...parsed.warnings.map((message) => ({ type: "parse", message })),
      ...result.warnings,
    ];
    const normalizedDate = normalizeDate(vertrekDatum) || new Date().toISOString().slice(0, 10);
    const run = normalizeRunEntry({
      id: state.next_run_id,
      created_at: new Date().toISOString(),
      user,
      row_count: parsed.rows.length,
      total_som: result.summary.total_som,
      status: warnings.length ? "warnings" : "ok",
      vertrek_datum: normalizedDate,
      label: String(label || "").trim(),
      warnings,
      result: {
        ...result,
        warnings,
        vertrekDatum: normalizedDate,
        generatedAt: new Date().toISOString(),
      },
    });
    state.next_run_id += 1;
    state.runs.unshift(run);
    await writeState(state);
    return summarizeRun(run);
  }

  async function updateRunDate(runId, vertrekDatum) {
    const state = await readState();
    const run = state.runs.find((entry) => entry.id === runId);
    if (!run) {
      throw new Error("Run not found");
    }
    const normalizedDate = normalizeDate(vertrekDatum);
    if (!normalizedDate) {
      throw new Error("Date is required");
    }
    run.vertrek_datum = normalizedDate;
    if (run.result) {
      run.result.vertrekDatum = normalizedDate;
    }
    await writeState(state);
    return summarizeRun(run);
  }

  async function updateRunLabel(runId, label) {
    const state = await readState();
    const run = state.runs.find((entry) => entry.id === runId);
    if (!run) {
      throw new Error("Run not found");
    }
    run.label = String(label || "").trim();
    await writeState(state);
    return summarizeRun(run);
  }

  async function deleteRun(runId) {
    const state = await readState();
    const nextRuns = state.runs.filter((entry) => entry.id !== runId);
    if (nextRuns.length === state.runs.length) {
      throw new Error("Run not found");
    }
    state.runs = nextRuns;
    await writeState(state);
  }

  async function upsertArticle(article) {
    const state = await readState();
    const normalized = normalizeArticleEntry(article);
    if (normalized.broncode == null || !normalized.omschrijving || !normalized.variant_code) {
      throw new Error("Broncode, omschrijving, and variant are required");
    }
    normalized.active = true;
    normalized.updated_at = new Date().toISOString();
    const index = state.articles.findIndex((entry) => entry.broncode === normalized.broncode);
    if (index >= 0) {
      state.articles[index] = { ...state.articles[index], ...normalized };
    } else {
      state.articles.push(normalized);
    }
    await writeState(state);
    return normalized;
  }

  async function deactivateArticle(broncode) {
    const state = await readState();
    const article = state.articles.find((entry) => entry.broncode === broncode);
    if (!article) {
      throw new Error("Article not found");
    }
    article.active = false;
    article.updated_at = new Date().toISOString();
    await writeState(state);
  }

  async function bulkSetZonderTak(broncodes, value) {
    const state = await readState();
    const found = [];
    const missing = [];
    for (const broncode of broncodes) {
      const article = state.articles.find((entry) => entry.broncode === broncode);
      if (!article) {
        missing.push(broncode);
        continue;
      }
      article.zonder_tak = Boolean(value);
      article.updated_at = new Date().toISOString();
      found.push(broncode);
    }
    await writeState(state);
    return { updated: found.length, missing };
  }

  async function upsertApe(entry) {
    const state = await readState();
    const normalized = normalizeApeEntry(entry);
    if (!normalized.omschrijving || normalized.ape <= 0) {
      throw new Error("Omschrijving and APE are required");
    }
    normalized.updated_at = new Date().toISOString();
    const index = state.ape.findIndex((item) => item.omschrijving === normalized.omschrijving);
    if (index >= 0) {
      state.ape[index] = normalized;
    } else {
      state.ape.push(normalized);
    }
    await writeState(state);
    return normalized;
  }

  async function deleteApe(omschrijving) {
    const state = await readState();
    const nextApe = state.ape.filter((entry) => entry.omschrijving !== omschrijving);
    if (nextApe.length === state.ape.length) {
      throw new Error("APE entry not found");
    }
    state.ape = nextApe;
    await writeState(state);
  }

  async function findRun(runId) {
    const state = await readState();
    const run = state.runs.find((entry) => entry.id === runId);
    if (!run || !run.result) {
      return null;
    }
    return run;
  }

  async function downloadFile(runId, kind, sheet) {
    const run = await findRun(runId);
    if (!run) {
      throw new Error("Run not found");
    }
    if (kind === "inlezen") {
      return {
        contentType: "text/csv; charset=utf-8",
        filename: "Inlezen.csv",
        body: generateInlezenCsv(run.result.inlezen || []),
      };
    }
    if (kind === "yybu") {
      const sheetLines = run.result.yybu_files?.[sheet];
      if (!sheetLines) {
        throw new Error("YYBU sheet not found");
      }
      return {
        contentType: "text/csv; charset=utf-8",
        filename: `${sheet}.csv`,
        body: generateUniFile(sheet, sheetLines, run.result.vertrekDatum || run.vertrek_datum),
      };
    }
    throw new Error("Unknown download");
  }

  async function renderPrintlijst(runId, hoes, tak) {
    const run = await findRun(runId);
    if (!run) {
      throw new Error("Run not found");
    }
    const normalizedHoes = String(hoes || "").toLowerCase();
    const isPlast = normalizedHoes === "plast";
    const isKraft = normalizedHoes === "kraft";
    if (!isPlast && !isKraft) {
      throw new Error("Onbekende hoes");
    }
    let items = isPlast ? [...(run.result.printlijst_plast || [])] : [...(run.result.printlijst_kraft || [])];
    if (tak) {
      items = items.filter((item) => item.tak === tak);
    }
    const title = tak
      ? `Printlijst ${isPlast ? "Plastic" : "Kraft"} ${tak}`
      : `Printlijst ${isPlast ? "Plastic" : "Kraft"}`;
    return generatePrintlijstHtml(title, items, run.result.vertrekDatum || run.vertrek_datum);
  }

  return {
    getAppState,
    processImport,
    updateRunDate,
    updateRunLabel,
    deleteRun,
    upsertArticle,
    deactivateArticle,
    bulkSetZonderTak,
    upsertApe,
    deleteApe,
    downloadFile,
    renderPrintlijst,
  };
}
