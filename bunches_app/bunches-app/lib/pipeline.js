// pipeline.js — kern van de hele app. Algoritme gevalideerd 48/48 tegen Brad's Excel.
//
// Input:  excel.bradley rijen (uit parser)
// Output: { inlezen, printlijst_plast, printlijst_kraft, yybu_files, warnings }

const VARIANT_TAK_REGEX = /(\d+)/;  // pakt eerste getal uit variant_code (YBU3K → 3)

// Mapping van Hoes + Tak naar YYBU sheet naam
const YYBU_SHEETS = {
  'Plast|3T': 'YYBU3P',
  'Plast|4T': 'YYBU4P',
  'Plast|5T': 'YYBU5P',
  'Plast|10T': 'YYB10P',
  'Kraft|3T': 'YYBU3K',
  'Kraft|4T': 'YYBU4K',
  'Kraft|5T': 'YYBU5K',
  'Kraft|10T': 'YYB10K',
};

function takToInt(tak) {
  if (!tak) return 0;
  const m = String(tak).match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 0;
}

function extractTakFromName(name) {
  if (!name) return null;
  const n = String(name);
  // Volgorde: 10T eerst (anders matcht 0T), dan 5T/4T/3T
  for (const p of ['10T', 'x10', '5T', 'x5', '4T', '3T']) {
    if (n.includes(p)) return p;
  }
  return null;
}

/**
 * Hoofdpipeline. db is een instance van db.js createDb().
 */
function processData(rows, db) {
  const warnings = [];

  // ===== Stap 1: enrich met article master =====
  const enriched = [];
  const missingArticles = new Set();

  for (const r of rows) {
    const article = db.getArticleByBroncode(r.broncode);
    if (!article) {
      missingArticles.add(r.broncode);
      continue;
    }
    enriched.push({ ...r, article });
  }

  if (missingArticles.size > 0) {
    warnings.push({
      type: 'missing_article',
      message: `${missingArticles.size} onbekende broncode(s) — voeg toe in Artikel-beheer`,
      details: [...missingArticles],
    });
  }

  // ===== Stap 2: Inlezen output =====
  // Algoritme (gevalideerd 48/48 tegen Brad's Excel):
  //   a) SUM som per broncode
  //   b) Per broncode: raw_stems = sum_som × tak_int, rond up naar afronden indien gezet
  //   c) Groepeer per (ytype, broncode_inlezen) → sum

  const somPerBroncode = new Map();
  for (const r of enriched) {
    const cur = somPerBroncode.get(r.broncode) || 0;
    somPerBroncode.set(r.broncode, cur + (r.som || 0));
  }

  const inlezenGroups = new Map();  // key: `${ytype}|${broncode_inlezen}` → total
  for (const [broncode, totalSom] of somPerBroncode) {
    const article = db.getArticleByBroncode(broncode);
    if (!article) continue;

    const variant = article.variant_code || '';
    const m = variant.match(VARIANT_TAK_REGEX);
    const takInt = m ? parseInt(m[1], 10) : 0;

    // Bij zonder_tak-artikelen: geen vermenigvuldiging met tak
    // (bv. "Bun Rose. Sup Mix" — 88 bunches blijft 88, geen 88×5)
    const multiplier = article.zonder_tak ? 1 : takInt;
    let rawStems = totalSom * multiplier;
    if (article.afronden) {
      rawStems = Math.ceil(rawStems / article.afronden) * article.afronden;
    }

    const ytype = 'Y' + variant;
    const key = `${ytype}|${article.broncode_inlezen ?? ''}`;
    inlezenGroups.set(key, (inlezenGroups.get(key) || 0) + rawStems);
  }

  const inlezen = [];
  for (const [key, total] of inlezenGroups) {
    if (total <= 0) continue;
    const [variant, broncode_inlezen] = key.split('|');
    inlezen.push({
      variant,
      total,
      broncode_inlezen: broncode_inlezen ? parseInt(broncode_inlezen, 10) : null,
    });
  }

  // ===== Stap 3: Printlijst per Hoes =====
  // Groepeer per artikelnaam, sum som = totaal bossen
  // aantal_eenheden = totaal_bossen / APE (uit BUN HVLH Lijst)

  const missingApe = new Set();

  function buildPrintlijst(hoesFilter) {
    const groups = new Map();
    for (const r of enriched) {
      if (r.hoes !== hoesFilter) continue;
      const name = (r.naam || '').trim();
      if (!name) continue;
      const g = groups.get(name) || { totaal_bossen: 0, lengte: null };
      g.totaal_bossen += (r.som || 0);
      if (g.lengte == null && r.lengte != null) g.lengte = r.lengte;
      groups.set(name, g);
    }

    const result = [];
    for (const [name, g] of groups) {
      if (g.totaal_bossen <= 0) continue;
      const ape = db.getApe(name);
      if (ape == null) {
        missingApe.add(name);
      }
      result.push({
        naam: name,
        tak: extractTakFromName(name),
        lengte: g.lengte,
        aantal_eenheden: ape ? g.totaal_bossen / ape : null,
        ape: ape || 0,
        totaal_bossen: g.totaal_bossen,
      });
    }
    // Sort alfabetisch
    result.sort((a, b) => a.naam.localeCompare(b.naam));
    return result;
  }

  const printlijst_plast = buildPrintlijst('Plast');
  const printlijst_kraft = buildPrintlijst('Kraft');

  if (missingApe.size > 0) {
    warnings.push({
      type: 'missing_ape',
      message: `${missingApe.size} artikel(en) zonder APE — voeg toe in APE-beheer`,
      details: [...missingApe],
    });
  }

  // ===== Stap 4: YYBU* files (per Hoes × Tak combinatie) =====
  // Een file per (hoes, tak), met UNI header + één regel per unieke broncode

  const yybuGroups = new Map();  // sheet_name → Map(broncode → {naam, broncode_inlezen, total_som, total_stems})

  for (const r of enriched) {
    if (!r.hoes || !r.tak) continue;
    const sheet = YYBU_SHEETS[`${r.hoes}|${r.tak}`];
    if (!sheet) continue;

    if (!yybuGroups.has(sheet)) yybuGroups.set(sheet, new Map());
    const sheetMap = yybuGroups.get(sheet);

    const bronEntry = sheetMap.get(r.broncode) || {
      naam: r.naam,
      broncode_inlezen: r.article.broncode_inlezen,
      total_som: 0,
      total_stems: 0,
    };
    bronEntry.total_som += (r.som || 0);
    // Bij zonder_tak-artikelen: som = stems (geen × tak-vermenigvuldiging)
    const multiplier = r.article.zonder_tak ? 1 : takToInt(r.tak);
    bronEntry.total_stems += (r.som || 0) * multiplier;
    sheetMap.set(r.broncode, bronEntry);
  }

  // Build per-sheet structure
  const yybu_files = {};
  for (const [sheet, broncodes] of yybuGroups) {
    const lines = [];
    for (const [bron, data] of broncodes) {
      lines.push({
        broncode: bron,
        broncode_inlezen: data.broncode_inlezen,
        naam: data.naam,
        total_som: data.total_som,
        total_stems: data.total_stems,
      });
    }
    yybu_files[sheet] = lines;
  }

  // ===== Stap 5: samenvatting =====
  const summary = {
    input_rows: rows.length,
    enriched_rows: enriched.length,
    total_som: rows.reduce((s, r) => s + (r.som || 0), 0),
    inlezen_count: inlezen.length,
    inlezen_total_stems: inlezen.reduce((s, r) => s + r.total, 0),
    printlijst_plast_count: printlijst_plast.length,
    printlijst_kraft_count: printlijst_kraft.length,
    yybu_sheet_count: Object.keys(yybu_files).length,
  };

  return {
    summary,
    inlezen,
    printlijst_plast,
    printlijst_kraft,
    yybu_files,
    warnings,
  };
}

module.exports = { processData, YYBU_SHEETS };
