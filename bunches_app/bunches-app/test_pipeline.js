// test_pipeline.js — valideer JS-pipeline tegen verwachte output (zonder SQLite)
const fs = require('fs');
const path = require('path');
const { processData } = require('./lib/pipeline');
const { parsePaste, REQUIRED_HEADERS } = require('./lib/parser');

// Mock db: gewone JS Maps in plaats van SQLite
const articles = JSON.parse(fs.readFileSync(path.join(__dirname, 'data', 'articles.json'), 'utf-8'));
const apeData = JSON.parse(fs.readFileSync(path.join(__dirname, 'data', 'ape.json'), 'utf-8'));

const articleByBron = new Map();
for (const a of articles) {
  if (a.broncode != null) articleByBron.set(a.broncode, a);
}

const apeMap = new Map();
for (const e of apeData) {
  if (e.omschrijving) apeMap.set(e.omschrijving.trim(), e.ape);
}

const mockDb = {
  getArticleByBroncode: (b) => articleByBron.get(b) || null,
  getApe: (name) => apeMap.get(name) ?? null,
};

// === Laad input bradley data en converteer naar pipeline-formaat ===
const inputs = JSON.parse(fs.readFileSync(path.join(__dirname, '../extracted/input_bradley.json'), 'utf-8'));

// inputs zijn al gestructureerd (niet uit paste tekst), dus geef direct door
const result = processData(inputs, mockDb);

console.log('=== Summary ===');
console.log(JSON.stringify(result.summary, null, 2));

console.log('\n=== Warnings ===');
for (const w of result.warnings) {
  console.log(`  ${w.type}: ${w.message} (${w.details?.length || 0} details)`);
}

// === Vergelijk met verwachte output ===
const expected = JSON.parse(fs.readFileSync(path.join(__dirname, '../extracted/expected_inlezen.json'), 'utf-8'));

function key(r) { return `${r.variant.trim()}|${r.broncode_inlezen}|${r.total}`; }

const computedSet = new Set(result.inlezen.map(key));
const expectedSet = new Set(expected.map(key));

const match = [...computedSet].filter(k => expectedSet.has(k)).length;
const onlyInComputed = [...computedSet].filter(k => !expectedSet.has(k));
const onlyInExpected = [...expectedSet].filter(k => !computedSet.has(k));

console.log(`\n=== Inlezen validatie ===`);
console.log(`Computed: ${computedSet.size}, Expected: ${expectedSet.size}`);
console.log(`Match: ${match}/${expectedSet.size}`);
if (onlyInComputed.length) {
  console.log(`Only in computed (${onlyInComputed.length}):`);
  for (const k of onlyInComputed.slice(0, 5)) console.log(`  + ${k}`);
}
if (onlyInExpected.length) {
  console.log(`Only in expected (${onlyInExpected.length}):`);
  for (const k of onlyInExpected.slice(0, 5)) console.log(`  - ${k}`);
}

// Printlijst validatie
const expPlast = JSON.parse(fs.readFileSync(path.join(__dirname, '../extracted/expected_printlijst_plast.json'), 'utf-8'));
const expKraft = JSON.parse(fs.readFileSync(path.join(__dirname, '../extracted/expected_printlijst_kraft.json'), 'utf-8'));

function plKey(r) { return `${r.naam}|${r.tak}|${r.lengte}|${r.ape}|${r.totaal_bossen ?? r.totaal_stuks}`; }

function comparePrintlijst(label, computed, expected) {
  const c = new Set(computed.map(plKey));
  const e = new Set(expected.map(r => `${r.naam}|${r.tak}|${r.lengte}|${r.ape}|${r.totaal_stuks}`));
  const m = [...c].filter(k => e.has(k)).length;
  console.log(`\n=== Printlijst ${label} ===`);
  console.log(`Match: ${m}/${e.size}`);
  const missing = [...e].filter(k => !c.has(k));
  if (missing.length) {
    console.log(`Missing (${missing.length}):`);
    for (const k of missing.slice(0, 5)) console.log(`  - ${k}`);
  }
}

comparePrintlijst('Plast', result.printlijst_plast, expPlast);
comparePrintlijst('Kraft', result.printlijst_kraft, expKraft);

// YYBU sample
console.log('\n=== YYBU files ===');
for (const sheet of Object.keys(result.yybu_files).sort()) {
  console.log(`  ${sheet}: ${result.yybu_files[sheet].length} broncodes`);
}

// Test parser ook
console.log('\n=== Parser test (tab-separated paste) ===');
const sampleHeader = REQUIRED_HEADERS.join('\t');
const sampleRow = ['5FJ17D', 'Bun Als. Kenia Mix Plast 3T', '60', '25', 'mix', 'KE', '1 x 20 =', '20', '2026-07-01', '2026-05-28', '2646428', '3T', 'Plast', 'Bun'].join('\t');
const samplePaste = sampleHeader + '\n' + sampleRow;
const { rows, warnings } = parsePaste(samplePaste);
console.log(`Parsed rows: ${rows.length}, warnings: ${warnings.length}`);
if (rows.length) console.log('First row:', rows[0]);
