// seed.js — laad initial master data uit data/articles.json + data/ape.json.
// Stand-alone uitvoerbaar: `node lib/seed.js`

const path = require('path');
const fs = require('fs');
const { createDb } = require('./db');

function seedFromJson(dbPath, articlesPath, apePath) {
  const db = createDb(dbPath);

  if (articlesPath && fs.existsSync(articlesPath)) {
    const articles = JSON.parse(fs.readFileSync(articlesPath, 'utf-8'));
    // Filter ongeldige rijen (broncode of variant_code = null)
    const valid = articles.filter(a => a.broncode != null && a.variant_code);
    db.bulkUpsertArticles(valid);
    console.log(`Articles seeded: ${valid.length}`);
  } else {
    console.warn('articles.json niet gevonden, sla over');
  }

  if (apePath && fs.existsSync(apePath)) {
    const apeRaw = JSON.parse(fs.readFileSync(apePath, 'utf-8'));
    // Dedupe: verzamel alle unieke waardes per naam
    const valuesSeen = new Map();  // name -> Set of values
    const finalValue = new Map();  // name -> last value seen
    for (const e of apeRaw) {
      if (!e.omschrijving || e.ape == null) continue;
      const name = e.omschrijving.trim();
      if (!valuesSeen.has(name)) valuesSeen.set(name, new Set());
      valuesSeen.get(name).add(e.ape);
      finalValue.set(name, e.ape);  // last-write-wins
    }
    const apeEntries = [...finalValue.entries()].map(([omschrijving, ape]) => ({ omschrijving, ape }));
    db.bulkUpsertApe(apeEntries);
    console.log(`APE seeded: ${apeEntries.length}`);
    // Toon echte conflicten (naam met meer dan 1 unieke waarde)
    const conflicts = [];
    for (const [name, values] of valuesSeen) {
      if (values.size > 1) {
        conflicts.push({ name, values: [...values].sort((a, b) => a - b), winner: finalValue.get(name) });
      }
    }
    if (conflicts.length > 0) {
      console.warn(`APE conflicten (${conflicts.length}, laatste waarde wint):`);
      for (const c of conflicts) {
        console.warn(`  ${c.name}: waardes [${c.values.join(', ')}] → ${c.winner}`);
      }
    }
  } else {
    console.warn('ape.json niet gevonden, sla over');
  }
}

// Als rechtstreeks aangeroepen
if (require.main === module) {
  const root = path.join(__dirname, '..');
  seedFromJson(
    process.env.DB_PATH || path.join(root, 'data', 'bunches.sqlite'),
    path.join(root, 'data', 'articles.json'),
    path.join(root, 'data', 'ape.json'),
  );
}

module.exports = { seedFromJson };
