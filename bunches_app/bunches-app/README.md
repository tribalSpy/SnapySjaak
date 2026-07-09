# Bunches App

Node app die `excel.bradley` data (of ruwe systeem-export) omzet naar:
- **Inlezen.csv** — voor systeem-import (geen header, rij 4-7 = 0, alfabetisch gesorteerd)
- **YYBU*.csv** — Universal/UNI-formaat per (Hoes × Tak)
- **Printlijst Plast / Kraft** — volledig én per Tak, met checkbox per regel voor het picken

## Input formaten (auto-detect)

**Ruwe systeem-export**: het CSV bestand direct uit het systeem (met `====` separator lines en quoted velden). App:
- Skipt separator/header rijen
- Filtert alleen "Bun" rijen (naam bevat "bun ")
- Splitst Aantal/Som op positie 12
- Leidt Tak/Hoes af uit de naam
- Mapt naar excel.bradley formaat

**Excel.bradley**: tab-separated, kolommen: Klantcode | Naam | Lengte | Cr2 | Cr3 | Land | Aantal | Som | Vertrek | Bestel | Broncode | Tak | Hoes | Bun

## Validatie

Algoritme gevalideerd 48/48 tegen originele Excel. Ruwe parser reproduceert exact wat de Excel-macro doet.

## Structuur

```
bunches-app/
├── package.json
├── index.js          ← standalone entry + module exports
├── router.js          ← Express router (self-contained, geen view engine setup)
├── lib/
│   ├── db.js
│   ├── parser.js      ← auto-detect ruw / bradley
│   ├── pipeline.js
│   ├── outputs.js
│   └── seed.js
├── views/             ← EJS templates
└── data/
    ├── articles.json
    ├── ape.json
    └── bunches.sqlite (gegenereerd na seed)
```

## Lokaal draaien

```bash
npm install
node lib/seed.js
node index.js
```

Open `http://localhost:3000/bunches`.

## Embedden in bestaande Express app

```js
const { createBunchesRouter } = require('./bunches-app');

app.use('/bunches', createBunchesRouter({
  dbPath: '/var/data/bunches.sqlite',
  basePath: '/bunches',
}));
```

Router rendert zijn eigen views — geen view engine setup in parent app nodig. Auth wordt door parent afgehandeld.

## Deploy op Render

Starter plan ($7/mnd) + 1GB persistent disk mount op `/var/data`. Env var `DB_PATH=/var/data/bunches.sqlite`. Build: `npm install && node lib/seed.js`. Start: `node index.js`.

## Workflow

1. Open `/bunches`
2. Sleep het ruwe CSV bestand in het drop-veld (of plak inhoud, of upload) + kies vertrekdatum
3. Verwerk → preview + downloads
4. Download **Inlezen.csv** + **YYBU*.csv** → upload naar systeem
5. Klik op een printlijst-knop → nieuw tabblad met printbare lijst inclusief checkboxes → Ctrl+P

## Master data beheer

- `/bunches/admin/articles` — Sorteerblad inlezen. **Inline edit**: wijzig velden in de rij en klik Opslaan.
- `/bunches/admin/ape` — BUN HVLH Lijst. **Inline edit** voor APE waarde.
- `/bunches/admin/runs` — geschiedenis van imports.
