// router.js — Express router. Mount in je bestaande app:
//
//   const { createBunchesRouter } = require('./bunches-app');
//   app.use('/bunches', createBunchesRouter({ dbPath: './data/bunches.sqlite' }));
//
// Router is self-contained: eigen EJS renderer met layout-helper.
// Geen dependencies op parent app's view engine.

const express = require('express');
const path = require('path');
const ejs = require('ejs');
const { createDb } = require('./lib/db');
const { parsePaste } = require('./lib/parser');
const { processData } = require('./lib/pipeline');
const {
  generateInlezenCsv,
  generateUniFile,
  generatePrintlijstHtml,
} = require('./lib/outputs');

const VIEWS_DIR = path.join(__dirname, 'views');

/**
 * Render een view met layout-support. `<% layout('_layout') %>` bovenaan een view
 * wordt hier verwerkt: inner view wordt gerenderd, layout wordt gerenderd met `body`.
 */
function renderView(viewName, locals, callback) {
  let requestedLayout = null;
  const localsWithLayout = {
    ...locals,
    layout: (name) => { requestedLayout = name; return ''; },
  };
  const viewPath = path.join(VIEWS_DIR, viewName + '.ejs');
  ejs.renderFile(viewPath, localsWithLayout, {}, (err, body) => {
    if (err) return callback(err);
    if (!requestedLayout) return callback(null, body);
    const layoutPath = path.join(VIEWS_DIR, requestedLayout + '.ejs');
    ejs.renderFile(layoutPath, { ...locals, body }, {}, (err2, html) => {
      if (err2) return callback(err2);
      callback(null, html);
    });
  });
}

function sendView(res, viewName, locals) {
  // Merge res.locals (bv. basePath gezet door router middleware) in template locals
  const mergedLocals = { ...(res.locals || {}), ...locals };
  renderView(viewName, mergedLocals, (err, html) => {
    if (err) {
      console.error('View render error:', err);
      return res.status(500).send('View error: ' + (err.message || err));
    }
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.send(html);
  });
}

function createBunchesRouter(options = {}) {
  const {
    dbPath = path.join(__dirname, 'data', 'bunches.sqlite'),
    basePath = '',  // optioneel: voor link-generatie
  } = options;

  const db = createDb(dbPath);
  const router = express.Router();

  // Views config — gebruik EJS uit deze app's views/
  router.use((req, res, next) => {
    res.locals.basePath = basePath;
    next();
  });

  // Body parsers (alleen binnen deze router)
  router.use(express.urlencoded({ extended: true, limit: '5mb' }));
  router.use(express.json({ limit: '5mb' }));

  // Alle runs worden persistent opgeslagen in de DB.
  // Token is de string 'run-{id}' — geldig zolang de run in de DB staat.
  function getResultByToken(token) {
    if (!token || !token.startsWith('run-')) return null;
    const id = parseInt(token.slice(4), 10);
    if (isNaN(id)) return null;
    const row = db.getRun(id);
    if (!row || !row.result_json) return null;
    try {
      const result = JSON.parse(row.result_json);
      result._runId = id;
      result._createdAt = row.created_at;
      result._label = row.label;
      return result;
    } catch (e) {
      console.error('Failed to parse result_json for run', id, e.message);
      return null;
    }
  }

  // ===== Routes =====

  // Home: paste form
  router.get('/', (req, res) => {
    sendView(res, 'paste', {
      title: 'Bunches import',
      articleCount: db.listArticles().length,
      apeCount: db.listApe().length,
    });
  });

  // Process paste
  router.post('/process', (req, res) => {
    const text = req.body.paste_text || '';
    const vertrekDatum = (req.body.vertrek_datum || '').trim();  // yyyy-mm-dd
    const { rows, warnings: parseWarnings } = parsePaste(text);

    if (rows.length === 0) {
      return sendView(res, 'paste', {
        title: 'Bunches import',
        articleCount: db.listArticles().length,
        apeCount: db.listApe().length,
        error: parseWarnings.join('; '),
        previous: text,
        defaultDate: vertrekDatum || new Date().toISOString().slice(0, 10),
      });
    }

    const result = processData(rows, db);
    const warnings = [...parseWarnings.map(m => ({ type: 'parse', message: m })), ...result.warnings];

    // Bepaal beschikbare Hoes×Tak combinaties
    const taksByHoes = { Plast: new Set(), Kraft: new Set() };
    for (const it of result.printlijst_plast) if (it.tak) taksByHoes.Plast.add(it.tak);
    for (const it of result.printlijst_kraft) if (it.tak) taksByHoes.Kraft.add(it.tak);
    const takOrder = ['3T', '4T', '5T', '10T'];
    const sortedTaks = (set) => takOrder.filter(t => set.has(t));

    const finalDate = vertrekDatum || new Date().toISOString().slice(0, 10);
    const fullResult = {
      ...result,
      warnings,
      generatedAt: new Date().toISOString(),
      vertrekDatum: finalDate,
      availableTaks: {
        Plast: sortedTaks(taksByHoes.Plast),
        Kraft: sortedTaks(taksByHoes.Kraft),
      },
    };

    // Opslaan in DB
    let runInsert;
    try {
      runInsert = db.logRun({
        user: req.user?.email || req.user?.id || 'anon',
        row_count: rows.length,
        total_som: result.summary.total_som,
        status: warnings.length > 0 ? 'warnings' : 'ok',
        warnings_json: JSON.stringify(warnings),
        raw_json: JSON.stringify(rows.slice(0, 500)),
        result_json: JSON.stringify(fullResult),
        vertrek_datum: finalDate,
        label: null,
      });
    } catch (e) {
      console.error('Failed to log run:', e);
      return res.status(500).send('Opslag mislukt: ' + e.message);
    }

    const runId = runInsert.lastInsertRowid;
    res.redirect(`${basePath}/result/run-${runId}`);
  });

  // Update date for an existing run (herbereken niks, gewoon datum bijwerken)
  router.post('/result/:token/date', (req, res) => {
    if (!req.params.token.startsWith('run-')) return res.redirect(`${basePath}/`);
    const runId = parseInt(req.params.token.slice(4), 10);
    const row = db.getRun(runId);
    if (!row) return res.redirect(`${basePath}/`);
    const newDate = (req.body.vertrek_datum || '').trim();
    if (newDate) {
      // Update vertrekDatum in het opgeslagen result JSON
      let result;
      try { result = JSON.parse(row.result_json); } catch (e) { result = {}; }
      result.vertrekDatum = newDate;
      db.updateRunDate(runId, newDate, JSON.stringify(result));
    }
    res.redirect(`${basePath}/result/${req.params.token}`);
  });

  // Update run label (optioneel, voor herkenning in de lijst)
  router.post('/result/:token/label', (req, res) => {
    if (!req.params.token.startsWith('run-')) return res.redirect(`${basePath}/`);
    const runId = parseInt(req.params.token.slice(4), 10);
    const label = (req.body.label || '').trim() || null;
    db.updateRunLabel(runId, label);
    res.redirect(`${basePath}/result/${req.params.token}`);
  });

  // Delete run
  router.post('/admin/runs/:id/delete', (req, res) => {
    const id = parseInt(req.params.id, 10);
    if (!isNaN(id)) db.deleteRun(id);
    res.redirect(`${basePath}/admin/runs`);
  });

  // Result page
  router.get('/result/:token', (req, res) => {
    const result = getResultByToken(req.params.token);
    if (!result) {
      res.status(404);
      return sendView(res, 'paste', {
        title: 'Bunches import',
        articleCount: db.listArticles().length,
        apeCount: db.listApe().length,
        error: 'Resultaat verlopen of niet gevonden. Plak opnieuw.',
      });
    }
    sendView(res, 'result', {
      title: 'Resultaat',
      token: req.params.token,
      result,
      yybuSheets: Object.keys(result.yybu_files).sort(),
    });
  });

  // Downloads
  router.get('/download/:token/inlezen.csv', (req, res) => {
    const result = getResultByToken(req.params.token);
    if (!result) return res.status(404).send('Niet gevonden');
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="Inlezen.csv"');
    res.send(generateInlezenCsv(result.inlezen));
  });

  router.get('/download/:token/yybu/:sheet.csv', (req, res) => {
    const result = getResultByToken(req.params.token);
    if (!result) return res.status(404).send('Niet gevonden');
    const sheet = req.params.sheet;
    const lines = result.yybu_files[sheet];
    if (!lines) return res.status(404).send('Sheet niet gevonden');
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${sheet}.csv"`);
    res.send(generateUniFile(sheet, lines, result.vertrekDatum));
  });

  // Printlijst — optioneel filter op Tak
  // /printlijst/:token/:hoes          → alle (Plast of Kraft)
  // /printlijst/:token/:hoes/:tak     → gefilterd op Tak (bv. 3T, 5T)
  // Query ?format=pdf → download als PDF
  router.get('/printlijst/:token/:hoes/:tak?', (req, res) => {
    const result = getResultByToken(req.params.token);
    if (!result) return res.status(404).send('Niet gevonden');
    const hoesParam = req.params.hoes.toLowerCase();
    const tak = req.params.tak;
    let items;
    let hoesLabel;
    if (hoesParam === 'plast') {
      items = result.printlijst_plast;
      hoesLabel = 'Plastic';
    } else if (hoesParam === 'kraft') {
      items = result.printlijst_kraft;
      hoesLabel = 'Kraft';
    } else {
      return res.status(404).send('Onbekende hoes (gebruik plast of kraft)');
    }
    if (tak) {
      items = items.filter(i => i.tak === tak);
    }
    const title = tak ? `Printlijst ${hoesLabel} ${tak}` : `Printlijst ${hoesLabel}`;

    if (req.query.format === 'pdf') {
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      res.send(generatePrintlijstHtml(title, items, result.vertrekDatum, { autoPrint: true }));
      return;
    }

    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.send(generatePrintlijstHtml(title, items, result.vertrekDatum));
  });

  // ===== Admin pages =====

  router.get('/admin/articles', (req, res) => {
    const q = (req.query.q || '').toLowerCase();
    let articles = db.listArticles();
    if (q) {
      articles = articles.filter(a =>
        String(a.broncode).includes(q) ||
        (a.omschrijving || '').toLowerCase().includes(q) ||
        (a.variant_code || '').toLowerCase().includes(q));
    }
    // Bulk-melding uit query decoderen (rudimentair; volstaat voor deze use)
    let bulkMsg = null;
    if (req.query.bulk) {
      bulkMsg = String(req.query.bulk).replace(/\+/g, ' ');
    }
    sendView(res, 'articles', {
      title: 'Artikel-beheer',
      articles: articles.slice(0, 200),
      total: db.listArticles().length,
      query: req.query.q || '',
      bulkMsg,
    });
  });

  router.post('/admin/articles', (req, res) => {
    const b = req.body;
    if (!b.broncode || !b.variant_code) {
      return res.redirect(`${basePath}/admin/articles?error=ontbrekende+velden`);
    }
    db.upsertArticle({
      broncode: parseInt(b.broncode, 10),
      omschrijving: (b.omschrijving || '').trim(),
      variant_code: b.variant_code.trim(),
      broncode_inlezen: b.broncode_inlezen ? parseInt(b.broncode_inlezen, 10) : null,
      afronden: b.afronden ? parseInt(b.afronden, 10) : null,
      zonder_tak: b.zonder_tak === '1' || b.zonder_tak === 'on' || b.zonder_tak === 'true',
    });
    res.redirect(`${basePath}/admin/articles?ok=1`);
  });

  router.post('/admin/articles/:broncode/delete', (req, res) => {
    db.deactivateArticle(parseInt(req.params.broncode, 10));
    res.redirect(`${basePath}/admin/articles?deleted=1`);
  });

  // Bulk zonder_tak aan/uit voor een lijst broncodes
  router.post('/admin/articles/bulk-zonder-tak', (req, res) => {
    const broncodesText = (req.body.broncodes || '').trim();
    const action = req.body.action;  // "on" or "off"
    if (!broncodesText || !['on', 'off'].includes(action)) {
      return res.redirect(`${basePath}/admin/articles?error=bulk+invalid`);
    }
    const broncodes = broncodesText.split(/[\s,;]+/).map(s => parseInt(s, 10)).filter(n => !isNaN(n));
    if (broncodes.length === 0) {
      return res.redirect(`${basePath}/admin/articles?error=geen+geldige+broncodes`);
    }
    const value = action === 'on' ? 1 : 0;
    const stmt = db.db.prepare('UPDATE articles SET zonder_tak = ?, updated_at = datetime(\'now\') WHERE broncode = ?');
    const tx = db.db.transaction((codes) => {
      let updated = 0;
      const notFound = [];
      for (const bc of codes) {
        const r = stmt.run(value, bc);
        if (r.changes > 0) updated++;
        else notFound.push(bc);
      }
      return { updated, notFound };
    });
    const result = tx(broncodes);
    const msg = `${result.updated}+bijgewerkt` + (result.notFound.length ? `,+niet+gevonden:+${result.notFound.join(',')}` : '');
    res.redirect(`${basePath}/admin/articles?bulk=${msg}`);
  });

  router.get('/admin/ape', (req, res) => {
    const q = (req.query.q || '').toLowerCase();
    let ape = db.listApe();
    if (q) ape = ape.filter(a => (a.omschrijving || '').toLowerCase().includes(q));
    sendView(res, 'ape', {
      title: 'APE-beheer (BUN HVLH Lijst)',
      ape: ape.slice(0, 300),
      total: db.listApe().length,
      query: req.query.q || '',
    });
  });

  router.post('/admin/ape', (req, res) => {
    const b = req.body;
    if (!b.omschrijving || !b.ape) {
      return res.redirect(`${basePath}/admin/ape?error=ontbrekende+velden`);
    }
    db.upsertApe({
      omschrijving: b.omschrijving.trim(),
      ape: parseInt(b.ape, 10),
    });
    res.redirect(`${basePath}/admin/ape?ok=1`);
  });

  router.post('/admin/ape/delete', (req, res) => {
    if (req.body.omschrijving) db.deleteApe(req.body.omschrijving);
    res.redirect(`${basePath}/admin/ape?deleted=1`);
  });

  // Audit log
  router.get('/admin/runs', (req, res) => {
    sendView(res, 'runs', {
      title: 'Geschiedenis',
      runs: db.listRuns(),
    });
  });

  return router;
}

module.exports = { createBunchesRouter };
