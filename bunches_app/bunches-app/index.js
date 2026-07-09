// index.js — twee modi:
//   1. Standalone: `node index.js` start eigen Express server
//   2. Embed: `require('./bunches-app').createBunchesRouter(opts)` mount in bestaande app

const express = require('express');
const path = require('path');
const fs = require('fs');
const { createBunchesRouter } = require('./router');
const { seedFromJson } = require('./lib/seed');

// === Voor embedding ===
module.exports = {
  createBunchesRouter,
  seedFromJson,
};

// === Standalone server ===
if (require.main === module) {
  const PORT = process.env.PORT || 3000;
  const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'data', 'bunches.sqlite');

  // Eerste run: seed als DB nieuw is
  const isNew = !fs.existsSync(DB_PATH);
  if (isNew && process.env.AUTO_SEED !== 'false') {
    console.log('Nieuwe DB, seeding vanuit data/...');
    seedFromJson(
      DB_PATH,
      path.join(__dirname, 'data', 'articles.json'),
      path.join(__dirname, 'data', 'ape.json'),
    );
  }

  const app = express();

  // Mount router op /bunches. Router rendert zelf zijn views — geen view engine setup nodig.
  app.use('/bunches', createBunchesRouter({
    dbPath: DB_PATH,
    basePath: '/bunches',
  }));

  // Health check
  app.get('/health', (req, res) => res.json({ ok: true, time: new Date().toISOString() }));

  // Root → redirect naar app
  app.get('/', (req, res) => res.redirect('/bunches'));

  app.listen(PORT, () => {
    console.log(`Bunches app draait op http://localhost:${PORT}/bunches`);
  });
}
