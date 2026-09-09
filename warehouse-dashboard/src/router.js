const express = require('express');
const fs = require('fs');
const path = require('path');

// The "slot": everything the dashboard needs, as one mountable Express Router.
// Standalone today via server.js (mounted at '/'); can be require()'d and
// app.use('/warehouse', createWarehouseDashboardRouter(...)) 'd into another
// app later with no changes here — that's why every asset/fetch path in the
// front-end is relative, not root-absolute.
function createWarehouseDashboardRouter(options = {}) {
  const backendUrl = (options.backendUrl || process.env.WAREHOUSE_BACKEND_URL || 'http://127.0.0.1:4000')
    .replace(/\/+$/, '');

  // Dashboard-local settings (currently just hidden/excluded references), kept
  // separate from the backend entirely — this only ever changes what THIS
  // dashboard shows, never anything scan stations do. Defaults to the same
  // path the Streamlit dashboard used, so the two can share settings if both
  // run side by side during the transition.
  const settingsPath = options.settingsPath
    || process.env.SETTINGS_PATH
    || path.join('C:', 'RFID', 'data', 'scan_dashboard_settings.json');

  const router = express.Router();

  // Non-sensitive: just lets the front-end show which backend/settings file it's using.
  router.get('/api/config', (req, res) => {
    res.json({ backendHost: safeHost(backendUrl), settingsPath });
  });

  router.get('/api/settings', (req, res) => {
    res.json({ excludedReferences: readExcludedReferences(settingsPath) });
  });

  router.post('/api/settings', express.json(), (req, res) => {
    const list = Array.isArray(req.body?.excludedReferences) ? req.body.excludedReferences : [];
    const normalized = Array.from(new Set(
      list.map(r => String(r).trim().toUpperCase()).filter(Boolean)
    )).sort();

    try {
      fs.mkdirSync(path.dirname(settingsPath), { recursive: true });
      fs.writeFileSync(settingsPath, JSON.stringify({ excluded_references: normalized }, null, 2));
      res.json({ ok: true, excludedReferences: normalized });
    } catch (err) {
      res.status(500).json({ ok: false, error: String(err.message || err) });
    }
  });

  // Server-side proxy for both API calls below. Two reasons this isn't a direct
  // browser->backend fetch: (1) serverBackend.js doesn't send CORS headers, so a
  // cross-origin browser fetch would just fail; (2) it keeps the backend's real
  // address out of client-side JS, so this router stays portable across
  // environments without editing the front-end.
  router.get('/api/status', async (req, res) => {
    try {
      const upstream = await fetch(`${backendUrl}/status`);
      if (!upstream.ok) throw new Error(`backend responded ${upstream.status}`);
      const data = await upstream.json();
      res.json({ ok: true, data });
    } catch (err) {
      res.status(502).json({ ok: false, error: String(err.message || err) });
    }
  });

  router.get('/api/activity-log', async (req, res) => {
    try {
      const params = new URLSearchParams();
      if (req.query.date) params.set('date', String(req.query.date));
      if (req.query.limit) params.set('limit', String(req.query.limit));
      const qs = params.toString();

      const upstream = await fetch(`${backendUrl}/activity-log${qs ? `?${qs}` : ''}`);
      if (!upstream.ok) throw new Error(`backend responded ${upstream.status}`);
      const data = await upstream.json();
      res.json(data);
    } catch (err) {
      res.status(502).json({ ok: false, error: String(err.message || err), events: [] });
    }
  });

  router.use(express.static(path.join(__dirname, '..', 'public')));

  return router;
}

function readExcludedReferences(settingsPath) {
  try {
    if (!fs.existsSync(settingsPath)) return [];
    const payload = JSON.parse(fs.readFileSync(settingsPath, 'utf8'));
    const list = Array.isArray(payload.excluded_references) ? payload.excluded_references : [];
    return Array.from(new Set(list.map(r => String(r).trim().toUpperCase()).filter(Boolean))).sort();
  } catch {
    return [];
  }
}

function safeHost(url) {
  try {
    return new URL(url).host;
  } catch {
    return 'unknown';
  }
}

module.exports = { createWarehouseDashboardRouter };
