# Warehouse Dashboard

Read-only view of the live warehouse state and activity log. Phase 1 of the
scan-station rebuild plan: reuses the existing `serverBackend.js` endpoints
(`/status`, `/activity-log`) with zero changes to the backend or any scan
station — this is purely a new way to look at data that already exists.

## Live deployment: served from shadow-app, backed by Postgres

In production this folder's `public/` assets are served directly by
shadow-app's own server (`shadow-app/server/index.js`) under `/warehouse/` —
**not** via this folder's own `server.js`/`src/router.js`. The LAN machine
running `serverBackend.js` pushes its status snapshot and new activity events
to `POST /warehouse/ingest` on the Render app (shared-secret header
`X-Ingest-Secret` checked against `RENDER_INGEST_SECRET`), which upserts them
into the `warehouse_status` / `warehouse_activity_log` tables on shadow-app's
existing Postgres database. The dashboard's own reads (`/warehouse/api/status`,
`/warehouse/api/activity-log`) are served by shadow-app too, straight from
those tables — no `WAREHOUSE_BACKEND_URL` proxy involved, and gated behind a
logged-in shadow-app user with the `warehouse:view` permission.

`server.js` and `src/router.js` in this folder are unused in that
deployment — they're kept only for standalone/local testing against a live
`serverBackend.js` directly (see below), and were intentionally left
untouched when the Postgres wiring was added to shadow-app.

## Run it standalone

```bash
npm install
cp .env.example .env    # edit WAREHOUSE_BACKEND_URL if needed
npm start
```

Opens on `http://localhost:4500` by default. It talks to the backend
server-side (`src/router.js` proxies `/status` and `/activity-log`), so the
backend's address never has to be reachable from the browser directly, and
never appears in client-side JS.

## Why it's built this way

This is deliberately kept separate from any other app (including the
Render.com one) for now, but shaped so it can become a "slot" inside a
bigger app later without a rewrite:

- All the actual logic lives in `src/router.js` as one `createWarehouseDashboardRouter()`
  function that returns a standard Express `Router`. `server.js` is just a
  thin wrapper that mounts it at `/` and listens on a port — that's the only
  part that's specific to running this standalone.
- To mount it inside another Express app later as one more route:

  ```js
  const { createWarehouseDashboardRouter } = require('./warehouse-dashboard/src/router');
  app.use('/warehouse', createWarehouseDashboardRouter({
    backendUrl: 'http://10.1.100.47:4000'
  }));
  ```

  No changes needed in `router.js` or the front-end for that to work — every
  asset link and `fetch()` call in `public/` uses a relative path (`api/status`,
  `styles.css`, ...), not a root-absolute one, specifically so the whole thing
  keeps working regardless of which prefix it's mounted under.

## What's here

- `GET /api/status` — proxies the backend's `/status` (warehouse location grid).
- `GET /api/activity-log` — proxies the backend's `/activity-log`, passing through `date`/`limit`.
- `GET /api/config` — just the backend's hostname, so the page can show what it's connected to.
- `public/` — the dashboard itself: stat tiles, a filterable location table, a
  recent-activity feed, and the same "upload issue" banner logic added to the
  Streamlit dashboard (Google Drive / upload failures surfaced within 30 min).

## What this intentionally doesn't do (yet)

Read-only, by design — it doesn't reset scans, upload CSVs, or write anything
back to the backend. That's the whole point of doing this as Phase 1: prove
the web stack talks cleanly to the real backend before any scan-station logic
gets touched.
