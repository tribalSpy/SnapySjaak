# Sjaak vd Vijver Dashboard

Main operations app for Sjaak vd Vijver workflows.

This repository now contains the Shadow web app used on Render, plus older helper tools and source folders that still support parts of the workflow.

## Start Here

- Full app guide: [docs/app-book.md](docs/app-book.md)
- Module list: [docs/module-map.md](docs/module-map.md)

## Main Shadow Modules

- `Photos`
- `Fust`
- `CMR Print`
- `Hal Locations`
- `Expedition Sticker`
- `Bunches`
- `Fout Registratie`
- `Fouten Overzicht`
- `UKDocs Exportdocs`
- `Phyto Inspection`
- `Inklokken`
- `Users`
- `Settings`

## Tech Stack

- Frontend: React + Vite
- Backend: Node.js
- Database: PostgreSQL
- External services: Google Sheets, Google Drive, Gmail
- Deployment: Render
- Persistent storage: Render disk for cache, files, and local snapshots

## Main App Folder

The active Shadow app lives here:

```text
shadow-app/
```

Important files:

```text
shadow-app/src/main.jsx
shadow-app/src/styles.css
shadow-app/server/index.js
shadow-app/server/bunches.js
shadow-app/public/bunches.html
render.yaml
Dockerfile
```

## Notes

- Some older folders still exist in the repo for source history or helper tooling.
- The Shadow app is the main user-facing system.
- The app now uses PostgreSQL for durable app-owned records, while Google Sheets and Gmail still stay part of the operational workflow where needed.
- Render persistent disk is still used for generated files, uploads, snapshots, and cache-style shared state.
- `UKDocs Exportdocs` is the shipment collection page, while `Phyto Inspection` is the separate inspection-paper workflow.
- `Bunches` print lists now use the browser print-preview flow so users can print or save as PDF from the cleaner HTML layout.
- For setup, workflows, and module behavior, use the app book in `docs/app-book.md`.
