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
- `Fout Registratie`
- `Fouten Overzicht`
- `UKdocs Print`
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
render.yaml
Dockerfile
```

## Notes

- Some older folders still exist in the repo for source history or helper tooling.
- The Shadow app is the main user-facing system.
- For setup, workflows, and module behavior, use the app book in `docs/app-book.md`.
