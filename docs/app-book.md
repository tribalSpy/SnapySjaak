# Sjaak vd Vijver Dashboard Book

## 1. Purpose

This document describes the Shadow app inside the Sjaak vd Vijver Dashboard repository.

The app combines several daily operations into one web system:

- photo and run browsing
- fust registration and overview
- CMR printing
- hal location tools
- expedition sticker generation
- mistake registration
- mistake reporting
- UK export document workflows
- employee clocking
- user permissions
- settings and external connections

The goal is to keep the daily operational tools in one app, while still allowing different users to see only the menus they are allowed to use.

## 2. Main Architecture

## 2.1 Frontend

The frontend is a React app in:

```text
shadow-app/src/
```

Main UI logic is mostly in:

```text
shadow-app/src/main.jsx
```

Main styling is in:

```text
shadow-app/src/styles.css
```

## 2.2 Backend

The backend is a Node.js server in:

```text
shadow-app/server/index.js
```

This server:

- serves the React app
- exposes API endpoints for every module
- reads and writes persistent app state
- talks to PostgreSQL
- talks to Google Sheets
- talks to Google Drive
- talks to Gmail
- manages local cached files and backups

## 2.3 Deployment

The production app is intended to run on Render.

Main deployment files:

```text
render.yaml
Dockerfile
```

Render uses:

- a web service for the Shadow app
- a persistent disk mounted at `/var/data`
- environment variables for credentials and external settings

## 2.4 Storage Model

The app currently uses a mixed storage model:

- PostgreSQL for durable structured data
- Google Sheets for spreadsheet-based workflows and backup-style visibility
- local persistent disk for cache, uploaded/generated files, and snapshots

This is important because some modules are operationally driven by Sheets, while newer durable app data is moving into PostgreSQL.

## 3. Main Menus

The visible app menus are permission-based.

Current main menus include:

1. `Photos`
2. `Fust`
3. `CMR Print`
4. `Hal Locations`
5. `Expedition Sticker`
6. `Fout Registratie`
7. `Fouten Overzicht`
8. `UKdocs Print`
9. `Inklokken`
10. `Users`
11. `Settings`

## 4. Module Overview

## 4.1 Photos

Purpose:

- browse run folders and images
- inspect QR information
- work with recent Google Drive runs and older archived runs

Main behavior:

- recent runs can be loaded from Google Drive
- older runs can come from local archive paths
- cached image metadata helps speed up reloads

## 4.2 Fust

Purpose:

- capture IN and OUT packaging movements
- manage CMR and fustbon references
- show balances and transaction history

Main behavior:

- actions are stored durably
- overview pages aggregate movements by customer and country
- spreadsheet sync is still part of the operational flow
- local backups can be created and restored for missing metadata recovery

Important note:

- this module is one of the most sensitive for data consistency because edits, deletes, spreadsheet sync, and database state all need to stay aligned

## 4.3 CMR Print

Purpose:

- manage CMR customers, exporters, transport info, loading places, and templates
- print single or batch CMR documents

Main behavior:

- template editor controls field positions and sizes
- customer records link to exporter, transport, and loading place profiles
- manual fields 5, 7, 9, and 17 remain editable
- batch print supports per-customer overrides before preview or print

Current important data groups:

- customer info
- exporter info
- transport info
- loading places
- templates
- default template settings

## 4.4 Hal Locations

Purpose:

- inspect and generate location-based output from `ERP_PASTE`

Main behavior:

- spreadsheet data is loaded from Google Sheets
- helper logic groups location and customer prefixes
- used directly by sticker-generation workflows

## 4.5 Expedition Sticker

Purpose:

- generate expedition sticker PDFs from planning and split files combined with live `ERP_PASTE`

Main behavior:

- planner uploads shared planning and split files once
- those files are saved for reuse by others
- `ERP_PASTE` is pulled from Google Sheets
- sticker PDFs are generated from the combined source

Important note:

- the sticker workflow depends on clean source files and valid customer/location matching

## 4.6 Fout Registratie

Purpose:

- register daily mistakes for the active team

Main behavior:

- supports 3 UI languages
- team for the day is selected first
- users register mistakes by person and type
- past data syncs to the `fouten` sheet
- stored mistake type keys stay canonical, while UI labels can be translated

Important note:

- the shared overview/reporting side should keep type reporting in English for consistency

## 4.7 Fouten Overzicht

Purpose:

- review mistake data across people and time periods

Main behavior:

- shows person ranking
- shows mistake counts by type
- supports day, week, and month analysis
- intended as a separate permission-controlled reporting page

Use cases:

- see who makes the most mistakes
- see which type happens most often
- compare days, weeks, and months
- look for patterns by period or season

## 4.8 UKdocs Print

Purpose:

- collect, prepare, and track UK export document packages

Main behavior:

- shipments are date-driven
- saved collections remain stored but daily views show only the selected day
- Gmail sync picks up matching files for the selected export date
- generated UKdocs files stay linked to the correct shipment
- users can download all shipment files from one place

Important note:

- this module combines spreadsheet source data, generated Excel files, Gmail attachments, uploaded PDFs, and operational send-ready actions

## 4.9 Inklokken

Purpose:

- register employee IN and OUT records
- keep employee and clock records synced with the configured spreadsheet

Main behavior:

- employees come from the badge spreadsheet
- records are written to the configured backup tab
- work duration can be derived from IN and OUT pairs

## 4.10 Users

Purpose:

- manage who can access which parts of the app

Main behavior:

- menu access is permission-based
- some pages are intentionally separated so they can be blocked independently

Examples:

- `Fout Registratie` and `Fouten Overzicht` can be split
- operational users can get only the tools they need

## 4.11 Settings

Purpose:

- central place for spreadsheet IDs, mail settings, CMR/Drive settings, and system behavior

Main behavior:

- stores Google Sheets settings
- stores UKdocs/Gmail settings
- stores CMR/Drive settings
- stores SMTP mail settings
- exposes connection tests
- includes Fust database backfill and backup tools

## 5. Main Integrations

## 5.1 PostgreSQL

Used for durable structured data.

Recommended production direction:

- keep PostgreSQL as the main source for app-owned data
- keep spreadsheets as operational input/output and secondary visibility where needed

## 5.2 Google Sheets

Used for:

- clock employee source
- clock backup rows
- mistake sheet sync
- hal locations / `ERP_PASTE`
- UKdocs source spreadsheet
- some operational source-of-truth flows

## 5.3 Google Drive

Used for:

- photo/run browsing
- file upload/download flows
- CMR-related drive integration where configured

## 5.4 Gmail

Used in UKdocs workflows to:

- read matching export attachments
- connect export files to the correct shipment
- reduce manual document chasing

## 5.5 SMTP Email

Used to send ready-paper notifications where configured.

## 6. Daily Operational Workflows

## 6.1 Fust

Typical flow:

1. choose IN or OUT
2. enter customer/transport and package counts
3. attach or link document reference
4. save action
5. review in overview and last actions

## 6.2 CMR Print

Typical flow:

1. maintain customer and linked profile data
2. choose template
3. fill manual fields where needed
4. preview CMR
5. print single or batch documents

## 6.3 UKdocs Print

Typical flow:

1. load sendings for the selected date
2. open a shipment
3. collect generated files, uploaded files, and Gmail-picked files
4. verify progress
5. send ready notification

## 6.4 Fout Registratie

Typical flow:

1. select today team
2. choose worker
3. choose mistake type
4. add optional or required comment
5. save entry
6. review later in summary or overview pages

## 7. Backups and Recovery

There are several backup concepts in the system:

- Render persistent disk
- Fust JSON-style backups/snapshots
- spreadsheet copies
- database backfill tools

Important rule:

- cache reset is not the same as safe backup
- app-owned important data should live in PostgreSQL or another durable system, not only in cache

## 8. Current Risks and Sensitive Areas

The most sensitive parts of the app are:

1. Fust consistency between app state, spreadsheet rows, and database rows
2. CMR template persistence and imported customer data
3. UKdocs file matching and duplicate prevention
4. Gmail and Drive token expiry
5. shared state that used to live only on disk/cache

## 9. Recommended Direction

Recommended long-term direction:

1. move durable app-owned records fully into PostgreSQL
2. keep Google Sheets only where the business truly needs spreadsheet access
3. keep generated files on persistent storage, but keep references in the database
4. keep module permissions separate and explicit
5. document each workflow as it stabilizes

## 10. Important Files

Core app files:

```text
shadow-app/src/main.jsx
shadow-app/src/styles.css
shadow-app/server/index.js
shadow-app/public/dag-foutjes.html
```

Deployment files:

```text
render.yaml
Dockerfile
```

Support areas still present in the repository:

```text
cmrprint/
StickerPrinter/
foutjeskoelcel/
InKlokken/
src/
app.py
```

## 11. How To Extend This Book

Good next chapters to add later:

1. full permission matrix
2. API endpoint list
3. database table list
4. environment variable reference
5. backup and restore procedures
6. Render deployment checklist
7. user guide per department

## 12. Suggested Writing Workflow In VS Code

With `Markdown All in One`, this works well:

1. open `docs/app-book.md`
2. use heading structure for chapters
3. auto-generate a table of contents if you want
4. split future details into separate markdown files
5. link them back into this book

That lets this file act as the main app manual, while smaller docs hold the deep technical details.
