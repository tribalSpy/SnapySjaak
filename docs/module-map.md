# Module Map

Quick module reference for the Shadow app.

## Main User Menus

| Menu | Purpose |
| --- | --- |
| `Photos` | Browse run folders, images, and QR info |
| `Fust` | Register IN/OUT actions and review balances |
| `CMR Print` | Manage templates, customers, and print CMRs |
| `Hal Locations` | Load and process `ERP_PASTE` location data |
| `Expedition Sticker` | Generate sticker PDFs from planning, split, and live sheet data |
| `Fout Registratie` | Register daily mistakes for the active team |
| `Fouten Overzicht` | Report on mistakes by person, type, and period |
| `UKdocs Print` | Manage UK export document collections |
| `Inklokken` | Employee clock IN/OUT records |
| `Users` | Access and permission management |
| `Settings` | External connections and system configuration |

## Main Technical Files

| File | Role |
| --- | --- |
| `shadow-app/src/main.jsx` | Main frontend screens and client logic |
| `shadow-app/src/styles.css` | Shared frontend styling |
| `shadow-app/server/index.js` | Main backend API and app server |
| `shadow-app/public/dag-foutjes.html` | Embedded mistake registration app |
| `render.yaml` | Render service configuration |
| `Dockerfile` | Container build for production |

## Main External Systems

| System | Used For |
| --- | --- |
| PostgreSQL | Durable structured app data |
| Google Sheets | Operational inputs, sync outputs, and visibility |
| Google Drive | Run files, uploads, and browsing |
| Gmail | UKdocs attachment pickup |
| SMTP | Ready-mail sending |
| Render disk | Cache, backups, generated files |
