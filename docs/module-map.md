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
| `Bunches` | Import bunches data, generate Inlezen / YYBU files, and print picking lists |
| `Fout Registratie` | Register daily mistakes for the active team |
| `Fouten Overzicht` | Report on mistakes by person, type, and period |
| `UKDocs Exportdocs` | Manage UK export zending collections and send-ready workflows |
| `Phyto Inspection` | Handle inspection-only papers for voorraad and nakeuring flows |
| `Inklokken` | Employee clock IN/OUT records |
| `Users` | Access and permission management |
| `Settings` | External connections and system configuration |

## Main Technical Files

| File | Role |
| --- | --- |
| `shadow-app/src/main.jsx` | Main frontend screens and client logic |
| `shadow-app/src/styles.css` | Shared frontend styling |
| `shadow-app/server/index.js` | Main backend API and app server |
| `shadow-app/server/bunches.js` | Bunches import, run history, downloads, and print-list backend logic |
| `shadow-app/public/bunches.html` | Embedded Bunches frontend used inside the Shadow app |
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

## Notes

- `UKDocs Exportdocs` and `Phyto Inspection` are intentionally separate menus so permissions can block one without blocking the other.
- `Bunches` print lists use browser print preview for the clean printable layout instead of relying only on a generated PDF.
- Viewer users should see shared UKDocs connection status through the UKDocs state endpoint, not only through admin settings pages.
