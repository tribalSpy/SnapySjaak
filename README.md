# Google Drive Run Dashboard

Streamlit dashboard for browsing recent run folders from one or more Google Drive accounts and older runs from a local archive.

## Features

- Reads recent run folders from one or more Google Drive root folders
- Reads older run folders from an optional local archive path
- Parses run folder names in the format `customer_YYYYMMDD` or `customer_YYYYMMDD_runid`
- Filters runs by date and customer code
- Groups runs by customer
- Shows carrier, run ID, folder name, QR info, and images for each run
- Reports malformed folder names in the sidebar
- Persists run metadata and downloaded Drive images in a local cache for faster reloads
- Includes a sidebar button to clear the saved cache and reload from source
- Includes a Hal Locations menu with integrated sticker PDF generation from halindeling Excel files

## Project Structure

```text
app.py
src/
  drive_service.py
  models.py
  parser.py
  ui_helpers.py
credentials/
  service_account.json
.streamlit/
  secrets.toml
requirements.txt
StickerPrinter/
.env.example
```

## Setup

1. Create and activate a virtual environment.
2. Install dependencies from `requirements.txt`.
3. Copy `.env.example` to `.env`.
4. Fill in:
   - `GOOGLE_DRIVE_ROOT_FOLDER_ID`
   - `GOOGLE_APPLICATION_CREDENTIALS`
   - `GOOGLE_DRIVE_ACCOUNT_NAMES` if you want extra Google Drive accounts
   - `GOOGLE_DRIVE_ROOT_FOLDER_ID_<NAME>` for each extra account
   - `GOOGLE_APPLICATION_CREDENTIALS_<NAME>` or `GOOGLE_SERVICE_ACCOUNT_JSON_<NAME>` only if an extra account needs different credentials
   - `LOCAL_ARCHIVE_ROOT` (optional)
   - `LOCAL_ARCHIVE_AFTER_DAYS` (optional, defaults to `7`)
5. Place your service account JSON at the configured credentials path.
6. Share the relevant Google Drive folders with the service account if needed.

Example for a second account:

```env
GOOGLE_DRIVE_ACCOUNT_NAMES=second
GOOGLE_DRIVE_ROOT_FOLDER_ID_SECOND=your_second_google_drive_root_folder_id
```

If the second folder is shared with the same service account, you can reuse the default
`GOOGLE_APPLICATION_CREDENTIALS` and only add another folder ID. Add per-account credentials
only when the extra Google Drive source truly needs a different login.

You can also run the app with Streamlit secrets instead of a local `.env` file.
For Community Cloud, use `.streamlit/secrets.toml.example` as your template and paste
the full service account JSON into `GOOGLE_SERVICE_ACCOUNT_JSON`.

## Run

```powershell
streamlit run app.py
```

Use the sidebar menu to switch between `Photo Dashboard` and `Hal Locations`.

## Deploy To Streamlit Community Cloud

1. Push this project to GitHub.
2. In Streamlit Community Cloud, create a new app from your GitHub repository.
3. Set the main file path to `app.py`.
4. In the app's Secrets settings, add the values from `.streamlit/secrets.toml.example`.

Do not commit these private files:

- `.env`
- `.streamlit/secrets.toml`
- `credentials/service_account.json`

## Source Rules

- Recent runs are loaded from all configured Google Drive accounts.
- Runs older than `LOCAL_ARCHIVE_AFTER_DAYS` are loaded from `LOCAL_ARCHIVE_ROOT`.
- This lets you keep the current week in Drive and older runs in a local archive.

## Folder Assumptions

- Structure: `RootFolder / CarrierFolder / RunFolder / files`
- Also supported: `RootFolder / RunFolder / files`
- Local archive runs are regrouped into `RootFolder / YYYY-MM-DD / RunFolder / files`
- Carrier archive runs are regrouped into `RootFolder / CarrierFolder / YYYY-MM-DD / RunFolder / files`
- If the local archive root has one carrier folder, new root-level run folders are moved into that carrier's date folders.
- Carrier name comes from the parent folder name
- Run folder names must match:
  - `cust123_20260310`
  - `cust123_20260310_232544`
- QR info is resolved in this order:
  - `qr.txt` or `qr.json`
  - filenames containing `qr`
  - fallback message

## Notes

- No database is used.
- The app can combine multiple Google Drive accounts and a local archive.
- Malformed folders are skipped and shown in the sidebar.
