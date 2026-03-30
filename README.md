# Google Drive Run Dashboard

Streamlit dashboard for browsing recent run folders from Google Drive and older runs from a local archive.

## Features

- Reads recent run folders from a Google Drive root folder
- Reads older run folders from an optional local archive path
- Parses run folder names in the format `customer_YYYYMMDD` or `customer_YYYYMMDD_runid`
- Filters runs by date and customer code
- Groups runs by customer
- Shows carrier, run ID, folder name, QR info, and images for each run
- Reports malformed folder names in the sidebar
- Persists run metadata and downloaded Drive images in a local cache for faster reloads
- Includes a sidebar button to clear the saved cache and reload from source

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
.env.example
```

## Setup

1. Create and activate a virtual environment.
2. Install dependencies from `requirements.txt`.
3. Copy `.env.example` to `.env`.
4. Fill in:
   - `GOOGLE_DRIVE_ROOT_FOLDER_ID`
   - `GOOGLE_APPLICATION_CREDENTIALS`
   - `LOCAL_ARCHIVE_ROOT` (optional)
   - `LOCAL_ARCHIVE_AFTER_DAYS` (optional, defaults to `7`)
5. Place your service account JSON at the configured credentials path.
6. Share the relevant Google Drive folders with the service account if needed.

You can also run the app with Streamlit secrets instead of a local `.env` file.
For Community Cloud, use `.streamlit/secrets.toml.example` as your template and paste
the full service account JSON into `GOOGLE_SERVICE_ACCOUNT_JSON`.

## Run

```powershell
streamlit run app.py
```

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

- Recent runs are loaded from Google Drive.
- Runs older than `LOCAL_ARCHIVE_AFTER_DAYS` are loaded from `LOCAL_ARCHIVE_ROOT`.
- This lets you keep the current week in Drive and older runs in a local archive.

## Folder Assumptions

- Structure: `RootFolder / CarrierFolder / RunFolder / files`
- Also supported: `RootFolder / RunFolder / files`
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
- The app can combine Google Drive and a local archive.
- Malformed folders are skipped and shown in the sidebar.
