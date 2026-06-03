# SnappySjaak Shadow App

React + Node version of the dashboard. It sits next to the Streamlit app and reads the same cache files.

## Run

Install once:

```powershell
cd shadow-app
npm install
```

Build the React app:

```powershell
npm run build
```

Start the local Node server:

```powershell
npm start
```

Open:

```text
http://127.0.0.1:4174
```

## Notes

- The Streamlit app is unchanged.
- Local archive images are served by the Node backend.
- Rebuild and refresh date start the existing `sync_index.py` script.
- Opening the app starts a throttled background refresh for today, or for the selected date.
- Google Drive image downloading still lives in the Python app. The shadow app is mainly for comparing the interface and local archive workflow.
- The first browser visit creates the first admin account. Users are stored in `SNAPPYSJAAK_CACHE_DIR/shadow-users.json` or `.cache/shadow-users.json` by default, with hashed passwords.
- Admins can manage users from the sidebar Users page.

## Deploy On Render

This app is a mixed Node + Python service, so the easiest Render setup is a Docker web service.

Environment variables to set:

- `GOOGLE_DRIVE_ROOT_FOLDER_ID`
- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `GOOGLE_DRIVE_ACCOUNT_NAMES` if you use extra Google Drive accounts
- `GOOGLE_DRIVE_ROOT_FOLDER_ID_<NAME>` for each extra account
- `GOOGLE_SERVICE_ACCOUNT_JSON_<NAME>` only if an extra account needs different credentials
- `SNAPPYSJAAK_CACHE_DIR=/var/data/snappysjaak-cache`
- `PYTHON=python3`
- `TRIGGER_POLLER_ON_SYNC=0`
- `LOCAL_ARCHIVE_ROOT=` to disable the old local archive path in production
- `SHADOW_USERS_SEED_PATH=/etc/secrets/shadow-users.json` if you want to seed existing users from a Render Secret File

Recommended Render settings:

- Use the repo root as the Docker build context
- Attach a persistent disk
- Mount it at `/var/data`
- Keep `SNAPPYSJAAK_CACHE_DIR` inside that mount path

With a persistent disk, these survive restarts and deploys:

- saved user accounts
- generated run index cache
- downloaded Google Drive image cache

Without a persistent disk, you can still seed users from a Render Secret File:

- create a secret file named `shadow-users.json`
- paste the contents of your local `.cache/shadow-users.json`
- set `SHADOW_USERS_SEED_PATH=/etc/secrets/shadow-users.json`

On first startup without a local users file, the app will copy those seeded users into its normal cache file.
