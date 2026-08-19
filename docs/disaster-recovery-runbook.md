# Disaster Recovery Runbook — Office PC Standby

This covers the backup/standby system: an office PC that continuously receives an
encrypted copy of everything Render holds, and can be switched on with one click
to become the live system if Render goes down.

Read this once fully before you need it. During a real outage is not the time to
be reading for the first time.

## 1. One-time setup on the office PC

Yes — the office PC needs a full copy of this repository, because it runs the
exact same app, just pointed at its own local data instead of Render's.

1. **Install prerequisites**: Node 22, Python 3, and Postgres (a local instance —
   see step 5). On Windows, also install `postgresql-client` tools so `pg_restore`
   is on PATH.
2. **Copy the repo** to the office PC (git clone, or copy the folder). You need at
   least: `shadow-app/`, `src/`, `requirements.txt`, `Dockerfile` is not needed
   (you're running Node directly, not Docker, on this PC).
3. **Install dependencies**:
   ```
   cd shadow-app && npm ci
   python -m venv .venv && .venv\Scripts\activate && pip install -r ..\requirements.txt
   ```
4. **Generate the backup keypair** (do this once, only on the office PC):
   ```
   node shadow-app/server/backup/generate-keypair.js
   ```
   This prints a `BACKUP_PUBLIC_KEY` and a `BACKUP_PRIVATE_KEY`.
   - Put `BACKUP_PUBLIC_KEY` in Render's environment variables (Render dashboard,
     since it's marked `sync: false` in `render.yaml` and won't be picked up from git).
   - Keep `BACKUP_PRIVATE_KEY` **only on this PC** — never commit it, never put it
     on Render. Also save a copy somewhere safe (e.g. a password manager) — if
     this key is lost, every backup ever taken becomes unreadable.
5. **Set up local Postgres**: install Postgres, create an empty database, and set
   `DATABASE_URL` (office PC's local env) to point at it. This lets the standby
   fully match Render, including the LLM-poller job queue.
6. **Generate a backup agent key** (any long random string works, e.g.
   `openssl rand -hex 32`), and set it as `BACKUP_AGENT_API_KEY` in **both**
   Render's environment variables and the office PC's local environment — it has
   to be the identical value on both sides.
7. **Set the office PC's environment variables** (e.g. in a `.env` file loaded by
   your process manager, or system environment variables):
   ```
   SNAPPYSJAAK_CACHE_DIR=D:\SnappySjaakBackup\cache      (or wherever you want the data)
   RENDER_BACKUP_BASE_URL=https://<your-render-service>.onrender.com
   BACKUP_AGENT_API_KEY=<the shared key from step 6>
   BACKUP_PRIVATE_KEY=<from step 4, THIS PC ONLY>
   DATABASE_URL=<local Postgres connection string from step 5>
   GOOGLE_APPLICATION_CREDENTIALS=credentials/service_account.json   (copy this file from local dev setup)
   PORT=4174
   ```
8. **Start the app** the same way local dev already does (`npm start` inside
   `shadow-app/`, or the existing `start_shadow_app.bat`). Leave it running.

That's it — this instance now runs continuously in **Backup** mode (the default),
quietly pulling data every 15 minutes. You'll see a "BACKUP MODE" banner across
the app and the two normal background jobs (reminder emails, sheet sync) stay off.

## 2. Day to day

Nothing to do. Every 15 minutes the office PC pulls whatever changed on Render
(JSON state files, uploaded documents) and verifies it; every 6 hours it also
pulls and restores a fresh Postgres dump. If a sync fails repeatedly or goes
stale for more than ~45 minutes, this PC emails the ICT support recipients
directly (it does not rely on Render being reachable to raise that alarm).

Check in on it occasionally: open the app on this PC and look at
Settings > System mode to see when it last changed, and its `/api/health`
endpoint for a quick status snapshot.

## 3. Failover — Render is down, you need to keep working

1. Confirm Render is actually down (try opening its normal URL).
2. On the office PC, open the app (it's already running), go to
   **Settings > System mode**, and click **Online**. Confirm the prompt.
3. The two background jobs (reminders, sheet sync) start running locally, and
   this PC now behaves exactly like Render did.
4. Share the LAN address with whoever else needs it — the console this app was
   started from already prints it (`Open from another PC on the same network:`).
   Everyone on the office Wi-Fi/LAN can use it; there's no remote access to this
   standby from outside the office.
5. Everyone logs in again (sessions aren't shared between instances — this is
   expected, not a bug).
6. Work normally. Fust, UKDocs, clock records etc. only need this PC and the
   internet (for Google Drive/Sheets/Gmail/SMTP) — none of it depends on Render.

**Do not run both Render and the office PC in Online mode at the same time on
purpose.** A safety check will make each side skip its background jobs and email
an alert if it detects the other side also claiming to be Online, but that's a
safety net for the moment Render comes back unexpectedly — not something to rely
on deliberately.

## 4. Render is back — bringing it up to date

1. Confirm Render is genuinely healthy again (it's serving normally).
2. **Immediately switch the office PC's Settings > System mode back to Backup.**
   Do this before anything else, so nobody keeps working on two systems at once.
3. Run the reconciliation tool from the office PC:
   ```
   cd shadow-app/server/backup
   set RENDER_BASE_URL=https://<your-render-service>.onrender.com
   set RECONCILE_USERNAME=<an admin username>
   set RECONCILE_PASSWORD=<that admin's password>
   node reconcile.js
   ```
   This is a **dry run** by default — it reports what would change without
   touching anything. Read the summary:
   - New Fust actions created during the outage → will be **added** to Render.
   - Existing Fust actions edited during the outage (confirmed, document
     attached) → will be **updated** (only filling in what Render is missing;
     it never overwrites something Render already has).
   - `fust-settings.json` and `shadow-users.json` are only **diffed and
     reported** — nothing here is applied automatically, since these hold live
     credentials and account permissions. Review any differences shown and
     decide by hand whether to apply them (e.g. through Settings normally on
     Render, or via `POST /api/backup/reconcile/settings-apply` with just the
     specific fields you approved).
4. Once you're happy with the dry-run report, apply it for real:
   ```
   node reconcile.js --apply
   ```
5. Confirm Render's data now looks right, then resume normal operation on Render.
   The office PC, back in Backup mode, will start pulling fresh data from Render
   again on its usual 15-minute cycle.

## 5. Known limitations (accepted trade-offs)

- Edits to secondary modules (clock records, expedition stickers, dag-foutjes,
  bunches) made on the standby during an outage are **not** auto-merged back —
  only Fust actions are. If those were used during an outage, re-enter that data
  by hand afterward; daily volume there is normally low enough that this is a
  short manual step, not a project.
- The reconciliation tool only looks at what changed after this PC's last
  successful sync from Render — it assumes Render itself made no conflicting
  changes while it was down (true by definition, since it was down).
- Losing `BACKUP_PRIVATE_KEY` means losing the ability to read any backup ever
  taken. Keep a second copy somewhere safe.
