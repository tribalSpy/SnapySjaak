# Disaster Recovery Runbook — Office PC Standby

This covers the backup/standby system: an office PC that continuously receives an
encrypted copy of everything Render holds, and can be switched on with one click
to become the live system if Render goes down.

Read this once fully before you need it. During a real outage is not the time to
be reading for the first time.

## 1. One-time setup on the office PC

Yes — the office PC needs a full copy of this repository, because it runs the
exact same app, just pointed at its own local data instead of Render's.

1. **`git clone` the repo** to the office PC — use a clone specifically, not a
   plain folder copy, so `update_standby_pc.bat` (section 3) can pull future
   code changes. Cloning won't bring `credentials/service_account.json`
   (it's deliberately excluded from git) — copy that file over separately
   from your local dev setup, into the same relative path in the clone.
2. **Run `setup_standby_pc.bat`** (repo root, double-click it). It will:
   - Check for Node.js, Python, and PostgreSQL, and offer to install anything
     missing via `winget`. PostgreSQL's installer needs a superuser password
     set interactively, so that one step isn't silent — remember whatever
     password you set, you'll need it in step 4 below.
   - Install the app's Node and Python dependencies.
   - Generate the backup secrets and write a **`start_standby.bat`** at the
     repo root with two of them already filled in. **Copy the printed
     `BACKUP_PUBLIC_KEY` and `BACKUP_AGENT_API_KEY` now** — the console output
     is the only time the private key is shown in full (it's also saved into
     `start_standby.bat` for you, so you don't need to copy that one).
   - Running the script again later won't regenerate these secrets — it
     detects `start_standby.bat` already exists and leaves it alone, so you
     can't accidentally invalidate a key you've already given Render.
3. **Add to Render's dashboard environment variables**: paste in the
   `BACKUP_PUBLIC_KEY` and `BACKUP_AGENT_API_KEY` the script just printed.
   (`PGDUMP_MIN_INTERVAL_MINUTES` already defaults on its own — nothing to do there.)
4. **Create an empty Postgres database** (in the Postgres you just installed)
   and note its connection string — this lets the standby fully match Render,
   including the LLM-poller job queue.
5. **Open `start_standby.bat`** and fill in the two placeholders:
   `RENDER_BACKUP_BASE_URL` (your Render service's URL) and `DATABASE_URL`
   (from step 4). Everything else in that file is already set.
6. **Double-click `start_standby.bat`** and leave it running.

That's it — this instance now runs continuously in **Backup** mode (the default),
quietly pulling data every 15 minutes. You'll see a "BACKUP MODE" banner across
the app and the two normal background jobs (reminder emails, sheet sync) stay off.

Also save a second copy of the private key (shown once during setup, and stored
in `start_standby.bat`) somewhere safe, e.g. a password manager — losing it means
losing the ability to read every backup ever taken.

## 2. Day to day

Nothing to do. Every 15 minutes the office PC pulls whatever changed on Render
(JSON state files, uploaded documents) and verifies it; every 6 hours it also
pulls and restores a fresh Postgres dump. If a sync fails repeatedly or goes
stale for more than ~45 minutes, this PC emails the ICT support recipients
directly (it does not rely on Render being reachable to raise that alarm).

Check in on it occasionally: open the app on this PC and look at
Settings > System mode to see when it last changed, and its `/api/health`
endpoint for a quick status snapshot.

## 3. Updating the standby when the app changes

When you push code changes (bug fixes, new features) to the main app, the
office PC doesn't get them automatically — it's running its own local copy.

1. Close the window running `start_standby.bat` (or press Ctrl+C in it).
2. Double-click `update_standby_pc.bat`. It pulls the latest code (`git pull`)
   and reinstalls Node/Python dependencies if anything changed. Your data
   (`standby-cache\`, `start_standby.bat`, `credentials\`) is untouched.
3. Double-click `start_standby.bat` again to resume.

This only works if the office PC was set up via `git clone` (section 1, step 1)
rather than a plain folder copy — a plain copy has no way to pull updates.

## 4. Failover — Render is down, you need to keep working

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

## 5. Render is back — bringing it up to date

1. Confirm Render is genuinely healthy again (it's serving normally).
2. **Immediately switch the office PC's Settings > System mode back to Backup.**
   Do this before anything else, so nobody keeps working on two systems at once.
   Once you do, a reminder banner appears on the office PC's Settings page —
   it stays until you actually run reconciliation, and switching back to Online
   while it's showing asks for an extra confirmation, so it's hard to miss.
3. Run reconciliation — two ways to do this, same result either way:
   - **From the Settings page (easiest)**: on the office PC, scroll to the
     "Reconciliation" card, fill in the Render base URL and an admin
     username/password, and click **Dry run**. Review the summary — new Fust
     actions show as "added", edited ones as "updated" (only filling in what
     Render is missing; it never overwrites something Render already has).
     `fust-settings.json`/`shadow-users.json` differences are only reported,
     never applied automatically, since they hold live credentials and account
     permissions — review those by hand. Once you're happy, click **Apply**.
   - **From the command line**, if you prefer:
     ```
     cd shadow-app/server/backup
     set RENDER_BASE_URL=https://<your-render-service>.onrender.com
     set RECONCILE_USERNAME=<an admin username>
     set RECONCILE_PASSWORD=<that admin's password>
     node reconcile.js
     node reconcile.js --apply
     ```
4. Confirm Render's data now looks right, then resume normal operation on Render.
   The office PC, back in Backup mode, will start pulling fresh data from Render
   again on its usual 15-minute cycle.

## 6. Known limitations (accepted trade-offs)

- Edits to secondary modules (clock records, expedition stickers, dag-foutjes,
  bunches) made on the standby during an outage are **not** auto-merged back —
  only Fust actions are. If those were used during an outage, re-enter that data
  by hand afterward; daily volume there is normally low enough that this is a
  short manual step, not a project.
- The reconciliation tool only looks at what changed after this PC's last
  successful sync from Render — it assumes Render itself made no conflicting
  changes while it was down (true by definition, since it was down).
- Losing the private key (`start_standby.bat`'s `BACKUP_PRIVATE_KEY`) means
  losing the ability to read any backup ever taken. Keep a second copy
  somewhere safe.
- `setup_standby_pc.bat` refuses to regenerate secrets once `start_standby.bat`
  exists, specifically to prevent silently invalidating a key already given to
  Render. If you genuinely need to rotate keys, delete `start_standby.bat`,
  re-run the setup script, and update `BACKUP_PUBLIC_KEY`/`BACKUP_AGENT_API_KEY`
  on Render to match the newly printed values.
