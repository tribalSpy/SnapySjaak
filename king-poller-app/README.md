# King poller

Delivers the **Inkoop Controle → Import naar King** exports from the shadow-app
(on Render) into the folder King ERP's scheduler reads on the office network.
It only polls for `king_export` jobs; no other poller ever gets those.

## What it does per export
1. Writes each invoice PDF (`<invoice number without dots>.pdf`, e.g.
   `049876FC20260181.pdf`) into `pdf_dir`.
2. Writes the archive XML (`archief_<batch>.xml`, `KING_DIGITAAL_ARCHIEF`) into
   `import_dir`.
3. Waits until King has read it (the file disappears), up to
   `archive_wait_minutes`.
4. Writes the journal XML (`journaal_<batch>.xml`, `KING_JOURNAAL`) into
   `import_dir`. King links every journal post to its PDF through
   `JR_ARCHIEFSTUK_EXTERN_ID` = `DAR_EXTERN_ID`.

Files are written as `*.part` and renamed when complete, so King never reads a
half-written file. The app shows each export as queued → delivered (or failed
with the reason).

## Setup (once, on an always-on office PC with access to the King folder)
1. Python 3 (standard library only, nothing to install).
2. Copy `config.example.json` to `config.json` and fill in:
   - `server_url`: the shadow-app URL.
   - `api_key`: the same value as `SHADOW_LLM_POLLER_API_KEY` on Render.
   - `import_dir`: the folder King's scheduler checks for XML files.
   - `pdf_dir`: where the PDFs go. In the app's **King settings**, set
     "King PDF folder" to this same folder **as King's server sees it**, because
     the archive XML tells King to read the PDF from that path.
   - `archive_wait_minutes`: how long to wait for King to read the archive XML
     before writing the journal (King checks about every 15 minutes).
3. Run `run_poller.bat`. To start it automatically:
   `schtasks /Create /TN "King poller" /TR "%CD%\run_poller.bat" /SC ONLOGON`
