# Shelf Count Poller

A separate, dedicated poller for exactly one job type: `shelf_count`. It
talks to the same Shadow App job queue as `llm-poller-app/poller.py`, but is
its own independent process -- deliberately kept apart so the existing
invoice-PDF/CSI poller is never touched or put at risk by shelf-counting
changes.

It can run on the same PC (and against the same local Ollama instance) as
the other poller, or on its own machine -- both are just HTTP clients of the
same Render app.

## Why a second poller, and why it's safe to run alongside the other one

Every poll this app sends includes `job_types: ["shelf_count"]`. The server
only ever hands a `shelf_count` job to a poller that asks for it by name, so:

- This poller can never accidentally claim (and fail) an `excel_to_pdf` or
  `ukdocs_csi_audit` job meant for the other poller.
- The other poller (which never asks for `shelf_count`) can never
  accidentally claim one of these either.

## Setup on the GPU PC

1. Install Python 3.11 or newer (if not already installed for the other
   poller).
2. Make sure Ollama is running locally with a vision-capable model pulled
   (the same one `llm-poller-app` already uses is fine).
3. Copy this whole `shelf-poller-app` folder onto the PC.
4. Copy `config.example.json` to `config.json`.
5. Fill in `server_url`, `api_key` (same value as `SHADOW_LLM_POLLER_API_KEY`
   on the server), a unique `agent_name` (must be different from the other
   poller's), and `model_name`.
   **Run `ollama list` first and copy the exact name from there** (e.g.
   `qwen3-vl:32b`) -- confirmed real: `config.example.json`'s own
   `model_name` used to be a placeholder, `gwen32b-vision`, that was never
   an actual installed model anywhere. Left as-is, every job failed
   instantly with `{"error":"model 'gwen32b-vision' not found"}`, and
   nothing in the Shelf Count page said why until `error_text` was added
   to that table. A model without `-vl`/`vision` in its name (most of what
   `ollama list` shows on a general-purpose PC) cannot take image input at
   all, regardless of its name matching -- it has to be a real
   vision-capable pull.
6. Start it with `run_poller.bat`.

## Optional: trained model side by side

Set `trained_model_python` (e.g. `...\shelf-training\.venv\Scripts\python.exe`)
and `trained_model_infer` (e.g. `...\shelf-training\models\infer.py`) in
`config.json` to also count every job's photos with the trained YOLO model
(the active one in `shelf-training/models/registry.json`). Its result is
sent along for comparison only; it can never fail a job. Leave both empty to
switch it off.

## What this does

- Sends heartbeat updates to Shadow (capabilities always `["shelf_count"]`).
- Polls only for `shelf_count` jobs.
- Runs the job's prompt + attached trolley photos against local Ollama.
- Sends the raw model response back -- shadow-app's own job-completion
  handler parses the JSON result (shelves/levels/confidence) and updates
  the `shelf_counts` table.

## Files

- `poller.py`
- `config.example.json`
- `run_poller.bat`
