# Shelf Count Training Pipeline

Trains the shelf/level-counting model that `shelf-poller-app/` runs, using
real trolley photos from Google Drive and real results already recorded by
`shadow-app`.

All 7 steps are built. Run `setup_gpu_pc.bat` (repo root) on the GPU PC
first -- see [Setup on the GPU PC](#setup-on-the-gpu-pc) below.

## Folder structure

- `data/` -- Step 1: dataset collection.
- `labeling/` -- Step 2: Label Studio setup + import/export scripts.
- `training/` -- Steps 3, 4 & 6: dataset prep, training, model registry.
- `evaluation/` -- Step 5: counting-accuracy reports.
- `models/` -- registry + inference helper (filled in by Step 6).
- `feedback/` -- Step 7: pulls newly-flagged needs_review photos back into labeling.

## Pipeline, end to end

```
data/collect.py            Step 1  -- download photos (priority + general sample)
labeling/import_tasks.py   Step 2  -- push them into Label Studio
   (label in the browser)
labeling/export_annotations.py     -- pull finished boxes back out
training/prepare_dataset.py Step 3 -- convert to a YOLO dataset (train/val split)
training/train.py          Step 4  -- train
evaluation/evaluate.py     Step 5  -- counting-accuracy report on the val set
training/promote_model.py  Step 6  -- register the model as active
feedback/run_feedback_cycle.py  Step 7  -- (recurring) queue new needs_review photos for labeling
```

Each script's own README covers its details. Retraining is always a manual
decision (run Step 3 → 6 again) once enough new labels exist -- nothing in
this pipeline retrains or redeploys automatically.

**The trained model is not wired into the live nightly pipeline.**
`shelf-poller-app/` still counts via the Ollama vision-language model,
unchanged. See [`models/README.md`](models/README.md) for why that's a
deliberate, separate decision.

## Setup on the GPU PC

Run `setup_gpu_pc.bat` from the repo root (see the repo root's
`GPU_PC_SETUP.md` for the full plan and what to send). It creates a shared
virtual environment (`shelf-training/.venv`) and installs every step's
dependencies, and scaffolds `config.json` files from each step's
`config.example.json` (never overwriting ones that already exist). It does
**not** fill in secrets (Drive credentials, API keys, the Label Studio
token) -- those still need a human to type them in once.

## Step 1: Dataset collection (`data/collect.py`)

Downloads a sample of trolley photos into `data/dataset/raw/<folder>/`,
tracked in `data/dataset/manifest.csv`. Two tiers, run together in one pass:

1. **Priority**: every reference shadow-app has already flagged
   `needs_review` (low model confidence) in the given date range -- these
   are the hard cases worth labeling first. Fetched from shadow-app itself
   (`GET /api/shelf-count/dataset-candidates`), not by scanning Drive.
   (There is no "deviation" tier yet -- that needs `expected_average` to
   exist, which isn't populated until a later phase.)
2. **General**: a stratified random sample of the remaining target count,
   spread evenly across customer codes, drawn directly from every Drive
   folder dated in range (regardless of whether shadow-app ever processed
   it).

Duplicate photos (by content hash, not filename) are skipped automatically,
so re-running `collect.py` later with an overlapping date range only adds
what's genuinely new.

### Setup

After running `setup_gpu_pc.bat` (which creates `data/config.json` from the
example), fill in:
- `server_url` / `api_key` -- same value as shadow-app's
  `SHADOW_LLM_POLLER_API_KEY`.
- `drive_root_folder_id` -- same value as shadow-app's
  `GOOGLE_DRIVE_ROOT_FOLDER_ID`.

Also make sure the repo root's `.env` has `GOOGLE_SERVICE_ACCOUNT_JSON` set
(the same credential `drive_bridge.py`/`src/drive_service.py` already use)
-- this script loads that same `.env` file, no separate credential setup
needed.

### Run it

```
data\collect.bat --from-date 2026-09-01 --to-date 2026-09-29 --count 200
```
(or `python data/collect.py ...` directly, once the venv is activated)

- `--count` is the GENERAL tier's target; priority (needs_review) photos are
  always collected in full on top of that, since there should never be many
  of them at once.
- `--customer-references A,B,C` restricts the general tier to specific
  customer codes (the priority tier is never restricted -- a hard case is a
  hard case regardless of customer).

Output: `data/dataset/raw/<folder_name>/<photo>.jpg` plus
`data/dataset/manifest.csv` (file_id, folder_name, customer_reference, date,
local_path, sha256, priority, status, confidence) -- this manifest is what
`labeling/import_tasks.py` reads from next.

## Step 2: Labeling (`labeling/`)

Local Label Studio instance on the GPU PC. See
[`labeling/README.md`](labeling/README.md) for setup and the day-to-day
import → label → export workflow.

## Steps 3, 4 & 6: Dataset prep, training, model registry (`training/`)

See [`training/README.md`](training/README.md). In short:
`prepare_dataset.py` (Step 3) turns Label Studio's export into a YOLO
dataset, `train.py` (Step 4) trains on it, `promote_model.py` (Step 6)
registers a run you're happy with into `models/`.

## Step 5: Evaluation (`evaluation/`)

See [`evaluation/README.md`](evaluation/README.md) -- a counting-accuracy
report (exact-match rate, mean absolute error) on the trained model's val
split, run before deciding whether to promote it.

## Step 7: Feedback loop (`feedback/`)

See [`feedback/README.md`](feedback/README.md) -- a recurring job that
pulls freshly-flagged `needs_review` photos back into the labeling queue,
so accuracy keeps improving without anyone manually hunting for hard cases.
