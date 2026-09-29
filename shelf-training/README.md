# Shelf Count Training Pipeline

Trains the shelf/level-counting model that `shelf-poller-app/` runs, using
real trolley photos from Google Drive and real results already recorded by
`shadow-app`.

## Folder structure

- `data/` -- Step 1: dataset collection.
- `labeling/` -- Step 2: Label Studio setup + import/export scripts.
- `training/` -- Step 4: Ultralytics YOLO training scripts (not built yet).
- `evaluation/` -- Step 5: counting-accuracy reports (not built yet).
- `models/` -- Step 6: trained model registry (not built yet).

(Step 3, dataset prep -- converting Label Studio's export into YOLO-format
labels + a train/val split -- is also not built yet; it's the next thing to
add, sitting between `labeling/` and `training/`.)

Each step is being implemented one at a time, per its own brief.

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

1. Install Python 3.11+.
2. From this folder: `pip install -r requirements.txt`
3. Copy `data/config.example.json` to `data/config.json` and fill in:
   - `server_url` / `api_key` -- same value as shadow-app's
     `SHADOW_LLM_POLLER_API_KEY`.
   - `drive_root_folder_id` -- same value as shadow-app's
     `GOOGLE_DRIVE_ROOT_FOLDER_ID`.
4. Make sure the repo root's `.env` has `GOOGLE_SERVICE_ACCOUNT_JSON` set
   (the same credential `drive_bridge.py`/`src/drive_service.py` already
   use) -- this script loads that same `.env` file, no separate credential
   setup needed.

### Run it

```bash
python data/collect.py --from-date 2026-09-01 --to-date 2026-09-29 --count 200
```

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
