# Step 2: Labeling (Label Studio)

Runs on the GPU PC (same machine as `shelf-poller-app/`), so labeled photos
never leave the LAN and there's nothing to sync.

Annotation task: draw one bounding box around **each visible shelf level**
in a photo, labeled `shelf_level`. Counting boxes per image is what turns
into `shelf_count`/`level_count` at inference time -- this is why the model
is an Ultralytics YOLO **detector**, not a classifier.

## Setup (one-time, on the GPU PC)

1. Install Label Studio and this folder's scripts' dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Point Label Studio at the dataset folder and turn on local file serving,
   then start it (PowerShell):
   ```powershell
   $env:LABEL_STUDIO_LOCAL_FILES_SERVING_ENABLED = "true"
   $env:LOCAL_FILES_DOCUMENT_ROOT = (Resolve-Path ..\data\dataset).Path
   label-studio start
   ```
   (bash equivalent: `export LABEL_STUDIO_LOCAL_FILES_SERVING_ENABLED=true; export LOCAL_FILES_DOCUMENT_ROOT=$(realpath ../data/dataset); label-studio start`)
3. Open `http://localhost:8080`, create your account, then go to
   **Account & Settings > Personal Access Token** and copy it.
4. Copy `config.example.json` to `config.json` and fill in `api_token`.
   Leave `project_id` as `null` the first time -- `import_tasks.py` creates
   the project for you and fills it in.

`dataset_dir` in `config.json` must point at the same folder as
`LOCAL_FILES_DOCUMENT_ROOT` above (default `../data/dataset`, i.e.
`shelf-training/data/dataset` -- the folder `data/collect.py` fills).

## Day-to-day workflow

1. Collect more photos: `python ../data/collect.py --from-date ... --to-date ...`
2. Push any newly-collected photos into Label Studio as tasks:
   ```bash
   python import_tasks.py
   ```
   Safe to re-run any time -- it only imports rows from `manifest.csv` it
   hasn't sent before (tracked in `imported_tasks.json`).
3. Label in the browser at `http://localhost:8080` -- open the project,
   draw a `shelf_level` box around every visible shelf level in each photo,
   submit.
4. Pull finished annotations back out whenever you want a snapshot to hand
   to Step 3 (dataset prep):
   ```bash
   python export_annotations.py
   ```
   Writes a dated file to `exports/export_<timestamp>.json` -- Label
   Studio's native JSON format (image path + one entry per drawn box, in
   percentage coordinates). Nothing consumes this yet; Step 3 will convert
   it into YOLO-format label files and a train/val split.

## Notes

- `imported_tasks.json` and `exports/` are local working state, not meant
  to be committed (see `.gitignore`).
- If a Label Studio upgrade changes the export API shape,
  `export_annotations.py` prints instructions to use the UI's Export button
  as a fallback -- this script was written against Label Studio 1.23.1's
  documented `/api/projects/{id}/export` endpoint but hasn't been run
  against a live instance yet, since none exists in this environment.
