# Step 2: Labeling (Label Studio)

Runs on the GPU PC (same machine as `shelf-poller-app/`), so labeled photos
never leave the LAN and there's nothing to sync.

Annotation task: draw one bounding box around **each visible shelf level**
in a photo, labeled `shelf_level`. Counting boxes per image is what turns
into `shelf_count`/`level_count` at inference time -- this is why the model
is an Ultralytics YOLO **detector**, not a classifier.

## Setup (one-time, on the GPU PC)

`setup_gpu_pc.bat` (repo root) already installed Label Studio into the
shared venv and copied `config.example.json` to `config.json` for you.
What's left is the part that needs a human:

1. Start Label Studio: double-click `start_label_studio.bat` (it activates
   the venv, points `LOCAL_FILES_DOCUMENT_ROOT` at `data/dataset`, and runs
   `label-studio start`).
2. Open `http://localhost:8080`, create your account, then go to
   **Account & Settings > Personal Access Token** and copy it.
3. Paste that token into `config.json`'s `api_token`. Leave `project_id` as
   `null` the first time -- `import_tasks.py` creates the project for you
   and fills it in.

## Day-to-day workflow

1. `start_label_studio.bat` (leave it running).
2. Collect more photos: `..\data\collect.bat --from-date ... --to-date ...`
3. Push any newly-collected photos into Label Studio as tasks:
   ```
   import_tasks.bat
   ```
   Safe to re-run any time -- it only imports rows from `manifest.csv` it
   hasn't sent before (tracked in `imported_tasks.json`).
4. Label in the browser at `http://localhost:8080` -- open the project,
   draw a `shelf_level` box around every visible shelf level in each photo,
   submit.
5. Pull finished annotations back out whenever you want a snapshot to hand
   to Step 3 (dataset prep):
   ```
   export_annotations.bat
   ```
   Writes a dated file to `exports/export_<timestamp>.json` -- Label
   Studio's native JSON format (image path + one entry per drawn box, in
   percentage coordinates). `../training/prepare_dataset.py` (Step 3)
   converts this into YOLO-format labels and a train/val split.

## Pre-labeling with a trained model (faster labeling)

Once a model exists, let it draw the boxes first and only correct them:

```
prelabel_tasks.bat --run shelf_level_20261005_110950
```

(or without `--run` once a model is promoted). For every task with no
annotation yet it runs the model and stores its boxes in Label Studio as a
*prediction* (`model_version` = the run name). Open the task, fix the boxes
(move, resize, delete wrong ones, add missed shelves) and **Submit** -- only
submitted annotations are exported and trained on, predictions never are.

Options: `--conf 0.3` (only surer boxes), `--limit 50` (first 50 tasks),
`--overwrite` (also tasks that already got predictions, e.g. from an older
model), `--min-shelf-height 1.5` (make predicted shelf boxes at least 1.5%
of the photo height, the thicker shelf-board-plus-front-edge style).

If the boxes don't show up when you open a task: project **Settings >
Annotation**, enable showing predictions to annotators.

## Notes

- `imported_tasks.json` and `exports/` are local working state, not meant
  to be committed (see `.gitignore`).
- Setting `LOCAL_FILES_DOCUMENT_ROOT`/`LABEL_STUDIO_LOCAL_FILES_SERVING_ENABLED`
  (what `start_label_studio.bat` does) is necessary but **not sufficient** on
  this Label Studio version -- confirmed from its own source: every
  `/data/local-files/` request 404s unless a Local Storage connection is
  also registered for the project. `import_tasks.py` now does this for you
  automatically (`ensure_local_storage`, registered at `data/dataset/raw`,
  since Label Studio refuses a storage whose path is the same as
  `LOCAL_FILES_DOCUMENT_ROOT` itself) -- if photos still don't load after
  re-running `import_tasks.bat`, check **Settings > Cloud Storage** in the
  project UI for that entry, or check the server console log for the exact
  404'd path.
- The token from **Account & Settings > Personal Access Token** is a JWT
  *refresh* token on current Label Studio versions, not a static API key --
  `_auth.py` exchanges it for a short-lived access token (`/api/token/refresh`)
  before every run of `import_tasks.py`/`export_annotations.py`. If that
  exchange 404s, this Label Studio instance is old enough to still expect a
  static `Authorization: Token <token>` header instead; open an issue/ask
  before reverting, since which one applies depends on the installed version.
- If a Label Studio upgrade changes the export API shape,
  `export_annotations.py` prints instructions to use the UI's Export button
  as a fallback.
