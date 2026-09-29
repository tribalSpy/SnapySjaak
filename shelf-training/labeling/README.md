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

## Notes

- `imported_tasks.json` and `exports/` are local working state, not meant
  to be committed (see `.gitignore`).
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
