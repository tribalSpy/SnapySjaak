# Steps 3, 4 & 6: Dataset prep, training, and model registry

## Setup

`setup_gpu_pc.bat` (repo root) already ran `pip install -r requirements.txt`
into the shared venv and copied `config.example.json` to `config.json`.

`ultralytics` pulls in PyTorch as a dependency. If `setup_gpu_pc.bat` didn't
prompt you about it (or you skipped that prompt), make sure the CUDA-enabled
PyTorch build is installed, or training silently falls back to CPU:

1. Check the GPU is visible: `nvidia-smi`
2. Get the right install command for your CUDA version from
   https://pytorch.org/get-started/locally/, run it, then re-run
   `pip install -r requirements.txt` in the venv.

Adjust `batch`/`imgsz` in `config.json` to whatever the GPU PC's VRAM can
handle if training runs out of memory.

## Step 3: `prepare_dataset.py`

Converts every `labeling/exports/export_*.json` into a YOLO-format dataset
under `dataset/` (gitignored -- rebuild it any time from the exports):

```
prepare_dataset.bat
```

Run this again after every `labeling/export_annotations.bat` to pick up
newly-labeled photos before the next training run.

## Step 4: `train.py`

```
train.bat
```

Trains an Ultralytics YOLO model (`base_model` in config.json, default
`yolo11n.pt`) on `dataset/data.yaml`, then validates the best checkpoint and
writes `runs/<run_name>/summary.json` (mAP50, mAP50-95, precision, recall).
Prints the run name to use with the next two steps.

## Step 5: evaluate it

See [`../evaluation/README.md`](../evaluation/README.md) -- run this before
deciding whether a model is worth promoting.

## Step 6: `promote_model.py`

Once a run's counting-accuracy report looks good:

```
promote_model.bat --run shelf_level_20260910_120000
```

Copies that run's `best.pt` into `../models/` and marks it active in
`../models/registry.json`. See [`../models/README.md`](../models/README.md)
for how the registry works and how (and whether) to point anything at it.
