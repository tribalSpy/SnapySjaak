# Step 6: Model registry

`registry.json` tracks every model version promoted by
`training/promote_model.py` -- version name, base model, validation metrics
(mAP50, mAP50-95, precision, recall), when it was trained/promoted, and
which one is currently `active`. It's small and text-only, so it's tracked
in git.

The actual weight files (`*.pt`) are **not** tracked in git (see
`../.gitignore`) -- they're tens of MB and belong on the GPU PC, not in the
app repo. If you need to move a model to another machine, copy the `.pt`
file referenced by the active registry entry's `weights_path` by hand (or
zip the whole `models/` folder).

## Using the active model

```python
from infer import count_photo
count_photo("path/to/photo.jpg")  # {"shelf_level": 4, "extension": 0, "confidence": 0.83}
```

From the command line: `python infer.py path/to/photo.jpg`, or
`python infer.py --batch photo1.jpg photo2.jpg ...` for one JSON document
covering several photos (what `shelf-poller-app` runs).

## Side by side with the live count

`shelf-poller-app/` still produces the official nightly count with the
Ollama vision-language model. Once `trained_model_python` (this folder's
`.venv\Scripts\python.exe`) and `trained_model_infer` (this `infer.py`) are
set in the poller's `config.json`, it also runs the **active** model on the
same photos and sends that along. The Shelf count page shows it in the
"Trained model (shelves / ext.)" column, orange when it disagrees with the
official count, with the per-photo counts in the tooltip (reference total =
median photo count x trolley count). While no model is promoted, that
column just shows "error" (hover: "No active model...") and nothing else
changes.

Switching the official count over to this model is a later, separate
decision, once the side-by-side column has shown it is right more often
than Ollama on real nights.
