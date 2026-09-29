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
from infer import count_shelf_levels
result = count_shelf_levels("path/to/photo.jpg")  # {"shelf_count": 4, "confidence": 0.83}
```

or from the command line: `python infer.py path/to/photo.jpg`.

## This is intentionally not wired into shelf-poller-app

The live nightly shelf-count pipeline (`shelf-poller-app/`) still counts
via the Ollama vision-language model, unchanged. Swapping the live job over
to this trained detector is a separate decision to make once
`evaluation/evaluate.py`'s counting-accuracy report shows it's actually
more reliable than the current prompt-based approach on real trolley
photos -- not something to flip on by default the moment a model exists.
