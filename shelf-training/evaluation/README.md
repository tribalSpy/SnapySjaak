# Step 5: Counting-accuracy report

Ultralytics' own validation metrics (mAP50, precision, recall) describe how
well the model localizes shelf levels -- not directly how often the count
it produces is the *right number*, which is what the nightly pipeline
actually cares about. This re-runs the val split and compares predicted vs.
ground-truth box counts per image.

```
evaluate.bat --run shelf_level_20260910_120000
```

Writes:
- `reports/eval_<run>_<timestamp>.csv` -- one row per val image (predicted
  count, ground-truth count, absolute difference).
- `reports/eval_<run>_<timestamp>_summary.json` -- exact-match rate, mean
  absolute error, RMSE, max difference.

Use this to decide whether a run is actually worth promoting with
`../training/promote_model.py`, and to compare successive runs after a
`../feedback/run_feedback_cycle.py` round of new labels.
