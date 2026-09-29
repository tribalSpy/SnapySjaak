# Step 7: Feedback loop

Run `run_feedback_cycle.py` on a recurring schedule (weekly is reasonable)
to pull freshly-flagged `needs_review` photos into the Label Studio queue
automatically:

```
run_feedback_cycle.bat --days-back 14
```

To automate it, add a Windows Scheduled Task on the GPU PC that runs
(adjust the path to wherever this repo lives on that machine):

```
schtasks /Create /TN "ShelfCount Feedback Cycle" /SC WEEKLY /D MON /ST 07:00 ^
  /TR "C:\path\to\SnappySjaak\shelf-training\feedback\run_feedback_cycle.bat"
```

This only queues photos for labeling -- it never retrains or redeploys
anything by itself. Retraining is a deliberate, manual step (see the
top-level [README](../README.md)) once enough new labels have piled up.
