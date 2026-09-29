# GPU PC setup plan

Everything the shelf-count system needs on the GPU PC: the two poller apps
(`llm-poller-app`, `shelf-poller-app`) and the training pipeline
(`shelf-training`). One script, `setup_gpu_pc.bat`, does the installable
parts; a short list of manual steps below covers the rest (secrets and
things only a human can verify).

## 1. Get the code onto the GPU PC

That PC doesn't sync this Synology Drive folder, so:

1. Copy `pull_code.bat` (repo root, this one file only) onto the GPU PC --
   USB stick, email, or download it directly from GitHub.
2. Run it there. First run clones the repo (into `.\SnappySjaak` next to
   wherever you saved the script, or pass a target folder as an argument);
   every run after that just pulls the latest changes.
3. If that PC hasn't cloned this repo before, Git will prompt for GitHub
   credentials on first clone/pull (a private repo) -- sign in when asked.

Re-run `pull_code.bat` any time to update the code later; it never
overwrites local changes -- if it can't fast-forward, it tells you to
`git status`/`git stash` first instead of discarding anything.

## 2. Send these separately (never through git, never through chat)

- The repo root `.env` file (Google Drive service-account credentials +
  `GOOGLE_DRIVE_ROOT_FOLDER_ID`). Copy it directly (USB, a secure file
  share) -- it's already gitignored, so it isn't in the repo.
- The value of shadow-app's `SHADOW_LLM_POLLER_API_KEY` (used by both
  poller apps and by `shelf-training/data/config.json` and
  `shelf-training/labeling` isn't gated by it, but `data/collect.py` is).
- Whatever `llm-poller-app`/`shelf-poller-app` `agent_name`/`pc_name`/
  `model_name` values you want this specific PC to identify itself as.

## 3. Run the setup script

Double-click `setup_gpu_pc.bat` (repo root). It:
- Creates one shared virtual environment at `shelf-training\.venv`.
- Installs Google Drive/dotenv, Label Studio, and Ultralytics/PyTorch
  dependencies into it.
- Pauses once to ask whether CUDA-enabled PyTorch is already installed --
  if not, it points you at https://pytorch.org/get-started/locally/ so
  training actually uses the GPU instead of silently falling back to CPU.
- Copies every `config.example.json` to `config.json` next to it, wherever
  one doesn't already exist (never overwrites one that's already set up).
- Warns if `.env` is missing.

It does not touch `llm-poller-app/poller.py` or `shelf-poller-app/poller.py`
themselves -- those only use Python's standard library and need no
dependency installs, matching how they've always been set up.

## 4. Manual steps (secrets + human verification)

1. Fill in `shelf-training\data\config.json`: `server_url`, `api_key`
   (= `SHADOW_LLM_POLLER_API_KEY`), `drive_root_folder_id`.
2. Fill in `llm-poller-app\config.json` and `shelf-poller-app\config.json`
   if this is a new PC (see each folder's own `README.md`).
3. `ollama list` -- confirm the vision model the poller configs reference
   is actually pulled on this machine.
4. Run `shelf-training\labeling\start_label_studio.bat` once, create a
   Label Studio account, copy a Personal Access Token from
   **Account & Settings**, and paste it into
   `shelf-training\labeling\config.json`.

## 5. Start the long-running pieces

- `llm-poller-app\run_poller.bat`
- `shelf-poller-app\run_poller.bat`
- `shelf-training\labeling\start_label_studio.bat` (only while actively
  labeling -- no need to keep it running otherwise)

Consider Windows Scheduled Tasks (`schtasks /Create ...`) so the two
pollers restart automatically on reboot -- see each poller's own README.

## 6. Optional: automate Step 7 (feedback loop)

```
schtasks /Create /TN "ShelfCount Feedback Cycle" /SC WEEKLY /D MON /ST 07:00 ^
  /TR "C:\path\to\SnappySjaak\shelf-training\feedback\run_feedback_cycle.bat"
```

See [`shelf-training/feedback/README.md`](shelf-training/feedback/README.md).

## What's already handled without any of this

- Everything upstream of the GPU PC (RFID portal ingest, Drive uploads,
  shadow-app's nightly trigger, the `llm_jobs` queue) lives on the server
  side and needs no GPU-PC setup at all.
- The live nightly shelf-count job still runs through the Ollama
  vision-language model via `shelf-poller-app` -- the trained YOLO model
  this pipeline produces is not wired into that job. See
  [`shelf-training/models/README.md`](shelf-training/models/README.md).
