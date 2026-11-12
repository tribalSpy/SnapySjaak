# Media poller (BEAST PC)

Makes the Campagnes photos and videos on the BEAST PC (RTX 5090). SnappySjaak on Render only queues the work. This app claims `media_*` jobs and does all the heavy work here:

- reads the photos from Google Drive;
- draws the frames with Pillow and encodes them with ffmpeg (NVENC on the RTX);
- writes French and English captions with the local Ollama model;
- uploads the results to Drive.

It sends back only ids and texts. The other pollers never take `media_*` jobs.

## Setup

0. **Easiest:** run `setup_media_pc.bat` (repo root). It does steps 1–3 and, if you want, the AI part.
1. **ffmpeg:** install a build with NVENC (e.g. `winget install Gyan.FFmpeg`) so `ffmpeg` is on PATH, or set `ffmpeg_path`. The poller checks NVENC with a test encode at start and falls back to libx264.
2. **Config:** copy `config.example.json` to `config.json` and fill in:
   - `server_url`, `api_key` (same as `SHADOW_LLM_POLLER_API_KEY` on Render), `agent_name`;
   - `drive`: either a service-account JSON or a Google user's OAuth refresh token.
     - **Service account:** it has no storage of its own, so put the Social folders in a **Shared Drive** and add the service account as *Content manager*.
     - **OAuth:** use it when the folders are in a normal My Drive.
   - `caption_model`: the Ollama model for captions (e.g. the one already on this PC). Leave it empty for simple captions.
   - `music_dir`: a folder of royalty-free music. A theme or item names a file from it.
3. **Start:** run `run_poller.bat`. It makes `.venv` with Python 3.12 and keeps the poller running.

## Drive folders (set in Campagnes > Settings)

- **Library:** the Social folder where people drop HD photos. Subfolders become tags, e.g. `Plant of the Month/Calathea`. Folders that start with `_` are skipped.
- **Output:**
  - finished photos and videos go in `<Output>/<Theme>/<yyyy-mm>/`;
  - library thumbnails go in `<Output>/_thumbs/`.
  - Both are shared "anyone with the link" so the app can preview them without the files passing through Render.

## Try a recipe locally (no server, no Drive)

```
python poller.py --test-recipe zoom_in --photo C:\photos\raphis.jpg --title "Raphis" --text-fr "Palmier Raphis" --text-en "Lady palm"
python poller.py --test-recipe slideshow --photo a.jpg --photo b.jpg --photo c.jpg --variants 9x16,4x5 --duration 12
python poller.py --test-recipe catalogue_card --photo a.jpg --title "Raphis" --variants 4x5
```

The output goes in `test_output/`. Add `--captions` to also try the Ollama captions.

**Recipes:**

| Recipe | What it makes |
| --- | --- |
| `zoom_in` | video, slow zoom in |
| `zoom_out` | video, slow zoom out |
| `pan` | video, pan across the photo |
| `detail_tour` | video, zooms into details; points can be given |
| `slideshow` | video, several photos with crossfades |
| `offer_reel` | video, one name card per article |
| `photo_post` | photo, framed with text |
| `carousel` | photos, one card per photo |
| `catalogue_card` | photo, white card with name FR/EN and the conditions |

**Formats:** `9x16` (Reels/Shorts/TikTok, 1080×1920), `4x5` (feed, 1080×1350), `1x1` (1080×1080).

## AI recipes (RTX 5090)

`setup_media_pc.bat`, step 6, installs the AI part: torch for CUDA 12.8 and diffusers. It checks the GPU, switches `ai.enabled` on and can download the models (about 40 GB, into the Hugging Face cache; set `HF_HOME` first to use another disk).

| Recipe | Model | What it makes |
| --- | --- | --- |
| `ai_flower_drop` | Wan 2.2 TI2V-5B (Apache-2.0) | petals and flowers falling around the product, 2× slow motion (ffmpeg motion interpolation), about 10 s |
| `ai_living` | Wan 2.2 | the photo gently comes alive |
| `ai_art` | SDXL + ControlNet canny (OpenRAIL++) | pencil sketch / watercolour / oil that keeps the photo's lines; a morph video plus the art stills |

- **Text:** every AI video gets the same FR/EN text, brand name and formats as the other recipes.
- **Licences:** all three models allow commercial use. Flux-dev is not used (non-commercial).
- **Night window:** AI jobs wait for `ai.window` (default `19:00-07:00`), so they don't fight Ollama or shelf-training for the GPU. "Make it now" in the app skips the wait. Ollama is unloaded before an AI job, and the GPU is freed after it.
- **Something not working?**
  - `python ai.py` checks torch, CUDA and the GPU.
  - `python poller.py --test-recipe ai_flower_drop --photo foto.jpg --title "Viola"` tries a real AI render.
  - If Wan 2.2 5B won't load as image-to-video in your diffusers version, run `pip install -U diffusers`. Or set `ai.video_model` to `Wan-AI/Wan2.2-I2V-A14B-Diffusers` with `ai.cpu_offload: true` (bigger, slower, needs 64 GB+ RAM).

## Posting to Instagram and Facebook (automatic, after approval)

SnappySjaak queues an approved post when it is due. The BEAST (`media_publish` job) then does the posting:

1. downloads the video or photos from Drive;
2. gets the page token from SnappySjaak right before posting;
3. posts with Meta's Graph API (`meta_publish.py`):
   - **Instagram:** a Reel with a resumable upload, or a photo / carousel. Photos go up first as temporary unpublished Page photos, because Instagram fetches photos from a link; those temporary photos are deleted afterwards.
   - **Facebook Page:** a Reel for 9:16, a page video, or one post with several photos.

The post link comes back to the app.

**Stats:** twice a day a `media_stats` job reads views, reach, likes, comments, shares and saves of the last 35 days of posts.

**No double posts:** a publish job is never retried by itself. If it fails, the post shows the error. Check on Instagram or Facebook whether it went out before you press **Post now** again.

The Meta app setup is in SnappySjaak › Campagnes › Settings › Instagram & Facebook.
