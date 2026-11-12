# Media poller (BEAST PC)

Makes the Campagnes photos and videos on the BEAST PC (RTX 5090). SnappySjaak on Render only queues the work. This app claims `media_*` jobs and does all the heavy work here:

- reads the photos from Google Drive;
- draws the frames with Pillow and encodes them with ffmpeg (NVENC on the RTX);
- writes French and English captions with the local Ollama model;
- uploads the results to Drive.

It sends back only ids and texts. The other pollers never take `media_*` jobs.

## Setup

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

The AI recipes (Flower Drop slow motion, living photo, photo → art) come with the ComfyUI setup in phase 4.
