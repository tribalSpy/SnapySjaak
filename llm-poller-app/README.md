# Shadow LLM Poller

Standalone poller app for a Windows PC that has a local Ollama model running on port `11434`.

This app does not need a browser session. It talks directly to Shadow App on Render with an API key, claims queued LLM jobs, runs them on the local Ollama instance, and posts the results back.

For UKDocs CSI temp-phyto PDF vision checks, install `PyMuPDF` on the poller PC too:

```powershell
pip install pymupdf
```

## What this first version does

- Sends heartbeat updates to Shadow
- Polls for queued jobs
- Runs `ollama_chat` jobs against local Ollama
- Can render temp-phyto PDF pages into images for Ollama vision models when `PyMuPDF` is installed
- Sends success or failure back to Shadow

## Files

- `poller.py`
- `config.example.json`
- `run_poller.bat`

## Setup on the new PC

1. Install Python 3.11 or newer.
2. Make sure Ollama is running locally on that PC.
3. Copy this whole `llm-poller-app` folder onto the PC.
4. Copy `config.example.json` to `config.json`.
5. Fill in the values in `config.json`.
6. Start the poller with `run_poller.bat`.

## Example config

```json
{
  "server_url": "https://your-render-app.onrender.com",
  "api_key": "set-the-same-value-as-SHADOW_LLM_POLLER_API_KEY",
  "agent_name": "office-ollama-1",
  "pc_name": "Office-PC",
  "model_name": "gwen32b-vision",
  "version": "1.0.0",
  "poll_interval_seconds": 10,
  "ollama_url": "http://127.0.0.1:11434",
  "capabilities": [
    "ollama_chat"
  ]
}
```

## Render env var

Set this on the Shadow Render service:

- `SHADOW_LLM_POLLER_API_KEY`

Use the exact same value in the poller `config.json`.

## Test the connection

After deploy, you can create a test job from the backend with:

- `POST /api/llm/jobs`

Example payload:

```json
{
  "job_type": "ollama_chat",
  "priority": 10,
  "payload_json": {
    "model": "gwen32b-vision",
    "messages": [
      {
        "role": "user",
        "content": "Reply with only: poller ok"
      }
    ]
  }
}
```

## Supported job payload

`ollama_chat`

```json
{
  "model": "gwen32b-vision",
  "messages": [
    { "role": "user", "content": "Hello" }
  ],
  "format": "json",
  "options": {},
  "keep_alive": "5m"
}
```

## Notes

- The poller keeps running without browser login.
- If Ollama is down, the job is marked failed and the reason is sent back to Shadow.
- This version is the infrastructure layer. UKDocs CSI job types can be added on top of this without changing the install method.
