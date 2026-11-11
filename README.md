# Inbox Assistant

A free, personal Gmail helper that runs entirely inside your own Google account (Google Apps Script + the Gemini API free tier).

- Plain-English rules in a Google Sheet, plus a **Rule Library** of ready-made rules by account type
- Auto-labels, archives (now or after N days), flags what needs you, alerts you immediately
- Morning (and optional midday) digest email listing everything it did, with links
- A private **phone app**: Today, Rules, Settings, and all your linked Gmail accounts in one place
- Starts in test mode (`DRY_RUN`) so you can check its choices before it touches anything
- Star any email to keep the assistant's hands off it

## Install

Use the setup guide: https://claude.ai/artifact/MqpvLwnKzfprXRwC25EjDm

Short version: make a copy of the template sheet → **📬 Inbox Assistant → Set up** → Allow → paste a free Gemini key from https://aistudio.google.com/apikey into the Settings tab. Then publish the phone app once (Deploy → New deployment → Web app → Only myself).

## How updates work

Each project holds only `Loader.gs`. The loader downloads `Core.gs` from this repo, keeps a copy in the project's script properties, and runs it.

- **Install the latest:** app → Settings → Check for update, the sheet menu, or `UPDATE_CODE` = `TRUE`
- **Automatic:** `AUTO_UPDATE` = `TRUE` (checks once a day)
- **Channel:** `UPDATE_CHANNEL` = `main` (stable, default) or `beta` (newest first)
- **Undo:** keeps the previous version; if a stored version won't start, the loader falls back by itself

No Apps Script API, Cloud project, or extra permissions are needed for updates.

## Files

| File | What it is |
|---|---|
| `Loader.gs` | The only code in each Apps Script project. Rarely changes. |
| `Core.gs` | The assistant itself, downloaded by the loader. Generated: `App.html` is embedded by `build.py`. |
| `App.html` | The phone app's page (edit this, then run `python3 build.py`). |
| `appsscript.json` | Project settings (permissions) for hand installs. |
| `Stub.gs` | Older library-based setup. Not needed with the loader. |

## Privacy

Your Gemini key, rules, settings, and log stay in your own Google Sheet. Email content goes only to Gmail and Google's Gemini API. On Gemini's free tier, Google may use what's sent to improve its products; turn on billing for the key if that matters for the account (about $0–2/month).
