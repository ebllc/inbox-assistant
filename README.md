# Inbox Assistant

A free, personal Gmail helper that runs entirely inside your own Google account (Google Apps Script + the Gemini API free tier).

- Plain-English rules in a Google Sheet, plus a **Rule Library** of ready-made rules by account type
- Auto-labels, archives (now or after N days), flags what needs you, alerts you immediately
- Morning (and optional midday) digest email listing everything it did, with links
- Starts in test mode (`DRY_RUN`) so you can check its choices before it touches anything
- Star any email to keep the assistant's hands off it

## Install

Use the setup guide: https://claude.ai/artifact/MqpvLwnKzfprXRwC25EjDm

The quick version: make a copy of the template sheet, choose **📬 Inbox Assistant → Set up** from the sheet's menu, click Allow, and paste a free Gemini key from https://aistudio.google.com/apikey into the Settings tab.

## Updates

Each copy can update itself from this repo.

- **Once per Google account:** turn on the Apps Script API at https://script.google.com/home/usersettings
- **Get the latest now:** menu → *Install latest version*, or set `UPDATE_CODE` to `TRUE` in Settings
- **Automatic:** set `AUTO_UPDATE` to `TRUE`; it checks once a day
- **Channel:** `UPDATE_CHANNEL` = `main` (stable, default) or `beta` (newest first)
- **Undo:** menu → *Undo last update* (keeps the previous version as a backup)

Each copy keeps its own `appsscript.json` (time zone, permissions). Updates replace only the code.

## Files

| File | What it is |
|---|---|
| `Code.gs` | The whole assistant. This is what copies download when they update. |
| `appsscript.json` | Project settings (permissions). Needed for self-update. |
| `Stub.gs` | For running extra accounts off one main project as an Apps Script library. Optional. |

## Privacy

Your Gemini key, rules, settings, and log stay in your own Google Sheet. Email content goes only to Gmail and Google's Gemini API. On Gemini's free tier, Google may use what's sent to improve its products; turn on billing for the key if that matters for the account (about $0–2/month).
