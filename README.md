# Corporate Jargon Translator — Outlook Add-in

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Turns your boring plain English emails into elite executive doublespeak using AI.

## What it does

Adds a **"Translate to Corporate"** button to your Outlook ribbon. Click it, pick a jargon intensity level, and watch your email get transformed into incomprehensible corporate prose.

### AI Providers (cheapest models)

| Provider | Model | Cost |
|----------|-------|------|
| **OpenAI** | gpt-4o-mini | ~$0.15 / 1M input tokens |
| **Google Gemini** | gemini-2.0-flash | Free tier available, or ~$0.075 / 1M input tokens |

Each user provides their own API key — enter it once and it's saved locally in your browser.

### Jargon Levels

| Level | Example Output |
|-------|---------------|
| **Mild** | "Let's discuss the blockers on this project" |
| **Medium** | "Leverage cross-functional alignment to action the deliverables" |
| **Extreme** | "Drive holistic paradigm shift through value-added ecosystem synergies" |
| **Unhinged** | "Synergize the bandwidth-limited cross-pollination of our action-oriented deliverable pipeline" |

## Prerequisites

- **Node.js** (v16+) — [Download](https://nodejs.org/)
- **Outlook** — Desktop (Windows/Mac), Outlook on the Web, or New Outlook
- An API key from **one or both** providers:
  - [OpenAI API Key](https://platform.openai.com/api-keys)
  - [Gemini API Key](https://aistudio.google.com/apikey) (free tier available)

## Setup

### 1. Install dependencies

```bash
cd outlook-corporate-translator
npm install
```

### 2. Start the dev server

```bash
npm start
```

This will:
- Start a local HTTPS server (on `https://localhost:3000`)
- Auto-install dev certificates if needed
- Sideload the add-in into Outlook

### 3. Use it

1. Open an email in Outlook (read or compose mode)
2. Look for the **"Corporate Translator"** group in the ribbon/toolbar
3. Click **"Translate to Corporate"**
4. Select your AI provider (OpenAI or Gemini)
5. Enter your API key (stored locally in your browser — never sent to our servers)
6. Pick your jargon intensity
7. Click **"Translate to Corporate"**
8. Review the translation, then choose your insert mode:
   - **Both** — Keeps your original text and appends the corporate translation below it
   - **Corporate Only** — Replaces the email with just the translation
9. Click **Insert into Email** or **Copy to Clipboard**

## Manual Sideloading (if `npm start` doesn't auto-sideload)

### Outlook Desktop (Windows)
1. File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs
2. Add `https://localhost:3000` and check "Show in Menu"
3. Restart Outlook
4. Home → Get Add-ins → My Add-ins → find "Corporate Jargon Translator"

### Outlook on the Web
1. Open an email → Click **...** → Get Add-ins
2. Click "My add-ins" → "Custom Add-ins" → **Add from a URL**
3. Paste: `https://jimdawdy-hub.github.io/outlook-corporate-translator/manifest.xml`

> **Note:** If your organization locks down add-in installation, you may need admin approval. The URL method works on personal accounts and unlocked org accounts.

## Project Structure

```
outlook-corporate-translator/
├── manifest.xml          # Add-in manifest (tells Outlook what this plugin does)
├── package.json
├── server.js             # Local HTTPS dev server
├── src/
│   ├── taskpane.html     # Main UI panel
│   ├── taskpane.css      # Styles
│   ├── taskpane.js       # Core logic (Office.js + AI APIs)
│   ├── commands.html     # Command handler stub
│   └── commands.js       # Reserved for direct-action commands
└── README.md
```

## How It Works

1. **Office.js** — Microsoft's library for building add-ins. Gives us access to the email content.
2. **Task Pane** — A side panel that opens in Outlook. This is where the UI lives.
3. **AI APIs** — Sends the email text to your chosen provider (OpenAI or Gemini) with a system prompt tuned to the selected jargon level.
4. **Insert Modes** — "Both" keeps original + translation stacked. "Corporate Only" replaces entirely.

## Security Notes

- API keys are stored in `localStorage` (your browser's local storage). They are only sent to the provider you selected (OpenAI or Google) — never to our servers.
- No analytics, no tracking, no data collection.
- The add-in only has `ReadWriteItem` permission — it can read/modify the current email but nothing else.

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "Translate" button doesn't appear | Make sure the add-in is sideloaded. Check Outlook's add-in settings. |
| "Could not read email body" | Make sure you have an email open. Try clicking on the email body first. |
| API key error (OpenAI) | Double-check your key starts with `sk-` and has not expired. |
| API key error (Gemini) | Double-check your key starts with `AIza` and is enabled for the Gemini API. |
| CORS error in console | Shouldn't happen (both providers allow browser calls), but check your API key permissions. |

## License

MIT — see [LICENSE](LICENSE).
