# Changelog

All notable changes to the Corporate Jargon Translator Outlook Add-in will be documented in this file.

## [1.1.0] - 2026-06-16

### Added
- **Insert Mode selector** — choose between "Both" (original email + corporate translation stacked) or "Corporate Only" (replace entirely with translation)
- Insert mode preference persists in localStorage between sessions
- Both "Insert into Email" and "Copy to Clipboard" respect the selected insert mode
- Original email text is preserved in memory during translation for the "Both" mode
- Separator line (`--- Corporate Translation ---`) between original and translated text in "Both" mode

## [1.0.0] - 2026-06-16

### Added
- Initial release
- Outlook Add-in with ribbon button ("Translate to Corporate")
- Task pane UI with OpenAI API key input (stored locally, never transmitted to third parties)
- Four jargon intensity levels:
  - **Mild** — Light buzzword swap
  - **Medium** — Heavy corporate doublespeak (default)
  - **Extreme** — Full McKinsey word salad
  - **Unhinged** — Barely comprehensible executive nonsense
- Preview panel showing original text and translated output side by side
- Insert translated text directly into email compose window
- Copy to clipboard functionality
- Uses GPT-4o-mini for fast, cost-effective translations
- Compatible with Outlook Desktop (Windows/Mac), Outlook on the Web, and New Outlook
- Self-signed dev certificate support for local HTTPS development
