# Smart Page Digest

Chrome extension that summarizes, quizzes, and answers questions about any web page using Chrome's built-in AI (Gemini Nano). All processing happens on your device — no data is sent to external servers.

## Features

### Core

- **One-click summarization**: Click the extension icon to open the side panel and summarize the current page
- **Dual summary**: Displays both TL;DR and Key Points sections simultaneously (configurable)
- **Language selection**: Summarize in your preferred language (Auto / Japanese / English / Spanish)
- **Custom Q&A chat**: Ask follow-up questions about the page in an interactive chat interface
- **Streaming output**: Results appear in real-time as the AI generates them
- **Translation**: Translate summaries into other languages with on-device translation models
- **Read aloud**: Text-to-speech with voice selection and adjustable speed (0.5x–2x)
- **Copy**: Copy summaries to clipboard; copy page title as a Markdown link
- **Keyboard shortcut**: Alt+S (Option+S on Mac) to open panel and summarize
- **Context menu**: Right-click on any page to summarize
- **Tab-aware caching**: Summaries, quizzes, and chat history are preserved per-tab; switching tabs restores previous results instantly
- **Auto-summarize**: Optionally auto-summarize when switching tabs while the panel is open

### Research-Based Features

- **Comprehension Quiz** (Retrieval Practice)
  - Generates 3 quiz questions from page content
  - Click questions to reveal answers — active recall strengthens memory
  - Works independently of summarization (extracts text on demand)
  - Based on: Roediger & Karpicke (2006) — retrieval practice improves long-term retention by 50% vs re-reading

- **Importance-Highlighted Key Points**
  - Each key point tagged with importance: HIGH / MEDIUM / LOW
  - Visual indicators: red (high), yellow (medium), gray (low) border and badge
  - Helps focus on critical information first
  - Based on: Nielsen (2006) F-shaped reading pattern research

## Privacy

- **100% on-device processing**: All summarization and AI queries are handled by Gemini Nano, running entirely on your device
- **No data transmission**: Page content is never sent to external servers
- **No training on your data**: Gemini Nano is a pre-trained, inference-only model
- **Offline capable**: Works without an internet connection after the initial model download

## Requirements

- Chrome 138 or later (Chrome Canary recommended)
- 22GB free disk space (for Gemini Nano model, downloaded on first use)
- 4GB+ GPU VRAM or 16GB+ RAM with 4+ CPU cores

### Chrome Flags (if needed)

If the Summarizer API is not available, enable these flags in `chrome://flags`:

1. `#optimization-guide-on-device-model` → Enabled
2. `#prompt-api-for-gemini-nano-multimodal-input` → Enabled (for custom prompts)

## Installation

1. Clone this repository
2. Open `chrome://extensions/` in Chrome
3. Enable "Developer mode" (top right)
4. Click "Load unpacked" and select the `src/` directory

## Architecture

```
Extension Icon → Side Panel opens
                    ↓
              Summarize Button
                    ↓
    Background SW → Content Script (extract text)
                    ↓
    Side Panel ← Summarizer API (TL;DR + Key Points)
               ← Prompt API (Quiz / Q&A / Importance)
               ← Translator API (language translation)
               ← Web Speech API (read aloud)
```

### APIs Used

| API | Purpose |
|-----|---------|
| Summarizer API | Standard TL;DR and key-points summaries |
| Prompt API (LanguageModel) | Cross-language summarization, custom Q&A, quiz generation, importance annotation |
| Translator API | On-device translation of summaries |
| Web Speech API | Text-to-speech read aloud |
| Side Panel API | Right-side panel UI |
| Context Menus API | Right-click page summarization |
| Commands API | Keyboard shortcut (Alt+S) |

### Files

| File | Purpose |
|------|---------|
| `src/manifest.json` | Extension configuration (Manifest V3) |
| `src/config.js` | Shared constants and default settings |
| `src/content.js` | Page text extraction from DOM |
| `src/background.js` | Service worker, message routing |
| `src/sidepanel/` | Side panel UI (HTML, CSS, JS) |
| `src/options/` | Settings page |
| `src/_locales/` | Internationalization (en, ja, es) |

## License

MIT
