# 🤖 Medusa AI Telegram Bot

A state-of-the-art Telegram Bot powered by **Medusa AI**, featuring dual intelligence engines, multimodal analysis, live web knowledge, and strict model secrecy.

---

## 🌟 Key Features

1. **Real-Time Web Knowledge:**
   - Automatically detects queries regarding live events, current news, dates, and prices.
   - Queries live web sources to provide fresh, up-to-date answers.

2. **Multimodal Analysis:**
   - **Photos & Images:** Send any image for object recognition, scene breakdown, and visual Q&A.
   - **PDF Documents:** Comprehensive summaries, research paper breakdowns, and document Q&A.
   - **ZIP Archives:** Deep inspection of zip files, complete directory tree analysis, and code review.
   - **Voice & Audio Notes:** Automatic transcription and conversational voice responses.
   - **Code & Text Files:** Script debugging, code explanations, and file reviews.

3. **Clean HTML Formatting:**
   - Robust Markdown-to-Telegram HTML rendering.
   - Proper bold `<b>...</b>`, italics, code blocks `<pre><code>...</code></pre>`, without raw asterisks (`***`).

4. **Dual Engine Modes:**
   - **⚡ Default Mode (`/default`):** Ultra-fast, long-context conversational chat with live web search.
   - **🐍 Medusa Mode (`/medusa`):** Secondary local dedicated neural engine.
   - If the Medusa engine is inactive or offline, the bot cleanly informs the user: *"Models are offline now"*.

5. **Strict Model Secrecy:**
   - The bot persona is exclusively **Medusa AI**.
   - Underlying model names, architectures, weights, and backend providers are strictly protected and never disclosed to users.

---

## 💬 Bot Commands

| Command | Description |
| :--- | :--- |
| `/start` | Open the welcome menu with interactive buttons |
| `/help` | Detailed guide on commands, modes, and file uploads |
| `/default` | Switch to Default Mode (Fast chat + live web search) |
| `/medusa` | Switch to Medusa Secondary Mode |
| `/reset` or `/new` | Clear current chat memory and start fresh |
| `/system <prompt>` | Set a custom persona or role |
| `/system reset` | Restore default Medusa persona |

---

## ⚙️ Configuration (`.env`)

```env
TELEGRAM_BOT_TOKEN=your_telegram_bot_token
GROQ_API_KEY=your_groq_api_key
GEMINI_API_KEY=your_gemini_api_key
```

---

## ▶️ Running the Bot

### 1. Telegram Bot (Public Access)
Double-click `run_bot.bat` or run:
```powershell
python bot.py
```

### 2. Interactive CLI Terminal (Text Only)
Double-click `run_cli.bat` or run:
```powershell
python cli.py
```
- Includes animated ASCII Figlet banner, live date & time, Tsukiziro author tag, colorized terminal dialog, and real-time live web search.
- Dual Modes:
  - `/default` — Cloud engine with real-time web search.
  - `/medusa` — Local GPU model (`medusa-1.8-flash.gguf`).
- Memory & Control:
  - `/clear` — Clear memory & refresh banner.
  - `/exit` — Quit CLI.
