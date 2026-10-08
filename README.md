# Interview Assistant

An AI-powered mock interview assistant with voice input, dual-window answers, and Excel-backed Q&A storage.

## Features

- 🎙️ **Voice listening** — speak your interview question directly (Chrome/Edge)
- ⌨️ **Manual input** — type your question as a fallback
- 🤖 **AI-generated answers** — powered by OpenAI GPT-4o-mini
- 📋 **Dual-window output** — Key Points (left) + Illustrated Answer (right)
- 📊 **Excel storage** — all Q&A saved per role to `data/interview_qa_database.xlsx`
- 🔁 **Smart caching** — Excel is always checked first (fast & free); AI is only called when no similar question exists
- 6 roles: Test Manager, Scrum Master, Defect Manager, SAP Tester, SAP Test Manager, Project Manager

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Backend | **Python 3.11+**, FastAPI, uvicorn |
| AI | OpenAI `gpt-4o-mini` (async) |
| Excel | `openpyxl` |
| Fuzzy matching | `rapidfuzz` |
| Frontend | HTML + CSS + Vanilla JS (Web Speech API) |
| Java required? | ❌ No |
| Node.js required? | ❌ No |

---

## Setup

### 1. Prerequisites

- Python 3.11 or newer — https://www.python.org/downloads/
- Google Chrome or Microsoft Edge (for voice input)

### 2. Install Python dependencies

```bash
cd interview-assistant/backend
pip install -r requirements.txt
```

Or using a virtual environment (recommended):

```bash
cd interview-assistant/backend
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
```

### 3. Configure your OpenAI API key

```bash
copy .env.example .env
```

Edit `.env`:

```
OPENAI_API_KEY=sk-your-key-here
PORT=3001
HOST=127.0.0.1
```

### 4. Start the backend

```bash
python server.py
```

Server starts at `http://127.0.0.1:3001`

### 5. Open the frontend

Open `interview-assistant/frontend/index.html` in **Chrome or Edge**.

Or access it via the server directly at: `http://127.0.0.1:3001/app`

---

## Answer Lookup Flow

```
Question asked
     │
     ▼
Search Excel (rapidfuzz token_set_ratio ≥ 65%)
     │
     ├─── Match found ──► Return from cache  ✅  (instant, no API cost)
     │
     └─── No match    ──► Call OpenAI AI  🤖  → Save result to Excel
```

Excel is **always checked first** regardless of how many sessions have been completed.

---

## Data Storage

| File | Purpose |
|------|---------|
| `data/interview_qa_database.xlsx` | All Q&A — one sheet per role |
| `data/stats.json` | Interview session counts per role |

---

## Project Structure

```
interview-assistant/
├── backend/
│   ├── server.py              # FastAPI application
│   ├── config.py              # Settings (loaded from .env)
│   ├── requirements.txt
│   ├── .env.example
│   └── services/
│       ├── ai_service.py      # OpenAI async integration
│       ├── excel_service.py   # openpyxl read/write
│       └── cache_service.py   # rapidfuzz matching + session stats
├── frontend/
│   ├── index.html             # Main UI
│   ├── style.css              # Styles
│   └── app.js                 # Voice + API + render logic
└── data/                      # Auto-created on first run
    ├── interview_qa_database.xlsx
    └── stats.json
```
