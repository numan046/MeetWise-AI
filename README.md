<div align="center">

# 📋 MeetWise AI

### AI Meeting Notes Generator

Turn long, messy meeting notes into a clean structured summary with **key points, decisions, action items and deadlines**, ready to download as Markdown, PDF, JSON or CSV.

![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-App-FF4B4B?logo=streamlit&logoColor=white)
![Groq](https://img.shields.io/badge/LLM-Groq%20%7C%20Llama%203.3%2070B-F55036)
![pandas](https://img.shields.io/badge/pandas-Data-150458?logo=pandas&logoColor=white)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen)

</div>

---

## 📖 Overview

**MeetWise AI** is a web application built with Python and Streamlit. The user pastes meeting notes or uploads a file, and the app sends the text to an AI model. The model returns a structured summary. The app can also rewrite the raw notes in a cleaner, more readable form, and every result can be downloaded in several formats.

> **Internship task:** Create an AI application that converts lengthy meeting notes into structured summaries, identifying key discussion points, action items, important decisions and deadlines while improving readability.
>
> **Level:** Intermediate Level – Project 2 · **Domain:** Artificial Intelligence · **Organisation:** CyberSNova

---

## ✨ Features

| Feature | Description |
|---|---|
| 📝 **Flexible input** | Paste notes or upload `.txt`, `.md`, `.pdf` or `.docx` files |
| 🧩 **Structured summary** | Executive summary, key discussion points, decisions, action items, deadlines, open questions and risks/blockers |
| ✅ **Action items** | Each item includes owner, due date and priority |
| 🎭 **Meeting templates** | 8 templates that tune the analysis to the meeting type |
| ✨ **Readability boost** | Optional rewrite of the raw notes into a cleaner version |
| 🌐 **Bilingual output** | Output language: English or Roman Urdu |
| 📚 **Meeting history** | Keeps up to 20 analysed meetings in the sidebar |
| 🔁 **Model fallback** | Automatically tries other Groq models if one is unavailable |
| 📤 **Multi-format export** | Markdown, PDF, JSON and an action-items CSV |
| 🔌 **Connection test** | Check that your API key works from the sidebar |

### 🎭 Meeting Templates

| Template | Best for |
|---|---|
| 🔍 **Auto-detect** | Let the AI work out the meeting type |
| ⚡ **Daily Standup** | Yesterday, today and blockers |
| 💡 **Brainstorm** | Idea sessions |
| 🤝 **Client Call** | Customer and stakeholder meetings |
| 👥 **1-on-1** | Manager and team-member check-ins |
| 📊 **Review / Retro** | Reviews and retrospectives |
| 🏛 **Board / Exec** | Financials, risks, governance and approvals |
| 🎓 **Training / Workshop** | Learning sessions |

---

## 🛠️ Tech Stack

| Tool | Used for |
|---|---|
| **Python 3** | Main programming language |
| **Streamlit** | Web interface: tabs, text box, file upload, sidebar settings |
| **Groq API (Llama 3.3 70B)** | AI model that analyses the notes |
| **OpenAI Python library** | Client used to call Groq's OpenAI-compatible API |
| **PyMuPDF** | Reading text from PDF files |
| **python-docx** | Reading text from Word (`.docx`) files |
| **pandas** | Building the action-items CSV |
| **fpdf2** | Generating the PDF report |

---

## 📁 Project Structure

```text
MeetWise-AI/
├── app.py               # Main Streamlit application
├── requirements.txt     # Python dependencies
├── README.md            # Project documentation
└── .streamlit/
    └── secrets.toml     # (optional, not committed) API key
```

---

## 🚀 Getting Started

### Prerequisites
- Python 3.9 or higher
- A free [Groq API key](https://console.groq.com/keys)

### 1. Clone the repository
```bash
git clone https://github.com/<your-username>/MeetWise-AI.git
cd MeetWise-AI
```

### 2. Create a virtual environment *(recommended)*
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

`requirements.txt`:
```text
streamlit
openai
pandas
pymupdf
python-docx
fpdf2
```

> PDF and DOCX reading and PDF export are optional features. If a library is missing, the app still runs and disables that feature.

### 4. Set your Groq API key

**Option A – Environment variable**
```bash
# Windows (PowerShell)
$env:GROQ_API_KEY="your_api_key_here"

# macOS / Linux
export GROQ_API_KEY="your_api_key_here"
```

**Option B – Streamlit secrets** (`.streamlit/secrets.toml`)
```toml
GROQ_API_KEY = "your_api_key_here"
```

### 5. Run the app
```bash
streamlit run app.py
```
The app opens at **http://localhost:8501**.

---

## 🖥️ Usage

1. Open the **📝 Input** tab and choose **Paste notes** or **Upload file**.
2. Select a **meeting template** (or leave it on Auto-detect) and the output language.
3. Optionally tick **Enhance readability**.
4. Run the analysis and open **📋 Structured Summary** to read the result.
5. Go to **📤 Export** to download the report.

### 📤 Export Formats

| Format | File | Content |
|---|---|---|
| Markdown | `meeting_<date>.md` | Full formatted report |
| PDF | `meeting_<date>.pdf` | Styled printable report |
| JSON | `meeting_<date>.json` | Structured data for other tools |
| CSV | `action_items.csv` | Action items with owner, due date and priority |

**Limits:** up to 40,000 characters of notes per analysis · 20 meetings kept in history.

---


---

## ⚠️ Limitations

- Very long notes are limited to 40,000 characters.
- The PDF export uses a basic Latin character set, so some special characters may be simplified.
- AI summaries should be reviewed before being shared.
- Free Groq usage is subject to rate limits.

---

## 🔮 Future Improvements

- Audio and video transcription as input
- Calendar integration for deadlines
- Email sharing of summaries

---

## 👤 Author

**Muhammad Numan**
AI Engineering Intern at **[CyberSNova](https://cybersnova.co.in)**


---

<div align="center">

⭐ If you found this project useful, consider giving it a star.

</div>
