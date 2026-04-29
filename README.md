# Pustaka Agreement Generator

An AI-powered desktop application that auto-generates publishing agreements for authors. Built for the operations team at Pustaka Media — eliminating manual form-filling across 7 agreement types.

![Python](https://img.shields.io/badge/Python-3.10+-blue) ![Claude API](https://img.shields.io/badge/Claude-Anthropic-orange) ![License](https://img.shields.io/badge/License-MIT-green)

---

## Features

- **7 agreement types** — Katha (Pearl, Sapphire, Silver++, SGP), Pustaka (New Author, Legal Heir, Licensor), and Addendum
- **AI-powered extraction** — Paste or upload author details (PDF / DOCX / TXT), Claude API auto-fills all form fields
- **Word template population** — Fills DOCX templates with extracted data, ready to download
- **Desktop app** — No browser needed, runs as a standalone Windows EXE
- **Secure API key storage** — Key stored in per-user config, never hardcoded
- **Clean UI** — Built with CustomTkinter, designed for non-technical office staff

---

## Tech Stack

- **Frontend:** Python, CustomTkinter (Desktop GUI)
- **AI:** Anthropic Claude API (field extraction from unstructured text)
- **Document Processing:** python-docx, PyPDF2, lxml
- **Packaging:** PyInstaller (Windows EXE)

---

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/shlab7777/pustaka-agreement-generator.git
cd pustaka-agreement-generator
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Set up API key

Copy the example env file and add your Claude API key:

```bash
cp .env.example .env
```

Edit `.env`:

```
ANTHROPIC_API_KEY=your_claude_api_key_here
```

Get your API key from: https://console.anthropic.com

### 4. Add agreement templates

Place your DOCX agreement templates in a `templates/` folder:

```
templates/
├── Silver ++ template - katha agreement.docx
├── silver gold plan template - katha agreement.docx
├── Template_New_Authors_Paperback - pustaka agreement.docx
├── Template_Legal Heir_Digital - pustaka agreement.docx
├── Template_Licensor_Paperback new - pustaka agreement.docx
├── pearl plan template (addendum).docx
├── sapphire plan template (addendum).docx
└── pustaka- defualt addendum.docx
```

### 5. Run the app

```bash
python pustaka_app.py
```

---

## How It Works

1. Select the agreement type (Katha / Pustaka / Addendum)
2. Choose the publishing plan
3. Paste author details or upload a PDF/DOCX/TXT file
4. Claude API extracts and fills all form fields automatically
5. Review and edit fields if needed
6. Generate and download the completed Word agreement

---

## Project Structure

```
pustaka-agreement-generator/
├── pustaka_app.py       # Main application
├── .env.example         # API key template (copy to .env)
├── .gitignore           # Git ignore rules
├── requirements.txt     # Python dependencies
├── templates/           # DOCX agreement templates (not included in repo)
└── README.md            # This file
```

---

## Security Note

The Claude API key is loaded from a `.env` file and never hardcoded in source code. The `.env` file is excluded from version control via `.gitignore`. The app also supports storing the key in a per-user system config directory for packaged EXE distribution.

---

## Built By

**Sharan Shriney D** — Full Stack Developer & AI Engineer  
[LinkedIn](https://linkedin.com/in/sharan-shriney) · [GitHub](https://github.com/shlab7777)
