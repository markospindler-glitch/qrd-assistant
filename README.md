# QRD Assistant

A small Streamlit app for EMA QRD English-to-Slovenian review projects.

## What It Does

- Upload an English DOCX source with tracked changes.
- Upload a Slovenian DOCX target.
- Optionally upload English/Slovenian QRD templates and an instructions DOCX.
- Generate a ZIP review package containing:
  - source tracked-change map
  - source and target paragraph exports
  - optional AI Slovenian draft translations for revised English rows
  - optional AI tracked-change draft DOCX based on the Slovenian target
  - terminology/template flags
  - review brief
  - ChatGPT review prompt
  - query log
  - clean target review copy

## Accuracy Position

Without AI mode, this app is analysis-first. With AI mode, it creates a Slovenian tracked-change draft DOCX, but that file is still a review draft. Final delivery should be checked in Microsoft Word against the English source, QRD templates, instructions, and terminology rules.

## Optional OpenAI Draft Translation

The app can generate draft Slovenian updates and apply them to a copy of the Slovenian target DOCX with Word tracked changes. This requires an OpenAI API key.

For Streamlit Community Cloud:

1. Open the app settings.
2. Go to **Secrets**.
3. Add:

```toml
OPENAI_API_KEY = "your_api_key_here"
```

The key should never be committed to GitHub. AI output is included in the ZIP as CSV, Markdown, and an `AI-TRACKED-DRAFT.docx` file when changes can be matched to target paragraphs. Review every tracked change in Microsoft Word before client delivery.

## Local Run

```powershell
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Streamlit Community Cloud

Deploy settings:

- Repository: this GitHub repository
- Branch: `main`
- Main file path: `streamlit_app.py`
- Python version: current default is fine
