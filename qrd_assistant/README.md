# QRD Assistant

A small local app for EMA QRD English-to-Slovenian review projects.

## What It Does

- Upload English source DOCX with tracked changes.
- Upload Slovenian target DOCX.
- Optionally upload English/Slovenian QRD templates and instructions.
- Generate a ZIP package with:
  - `source_change_map.csv`
  - `source_paragraphs.csv`
  - `target_paragraphs.csv`
  - optional AI tracked-change draft DOCX in the Streamlit app
  - `terminology_and_template_flags.csv`
  - `review_brief.md`
  - `chatgpt_prompt.md`
  - `query_log.md`
  - a clean target review copy

## Accuracy Position

The local no-AI app is intentionally analysis-first. The Streamlit app can also create an AI tracked-change draft DOCX, but that draft still needs Word review before client delivery.

## Run

Use the bundled Codex Python runtime:

```powershell
& "C:\Users\marko\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" qrd_assistant\app.py
```

Then open:

`http://127.0.0.1:8765`
