# QRD Assistant

A small Streamlit app for EMA QRD English-to-Slovenian review projects.

## What It Does

- Upload an English DOCX source with tracked changes.
- Upload a Slovenian DOCX target.
- Optionally upload English/Slovenian QRD templates and an instructions DOCX.
- Generate a ZIP review package containing:
  - source tracked-change map
  - source and target paragraph exports
  - terminology/template flags
  - review brief
  - ChatGPT review prompt
  - query log
  - clean target review copy

## Accuracy Position

This app is analysis-first. It does not silently translate or overwrite the client document. Final tracked changes should still be applied and validated in Microsoft Word.

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

