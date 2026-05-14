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
  - `terminology_and_template_flags.csv`
  - `review_brief.md`
  - `chatgpt_prompt.md`
  - `query_log.md`
  - a clean target review copy

## Accuracy Position

This app is intentionally analysis-first. It does not silently translate or overwrite the Slovenian target document. For QRD work, the final client deliverable should still be prepared in Microsoft Word with Track Changes enabled and validated in Word.

## Run

Use the bundled Codex Python runtime:

```powershell
& "C:\Users\marko\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" qrd_assistant\app.py
```

Then open:

`http://127.0.0.1:8765`

