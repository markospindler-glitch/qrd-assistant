from __future__ import annotations

import csv
import json
import os
import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path

import streamlit as st

from qrd_assistant.app import UploadedFiles, build_reports


st.set_page_config(
    page_title="QRD Assistant",
    layout="wide",
)


def save_upload(upload, folder: Path) -> Path | None:
    if upload is None:
        return None
    path = folder / upload.name
    path.write_bytes(upload.getbuffer())
    return path


def get_openai_key() -> str | None:
    try:
        secret_key = st.secrets.get("OPENAI_API_KEY")
    except Exception:
        secret_key = None
    return secret_key or os.environ.get("OPENAI_API_KEY")


def extract_json_array(text: str) -> list[dict[str, str]]:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.strip("`")
        cleaned = cleaned.removeprefix("json").strip()
    start = cleaned.find("[")
    end = cleaned.rfind("]")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("The model did not return a JSON array.")
    data = json.loads(cleaned[start : end + 1])
    if not isinstance(data, list):
        raise ValueError("The model response was not a JSON array.")
    return data


def write_ai_drafts(project_dir: Path, api_key: str, model: str, max_rows: int) -> None:
    from openai import OpenAI

    out_dir = project_dir / "output"
    change_map = out_dir / "source_change_map.csv"
    rows: list[dict[str, str]] = []
    with change_map.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            revised = (row.get("revised_text") or "").strip()
            original = (row.get("original_text") or "").strip()
            if revised and revised != original:
                rows.append(row)
            if len(rows) >= max_rows:
                break

    drafts: list[dict[str, str]] = []
    client = OpenAI(api_key=api_key)
    batch_size = 8
    instructions_text = """
You are an expert EMA QRD English-to-Slovenian regulatory translator.
Translate revised English source changes into Slovenian draft wording.
Preserve EMA QRD terminology, product names, section references, units, punctuation conventions, and medical meaning.
Do not invent missing context. If the source row is a table fragment or ambiguous, provide the best draft and mark the safety field as "needs_review".
Return only a JSON array. Each item must contain:
source_row, section, slovenian_draft, safety, notes.
"""
    for start in range(0, len(rows), batch_size):
        batch = rows[start : start + batch_size]
        prompt = "Translate these changed English QRD rows into Slovenian draft wording:\n\n"
        prompt += json.dumps(
            [
                {
                    "source_row": row["idx"],
                    "section": row["section"],
                    "category_guess": row["category_guess"],
                    "english_original": row["original_text"],
                    "english_revised": row["revised_text"],
                }
                for row in batch
            ],
            ensure_ascii=False,
            indent=2,
        )
        response = client.responses.create(
            model=model,
            instructions=instructions_text,
            input=prompt,
        )
        drafts.extend(extract_json_array(response.output_text))

    csv_path = out_dir / "ai_slovenian_draft_translations.csv"
    fieldnames = ["source_row", "section", "slovenian_draft", "safety", "notes"]
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in drafts:
            writer.writerow({name: row.get(name, "") for name in fieldnames})

    md_lines = [
        "# AI Slovenian Draft Translations",
        "",
        "These are draft translations of revised English source rows. They are not a client-final tracked-change document.",
        "",
    ]
    for row in drafts:
        md_lines.extend(
            [
                f"## Source row {row.get('source_row', '')}",
                "",
                f"Section: {row.get('section', '')}",
                "",
                f"Safety: `{row.get('safety', 'needs_review')}`",
                "",
                row.get("slovenian_draft", ""),
                "",
                f"Notes: {row.get('notes', '')}",
                "",
            ]
        )
    (out_dir / "ai_slovenian_draft_translations.md").write_text("\n".join(md_lines), encoding="utf-8")

    zip_path = project_dir / "qrd_assistant_output.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in out_dir.iterdir():
            zf.write(item, item.name)


def run_package(
    source,
    target,
    template_en,
    template_sl,
    instructions,
    make_ai_drafts: bool,
    api_key: str | None,
    model: str,
    max_ai_rows: int,
) -> tuple[Path, str]:
    project_id = uuid.uuid4().hex[:10]
    project_dir = Path(tempfile.gettempdir()) / "qrd_assistant_streamlit" / project_id
    upload_dir = project_dir / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)

    source_path = save_upload(source, upload_dir)
    target_path = save_upload(target, upload_dir)
    if source_path is None or target_path is None:
        raise ValueError("Upload both the English source and Slovenian target documents.")

    files = UploadedFiles(
        source_docx=source_path,
        target_docx=target_path,
        template_en=save_upload(template_en, upload_dir),
        template_sl=save_upload(template_sl, upload_dir),
        instructions=save_upload(instructions, upload_dir),
    )
    outputs = build_reports(files, project_dir)
    if make_ai_drafts:
        if not api_key:
            raise ValueError("OPENAI_API_KEY is not configured in Streamlit secrets.")
        write_ai_drafts(project_dir, api_key, model, max_ai_rows)
    brief = outputs["brief"].read_text(encoding="utf-8")
    return outputs["zip"], brief


st.title("QRD Assistant")
st.caption("EMA QRD English-to-Slovenian review package generator")

left, right = st.columns([2, 1], gap="large")

with right:
    st.subheader("Use This App")
    st.markdown(
        """
1. **Upload the English source.** Use the Word file with tracked changes.
2. **Upload the Slovenian target.** Use the file you need to update.
3. **Add QRD templates if available.** The app still works without them.
4. **Optional:** turn on AI draft translations if an OpenAI API key is configured.
5. **Create the package.** The app reads the files and prepares review materials.
6. **Work from the ZIP.** Use the change map, draft translations, flags, and query log while editing in Word.

**Important:** the app does not overwrite your client document. Final tracked changes should still be made and checked in Microsoft Word.
"""
    )

with left:
    st.subheader("Documents")
    source = st.file_uploader(
        "English source with tracked changes",
        type=["docx"],
        help="Required. This should be the English QRD document with visible tracked changes.",
    )
    target = st.file_uploader(
        "Slovenian target document",
        type=["docx"],
        help="Required. This is the Slovenian file you need to update for the client.",
    )

    st.divider()
    st.subheader("Optional files")
    template_en = st.file_uploader("English QRD template", type=["docx"])
    template_sl = st.file_uploader("Slovenian QRD template", type=["docx"])
    instructions = st.file_uploader("Instructions document", type=["docx"])

    st.divider()
    st.subheader("AI draft translation")
    api_key = get_openai_key()
    make_ai_drafts = st.checkbox(
        "Generate automatic Slovenian draft translations",
        help="Creates draft Slovenian wording for revised English source rows. Requires OPENAI_API_KEY in Streamlit secrets.",
    )
    model = st.text_input("OpenAI model", value="gpt-5.2", disabled=not make_ai_drafts)
    max_ai_rows = st.number_input(
        "Maximum changed rows to translate",
        min_value=1,
        max_value=400,
        value=40,
        step=10,
        disabled=not make_ai_drafts,
        help="Use a smaller number for quick tests, then increase for the full document.",
    )
    if make_ai_drafts and not api_key:
        st.warning("Add OPENAI_API_KEY in Streamlit secrets before deploying/running AI draft translations.")

    ready = source is not None and target is not None
    if not ready:
        st.info("Upload the English source and Slovenian target to enable package creation.")

    if st.button("Create QRD review package", type="primary", disabled=not ready):
        with st.spinner("Reading DOCX files and preparing the review package..."):
            try:
                zip_path, brief = run_package(
                    source,
                    target,
                    template_en,
                    template_sl,
                    instructions,
                    make_ai_drafts,
                    api_key,
                    model.strip() or "gpt-5.2",
                    int(max_ai_rows),
                )
                st.session_state["zip_bytes"] = zip_path.read_bytes()
                st.session_state["brief"] = brief
                st.session_state["zip_name"] = "qrd_assistant_output.zip"
            except Exception as exc:
                st.error(f"Processing stopped: {exc}")

if "zip_bytes" in st.session_state:
    st.success("Review package ready.")
    st.download_button(
        "Download ZIP package",
        data=st.session_state["zip_bytes"],
        file_name=st.session_state.get("zip_name", "qrd_assistant_output.zip"),
        mime="application/zip",
    )
    with st.expander("Review brief", expanded=True):
        st.markdown(st.session_state.get("brief", ""))


def cleanup_old_tmp() -> None:
    root = Path(tempfile.gettempdir()) / "qrd_assistant_streamlit"
    if not root.exists():
        return
    # Streamlit reruns often; keep cleanup intentionally conservative.
    for child in root.iterdir():
        if child.is_dir() and len(list(root.iterdir())) > 30:
            shutil.rmtree(child, ignore_errors=True)


cleanup_old_tmp()
