from __future__ import annotations

import shutil
import tempfile
import uuid
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


def run_package(source, target, template_en, template_sl, instructions) -> tuple[Path, str]:
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
4. **Create the package.** The app reads the files and prepares review materials.
5. **Work from the ZIP.** Use the change map, flags, and query log while editing in Word.

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

    ready = source is not None and target is not None
    if not ready:
        st.info("Upload the English source and Slovenian target to enable package creation.")

    if st.button("Create QRD review package", type="primary", disabled=not ready):
        with st.spinner("Reading DOCX files and preparing the review package..."):
            try:
                zip_path, brief = run_package(source, target, template_en, template_sl, instructions)
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
