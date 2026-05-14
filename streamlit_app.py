from __future__ import annotations

import json
import os
import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path

import streamlit as st

from qrd_assistant.app import (
    UploadedFiles,
    build_reports,
    create_tracked_change_docx,
    extract_paragraphs,
    norm,
    paired_change_rows,
    write_csv,
)


st.set_page_config(
    page_title="QRD Assistant",
    layout="wide",
)


def apply_translat_style() -> None:
    st.markdown(
        """
        <style>
        :root {
            --qrd-red: #df1738;
            --qrd-red-dark: #b9152f;
            --qrd-red-soft: #fde9ed;
            --qrd-gray: #666666;
            --qrd-gray-dark: #3f3f3f;
            --qrd-gray-soft: #f5f5f5;
            --qrd-line: #dddddd;
            --qrd-text: #3a3a3a;
            --qrd-muted: #777777;
        }
        html,
        body,
        [data-testid="stAppViewContainer"] {
            background: #ffffff;
            color: var(--qrd-text);
        }
        .block-container {
            max-width: 1180px;
            padding-top: 0;
            padding-bottom: 2rem;
        }
        [data-testid="stSidebar"] {
            background: #f7f7f7;
            border-right: 1px solid #e2e2e2;
        }
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] li {
            color: #666666;
            font-size: 0.94rem;
        }
        h1,
        h2,
        h3,
        p {
            letter-spacing: 0;
        }
        h1,
        h2,
        h3 {
            color: var(--qrd-gray-dark);
            font-weight: 700;
        }
        p,
        .stCaptionContainer,
        [data-testid="stMarkdownContainer"] p {
            color: var(--qrd-muted);
        }
        .qrd-top-strip {
            margin: 0 calc(50% - 50vw);
            padding: 0.38rem max(1rem, calc((100vw - 1180px) / 2));
            background: var(--qrd-red);
            color: #ffffff;
            font-size: 0.82rem;
        }
        .qrd-top-strip-inner {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
        }
        .qrd-top-strip span {
            color: #ffffff;
            white-space: nowrap;
        }
        .qrd-brandbar {
            padding: 2.3rem 0 2rem;
            background: #ffffff;
            border-bottom: 1px solid #eeeeee;
        }
        .qrd-brandbar-main {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 2rem;
        }
        .qrd-logo {
            display: flex;
            align-items: center;
            gap: 0.25rem;
            line-height: 1;
        }
        .qrd-logo-mark {
            color: var(--qrd-red);
            font-size: 3.1rem;
            font-weight: 300;
            transform: translateY(-0.08rem);
        }
        .qrd-logo-word {
            color: #e01a3d;
            font-size: 2.1rem;
            font-weight: 800;
            letter-spacing: 0;
        }
        .qrd-logo-word span {
            color: #777777;
            font-weight: 700;
        }
        .qrd-tagline {
            margin: 0.4rem 0 0 2.3rem;
            color: #777777;
            font-family: Georgia, "Times New Roman", serif;
            font-size: 0.9rem;
        }
        .qrd-tagline strong {
            color: var(--qrd-red);
            font-weight: 700;
        }
        .qrd-nav {
            display: flex;
            flex-wrap: wrap;
            justify-content: flex-end;
            gap: 1.1rem;
            color: #666666;
            font-size: 0.88rem;
        }
        .qrd-nav strong {
            color: var(--qrd-red);
            font-weight: 500;
        }
        .qrd-service-line {
            margin-top: 1.7rem;
            color: #666666;
            font-size: 0.91rem;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-align: right;
            text-transform: uppercase;
        }
        .qrd-page-title {
            margin: 3rem 0 1.3rem;
            padding-top: 1.3rem;
            border-top: 1px solid var(--qrd-red);
            text-align: center;
        }
        .qrd-page-title h1 {
            margin: 0;
            color: var(--qrd-red);
            font-size: 2.15rem;
            line-height: 1.18;
        }
        .qrd-page-title p {
            margin: 0.55rem auto 0;
            max-width: 760px;
            color: #666666;
            font-size: 1rem;
        }
        .qrd-header {
            display: none;
        }
        .qrd-summary-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(120px, 1fr));
            gap: 12px;
            min-width: min(520px, 48vw);
        }
        .qrd-summary-tile {
            display: grid;
            gap: 8px;
            min-height: 86px;
            padding: 14px;
            background: var(--qrd-gray-panel);
            border: 1px solid var(--qrd-line);
            border-radius: 8px;
        }
        .qrd-summary-tile span {
            color: var(--qrd-muted);
            font-size: 13px;
            font-weight: 700;
        }
        .qrd-summary-tile strong {
            color: var(--qrd-text);
            font-size: 24px;
            line-height: 1.1;
        }
        .qrd-summary-tile.warning strong {
            color: var(--qrd-red);
        }
        .qrd-panel-title {
            margin: 1.1rem 0 0.35rem;
            padding: 0.85rem 1rem;
            background: #ffffff;
            border-top: 1px solid #d8d8d8;
            border-bottom: 1px solid #eeeeee;
        }
        .qrd-panel-title strong {
            display: block;
            color: var(--qrd-gray-dark);
            font-size: 16px;
        }
        .qrd-panel-title span {
            display: block;
            margin-top: 3px;
            color: var(--qrd-muted);
            font-size: 14px;
        }
        div[data-testid="stFileUploader"] {
            padding: 0.75rem;
            background: #ffffff;
            border: 1px solid #dcdcdc;
            border-radius: 3px;
        }
        div[data-testid="stExpander"],
        div[data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #dddddd;
            border-radius: 3px;
            overflow: hidden;
        }
        .stButton > button,
        .stDownloadButton > button {
            min-height: 40px;
            border-radius: 3px;
            border: 1px solid #cfcfcf;
            background: #ffffff;
            color: #555555;
            font-weight: 700;
            box-shadow: none;
        }
        .stButton > button:hover,
        .stDownloadButton > button:hover {
            border-color: var(--qrd-red);
            color: var(--qrd-red-dark);
            background: var(--qrd-red-soft);
        }
        .stButton > button[kind="primary"] {
            border-color: var(--qrd-red);
            background: var(--qrd-red);
            color: #ffffff;
        }
        .stButton > button[kind="primary"]:hover {
            border-color: var(--qrd-red-dark);
            background: var(--qrd-red-dark);
            color: #ffffff;
        }
        div[data-testid="stAlert"] {
            border-radius: 3px;
            border: 1px solid var(--qrd-line);
        }
        .qrd-side-note {
            padding: 0.9rem 1rem;
            background: #ffffff;
            border: 1px solid var(--qrd-line);
            border-top: 3px solid var(--qrd-red);
            border-radius: 3px;
        }
        @media (max-width: 900px) {
            .qrd-brandbar-main,
            .qrd-top-strip-inner {
                align-items: stretch;
                flex-direction: column;
            }
            .qrd-nav,
            .qrd-service-line {
                justify-content: flex-start;
                text-align: left;
            }
            .qrd-summary-grid {
                grid-template-columns: 1fr;
                min-width: 0;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_app_header() -> None:
    st.markdown(
        """
        <div class="qrd-top-strip">
          <div class="qrd-top-strip-inner">
            <span>Slovenščina | Angleščina</span>
            <span>+386 1 430 57 48 &nbsp;&nbsp; info@translat.si &nbsp;&nbsp; Komenskega 12, Slovenija</span>
          </div>
        </div>
        <header class="qrd-brandbar">
          <div class="qrd-brandbar-main">
            <div>
              <div class="qrd-logo" aria-label="QRD Assistant">
                <span class="qrd-logo-mark">{</span>
                <span class="qrd-logo-word">qrd<span>assistant</span></span>
              </div>
              <p class="qrd-tagline">Regulatory translation <strong>without surprises.</strong></p>
            </div>
            <nav class="qrd-nav" aria-label="Project navigation">
              <span>EMA QRD</span>
              <span>Tracked changes</span>
              <strong>Slovenian review</strong>
              <span>Template check</span>
            </nav>
          </div>
          <div class="qrd-service-line">Prevajanje | Lektoriranje | EMA besedila | QRD pregled | DTP</div>
        </header>
        <section class="qrd-page-title">
          <h1>Prepare your QRD output package</h1>
          <p>Upload the English changed source and Slovenian target, then create a ZIP with review files and an optional tracked-change draft DOCX.</p>
        </section>
        """,
        unsafe_allow_html=True,
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


def collect_reference_text(paths: list[Path | None], max_chars: int = 4000) -> str:
    parts: list[str] = []
    for path in paths:
        if not path:
            continue
        try:
            text = "\n".join(row["accepted"] for row in extract_paragraphs(path) if norm(row["accepted"]))
        except Exception:
            continue
        if text:
            parts.append(f"{path.name}:\n{text}")
    return "\n\n".join(parts)[:max_chars]


def write_ai_drafts(
    project_dir: Path,
    target_path: Path,
    reference_paths: list[Path | None],
    api_key: str,
    model: str,
    max_rows: int,
) -> dict[str, int | str]:
    from openai import OpenAI

    out_dir = project_dir / "output"
    rows: list[dict[str, str]] = []
    for row in paired_change_rows(project_dir):
        revised = norm(row.get("revised_text") or "")
        original = norm(row.get("original_text") or "")
        if revised != original and row.get("target_idx"):
            rows.append(row)
        if len(rows) >= max_rows:
            break

    drafts: list[dict[str, str]] = []
    client = OpenAI(api_key=api_key)
    batch_size = 8
    reference_text = collect_reference_text(reference_paths)
    instructions_text = """
You are an expert EMA QRD English-to-Slovenian regulatory translator.
Apply English tracked changes to the matching Slovenian target paragraph.
Preserve EMA QRD terminology, product names, section references, units, punctuation conventions, Slovenian grammar, and medical meaning.
Use the current Slovenian target paragraph as the base text. Change only what is needed to reflect the English revision.
If the English row is a full deletion, return an empty slovenian_revised value and safety "needs_review".
If the source and target paragraph do not correspond, return safety "do_not_apply".
Return only a JSON array. Each item must contain:
source_row, target_row, section, slovenian_revised, safety, change_type, notes.
Safety must be exactly one of: apply, needs_review, do_not_apply.
"""
    for start in range(0, len(rows), batch_size):
        batch = rows[start : start + batch_size]
        prompt = "Apply these changed English QRD rows to the matched Slovenian target paragraphs.\n\n"
        if reference_text:
            prompt += "Reference templates/instructions excerpt:\n"
            prompt += reference_text
            prompt += "\n\n"
        prompt += json.dumps(
            [
                {
                    "source_row": row["idx"],
                    "source_ordinal": row.get("source_ordinal", ""),
                    "target_row": row.get("target_idx", ""),
                    "target_ordinal": row.get("target_ordinal", ""),
                    "section": row["section"],
                    "category_guess": row["category_guess"],
                    "english_original": row["original_text"],
                    "english_revised": row["revised_text"],
                    "english_inserted": row.get("inserted_text", ""),
                    "english_deleted": row.get("deleted_text", ""),
                    "current_slovenian_target": row.get("target_current_text", ""),
                    "target_match_method": row.get("target_match_method", ""),
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
        parsed = extract_json_array(response.output_text)
        parsed_by_source = {str(item.get("source_row", "")): item for item in parsed}
        for input_row in batch:
            item = parsed_by_source.get(str(input_row["idx"]), {})
            safety = str(item.get("safety") or "needs_review").strip().lower()
            if safety not in {"apply", "needs_review", "do_not_apply"}:
                safety = "needs_review"
            slovenian_revised = str(item.get("slovenian_revised") or "").strip()
            is_deletion = input_row.get("category_guess") == "DELETION"
            drafts.append(
                {
                    "source_row": input_row["idx"],
                    "source_ordinal": input_row.get("source_ordinal", ""),
                    "target_row": input_row.get("target_idx", ""),
                    "target_ordinal": input_row.get("target_ordinal", ""),
                    "section": input_row.get("section", ""),
                    "slovenian_current": input_row.get("target_current_text", ""),
                    "slovenian_revised": slovenian_revised,
                    "safety": safety,
                    "change_type": str(item.get("change_type") or ("deletion" if is_deletion else "revision")),
                    "delete_target": str(is_deletion and not slovenian_revised),
                    "applied_to_tracked_docx": "False",
                    "notes": str(item.get("notes") or ""),
                }
            )

    revision_candidates = [
        {
            "target_idx": row["target_row"],
            "slovenian_revised": row["slovenian_revised"],
            "delete_target": row["delete_target"],
        }
        for row in drafts
        if row["target_row"] and row["safety"] != "do_not_apply"
    ]
    tracked_docx_path = out_dir / f"{target_path.stem}-AI-TRACKED-DRAFT.docx"
    applied_targets: set[str] = set()
    if revision_candidates:
        applied_targets = create_tracked_change_docx(target_path, revision_candidates, tracked_docx_path)
        for row in drafts:
            if row["target_row"] in applied_targets:
                row["applied_to_tracked_docx"] = "True"

    csv_path = out_dir / "ai_slovenian_draft_translations.csv"
    fieldnames = [
        "source_row",
        "source_ordinal",
        "target_row",
        "target_ordinal",
        "section",
        "slovenian_current",
        "slovenian_revised",
        "safety",
        "change_type",
        "delete_target",
        "applied_to_tracked_docx",
        "notes",
    ]
    write_csv(csv_path, [{name: row.get(name, "") for name in fieldnames} for row in drafts])

    md_lines = [
        "# AI Slovenian Draft Translations",
        "",
        "These are draft target-language updates generated from the English tracked source. Review them before client delivery.",
        "",
        f"Tracked-change draft DOCX: `{tracked_docx_path.name if applied_targets else 'not created'}`",
        "",
    ]
    for row in drafts:
        md_lines.extend(
            [
                f"## Source row {row.get('source_row', '')}",
                "",
                f"Section: {row.get('section', '')}",
                "",
                f"Target row: {row.get('target_row', '')}",
                "",
                f"Safety: `{row.get('safety', 'needs_review')}`",
                "",
                f"Applied to tracked DOCX: `{row.get('applied_to_tracked_docx', 'False')}`",
                "",
                "Current Slovenian:",
                "",
                row.get("slovenian_current", ""),
                "",
                "Revised Slovenian:",
                "",
                row.get("slovenian_revised", ""),
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
    return {
        "draft_count": len(drafts),
        "applied_count": len(applied_targets),
        "tracked_docx": tracked_docx_path.name if applied_targets else "",
    }


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
    ai_summary: dict[str, int | str] = {}
    if make_ai_drafts:
        if not api_key:
            raise ValueError("OPENAI_API_KEY is not configured in Streamlit secrets.")
        ai_summary = write_ai_drafts(
            project_dir,
            files.target_docx,
            [files.template_en, files.template_sl, files.instructions],
            api_key,
            model,
            max_ai_rows,
        )
    brief = outputs["brief"].read_text(encoding="utf-8")
    if ai_summary:
        brief += "\n\n## AI Tracked Draft\n\n"
        brief += f"- Draft rows generated: {ai_summary.get('draft_count', 0)}\n"
        brief += f"- Target paragraphs updated with tracked changes: {ai_summary.get('applied_count', 0)}\n"
        tracked_docx = ai_summary.get("tracked_docx")
        if tracked_docx:
            brief += f"- Tracked draft file in ZIP: `{tracked_docx}`\n"
        brief += "- Open the DOCX in Microsoft Word and review every tracked change before client delivery.\n"
    return outputs["zip"], brief


apply_translat_style()
render_app_header()

with st.sidebar:
    st.header("Use This App")
    st.markdown(
        """
1. **Upload the English source.** Use the Word file with tracked changes.
2. **Upload the Slovenian target.** Use the file you need to update.
3. **Add QRD templates if available.** The app still works without them.
4. **Optional but recommended:** turn on AI tracked-draft creation if an OpenAI API key is configured.
5. **Create the package.** The ZIP will include review materials and, in AI mode, a Slovenian DOCX with tracked draft changes.
6. **Review in Word.** Open the tracked draft, compare it with the English source, and accept/edit changes only after checking.

**Important:** the AI DOCX is a review draft, not a file to send without checking.
"""
    )
    st.markdown(
        """
<div class="qrd-side-note">
  <strong>Required</strong><br>
  English source DOCX and Slovenian target DOCX.
</div>
""",
        unsafe_allow_html=True,
    )

st.markdown(
    """
<div class="qrd-panel-title">
  <strong>Documents</strong>
  <span>Start with the two Word documents from the QRD project.</span>
</div>
""",
    unsafe_allow_html=True,
)
doc_left, doc_right = st.columns(2, gap="large")
with doc_left:
    source = st.file_uploader(
        "English source with tracked changes",
        type=["docx"],
        help="Required. This should be the English QRD document with visible tracked changes.",
    )
with doc_right:
    target = st.file_uploader(
        "Slovenian target document",
        type=["docx"],
        help="Required. This is the Slovenian file you need to update for the client.",
    )

st.markdown(
    """
<div class="qrd-panel-title">
  <strong>Optional QRD material</strong>
  <span>Templates and instructions improve checks, but the app can run without them.</span>
</div>
""",
    unsafe_allow_html=True,
)
opt_a, opt_b, opt_c = st.columns(3, gap="large")
with opt_a:
    template_en = st.file_uploader("English QRD template", type=["docx"])
with opt_b:
    template_sl = st.file_uploader("Slovenian QRD template", type=["docx"])
with opt_c:
    instructions = st.file_uploader("Instructions document", type=["docx"])

st.markdown(
    """
<div class="qrd-panel-title">
  <strong>AI tracked draft</strong>
  <span>Use this when you want a Slovenian DOCX draft with visible Word revisions.</span>
</div>
""",
    unsafe_allow_html=True,
)
api_key = get_openai_key()
make_ai_drafts = st.checkbox(
    "Generate Slovenian tracked-change draft DOCX",
    help="Creates draft Slovenian wording and applies it to a copy of the target DOCX with Word tracked changes. Requires OPENAI_API_KEY in Streamlit secrets.",
)
settings_left, settings_right = st.columns(2, gap="large")
with settings_left:
    model = st.text_input("OpenAI model", value="gpt-5.2", disabled=not make_ai_drafts)
with settings_right:
    max_ai_rows = st.number_input(
        "Maximum changed rows to translate",
        min_value=1,
        max_value=400,
        value=40,
        step=10,
        disabled=not make_ai_drafts,
        help="Use a smaller number for quick tests, then increase for the full document. These rows can be applied to the tracked draft.",
    )
if make_ai_drafts and not api_key:
    st.warning("Add OPENAI_API_KEY in Streamlit secrets before deploying/running AI draft translations.")

ready = source is not None and target is not None
if not ready:
    st.info("Upload the English source and Slovenian target to enable package creation.")

if st.button("Create QRD output package", type="primary", disabled=not ready):
    with st.spinner("Reading DOCX files and preparing the QRD output..."):
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
    st.success("QRD output package ready.")
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
