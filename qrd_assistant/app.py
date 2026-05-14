from __future__ import annotations

import csv
import html
import io
import json
import os
import re
import shutil
import sys
import uuid
import zipfile
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


APP_ROOT = Path(__file__).resolve().parent
WORKSPACE_ROOT = APP_ROOT.parent
PROJECTS_DIR = APP_ROOT / "projects"
PROJECTS_DIR.mkdir(exist_ok=True)

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = f"{{{NS['w']}}}"
XML_NS = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("w", NS["w"])

# Keep Slovenian terms in ASCII escape form so the deployed app is not affected
# by editor, terminal, or keyboard encoding differences.
SL_RULES = [
    ("solution for injection", "raztopina za injiciranje", "EMA form wording"),
    ("subcutaneous use", "subkutana uporaba", "Route wording"),
    ("proline", "prolin", "Excipient terminology"),
    ("L-proline", "L-prolin", "Flag when English removes L-"),
    ("patient", "bolnik", "Prefer patient wording where English says patient"),
    ("participants", "udele\u017eenci", "Often becomes bolniki in clinical sections"),
    ("subjects", "preiskovanci", "Often becomes bolniki in clinical sections"),
    ("vial", "viala", "Container terminology"),
    ("package leaflet", "navodilo za uporabo", "QRD document part"),
    ("summary of product characteristics", "povzetek glavnih zna\u010dilnosti zdravila", "QRD document part"),
    ("section", "poglavje", "Cross-reference wording"),
]

TEMPLATE_HINTS = [
    "NAME OF THE MEDICINAL PRODUCT",
    "STATEMENT OF ACTIVE SUBSTANCE",
    "LIST OF EXCIPIENTS",
    "PHARMACEUTICAL FORM",
    "METHOD AND ROUTE",
    "SPECIAL STORAGE CONDITIONS",
    "NAME AND ADDRESS OF THE MARKETING AUTHORISATION HOLDER",
    "POVZETEK GLAVNIH ZNA\u010cILNOSTI ZDRAVILA",
    "IME ZDRAVILA",
    "NAVODILO ZA UPORABO",
    "SEZNAM POMO\u017dNIH SNOVI",
]


@dataclass
class UploadedFiles:
    source_docx: Path
    target_docx: Path
    template_en: Path | None
    template_sl: Path | None
    instructions: Path | None


def safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._#() -]+", "_", name).strip(" .")
    return cleaned or "uploaded.docx"


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\u00a0", " ")).strip()


def read_docx_xml(docx_path: Path, member: str = "word/document.xml") -> ET.Element:
    with zipfile.ZipFile(docx_path) as zf:
        return ET.fromstring(zf.read(member))


def iter_paragraphs(root: ET.Element) -> Iterable[ET.Element]:
    yield from root.iter(f"{W}p")


def node_text(node: ET.Element, mode: str = "accepted") -> str:
    parts: list[str] = []
    for child in node.iter():
        tag = child.tag
        if tag == f"{W}t" and child.text:
            in_del = any_parent_tag(child, f"{W}del")
            in_ins = any_parent_tag(child, f"{W}ins")
            if mode == "accepted" and not in_del:
                parts.append(child.text)
            elif mode == "original" and not in_ins:
                parts.append(child.text)
            elif mode == "all":
                parts.append(child.text)
        elif tag == f"{W}delText" and child.text:
            if mode in {"original", "all"}:
                parts.append(child.text)
    return "".join(parts)


def any_parent_tag(node: ET.Element, tag: str) -> bool:
    # ElementTree does not keep parent pointers. For DOCX revision extraction we
    # avoid this helper in hot paths and use recursive traversal below.
    return False


def paragraph_text_by_revision(p: ET.Element) -> tuple[str, str, str, str, str]:
    original: list[str] = []
    accepted: list[str] = []
    all_text: list[str] = []
    inserted: list[str] = []
    deleted: list[str] = []

    def walk(node: ET.Element, in_ins: bool = False, in_del: bool = False) -> None:
        tag = node.tag
        if tag == f"{W}ins":
            in_ins = True
        elif tag == f"{W}del":
            in_del = True

        if tag == f"{W}t" and node.text:
            all_text.append(node.text)
            if not in_del:
                accepted.append(node.text)
            if not in_ins:
                original.append(node.text)
            if in_ins:
                inserted.append(node.text)
        elif tag == f"{W}delText" and node.text:
            all_text.append(node.text)
            original.append(node.text)
            deleted.append(node.text)

        for child in list(node):
            walk(child, in_ins, in_del)

    walk(p)
    return (
        "".join(original),
        "".join(accepted),
        "".join(all_text),
        " ".join(x for x in inserted if x.strip()),
        " ".join(x for x in deleted if x.strip()),
    )


def extract_paragraphs(docx_path: Path) -> list[dict[str, str]]:
    root = read_docx_xml(docx_path)
    rows: list[dict[str, str]] = []
    for idx, p in enumerate(iter_paragraphs(root), start=1):
        original, accepted, all_text, inserted, deleted = paragraph_text_by_revision(p)
        if not norm(all_text):
            continue
        style_el = p.find(f"{W}pPr/{W}pStyle", NS)
        style = style_el.get(f"{W}val") if style_el is not None else ""
        rows.append(
            {
                "idx": str(idx),
                "style": style or "",
                "original": original,
                "accepted": accepted,
                "all": all_text,
                "inserted": inserted,
                "deleted": deleted,
                "has_revision": str(bool(norm(inserted) or norm(deleted))),
            }
        )
    return rows


def guess_section(text: str, current: str) -> str:
    value = norm(text)
    if not value:
        return current
    words = re.findall(r"[A-Za-z]+", value)
    tableish = bool(re.search(r"\b(mg|kg|ml|N|n|CI|SD|SE|LS|QMG|MG-ADL|MG-C)\b", value)) or bool(
        re.match(r"^[\u2248<>\u2264\u2265\d\s.,()%/-]+$", value)
    )
    if re.match(r"^\d{1,2}(?:\.\d+)?(?:\.|\s)[A-Z][A-Za-z].+", value):
        return value[:120]
    if value.upper() == value and len(value) > 16 and len(words) >= 3 and not tableish:
        return value[:120]
    if value in {"Mechanism of action", "Pharmacodynamic effects", "Clinical efficacy and safety", "Study MG0003", "Paediatric population"}:
        return value
    return current


def change_map(source_docx: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    section = ""
    for source_ordinal, para in enumerate(extract_paragraphs(source_docx), start=1):
        section = guess_section(para["accepted"] or para["original"], section)
        if para["has_revision"] == "True":
            original = norm(para["original"])
            revised = norm(para["accepted"])
            category = "TEMPLATE_OR_LABEL" if any(h.lower() in original.lower() for h in TEMPLATE_HINTS) else "PRODUCT_OR_MIXED"
            if not revised:
                category = "DELETION"
            rows.append(
                {
                    "idx": para["idx"],
                    "source_ordinal": str(source_ordinal),
                    "section": section,
                    "category_guess": category,
                    "original_text": original,
                    "revised_text": revised,
                    "inserted_text": norm(para["inserted"]),
                    "deleted_text": norm(para["deleted"]),
                }
            )
    return rows


def template_texts(paths: Iterable[Path | None]) -> list[str]:
    texts: list[str] = []
    for path in paths:
        if not path or not path.exists():
            continue
        try:
            for row in extract_paragraphs(path):
                value = norm(row["accepted"])
                if value:
                    texts.append(value)
        except Exception:
            continue
    return texts


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def build_reports(files: UploadedFiles, project_dir: Path) -> dict[str, Path]:
    out_dir = project_dir / "output"
    out_dir.mkdir(exist_ok=True)

    changes = change_map(files.source_docx)
    target_rows = extract_paragraphs(files.target_docx)
    source_rows = extract_paragraphs(files.source_docx)
    templates = template_texts([files.template_en, files.template_sl])

    write_csv(out_dir / "source_change_map.csv", changes)
    write_csv(out_dir / "source_paragraphs.csv", source_rows)
    write_csv(out_dir / "target_paragraphs.csv", target_rows)

    flags = terminology_flags(changes, target_rows, templates)
    write_csv(out_dir / "terminology_and_template_flags.csv", flags)

    draft_target = out_dir / f"{files.target_docx.stem}-REVIEW-COPY.docx"
    shutil.copy2(files.target_docx, draft_target)

    (out_dir / "review_brief.md").write_text(review_brief(changes, flags, files), encoding="utf-8")
    (out_dir / "chatgpt_prompt.md").write_text(chatgpt_prompt(changes, flags), encoding="utf-8")
    (out_dir / "query_log.md").write_text(query_log(changes, flags), encoding="utf-8")
    (out_dir / "project_manifest.json").write_text(
        json.dumps(
            {
                "created": datetime.now().isoformat(timespec="seconds"),
                "source": files.source_docx.name,
                "target": files.target_docx.name,
                "template_en": files.template_en.name if files.template_en else None,
                "template_sl": files.template_sl.name if files.template_sl else None,
                "change_count": len(changes),
                "flag_count": len(flags),
                "safety_model": "analysis-first; no unreviewed automatic translation applied",
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    zip_path = project_dir / "qrd_assistant_output.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in out_dir.iterdir():
            zf.write(item, item.name)

    return {
        "zip": zip_path,
        "brief": out_dir / "review_brief.md",
        "changes": out_dir / "source_change_map.csv",
        "flags": out_dir / "terminology_and_template_flags.csv",
    }


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists() or not path.read_text(encoding="utf-8-sig").strip():
        return []
    with path.open(newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def paired_change_rows(project_dir: Path) -> list[dict[str, str]]:
    out_dir = project_dir / "output"
    changes = read_csv_rows(out_dir / "source_change_map.csv")
    target_rows = read_csv_rows(out_dir / "target_paragraphs.csv")
    target_by_ordinal = {str(ordinal): row for ordinal, row in enumerate(target_rows, start=1)}
    paired: list[dict[str, str]] = []
    for change in changes:
        source_ordinal = change.get("source_ordinal") or ""
        target = target_by_ordinal.get(source_ordinal)
        row = dict(change)
        row["target_ordinal"] = source_ordinal
        row["target_idx"] = target.get("idx", "") if target else ""
        row["target_current_text"] = norm(target.get("accepted", "")) if target else ""
        row["target_match_method"] = "paragraph_ordinal" if target else "not_matched"
        paired.append(row)
    return paired


def _text_run(text: str, deleted: bool = False) -> ET.Element:
    run = ET.Element(f"{W}r")
    text_el = ET.SubElement(run, f"{W}delText" if deleted else f"{W}t")
    if text[:1].isspace() or text[-1:].isspace():
        text_el.set(f"{{{XML_NS}}}space", "preserve")
    text_el.text = text
    return run


def _revision_element(kind: str, revision_id: int, author: str, date: str, text: str) -> ET.Element:
    wrapper = ET.Element(
        f"{W}{kind}",
        {
            f"{W}id": str(revision_id),
            f"{W}author": author,
            f"{W}date": date,
        },
    )
    wrapper.append(_text_run(text, deleted=kind == "del"))
    return wrapper


def _diff_revision_children(original: str, revised: str, author: str, date: str, start_id: int) -> tuple[list[ET.Element], int]:
    old_tokens = re.findall(r"\s+|\S+", original)
    new_tokens = re.findall(r"\s+|\S+", revised)
    matcher = SequenceMatcher(None, old_tokens, new_tokens)
    children: list[ET.Element] = []
    revision_id = start_id

    for tag, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        old_text = "".join(old_tokens[old_start:old_end])
        new_text = "".join(new_tokens[new_start:new_end])
        if tag == "equal" and old_text:
            children.append(_text_run(old_text))
        elif tag == "delete" and old_text:
            children.append(_revision_element("del", revision_id, author, date, old_text))
            revision_id += 1
        elif tag == "insert" and new_text:
            children.append(_revision_element("ins", revision_id, author, date, new_text))
            revision_id += 1
        elif tag == "replace":
            if old_text:
                children.append(_revision_element("del", revision_id, author, date, old_text))
                revision_id += 1
            if new_text:
                children.append(_revision_element("ins", revision_id, author, date, new_text))
                revision_id += 1

    return children, revision_id


def _replace_paragraph_with_revision(p: ET.Element, revised: str, author: str, date: str, start_id: int) -> int:
    original = paragraph_text_by_revision(p)[1]
    p_pr = p.find(f"{W}pPr")
    for child in list(p):
        p.remove(child)
    if p_pr is not None:
        p.append(p_pr)
    children, next_id = _diff_revision_children(original, revised, author, date, start_id)
    for child in children:
        p.append(child)
    return next_id


def _enable_track_revisions(settings_xml: bytes) -> bytes:
    try:
        root = ET.fromstring(settings_xml)
    except ET.ParseError:
        return settings_xml
    if root.find(f"{W}trackRevisions") is None:
        root.append(ET.Element(f"{W}trackRevisions"))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def create_tracked_change_docx(
    target_docx: Path,
    revisions: list[dict[str, str]],
    output_docx: Path,
    author: str = "QRD Assistant",
) -> set[str]:
    applied_target_rows: set[str] = set()
    revision_date = datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
    revision_id = 1

    with zipfile.ZipFile(target_docx, "r") as zin:
        document_root = ET.fromstring(zin.read("word/document.xml"))
        paragraphs = {str(idx): p for idx, p in enumerate(iter_paragraphs(document_root), start=1)}

        for revision in revisions:
            target_idx = str(revision.get("target_idx") or revision.get("target_row") or "")
            revised = norm(revision.get("slovenian_revised") or "")
            delete_target = str(revision.get("delete_target") or "").lower() in {"true", "1", "yes"}
            if not target_idx or target_idx in applied_target_rows or (not revised and not delete_target):
                continue
            paragraph = paragraphs.get(target_idx)
            if paragraph is None:
                continue
            current = norm(paragraph_text_by_revision(paragraph)[1])
            if current == revised and not delete_target:
                continue
            revision_id = _replace_paragraph_with_revision(paragraph, revised, author, revision_date, revision_id)
            applied_target_rows.add(target_idx)

        document_xml = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = b"" if item.is_dir() else zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = document_xml
                elif item.filename == "word/settings.xml":
                    data = _enable_track_revisions(data)
                zout.writestr(item, data)

    return applied_target_rows


def terminology_flags(changes: list[dict[str, str]], target_rows: list[dict[str, str]], templates: list[str]) -> list[dict[str, str]]:
    target_blob = "\n".join(row["accepted"] for row in target_rows).lower()
    template_blob = "\n".join(templates).lower()
    flags: list[dict[str, str]] = []

    for change in changes:
        revised_lower = change["revised_text"].lower()
        original_lower = change["original_text"].lower()
        for en, sl, reason in SL_RULES:
            if en.lower() in revised_lower or en.lower() in original_lower:
                flags.append(
                    {
                        "source_row": change["idx"],
                        "section": change["section"],
                        "english_trigger": en,
                        "expected_slovenian": sl,
                        "reason": reason,
                        "target_contains_expected": str(sl.lower() in target_blob),
                        "template_contains_expected": str(sl.lower() in template_blob),
                        "action": "Check corresponding Slovenian paragraph and apply with Track Changes if relevant.",
                    }
                )

    for risky in ["L-prolin", "udeležencev", "preiskovancev", "razpredelnica"]:
        if risky.lower() in target_blob:
            flags.append(
                {
                    "source_row": "",
                    "section": "Global Slovenian consistency",
                    "english_trigger": "",
                    "expected_slovenian": risky,
                    "reason": "Potential consistency/template issue found in target.",
                    "target_contains_expected": "True",
                    "template_contains_expected": str(risky.lower() in template_blob),
                    "action": "Review occurrences manually before changing; may be context-dependent.",
                }
            )
    return flags


def review_brief(changes: list[dict[str, str]], flags: list[dict[str, str]], files: UploadedFiles) -> str:
    deletions = sum(1 for row in changes if row["category_guess"] == "DELETION")
    by_section: dict[str, int] = {}
    for row in changes:
        by_section[row["section"] or "(unknown)"] = by_section.get(row["section"] or "(unknown)", 0) + 1
    section_lines = "\n".join(f"- {section}: {count}" for section, count in sorted(by_section.items(), key=lambda x: x[1], reverse=True)[:12])
    return f"""# QRD Assistant Review Brief

Created: {datetime.now().strftime("%Y-%m-%d %H:%M")}

## Files

- Source EN: `{files.source_docx.name}`
- Target SL: `{files.target_docx.name}`
- EN template: `{files.template_en.name if files.template_en else "not uploaded"}`
- SL template: `{files.template_sl.name if files.template_sl else "not uploaded"}`

## Safety Position

This package is analysis-first. It does not claim to produce a client-final translation automatically.
Use it to identify source changes, template-sensitive areas, terminology flags, and safe review anchors.
The final client deliverable should still be prepared in Microsoft Word with Track Changes enabled.

## Summary

- Source paragraphs with tracked changes: {len(changes)}
- Full-deletion rows in source: {deletions}
- Terminology/template flags: {len(flags)}

## Highest-Volume Sections

{section_lines or "- No changed sections detected."}

## Recommended Workflow

1. Open `source_change_map.csv` beside the English source.
2. Open the Slovenian target in Microsoft Word with Track Changes enabled.
3. Work section by section, starting with high-volume sections and full deletions.
4. Use `terminology_and_template_flags.csv` to check EMA-required wording.
5. Use `query_log.md` for items that should not be guessed.
6. Save the client deliverable only from Microsoft Word, not from raw XML tools.
"""


def chatgpt_prompt(changes: list[dict[str, str]], flags: list[dict[str, str]]) -> str:
    preview = changes[:20]
    preview_lines = "\n".join(
        f"- Row {row['idx']} [{row['section']}]\n  EN old: {row['original_text']}\n  EN new: {row['revised_text']}"
        for row in preview
    )
    flag_lines = "\n".join(
        f"- {row['expected_slovenian']} ({row['reason']})" for row in flags[:25]
    )
    return f"""# Prompt For QRD Review In ChatGPT

You are assisting with an EMA QRD English-to-Slovenian regulatory update.

Requirements:
- Preserve EMA QRD template wording where applicable.
- Translate only changes visible in the English tracked source.
- Keep product-specific wording consistent across the document.
- Use Slovenian regulatory terminology.
- Do not guess ambiguous table/figure restructuring; mark it as a query.
- The client deliverable is a Slovenian Word document with Track Changes visible.

Important terminology/template flags:
{flag_lines or "- No flags generated."}

First source changes to review:
{preview_lines or "- No tracked changes detected."}

For each change, return:
- Whether it is template text, product text, or mixed.
- The Slovenian wording to apply.
- Whether the change is safe to apply directly.
- Any query/comment for the client or reviewer.
"""


def query_log(changes: list[dict[str, str]], flags: list[dict[str, str]]) -> str:
    deletion_rows = [row for row in changes if row["category_guess"] == "DELETION"]
    lines = ["# QRD Query Log", ""]
    if deletion_rows:
        lines.append("## Full Deletions To Verify In Word")
        for row in deletion_rows[:100]:
            lines.append(f"- Source row {row['idx']} in `{row['section']}` deletes: {row['original_text'][:220]}")
        lines.append("")
    if flags:
        lines.append("## Terminology/Template Checks")
        for row in flags[:150]:
            lines.append(f"- {row['section']}: check `{row['expected_slovenian']}`. {row['action']}")
    if not deletion_rows and not flags:
        lines.append("No automatic queries generated.")
    return "\n".join(lines) + "\n"


class QRDHandler(BaseHTTPRequestHandler):
    server_version = "QRDAssistant/0.1"

    def do_GET(self) -> None:
        if self.path in {"/", "/index.html"}:
            self.send_html(index_page())
            return
        if self.path.startswith("/download/"):
            project_id = self.path.rsplit("/", 1)[-1]
            zip_path = PROJECTS_DIR / project_id / "qrd_assistant_output.zip"
            if not zip_path.exists():
                self.send_error(HTTPStatus.NOT_FOUND, "Output package not found")
                return
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Disposition", 'attachment; filename="qrd_assistant_output.zip"')
            self.send_header("Content-Length", str(zip_path.stat().st_size))
            self.end_headers()
            with zip_path.open("rb") as f:
                shutil.copyfileobj(f, self.wfile)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        if self.path != "/process":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        try:
            project_id = uuid.uuid4().hex[:10]
            project_dir = PROJECTS_DIR / project_id
            upload_dir = project_dir / "uploads"
            upload_dir.mkdir(parents=True)
            files = self.parse_upload(upload_dir)
            outputs = build_reports(files, project_dir)
            brief = html.escape(outputs["brief"].read_text(encoding="utf-8")[:5000])
            self.send_html(result_page(project_id, brief))
        except Exception as exc:
            self.send_response(HTTPStatus.INTERNAL_SERVER_ERROR)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(error_page(exc).encode("utf-8"))

    def parse_upload(self, upload_dir: Path) -> UploadedFiles:
        import cgi

        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})

        def save_field(name: str, required: bool = False) -> Path | None:
            item = form[name] if name in form else None
            if item is None or not getattr(item, "filename", ""):
                if required:
                    raise ValueError(f"Missing required upload: {name}")
                return None
            filename = safe_filename(item.filename)
            path = upload_dir / filename
            with path.open("wb") as f:
                shutil.copyfileobj(item.file, f)
            if path.suffix.lower() != ".docx":
                raise ValueError(f"{filename} is not a .docx file")
            return path

        source = save_field("source_docx", required=True)
        target = save_field("target_docx", required=True)
        assert source and target
        return UploadedFiles(
            source_docx=source,
            target_docx=target,
            template_en=save_field("template_en"),
            template_sl=save_field("template_sl"),
            instructions=save_field("instructions"),
        )

    def send_html(self, body: str) -> None:
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(body.encode("utf-8"))

    def log_message(self, fmt: str, *args: object) -> None:
        sys.stderr.write("%s - %s\n" % (self.log_date_time_string(), fmt % args))


def page_shell(title: str, body: str) -> str:
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>{CSS}</style>
</head>
<body>
  <main class="shell">{body}</main>
</body>
</html>"""


def index_page() -> str:
    return page_shell(
        "QRD Assistant",
        """<section class="header">
  <div>
    <p class="eyebrow">EMA QRD Slovenian workflow</p>
    <h1>QRD Assistant</h1>
  </div>
  <span class="badge">accuracy-first</span>
</section>

<section class="workspace">
<form class="panel" method="post" action="/process" enctype="multipart/form-data">
  <label>
    <span>English source with tracked changes</span>
    <input required type="file" name="source_docx" accept=".docx">
  </label>
  <label>
    <span>Slovenian target document</span>
    <input required type="file" name="target_docx" accept=".docx">
  </label>
  <label>
    <span>English QRD template <em>optional</em></span>
    <input type="file" name="template_en" accept=".docx">
  </label>
  <label>
    <span>Slovenian QRD template <em>optional</em></span>
    <input type="file" name="template_sl" accept=".docx">
  </label>
  <label>
    <span>Instructions document <em>optional</em></span>
    <input type="file" name="instructions" accept=".docx">
  </label>
  <button type="submit">Create QRD review package</button>
</form>

<aside class="sidebar" aria-label="Instructions">
  <h2>Use This App</h2>
  <ol class="steps">
    <li><strong>Upload the English source.</strong><span>Use the Word file that contains tracked changes.</span></li>
    <li><strong>Upload the Slovenian target.</strong><span>Use the file you need to update for the client.</span></li>
    <li><strong>Add QRD templates if available.</strong><span>The button still works without templates or instructions.</span></li>
    <li><strong>Create the package.</strong><span>The app reads the files and prepares review materials.</span></li>
    <li><strong>Work from the ZIP.</strong><span>Use the change map, flags, and query log while editing the Slovenian file in Word.</span></li>
  </ol>
  <div class="note">
    <h3>Important</h3>
    <p>The app does not overwrite your client document. Final tracked changes should still be made and checked in Microsoft Word.</p>
  </div>
</aside>
</section>

<section class="grid">
  <article>
    <h2>What it produces</h2>
    <p>Change map, paragraph exports, terminology/template flags, query log, ChatGPT prompt, and a clean target review copy.</p>
  </article>
  <article>
    <h2>Safety rule</h2>
    <p>The app does not silently translate or overwrite client files. Final tracked changes should be applied and validated in Microsoft Word.</p>
  </article>
</section>""",
    )


def result_page(project_id: str, brief: str) -> str:
    return page_shell(
        "QRD Assistant Output",
        f"""<section class="header">
  <div>
    <p class="eyebrow">Package ready</p>
    <h1>Review files generated</h1>
  </div>
  <a class="buttonlink" href="/download/{html.escape(project_id)}">Download ZIP</a>
</section>
<pre class="brief">{brief}</pre>
<a class="back" href="/">Start another project</a>""",
    )


def error_page(exc: Exception) -> str:
    return page_shell(
        "QRD Assistant Error",
        f"""<section class="header"><div><p class="eyebrow">Error</p><h1>Processing stopped</h1></div></section>
<pre class="brief">{html.escape(str(exc))}</pre>
<a class="back" href="/">Back</a>""",
    )


CSS = """
:root {
  color-scheme: light;
  font-family: Inter, Segoe UI, Arial, sans-serif;
  background: #f5f7f8;
  color: #172026;
}
body { margin: 0; }
.shell { width: min(1180px, calc(100% - 32px)); margin: 32px auto; }
.header { display: flex; justify-content: space-between; gap: 24px; align-items: end; margin-bottom: 22px; }
.eyebrow { margin: 0 0 6px; color: #53616b; font-size: 13px; text-transform: uppercase; letter-spacing: .08em; }
h1 { margin: 0; font-size: 36px; line-height: 1.1; }
h2 { margin: 0 0 8px; font-size: 18px; }
.badge, .buttonlink, button {
  border: 1px solid #0f6b63;
  background: #0f6b63;
  color: white;
  border-radius: 6px;
  padding: 10px 14px;
  font-weight: 700;
  text-decoration: none;
}
.badge { background: #e8f3f1; color: #0f6b63; }
.workspace {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 340px;
  gap: 16px;
  align-items: start;
}
.panel {
  display: grid;
  gap: 14px;
  padding: 20px;
  background: white;
  border: 1px solid #d8e0e4;
  border-radius: 8px;
  box-shadow: 0 8px 20px rgba(23,32,38,.06);
}
label { display: grid; gap: 7px; font-weight: 700; color: #25313a; }
label em {
  color: #6b7a84;
  font-style: normal;
  font-weight: 600;
  font-size: 13px;
}
input[type=file] {
  width: 100%;
  box-sizing: border-box;
  border: 1px solid #c8d2d8;
  border-radius: 6px;
  padding: 10px;
  background: #fbfcfd;
}
button { cursor: pointer; font-size: 15px; margin-top: 4px; }
.sidebar {
  background: #ffffff;
  border: 1px solid #d8e0e4;
  border-radius: 8px;
  padding: 18px;
  box-shadow: 0 8px 20px rgba(23,32,38,.06);
  position: sticky;
  top: 18px;
}
.sidebar h2 { margin-bottom: 14px; }
.sidebar h3 { margin: 0 0 6px; font-size: 15px; }
.steps {
  list-style: none;
  margin: 0;
  padding: 0;
  display: grid;
  gap: 12px;
  counter-reset: step;
}
.steps li {
  display: grid;
  grid-template-columns: 28px minmax(0, 1fr);
  column-gap: 10px;
  row-gap: 3px;
  color: #25313a;
}
.steps li::before {
  counter-increment: step;
  content: counter(step);
  width: 28px;
  height: 28px;
  border-radius: 50%;
  display: inline-grid;
  place-items: center;
  background: #e8f3f1;
  color: #0f6b63;
  font-weight: 800;
}
.steps strong { align-self: center; }
.steps span {
  grid-column: 2;
  color: #53616b;
  line-height: 1.42;
  font-size: 14px;
}
.note {
  margin-top: 16px;
  padding-top: 14px;
  border-top: 1px solid #e2e8eb;
}
.note p {
  margin: 0;
  color: #43515a;
  line-height: 1.5;
  font-size: 14px;
}
.grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; margin-top: 14px; }
article { background: white; border: 1px solid #d8e0e4; border-radius: 8px; padding: 16px; }
article p { margin: 0; color: #43515a; line-height: 1.5; }
.brief {
  white-space: pre-wrap;
  background: white;
  border: 1px solid #d8e0e4;
  border-radius: 8px;
  padding: 18px;
  overflow: auto;
  line-height: 1.45;
}
.back { display: inline-block; margin-top: 14px; color: #0f6b63; font-weight: 700; }
@media (max-width: 720px) {
  .header, .grid, .workspace { display: block; }
  .badge, .buttonlink { display: inline-block; margin-top: 14px; }
  article, .sidebar { margin-top: 14px; }
  .sidebar { position: static; }
}
"""


def main() -> None:
    port = int(os.environ.get("QRD_ASSISTANT_PORT", "8765"))
    server = ThreadingHTTPServer(("127.0.0.1", port), QRDHandler)
    print(f"QRD Assistant running at http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
