"""Flask web app for Sunday worship PPT generator."""

import logging
import os
import shutil
import uuid

from flask import Flask, jsonify, render_template, request, send_file, session
from pptx import Presentation

import bible_pages
import deck_sources
from bible_fetcher import get_testament
from file_converter import convert_directory, convert_legacy
from pdf_parser import parse_agenda
from ppt_builder import build_pptx
from slide_planner import plan_match_items, plan_slides, unused_inputs

logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or os.urandom(24)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Use /tmp for serverless (read-only filesystem), local uploads/ for development
UPLOAD_DIR = (os.path.join("/tmp", "uploads") if os.environ.get("VERCEL")
              else os.path.join(BASE_DIR, "uploads"))

ALLOWED_EXTENSIONS = {".pdf", ".ppt", ".pptx", ".doc", ".docx"}

os.makedirs(UPLOAD_DIR, exist_ok=True)


def _session_dir():
    sid = session.get("sid")
    if not sid:
        sid = str(uuid.uuid4())
        session["sid"] = sid
    d = os.path.join(UPLOAD_DIR, sid)
    os.makedirs(d, exist_ok=True)
    return d


def _discover_input_files(upload_dir):
    """Split uploaded files into the agenda PDF path and the rest (by name)."""
    pdf_path = None
    input_files = {}
    for fname in sorted(os.listdir(upload_dir)):
        fpath = os.path.join(upload_dir, fname)
        if not os.path.isfile(fpath):
            continue
        ext = os.path.splitext(fname)[1].lower()
        if ext == ".pdf":
            pdf_path = fpath
        elif ext in ALLOWED_EXTENSIONS:
            input_files[fname] = fpath
    return pdf_path, input_files


def _error_response(e):
    import traceback
    logger.exception("Request failed")
    return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


def _build_plan_summary(slides_spec):
    """Collapse a slide spec list into a per-section summary for the UI."""
    plan_summary = []
    state = {"section": None, "count": 0, "label": None}

    def flush(source):
        if state["section"] and state["count"]:
            plan_summary.append({
                "section": state["section"],
                "label": state["label"] or state["section"],
                "source": source,
                "slides": state["count"],
                "status": "ok",
            })
        state["section"] = None
        state["count"] = 0
        state["label"] = None

    def run(section, source, label=None):
        if state["section"] != section:
            flush(source)
            state["section"] = section
            state["label"] = label
        state["count"] += 1

    for spec in slides_spec:
        stype = spec["type"]

        if stype == "copy_template":
            run("fixed", "template", "固定投影片")

        elif stype == "copy_external":
            run("input", "input", "輸入檔案")

        elif stype == "blank":
            continue

        elif stype == "anthem_lyrics":
            run("anthem_lyrics", "docx", f"獻詩歌詞: {spec.get('title', '')}")

        elif stype == "scripture_verses":
            run("scripture_verses", "bible-api.com", f"經文: {spec.get('ref', '')}")

        elif stype == "sermon_point":
            run("sermon_points", "generated", "今日信息大綱")

        else:
            flush("template" if state["section"] == "fixed" else "input")

            if stype == "hymn_placeholder":
                plan_summary.append({
                    "section": "hymn", "label": spec.get("label", "Hymn"),
                    "source": "⚠ 找不到檔案 — 使用佔位符",
                    "slides": 1, "status": "warning",
                })

            elif stype == "anthem_title":
                plan_summary.append({
                    "section": "anthem", "label": f"獻詩: {spec.get('title', '')}",
                    "source": "docx or library", "slides": 1, "status": "ok",
                })

            elif stype == "scripture_title":
                ref = spec.get("item", {}).get("title", "")
                testament = get_testament(ref.split()[0]) if ref else "新約"
                plan_summary.append({
                    "section": "scripture", "label": ref,
                    "source": "bible-api.com", "slides": 1, "status": "ok",
                    "testament": testament, "page": spec.get("bible_page"),
                })

            elif stype == "sermon_title":
                plan_summary.append({
                    "section": "sermon", "label": f"今日信息: {spec.get('title', '')}",
                    "source": "generated", "slides": 1, "status": "ok",
                })

            elif stype == "announcement":
                plan_summary.append({
                    "section": "announcement", "label": spec.get("section", "報告"),
                    "source": "generated", "slides": 1, "status": "ok",
                })

    flush("template" if state["section"] == "fixed" else "input")
    return plan_summary


def _load_libraries(fixed_deck):
    """Fixed deck first, then the past output decks, as Presentations."""
    libraries = [fixed_deck]
    for p in deck_sources.library_paths(exclude=[deck_sources.fixed_path()]):
        try:
            libraries.append(Presentation(p))
        except Exception:
            logger.exception("Failed to load library PPTX %s; skipping", p)
    return libraries


@app.route("/")
def index():
    return render_template("index.html")


# ── File upload ────────────────────────────────────────────────────────────────

@app.route("/api/upload", methods=["POST"])
def upload_file():
    """Upload one or more input files."""
    if "files" not in request.files:
        return jsonify({"error": "No files provided"}), 400

    upload_dir = _session_dir()
    saved = []
    unconverted = []

    for f in request.files.getlist("files"):
        name = f.filename
        ext = os.path.splitext(name)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            continue
        dest = os.path.join(upload_dir, name)
        f.save(dest)

        # Auto-convert legacy .ppt/.doc to modern formats
        if ext in (".ppt", ".doc"):
            new_path = convert_legacy(dest, upload_dir)
            if new_path:
                os.remove(dest)
                dest = new_path
                name = os.path.basename(new_path)
            else:
                unconverted.append(name)

        saved.append({"name": name, "size": os.path.getsize(dest)})

    body = {"uploaded": saved}
    if unconverted:
        body["warning"] = (
            "無法轉換舊格式檔案（需要安裝 LibreOffice）："
            + "、".join(unconverted)
            + "。請在 PowerPoint 另存為 .pptx / .docx 後重新上傳。"
        )
    return jsonify(body)


@app.route("/api/files", methods=["GET"])
def list_uploaded():
    """List uploaded files for this session."""
    upload_dir = _session_dir()
    files = []
    for fname in sorted(os.listdir(upload_dir)):
        fpath = os.path.join(upload_dir, fname)
        if os.path.isfile(fpath):
            files.append({"name": fname, "size": os.path.getsize(fpath),
                          "source": "local"})
    return jsonify({"files": files})


@app.route("/api/files/<filename>", methods=["DELETE"])
def delete_file(filename):
    """Remove an uploaded file."""
    upload_dir = _session_dir()
    fpath = os.path.join(upload_dir, os.path.basename(filename))
    if os.path.exists(fpath):
        os.remove(fpath)
    return jsonify({"deleted": filename})


@app.route("/api/clear", methods=["POST"])
def clear_files():
    """Clear all uploaded files for this session."""
    upload_dir = _session_dir()
    shutil.rmtree(upload_dir, ignore_errors=True)
    os.makedirs(upload_dir, exist_ok=True)
    return jsonify({"cleared": True})


# ── Plan (preview) ────────────────────────────────────────────────────────────

@app.route("/api/plan", methods=["POST"])
def plan():
    """Preview the slide plan for the worship service (without generating PPTX)."""
    upload_dir = _session_dir()
    convert_directory(upload_dir)

    pdf_path, input_files = _discover_input_files(upload_dir)
    if not pdf_path:
        return jsonify({"error": "No PDF agenda found. Please upload the agenda PDF."}), 400

    fixed_path = deck_sources.fixed_path()
    if not fixed_path:
        return jsonify({"error": "template/fixed.pptx not found."}), 500

    try:
        fixed_deck = Presentation(fixed_path)
        libraries = _load_libraries(fixed_deck)
        agenda = parse_agenda(pdf_path)

        overrides = None
        if request.is_json:
            overrides = (request.get_json(silent=True) or {}).get("overrides")

        match_items = plan_match_items(agenda, input_files)
        slides_spec = plan_slides(fixed_deck, libraries, agenda, input_files,
                                  skip_intro=bool(deck_sources.intro_path()),
                                  bible_page=None, overrides=overrides)

        return jsonify({
            "plan": _build_plan_summary(slides_spec),
            "match_items": match_items,
            "unused_inputs": unused_inputs(slides_spec, input_files),
            "total_slides": len(slides_spec),
        })

    except Exception as e:
        return _error_response(e)


# ── Generate ───────────────────────────────────────────────────────────────────

@app.route("/api/generate", methods=["POST"])
def generate():
    """Parse the PDF and generate the worship PPTX."""
    upload_dir = _session_dir()
    convert_directory(upload_dir)

    pdf_path, input_files = _discover_input_files(upload_dir)
    if not pdf_path:
        return jsonify({"error": "No PDF agenda found. Please upload the agenda PDF."}), 400

    fixed_path = deck_sources.fixed_path()
    if not fixed_path:
        return jsonify({"error": "template/fixed.pptx not found."}), 500

    try:
        agenda = parse_agenda(pdf_path)

        # Name the file the way the archive does: YYYYMMDD, not MMDDYYYY.
        date_str = agenda.get("date", "")
        parts = date_str.split("/")
        stamp = f"{parts[2]}{parts[0]}{parts[1]}" if len(parts) == 3 else date_str.replace("/", "")
        out_name = f"Sunday Worship {stamp}.pptx" if stamp else "Sunday Worship.pptx"
        out_path = os.path.join(upload_dir, out_name)

        body = request.get_json(silent=True) or {}
        bible_page = body.get("bible_page")
        overrides = body.get("overrides")

        report = build_pptx(
            fixed_path, agenda, input_files, out_path,
            library_paths=deck_sources.library_paths(exclude=[fixed_path]),
            intro_path=deck_sources.intro_path(),
            bible_page=bible_page, overrides=overrides,
        )

        # Keep the page number the operator typed, so next time this passage
        # comes round it is already filled in.
        if bible_page:
            for item in agenda.get("worship_order", []):
                if item.get("type") == "scripture":
                    bible_pages.remember(item.get("title", ""), bible_page)
                    break

        session["output_file"] = out_path
        session["output_name"] = out_name

        return jsonify({
            "success": True,
            "filename": out_name,
            "unused_inputs": report["unused_inputs"],
            "agenda_summary": {
                "date": agenda.get("date"),
                "worship_order": [
                    {"type": i.get("type"), "title": i.get("title"),
                     "number": i.get("number")}
                    for i in agenda.get("worship_order", [])
                ],
                "sermon_title": agenda.get("sermon_outline", {}).get("title"),
                "announcement_sections": list(agenda.get("announcements", {}).keys()),
            },
        })

    except Exception as e:
        return _error_response(e)


@app.route("/api/download")
def download():
    """Download the generated PPTX."""
    out_path = session.get("output_file")
    out_name = session.get("output_name", "output.pptx")
    if not out_path or not os.path.exists(out_path):
        return jsonify({"error": "No generated file found. Please generate first."}), 404
    return send_file(
        out_path,
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(debug=False, host="0.0.0.0", port=port)
