"""Flask web app for Sunday worship PPT generator."""

import json
import os
import shutil
import uuid

from flask import (
    Flask,
    jsonify,
    render_template,
    request,
    send_file,
    session,
)

from pdf_parser import parse_agenda
from ppt_builder import build_pptx
from slide_planner import plan_slides
from slide_finder import find_slide
from bible_fetcher import get_testament
from file_converter import convert_legacy

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or os.urandom(24)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Use /tmp for Vercel serverless (read-only filesystem), local uploads/ for development
UPLOAD_DIR = os.path.join("/tmp", "uploads") if os.environ.get("VERCEL") else os.path.join(BASE_DIR, "uploads")
# Local fixed intro template (first few pre-worship slides)
_intro_candidate = os.path.join(BASE_DIR, "template", "intro.pptx")
INTRO_PATH = _intro_candidate if os.path.exists(_intro_candidate) else None
# Default template: latest example output
_example_dir = os.path.join(BASE_DIR, "example")
TEMPLATE_PATH = None
if os.path.isdir(_example_dir):
    for _sub in sorted(os.listdir(_example_dir), reverse=True):
        _out = os.path.join(_example_dir, _sub, "output")
        if os.path.isdir(_out):
            for _f in os.listdir(_out):
                if _f.endswith(".pptx"):
                    TEMPLATE_PATH = os.path.join(_out, _f)
                    break
        if TEMPLATE_PATH:
            break

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

        saved.append({"name": name, "size": os.path.getsize(dest)})

    return jsonify({"uploaded": saved})


@app.route("/api/files", methods=["GET"])
def list_uploaded():
    """List uploaded files for this session."""
    upload_dir = _session_dir()
    files = []
    for fname in os.listdir(upload_dir):
        fpath = os.path.join(upload_dir, fname)
        files.append({
            "name": fname,
            "size": os.path.getsize(fpath),
            "source": "local",
        })
    return jsonify({"files": files})


@app.route("/api/files/<filename>", methods=["DELETE"])
def delete_file(filename):
    """Remove an uploaded file."""
    upload_dir = _session_dir()
    fpath = os.path.join(upload_dir, filename)
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

    # Find the PDF
    pdf_path = None
    input_files = {}

    for fname in os.listdir(upload_dir):
        fpath = os.path.join(upload_dir, fname)
        ext = os.path.splitext(fname)[1].lower()
        if ext == ".pdf":
            pdf_path = fpath
        else:
            input_files[fname] = fpath

    if not pdf_path:
        return jsonify({"error": "No PDF agenda found. Please upload the agenda PDF."}), 400

    if not TEMPLATE_PATH or not os.path.exists(TEMPLATE_PATH):
        return jsonify({"error": "Template PPTX not found."}), 500

    try:
        from pptx import Presentation
        template = Presentation(TEMPLATE_PATH)

        # Load libraries
        libraries = [template]
        library_paths = []
        if os.path.isdir(_example_dir):
            for sub in sorted(os.listdir(_example_dir)):
                out_dir = os.path.join(_example_dir, sub, "output")
                if os.path.isdir(out_dir):
                    for f in os.listdir(out_dir):
                        if f.endswith(".pptx"):
                            library_paths.append(os.path.join(out_dir, f))
        for p in library_paths:
            try:
                libraries.append(Presentation(p))
            except Exception:
                pass

        # Parse the agenda
        agenda = parse_agenda(pdf_path)

        # Plan slides (no bible_page yet, user will fill it in)
        slides_spec = plan_slides(template, libraries, agenda, input_files, skip_intro=False, bible_page=None)

        # Build a summary per section
        plan_summary = []
        current_section = None
        current_count = 0
        current_source = None

        for spec in slides_spec:
            stype = spec["type"]

            # Track sections
            if stype == "copy_template":
                if current_section != "fixed":
                    if current_section and current_count:
                        plan_summary.append({
                            "section": current_section,
                            "label": current_source or current_section,
                            "source": "template",
                            "slides": current_count,
                            "status": "ok"
                        })
                    current_section = "fixed"
                    current_count = 0
                current_count += 1

            elif stype == "copy_external":
                if current_section != "input":
                    if current_section and current_count:
                        plan_summary.append({
                            "section": current_section,
                            "label": current_source or current_section,
                            "source": "input",
                            "slides": current_count,
                            "status": "ok"
                        })
                    current_section = "input"
                    current_count = 0
                current_count += 1

            elif stype == "blank":
                pass

            elif stype == "hymn_placeholder":
                plan_summary.append({
                    "section": "hymn",
                    "label": spec.get("label", "Hymn"),
                    "source": "⚠ not found — placeholder",
                    "slides": 1,
                    "status": "warning"
                })

            elif stype == "anthem_title":
                plan_summary.append({
                    "section": "anthem",
                    "label": f"獻詩: {spec.get('title', '')}",
                    "source": "docx or library",
                    "slides": 1,
                    "status": "ok"
                })

            elif stype == "scripture_title":
                item = spec.get("item", {})
                ref = item.get("title", "")
                testament = get_testament(ref.split()[0]) if ref else "新約"
                plan_summary.append({
                    "section": "scripture",
                    "label": ref,
                    "source": "bible-api.com or library",
                    "slides": 1,
                    "status": "ok",
                    "testament": testament,
                    "page": spec.get("bible_page")
                })

            elif stype == "sermon_title":
                plan_summary.append({
                    "section": "sermon",
                    "label": f"今日信息: {spec.get('title', '')}",
                    "source": "generated",
                    "slides": 1,
                    "status": "ok"
                })

            elif stype == "announcement":
                plan_summary.append({
                    "section": "announcement",
                    "label": spec.get("section", "報告"),
                    "source": "generated",
                    "slides": 1,
                    "status": "ok"
                })

        # Append final section if one is being tracked
        if current_section and current_count:
            plan_summary.append({
                "section": current_section,
                "label": current_source or current_section,
                "source": "template" if current_section == "fixed" else "input",
                "slides": current_count,
                "status": "ok"
            })

        return jsonify({"plan": plan_summary})

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


# ── Generate ───────────────────────────────────────────────────────────────────

@app.route("/api/generate", methods=["POST"])
def generate():
    """Parse the PDF and generate the worship PPTX."""
    upload_dir = _session_dir()

    # Find the PDF
    pdf_path = None
    input_files = {}

    for fname in os.listdir(upload_dir):
        fpath = os.path.join(upload_dir, fname)
        ext = os.path.splitext(fname)[1].lower()
        if ext == ".pdf":
            pdf_path = fpath
        else:
            input_files[fname] = fpath

    if not pdf_path:
        return jsonify({"error": "No PDF agenda found. Please upload the agenda PDF."}), 400

    if not TEMPLATE_PATH or not os.path.exists(TEMPLATE_PATH):
        return jsonify({"error": "Template PPTX not found. Place a completed PPTX in example/<date>/output/"}), 500

    try:
        # Parse the agenda PDF
        agenda = parse_agenda(pdf_path)

        # Generate output filename from date
        date_str = agenda.get("date", "").replace("/", "")
        out_name = f"Sunday Worship {date_str}.pptx"
        out_path = os.path.join(upload_dir, out_name)

        # Collect all example output PPTX files as slide library
        library_paths = []
        if os.path.isdir(_example_dir):
            for sub in sorted(os.listdir(_example_dir)):
                out_dir = os.path.join(_example_dir, sub, "output")
                if os.path.isdir(out_dir):
                    for f in os.listdir(out_dir):
                        if f.endswith(".pptx"):
                            library_paths.append(os.path.join(out_dir, f))

        # Get bible page number from request (optional)
        bible_page = request.json.get("bible_page") if request.json else None

        # Build the PPTX
        build_pptx(TEMPLATE_PATH, agenda, input_files, out_path,
                   library_paths=library_paths, intro_path=INTRO_PATH, bible_page=bible_page)

        # Store output path in session
        session["output_file"] = out_path
        session["output_name"] = out_name

        return jsonify({
            "success": True,
            "filename": out_name,
            "agenda_summary": {
                "date": agenda.get("date"),
                "worship_order": [
                    {"type": i.get("type"), "title": i.get("title"), "number": i.get("number")}
                    for i in agenda.get("worship_order", [])
                ],
                "sermon_title": agenda.get("sermon_outline", {}).get("title"),
                "announcement_sections": list(agenda.get("announcements", {}).keys()),
            },
        })

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


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
    import os
    port = int(os.environ.get("PORT", 5001))
    app.run(debug=False, host="0.0.0.0", port=port)
