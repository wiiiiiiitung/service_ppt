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

app = Flask(__name__)
app.secret_key = os.urandom(24)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
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

ALLOWED_EXTENSIONS = {".pdf", ".ppt", ".pptx", ".docx"}

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

        # Build the PPTX
        build_pptx(TEMPLATE_PATH, agenda, input_files, out_path,
                   library_paths=library_paths, intro_path=INTRO_PATH)

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
