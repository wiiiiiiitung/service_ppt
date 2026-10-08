"""
Shared fixture plumbing for the regression harness.

A fixture is one `example_2/<date>/` folder holding the week's input files
(agenda PDF, hymn/reading PPTXs, anthem DOCX) alongside the hand-finished
`Sunday Worship ....pptx` reference deck that the generator should reproduce.
"""

import glob
import os
import shutil

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIXTURE_DIR = os.path.join(BASE_DIR, "example_2")
LEGACY_EXTS = (".ppt", ".doc")
MODERN_EXTS = (".pptx", ".docx")


def weeks():
    """Fixture folder names (dates) that have both inputs and a reference deck."""
    if not os.path.isdir(FIXTURE_DIR):
        return []
    out = []
    for name in sorted(os.listdir(FIXTURE_DIR)):
        d = os.path.join(FIXTURE_DIR, name)
        if os.path.isdir(d) and reference_deck(name):
            out.append(name)
    return out


def reference_deck(week):
    """Path to the hand-finished reference deck for a week, or None."""
    hits = sorted(glob.glob(os.path.join(FIXTURE_DIR, week, "Sunday Worship*.pptx")))
    return hits[0] if hits else None


def stage_inputs(week, work_dir):
    """
    Copy a week's input files into work_dir, skipping the reference deck.

    Several fixtures ship both a legacy `.ppt`/`.doc` and a manually re-saved
    modern twin. Prefer the twin so the fixture exercises the same files the
    user would upload; convert anything left over, which is what the app does
    on upload.

    Returns (pdf_path, {filename: path}).
    """
    src = os.path.join(FIXTURE_DIR, week)
    shutil.rmtree(work_dir, ignore_errors=True)
    os.makedirs(work_dir, exist_ok=True)

    names = sorted(os.listdir(src))
    stems_with_twin = {
        os.path.splitext(n)[0] for n in names
        if os.path.splitext(n)[1].lower() in MODERN_EXTS
    }

    for name in names:
        if name.startswith("Sunday Worship"):
            continue
        stem, ext = os.path.splitext(name)
        if ext.lower() in LEGACY_EXTS and stem in stems_with_twin:
            continue
        shutil.copy2(os.path.join(src, name), os.path.join(work_dir, name))

    from file_converter import convert_directory
    convert_directory(work_dir)

    pdf_path = None
    input_files = {}
    for name in sorted(os.listdir(work_dir)):
        path = os.path.join(work_dir, name)
        ext = os.path.splitext(name)[1].lower()
        if ext == ".pdf":
            pdf_path = path
        elif ext in MODERN_EXTS:
            input_files[name] = path
    return pdf_path, input_files
