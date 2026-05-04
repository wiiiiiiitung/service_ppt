"""
Convert legacy .ppt and .doc files to modern .pptx/.docx using LibreOffice.

Falls back gracefully if LibreOffice is not available — caller should check
the return value to know if conversion succeeded.
"""

import os
import shutil
import subprocess


def find_soffice():
    """Locate the LibreOffice CLI binary, or return None."""
    return shutil.which("soffice") or shutil.which("libreoffice")


def convert_legacy(file_path, output_dir=None):
    """
    Convert a .ppt or .doc file to .pptx/.docx in the same directory.

    Returns the new file path on success, or None if conversion failed
    (e.g. LibreOffice not installed, file already modern, unsupported format).
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in (".ppt", ".doc"):
        return None

    soffice = find_soffice()
    if not soffice:
        return None

    target_ext = "pptx" if ext == ".ppt" else "docx"
    output_dir = output_dir or os.path.dirname(file_path)

    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", target_ext,
             "--outdir", output_dir, file_path],
            check=True, capture_output=True, timeout=60
        )
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return None

    base = os.path.splitext(os.path.basename(file_path))[0]
    new_path = os.path.join(output_dir, f"{base}.{target_ext}")
    return new_path if os.path.exists(new_path) else None


def convert_directory(input_dir):
    """
    Convert all .ppt/.doc files in a directory to modern formats in place.

    Removes the originals on successful conversion. Returns a list of the
    converted file paths.
    """
    converted = []
    for fname in os.listdir(input_dir):
        fpath = os.path.join(input_dir, fname)
        if not os.path.isfile(fpath):
            continue
        new_path = convert_legacy(fpath, input_dir)
        if new_path:
            converted.append(new_path)
            try:
                os.remove(fpath)
            except OSError:
                pass
    return converted
