"""Assemble the Sunday worship PPTX from the parsed agenda and input files."""

import copy
import glob
import os
import re
import shutil

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

from docx import Document

from bible_fetcher import fetch_verses, group_verses_for_slides
from styles import (
    get_style, apply_run_style, apply_paragraph_style, wrap_chinese_text,
    CYAN_BRIGHT, CYAN_SOFT, WHITE,
)

# Slide index mapping in the template PPTX (0-based)
TEMPLATE_SLIDES = {
    "logo": 0,           # Slide 1: logo/cover
    "welcome": 1,        # Slide 2: 歡迎
    "blank": 2,          # Slide 3: blank spacer
    "zoom_info": 3,      # Slide 4: Zoom translation
    "prepare": 4,        # Slide 5: 安靜敬虔
    "opening": 5,        # Slide 6: 開會詩
    "call_to_worship": 6,  # Slide 7: 宣召
    "blank2": 7,         # Slide 8: blank
    # Slides 9-11: Hymn 1 (dynamic)
    # Slide 12: blank
    "prayer": 12,        # Slide 13: 祈禱
    "lords_prayer_1": 13,  # Slide 14: 主禱文 part 1
    "lords_prayer_2": 14,  # Slide 15: 主禱文 part 2
    "creed_1": 15,       # Slide 16: 信仰告白 part 1
    "creed_2": 16,       # Slide 17: 信仰告白 part 2
    # Slide 18: blank
    # Slides 19-31: Responsive Reading (dynamic)
    # Slide 32: blank
    # Slide 33: Anthem title
    # Slides 34-40: Anthem lyrics
    # Slide 41: blank
    # Slide 42: Scripture title
    # Slides 43-47: Scripture text (we skip for now)
    # Slides 48-51: Sermon outline
    # Slide 52: blank
    # Slides 53-61: Hymn 2 (dynamic)
    # Slide 62: blank
    "offering_1": 62,    # Slide 63: 奉獻 intro
    "offering_2": 63,    # Slide 64: 奉獻 song
    "blank_offer": 64,   # Slide 65: blank
    "announce_title": 65,  # Slide 66: 報告
    # Slides 67-75: Announcements (dynamic)
    "blank_ann": 75,     # Slide 76: blank
    # Slides 77: Doxology (dynamic)
    "benediction": 77,   # Slide 78: 祝禱
    "quiet": 78,         # Slide 79: 默禱散會
    "website": 79,       # Slide 80: website
}


def build_pptx(template_path, agenda, input_files, output_path, library_paths=None, intro_path=None):
    """
    Build the worship PPTX.

    Args:
        template_path: Path to the template PPTX (for slide master/layouts)
        agenda: Parsed agenda dict from pdf_parser
        input_files: Dict of filename → filepath for input files
        output_path: Where to save the result
        library_paths: Additional PPTX files to search for hymn/reading slides
        intro_path: Optional local fixed-intro PPTX whose slides replace the
                    first few pre-worship slides of the template.
    """
    template = Presentation(template_path)
    use_intro = intro_path and os.path.exists(intro_path)

    # Build slide library: template + any extra library files
    libraries = [template]
    for p in (library_paths or []):
        if p != template_path and os.path.exists(p):
            try:
                libraries.append(Presentation(p))
            except Exception:
                pass

    # Plan all slides; when using intro template, skip the fixed intro specs
    slides_to_add = _plan_slides(template, libraries, agenda, input_files,
                                 skip_intro=use_intro)

    # Choose output base: intro PPTX (keeps its images intact) or template
    base_path = intro_path if use_intro else template_path
    shutil.copy2(base_path, output_path)
    out_prs = Presentation(output_path)

    if not use_intro:
        _clear_slides(out_prs)
    # If using intro, keep its existing slides and append the rest

    for slide_spec in slides_to_add:
        _add_slide(out_prs, template, slide_spec, agenda)

    out_prs.save(output_path)
    return output_path


def _plan_slides(template, libraries, agenda, input_files, skip_intro=False):
    """Build the ordered list of slide specs for the service."""
    order = agenda.get("worship_order", [])
    slides = []

    # === Fixed pre-worship slides ===
    # When skip_intro is True, those slides come from the intro PPTX (used as
    # the output base) and don't need to be appended here.
    if not skip_intro:
        slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["logo"]})
        slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["welcome"]})
        slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["blank"]})
        slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["zoom_info"]})
        slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["prepare"]})
        slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["opening"]})

    # === Worship order items ===
    for item in order:
        itype = item.get("type")

        if itype == "call_to_worship":
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["call_to_worship"]})
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["blank2"]})

        elif itype == "hymn":
            hymn_slides = _get_hymn_slides(libraries, item, input_files)
            slides.extend(hymn_slides)
            slides.append({"type": "blank"})

        elif itype == "prayer":
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["prayer"]})
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["lords_prayer_1"]})
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["lords_prayer_2"]})

        elif itype == "creed":
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["creed_1"]})
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["creed_2"]})
            slides.append({"type": "blank"})

        elif itype == "responsive":
            reading_slides = _get_reading_slides(libraries, item, input_files)
            slides.extend(reading_slides)
            slides.append({"type": "blank"})

        elif itype == "anthem":
            anthem_slides = _get_anthem_slides(libraries, item, input_files)
            slides.extend(anthem_slides)
            slides.append({"type": "blank"})

        elif itype == "scripture":
            scripture_slides = _get_scripture_slides(libraries, item)
            slides.extend(scripture_slides)
            # No blank between scripture and sermon (they follow directly)

        elif itype == "sermon":
            sermon_slides = _get_sermon_slides(agenda)
            slides.extend(sermon_slides)
            slides.append({"type": "blank"})

        elif itype == "offering":
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["offering_1"]})
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["offering_2"]})
            slides.append({"type": "blank"})

        elif itype == "announcements":
            ann_slides = _get_announcement_slides(agenda)
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["announce_title"]})
            slides.extend(ann_slides)
            slides.append({"type": "blank"})

        elif itype == "doxology":
            dox_slides = _get_hymn_slides(libraries, item, input_files, is_doxology=True)
            slides.extend(dox_slides)

        elif itype == "benediction":
            slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["benediction"]})

    slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["quiet"]})
    slides.append({"type": "copy_template", "index": TEMPLATE_SLIDES["website"]})

    return slides


def _get_hymn_slides(libraries, item, input_files, is_doxology=False):
    """Get slides for a hymn item, from input file or slide library."""
    num = item.get("number")
    title = item.get("title", "")

    # Try readable input file (PPTX only)
    matched = _match_file(num, title, input_files, [".pptx"])
    if matched:
        try:
            src = Presentation(matched)
            if src.slides:
                return [{"type": "copy_external", "prs": src, "index": i}
                        for i in range(len(src.slides))]
        except Exception:
            pass

    # Search all libraries
    for lib in libraries:
        indices = _find_hymn_slides_in_library(lib, num, title)
        if indices:
            return [{"type": "copy_external", "prs": lib, "index": i} for i in indices]

    label = f"{'頌榮' if is_doxology else '聖詩'} {num}: {title}" if num else title
    return [{"type": "hymn_placeholder", "label": label}]


def _get_reading_slides(libraries, item, input_files):
    """Get slides for the responsive reading."""
    num = item.get("number")
    title = item.get("title", "")

    matched = _match_file(num, title, input_files, [".pptx"])
    if matched:
        try:
            src = Presentation(matched)
            if src.slides:
                return [{"type": "copy_external", "prs": src, "index": i}
                        for i in range(len(src.slides))]
        except Exception:
            pass

    for lib in libraries:
        indices = _find_reading_slides_in_library(lib, num, title)
        if indices:
            return [{"type": "copy_external", "prs": lib, "index": i} for i in indices]

    return [{"type": "hymn_placeholder", "label": f"啟應文 {num}: {title}"}]


def _get_anthem_slides(libraries, item, input_files):
    """Get anthem title + lyrics slides from DOCX or library."""
    title = item.get("title", "")
    slides = [{"type": "anthem_title", "title": title}]

    # Try DOCX for lyrics (also match .doc/.DOC)
    matched = _match_file(None, title, input_files, [".docx", ".doc"])
    if matched:
        try:
            lyrics = _parse_docx_lyrics(matched)
            clean_title = re.sub(r"\s+", "", title)
            verses = [v for v in lyrics if re.sub(r"\s+", "", v) != clean_title]
            if verses:
                for chunk in _group_anthem_verses(verses, max_lines=6):
                    slides.append({"type": "anthem_lyrics", "title": title, "lyrics": chunk})
                return slides
        except Exception:
            pass

    # Search libraries for anthem slides
    for lib in libraries:
        indices = _find_anthem_slides_in_library(lib, title)
        if indices:
            # First index is the "獻詩:" title slide — skip since we have our own
            content_indices = [i for i in indices if i != indices[0]]
            if not content_indices:
                content_indices = indices
            return [{"type": "anthem_title", "title": title}] + \
                   [{"type": "copy_external", "prs": lib, "index": i} for i in content_indices]

    return slides


def _group_anthem_verses(verses, max_lines=20):
    """
    Pack verses into slide-sized chunks: combine short verses, split long ones.
    Each chunk is a string with \n-separated lines.
    """
    groups = []
    current = []
    current_lines = 0

    def flush():
        nonlocal current, current_lines
        if current:
            groups.append("\n".join(current))
            current = []
            current_lines = 0

    for v in verses:
        v_lines = v.split("\n")
        # Verse too tall — flush & split it across multiple slides
        if len(v_lines) > max_lines:
            flush()
            for i in range(0, len(v_lines), max_lines):
                groups.append("\n".join(v_lines[i:i + max_lines]))
            continue
        if current_lines + len(v_lines) > max_lines:
            flush()
        current.append(v)
        current_lines += len(v_lines)

    flush()
    return groups


def _get_scripture_slides(libraries, item):
    """
    Get scripture slides.

    Priority:
      1. Fetch text online from 和合本 (bible-api.com) — generates a title + verse slides.
      2. Search slide libraries for pre-made slides.
      3. Fallback: a generated title-only slide.
    """
    ref = item.get("title", "")

    # Try fetching from 和合本 online
    verses = fetch_verses(ref)
    if verses:
        slides = [{"type": "scripture_title", "item": item}]
        for group in group_verses_for_slides(verses):
            slides.append({"type": "scripture_verses", "ref": ref, "verses": group})
        return slides

    # Search libraries
    for lib in libraries:
        indices = _find_scripture_slides_in_library(lib, ref)
        if indices and len(indices) > 1:
            return [{"type": "copy_external", "prs": lib, "index": i} for i in indices]

    return [{"type": "scripture_title", "item": item}]


def _get_sermon_slides(agenda):
    """Build sermon outline slides from the parsed agenda."""
    outline = agenda.get("sermon_outline", {})
    title = outline.get("title", "")
    scripture = outline.get("scripture", "")
    main_points = outline.get("main_points", [])
    preacher = ""

    # Get preacher from worship order
    for item in agenda.get("worship_order", []):
        if item.get("type") == "sermon":
            preacher = item.get("presenter", "")
            break

    slides = []

    # Title slide
    header = f"{title}   {preacher}" if preacher else title
    slides.append({
        "type": "sermon_title",
        "header": header,
        "scripture": f"《{scripture}》" if scripture else "",
    })

    # One slide per main point; overflow into continuation slides if needed.
    for mp in main_points:
        heading = mp.get("heading", "")
        points = mp.get("points", [])
        groups = _group_sermon_points(heading, points)
        for i, group in enumerate(groups):
            slides.append({
                "type": "sermon_point",
                "header": header,
                "heading": heading,
                "points": group,
                "continuation": i > 0,
            })

    return slides


def _group_sermon_points(heading, points, chars_per_line=22, max_lines=10):
    """
    Split `points` into groups that fit within a single slide.

    Budget: max_lines total visual lines per slide. Each slide spends
      1 line on '今日信息' header + ceil(len(heading)/chars_per_line) on heading.
    Remaining budget is filled with points (each point costs its wrapped-line
    count). Overflow starts a new slide.
    """
    heading_lines = max(1, -(-len(heading) // chars_per_line)) if heading else 0
    reserved = 1 + heading_lines
    budget = max(1, max_lines - reserved)

    groups = []
    current = []
    current_lines = 0
    for pt in points:
        pt_lines = max(1, -(-len(pt) // chars_per_line))
        if current and current_lines + pt_lines > budget:
            groups.append(current)
            current = [pt]
            current_lines = pt_lines
        else:
            current.append(pt)
            current_lines += pt_lines
    if current:
        groups.append(current)
    return groups or [[]]


def _get_announcement_slides(agenda):
    """Build announcement slides from parsed agenda — 1 item per slide."""
    announcements = agenda.get("announcements", {})
    slides = []

    for section, items in announcements.items():
        section_label = f"報告： {section}"
        for item in items:
            slides.append({
                "type": "announcement",
                "section": section_label,
                "items": [item],
            })

    return slides


# ── File matching ─────────────────────────────────────────────────────────────

def _match_file(number, title, input_files, extensions):
    """Find an input file matching a hymn number or title."""
    for fname, fpath in input_files.items():
        ext = os.path.splitext(fname)[1].lower()
        if ext not in extensions:
            continue
        base = os.path.splitext(fname)[0]

        # Match by number prefix (e.g. "069-", "30-", "464-")
        if number is not None:
            patterns = [
                rf"^0*{number}[-_\s]",
                rf"^0*{number}$",
            ]
            for pat in patterns:
                if re.match(pat, base, re.IGNORECASE):
                    return fpath

        # Match by title substring
        if title:
            clean_title = re.sub(r"\s+", "", title)
            clean_base = re.sub(r"\s+", "", base)
            if clean_title in clean_base or clean_base in clean_title:
                return fpath

    return None


# ── Template slide finders ────────────────────────────────────────────────────

def _find_hymn_slides_in_library(prs, number, title):
    """Find hymn slides in a presentation by number prefix or title."""
    seen = set()
    for i, slide in enumerate(prs.slides):
        text = _slide_text(slide)
        if number and re.search(rf"{number}\s*[：:]", text):
            seen.add(i)
            continue
        if title and number is None:
            clean_title = re.sub(r"\s+", "", title)
            if clean_title in re.sub(r"\s+", "", text):
                seen.add(i)

    # Fall back to title search if number search found nothing
    if not seen and title:
        clean_title = re.sub(r"\s+", "", title)
        for i, slide in enumerate(prs.slides):
            if clean_title in re.sub(r"\s+", "", _slide_text(slide)):
                seen.add(i)

    return sorted(seen) or None


def _find_reading_slides_in_library(prs, number, title):
    """Find responsive reading slides in a presentation."""
    seen = set()
    for i, slide in enumerate(prs.slides):
        text = _slide_text(slide)
        if number and re.search(rf"啟應文\s*{number}", text):
            seen.add(i)
            continue
        if title and not number:
            if re.sub(r"\s+", "", title) in re.sub(r"\s+", "", text):
                seen.add(i)
    return sorted(seen) or None


def _find_anthem_slides_in_library(prs, title):
    """Find anthem slides in a presentation by title text."""
    seen = set()
    clean = re.sub(r"\s+", "", title)
    for i, slide in enumerate(prs.slides):
        if clean in re.sub(r"\s+", "", _slide_text(slide)):
            seen.add(i)
    return sorted(seen) or None


def _find_scripture_slides_in_library(prs, ref):
    """Find scripture slides by book name / reference."""
    seen = set()
    # Use the book name (first part before chapter:verse) as search key
    book = re.split(r"\s*\d", ref)[0].strip()
    clean_book = re.sub(r"\s+", "", book)
    if len(clean_book) < 2:
        return None
    for i, slide in enumerate(prs.slides):
        clean_text = re.sub(r"\s+", "", _slide_text(slide))
        if clean_book in clean_text:
            seen.add(i)
    return sorted(seen) or None


def _slide_text(slide):
    texts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            texts.append(shape.text_frame.text)
    return " ".join(texts)


# ── DOCX parsing ──────────────────────────────────────────────────────────────

def _parse_docx_lyrics(docx_path):
    """Parse a DOCX lyrics file and return list of verse strings."""
    doc = Document(docx_path)
    verses = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            # Split by newline within paragraph
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            if lines:
                verses.append("\n".join(lines))
    return verses


# ── Slide addition ────────────────────────────────────────────────────────────

def _clear_slides(prs):
    """Remove all slides from a presentation."""
    sldIdLst = prs.slides._sldIdLst
    for i in range(len(prs.slides) - 1, -1, -1):
        sld_id = sldIdLst[i]
        rId = sld_id.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        prs.part.drop_rel(rId)
        sldIdLst.remove(sld_id)


def _add_slide(out_prs, template, spec, agenda):
    """Add a slide to out_prs based on the spec."""
    stype = spec["type"]

    if stype == "copy_template":
        _copy_template_slide(out_prs, template, spec["index"])

    elif stype == "copy_external":
        _copy_template_slide(out_prs, spec["prs"], spec["index"])

    elif stype == "blank":
        layout = _get_layout(out_prs, "Blank")
        out_prs.slides.add_slide(layout)

    elif stype == "hymn_placeholder":
        _add_placeholder_slide(out_prs, spec["label"])

    elif stype == "anthem_title":
        _add_anthem_title_slide(out_prs, spec["title"])

    elif stype == "anthem_lyrics":
        _add_lyrics_slide(out_prs, spec["title"], spec["lyrics"])

    elif stype == "scripture_title":
        item = spec["item"]
        ref = item.get("title", "")
        _add_scripture_title_slide(out_prs, ref)

    elif stype == "scripture_verses":
        _add_scripture_verse_slide(out_prs, spec["ref"], spec["verses"])

    elif stype == "sermon_title":
        _add_sermon_title_slide(out_prs, spec)

    elif stype == "sermon_point":
        _add_sermon_point_slide(out_prs, spec)

    elif stype == "announcement":
        _add_announcement_slide(out_prs, spec)


def _copy_template_slide(out_prs, src_prs, index):
    """Copy a slide from src_prs at index into out_prs."""
    if index >= len(src_prs.slides):
        return

    src_slide = src_prs.slides[index]

    # Use a matching layout if possible
    layout_name = src_slide.slide_layout.name
    layout = _get_layout(out_prs, layout_name)

    new_slide = out_prs.slides.add_slide(layout)

    # Replace shape tree content
    src_sp_tree = src_slide._element.find(qn("p:cSld")).find(qn("p:spTree"))
    new_sp_tree = new_slide._element.find(qn("p:cSld")).find(qn("p:spTree"))

    # Remove auto-generated placeholders
    for child in list(new_sp_tree):
        new_sp_tree.remove(child)

    # Copy shapes from source
    for child in src_sp_tree:
        new_sp_tree.append(copy.deepcopy(child))

    # Copy background
    src_cSld = src_slide._element.find(qn("p:cSld"))
    new_cSld = new_slide._element.find(qn("p:cSld"))
    src_bg = src_cSld.find(qn("p:bg"))
    if src_bg is not None:
        new_bg = new_cSld.find(qn("p:bg"))
        if new_bg is not None:
            new_cSld.remove(new_bg)
        new_cSld.insert(0, copy.deepcopy(src_bg))


def _get_layout(prs, name):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    return prs.slide_layouts[0]


# ── Dynamic slide creators ────────────────────────────────────────────────────

def _add_placeholder_slide(out_prs, label):
    """Add a simple placeholder slide for missing content."""
    layout = _get_layout(out_prs, "Blank")
    slide = out_prs.slides.add_slide(layout)
    tf = slide.shapes.add_textbox(
        Inches(1), Inches(2.5), Inches(11.33), Inches(2)
    ).text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = 2  # center
    run = p.add_run()
    run.text = f"[{label}]"
    run.font.size = Pt(32)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x00, 0xFF, 0xFF)


def _add_section_slide(out_prs, label):
    """Add a section header slide (cyan text, centered)."""
    layout = _get_layout(out_prs, "Blank")
    slide = out_prs.slides.add_slide(layout)
    tb = slide.shapes.add_textbox(
        Emu(3359150), Emu(2492375), Emu(6673200), Emu(1219200)
    )
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = 2  # center
    run = p.add_run()
    run.text = label
    run.font.size = Pt(40)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x00, 0xFF, 0xFF)


def _add_anthem_title_slide(out_prs, title):
    """Anthem title slide — '獻詩: <title>' centered in mid-slide."""
    layout = _get_layout(out_prs, "Blank")
    slide = out_prs.slides.add_slide(layout)

    style = get_style("anthem_title") or {}
    box = style.get("box", {})
    pos  = box.get("pos",  (Emu(981777), Emu(2786743)))
    size = box.get("size", (Emu(10501161), Emu(1015663)))

    tb = slide.shapes.add_textbox(pos[0], pos[1], size[0], size[1])
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    apply_paragraph_style(p, style)
    if p.alignment is None:
        p.alignment = 2  # default center
    run = p.add_run()
    run.text = f"獻詩: {title}"
    apply_run_style(run, style)
    if run.font.size is None:
        run.font.size = Pt(60)
        run.font.bold = True
        run.font.color.rgb = CYAN_SOFT


def _add_scripture_title_slide(out_prs, ref):
    """Scripture title slide — '經文 <reference>' with style applied."""
    layout = _get_layout(out_prs, "Blank")
    slide = out_prs.slides.add_slide(layout)

    style = get_style("scripture_title") or {}
    box = style.get("box", {})
    pos  = box.get("pos",  (Emu(154746), Emu(2086708)))
    size = box.get("size", (Emu(12037254), Emu(2000250)))

    tb = slide.shapes.add_textbox(pos[0], pos[1], size[0], size[1])
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    apply_paragraph_style(p, style)
    if p.alignment is None:
        p.alignment = 2

    # Reference only (no "經文" label)
    reference = style.get("reference", {})
    run = p.add_run()
    run.text = ref
    if reference:
        apply_run_style(run, reference)
    else:
        run.font.size = Pt(72); run.font.bold = True
        run.font.color.rgb = CYAN_SOFT


def _add_scripture_verse_slide(out_prs, ref, verses):
    """Scripture verse slide with reference top-bar and wrapped verse body."""
    layout = _get_layout(out_prs, "Blank")
    slide = out_prs.slides.add_slide(layout)

    # ── Top reference bar ──
    bar_style = get_style("scripture_verse_bar") or {}
    bar_box = bar_style.get("box", {})
    bpos  = bar_box.get("pos",  (Emu(0), Emu(0)))
    bsize = bar_box.get("size", (Emu(12192000), Emu(1561514)))

    tb_ref = slide.shapes.add_textbox(bpos[0], bpos[1], bsize[0], bsize[1])
    tf_ref = tb_ref.text_frame
    p_ref = tf_ref.paragraphs[0]
    apply_paragraph_style(p_ref, bar_style)
    if p_ref.alignment is None:
        p_ref.alignment = 2
    run_ref = p_ref.add_run()
    run_ref.text = ref
    if bar_style:
        apply_run_style(run_ref, bar_style)
    else:
        run_ref.font.size = Pt(40); run_ref.font.bold = True
        run_ref.font.color.rgb = CYAN_SOFT

    # ── Verse body ──
    body_style = get_style("scripture_verse_body") or {}
    body_box = body_style.get("box", {})
    dpos  = body_box.get("pos",  (Emu(304800), Emu(1600000)))
    dsize = body_box.get("size", (Emu(11582400), Emu(5100000)))
    chars_per_line = body_style.get("chars_per_line", 0)

    tb_body = slide.shapes.add_textbox(dpos[0], dpos[1], dsize[0], dsize[1])
    tf_body = tb_body.text_frame
    tf_body.word_wrap = True

    first = True
    for v in verses:
        # Dynamic indent: align continuation with content after verse number.
        # "1." → 2-space indent; "10." → 3-space indent.
        prefix = f"{v['verse']}."
        indent = " " * len(prefix)
        line = prefix + v['text']
        if chars_per_line:
            line = wrap_chinese_text(line, chars_per_line, indent=indent)
        p = tf_body.paragraphs[0] if first else tf_body.add_paragraph()
        first = False
        run = p.add_run()
        run.text = line
        if body_style:
            apply_run_style(run, body_style)
        else:
            run.font.size = Pt(32)
            run.font.color.rgb = WHITE


def _add_lyrics_slide(out_prs, title, lyrics):
    """Add a lyric/content slide using 詩歌 layout style."""
    layout = _get_layout(out_prs, "詩歌")
    slide = out_prs.slides.add_slide(layout)

    ph_list = list(slide.placeholders)
    lines = lyrics.split("\n")

    if len(ph_list) >= 2:
        ph_list[0].text = title
        tf = ph_list[1].text_frame
        tf.clear()
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
    else:
        tb = slide.shapes.add_textbox(Emu(0), Emu(126609), Emu(12192000), Emu(436100))
        tb.text_frame.paragraphs[0].text = title

        tb2 = slide.shapes.add_textbox(Emu(0), Emu(745435), Emu(12192000), Emu(6112565))
        tf = tb2.text_frame
        tf.word_wrap = True
        tf.clear()
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line


def _add_sermon_title_slide(out_prs, spec):
    """Sermon title slide: '今日信息' header + sermon title + preacher using layout placeholders."""
    layout = _get_layout(out_prs, "詩歌")
    slide = out_prs.slides.add_slide(layout)
    ph_list = list(slide.placeholders)

    # Split header back into title + preacher
    header = spec.get("header", "")
    parts = header.split("   ", 1) if "   " in header else [header, ""]
    sermon_title = parts[0].strip()
    preacher = parts[1].strip() if len(parts) > 1 else ""

    hdr_style = get_style("sermon_title_header") or {}
    body_style = get_style("sermon_title_body") or {}

    if len(ph_list) >= 2:
        # Title placeholder (idx=0): "今日信息" with cyan color
        tf0 = ph_list[0].text_frame
        tf0.clear()
        p0 = tf0.paragraphs[0]
        run0 = p0.add_run()
        run0.text = "今日信息"
        apply_run_style(run0, hdr_style)

        # Body placeholder (idx=1): sermon title + preacher
        tf = ph_list[1].text_frame
        tf.clear()
        p_title = tf.paragraphs[0]

        # Sermon title
        run_t = p_title.add_run()
        run_t.text = sermon_title
        apply_run_style(run_t, body_style.get("title", {}))

        # Add preacher if present
        if preacher:
            run_sp = p_title.add_run()
            run_sp.text = "             "  # visual spacer
            apply_run_style(run_sp, body_style.get("spacer", {}))
            run_p = p_title.add_run()
            run_p.text = preacher
            apply_run_style(run_p, body_style.get("preacher", {}))


def _add_sermon_point_slide(out_prs, spec):
    """Add sermon point slide using layout placeholders."""
    layout = _get_layout(out_prs, "詩歌")
    slide = out_prs.slides.add_slide(layout)

    heading = spec.get("heading", "")
    points = spec.get("points", [])
    if spec.get("continuation") and heading and "(續)" not in heading:
        heading = f"{heading} (續)"

    # Find body placeholder (idx=1) and title placeholder (idx=0)
    body_ph = None
    title_ph = None
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 1:
            body_ph = ph
        elif ph.placeholder_format.idx == 0:
            title_ph = ph

    # Remove title placeholder to have single-shape slide
    if title_ph is not None:
        sp = title_ph._element
        sp.getparent().remove(sp)

    hdr_style = get_style("sermon_point_header") or {}
    body_style = get_style("sermon_point_body") or {}

    if body_ph is not None:
        tf = body_ph.text_frame
        tf.clear()

        # Header: "今日信息"
        p0 = tf.paragraphs[0]
        run0 = p0.add_run()
        run0.text = "今日信息"
        apply_run_style(run0, hdr_style)

        # Heading
        p_heading = tf.add_paragraph()
        apply_paragraph_style(p_heading, body_style)
        run_heading = p_heading.add_run()
        run_heading.text = heading

        # Points
        for pt in points:
            p_pt = tf.add_paragraph()
            apply_paragraph_style(p_pt, body_style)
            run_pt = p_pt.add_run()
            run_pt.text = pt
    else:
        # Fallback: manual textbox
        tb = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(12192000), Emu(6858000))
        tf = tb.text_frame
        p0 = tf.paragraphs[0]
        run0 = p0.add_run()
        run0.text = "今日信息"
        apply_run_style(run0, hdr_style)
        for text in [heading] + points:
            p = tf.add_paragraph()
            apply_paragraph_style(p, body_style)
            run = p.add_run()
            run.text = text


def _add_announcement_slide(out_prs, spec):
    """Add an announcement slide."""
    layout = _get_layout(out_prs, "Blank")
    # Try 詩歌 layout for consistency
    for l in out_prs.slide_layouts:
        if l.name == "1_詩歌" or l.name == "詩歌":
            layout = l
            break

    slide = out_prs.slides.add_slide(layout)
    ph_list = list(slide.placeholders)
    section = spec.get("section", "報告")
    items = spec.get("items", [])
    content = "\n".join(items)

    if len(ph_list) >= 2:
        ph_list[0].text = section
        tf = ph_list[1].text_frame
        tf.clear()
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = item
            p.alignment = 1  # left
    else:
        tb = slide.shapes.add_textbox(
            Emu(998220), Emu(-152400), Emu(9403080), Emu(1143000)
        )
        tb.text_frame.paragraphs[0].text = section

        tb2 = slide.shapes.add_textbox(
            Emu(0), Emu(990600), Emu(12355033), Emu(5360964)
        )
        tf = tb2.text_frame
        tf.word_wrap = True
        tf.clear()
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = 1  # left
            run = p.add_run()
            run.text = item
            run.font.size = Pt(18)
            run.font.bold = True
