"""
Slide planning logic: translate agenda + input files → ordered list of slide specs.

Determines what slides should be created and in what order, before generation.
Specs that copy an input deck carry `source_file`, so `unused_inputs()` can
report an uploaded file that nothing consumed — a guest preacher's sermon deck
used to be dropped in silence.
"""

import logging
import os
import re

from docx import Document
from pptx import Presentation

import bible_pages
from bible_fetcher import fetch_verses
from service_config import SLIDE_STYLES
from slide_finder import find_slide
from text_layout import capacity, pack_lines, wrap_text

logger = logging.getLogger(__name__)

_UNSET = object()


def plan_slides(fixed_deck, libraries, agenda, input_files, skip_intro=False,
                bible_page=None, overrides=None):
    """
    Build the ordered list of slide specs for the service.

    Args:
        fixed_deck: Presentation holding the recurring service slides
        libraries: List of library Presentations
        agenda: Parsed agenda dict with worship_order, sermon_outline, announcements
        input_files: Dict of filename → filepath
        skip_intro: If True, don't add intro slides (they come from intro PPTX)
        bible_page: Page number for the 經文 title. None falls back to
            bible_pages.lookup(); only a value the operator typed wins over it.
        overrides: Optional dict of item_id → filename (from input_files) to
            force a specific file. Use empty string to force "no file"
            (library/placeholder fallback). Ids match plan_match_items().

    Returns:
        List of slide spec dicts describing what to create
    """
    order = agenda.get("worship_order", [])
    slides = []
    overrides = overrides or {}
    hymn_idx = reading_idx = anthem_idx = 0

    # === Fixed pre-worship slides ===
    if not skip_intro:
        call_to_worship_idx = find_slide(fixed_deck, "call_to_worship")
        if call_to_worship_idx is None:
            call_to_worship_idx = 6  # fallback
        for i in range(call_to_worship_idx):
            slides.append({"type": "copy_template", "prs": fixed_deck, "index": i})

    def copy_fixed(*keys):
        for key in keys:
            idx = find_slide(fixed_deck, key)
            if idx is not None:
                slides.append({"type": "copy_template", "prs": fixed_deck, "index": idx})

    # === Worship order items ===
    for pos, item in enumerate(order):
        itype = item.get("type")
        next_type = order[pos + 1].get("type") if pos + 1 < len(order) else None

        if itype == "call_to_worship":
            copy_fixed("call_to_worship")
            slides.append({"type": "blank"})

        elif itype == "hymn":
            slides.extend(_get_hymn_slides(
                libraries, item, input_files,
                override=overrides.get(f"hymn-{hymn_idx}")))
            hymn_idx += 1
            slides.append({"type": "blank"})

        elif itype == "prayer":
            copy_fixed("prayer", "lords_prayer_1", "lords_prayer_2")

        elif itype == "creed":
            copy_fixed("creed_1", "creed_2")
            slides.append({"type": "blank"})

        elif itype == "responsive":
            slides.extend(_get_reading_slides(
                libraries, item, input_files,
                override=overrides.get(f"reading-{reading_idx}")))
            reading_idx += 1
            slides.append({"type": "blank"})

        elif itype == "anthem":
            slides.extend(_get_anthem_slides(
                libraries, item, input_files,
                override=overrides.get(f"anthem-{anthem_idx}")))
            anthem_idx += 1
            slides.append({"type": "blank"})

        elif itype == "scripture":
            slides.extend(_get_scripture_slides(item, bible_page))
            # No blank between scripture and sermon

        elif itype == "sermon":
            slides.extend(_get_sermon_slides(
                agenda, input_files, override=overrides.get("sermon-0")))
            slides.append({"type": "blank"})

        elif itype == "offering":
            copy_fixed("offering_1", "offering_2")
            # On communion Sundays the offering runs straight into 聖餐 with no
            # divider, the way the reference decks do it.
            if next_type != "communion":
                slides.append({"type": "blank"})

        elif itype == "communion":
            _warn_on_fixed_mismatch(fixed_deck, item)
            copy_fixed("communion_1", "communion_2", "communion_3")
            slides.append({"type": "blank"})

        elif itype == "announcements":
            ann_slides = _get_announcement_slides(agenda)
            if ann_slides:
                copy_fixed("announce_title")
                slides.extend(ann_slides)
                slides.append({"type": "blank"})

        elif itype in ("doxology", "benediction"):
            pass  # handled in the closing section below

    # === Fixed closing slides ===
    copy_fixed("doxology", "benediction", "quiet", "website")

    return slides


def _warn_on_fixed_mismatch(fixed_deck, item):
    """Warn when the agenda's communion hymn isn't the one in the fixed deck."""
    num = item.get("number")
    if not num:
        return
    idx = find_slide(fixed_deck, "communion_2")
    if idx is None:
        return
    text = _slide_text(fixed_deck.slides[idx])
    if num not in text:
        logger.warning(
            "Agenda asks for 聖餐 hymn %s but template/fixed.pptx has %r; "
            "using the template slide.", num, text[:40])


def unused_inputs(slides_spec, input_files):
    """
    Input files that no planned slide consumed.

    An unmatched PPTX is usually a guest preacher's own deck, which has no
    agenda line to match against; surfacing it lets the operator attach it
    rather than discover the gap during the service.
    """
    used = {os.path.basename(s["source_file"])
            for s in slides_spec if s.get("source_file")}
    return sorted(name for name in input_files if name not in used)


# ── External deck resolution ─────────────────────────────────────────────────

def _get_external_slides(libraries, item, input_files, find_fn, override=None,
                         restyle=None, placeholder_label=""):
    """
    Shared resolution flow for hymn/reading items: try a matched input PPTX,
    then search libraries via find_fn, then fall back to a placeholder.
    """
    num = item.get("number")
    title = item.get("title", "")
    extra = {"restyle": restyle} if restyle else {}

    matched = _resolve_input(override, input_files, [".pptx"])
    if matched is _UNSET:
        matched = _match_file(num, title, input_files, [".pptx"])
    if matched:
        try:
            src = Presentation(matched)
            if src.slides:
                return [{"type": "copy_external", "prs": src, "index": i,
                         "source_file": matched, **extra}
                        for i in range(len(src.slides))]
        except Exception:
            logger.exception("Failed to load matched input PPTX %s; falling back "
                             "to library search", matched)

    for lib in libraries:
        indices = find_fn(lib, num, title)
        if indices:
            return [{"type": "copy_external", "prs": lib, "index": i, **extra}
                    for i in indices]

    return [{"type": "hymn_placeholder", "label": placeholder_label}]


def _get_hymn_slides(libraries, item, input_files, is_doxology=False, override=None):
    """Get slides for a hymn item, from input file or slide library."""
    num = item.get("number")
    title = item.get("title", "")
    label = f"{'頌榮' if is_doxology else '聖詩'} {num}: {title}" if num else title
    return _get_external_slides(libraries, item, input_files,
                                _find_hymn_slides_in_library,
                                override=override, placeholder_label=label)


def _get_reading_slides(libraries, item, input_files, override=None):
    """Get slides for the responsive reading."""
    num = item.get("number")
    title = item.get("title", "")
    label = f"啟應文 {num}: {title}"
    return _get_external_slides(libraries, item, input_files,
                                _find_reading_slides_in_library,
                                override=override, restyle="responsive",
                                placeholder_label=label)


def _get_anthem_slides(libraries, item, input_files, override=None):
    """Get anthem title + lyrics slides from DOCX or library."""
    title = item.get("title", "")
    slides = [{"type": "anthem_title", "title": title}]

    matched = _resolve_input(override, input_files, [".docx", ".doc"])
    if matched is _UNSET:
        matched = _match_file(None, title, input_files, [".docx", ".doc"])
    if matched:
        try:
            lyrics = _parse_docx_lyrics(matched)
            clean_title = re.sub(r"\s+", "", title)
            verses = []
            title_seen = False
            for v in lyrics:
                if not title_seen and re.sub(r"\s+", "", v) == clean_title:
                    title_seen = True
                    continue
                verses.append(v)
            if verses:
                for chunk in _group_anthem_verses(verses):
                    slides.append({"type": "anthem_lyrics", "title": title,
                                   "lines": chunk, "source_file": matched})
                return slides
        except Exception:
            logger.exception("Failed to parse anthem lyrics from %s; falling back "
                             "to library search", matched)

    for lib in libraries:
        indices = _find_anthem_slides_in_library(lib, title)
        if indices:
            content_indices = [i for i in indices if i != indices[0]] or indices
            return [{"type": "anthem_title", "title": title}] + \
                   [{"type": "copy_external", "prs": lib, "index": i}
                    for i in content_indices]

    return slides


def _group_anthem_verses(verses):
    """
    Pack lyric lines into slide-sized chunks using the measured 詩歌 body
    capacity (6 lines at 54pt, not the 7 a 1.2 line-height would suggest —
    that off-by-one is what pushed the last line of each slide off the bottom).

    Returns a list of line-lists.
    """
    style = SLIDE_STYLES["lyrics"]
    cols, max_lines = capacity(style["box"], style["size_pt"])

    all_lines = []
    for v in verses:
        for line in v.split("\n"):
            if line.strip():
                all_lines.append(line)

    # A lyric line longer than the box still has to wrap; count it honestly.
    expanded = []
    for line in all_lines:
        expanded.extend(wrap_text(line, cols))
    return pack_lines(expanded, max_lines)


# ── Scripture ────────────────────────────────────────────────────────────────

def _get_scripture_slides(item, bible_page=None):
    """
    Build the 經文 title slide plus verse slides.

    The library fallback this used to have searched on book name alone, so a
    reference like "詩篇 90" matched any slide containing 詩篇 and pulled in a
    完全 unrelated 啟應文 reading. There is no safe way to recover verse text
    from the library, so when the online lookup fails this now emits the title
    slide only and says so in the log.
    """
    ref = item.get("title", "")
    page = bible_page if bible_page not in (None, "") else bible_pages.lookup(ref)
    slides = [{"type": "scripture_title", "item": item, "bible_page": page}]

    verses = fetch_verses(ref)
    if not verses:
        logger.warning("No verse text for %r; emitting the 經文 title slide only. "
                       "Add the passage manually or retry with network access.", ref)
        return slides

    style = SLIDE_STYLES["scripture_verse_body"]
    cols, max_lines = capacity(style["size"], style["size_pt"])

    for group in _group_verses(verses, cols, max_lines):
        slides.append({"type": "scripture_verses", "ref": ref, "verses": group})
    return slides


def _group_verses(verses, cols, max_lines):
    """
    Group verses onto slides by their *wrapped* line count.

    Each verse is laid out as "25.<text>" with continuation lines indented to
    clear the number, so the line count is computed from the formatted text,
    not the raw verse. Each verse after the first also costs ~0.2 of a line in
    paragraph spacing (the master sets spcBef to 20%).
    """
    groups = []
    current = []
    current_cost = 0.0

    for v in verses:
        lines = _verse_lines(v, cols)
        cost = len(lines) + (0.2 if current else 0.0)
        if current and current_cost + cost > max_lines:
            groups.append(current)
            current = []
            current_cost = 0.0
            cost = len(lines)
        current.append(dict(v, lines=lines))
        current_cost += cost

    if current:
        groups.append(current)
    return groups


def _verse_lines(verse, cols):
    """Wrapped lines for one verse, with continuation lines indented."""
    prefix = f"{verse['verse']}."
    indent = " " * (len(prefix) + 1)
    return wrap_text(prefix + verse["text"], cols, indent=indent)


# ── Sermon ───────────────────────────────────────────────────────────────────

def _get_sermon_slides(agenda, input_files=None, override=None):
    """
    Build sermon slides.

    Normally the outline from the agenda's 講台綱要 page. When the operator
    attaches a preacher's own deck (overrides["sermon-0"]), its slides are
    copied in after the title slide instead — some guest preachers supply a
    full deck that has no agenda line to match on.
    """
    outline = agenda.get("sermon_outline", {})
    title = outline.get("title", "")
    scripture = outline.get("scripture", "")
    main_points = outline.get("main_points", [])

    preacher = ""
    for item in agenda.get("worship_order", []):
        if item.get("type") == "sermon":
            preacher = item.get("presenter", "")
            break

    slides = [{
        "type": "sermon_title",
        "title": title,
        "preacher": preacher,
        "scripture": f"《{scripture}》" if scripture else "",
    }]

    attached = _resolve_input(override, input_files or {}, [".pptx"])
    if attached not in (_UNSET, None):
        try:
            src = Presentation(attached)
            if src.slides:
                slides.extend({"type": "copy_external", "prs": src, "index": i,
                               "source_file": attached}
                              for i in range(len(src.slides)))
                return slides
        except Exception:
            logger.exception("Failed to load sermon deck %s; falling back to the "
                             "agenda outline", attached)

    style = SLIDE_STYLES["sermon_point"]
    cols, max_lines = capacity(style["box"], style["size_pt"])

    for mp in main_points:
        heading = mp.get("heading", "")
        # Tighten numbering: "1. text" → "1.text" (matches reference deck)
        points = [re.sub(r"^(\d+)\.\s+", r"\1.", pt) for pt in mp.get("points", [])]
        for i, group in enumerate(_group_sermon_points(heading, points, cols, max_lines)):
            slides.append({
                "type": "sermon_point",
                "heading": heading,
                "points": group,
                "continuation": i > 0,
            })

    return slides


def _group_sermon_points(heading, points, cols, max_lines):
    """
    Split a main point's sub-points into slides that actually fit.

    Budget: the whole box, minus the 今日信息 header line, minus the heading
    (which repeats on continuation slides). Everything is measured in wrapped
    lines at the real box width.
    """
    header_lines = 1
    heading_lines = len(wrap_text(heading, cols)) if heading else 0
    budget = max(1, max_lines - header_lines - heading_lines)

    groups = []
    current = []
    current_lines = 0
    for pt in points:
        n = len(wrap_text(pt, cols))
        if current and current_lines + n > budget:
            groups.append(current)
            current = []
            current_lines = 0
        current.append(pt)
        current_lines += n
    if current:
        groups.append(current)
    return groups or [[]]


# ── Announcements ────────────────────────────────────────────────────────────

def _get_announcement_slides(agenda):
    """
    One slide per announcement item.

    Long items used to be chopped at a fixed character count, which produced
    tail slides holding a fragment ("報告：台語部 -8656。") and cut through
    English names mid-word. Instead the item stays whole and the generator
    steps the font size down until it fits, which is what the reference decks
    do by hand.
    """
    slides = []
    for section, items in agenda.get("announcements", {}).items():
        section_label = f"報告： {section}"
        for item in items:
            slides.append({
                "type": "announcement",
                "section": section_label,
                "text": item,
            })
    return slides


# ── Input file matching ──────────────────────────────────────────────────────

def _resolve_input(override, input_files, extensions):
    """
    Resolve a user-supplied override into a file path.

    Returns:
        - _UNSET if no override (caller should run auto-match)
        - None if override is "" or unknown (force "no file" fallback)
        - filepath str otherwise
    """
    if override is None:
        return _UNSET
    if override == "":
        return None
    if override in input_files:
        ext = os.path.splitext(override)[1].lower()
        if ext in extensions:
            return input_files[override]
    return None


def _match_file(number, title, input_files, extensions):
    """
    Find an input file matching a hymn number or title.

    `number` is a string and may carry a letter suffix ("254A"), so it is
    regex-escaped rather than interpolated raw.
    """
    num = str(number) if number not in (None, "") else None
    num_pat = re.escape(num) if num else None

    # First pass: number prefix, or whole-string title containment.
    for fname, fpath in input_files.items():
        ext = os.path.splitext(fname)[1].lower()
        if ext not in extensions:
            continue
        base = os.path.splitext(fname)[0].strip()

        if num_pat:
            for pat in (rf"^0*{num_pat}[-_\s]", rf"^0*{num_pat}$"):
                if re.match(pat, base, re.IGNORECASE):
                    return fpath

        if title:
            clean_title = re.sub(r"\s+", "", title)
            clean_base = re.sub(r"\s+", "", base)
            if clean_title and (clean_title in clean_base or clean_base in clean_title):
                return fpath

    # Second pass: number anywhere with a non-digit boundary, or fuzzy title.
    best = None
    best_score = 0.0
    for fname, fpath in input_files.items():
        ext = os.path.splitext(fname)[1].lower()
        if ext not in extensions:
            continue
        base = os.path.splitext(fname)[0].strip()

        if num_pat and re.search(rf"(?<!\d)0*{num_pat}(?![0-9A-Za-z])", base,
                                 re.IGNORECASE):
            return fpath

        if title:
            clean_title = re.sub(r"[\s\-_().\[\]【】《》，,]", "", title)
            clean_base = re.sub(r"[\s\-_().\[\]【】《》，,]", "", base)
            if clean_title:
                shared = sum(1 for ch in clean_title if ch in clean_base)
                score = shared / len(clean_title)
                if score >= 0.6 and score > best_score:
                    best = fpath
                    best_score = score

    return best


def plan_match_items(agenda, input_files):
    """
    Build the list of agenda items that need an input file, with auto-match
    results and candidate files for each.

    Item ids are stable for a given agenda order: hymn-0, hymn-1, reading-0,
    anthem-0, sermon-0. They line up with the overrides keys consumed by
    plan_slides().
    """
    pptx_candidates = sorted(
        fname for fname in input_files
        if os.path.splitext(fname)[1].lower() == ".pptx"
    )
    docx_candidates = sorted(
        fname for fname in input_files
        if os.path.splitext(fname)[1].lower() in (".docx", ".doc")
    )

    items = []
    hymn_idx = reading_idx = anthem_idx = 0
    for entry in agenda.get("worship_order", []):
        itype = entry.get("type")
        num = entry.get("number")
        title = entry.get("title", "")

        if itype == "hymn":
            matched = _match_file(num, title, input_files, [".pptx"])
            items.append({
                "id": f"hymn-{hymn_idx}",
                "kind": "hymn",
                "label": f"聖詩 {num}: {title}" if num else f"聖詩: {title}",
                "matched_file": os.path.basename(matched) if matched else None,
                "candidates": pptx_candidates,
            })
            hymn_idx += 1
        elif itype == "responsive":
            matched = _match_file(num, title, input_files, [".pptx"])
            items.append({
                "id": f"reading-{reading_idx}",
                "kind": "responsive",
                "label": f"啟應文 {num}: {title}" if num else f"啟應文: {title}",
                "matched_file": os.path.basename(matched) if matched else None,
                "candidates": pptx_candidates,
            })
            reading_idx += 1
        elif itype == "anthem":
            matched = _match_file(None, title, input_files, [".docx", ".doc"])
            items.append({
                "id": f"anthem-{anthem_idx}",
                "kind": "anthem",
                "label": f"獻詩: {title}",
                "matched_file": os.path.basename(matched) if matched else None,
                "candidates": docx_candidates,
            })
            anthem_idx += 1
        elif itype == "sermon":
            # Never auto-matched: a preacher's deck is named after the passage,
            # not the sermon title. Offered so it can be attached by hand.
            items.append({
                "id": "sermon-0",
                "kind": "sermon",
                "label": f"證道: {title}（可選：講道投影片）",
                "matched_file": None,
                "candidates": pptx_candidates,
            })
    return items


# ── Library search ───────────────────────────────────────────────────────────

def _find_hymn_slides_in_library(prs, number, title):
    """Find hymn slides in a presentation by number prefix or title."""
    num_pat = re.escape(str(number)) if number not in (None, "") else None
    seen = set()
    for i, slide in enumerate(prs.slides):
        text = _slide_text(slide)
        if num_pat and re.search(rf"(?<!\d){num_pat}\s*[：:]", text):
            seen.add(i)
            continue
        if title and not num_pat:
            if re.sub(r"\s+", "", title) in re.sub(r"\s+", "", text):
                seen.add(i)

    if not seen and title:
        clean_title = re.sub(r"\s+", "", title)
        for i, slide in enumerate(prs.slides):
            if clean_title in re.sub(r"\s+", "", _slide_text(slide)):
                seen.add(i)

    return sorted(seen) or None


def _find_reading_slides_in_library(prs, number, title):
    """Find responsive reading slides in a presentation."""
    num_pat = re.escape(str(number)) if number not in (None, "") else None
    seen = set()
    for i, slide in enumerate(prs.slides):
        text = _slide_text(slide)
        if num_pat and re.search(rf"啟應文\s*{num_pat}(?!\d)", text):
            seen.add(i)
            continue
        if title and not num_pat:
            if re.sub(r"\s+", "", title) in re.sub(r"\s+", "", text):
                seen.add(i)
    return sorted(seen) or None


def _find_anthem_slides_in_library(prs, title):
    """Find anthem slides in a presentation by title text."""
    seen = set()
    clean = re.sub(r"\s+", "", title)
    if not clean:
        return None
    for i, slide in enumerate(prs.slides):
        if clean in re.sub(r"\s+", "", _slide_text(slide)):
            seen.add(i)
    return sorted(seen) or None


def _slide_text(slide):
    """Extract all text from a slide."""
    return " ".join(shape.text_frame.text for shape in slide.shapes
                    if shape.has_text_frame)


def _parse_docx_lyrics(docx_path):
    """Parse a DOCX lyrics file and return list of verse strings."""
    doc = Document(docx_path)
    verses = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            if lines:
                verses.append("\n".join(lines))
    return verses
