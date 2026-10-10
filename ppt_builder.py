"""
Main entry point: assemble the Sunday worship PPTX from agenda + inputs.
"""

import logging
import os
import shutil

from pptx import Presentation

from autofit import shrink_overflowing_text, strip_trailing_empty_paragraphs
from slide_copier import clear_slides, copy_slide, get_layout, restyle_responsive_slide
from slide_generators import (
    add_announcement_slide, add_anthem_title_slide, add_lyrics_slide,
    add_placeholder_slide, add_scripture_title_slide, add_scripture_verse_slide,
    add_sermon_point_slide, add_sermon_title_slide,
)
from slide_planner import plan_slides, unused_inputs

logger = logging.getLogger(__name__)


def build_pptx(fixed_path, agenda, input_files, output_path, library_paths=None,
               intro_path=None, bible_page=None, overrides=None):
    """
    Build the worship PPTX.

    Args:
        fixed_path: Deck holding the recurring service slides (template/fixed.pptx)
        agenda: Parsed agenda dict from pdf_parser
        input_files: Dict of filename → filepath for input files
        output_path: Where to save the result
        library_paths: Additional PPTX files to search for hymn/reading slides
        intro_path: Fixed-intro PPTX; the output deck starts as a copy of it,
            so it also supplies the slide masters and layouts
        bible_page: Page number for the 經文 title slide, when the operator
            supplied one; otherwise looked up from bible_pages
        overrides: item_id → input filename, from the UI's file pickers

    Returns:
        dict with `path` and `unused_inputs` (uploaded files nothing consumed)
    """
    fixed_deck = Presentation(fixed_path)
    use_intro = bool(intro_path and os.path.exists(intro_path))

    libraries = [fixed_deck]
    for p in (library_paths or []):
        if p != fixed_path and os.path.exists(p):
            try:
                libraries.append(Presentation(p))
            except Exception:
                logger.exception("Failed to load library PPTX %s; skipping", p)

    slides_to_add = plan_slides(
        fixed_deck, libraries, agenda, input_files,
        skip_intro=use_intro, bible_page=bible_page, overrides=overrides,
    )

    base_path = intro_path if use_intro else fixed_path
    shutil.copy2(base_path, output_path)
    out_prs = Presentation(output_path)

    if not use_intro:
        clear_slides(out_prs)

    for spec in slides_to_add:
        _add_slide(out_prs, fixed_deck, spec, agenda)

    out_prs.save(output_path)
    return {
        "path": output_path,
        "unused_inputs": unused_inputs(slides_to_add, input_files),
    }


def _add_slide(out_prs, fixed_deck, spec, agenda):
    """Dispatch slide creation based on spec type."""
    stype = spec["type"]

    if stype == "copy_template":
        prs = spec.get("prs", fixed_deck)
        index = spec.get("index")
        if index is not None:
            # Copied verbatim: the fixed deck is curated by hand, so its
            # slides are reproduced exactly rather than re-measured.
            copy_slide(out_prs, prs, index)

    elif stype == "copy_external":
        prs = spec.get("prs")
        index = spec.get("index")
        if index is not None and prs:
            new_slide = copy_slide(out_prs, prs, index)
            if new_slide is not None:
                if spec.get("restyle") == "responsive":
                    restyle_responsive_slide(new_slide, out_prs.slide_width)
                strip_trailing_empty_paragraphs(new_slide)
                shrink_overflowing_text(new_slide)

    elif stype == "blank":
        out_prs.slides.add_slide(get_layout(out_prs, "Blank"))

    elif stype == "hymn_placeholder":
        add_placeholder_slide(out_prs, spec.get("label", ""))

    elif stype == "anthem_title":
        add_anthem_title_slide(out_prs, spec.get("title", ""))

    elif stype == "anthem_lyrics":
        add_lyrics_slide(out_prs, spec.get("title", ""), spec.get("lines", []))

    elif stype == "scripture_title":
        add_scripture_title_slide(out_prs, spec.get("item", {}), spec.get("bible_page"))

    elif stype == "scripture_verses":
        add_scripture_verse_slide(out_prs, spec.get("ref", ""), spec.get("verses", []))

    elif stype == "sermon_title":
        add_sermon_title_slide(out_prs, spec)

    elif stype == "sermon_point":
        add_sermon_point_slide(out_prs, spec)

    elif stype == "announcement":
        add_announcement_slide(out_prs, spec)

    else:
        logger.warning("Unknown slide spec type %r; skipping", stype)
