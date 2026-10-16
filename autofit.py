"""
Shrink text that would overflow its box.

These decks don't enable PowerPoint's shrink-on-overflow autofit, so text that
is too tall for its placeholder simply runs off the bottom of the slide and is
never seen. That happens a lot in the source material: hymn decks routinely
carry more 54pt lines than the box holds, and rendering the hand-finished
reference decks shows text falling as much as 250pt — three and a half lines —
below the slide edge.

Copied slides therefore get measured and, where needed, scaled down. The text
itself is never touched, only the run sizes, so nothing is lost or re-wrapped
differently; it just becomes visible.

Height, not line count
----------------------
Measuring in "lines" only works if every line is the same size, and in these
source decks they are not — a hymn body often mixes 54pt lyric lines with 32pt
romanisation. So each paragraph is measured at its own effective size and the
heights are summed, which is what the renderer actually does.

Only slides copied from a week's *input* decks are measured. The curated fixed
deck is left exactly as the operator maintains it: its slides were being
shrunk — 宣召 from 66pt to 55pt, 信仰告白 from 50pt to 33pt — because the
height model over-predicted, and even with that fixed, second-guessing a deck
the operator controls is not this pass's job.
"""

import logging

from pptx.oxml.ns import qn
from pptx.util import Pt

from slide_copier import placeholder_idx
from text_layout import (DEFAULT_INSET, LINE_PITCH, PARA_SPACING,
                         WIDTH_MARGIN_EM, estimate_lines)

logger = logging.getLogger(__name__)

# Don't shrink past this fraction of the original size — beyond it the slide is
# unreadable from the back of the room and a human should re-split the content.
MIN_SCALE = 0.62
# Steps to try, so sizes stay on tidy values rather than arbitrary fractions.
SCALE_STEPS = (0.92, 0.84, 0.78, 0.72, 0.66, MIN_SCALE)
# A body placeholder that inherits the 詩歌 master's body style renders its
# first line about a full line below the top of the box — the master's spcBef
# plus a full leading. Measured: a 6-line 54pt body whose text is 438pt tall
# starts 59pt (1.09em) down a 455pt box and so runs 30pt off the slide, even
# though the text alone fits. A plain text box gets no such offset: 信仰告白
# spans 5.6-537.1pt in a 0-540pt box. Budget for it, or those slides escape
# shrinking and lose their last line.
TOP_LEADING = 1.0

# Copied shapes often can't be matched to a destination placeholder, so they
# inherit the master's body style — including its hanging-indent bullet
# (marL 342900 EMU, exactly 0.5em at 54pt) and the bullet glyph itself. None of
# that shows in the box geometry, so allow for it on top of the usual margin.
# Under-allowing here means a line is thought to fit when it actually wraps,
# and the slide escapes shrinking and overflows by about half a line.
COPIED_WIDTH_MARGIN_EM = WIDTH_MARGIN_EM + 1.0


def strip_trailing_empty_paragraphs(slide):
    """
    Drop empty paragraphs at the end of each text frame on `slide`.

    Source hymn and reading decks routinely end a body with one or two blank
    paragraphs left over from editing. They are not harmless: each takes a full
    line of vertical space, and because these shapes often can't be matched to
    the destination layout's placeholders they inherit the master's bullet
    character, so a blank line renders as a stray "•". Blank paragraphs
    *between* stanzas are kept; those are deliberate spacing.

    Returns the number of paragraphs removed.
    """
    removed = 0
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        paragraphs = shape.text_frame.paragraphs
        if len(paragraphs) <= 1:
            continue
        for p in reversed(paragraphs[1:]):
            if p.text.strip():
                break
            p._p.getparent().remove(p._p)
            removed += 1
    return removed


def shrink_overflowing_text(slide, min_scale=MIN_SCALE):
    """
    Scale down any text frame on `slide` whose content doesn't fit its box.

    Returns the number of shapes that were changed.
    """
    changed = 0
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if _shrink_shape(shape, min_scale):
            changed += 1
    return changed


def _shrink_shape(shape, min_scale):
    box = _resolved_box(shape)
    if box is None:
        return False

    inset = _insets(shape)
    usable_w = max(0, int(box[0]) - 2 * int(inset[0]))
    usable_h = max(0, int(box[1]) - 2 * int(inset[1]))
    if usable_w <= 0 or usable_h <= 0:
        return False

    default = _inherited_size(shape)
    sizes = _paragraph_sizes(shape, default)
    if not sizes or not any(sizes):
        return False

    needed = _needed_height(shape, sizes, usable_w)
    if needed <= usable_h:
        return False

    for scale in SCALE_STEPS:
        if scale < min_scale:
            break
        scaled = [_scaled(s, scale) for s in sizes]
        if not all(scaled):
            break
        if _needed_height(shape, scaled, usable_w) <= usable_h:
            _apply_sizes(shape, scaled)
            logger.info("Shrank %r to %d%% to fit its box", shape.name,
                        round(scale * 100))
            return True

    floor = [_scaled(s, min_scale) for s in sizes]
    _apply_sizes(shape, floor)
    logger.warning(
        "%r still overflows at %d%% of its original size (needs %.2fin of "
        "%.2fin); the source slide holds more text than the box fits — split "
        "it by hand.", shape.name, round(min_scale * 100),
        _needed_height(shape, floor, usable_w) / 914400, usable_h / 914400)
    return True


def _scaled(size, scale):
    if not size:
        return None
    pts = round(int(size) * scale / 12700)
    return Pt(pts) if pts > 0 else None


def _needed_height(shape, sizes, usable_w):
    """
    Rendered height of a shape's text, in EMU.

    Each paragraph is wrapped at its own size, so the columns available differ
    per paragraph; a 32pt line fits far more characters than a 54pt one. Lines
    are spaced by LINE_PITCH, and a paragraph gap is added only where the
    master actually applies one — see `_para_spacing`.
    """
    spacing = _para_spacing(shape)
    total = 0.0
    if spacing:
        first = next((int(x) for x in sizes if x), 0)
        total += first * TOP_LEADING
    for i, (p, size) in enumerate(zip(shape.text_frame.paragraphs, sizes)):
        em = int(size) if size else 0
        if em <= 0:
            continue
        # python-pptx reports an <a:br> as \x0b; each is its own rendered line.
        text = p.text.replace("\x0b", "\n")
        cols = max(0.0, usable_w / em - COPIED_WIDTH_MARGIN_EM)
        lines = 1 if not text.strip() else estimate_lines(text, cols)
        total += lines * em * LINE_PITCH
        if i:
            total += em * spacing
    return total


def _para_spacing(shape):
    """
    Extra gap before each paragraph, as a multiple of the font size.

    The 詩歌 master sets `spcBef` to 20% on its body style, so a body
    placeholder gets it. Its title style sets 0%, and a plain text box follows
    `otherStyle` with none — measured as exactly 1.2x pitch in a rendered deck,
    against 1.4x for a body placeholder.
    """
    ph_idx = placeholder_idx(shape)
    return PARA_SPACING if ph_idx not in (None, 0) else 0.0


def _paragraph_sizes(shape, default):
    """
    Effective font size per paragraph: the largest size any of its runs
    renders at.

    A paragraph's line height is set by its tallest glyphs, and a run with no
    explicit size still renders at the inherited default. These hymn decks put
    an explicit 32pt on the romanisation runs only — "性命無礙(gāi)，" is 54pt
    Chinese with a 32pt parenthetical — so resolving the paragraph to the
    explicit 32pt under-measured it by 40% and the slide was never shrunk.
    """
    sizes = []
    for p in shape.text_frame.paragraphs:
        candidates = [r.font.size or p.font.size or default for r in p.runs]
        candidates = [c for c in candidates if c]
        if not candidates:
            fallback = p.font.size or default
            candidates = [fallback] if fallback else []
        sizes.append(max(candidates, key=int) if candidates else None)
    return sizes


def _resolved_box(shape):
    """(width, height) of a shape, resolving placeholder inheritance."""
    try:
        w, h = shape.width, shape.height
    except (AttributeError, ValueError):
        return None
    if not w or not h or w <= 0 or h <= 0:
        return None
    return (w, h)


def _insets(shape):
    """Text-box insets, falling back to the master defaults."""
    tf = shape.text_frame
    left = tf.margin_left if tf.margin_left is not None else DEFAULT_INSET[0]
    top = tf.margin_top if tf.margin_top is not None else DEFAULT_INSET[1]
    return (left, top)


def _inherited_size(shape):
    """
    The size a run with no explicit size renders at.

    python-pptx returns None for anything inherited, so walk the chain the way
    PowerPoint does: the placeholder's own list style on the layout, then on
    the master, then the master's title/body/other text style. Title
    placeholders resolve against titleStyle — deciding that from the
    placeholder *type* matters, because a title's `idx` is optional and some
    decks write it as a 32-bit sentinel.
    """
    try:
        layout = shape.part.slide.slide_layout
    except AttributeError:
        return None

    ph_idx = placeholder_idx(shape)
    if ph_idx is not None:
        for source in (layout, layout.slide_master):
            size = _list_style_size(source, ph_idx)
            if size:
                return size
        style = "titleStyle" if ph_idx == 0 else "bodyStyle"
        size = _master_text_style_size(layout.slide_master, style)
        if size:
            return size

    for style in ("bodyStyle", "otherStyle"):
        size = _master_text_style_size(layout.slide_master, style)
        if size:
            return size
    return None


def _list_style_size(source, ph_idx):
    """lvl1 defRPr@sz from a layout/master placeholder's own <a:lstStyle>."""
    try:
        placeholders = source.placeholders
    except AttributeError:
        return None
    for ph in placeholders:
        if placeholder_idx(ph) != ph_idx:
            continue
        body = ph._element.find(qn("p:txBody"))
        if body is None:
            continue
        lst = body.find(qn("a:lstStyle"))
        if lst is None:
            continue
        lvl1 = lst.find(qn("a:lvl1pPr"))
        if lvl1 is None:
            continue
        def_rpr = lvl1.find(qn("a:defRPr"))
        if def_rpr is not None and def_rpr.get("sz"):
            return Pt(int(def_rpr.get("sz")) / 100)
    return None


def _master_text_style_size(master, style_name):
    """lvl1 defRPr@sz from the master's <p:txStyles>."""
    styles = master._element.find(qn("p:txStyles"))
    if styles is None:
        return None
    style = styles.find(qn(f"p:{style_name}"))
    if style is None:
        return None
    lvl1 = style.find(qn("a:lvl1pPr"))
    if lvl1 is None:
        return None
    def_rpr = lvl1.find(qn("a:defRPr"))
    if def_rpr is not None and def_rpr.get("sz"):
        return Pt(int(def_rpr.get("sz")) / 100)
    return None


def _apply_sizes(shape, sizes):
    """Set explicit per-paragraph sizes on a shape's runs and paragraphs."""
    for p, size in zip(shape.text_frame.paragraphs, sizes):
        if not size:
            continue
        p.font.size = size
        for r in p.runs:
            r.font.size = size
