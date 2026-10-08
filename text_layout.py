"""
Text measurement and slide-packing used by every generated slide.

All three generated sections (scripture, sermon, announcements) and the anthem
lyrics need the same thing: given a fixed-size text box and a font size, how
many lines fit, where should long lines wrap, and how should items be split
across slides. Getting that wrong is what pushes text off the bottom of a
slide or wraps it early, so the arithmetic lives in one place.

Calibration note
----------------
`DEFAULT_LINE_HEIGHT` is measured, not assumed. Rendering both the generated
and the hand-finished reference decks through LibreOffice and reading back the
text extents gives ~1.02in per line for 54pt 標楷體 — a factor of ~1.36, not
the ~1.2 a Latin face would use. The CJK faces this deck uses (標楷體 /
DFKai-SB) have unusually deep ascent+descent. With 1.2 the 詩歌 body box looks
like it holds 7 lines; it actually holds 6, which is exactly the off-by-one
that pushed the last lyric line off the slide.
"""

import math

EMU_PER_PT = 12700
EMU_PER_IN = 914400

# Measured: see the calibration note above.
DEFAULT_LINE_HEIGHT = 1.36

# Default text-box insets from the slide master (lIns/rIns 0.1in, tIns/bIns 0.05in).
DEFAULT_INSET = (91440, 45720)

# Usable width is shaved by this much before wrapping. A CJK glyph's advance is
# nominally 1em, but the real advance in 標楷體 / DFKai-SB runs slightly over,
# so a line computed to land exactly on the box edge rendered ~5pt past it.
WIDTH_SAFETY = 0.97


def _is_wide(ch):
    """True for characters that occupy a full em (CJK ideographs, full-width forms)."""
    cp = ord(ch)
    return (
        0x1100 <= cp <= 0x115F       # Hangul Jamo
        or 0x2E80 <= cp <= 0xA4CF    # CJK radicals … Yi
        or 0xAC00 <= cp <= 0xD7A3    # Hangul syllables
        or 0xF900 <= cp <= 0xFAFF    # CJK compatibility ideographs
        or 0xFE30 <= cp <= 0xFE4F    # CJK compatibility forms
        or 0xFF00 <= cp <= 0xFF60    # full-width forms
        or 0xFFE0 <= cp <= 0xFFE6
        or 0x20000 <= cp <= 0x3FFFD  # CJK extension planes
    )


def display_width(text):
    """
    Width of `text` in em units: 1.0 per CJK/full-width character, 0.5 per
    Latin character, digit or space.

    Counting raw `len()` instead badly overestimates Latin-heavy text — an
    announcement full of URLs, phone numbers and English names measured as
    twice its real width, which is why those items were being split across
    slides and cut mid-word.
    """
    return sum(1.0 if _is_wide(ch) else 0.5 for ch in text)


def capacity(box, size_pt, inset=DEFAULT_INSET, line_height=DEFAULT_LINE_HEIGHT):
    """
    How much text a box holds.

    Args:
        box: (width_emu, height_emu) of the text box
        size_pt: font size, as a python-pptx Length (Pt) or raw EMU int
        inset: (horizontal, vertical) inset per side, in EMU
        line_height: multiple of the font size one line occupies

    Returns:
        (cols, max_lines) — cols is the usable width in em units, so compare it
        against `display_width()`, not against `len()`.
    """
    width, height = int(box[0]), int(box[1])
    em = int(size_pt)
    if em <= 0:
        return (0, 0)
    usable_w = max(0, width - 2 * int(inset[0]))
    usable_h = max(0, height - 2 * int(inset[1]))
    cols = (usable_w / em) * WIDTH_SAFETY
    max_lines = int(usable_h // (em * line_height))
    return (cols, max(1, max_lines))


def estimate_lines(text, cols):
    """Wrapped line count for `text` in a box `cols` em wide."""
    if cols <= 0:
        return 1
    widest = max((display_width(line) for line in text.split("\n")), default=0)
    if "\n" in text:
        return sum(max(1, math.ceil(display_width(l) / cols)) for l in text.split("\n"))
    return max(1, math.ceil(widest / cols))


def wrap_text(text, cols, indent=""):
    """
    Hard-wrap `text` to `cols` em per line, prefixing continuation lines with
    `indent`. Returns a list of lines.

    Breaks at a CJK clause boundary near the limit when there is one, so a
    continuation line doesn't start mid-clause, and never breaks inside a run
    of Latin characters (a URL, a phone number, a name).
    """
    if cols <= 0 or display_width(text) <= cols:
        return [text]

    lines = []
    remaining = text
    indent_w = display_width(indent)
    first = True
    while remaining:
        budget = cols if first else cols - indent_w
        if display_width(remaining) <= budget:
            lines.append(remaining if first else indent + remaining)
            break
        cut = _cut_point(remaining, budget)
        lines.append(remaining[:cut] if first else indent + remaining[:cut])
        remaining = remaining[cut:]
        first = False
    return lines


def _cut_point(text, budget):
    """Index to break `text` at so the first piece is at most `budget` em wide."""
    width = 0.0
    hard = len(text)
    for i, ch in enumerate(text):
        width += 1.0 if _is_wide(ch) else 0.5
        if width > budget:
            hard = i
            break
    hard = max(1, hard)

    # Prefer a clause boundary within the last ~20% of the line.
    window = max(1, int(hard * 0.2))
    for i in range(hard, hard - window, -1):
        if text[i - 1] in "，、；。！？,;.!?":
            return i

    # Otherwise avoid splitting a Latin/digit run.
    i = hard
    while i > 1 and not _is_wide(text[i - 1]) and not text[i - 1].isspace() \
            and not _is_wide(text[i]) and not text[i].isspace():
        i -= 1
    return i if i > 1 else hard


def pack_lines(lines, max_lines):
    """Group already-wrapped lines into slide-sized chunks."""
    if max_lines <= 0:
        return [list(lines)] if lines else []
    return [list(lines[i:i + max_lines]) for i in range(0, len(lines), max_lines)]


def fit_size(text, box, ladder, inset=DEFAULT_INSET,
             line_height=DEFAULT_LINE_HEIGHT):
    """
    Largest size from `ladder` at which `text` fits `box` on one slide.

    Stands in for PowerPoint's shrink-on-overflow autofit, which these decks
    don't enable: an over-long announcement is set a step or two smaller rather
    than being chopped across slides.

    Returns (size, lines) using the smallest size in the ladder if nothing fits.
    """
    sizes = sorted(ladder, key=int, reverse=True)
    last = (sizes[-1], None)
    for size in sizes:
        cols, max_lines = capacity(box, size, inset, line_height)
        lines = []
        for para in text.split("\n"):
            lines.extend(wrap_text(para, cols))
        if len(lines) <= max_lines:
            return (size, lines)
        last = (size, lines)
    size = sizes[-1]
    cols, _ = capacity(box, size, inset, line_height)
    lines = []
    for para in text.split("\n"):
        lines.extend(wrap_text(para, cols))
    return (size, lines)
