"""
Text measurement and slide-packing used by every generated slide.

All three generated sections (scripture, sermon, announcements) and the anthem
lyrics need the same thing: given a fixed-size text box and a font size, how
many lines fit, where should long lines wrap, and how should items be split
across slides. Getting that wrong is what pushes text off the bottom of a
slide or wraps it early, so the arithmetic lives in one place.

Calibration note
----------------
Both constants are measured off rendered PDFs, not assumed.

`LINE_PITCH` is 1.2 — plain single spacing. Measuring consecutive baselines in
a rendered deck gives exactly 60.0pt for 50pt text in a plain text box.

`PARA_SPACING` is the *extra* gap between paragraphs, 0.2 of the font size,
and it applies only where the slide master's `spcBef` does: body placeholders
on the 詩歌 master. The same measurement on a 詩歌 placeholder gives 70.0pt
for 50pt text — 1.2 pitch plus 0.2 spacing. A plain text box gets no spacing
at all.

Keeping them separate matters. Rolling both into one 1.36 factor and *also*
adding paragraph spacing double-counted the gap and over-predicted the height
of anything with several paragraphs by ~30%, which made the autofit pass
shrink slides that fit perfectly well.
"""

import math

EMU_PER_PT = 12700
EMU_PER_IN = 914400

# Measured: see the calibration note above.
LINE_PITCH = 1.2
# Extra gap before each paragraph after the first, as a multiple of the font
# size. The 詩歌 master sets spcBef to 20% for body placeholders; a plain text
# box inherits none, so pass 0 for those.
PARA_SPACING = 0.2

# Backwards-compatible alias for the old single bundled factor.
DEFAULT_LINE_HEIGHT = LINE_PITCH

# Default text-box insets from the slide master (lIns/rIns 0.1in, tIns/bIns 0.05in).
DEFAULT_INSET = (91440, 45720)

# Usable width is shaved by this much (in em) before wrapping. A CJK glyph's
# advance is nominally 1em but runs slightly over in 標楷體 / DFKai-SB, so a
# line computed to land exactly on the box edge rendered ~5pt past it. An
# absolute margin rather than a percentage: a percentage scaled with the box
# and cost a whole character on the wide scripture box, wrapping a line earlier
# than the reference decks do.
WIDTH_MARGIN_EM = 0.3
WIDTH_SAFETY = 0.97  # retained for callers that still scale a ratio


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


def capacity(box, size_pt, inset=DEFAULT_INSET, line_pitch=LINE_PITCH,
             para_spacing=PARA_SPACING):
    """
    How much text a box holds.

    `max_lines` is worked out for the worst case of one line per paragraph, so
    it accounts for the paragraph gap: N lines occupy
    `N * line_pitch + (N - 1) * para_spacing` ems. Pass `para_spacing=0` for a
    plain text box, which inherits no `spcBef`.

    Args:
        box: (width_emu, height_emu) of the text box
        size_pt: font size, as a python-pptx Length (Pt) or raw EMU int
        inset: (horizontal, vertical) inset per side, in EMU
        line_pitch: baseline-to-baseline distance, in ems
        para_spacing: extra gap before each paragraph after the first, in ems

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

    cols = max(0.0, usable_w / em - WIDTH_MARGIN_EM)

    budget = usable_h / em
    max_lines = 0
    while True:
        n = max_lines + 1
        if n * line_pitch + (n - 1) * para_spacing > budget:
            break
        max_lines = n
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


# Kinsoku shori: characters that may not begin a line. A line that starts with
# a comma or a closing bracket reads as a typographic error, and the worst case
# — a line holding nothing but "」" — happened regularly.
NO_LINE_START = "，。、；：！？）」』】》〉·…ーヽヾ々,.;:!?)]}%‰°″′’”"
# Characters that may not end a line: an opening bracket belongs with the text
# it opens.
NO_LINE_END = "（「『【《〈([{‘“"


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

    # Prefer a clause boundary within the last ~20% of the line: breaking just
    # after the punctuation keeps it on the line it belongs to.
    window = max(1, int(hard * 0.2))
    for i in range(hard, hard - window, -1):
        if text[i - 1] in "，、；。！？,;.!?":
            nxt = text[i] if i < len(text) else ""
            if not nxt or nxt not in NO_LINE_START:
                return i

    # Otherwise avoid splitting a Latin/digit run.
    i = hard
    while i > 1 and not _is_wide(text[i - 1]) and not text[i - 1].isspace() \
            and not _is_wide(text[i]) and not text[i].isspace():
        i -= 1
    cut = i if i > 1 else hard
    return _apply_kinsoku(text, cut, hard)


def _apply_kinsoku(text, cut, hard):
    """
    Nudge a break point so it doesn't orphan punctuation.

    Pulls the break earlier while the next line would start with a character
    that may not begin one, or the current line would end with an opening
    bracket. Gives up rather than shrinking the line to nothing.
    """
    floor = max(1, hard - 4)
    while cut > floor:
        nxt = text[cut] if cut < len(text) else ""
        prev = text[cut - 1]
        if nxt and nxt in NO_LINE_START:
            cut -= 1
            continue
        if prev in NO_LINE_END:
            cut -= 1
            continue
        break
    return cut


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
