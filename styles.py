"""
Shared colours and the run-writing helpers every generated slide uses.

Positions, sizes and fonts per slide type live in service_config.SLIDE_STYLES.
"""

from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

# ── Colors (from reference output) ──
CYAN_BRIGHT = RGBColor(0x00, 0xFF, 0xFF)
CYAN_SOFT   = RGBColor(0x66, 0xFF, 0xFF)
YELLOW      = RGBColor(0xFF, 0xFF, 0x00)


def apply_font(font, spec):
    """
    Apply a {font, size_pt, bold, color} spec to a run's font.

    Keys absent from the spec are left alone so they inherit from the layout.
    """
    if not spec:
        return
    if spec.get("font"):
        font.name = spec["font"]
    if spec.get("size_pt"):
        font.size = spec["size_pt"]
    if spec.get("bold") is not None:
        font.bold = spec["bold"]
    if spec.get("color") is not None:
        font.color.rgb = spec["color"]


def add_break(paragraph):
    """
    Append a real `<a:br/>` line break to a paragraph.

    Inserted before `<a:endParaRPr>` when present: OOXML fixes the child order
    of a paragraph, and PowerPoint silently drops anything that follows
    endParaRPr.
    """
    br = paragraph._p.makeelement(qn("a:br"), {})
    end = paragraph._p.find(qn("a:endParaRPr"))
    if end is not None:
        end.addprevious(br)
    else:
        paragraph._p.append(br)
    return br


def write_lines(paragraph, lines, spec=None):
    """
    Write `lines` into a single paragraph, separated by real `<a:br/>` breaks.

    Two reasons this is a single paragraph rather than one paragraph per line:

    * A literal "\\n" assigned through `run.text` lands inside `<a:t>`, where
      OOXML collapses it. PowerPoint renders no break at all, so every computed
      wrap was being discarded and the text ran off the slide edge.
    * The slide master sets `spcBef` to 20% on body paragraphs, so each extra
      paragraph eats a fifth of a line of vertical budget. Breaks inside one
      paragraph cost nothing, which is how the hand-finished decks are built.

    Returns the runs that were added.
    """
    runs = []
    for i, line in enumerate(lines):
        if i:
            add_break(paragraph)
        run = paragraph.add_run()
        run.text = line
        apply_font(run.font, spec)
        runs.append(run)
    return runs
