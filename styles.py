"""
Style registry for dynamically-generated slides.

Each entry captures the position, size, font, and color that was observed in a
hand-crafted reference output. Builder functions call `get_style(name)` and fall
back to a sensible default when no entry exists.

Positions/sizes are in EMU (English Metric Units). 914400 EMU = 1 inch.
Font size values are in hundredths of a point multiplied by 100; we store them
as python-pptx `Pt(...)` values inside the builder.
"""

from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt


# ── Colors (from reference output) ──
CYAN_BRIGHT = RGBColor(0x00, 0xFF, 0xFF)
CYAN_SOFT   = RGBColor(0x66, 0xFF, 0xFF)
YELLOW      = RGBColor(0xFF, 0xFF, 0x00)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

# ── Font names ──
FONT_KAI    = "標楷體"
FONT_KAI_EN = "DFKai-SB"


# Each style describes a slide-element. Builders consult these and fall back to
# their own defaults when the key is missing.
STYLES = {
    # Scripture title: centered block, mid-slide
    "scripture_title": {
        "box": {
            "pos":  (Emu(154746), Emu(2086708)),
            "size": (Emu(12037254), Emu(2000250)),
        },
        "heading": {
            "font_name": FONT_KAI,
            "font_size": Pt(60),
            "bold": True,
            "color": CYAN_BRIGHT,
        },
        "reference": {
            "font_name": FONT_KAI_EN,
            "font_size": Pt(72),
            "bold": True,
            "color": CYAN_SOFT,
        },
        "align": "center",
    },

    # Scripture verse: top bar with reference, large body with verses
    "scripture_verse_bar": {
        "box": {
            "pos":  (Emu(0), Emu(0)),
            "size": (Emu(12192000), Emu(1561514)),
        },
        "font_name": FONT_KAI,
        "font_size": Pt(40),
        "bold": True,
        "color": CYAN_SOFT,
        "align": "center",
    },
    "scripture_verse_body": {
        "box": {
            "pos":  (Emu(0), Emu(815926)),  # adjusted from -22035 to 0 for safety
            "size": (Emu(12192000), Emu(6042074)),
        },
        "font_name": FONT_KAI_EN,
        "font_size": Pt(54),
        "bold": True,
        "color": WHITE,
        "chars_per_line": 16,  # Chinese chars per wrapped line; indent is
                                # dynamic (matches verse-number width)
    },

    # Sermon title slide: "今日信息" top + "<sermon title>   <preacher>" below
    "sermon_title_header": {
        "color": CYAN_BRIGHT,
    },
    "sermon_title_body": {
        "title": {
            "font_size": Pt(80),
        },
        "spacer": {
            "font_name": FONT_KAI_EN,
            "font_size": Pt(50),
            "bold": True,
        },
        "preacher": {
            "font_name": FONT_KAI,
            "font_size": Pt(44),
            "bold": True,
        },
    },

    # Sermon point slide: single body with styled "今日信息" + points
    "sermon_point_header": {
        "font_name": FONT_KAI,
        "font_size": Pt(44),
        "bold": True,
        "color": CYAN_BRIGHT,
    },
    "sermon_point_body": {
        "align": "left",
    },

    # Anthem title: centered one-liner '獻詩: <title>'
    "anthem_title": {
        "box": {
            "pos":  (Emu(981777), Emu(2786743)),
            "size": (Emu(10501161), Emu(1015663)),
        },
        "font_name": FONT_KAI,
        "font_size": Pt(60),
        "bold": True,
        "color": CYAN_SOFT,
        "align": "center",
    },
}


def get_style(name):
    """Return the style dict for `name`, or None if no entry exists."""
    return STYLES.get(name)


def apply_run_style(run, style):
    """Apply a subset of style keys (font_name/size/bold/color) to a run."""
    if not style:
        return
    if style.get("font_name"):
        run.font.name = style["font_name"]
    if style.get("font_size"):
        run.font.size = style["font_size"]
    if style.get("bold") is not None:
        run.font.bold = style["bold"]
    if style.get("color") is not None:
        run.font.color.rgb = style["color"]


_ALIGN_MAP = {"left": 1, "center": 2, "right": 3, "justify": 4}


def apply_paragraph_style(paragraph, style):
    """Apply alignment from a style to a paragraph."""
    if not style:
        return
    a = style.get("align")
    if a and a in _ALIGN_MAP:
        paragraph.alignment = _ALIGN_MAP[a]


def wrap_chinese_text(text, chars_per_line, indent=""):
    """
    Hard-wrap Chinese text at N characters per line.
    Continuation lines are prefixed with `indent`.
    """
    if chars_per_line <= 0 or len(text) <= chars_per_line:
        return text
    chunks = [text[i:i + chars_per_line] for i in range(0, len(text), chars_per_line)]
    return chunks[0] + "".join("\n" + indent + c for c in chunks[1:])
