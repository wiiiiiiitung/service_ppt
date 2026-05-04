"""
Service worship PPTX generation configuration.

Defines the section order (SECTIONS) and all slide styles/positions/fonts (SLIDE_STYLES).
This is the single source of truth for the entire PPTX structure.
"""

from pptx.util import Emu, Pt

# ── Worship service section sequence ──────────────────────────────────────────
# Each entry describes one block of slides in the output PPTX.
# Types:
#   "fixed"      → copy slides from template by text marker
#   "from_input" → copy all slides from the matched input PPTX as-is
#   "generated"  → build slides programmatically using SLIDE_STYLES

SECTIONS = [
    {"id": "intro",           "type": "fixed",      "markers": ["歡迎","Zoom","安靜","開  會  詩"], "count": 6},
    {"id": "call_to_worship", "type": "fixed",      "markers": ["宣  召"],    "after_blank": True},
    {"id": "hymn_1",          "type": "from_input", "agenda_type": "hymn",       "after_blank": True},
    {"id": "prayer_block",    "type": "fixed",      "markers": ["祈  禱","主禱文","信仰告白"], "after_blank": True},
    {"id": "responsive",      "type": "from_input", "agenda_type": "responsive", "after_blank": True},
    {"id": "anthem",          "type": "generated",  "agenda_type": "anthem",     "after_blank": True},
    {"id": "scripture",       "type": "generated",  "agenda_type": "scripture",  "after_blank": False},
    {"id": "sermon",          "type": "generated",  "agenda_type": "sermon",     "after_blank": True},
    {"id": "hymn_2",          "type": "from_input", "agenda_type": "hymn",       "after_blank": True},
    {"id": "offering",        "type": "fixed",      "markers": ["奉獻"],         "after_blank": False},
    {"id": "communion",       "type": "fixed",      "markers": ["聖餐"],         "conditional": "communion_in_agenda", "after_blank": True},
    {"id": "announcements",   "type": "generated",  "agenda_type": "announcements", "after_blank": True},
    {"id": "closing",         "type": "fixed",      "markers": ["頌榮","祝  禱","默 禱","rcnewtown"]},
]

# ── Per-section slide styles ──────────────────────────────────────────────────
# All positions/sizes in EMU (914400 EMU = 1 inch).
# Font: None = inherited from layout. Font size in Pt.

SLIDE_STYLES = {

    # 獻詩 title: "獻詩: {title}" centered mid-slide
    "anthem_title": {
        "layout":   "Blank",
        "pos":      (Emu(981777), Emu(2786743)),
        "size":     (Emu(10501161), Emu(1015663)),
        "text":     "獻詩: {title}",
        "font":     "標楷體",
        "size_pt":  Pt(60),
        "bold":     True,
        "align":    "center",
        "color":    None,  # inherit from layout
    },

    # 獻詩 / Hymn lyrics: use 詩歌 layout placeholders
    "lyrics": {
        "layout":       "詩歌",
        "title_ph_idx": 0,  # placeholder for song title
        "body_ph_idx":  1,  # placeholder for lyrics lines
        # fonts/colors all inherited from 詩歌 layout
    },

    # 報告 title: "報告" centered mid-slide
    "announcement_title": {
        "layout":   "Blank",
        "pos":      (Emu(3359150), Emu(2492375)),
        "size":     (Emu(5040313), Emu(1006475)),
        "text":     "報告",
        "font":     "標楷體",
        "size_pt":  Pt(60),
        "bold":     True,
        "align":    "center",
        "color":    None,
    },

    # 報告 item: one item per slide using 詩歌 layout
    "announcement_item": {
        "layout":       "詩歌",
        "title_ph_idx": 0,  # "報告： {section}"
        "body_ph_idx":  1,  # item text, align LEFT
        "align":        "left",
    },

    # 經文 title slide: mixed-run paragraph in one rectangle
    "scripture_title": {
        "layout":   "Blank",
        "pos":      (Emu(154746), Emu(2086708)),
        "size":     (Emu(12037254), Emu(2000250)),
        # runs in order: label | reference | page-hint
        "runs": [
            {"text": "經文",              "font": None,       "size_pt": Pt(60), "bold": True},
            {"text": "{ref}",            "font": "DFKai-SB", "size_pt": Pt(72), "bold": True},
            {"text": "  ",               "font": "DFKai-SB", "size_pt": Pt(72), "bold": True},
            {"text": "({testament}第{page}頁)", "font": "標楷體", "size_pt": Pt(54), "bold": True},
        ],
        "color": None,
    },

    # 經文 verse slides: reference bar on top + verse body
    "scripture_verse_bar": {
        "layout":   "Blank",
        "pos":      (Emu(0), Emu(0)),
        "size":     (Emu(12192000), Emu(1561514)),
        "font":     "標楷體",
        "bold":     True,
        "align":    "left",
        "color":    None,
    },
    "scripture_verse_body": {
        "pos":      (Emu(-22035), Emu(815926)),
        "size":     (Emu(12537195), Emu(6084277)),
        "font":     "DFKai-SB",
        "size_pt":  Pt(54),
        "bold":     True,
        "align":    "left",
        "color":    None,
        "verse_format": "{n}.{text}",  # one paragraph per verse
    },

    # 今日信息 title slide
    "sermon_title": {
        "layout":        "詩歌",
        "header_text":   "今日信息",
        "header_ph_idx": 0,
        "body_ph_idx":   1,
        "title_size_pt": Pt(80),
        "spacer_font":   "DFKai-SB",
        "spacer_size_pt": Pt(50),
        "spacer_bold":   True,
        "preacher_font": "標楷體",
        "preacher_size_pt": Pt(44),
        "preacher_bold": True,
    },

    # 今日信息 point slides (title placeholder removed, full-slide body)
    "sermon_point": {
        "layout":         "詩歌",
        "remove_title_ph": True,
        "body_ph_idx":    1,
        "header_text":    "今日信息",
        "header_font":    "標楷體",
        "header_size_pt": Pt(44),
        "header_bold":    True,
        "points_align":   "left",
    },
}
