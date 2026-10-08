"""
Service worship PPTX generation configuration.

Defines all slide styles/positions/fonts (SLIDE_STYLES) for generated slides.
Section ordering lives in slide_planner.plan_slides().

Every style that holds flowing text carries a `box` (and, where it differs from
the master default, an `inset` and `line_height`) so text_layout can work out
how much actually fits instead of relying on hard-coded line counts. The boxes
below are measured from the 詩歌 layout in template/intro.pptx.
"""

from pptx.util import Emu, Pt

from styles import CYAN_SOFT, YELLOW

# ── Text-cleanup toggles ────────────────────────────────────────────────────
# Church bulletin PDFs are justified, so pdfplumber extracts spurious spaces at
# line-wrap points in Chinese text. Collapse whitespace that sits between two
# CJK characters on announcement slides. Spaces bordering Latin text (URLs,
# phone numbers, "會號 994 2970 4298") are preserved — the reference decks keep
# them. Set to False to keep the raw PDF-extracted spacing.
ANNOUNCEMENT_STRIP_CJK_SPACES = True

# ── Measured layout geometry ────────────────────────────────────────────────
# The 詩歌 layout's body placeholder, which every lyric / sermon / announcement
# slide uses. 914400 EMU = 1 inch.
SONG_BODY_POS = (Emu(0), Emu(990600))
SONG_BODY_BOX = (Emu(12192000), Emu(5867400))
# Master text-box insets: lIns/rIns 0.1in, tIns/bIns 0.05in.
BODY_INSET = (Emu(91440), Emu(45720))

# Font sizes an over-long announcement may be stepped down through, standing in
# for the shrink-on-overflow autofit these decks don't enable.
ANNOUNCEMENT_SIZE_LADDER = [Pt(54), Pt(48), Pt(44), Pt(40), Pt(36), Pt(32), Pt(28)]

# ── Per-section slide styles ──────────────────────────────────────────────────
# Font: None = inherited from layout. Font size in Pt.

SLIDE_STYLES = {

    # 獻詩 title: "獻詩: {title}" centered mid-slide
    "anthem_title": {
        "layout":   "Blank",
        "pos":      (Emu(2677538), Emu(2786743)),
        "size":     (Emu(7109639), Emu(1015663)),
        "text":     "獻詩: {title}",
        "font":     "標楷體",
        "size_pt":  Pt(60),
        "bold":     True,
        "align":    "center",
        "color":    CYAN_SOFT,
    },

    # 獻詩 / Hymn lyrics: use 詩歌 layout placeholders
    "lyrics": {
        "layout":       "詩歌",
        "title_ph_idx": 0,
        "body_ph_idx":  1,
        "box":          SONG_BODY_BOX,
        "size_pt":      Pt(54),   # 詩歌 layout body default
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
        "box":          SONG_BODY_BOX,
        "size_pt":      Pt(54),
        "size_ladder":  ANNOUNCEMENT_SIZE_LADDER,
    },

    # 經文 title slide: three centered lines (label / reference / page hint)
    "scripture_title": {
        "layout":   "Blank",
        "pos":      (Emu(154746), Emu(2086708)),
        "size":     (Emu(12037254), Emu(3139321)),
        "align":    "center",
        # one entry per paragraph; each is a list of runs
        "lines": [
            [{"text": "經文", "font": "DFKai-SB", "size_pt": Pt(66), "bold": True, "color": CYAN_SOFT}],
            [{"text": "{ref}", "font": "DFKai-SB", "size_pt": Pt(72), "bold": True}],
            [
                {"text": "({testament}第", "font": "標楷體", "size_pt": Pt(54), "bold": True, "page_only": True},
                {"text": "{page}",         "font": "標楷體", "size_pt": Pt(54), "bold": True, "page_only": True, "color": YELLOW},
                {"text": "頁)",            "font": "標楷體", "size_pt": Pt(54), "bold": True, "page_only": True},
            ],
        ],
    },

    # 經文 verse slides: reference bar on top + verse body
    "scripture_verse_bar": {
        "layout":   "Blank",
        "pos":      (Emu(4080063), Emu(0)),
        "size":     (Emu(4031873), Emu(707886)),
        "font":     "標楷體",
        "size_pt":  Pt(40),
        "bold":     True,
        "align":    "center",
        "color":    CYAN_SOFT,
    },
    "scripture_verse_body": {
        # The reference decks place this box at x=-0.024in with a width of
        # 13.711in, which overhangs the 13.333in slide by 0.35in — enough to
        # clip the last glyph of a full line. Clamped to the slide so wrapping
        # happens where the text is still visible.
        "pos":      (Emu(0), Emu(815926)),
        "size":     (Emu(12192000), Emu(6042074)),
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
        # The reference decks open the body with an empty paragraph, which
        # drops the title to the vertical centre of the box.
        "lead_blank_line": True,
        "spacer_text":   " " * 51,
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
        "pos":            (Emu(0), Emu(-99152)),
        "size":           (Emu(12192000), Emu(6957152)),
        "box":            (Emu(12192000), Emu(6957152)),
        "header_text":    "今日信息",
        "header_font":    "標楷體",
        "header_size_pt": Pt(44),
        "header_bold":    True,
        "header_color":   CYAN_SOFT,
        "points_align":   "left",
        "size_pt":        Pt(54),   # points inherit the 詩歌 body size
    },
}
