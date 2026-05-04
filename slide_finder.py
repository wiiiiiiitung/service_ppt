"""
Find template slides by content marker instead of hardcoded indices.

Replaces the old TEMPLATE_SLIDES dict with a content-based lookup that is robust
to template edits.
"""

# Map slide semantic names to text markers that identify them in the template PPTX
SLIDE_MARKERS = {
    "logo":            "c",  # church logo or first intro slide marker
    "welcome":         "歡迎",
    "blank":           None,  # blank slide (no marker)
    "zoom_info":       "Zoom",
    "prepare":         "安靜",
    "opening":         "開  會  詩",
    "call_to_worship": "宣  召",
    "prayer":          "祈  禱",
    "lords_prayer_1":  "主禱文",
    "lords_prayer_2":  "勿得導阮",  # second 主禱文 slide
    "creed_1":         "信仰告白",
    "creed_2":         "第三日對死人中復活",
    "offering_1":      "捐得樂意",
    "offering_2":      "我的生命獻給祢",
    "communion_1":     "聖餐",
    "communion_2":     "215",
    "communion_3":     "與主同桌",
    "announce_title":  "報告",
    "doxology":        "頌榮",
    "benediction":     "祝  禱",
    "quiet":           "默 禱",
    "website":         "rcnewtown",
}


def find_slide(prs, key):
    """
    Find the first slide in prs whose text contains the marker for key.

    Args:
        prs: Presentation object
        key: key in SLIDE_MARKERS

    Returns:
        0-based slide index, or None if not found
    """
    if key not in SLIDE_MARKERS:
        return None

    marker = SLIDE_MARKERS[key]
    for i, slide in enumerate(prs.slides):
        text = " ".join(
            shape.text for shape in slide.shapes if shape.has_text_frame
        )
        if marker in text:
            return i

    return None


def find_consecutive(prs, keys):
    """
    Find a sequence of consecutive marker keys.

    Returns list of slide indices in order (some may be None if not found).
    """
    return [find_slide(prs, k) for k in keys]
