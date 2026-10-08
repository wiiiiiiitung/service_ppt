"""
Find slides in the fixed-slide deck by content marker instead of by index.

Markers must be *unambiguous*: `find_slide` returns the first match, so a
marker that also appears in an announcement or a hymn lyric silently resolves
to the wrong slide. `ambiguous_markers()` reports any that match more than one
slide, and the regression harness asserts it comes back empty.
"""

import logging

logger = logging.getLogger(__name__)

# Map slide semantic names to text markers identifying them in the fixed deck.
# A None marker means "no text to match on" (the blank slide).
SLIDE_MARKERS = {
    "logo":            None,  # church logo slide; carried by template/intro.pptx
    "welcome":         "歡迎來到",
    "blank":           None,
    "zoom_info":       "華語翻譯",
    "prepare":         "敬虔的心",
    "opening":         "開  會  詩",
    "call_to_worship": "宣  召",
    "prayer":          "祈  禱",
    "lords_prayer_1":  "阮在天裡的父",
    "lords_prayer_2":  "勿得導阮",
    "creed_1":         "我信上帝,全能的父",
    "creed_2":         "第三日對死人中復活",
    "offering_1":      "捐得樂意",
    "offering_2":      "我的生命獻給祢",
    "communion_1":     "聖餐",
    "communion_2":     "耶穌身軀替咱釘死",
    "communion_3":     "與主同桌",
    "announce_title":  "報告",
    "doxology":        "榮光歸聖父上帝",
    "benediction":     "祝  禱",
    "quiet":           "默 禱",
    "website":         "今天的講台信息",
}

# Markers that legitimately appear on several slides; the first match is the
# section title slide, which is the one we want.
_FIRST_MATCH_OK = {"communion_1", "announce_title"}


def _slide_texts(prs):
    return [
        " ".join(shape.text for shape in slide.shapes if shape.has_text_frame)
        for slide in prs.slides
    ]


def find_slide(prs, key):
    """
    Find the first slide in prs whose text contains the marker for key.

    Returns a 0-based slide index, or None when the key is unknown, has no
    marker, or nothing matches.
    """
    marker = SLIDE_MARKERS.get(key)
    if not marker:
        return None

    hits = [i for i, text in enumerate(_slide_texts(prs)) if marker in text]
    if not hits:
        logger.warning("Fixed-slide marker %r (%s) matched no slide", marker, key)
        return None
    if len(hits) > 1 and key not in _FIRST_MATCH_OK:
        logger.warning(
            "Fixed-slide marker %r (%s) matched slides %s; using the first. "
            "Make the marker more specific.", marker, key, hits
        )
    return hits[0]


def ambiguous_markers(prs):
    """
    Report markers that match more than one slide, or none at all.

    Returns {key: [slide indices]} for every marker that does not resolve to
    exactly one slide. Used by the regression harness to catch marker drift
    when the fixed deck is replaced.
    """
    texts = _slide_texts(prs)
    bad = {}
    for key, marker in SLIDE_MARKERS.items():
        if not marker or key in _FIRST_MATCH_OK:
            continue
        hits = [i for i, text in enumerate(texts) if marker in text]
        if len(hits) != 1:
            bad[key] = hits
    return bad
