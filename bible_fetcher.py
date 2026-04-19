"""Fetch 和合本 (Chinese Union Version) Bible passages from bible-api.com."""

import json
import re
import urllib.request
import urllib.parse


# Map common Chinese book abbreviations/full names to book names usable in the API
_BOOK_ALIASES = {
    "尼希米記": "尼希米記",
    "尼希米": "尼希米記",
    "創世記": "創世記",
    "出埃及記": "出埃及記",
    "利未記": "利未記",
    "民數記": "民數記",
    "申命記": "申命記",
    "約書亞記": "約書亞記",
    "士師記": "士師記",
    "路得記": "路得記",
    "撒母耳記上": "撒母耳記上",
    "撒母耳記下": "撒母耳記下",
    "列王紀上": "列王紀上",
    "列王紀下": "列王紀下",
    "歷代志上": "歷代志上",
    "歷代志下": "歷代志下",
    "以斯拉記": "以斯拉記",
    "以斯帖記": "以斯帖記",
    "約伯記": "約伯記",
    "詩篇": "詩篇",
    "箴言": "箴言",
    "傳道書": "傳道書",
    "雅歌": "雅歌",
    "以賽亞書": "以賽亞書",
    "耶利米書": "耶利米書",
    "耶利米哀歌": "耶利米哀歌",
    "以西結書": "以西結書",
    "但以理書": "但以理書",
    "何西阿書": "何西阿書",
    "約珥書": "約珥書",
    "阿摩司書": "阿摩司書",
    "俄巴底亞書": "俄巴底亞書",
    "約拿書": "約拿書",
    "彌迦書": "彌迦書",
    "那鴻書": "那鴻書",
    "哈巴谷書": "哈巴谷書",
    "西番雅書": "西番雅書",
    "哈該書": "哈該書",
    "撒迦利亞書": "撒迦利亞書",
    "瑪拉基書": "瑪拉基書",
    "馬太福音": "馬太福音",
    "馬可福音": "馬可福音",
    "路加福音": "路加福音",
    "約翰福音": "約翰福音",
    "使徒行傳": "使徒行傳",
    "羅馬書": "羅馬書",
    "哥林多前書": "哥林多前書",
    "哥林多後書": "哥林多後書",
    "加拉太書": "加拉太書",
    "以弗所書": "以弗所書",
    "腓立比書": "腓立比書",
    "歌羅西書": "歌羅西書",
    "帖撒羅尼迦前書": "帖撒羅尼迦前書",
    "帖撒羅尼迦後書": "帖撒羅尼迦後書",
    "提摩太前書": "提摩太前書",
    "提摩太後書": "提摩太後書",
    "提多書": "提多書",
    "腓利門書": "腓利門書",
    "希伯來書": "希伯來書",
    "雅各書": "雅各書",
    "彼得前書": "彼得前書",
    "彼得後書": "彼得後書",
    "約翰一書": "約翰一書",
    "約翰二書": "約翰二書",
    "約翰三書": "約翰三書",
    "猶大書": "猶大書",
    "啟示錄": "啟示錄",
}

API_BASE = "https://bible-api.com"


def parse_reference(ref):
    """
    Parse a Chinese scripture reference string.

    Examples:
        "尼希米記13:15~22" → ("尼希米記", 13, 15, 22)
        "羅馬書 4:1~12"    → ("羅馬書", 4, 1, 12)
        "詩篇 23:1"        → ("詩篇", 23, 1, 1)

    Returns (book, chapter, verse_start, verse_end) or None if unparseable.
    """
    ref = ref.strip()
    # Replace full-width colons and tildes
    ref = ref.replace("：", ":").replace("～", "~").replace("〜", "~").replace("－", "-")

    m = re.match(
        r"^([^\d]+?)\s*(\d+)\s*[:：]\s*(\d+)\s*[~\-～]\s*(\d+)",
        ref
    )
    if m:
        return m.group(1).strip(), int(m.group(2)), int(m.group(3)), int(m.group(4))

    m = re.match(r"^([^\d]+?)\s*(\d+)\s*[:：]\s*(\d+)", ref)
    if m:
        v = int(m.group(3))
        return m.group(1).strip(), int(m.group(2)), v, v

    return None


def fetch_verses(ref):
    """
    Fetch CUV (和合本) verses for a scripture reference string.

    Returns list of dicts: [{verse: int, text: str}, ...] or None on failure.
    """
    parsed = parse_reference(ref)
    if not parsed:
        return None

    book, chapter, v_start, v_end = parsed
    book_name = _BOOK_ALIASES.get(book, book)

    query = f"{book_name}{chapter}:{v_start}-{v_end}"
    url = f"{API_BASE}/{urllib.parse.quote(query)}?translation=cuv"

    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read())
        verses = data.get("verses", [])
        return [{"verse": v["verse"], "text": v["text"].strip()} for v in verses]
    except Exception:
        return None


def group_verses_for_slides(verses, chars_per_line=18, max_lines=7):
    """
    Group verses into slide-sized chunks.

    Each group contains verses that fit within max_lines lines of chars_per_line
    characters each.
    """
    groups = []
    current = []
    current_lines = 0

    for v in verses:
        verse_lines = max(1, -(-len(v["text"]) // chars_per_line))  # ceil division
        if current and current_lines + verse_lines > max_lines:
            groups.append(current)
            current = [v]
            current_lines = verse_lines
        else:
            current.append(v)
            current_lines += verse_lines

    if current:
        groups.append(current)

    return groups
