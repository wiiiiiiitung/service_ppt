"""Fetch 和合本 (Chinese Union Version) Bible passages from bible-api.com."""

import json
import re
import urllib.request
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeout


# Map abbreviations to full book names
_BOOK_ALIASES = {
    "尼希米": "尼希米記",
}

# New Testament books (for testament detection)
_NT_BOOKS = {
    "馬太福音", "馬可福音", "路加福音", "約翰福音", "使徒行傳",
    "羅馬書", "哥林多前書", "哥林多後書", "加拉太書", "以弗所書",
    "腓立比書", "歌羅西書", "帖撒羅尼迦前書", "帖撒羅尼迦後書",
    "提摩太前書", "提摩太後書", "提多書", "腓利門書", "希伯來書",
    "雅各書", "彼得前書", "彼得後書", "約翰一書", "約翰二書",
    "約翰三書", "猶大書", "啟示錄",
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


def get_testament(book_name):
    """Return '新約' or '舊約' based on book name."""
    return "新約" if book_name in _NT_BOOKS else "舊約"


def _do_fetch(url):
    """Helper: perform the actual HTTP fetch (blocking call for ThreadPoolExecutor)."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read())
        verses = data.get("verses", [])
        return [{"verse": v["verse"], "text": v["text"].strip()} for v in verses]
    except Exception:
        return None


def fetch_verses(ref, timeout=5):
    """
    Fetch CUV (和合本) verses for a scripture reference string (non-blocking).

    Returns list of dicts: [{verse: int, text: str}, ...] or None on failure/timeout.
    """
    parsed = parse_reference(ref)
    if not parsed:
        return None

    book, chapter, v_start, v_end = parsed
    book_name = _BOOK_ALIASES.get(book, book)

    query = f"{book_name}{chapter}:{v_start}-{v_end}"
    url = f"{API_BASE}/{urllib.parse.quote(query)}?translation=cuv"

    with ThreadPoolExecutor(max_workers=1) as pool:
        future = pool.submit(_do_fetch, url)
        try:
            return future.result(timeout=timeout)
        except (FuturesTimeout, Exception):
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
