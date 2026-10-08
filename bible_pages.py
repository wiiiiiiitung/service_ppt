"""
(book, chapter) → printed page number for the pew Bible used in the service.

The 經文 title slide carries a page hint — "(新約第7頁)" — that the operator
typed by hand every week. Every reference deck has one, so it is a required
part of the output, not an optional extra.

There is no formula for it: it depends on the physical edition. So this is a
lookup table that *learns*. It ships seeded with the values read off the
existing reference decks, and `remember()` adds each new confirmed entry. A
chapter the table has never seen returns None rather than a guess — the UI then
asks, exactly as before, and the answer is kept for next time.

Two seeded entries look wrong in the source decks and are kept as observed
rather than silently "corrected":
  * 馬太福音 6 → 7 but 馬太福音 7 → 4. A later chapter cannot be on an earlier
    page; one of the two was mistyped.
  * 詩篇 90 → 725 and 詩篇 146 → 725. The same page for both suggests 146 was
    copied from the previous week.
Verify those against the pew Bible and fix with `remember()` (or by editing the
JSON store); everything else is consistent.
"""

import json
import logging
import os

logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Writable store. /tmp on serverless, where the deployment filesystem is
# read-only; a repo-local data file when running normally.
STORE_PATH = os.environ.get("BIBLE_PAGES_PATH") or (
    os.path.join("/tmp", "bible_pages.json") if os.environ.get("VERCEL")
    else os.path.join(BASE_DIR, "data", "bible_pages.json")
)

# Harvested from example_2/*/Sunday Worship*.pptx — see the module docstring.
_SEED = {
    ("馬太福音", 6): 7,
    ("馬太福音", 7): 4,
    ("羅馬書", 6): 215,
    ("以斯帖記", 4): 620,
    ("以斯帖記", 5): 621,
    ("詩篇", 23): 673,
    ("詩篇", 90): 725,
    ("詩篇", 146): 725,
    ("箴言", 3): 770,
}

_cache = None


def _key(book, chapter):
    return f"{book}|{int(chapter)}"


def _load():
    global _cache
    if _cache is not None:
        return _cache
    _cache = {_key(b, c): p for (b, c), p in _SEED.items()}
    try:
        with open(STORE_PATH, encoding="utf-8") as fh:
            _cache.update(json.load(fh))
    except FileNotFoundError:
        pass
    except Exception:
        logger.exception("Could not read bible page store %s; using seeds only",
                         STORE_PATH)
    return _cache


def lookup(ref):
    """
    Page number for a scripture reference string, or None if not known.

    `ref` is the agenda's own text, e.g. "馬太福音 6:25~34" or "詩篇 90".
    """
    from bible_fetcher import parse_reference

    parsed = parse_reference(ref)
    if not parsed:
        return None
    book, chapter = parsed[0], parsed[1]
    return _load().get(_key(book, chapter))


def remember(ref, page):
    """
    Record the page number an operator confirmed for a reference.

    Best-effort: a read-only filesystem (serverless) just means the value
    applies to this run and the operator types it again next week.
    """
    from bible_fetcher import parse_reference

    parsed = parse_reference(ref)
    if not parsed or page in (None, ""):
        return False
    try:
        page = int(page)
    except (TypeError, ValueError):
        return False

    table = _load()
    table[_key(parsed[0], parsed[1])] = page
    try:
        os.makedirs(os.path.dirname(STORE_PATH), exist_ok=True)
        with open(STORE_PATH, "w", encoding="utf-8") as fh:
            json.dump(table, fh, ensure_ascii=False, indent=2, sort_keys=True)
        return True
    except Exception:
        logger.warning("Could not persist bible page for %s to %s", ref, STORE_PATH)
        return False
