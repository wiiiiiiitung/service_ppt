"""Parse Sunday service agenda PDF to extract worship order, sermon outline, and announcements."""

import re

import pdfplumber

from service_config import ANNOUNCEMENT_STRIP_CJK_SPACES

# Matches CJK ideographs and full-width punctuation. Used to strip spurious
# PDF-extraction spaces that sit next to Chinese text (see
# service_config.ANNOUNCEMENT_STRIP_CJK_SPACES).
_CJK_RE = r"[　-〿㐀-䶿一-鿿豈-﫿＀-￯]"

# Fallback presenter vocabulary, used only when the agenda's right-hand column
# can't be read off the page geometry. The primary mechanism is
# `_presenter_vocabulary`, which reads the actual column — the old hard-coded
# list quietly failed on anything new ("Elmhurst分區" ended up glued to an
# anthem title).
_PRESENTER_RE = re.compile(
    r"(司會/牧師|司會/會眾|眾立|會眾|司會|大詩班|陳牧師|"
    r"詩班[A-Za-z一-鿿]組?|"
    r"[一-鿿]{2,6}牧師|[一-鿿]{2,6}執事|[一-鿿]{2,6}長老)"
)

# A hymn number may carry a letter suffix (254A, 25A, 21B, 189B). Kept as a
# string throughout: parsing it as an int dropped the suffix and, before that,
# made the whole line fail to match, silently losing the entire hymn section.
_NUM = r"(\d+[A-Za-z]?)"

# Geometry thresholds for reading the presenter column (in PDF points).
_PRESENTER_MIN_GAP = 40      # blank space separating title column from presenter
_PRESENTER_MIN_X_FRAC = 0.70  # presenter starts in the right-hand 30% of the row


def parse_agenda(pdf_path):
    """
    Parse the service agenda PDF and return structured data.

    Returns a dict with keys:
        date, worship_order, sermon_outline, announcements
    """
    with pdfplumber.open(pdf_path) as pdf:
        pages = [p.extract_text() or "" for p in pdf.pages]
        order_page = next((p for p in pdf.pages
                           if "主日敬拜程序" in (p.extract_text() or "")), None)
        presenters = _presenter_vocabulary(order_page) if order_page else set()

    return {
        "date": _extract_date(pages),
        "worship_order": _extract_worship_order(
            _find_page(pages, "主日敬拜程序"), presenters),
        "sermon_outline": _extract_sermon_outline(_find_page(pages, "講台綱要")),
        "announcements": _extract_announcements(_find_page(pages, "報告事項")),
    }


def _find_page(pages, marker):
    """Return the first page containing the marker text, or empty string if not found."""
    return next((p for p in pages if marker in p), "")


def _presenter_vocabulary(page):
    """
    Read the agenda's right-hand presenter column off the page geometry.

    Every worship-order row is right-aligned to the same edge, with a wide gap
    between the item title and the presenter. Walking each row from the right
    and taking the contiguous run of words that sits in the right-hand portion
    gives the presenter strings actually used this week, whatever they are —
    no hard-coded list to fall behind.
    """
    vocab = set()
    try:
        words = page.extract_words(x_tolerance=1.5)
    except Exception:
        return vocab

    rows = {}
    for w in words:
        rows.setdefault(round(w["top"] / 3), []).append(w)

    for row in rows.values():
        row = sorted(row, key=lambda w: w["x0"])
        if len(row) < 2:
            continue
        right_edge = row[-1]["x1"]
        picked = [row[-1]]
        for i in range(len(row) - 1, 0, -1):
            gap = row[i]["x0"] - row[i - 1]["x1"]
            if gap > _PRESENTER_MIN_GAP:
                break
            picked.insert(0, row[i - 1])
        if picked[0]["x0"] < _PRESENTER_MIN_X_FRAC * right_edge:
            continue
        text = "".join(w["text"] for w in picked).strip()
        if text and len(text) <= 12 and not re.fullmatch(r"[\d\s./:~-]+", text):
            vocab.add(text)
            vocab.add(" ".join(w["text"] for w in picked).strip())
    return vocab


def _extract_worship_order(text, presenters=()):
    """
    Parse the worship order page.

    Returns a list of dicts with keys: type, number, title, presenter
    """
    items = []
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    start = 0
    for i, line in enumerate(lines):
        if "主日敬拜程序" in line:
            start = i + 1
            break

    end = len(lines)
    for i, line in enumerate(lines[start:], start):
        if "主日服事人員表" in line or "耶和華在祂的聖殿" in line:
            end = i
            break

    for line in lines[start:end]:
        # Scripture epigraphs wrap across lines and never parse as order items;
        # skipping them explicitly keeps the DROP diagnostics meaningful.
        if _is_epigraph(line):
            continue
        item = _parse_order_line(line, presenters)
        if item:
            items.append(item)

    return items


def _is_epigraph(line):
    """True for the scripture quotation block above the order, including wraps."""
    if line.startswith(("「", "」", "『", "』")):
        return True
    # A wrapped continuation ends the quote and cites the reference.
    return bool(re.search(r"[」』]\s*[（(][^)）]*\d+\s*[:：]\s*\d+", line))


def _parse_order_line(line, presenters=()):
    """Parse a single worship order line into structured data."""

    def split_presenter(body):
        """Split 'title presenter' into (title, presenter)."""
        body = body.strip()
        # Prefer the presenters actually read off this agenda's right column,
        # longest first so "司會/牧師" wins over "牧師".
        for cand in sorted(presenters, key=len, reverse=True):
            if not cand:
                continue
            if body.endswith(cand) and len(body) > len(cand):
                return body[: -len(cand)].strip(), cand
        m = re.search(r"\s+" + _PRESENTER_RE.pattern + r"$", body)
        if m:
            return body[: m.start()].strip(), m.group(1).strip()
        return body, ""

    def numbered(kind, keyword):
        m = re.match(rf"^{keyword}\s+{_NUM}\s+(.+)$", line)
        if not m:
            return None
        title, presenter = split_presenter(m.group(2))
        return {"type": kind, "number": m.group(1), "title": title,
                "presenter": presenter}

    def unnumbered(kind, keyword):
        m = re.match(rf"^{keyword}\s+(.+)$", line)
        if not m:
            return None
        title, presenter = split_presenter(m.group(1))
        return {"type": kind, "number": None, "title": title,
                "presenter": presenter}

    # 聖詩 71 我心讃美至高上帝 眾立   /   聖詩 254A 我認救主無驚見誚 會眾
    for kind, keyword in (("hymn", "聖詩"), ("responsive", "啟應文"),
                          ("doxology", "頌榮")):
        item = numbered(kind, keyword)
        if item:
            return item

    # 聖餐 215 耶穌身軀替咱釘死 (第1節) 牧師 — the number and the verse note are
    # both optional, and the line was previously not recognised at all, which
    # dropped the whole communion section on communion Sundays.
    m = re.match(rf"^聖餐(?:\s+{_NUM})?\s*(.*)$", line)
    if m and (m.group(1) or m.group(2)):
        rest, presenter = split_presenter(m.group(2) or "")
        verse = ""
        vm = re.search(r"[（(]\s*第\s*([\d,~\-、]+)\s*節\s*[）)]", rest)
        if vm:
            verse = vm.group(1)
            rest = rest[: vm.start()].strip()
        return {"type": "communion", "number": m.group(1), "title": rest,
                "presenter": presenter, "verse": verse}

    for kind, keyword in (("anthem", "獻詩"), ("scripture", "經文"),
                          ("sermon", "證道"), ("offering", "奉獻")):
        item = unnumbered(kind, keyword)
        if item:
            return item

    # Simple items: 宣召, 祈禱及主禱文, 信仰告白, 報告, 祝禱
    m = re.match(r"^(宣召|祈禱及主禱文|信仰告白|報告|祝禱)\s*(.*?)$", line)
    if m:
        return {"type": _map_type(m.group(1)), "number": None,
                "title": m.group(1).strip(), "presenter": m.group(2).strip()}

    return None


def _map_type(keyword):
    mapping = {
        "宣召": "call_to_worship",
        "祈禱及主禱文": "prayer",
        "信仰告白": "creed",
        "報告": "announcements",
        "祝禱": "benediction",
    }
    return mapping.get(keyword, keyword)


def _extract_date(pages):
    """Extract service date from PDF (format: MM/DD/YYYY)."""
    for page in pages:
        m = re.search(r"(\d{2}/\d{2}/\d{4})", page)
        if m:
            return m.group(1)
        m = re.search(r"NO\.\s*\d+\s+(\d{2}/\d{2}/\d{4})", page)
        if m:
            return m.group(1)
    return ""


def _extract_sermon_outline(text):
    """Parse the sermon outline page."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    title = ""
    scripture = ""
    main_points = []

    start = 0
    for i, line in enumerate(lines):
        if "講台綱要" in line or re.match(r"\d{2}/\d{2}/\d{4}", line):
            start = i + 1
            continue
        if "華語翻譯" in line or "號碼:" in line or "請自帶" in line:
            break

        if line.startswith("《") and "》" in line:
            scripture = line.strip("《》")
            continue

        if not title and start > 0 and not line.startswith(("一", "二", "三", "1.", "2.", "3.")):
            if not re.match(r"\d{2}/\d{2}/\d{4}", line):
                title = line
                continue

        if re.match(r"^[一二三四五]\.", line):
            main_points.append({"heading": line, "points": []})
            continue

        if re.match(r"^\d+\.", line) and main_points:
            main_points[-1]["points"].append(line)
            continue

        elif main_points:
            if main_points[-1]["points"]:
                main_points[-1]["points"][-1] += line
            else:
                main_points[-1]["heading"] += line

    return {
        "title": title,
        "scripture": scripture,
        "main_points": main_points,
    }


def _extract_announcements(text):
    """Parse the announcements page."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    sections = {}
    current_section = None
    current_items = []

    for line in lines:
        if "報告事項" in line:
            continue
        if "出席及奉獻" in line or "主日崇拜人數" in line:
            break

        if line.startswith("※") or re.match(r"^※\s+\S+", line):
            if current_section and current_items:
                sections[current_section] = current_items
            current_section = re.sub(r"^※\s*", "", line).strip()
            current_items = []
            continue

        if re.match(r"^\d+[.．]", line) and current_section is not None:
            current_items.append(line)
        elif current_items and current_section is not None:
            current_items[-1] = current_items[-1] + " " + line

    if current_section and current_items:
        sections[current_section] = current_items

    if ANNOUNCEMENT_STRIP_CJK_SPACES:
        sections = {
            section: [_strip_cjk_spacing(item) for item in items]
            for section, items in sections.items()
        }

    return sections


_ITEM_PREFIX_RE = re.compile(r"^(\d+[.．])\s*")


def _strip_cjk_spacing(text):
    """
    Collapse whitespace sitting *between two CJK characters*.

    The bulletin is justified, so pdfplumber reports spurious spaces at every
    line-wrap point in Chinese text. Removing those is right; removing a space
    that borders Latin text is not — "會號 994 2970 4298" became "會號994 2970
    4298" and "線上: https://…" lost its separator. The reference decks keep
    the CJK↔Latin spaces, so only the CJK↔CJK ones are collapsed.
    """
    prefix = ""
    m = _ITEM_PREFIX_RE.match(text)
    if m:
        prefix = m.group(1) + " "
        text = text[m.end():]
    text = re.sub(rf"(?<={_CJK_RE})\s+(?={_CJK_RE})", "", text)
    return prefix + text.strip()
