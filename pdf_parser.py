"""Parse Sunday service agenda PDF to extract worship order, sermon outline, and announcements."""

import re
import pdfplumber


def parse_agenda(pdf_path):
    """
    Parse the service agenda PDF and return structured data.

    Returns a dict with keys:
        date, worship_order, sermon_outline, announcements
    """
    with pdfplumber.open(pdf_path) as pdf:
        pages = [p.extract_text() or "" for p in pdf.pages]

    result = {
        "date": _extract_date(pages),
        "worship_order": _extract_worship_order(pages[2] if len(pages) > 2 else ""),
        "sermon_outline": _extract_sermon_outline(pages[3] if len(pages) > 3 else ""),
        "announcements": _extract_announcements(pages[0] if pages else ""),
    }
    return result


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


def _extract_worship_order(text):
    """
    Parse the worship order page (page 3).

    Returns a list of dicts with keys: type, number, title, presenter
    """
    items = []
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Find section start
    start = 0
    for i, line in enumerate(lines):
        if "主日敬拜程序" in line:
            start = i + 1
            break

    # Find section end (服事人員表)
    end = len(lines)
    for i, line in enumerate(lines[start:], start):
        if "主日服事人員表" in line or "耶和華在祂的聖殿" in line:
            end = i
            break

    order_lines = lines[start:end]

    for line in order_lines:
        # Skip the quote lines and separator lines
        if line.startswith("「") or line.startswith("」") or line.startswith("萬軍") or line.startswith("我，是否"):
            continue

        item = _parse_order_line(line)
        if item:
            items.append(item)

    return items


def _parse_order_line(line):
    """Parse a single worship order line into structured data."""

    # Known presenter suffixes — match from the end of the line
    PRESENTER_RE = (
        r"(司會/牧師|司會/會眾|眾立|會眾|司會|大詩班|陳牧師|"
        r"詩班[A-Za-z\u4e00-\u9fff]組?|"
        r"[\u4e00-\u9fff]{2,6}牧師|[\u4e00-\u9fff]{2,6}執事|[\u4e00-\u9fff]{2,6}長老)"
    )

    def split_presenter(body):
        """Split 'title presenter' into (title, presenter) using trailing presenter."""
        m = re.search(r"\s+" + PRESENTER_RE + r"$", body)
        if m:
            return body[:m.start()].strip(), m.group(1).strip()
        return body.strip(), ""

    # 聖詩 69 我的心神你當唱歌 眾立
    m = re.match(r"^聖詩\s+(\d+)\s+(.+)$", line)
    if m:
        title, presenter = split_presenter(m.group(2))
        return {"type": "hymn", "number": int(m.group(1)), "title": title, "presenter": presenter}

    # 啟應文 30 箴言 8 司會/會眾
    m = re.match(r"^啟應文\s+(\d+)\s+(.+)$", line)
    if m:
        title, presenter = split_presenter(m.group(2))
        return {"type": "responsive", "number": int(m.group(1)), "title": title, "presenter": presenter}

    # 頌榮 507 榮光歸聖父上帝 眾立
    m = re.match(r"^頌榮\s+(\d+)\s+(.+)$", line)
    if m:
        title, presenter = split_presenter(m.group(2))
        return {"type": "doxology", "number": int(m.group(1)), "title": title, "presenter": presenter}

    # 獻詩 耶和華是我牧者 大詩班
    m = re.match(r"^獻詩\s+(.+)$", line)
    if m:
        title, presenter = split_presenter(m.group(1))
        return {"type": "anthem", "number": None, "title": title, "presenter": presenter}

    # 經文 羅馬書 4:1~12 司會
    m = re.match(r"^經文\s+(.+)$", line)
    if m:
        title, presenter = split_presenter(m.group(1))
        return {"type": "scripture", "number": None, "title": title, "presenter": presenter}

    # 證道 亞伯拉罕 因信而被稱為義 陳信銘牧師
    m = re.match(r"^證道\s+(.+)$", line)
    if m:
        body = m.group(1)
        pm = re.search(r"\s+([\u4e00-\u9fff]{2,6}牧師|[\u4e00-\u9fff]{2,6}執事|[\u4e00-\u9fff]{2,6}長老)$", body)
        if pm:
            return {"type": "sermon", "number": None, "title": body[:pm.start()].strip(), "presenter": pm.group(1).strip()}
        return {"type": "sermon", "number": None, "title": body.strip(), "presenter": ""}

    # 奉獻 我的錢銀獻給你 眾立
    m = re.match(r"^奉獻\s+(.+)$", line)
    if m:
        title, presenter = split_presenter(m.group(1))
        return {"type": "offering", "number": None, "title": title, "presenter": presenter}

    # Simple items: 宣召, 祈禱及主禱文, 信仰告白, 報告, 祝禱
    m = re.match(r"^(宣召|祈禱及主禱文|信仰告白|報告|祝禱)\s*(.*?)$", line)
    if m:
        return {"type": _map_type(m.group(1)), "number": None, "title": m.group(1).strip(), "presenter": m.group(2).strip()}

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


def _extract_sermon_outline(text):
    """Parse the sermon outline page (page 4)."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Find the title block
    title = ""
    scripture = ""
    main_points = []

    # Skip header
    start = 0
    for i, line in enumerate(lines):
        if "講台綱要" in line or re.match(r"\d{2}/\d{2}/\d{4}", line):
            start = i + 1
            continue
        if "華語翻譯" in line or "號碼:" in line or "請自帶" in line:
            break

        # Scripture reference: 《...》
        if line.startswith("《") and "》" in line:
            scripture = line.strip("《》")
            continue

        # Check if it's a title (comes right after the date line)
        if not title and start > 0 and not line.startswith(("一", "二", "三", "1.", "2.", "3.")):
            if not re.match(r"\d{2}/\d{2}/\d{4}", line):
                title = line
                continue

        # Main points: 一. 二. 三.
        if re.match(r"^[一二三四五]\.", line):
            main_points.append({"heading": line, "points": []})
            continue

        # Sub-points: 1. 2. 3.
        if re.match(r"^\d+\.", line) and main_points:
            main_points[-1]["points"].append(line)
            continue

        # Continuation lines (wrap onto previous point or heading)
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
    """Parse the announcements page (page 1)."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    sections = {}
    current_section = None
    current_items = []

    for line in lines:
        if "報告事項" in line:
            continue
        if "出席及奉獻" in line or "主日崇拜人數" in line:
            break

        # Section headers: ※ 全教會, ※ 台語部
        if line.startswith("※") or re.match(r"^※\s+\S+", line):
            if current_section and current_items:
                sections[current_section] = current_items
            current_section = re.sub(r"^※\s*", "", line).strip()
            current_items = []
            continue

        # Numbered items
        if re.match(r"^\d+[.．]", line) and current_section is not None:
            current_items.append(line)
        elif current_items and current_section is not None:
            # Continuation of previous item
            current_items[-1] = current_items[-1] + " " + line

    if current_section and current_items:
        sections[current_section] = current_items

    return sections
