"""
Locate the PPTX assets the generator builds from.

Three distinct roles, previously conflated into one "template":

  intro    `template/intro.pptx` — the fixed pre-worship slides. The output
           deck starts as a copy of this file, so it also supplies the slide
           masters and the 詩歌 / Blank layouts every generated slide uses.

  fixed    `template/fixed.pptx` — the deck searched by text marker for every
           recurring slide (宣召, 祈禱, 主禱文, 信仰告白, 奉獻, 聖餐, 頌榮,
           祝禱, 默禱, website). Pinning this to an explicit, versioned file
           keeps the generated deck stable; it used to be whichever past output
           `os.listdir` happened to return first, so edits to an old week's
           deck silently changed this week's output.

  library  past output decks — searched for hymn / reading / anthem slides when
           the week's own input file is missing or unmatched.
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "template")
EXAMPLE_DIR = os.path.join(BASE_DIR, "example")


def intro_path():
    """Path to the fixed-intro deck, or None if it isn't installed."""
    p = os.path.join(TEMPLATE_DIR, "intro.pptx")
    return p if os.path.exists(p) else None


def fixed_path():
    """
    Path to the deck holding the recurring service slides.

    Prefers the pinned `template/fixed.pptx`. Falls back to the most recent
    example output, chosen by a full sort of (date, filename) so the result is
    deterministic rather than filesystem-order dependent.
    """
    pinned = os.path.join(TEMPLATE_DIR, "fixed.pptx")
    if os.path.exists(pinned):
        return pinned
    candidates = _example_outputs()
    return candidates[-1] if candidates else None


def library_paths(exclude=()):
    """Past output decks to search for hymn/reading/anthem slides."""
    skip = {os.path.abspath(p) for p in exclude if p}
    return [p for p in _example_outputs() if os.path.abspath(p) not in skip]


def _example_outputs():
    """All example output decks, sorted oldest → newest, deterministically."""
    out = []
    if not os.path.isdir(EXAMPLE_DIR):
        return out
    for sub in sorted(os.listdir(EXAMPLE_DIR)):
        out_dir = os.path.join(EXAMPLE_DIR, sub, "output")
        if not os.path.isdir(out_dir):
            continue
        for name in sorted(os.listdir(out_dir)):
            if name.endswith(".pptx") and not name.startswith("~$"):
                out.append(os.path.join(out_dir, name))
    return out
