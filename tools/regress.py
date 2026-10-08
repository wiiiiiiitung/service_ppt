#!/usr/bin/env python3
"""
Regression harness: generate each fixture week and compare it to the
hand-finished reference deck.

Three independent checks, because they catch different failures:

  structure  Per-slide normalised text, aligned with difflib against the
             reference. Catches dropped sections, wrong content and changed
             page splits.
  overflow   Renders the deck with LibreOffice and reads back every word's
             bounding box. Any word outside the slide is text the operator will
             never see — this is what "missing words" actually looks like, and
             no amount of structural diffing finds it.
  markers    Every fixed-slide marker resolves to exactly one slide in
             template/fixed.pptx, so swapping that deck can't silently
             re-point a section.

Usage:
    python3 tools/regress.py                    # structure + markers
    python3 tools/regress.py --render           # add the overflow check
    python3 tools/regress.py --week 20260802    # one week
    python3 tools/regress.py --save-baseline    # record current numbers
    python3 tools/regress.py --render --check   # fail if worse than baseline
    python3 tools/regress.py -v                 # list every differing slide
"""

import argparse
import difflib
import io
import json
import contextlib
import logging
import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pptx import Presentation

import deck_sources

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK_ROOT = os.path.join(
    os.environ.get("REGRESS_WORK", os.path.join(BASE_DIR, ".regress")))
BASELINE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "regress_baseline.json")
# Slack allowed when checking whether a word falls off the slide, in points.
# Kept tight on purpose: the reference decks are rendered too and reported in
# the same column, so a small overshoot the design itself has shows up on both
# sides rather than being hidden by a generous threshold.
OVERFLOW_MARGIN_PT = 2.0


def normalise(text):
    """Slide text reduced to what a reader would notice: characters, in order."""
    return re.sub(r"\s+", "", text.replace("\x0b", "\n"))


def slide_signatures(path):
    prs = Presentation(path)
    return [
        normalise("".join(sh.text_frame.text for sh in slide.shapes
                          if sh.has_text_frame))
        for slide in prs.slides
    ]


def compare(ref_path, gen_path):
    ref, gen = slide_signatures(ref_path), slide_signatures(gen_path)
    sm = difflib.SequenceMatcher(None, ref, gen, autojunk=False)
    ref_only, gen_only = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ("delete", "replace"):
            ref_only += [(i, ref[i]) for i in range(i1, i2)]
        if tag in ("insert", "replace"):
            gen_only += [(j, gen[j]) for j in range(j1, j2)]
    return {
        "ref_slides": len(ref),
        "gen_slides": len(gen),
        "ref_only": ref_only,
        "gen_only": gen_only,
        "similarity": round(sm.ratio(), 4),
    }


def render_pdf(pptx_path, out_dir):
    """Render a deck to PDF with LibreOffice. Returns the path, or None."""
    soffice = None
    for cand in ("soffice", "libreoffice"):
        from shutil import which
        soffice = which(cand)
        if soffice:
            break
    if not soffice:
        return None
    os.makedirs(out_dir, exist_ok=True)
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir,
             pptx_path],
            check=True, capture_output=True, timeout=600,
        )
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
        print(f"    ! render failed: {e}")
        return None
    pdf = os.path.join(
        out_dir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    return pdf if os.path.exists(pdf) else None


def find_overflow(pdf_path):
    """
    Text that falls outside the page, one entry per (slide, side).

    Every side is reported rather than just the worst on a slide: the same
    slide can run past the right edge *and* poke above the top, and comparing
    generated against reference only means something if both are counted the
    same way.

    Returns [(page_index, side, overshoot_pt, sample_text), ...].
    """
    import pdfplumber

    bad = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            words = page.extract_words()
            if not words:
                continue
            worst = {}
            for w in words:
                for side, over in (
                    ("right", w["x1"] - page.width),
                    ("bottom", w["bottom"] - page.height),
                    ("left", -w["x0"]),
                    ("top", -w["top"]),
                ):
                    if over > OVERFLOW_MARGIN_PT:
                        prev = worst.get(side)
                        if prev is None or over > prev[0]:
                            worst[side] = (round(over, 1), w["text"][:18])
            for side in ("bottom", "right", "top", "left"):
                if side in worst:
                    bad.append((i, side, worst[side][0], worst[side][1]))
    return bad


def reference_overflow(week, ref_path, cache_dir):
    """
    Off-slide text in the hand-finished reference deck for the same week.

    Reported next to the generated figure because some of it is inherent to
    the deck design — the 詩歌 title placeholder sits slightly above the slide
    edge — and the number that matters is whether the generator is better or
    worse than what the operator produces by hand.
    """
    os.makedirs(cache_dir, exist_ok=True)
    pdf = os.path.join(cache_dir, f"REF {week}.pdf")
    if not os.path.exists(pdf) or os.path.getmtime(pdf) < os.path.getmtime(ref_path):
        import shutil
        staged = os.path.join(cache_dir, f"REF {week}.pptx")
        shutil.copy2(ref_path, staged)
        rendered = render_pdf(staged, cache_dir)
        if not rendered:
            return None
        if rendered != pdf:
            os.replace(rendered, pdf)
    return len(find_overflow(pdf))


def generate(week, work_dir):
    """Run the real pipeline over a fixture week. Returns the build report."""
    from fixtures import stage_inputs

    from pdf_parser import parse_agenda
    from ppt_builder import build_pptx

    pdf_path, input_files = stage_inputs(week, work_dir)
    if not pdf_path:
        raise RuntimeError("no agenda PDF in fixture")

    agenda = parse_agenda(pdf_path)
    out_path = os.path.join(work_dir, f"GEN {week}.pptx")
    report = build_pptx(
        deck_sources.fixed_path(), agenda, input_files, out_path,
        library_paths=deck_sources.library_paths(
            exclude=[deck_sources.fixed_path()]),
        intro_path=deck_sources.intro_path(),
        bible_page=None, overrides=None,
    )
    report["agenda"] = agenda
    return report


def check_markers():
    from slide_finder import ambiguous_markers

    path = deck_sources.fixed_path()
    if not path:
        return {"error": "no fixed deck found"}
    bad = ambiguous_markers(Presentation(path))
    return {"deck": os.path.basename(path), "ambiguous": bad}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--week", action="append", help="limit to these fixture weeks")
    ap.add_argument("--render", action="store_true",
                    help="render with LibreOffice and check for off-slide text")
    ap.add_argument("--save-baseline", action="store_true")
    ap.add_argument("--check", action="store_true",
                    help="exit non-zero if any week is worse than the baseline")
    ap.add_argument("-v", "--verbose", action="store_true",
                    help="list every differing slide")
    args = ap.parse_args()

    logging.basicConfig(level=logging.ERROR,
                        format="    ! %(levelname)s %(name)s: %(message)s")

    sys.path.insert(0, os.path.join(BASE_DIR, "tools"))
    from fixtures import reference_deck, weeks

    marker_report = check_markers()
    print(f"fixed deck: {marker_report.get('deck')}")
    if marker_report.get("ambiguous"):
        print(f"  ! ambiguous markers: {marker_report['ambiguous']}")
    elif not marker_report.get("error"):
        print("  markers: all resolve uniquely")
    print()

    targets = args.week or weeks()
    results = {}

    hdr = f"{'week':10} {'ref':>4} {'gen':>4} {'sim':>6} {'ref-only':>9} {'gen-only':>9}"
    if args.render:
        hdr += f" {'overflow':>16}"
    print(hdr)
    print("-" * len(hdr))

    for week in targets:
        work = os.path.join(WORK_ROOT, week)
        try:
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                report = generate(week, work)
        except Exception as e:
            print(f"{week:10} ERROR {type(e).__name__}: {e}")
            results[week] = {"error": str(e)}
            continue

        cmp = compare(reference_deck(week), report["path"])
        row = {
            "ref_slides": cmp["ref_slides"],
            "gen_slides": cmp["gen_slides"],
            "similarity": cmp["similarity"],
            "ref_only": len(cmp["ref_only"]),
            "gen_only": len(cmp["gen_only"]),
            "unused_inputs": report["unused_inputs"],
        }

        line = (f"{week:10} {cmp['ref_slides']:>4} {cmp['gen_slides']:>4} "
                f"{cmp['similarity']:>6.3f} {len(cmp['ref_only']):>9} "
                f"{len(cmp['gen_only']):>9}")

        if args.render:
            pdf = render_pdf(report["path"], os.path.join(work, "render"))
            over = find_overflow(pdf) if pdf else None
            ref_over = reference_overflow(
                week, reference_deck(week), os.path.join(WORK_ROOT, "_ref"))
            row["overflow"] = len(over) if over is not None else None
            row["ref_overflow"] = ref_over
            shown = "-" if over is None else str(len(over))
            line += f" {shown + ' (ref ' + str(ref_over) + ')':>16}"
            if over:
                row["overflow_detail"] = over[:10]

        print(line)
        if report["unused_inputs"]:
            print(f"           ! unused input files: {report['unused_inputs']}")
        if args.verbose:
            for i, t in cmp["ref_only"]:
                print(f"           REF-ONLY [{i:3d}] {t[:60]}")
            for j, t in cmp["gen_only"]:
                print(f"           GEN-ONLY [{j:3d}] {t[:60]}")
            for item in row.get("overflow_detail", []):
                print(f"           OVERFLOW slide {item[0]:3d} {item[1]:>6} "
                      f"by {item[2]}pt  {item[3]!r}")
        results[week] = row

    print()
    _summarise(results, args.render)

    if args.save_baseline:
        with open(BASELINE_PATH, "w", encoding="utf-8") as fh:
            json.dump(results, fh, ensure_ascii=False, indent=2, sort_keys=True)
        print(f"baseline written to {BASELINE_PATH}")

    if args.check:
        return _check_against_baseline(results)
    return 0


def _summarise(results, rendered):
    ok = [r for r in results.values() if "error" not in r]
    if not ok:
        return
    print(f"{len(ok)} weeks: "
          f"mean similarity {sum(r['similarity'] for r in ok) / len(ok):.3f}, "
          f"{sum(r['ref_only'] for r in ok)} ref-only slides, "
          f"{sum(r['gen_only'] for r in ok)} gen-only slides, "
          f"{sum(len(r['unused_inputs']) for r in ok)} unused inputs")
    if rendered:
        counts = [r["overflow"] for r in ok if r.get("overflow") is not None]
        ref_counts = [r["ref_overflow"] for r in ok if r.get("ref_overflow") is not None]
        if counts:
            msg = f"  off-slide text on {sum(counts)} generated slides"
            if ref_counts:
                msg += f", vs {sum(ref_counts)} in the reference decks"
            print(msg)


def _check_against_baseline(results):
    try:
        with open(BASELINE_PATH, encoding="utf-8") as fh:
            base = json.load(fh)
    except FileNotFoundError:
        print("no baseline recorded; run with --save-baseline first")
        return 0

    regressions = []
    for week, row in results.items():
        prev = base.get(week)
        if not prev or "error" in row:
            continue
        if row["similarity"] < prev["similarity"] - 0.001:
            regressions.append(f"{week}: similarity {prev['similarity']} → {row['similarity']}")
        if (row.get("overflow") is not None and prev.get("overflow") is not None
                and row["overflow"] > prev["overflow"]):
            regressions.append(f"{week}: overflow {prev['overflow']} → {row['overflow']}")
        if len(row["unused_inputs"]) > len(prev.get("unused_inputs", [])):
            regressions.append(f"{week}: unused inputs grew")

    if regressions:
        print("REGRESSIONS:")
        for r in regressions:
            print(f"  - {r}")
        return 1
    print("no regressions against baseline")
    return 0


if __name__ == "__main__":
    sys.exit(main())
