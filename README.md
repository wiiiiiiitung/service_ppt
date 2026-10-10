# Sunday Worship PPT Generator

Flask app that assembles the Taiwanese-language Sunday worship deck from the
week's inputs: the agenda PDF, the hymn / responsive-reading PPTX files, and
the anthem lyrics DOCX.

```
python3 -m pip install -r requirements.txt
python3 app.py            # http://localhost:5001
```

## How a deck is assembled

Three PPTX assets, with distinct roles (`deck_sources.py`):

| Asset | Role |
| --- | --- |
| `template/intro.pptx` | The fixed pre-worship slides. The output deck *starts as a copy of this file*, so it also supplies the slide masters and the `詩歌` / `Blank` layouts every generated slide uses. |
| `template/fixed.pptx` | Searched by text marker for every recurring slide: 宣召, 祈禱, 主禱文, 信仰告白, 奉獻, 聖餐, 頌榮, 祝禱, 默禱, website. |
| `example/*/output/*.pptx` | Library, searched for hymn / reading / anthem slides when the week's own input file is missing. |

`template/fixed.pptx` is pinned deliberately. It used to be whichever past
output deck `os.listdir` returned first, which meant editing an old week's
deck silently changed this week's output. If you replace it, run the
regression harness — it asserts every marker in `slide_finder.SLIDE_MARKERS`
still resolves to exactly one slide.

Pipeline: `pdf_parser` → `slide_planner` (agenda + files → ordered slide specs)
→ `ppt_builder` (dispatch) → `slide_copier` (copy existing slides) or
`slide_generators` (build new ones). `text_layout` does all text measurement;
`autofit` shrinks anything that would still overflow.

## Text measurement

Where slides break is arithmetic, and it is all in `text_layout.py`:

- **Widths are measured in em**, not characters (`display_width`). A CJK glyph
  is 1 em, Latin is 0.5. Counting `len()` measured a URL-heavy announcement as
  twice its real width.
- **`DEFAULT_LINE_HEIGHT = 1.36` is measured, not assumed.** Rendering the decks
  through LibreOffice and reading back text extents gives ~1.02in per line for
  54pt 標楷體. With the usual 1.2 the `詩歌` body box looks like it holds 7
  lines; it holds 6 — exactly the off-by-one that pushed the last lyric line
  off every slide.
- Box geometry lives in `service_config.SLIDE_STYLES`, measured from the
  `詩歌` layout rather than hard-coded.

Multi-line text is written with `styles.write_lines`, which emits real
`<a:br/>` elements. A literal `\n` assigned through `run.text` lands inside
`<a:t>`, where OOXML collapses it — PowerPoint renders no break at all.

## Typography rules

Three rules the generated text follows, all derived from the hand-adjusted
reference decks:

**Verse numbers align.** Continuation lines are indented to clear the verse
number, and the number field is padded to the width of the widest number *on
that slide*, so a slide holding verses 9 and 10 reads

```
9.  因為我在人的權下，        not    9. 因為我在人的權下，
    下；對這個說                        下；對這個說
10. 耶穌聽見就希奇                  10.耶穌聽見就希奇
```

Grouping uses the widest number in the whole passage, then each slide is
re-wrapped against its own — which can only shorten lines, never overflow.

**No orphaned punctuation** (`text_layout.NO_LINE_START` / `NO_LINE_END`). A
line never begins with a comma, full stop or closing bracket, and never ends
with an opening one. Without this, a verse could end with a line holding
nothing but `」`.

**Chinese clauses lose their space.** The bulletin writes Chinese clauses with
an ASCII comma, so `信箱, 可填好`, `11/15, 積極會員` and `1559, 微信` all lose
the space. What the comma follows doesn't matter, only what it introduces —
which keeps `Apt.3F, Flushing` and `密碼 626954; 電話 646-558-8656` intact.

## Line metrics

`LINE_PITCH` is 1.2 and `PARA_SPACING` is 0.2, both measured off rendered PDFs:
consecutive baselines are exactly 60.0pt for 50pt text in a plain text box, and
70.0pt in a 詩歌 body placeholder, which adds the master's 20% `spcBef`. Keeping
them separate matters — rolling both into one 1.36 factor *and* adding
paragraph spacing double-counted the gap, over-predicted multi-paragraph
heights by ~30%, and made the autofit pass shrink 宣召 from 66pt to 55pt and
信仰告白 from 50pt to 33pt.

`autofit.TOP_LEADING` covers a separate effect: a body placeholder inheriting
the 詩歌 master renders its first line about a full line below the top of the
box, so a 6-line body whose text is 438pt tall starts 59pt down a 455pt box and
runs off the slide even though the text alone fits. Plain text boxes have no
such offset.

Slides copied from `template/fixed.pptx` are never re-measured — that deck is
curated by hand and is reproduced exactly.

## Regression harness

Eleven weeks of real input/output pairs live in `example_2/`. The harness
generates each week and compares it to the hand-finished deck:

```
python3 tools/regress.py                  # structural diff + marker check
python3 tools/regress.py --render         # + LibreOffice render, off-slide text
python3 tools/regress.py --week 20260802 -v
python3 tools/regress.py --save-baseline
python3 tools/regress.py --render --check # non-zero exit if worse than baseline
```

Three checks, because they catch different failures:

- **structure** — per-slide normalised text, aligned with `difflib`. Catches
  dropped sections, wrong content, changed page splits.
- **overflow** — renders to PDF and reads every word's bounding box. Text
  outside the slide is text nobody sees; no amount of structural diffing finds
  it. The reference decks are rendered too and reported in the same column,
  because some overshoot is inherent to the design.
- **markers** — every fixed-slide marker resolves to exactly one slide.

Current state across 11 weeks: 0 slides with off-slide text, against 45 in the
hand-finished reference decks; mean text similarity 0.876.

Reading the structural diff: a difference is not automatically a defect. The
references contain hand omissions the generator correctly does not reproduce —
`20260816` drops the 306A closing hymn the agenda lists, and several 啟應文
slides there lost words when continuation lines were merged by deleting them.
The generator follows the agenda.

## Legacy .ppt / .doc input

**Yes, the app takes `.ppt` / `.doc` directly** (any case: `.PPT`, `.DOC` too).
They are converted on upload and matched to agenda items like any other file.

This was A/B tested on the nine fixture weeks that ship both a raw `.ppt` and a
hand re-saved `.pptx`, building each week twice:

| | from your `.pptx` re-saves | from the raw `.ppt` |
| --- | --- | --- |
| slide count | — | identical, all 9 weeks |
| slide text | — | identical, 8 of 9 weeks |
| off-slide text | 0 slides | 0 slides |

Two things to know.

**LibreOffice flattens inheritance.** A placeholder that inherited its box
comes out of conversion with an explicit copy of one, which — taken literally —
pins a hymn body to the source deck's 10in width inside a 13.33in slide. That
is the narrow, left-shifted text the geometry fix exists to prevent, so
`slide_copier._is_flattened_inheritance` recognises a box that carries no real
positioning (it restates the source layout/master, or is full-bleed in the
source) and lets the destination layout supply it instead. The converter also
emits a bare `<p:ph/>` for the body, which OOXML would read as index 0 and so
inherit the *title* box; `_resolve_ph_idxs` names it and `_stamp_ph_idx` writes
the conclusion into the copy. Those three together took the geometry difference
between the two paths from 1180 attributes to 171, none of which cause
overflow.

**One character was lost in ~11 weeks of material.** On 09/06 the conversion
turned 「你的上帝」 into 「你上帝」. Nothing in the app can detect that, and it
is LibreOffice's fidelity, not ours. If a week's text is critical, a manual
re-save is still marginally safer; for routine use the trade against re-saving
every file by hand is yours to make.

## Deployment

Conversion shells out to LibreOffice (`file_converter.py`). **This is why
`Dockerfile` exists**: Vercel's
`@vercel/python` runtime has no LibreOffice, so `find_soffice()` returned None,
conversion silently did nothing, and every legacy file had to be re-saved by
hand in PowerPoint first. Run the container image (or run locally, where
`soffice` is on PATH) and that step disappears. The upload endpoint now says so
explicitly when a conversion fails rather than failing quietly.

## Bible page numbers

The 經文 title slide carries a page hint — "(新約第7頁)" — specific to the pew
Bible, with no formula behind it. `bible_pages.py` is a lookup table that
learns: seeded from the existing reference decks, and every page the operator
confirms is written back (`BIBLE_PAGES_PATH`, default `data/bible_pages.json`).
A passage it has not seen returns nothing and the UI asks, as before.

Two seeded entries look wrong in the source decks and are kept as observed
rather than quietly "corrected" — 馬太福音 6 → p.7 but 馬太福音 7 → p.4, and
詩篇 90 and 詩篇 146 both → p.725. See the module docstring.
