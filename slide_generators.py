"""
Dynamic slide generation: create slides with content.

All styles/positions/fonts are read from service_config.SLIDE_STYLES.
Multi-line text is written through styles.write_lines, which emits real
`<a:br/>` breaks inside a single paragraph — a literal "\\n" pushed through
`run.text` lands inside `<a:t>`, where OOXML collapses it and PowerPoint
renders no break at all.
"""

from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.util import Emu, Inches, Pt

from bible_fetcher import get_testament
from service_config import SLIDE_STYLES
from slide_copier import get_layout
from styles import CYAN_BRIGHT, apply_font, write_lines
from text_layout import fit_size

_ALIGN = {
    "center": PP_PARAGRAPH_ALIGNMENT.CENTER,
    "left": PP_PARAGRAPH_ALIGNMENT.LEFT,
    "right": PP_PARAGRAPH_ALIGNMENT.RIGHT,
}


def _new_slide(out_prs, style, default_layout="Blank"):
    layout = get_layout(out_prs, style.get("layout", default_layout))
    return out_prs.slides.add_slide(layout)


def _add_box(slide, style, pos_key="pos", size_key="size", word_wrap=True):
    pos, size = style[pos_key], style[size_key]
    box = slide.shapes.add_textbox(pos[0], pos[1], size[0], size[1])
    box.text_frame.word_wrap = word_wrap
    return box


def _placeholders(slide):
    """Title and body placeholders of a 詩歌-layout slide, or (None, None)."""
    title = body = None
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0 and title is None:
            title = ph
        elif idx == 1 and body is None:
            body = ph
    return title, body


def add_placeholder_slide(out_prs, label):
    """Add a simple placeholder slide for missing content."""
    layout = get_layout(out_prs, "Blank")
    slide = out_prs.slides.add_slide(layout)
    tf = slide.shapes.add_textbox(
        Inches(1), Inches(2.5), Inches(11.33), Inches(2)
    ).text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run()
    run.text = f"[{label}]"
    run.font.size = Pt(32)
    run.font.bold = True
    run.font.color.rgb = CYAN_BRIGHT


def add_anthem_title_slide(out_prs, title):
    """Anthem title slide: '獻詩: <title>' centered."""
    style = SLIDE_STYLES["anthem_title"]
    slide = _new_slide(out_prs, style)
    tb = _add_box(slide, style, word_wrap=False)
    p = tb.text_frame.paragraphs[0]
    p.alignment = _ALIGN.get(style.get("align"), PP_PARAGRAPH_ALIGNMENT.CENTER)
    run = p.add_run()
    run.text = style.get("text", "獻詩: {title}").format(title=title)
    apply_font(run.font, style)


def add_scripture_title_slide(out_prs, item, bible_page=None):
    """Scripture title slide: centered 經文 / reference / page-hint lines."""
    style = SLIDE_STYLES["scripture_title"]
    slide = _new_slide(out_prs, style)
    tf = _add_box(slide, style).text_frame

    ref = item.get("title", "")
    testament = get_testament(ref.split()[0]) if ref else "新約"
    align = _ALIGN.get(style.get("align"))

    first = True
    for line_cfg in style.get("lines", []):
        runs_cfg = [rc for rc in line_cfg
                    if not (rc.get("page_only") and not bible_page)]
        if not runs_cfg:
            continue

        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        if align is not None:
            p.alignment = align

        for run_cfg in runs_cfg:
            run = p.add_run()
            run.text = run_cfg.get("text", "").format(
                ref=ref, testament=testament, page=bible_page or "")
            apply_font(run.font, run_cfg)


def add_scripture_verse_slide(out_prs, ref, verses):
    """
    Scripture verse slide: reference bar on top, verse body below.

    Each verse is one paragraph; its wrapped lines (pre-computed by the planner,
    with continuation lines indented to clear the verse number) are joined with
    real line breaks.
    """
    layout = get_layout(out_prs, "Blank")
    slide = out_prs.slides.add_slide(layout)

    bar_style = SLIDE_STYLES["scripture_verse_bar"]
    tb_ref = _add_box(slide, bar_style, word_wrap=False)
    p_ref = tb_ref.text_frame.paragraphs[0]
    p_ref.alignment = _ALIGN.get(bar_style.get("align"), PP_PARAGRAPH_ALIGNMENT.LEFT)
    run_ref = p_ref.add_run()
    run_ref.text = ref
    apply_font(run_ref.font, bar_style)

    body_style = SLIDE_STYLES["scripture_verse_body"]
    tf_body = _add_box(slide, body_style).text_frame
    align = _ALIGN.get(body_style.get("align"))

    first = True
    for v in verses:
        lines = v.get("lines") or [f"{v['verse']}.{v['text']}"]
        p = tf_body.paragraphs[0] if first else tf_body.add_paragraph()
        first = False
        if align is not None:
            p.alignment = align
        write_lines(p, lines, body_style)


def add_lyrics_slide(out_prs, title, lines):
    """Hymn/anthem lyrics slide using 詩歌 layout placeholders."""
    style = SLIDE_STYLES["lyrics"]
    slide = _new_slide(out_prs, style, default_layout="詩歌")
    title_ph, body_ph = _placeholders(slide)

    if title_ph is not None and body_ph is not None:
        title_ph.text_frame.text = title
        tf = body_ph.text_frame
        tf.clear()
        write_lines(tf.paragraphs[0], lines)
        return

    tb = slide.shapes.add_textbox(Emu(0), Emu(126609), Emu(12192000), Emu(436100))
    tb.text_frame.paragraphs[0].text = title
    tb2 = slide.shapes.add_textbox(Emu(0), Emu(745435), Emu(12192000), Emu(6112565))
    tb2.text_frame.word_wrap = True
    write_lines(tb2.text_frame.paragraphs[0], lines)


def add_sermon_title_slide(out_prs, spec):
    """Sermon title slide: '今日信息' + title + preacher."""
    style = SLIDE_STYLES["sermon_title"]
    slide = _new_slide(out_prs, style, default_layout="詩歌")
    title_ph, body_ph = _placeholders(slide)
    if title_ph is None or body_ph is None:
        return

    sermon_title = spec.get("title", "")
    preacher = spec.get("preacher", "")

    tf0 = title_ph.text_frame
    tf0.clear()
    run0 = tf0.paragraphs[0].add_run()
    run0.text = style.get("header_text", "今日信息")
    run0.font.name = "標楷體"
    run0.font.bold = True

    tf = body_ph.text_frame
    tf.clear()
    # The reference decks open with an empty paragraph, which drops the title
    # to the vertical middle of the box.
    if style.get("lead_blank_line"):
        tf.paragraphs[0].text = ""
        p_title = tf.add_paragraph()
    else:
        p_title = tf.paragraphs[0]

    run_t = p_title.add_run()
    run_t.text = sermon_title
    run_t.font.size = style.get("title_size_pt", Pt(80))

    if preacher:
        run_sp = p_title.add_run()
        run_sp.text = style.get("spacer_text", " " * 51)
        apply_font(run_sp.font, {
            "font": style.get("spacer_font"),
            "size_pt": style.get("spacer_size_pt"),
            "bold": style.get("spacer_bold"),
        })
        run_p = p_title.add_run()
        run_p.text = preacher
        apply_font(run_p.font, {
            "font": style.get("preacher_font"),
            "size_pt": style.get("preacher_size_pt"),
            "bold": style.get("preacher_bold"),
        })


def add_sermon_point_slide(out_prs, spec):
    """Sermon point slide: 今日信息 header, then heading and bullet points."""
    style = SLIDE_STYLES["sermon_point"]
    slide = _new_slide(out_prs, style, default_layout="詩歌")

    heading = spec.get("heading", "")
    points = list(spec.get("points", []))

    title_ph, body_ph = _placeholders(slide)
    if style.get("remove_title_ph") and title_ph is not None:
        title_ph._element.getparent().remove(title_ph._element)

    if body_ph is not None:
        if style.get("pos"):
            body_ph.left, body_ph.top = style["pos"]
        if style.get("size"):
            body_ph.width, body_ph.height = style["size"]
        tf = body_ph.text_frame
    else:
        tb = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(12192000), Emu(6858000))
        tb.text_frame.word_wrap = True
        tf = tb.text_frame
    tf.clear()

    p0 = tf.paragraphs[0]
    run0 = p0.add_run()
    run0.text = style.get("header_text", "今日信息")
    apply_font(run0.font, {
        "font": style.get("header_font"),
        "size_pt": style.get("header_size_pt"),
        "bold": style.get("header_bold", True),
        "color": style.get("header_color"),
    })

    # Heading and points share one paragraph, separated by line breaks, so the
    # master's 20%-of-a-line paragraph spacing doesn't eat the height budget.
    body_lines = ([heading] if heading else []) + points
    if body_lines:
        p1 = tf.add_paragraph()
        p1.alignment = _ALIGN.get(style.get("points_align"),
                                  PP_PARAGRAPH_ALIGNMENT.LEFT)
        write_lines(p1, body_lines)


def add_announcement_slide(out_prs, spec):
    """
    Announcement slide: one item per slide, font stepped down to fit.

    These decks don't enable PowerPoint's shrink-on-overflow autofit, so an
    item that doesn't fit at the layout's 54pt is set at the largest size from
    the ladder that does. The item is never split: the old fixed-length split
    produced tail slides holding a fragment and cut through English names.
    """
    style = SLIDE_STYLES["announcement_item"]
    slide = _new_slide(out_prs, style, default_layout="詩歌")
    title_ph, body_ph = _placeholders(slide)

    section = spec.get("section", "報告")
    text = spec.get("text", "")
    size, _ = fit_size(text, style["box"], style["size_ladder"])

    if title_ph is not None and body_ph is not None:
        title_ph.text_frame.text = section
        tf = body_ph.text_frame
        tf.clear()
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = _ALIGN.get(style.get("align"), PP_PARAGRAPH_ALIGNMENT.LEFT)
        run = p.add_run()
        run.text = text
        run.font.size = size
        return

    tb = slide.shapes.add_textbox(Emu(998220), Emu(-152400), Emu(9403080), Emu(1143000))
    tb.text_frame.paragraphs[0].text = section
    tb2 = slide.shapes.add_textbox(Emu(0), Emu(990600), Emu(12192000), Emu(5360964))
    tb2.text_frame.word_wrap = True
    p = tb2.text_frame.paragraphs[0]
    p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = size
