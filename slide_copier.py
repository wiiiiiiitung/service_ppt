"""
Low-level XML manipulation for copying slides from the fixed deck, the slide
library, and the week's own input decks into the output presentation.
"""

import copy
import logging
from collections import Counter

from lxml import etree
from pptx.oxml.ns import qn
from pptx.util import Emu

logger = logging.getLogger(__name__)

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_SHAPE_TAGS = (qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"),
               qn("p:cxnSp"), qn("p:grpSp"))

# PowerPoint writes idx="4294967295" on some title placeholders, and
# python-pptx reports it verbatim. Comparing that against a layout's indices
# never matches, so a title would be treated as an unmatched placeholder.
PH_IDX_SENTINEL = 0xFFFFFFFF
_TITLE_TYPES = ("TITLE", "CENTER_TITLE")
# LibreOffice stores geometry in 1/100 mm and rounds on the way out, so a box
# it "preserved" can be a few hundred EMU off the original. Compare with slack.
_GEOM_TOLERANCE = 3600  # EMU, ~0.004in


def placeholder_idx(shape):
    """
    Placeholder index of a shape, normalised, or None if it isn't one.

    Three normalisations, each for a real shape seen in this material:

    * A title placeholder answers 0. Its `idx` is optional, and PowerPoint
      sometimes writes the 32-bit sentinel instead of omitting it.
    * A bare `<p:ph/>` carrying neither `type` nor `idx` answers None. OOXML
      defaults such an index to 0, which would make it collide with the title
      — LibreOffice emits exactly this for the body placeholder when it
      converts a legacy `.ppt`. "Unknown" is the honest answer.
    * The sentinel on anything else answers None.
    """
    try:
        if not shape.is_placeholder:
            return None
        fmt = shape.placeholder_format
        idx, ph_type = fmt.idx, fmt.type
    except (AttributeError, ValueError):
        return None

    if ph_type is not None and str(ph_type).split()[0] in _TITLE_TYPES:
        return 0

    el = _ph_element(shape)
    if el is not None and el.get("type") is None and el.get("idx") is None:
        return None

    if idx == PH_IDX_SENTINEL:
        return None
    return idx


def _ph_element(shape):
    """The shape's raw <p:ph> element, or None."""
    return _ph_element_of(shape._element)


def copy_slide(out_prs, src_prs, index):
    """
    Copy a slide from src_prs at index into out_prs.

    Deep-copies the shape tree and background, re-points relationship ids at
    copies of the source's related parts (images, media, hyperlinks), and
    normalises geometry for the destination slide size.

    Args:
        out_prs: output Presentation
        src_prs: source Presentation
        index: 0-based slide index in src_prs

    Returns:
        The newly added slide, or None if index is out of range.
    """
    if index is None or index < 0 or index >= len(src_prs.slides):
        return None

    src_slide = src_prs.slides[index]

    # Use a matching layout if possible; otherwise the first one (詩歌).
    layout = get_layout(out_prs, src_slide.slide_layout.name)
    new_slide = out_prs.slides.add_slide(layout)

    src_sp_tree = src_slide._element.find(qn("p:cSld")).find(qn("p:spTree"))
    new_sp_tree = new_slide._element.find(qn("p:cSld")).find(qn("p:spTree"))

    # Drop the placeholders add_slide() cloned from the layout; the source's
    # own shapes replace them wholesale.
    for child in list(new_sp_tree):
        new_sp_tree.remove(child)

    layout_ph_idxs = _placeholder_idxs(layout)
    src_shapes = list(src_slide.shapes)
    # A LibreOffice-converted deck sometimes gives two shapes the same
    # placeholder index; dropping geometry from both would stack them, so
    # ambiguous indices keep whatever box they carry.
    resolved_idxs = _resolve_ph_idxs(src_shapes, layout_ph_idxs)
    counts = Counter(i for i in resolved_idxs if i is not None)
    ambiguous = {i for i, n in counts.items() if n > 1}

    shape_idx = 0
    for child in src_sp_tree:
        new_child = copy.deepcopy(child)
        if child.tag in _SHAPE_TAGS:
            src_shape = src_shapes[shape_idx] if shape_idx < len(src_shapes) else None
            ph_idx = resolved_idxs[shape_idx] if shape_idx < len(resolved_idxs) else None
            shape_idx += 1
            if src_shape is not None:
                _stamp_ph_idx(new_child, src_shape, ph_idx)
                _normalise_geometry(new_child, src_shape, ph_idx, layout_ph_idxs,
                                    src_slide.slide_layout, ambiguous,
                                    src_prs.slide_width, out_prs.slide_width)
        new_sp_tree.append(new_child)

    _copy_background(src_slide, new_slide)
    _remap_relationships(new_slide, src_slide)

    return new_slide


def _placeholder_idxs(layout):
    """Placeholder indices the destination layout can supply geometry for."""
    idxs = set()
    for ph in layout.placeholders:
        idx = placeholder_idx(ph)
        if idx is not None:
            idxs.add(idx)
    return idxs


def _stamp_ph_idx(new_el, src_shape, ph_idx):
    """
    Write a resolved placeholder index onto a copied bare `<p:ph/>`.

    `_resolve_ph_idxs` works out that a bare placeholder is the body, but that
    conclusion has to be recorded in the XML or the renderer re-derives the
    OOXML default of 0 and the shape inherits the *title* box — a body of
    lyrics laid out in a 1.25in-tall title strip.
    """
    if ph_idx is None:
        return
    el = _ph_element_of(new_el)
    if el is None:
        return
    if el.get("type") is None and el.get("idx") is None and ph_idx != 0:
        el.set("idx", str(ph_idx))


def _ph_element_of(sp_el):
    """The <p:ph> element inside a raw <p:sp>, or None."""
    nv_sp_pr = sp_el.find(qn("p:nvSpPr"))
    if nv_sp_pr is None:
        return None
    nv_pr = nv_sp_pr.find(qn("p:nvPr"))
    return None if nv_pr is None else nv_pr.find(qn("p:ph"))


def _resolve_ph_idxs(src_shapes, layout_ph_idxs):
    """
    Placeholder index per source shape, resolving the bare `<p:ph/>` case.

    A LibreOffice-converted slide carries `<p:ph type="title"/>` for the title
    and a completely bare `<p:ph/>` for the body. The bare one can't be read on
    its own — OOXML would default its index to 0 and collide with the title —
    but in the context of the slide it is unambiguous: the body is the bare
    placeholder on a slide that already has a title. Naming it here lets it
    inherit the destination layout's body box like any other body placeholder.
    """
    idxs = [placeholder_idx(sh) for sh in src_shapes]
    claimed = {i for i in idxs if i is not None}

    for n, shape in enumerate(src_shapes):
        if idxs[n] is not None:
            continue
        el = _ph_element(shape)
        if el is None or el.get("type") is not None or el.get("idx") is not None:
            continue
        if 0 in claimed and 1 in layout_ph_idxs and 1 not in claimed:
            idxs[n] = 1
            claimed.add(1)
    return idxs


def _inherited_box(layout, ph_idx):
    """
    The box a shape would inherit for `ph_idx`, or None.

    Checks the layout, then its master — inheritance runs slide → layout →
    master, and a LibreOffice-converted deck often has the title box only on
    the master, its generated "Default" layout carrying nothing but the
    footer, date and slide-number placeholders.
    """
    sources = [layout]
    try:
        sources.append(layout.slide_master)
    except AttributeError:
        pass

    for source in sources:
        try:
            placeholders = source.placeholders
        except AttributeError:
            continue
        for ph in placeholders:
            if placeholder_idx(ph) != ph_idx:
                continue
            box = (ph.left, ph.top, ph.width, ph.height)
            if None not in box:
                return box
    return None


def _is_flattened_inheritance(box, src_layout, ph_idx, src_width):
    """
    True when a placeholder's explicit box carries no real positioning.

    Converting a legacy `.ppt` flattens inheritance: a placeholder that had no
    box of its own comes out with an explicit copy of one. Taken literally it
    pins the shape to the source deck's 10in width inside a 13.33in slide,
    which is the narrow, left-shifted text the whole geometry fix exists to
    avoid. Two signatures of a flattened box, either of which is enough:

    * It restates what the source layout or master would have supplied. The
      comparison allows a few hundred EMU, because LibreOffice round-trips
      geometry through 1/100 mm.
    * It is full-bleed in the source deck — flush left and exactly the source
      slide's width. "Fills the slide" is the intent, and it should keep
      filling the slide at the destination size. A box that merely overhangs
      (negative left, wider than the slide) is deliberate and is kept.
    """
    left, top, width, height = (int(v) for v in box)

    inherited = _inherited_box(src_layout, ph_idx)
    if inherited and all(
            abs(int(a) - b) <= _GEOM_TOLERANCE
            for a, b in zip(inherited, (left, top, width, height))):
        return True

    if src_width and abs(left) <= _GEOM_TOLERANCE \
            and abs(width - int(src_width)) <= _GEOM_TOLERANCE:
        return True

    return False


def _normalise_geometry(new_el, src_shape, ph_idx, layout_ph_idxs, src_layout,
                        ambiguous_idxs, src_width, dst_width):
    """
    Make a copied shape land where it should on the destination slide.

    Three distinct cases:

    * **Placeholder with no `<a:xfrm>` of its own.** It inherits geometry from
      whichever layout the slide uses. If the destination layout has the same
      placeholder index, leave the copy bare so it inherits the *destination*
      box — this is what keeps a hymn slide from a 4:3 (10in) source deck
      filling the 16:9 (13.33in) output slide, and it is how the hand-finished
      reference decks are built. Only when the destination layout has no such
      placeholder is there nothing to inherit from, and the source's resolved
      box has to be stamped on so the shape doesn't collapse to zero size.

    * **Placeholder whose explicit geometry only restates its source layout.**
      LibreOffice flattens inheritance when it converts a legacy `.ppt`: a
      placeholder that inherited its box comes out carrying an explicit copy
      of the layout's box. That is not real positioning, and taking it
      literally pinned a converted hymn title to the source deck's 9in width
      instead of letting it fill the 13.33in slide — so a deck built from
      `.ppt` did not match the same deck built from a hand re-saved `.pptx`.
      Geometry that merely repeats the source layout is dropped so the
      destination layout supplies it, as the first case does.

    * **Shape with explicit geometry, copied between different slide sizes.**
      Re-centre it horizontally by half the width difference, the way
      PowerPoint does when you paste a 4:3 slide into a 16:9 deck.
    """
    sp_pr = new_el.find(qn("p:spPr"))
    if sp_pr is None:
        return
    xfrm = sp_pr.find(qn("a:xfrm"))

    if xfrm is None:
        if ph_idx is not None and ph_idx in layout_ph_idxs:
            return  # inherit the destination layout's box
        if src_shape.left is None or src_shape.top is None:
            return
        xfrm = etree.Element(qn("a:xfrm"))
        etree.SubElement(xfrm, qn("a:off"),
                         x=str(src_shape.left), y=str(src_shape.top))
        etree.SubElement(xfrm, qn("a:ext"),
                         cx=str(src_shape.width), cy=str(src_shape.height))
        sp_pr.insert(0, xfrm)
        return

    if (ph_idx is not None and ph_idx in layout_ph_idxs
            and ph_idx not in ambiguous_idxs):
        actual = (src_shape.left, src_shape.top, src_shape.width, src_shape.height)
        if None not in actual and _is_flattened_inheritance(
                actual, src_layout, ph_idx, src_width):
            sp_pr.remove(xfrm)
            return  # inherit the destination layout's box

    if src_width and dst_width and src_width != dst_width:
        off = xfrm.find(qn("a:off"))
        if off is not None and off.get("x") is not None:
            shift = (dst_width - src_width) // 2
            off.set("x", str(int(off.get("x")) + shift))


def _copy_background(src_slide, new_slide):
    """Replace the new slide's background with the source slide's, if it has one."""
    src_cSld = src_slide._element.find(qn("p:cSld"))
    new_cSld = new_slide._element.find(qn("p:cSld"))
    src_bg = src_cSld.find(qn("p:bg"))
    if src_bg is None:
        return
    new_bg = new_cSld.find(qn("p:bg"))
    if new_bg is not None:
        new_cSld.remove(new_bg)
    new_cSld.insert(0, copy.deepcopy(src_bg))


def _remap_relationships(new_slide, src_slide):
    """
    Re-point relationship ids on the copied slide at its own relationships.

    A deep XML copy carries relationship ids (`r:embed` on an image blip,
    `r:id` on a chart or hyperlink) that only mean something relative to the
    *source* slide part. Left alone they dangle: PowerPoint shows a missing
    image, or offers to repair the file. Copy each referenced part into the new
    slide part and rewrite the id to the one that comes back.
    """
    src_rels = src_slide.part.rels
    new_part = new_slide.part
    remapped = {}

    for el in new_slide._element.iter():
        for name, value in list(el.attrib.items()):
            if not name.startswith(f"{{{_R_NS}}}") or not value:
                continue
            if value in remapped:
                el.set(name, remapped[value])
                continue
            rel = src_rels.get(value)
            if rel is None:
                logger.warning("Copied shape references unknown rId %s; dropping", value)
                del el.attrib[name]
                continue
            try:
                if rel.is_external:
                    new_rid = new_part.relate_to(
                        rel.target_ref, rel.reltype, is_external=True)
                else:
                    new_rid = new_part.relate_to(rel.target_part, rel.reltype)
            except Exception:
                logger.exception("Failed to copy relationship %s (%s)", value, rel.reltype)
                del el.attrib[name]
                continue
            remapped[value] = new_rid
            el.set(name, new_rid)


def _ensure_full_xfrm(shape):
    """
    Give a shape an explicit `<a:xfrm>` carrying its currently resolved box.

    Assigning `shape.left` or `shape.width` to a placeholder that inherits its
    geometry makes python-pptx create an `<a:xfrm>` holding *only* what was
    assigned — the unset half of `<a:off>`/`<a:ext>` is written as 0, which
    collapses the shape to zero height at the top of the slide. Resolving the
    inherited box first makes any later assignment safe.

    Returns False when the box can't be resolved, in which case the caller
    should leave the geometry alone.
    """
    sp_pr = shape._element.find(qn("p:spPr"))
    if sp_pr is None:
        return False
    if sp_pr.find(qn("a:xfrm")) is not None:
        return True

    left, top = shape.left, shape.top
    width, height = shape.width, shape.height
    if None in (left, top, width, height):
        return False

    xfrm = etree.Element(qn("a:xfrm"))
    etree.SubElement(xfrm, qn("a:off"), x=str(int(left)), y=str(int(top)))
    etree.SubElement(xfrm, qn("a:ext"), cx=str(int(width)), cy=str(int(height)))
    sp_pr.insert(0, xfrm)
    return True


def _drop_inheritable_xfrm(slide, shape):
    """
    Remove a shape's explicit `<a:xfrm>` so it inherits its layout's box.

    Only does so when the slide's layout actually has a placeholder with the
    same index to inherit from. Returns True when the box now comes from the
    layout — either because it was removed, or because there was none to begin
    with.
    """
    ph_idx = placeholder_idx(shape)
    if ph_idx is None or ph_idx not in _placeholder_idxs(slide.slide_layout):
        return False
    sp_pr = shape._element.find(qn("p:spPr"))
    if sp_pr is None:
        return False
    xfrm = sp_pr.find(qn("a:xfrm"))
    if xfrm is not None:
        sp_pr.remove(xfrm)
    return True


def restyle_responsive_slide(slide, slide_width):
    """
    Restyle a copied 啟應文 (responsive reading) slide to match the reference:

    - Title box: left as it is when it inherits the full-width 詩歌 title
      placeholder; only an explicitly narrower box gets centred.
    - Body box: paragraphs LEFT-aligned, manual line-wrap continuations merged
      back into their parent paragraph, box run full-bleed across the slide.
    """
    from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text
        if not text.strip():
            continue

        if "啟應文" in text:
            # Every reference deck runs this title across the full slide. The
            # source decks disagree about how to store its box — some inherit
            # it, and a LibreOffice-converted one carries an explicit 9in copy
            # — so normalise rather than trust what arrived: drop the explicit
            # box when the destination layout can supply one, and otherwise
            # centre what we have.
            if not _drop_inheritable_xfrm(slide, shape):
                if shape.width is not None and shape.width < slide_width \
                        and _ensure_full_xfrm(shape):
                    shape.left = (slide_width - shape.width) // 2
                    shape.top = 0
        else:
            _merge_continuation_paragraphs(shape.text_frame)
            for p in shape.text_frame.paragraphs:
                p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
            if _ensure_full_xfrm(shape):
                shape.left = 0
                shape.width = slide_width
            # The box runs to the slide edge, so CJK hanging punctuation —
            # which lets a trailing "；" or "." sit *outside* the text box —
            # puts those glyphs past the edge of the slide. The slide master
            # turns it off for body placeholders, but a copied reading body
            # isn't matched to one and picks up the default. Turn it off
            # explicitly, and leave a little more right inset besides.
            shape.text_frame.word_wrap = True
            shape.text_frame.margin_right = Emu(228600)
            for p in shape.text_frame.paragraphs:
                p._p.get_or_add_pPr().set("hangingPunct", "0")


def _merge_continuation_paragraphs(text_frame):
    """
    Merge paragraphs that start with whitespace (manual line-wrap continuations
    from narrow source boxes) into the preceding paragraph.
    """
    paragraphs = list(text_frame.paragraphs)
    for p in paragraphs[1:]:
        if not p.runs or not p.text or not p.text[0].isspace():
            continue
        prev = p._p.getprevious()
        if prev is None or prev.tag != qn("a:p"):
            continue

        # Drop leading whitespace-only runs, lstrip the first content run
        runs = list(p.runs)
        while runs and not runs[0].text.strip():
            runs[0]._r.getparent().remove(runs[0]._r)
            runs.pop(0)
        if runs:
            runs[0].text = runs[0].text.lstrip()

        # Move remaining runs into the previous paragraph, inserting before its
        # endParaRPr (if any) — a run appended after endParaRPr violates OOXML's
        # required child order and PowerPoint silently drops it on render.
        end_para_rpr = prev.find(qn("a:endParaRPr"))
        for r in runs:
            if end_para_rpr is not None:
                end_para_rpr.addprevious(r._r)
            else:
                prev.append(r._r)
        p._p.getparent().remove(p._p)


def clear_slides(prs):
    """Remove all slides from a presentation (no undo)."""
    sldIdLst = prs.slides._sldIdLst
    for i in range(len(prs.slides) - 1, -1, -1):
        sld_id = sldIdLst[i]
        rId = sld_id.get(f"{{{_R_NS}}}id")
        prs.part.drop_rel(rId)
        sldIdLst.remove(sld_id)


def get_layout(prs, name):
    """
    Get a slide layout by name from prs.

    Returns the matching layout, or the first layout if not found.
    """
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    return prs.slide_layouts[0]
