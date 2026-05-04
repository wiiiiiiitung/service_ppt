"""
Low-level XML manipulation for copying slides from template/library PPTXs.
"""

import copy
from lxml import etree
from pptx.oxml.ns import qn


def copy_slide(out_prs, src_prs, index):
    """
    Copy a slide from src_prs at index into out_prs.

    Performs deep XML copy of all shapes and background, preserving formatting
    and images. Relationships (embedded media, hyperlinks) are NOT copied due
    to python-pptx limitations.

    Args:
        out_prs: output Presentation
        src_prs: source Presentation
        index: 0-based slide index in src_prs
    """
    if index is None or index >= len(src_prs.slides):
        return

    src_slide = src_prs.slides[index]

    # Use a matching layout if possible
    layout_name = src_slide.slide_layout.name
    layout = get_layout(out_prs, layout_name)

    new_slide = out_prs.slides.add_slide(layout)

    # Replace shape tree content
    src_sp_tree = src_slide._element.find(qn("p:cSld")).find(qn("p:spTree"))
    new_sp_tree = new_slide._element.find(qn("p:cSld")).find(qn("p:spTree"))

    # Remove auto-generated placeholders
    for child in list(new_sp_tree):
        new_sp_tree.remove(child)

    # Copy shapes from source
    for child in src_sp_tree:
        new_sp_tree.append(copy.deepcopy(child))

    # Copy background
    src_cSld = src_slide._element.find(qn("p:cSld"))
    new_cSld = new_slide._element.find(qn("p:cSld"))
    src_bg = src_cSld.find(qn("p:bg"))
    if src_bg is not None:
        new_bg = new_cSld.find(qn("p:bg"))
        if new_bg is not None:
            new_cSld.remove(new_bg)
        new_cSld.insert(0, copy.deepcopy(src_bg))


def clear_slides(prs):
    """Remove all slides from a presentation (no undo)."""
    sldIdLst = prs.slides._sldIdLst
    for i in range(len(prs.slides) - 1, -1, -1):
        sld_id = sldIdLst[i]
        rId = sld_id.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
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
