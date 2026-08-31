"""Utilities to copy slides between python-pptx Presentation objects,
including their pictures (images) so nothing breaks visually."""
import copy
import io
from pptx.oxml.ns import qn


def _copy_images(dest_slide, src_slide):
    """For every picture/blip reference in the (already copied) dest_slide XML,
    make sure the underlying image part exists in dest's package and fix up
    the r:embed / r:link relationship ids to point at it."""
    src_part = src_slide.part
    dest_part = dest_slide.part
    for blip in dest_slide.shapes._spTree.findall('.//' + qn('a:blip')):
        for attr in (qn('r:embed'), qn('r:link')):
            old_rid = blip.get(attr)
            if not old_rid:
                continue
            if old_rid not in src_part.rels:
                continue
            rel = src_part.rels[old_rid]
            if 'image' not in rel.reltype:
                continue
            image_part = rel.target_part
            new_image_part, new_rid = dest_part.get_or_add_image_part(
                io.BytesIO(image_part.blob)
            )
            blip.set(attr, new_rid)


def append_slide(dest_prs, src_slide, layout_index=6):
    """Append a deep copy of src_slide (from another Presentation) to the
    end of dest_prs, preserving its shapes, text and pictures."""
    try:
        blank_layout = dest_prs.slide_layouts[layout_index]
    except IndexError:
        blank_layout = dest_prs.slide_layouts[-1]

    dest_slide = dest_prs.slides.add_slide(blank_layout)

    # Strip any placeholder shapes that came from the blank layout.
    for shape in list(dest_slide.shapes):
        shape._element.getparent().remove(shape._element)

    # Copy the background definition, if the source slide has one of its own.
    src_bg = src_slide._element.find(qn('p:cSld') + '/' + qn('p:bg'))
    if src_bg is not None:
        new_bg = copy.deepcopy(src_bg)
        cSld = dest_slide._element.find(qn('p:cSld'))
        cSld.insert(0, new_bg)

    # Copy every shape (text boxes, pictures, auto-shapes, ...).
    for shape in src_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        dest_slide.shapes._spTree.append(new_el)

    _copy_images(dest_slide, src_slide)
    return dest_slide


def delete_slide(prs, index):
    """Remove slide at position `index` from prs, including its relationship,
    so the part is fully dropped from the package (avoids leftover/duplicate
    slide parts when new slides are added afterwards)."""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    sldId = slides[index]
    rId = sldId.get(qn('r:id'))
    prs.part.drop_rel(rId)
    xml_slides.remove(sldId)
