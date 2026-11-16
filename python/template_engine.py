#!/usr/bin/env python
"""Template-based deck building: duplicate template pages and rewrite text.

The visual design of a downloaded template lives in each slide's shapes
(images, art, decoration) rather than in the master or layouts. The only
faithful way to reuse it is to duplicate a sample page and rewrite its text.
"""

from __future__ import annotations

import os
from copy import deepcopy

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Pt

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_ID = f"{{{R_NS}}}id"
R_EMBED = f"{{{R_NS}}}embed"
R_LINK = f"{{{R_NS}}}link"

PAGE_TYPES = ("cover", "toc", "section", "content", "end")

TOC_HINTS = ("目录", "contents", "content", "目 录")
END_HINTS = ("谢谢", "感谢", "thank", "thanks", "q&a", "汇报完毕")


def _iter_shapes(shapes):
    """Walk shapes recursively so grouped text blocks are visible too."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(shape.shapes)
        else:
            yield shape


def _text_blocks(slide) -> list[dict]:
    """Collect text frames with their dominant font size, largest first."""
    blocks = []
    for shape in _iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if not text:
            continue
        sizes = []
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.font.size is not None:
                    sizes.append(run.font.size.pt)
        size = max(sizes) if sizes else 0
        blocks.append({"shape": shape, "size": size, "text": text})
    blocks.sort(key=lambda block: block["size"], reverse=True)
    return blocks


def classify_page(slide, index: int, total: int, blocks: list[dict]) -> str:
    """Classify a page by its text layout, not just by keywords."""
    joined = " ".join(block["text"] for block in blocks).lower()
    if any(hint in joined for hint in TOC_HINTS) and len(joined) < 60:
        return "toc"
    if any(hint in joined for hint in END_HINTS):
        return "end"
    if index == 0:
        return "cover"
    # A real content page carries at least one body-sized text block.
    body_blocks = [
        block for block in blocks
        if block["size"] < 24 and len(block["text"]) >= 8
    ]
    if len(body_blocks) >= 1:
        return "content"
    return "section"


def analyze_template(prs) -> dict:
    """Return per-page type and text-block metadata for all template pages."""
    pages = []
    total = len(prs.slides)
    for index, slide in enumerate(prs.slides):
        blocks = _text_blocks(slide)
        pages.append(
            {
                "index": index,
                "slide": slide,
                "blocks": blocks,
                "kind": classify_page(slide, index, total, blocks),
            }
        )
    return {"pages": pages, "total": total}


def pick_samples(analysis: dict, overrides: dict | None = None) -> dict:
    """Pick one sample page per page type.

    Content samples prefer pages with a title plus 2-4 body blocks (a clean
    content layout) over dense collage pages.
    """
    overrides = overrides or {}
    samples = {}
    for page in analysis["pages"]:
        kind = page["kind"]
        if kind == "content":
            body = [
                block for block in page["blocks"]
                if block["size"] < 24 and len(block["text"]) >= 8
            ]
            if 2 <= len(body) <= 4 and page["index"] > 0 and "content" not in samples:
                samples["content"] = page
        elif kind not in samples:
            samples[kind] = page

    if "content" not in samples:
        candidates = [page for page in analysis["pages"] if page["index"] > 0]
        if candidates:
            samples["content"] = max(
                candidates,
                key=lambda page: sum(len(block["text"]) for block in page["blocks"]),
            )

    # explicit 1-based page overrides win
    for kind, page_number in overrides.items():
        try:
            index = int(page_number) - 1
        except (TypeError, ValueError):
            continue
        if 0 <= index < len(analysis["pages"]):
            samples[kind] = analysis["pages"][index]
    return samples


def duplicate_slide(prs, source):
    """Copy one slide's shapes and relationships into a new slide."""
    new_slide = prs.slides.add_slide(source.slide_layout)
    for shape in list(new_slide.shapes):
        shape._element.getparent().remove(shape._element)
    for shape in source.shapes:
        new_slide.shapes._spTree.insert_element_before(deepcopy(shape._element), "p:extLst")
    # Slide backgrounds are siblings of spTree, not shapes.
    background = source._element.cSld.find(qn("p:bg"))
    if background is not None:
        new_slide._element.cSld.insert(0, deepcopy(background))
    rel_map = {}
    for r_id, rel in source.part.rels.items():
        # Notes must belong to the new slide, never to the original sample.
        if rel.reltype == RT.NOTES_SLIDE:
            continue
        if rel.is_external:
            new_r_id = new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_r_id = new_slide.part.relate_to(rel.target_part, rel.reltype)
        rel_map[r_id] = new_r_id
    for element in new_slide._element.iter():
        for attr in (R_ID, R_EMBED, R_LINK):
            value = element.get(attr)
            if value and value in rel_map:
                element.set(attr, rel_map[value])
    return new_slide


def _set_block_text(shape, text: str) -> None:
    """Replace paragraphs and fields, retaining the first paragraph/run style."""
    frame = shape.text_frame
    first = frame.paragraphs[0]
    paragraph_style = deepcopy(first._p.find(qn("a:pPr")))
    run_style = deepcopy(first.runs[0]._r.find(qn("a:rPr"))) if first.runs else None
    frame.clear()
    for index, line in enumerate(text.split("\n")):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        if paragraph_style is not None:
            existing = paragraph._p.find(qn("a:pPr"))
            if existing is not None:
                paragraph._p.remove(existing)
            paragraph._p.insert(0, deepcopy(paragraph_style))
        run = paragraph.add_run()
        if run_style is not None:
            run._r.insert(0, deepcopy(run_style))
        run.text = line


def _clear_block(shape) -> None:
    _set_block_text(shape, "")


def _normalize_bullets(items) -> list[str]:
    texts = []
    for item in items or []:
        if isinstance(item, str):
            texts.append(item)
        elif isinstance(item, dict) and item.get("text"):
            texts.append(str(item["text"]))
    return texts


def _shrink_to_fit(shape, text: str, max_chars: int = 30) -> None:
    """Scale runs down proportionally when text runs long."""
    if len(text) <= max_chars:
        return
    ratio = max(0.55, max_chars / max(len(text), 1))
    try:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.font.size is not None and run.font.size.pt >= 14:
                    run.font.size = Pt(run.font.size.pt * ratio)
    except Exception:
        pass


def _pick_title_block(blocks: list[dict]) -> dict | None:
    """The most prominent block in the title range, ignoring huge art text."""
    candidates = [block for block in blocks if 30 <= block["size"] <= 60]
    if candidates:
        return max(candidates, key=lambda block: block["size"])
    smaller = [block for block in blocks if block["size"] < 60]
    if smaller:
        return max(smaller, key=lambda block: block["size"])
    return blocks[0] if blocks else None


def _body_blocks(blocks: list[dict], title_shape) -> list[dict]:
    """Body-sized blocks other than the chosen title."""
    return [
        block for block in blocks
        if block["shape"] is not title_shape and block["size"] < 24
    ]


def _write_entries(blocks: list[dict], entries: list[str]) -> None:
    """Use all entries, grouping overflow into real paragraphs in the last box."""
    if entries and not blocks:
        raise ValueError("template sample has no text box for the requested content; choose another samplePages page")
    for index, block in enumerate(blocks):
        values = entries[index:] if index == len(blocks) - 1 else entries[index:index + 1]
        text = "\n".join(values)
        _set_block_text(block["shape"], text)
        _shrink_to_fit(block["shape"], text, 30)


def rewrite_page(slide, spec: dict, kind: str) -> None:
    """Rewrite one duplicated page's text from an outline slide spec."""
    blocks = _text_blocks(slide)
    title = str(spec.get("title") or "")
    subtitle = str(spec.get("subtitle") or "")
    tag = str(spec.get("tag") or "")
    bullets = _normalize_bullets(spec.get("bullets"))
    items = [str(item) for item in (spec.get("items") or [])]

    title_block = _pick_title_block(blocks)
    if title and title_block is None:
        raise ValueError("template sample has no title text box; choose another samplePages page")
    # Capture metadata before clearing every original string, including fields
    # and table cells that are not selected as content targets.
    for shape in _iter_shapes(slide.shapes):
        if shape.has_text_frame:
            _clear_block(shape)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    cell.text = ""
    if title_block is not None:
        _set_block_text(title_block["shape"], title)
        _shrink_to_fit(title_block["shape"], title, 16 if kind == "cover" else 14)
    remaining = [block for block in blocks if title_block is None or block["shape"] is not title_block["shape"]]
    if kind == "cover":
        smalls = [block for block in remaining if block["size"] < 30]
        _write_entries(smalls, [value for value in (subtitle, tag) if value])
    elif kind == "section":
        number = str(spec.get("number") or "")
        numbers = [block for block in remaining if block["size"] >= 60 or block["text"].strip().isdigit()]
        if number and numbers:
            _set_block_text(numbers[0]["shape"], number)
        smalls = [block for block in remaining if block not in numbers and block["size"] < 30]
        entries = ([number] if number and not numbers else []) + ([subtitle] if subtitle else [])
        _write_entries(smalls, entries)
    elif kind == "end":
        _write_entries([block for block in remaining if block["size"] < 30], [subtitle] if subtitle else [])
    elif kind == "toc":
        _write_entries(remaining, items)
    else:
        _write_entries(_body_blocks(blocks, title_block["shape"] if title_block else None), bullets)

    notes = str(spec.get("notes") or "")
    if notes:
        slide.notes_slide.notes_text_frame.text = notes


def _remove_slides_except(prs, keep_slides) -> None:
    """Drop every slide whose slide_id is not in keep_slides."""
    keep_ids = {slide.slide_id for slide in keep_slides}
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        try:
            slide_id = int(sld_id.get("id"))
        except (TypeError, ValueError):
            continue
        if slide_id in keep_ids:
            continue
        r_id = sld_id.get(R_ID)
        if r_id is not None:
            prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)


def build_from_template(outline: dict, output_path: str, template_path: str) -> dict:
    """Render an outline by duplicating and rewriting template pages."""
    prs = Presentation(template_path)
    analysis = analyze_template(prs)
    if not analysis["pages"]:
        raise ValueError("template must contain at least one slide")
    overrides = outline.get("samplePages") or {}
    samples = pick_samples(analysis, overrides if isinstance(overrides, dict) else {})

    slides_spec = [spec for spec in (outline.get("slides") or []) if isinstance(spec, dict)]
    if not slides_spec:
        raise ValueError("outline.slides must be a non-empty array")

    kept = []
    used_types = []
    for spec in slides_spec:
        kind = str(spec.get("type") or "content")
        if kind not in PAGE_TYPES:
            raise ValueError(f"template mode does not support slide type {kind!r}; use cover/toc/section/content/end, or omit template for built-in layouts")
        if kind == "cover":
            spec = {**spec, "title": spec.get("title") or outline.get("title"),
                    "subtitle": spec.get("subtitle") or outline.get("subtitle")}
        sample = samples.get(kind) or samples.get("content") or analysis["pages"][0]
        new_slide = duplicate_slide(prs, sample["slide"])
        rewrite_page(new_slide, spec, kind)
        kept.append(new_slide)
        used_types.append(kind)

    _remove_slides_except(prs, kept)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)

    return {
        "output": output_path,
        "slideCount": len(kept),
        "theme": str(outline.get("theme") or ""),
        "mode": "template",
        "template": template_path,
        "pageTypes": used_types,
        "samples": {
            kind: analysis["pages"][page["index"]]["index"] + 1
            for kind, page in samples.items()
        },
    }
