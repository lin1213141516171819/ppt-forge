#!/usr/bin/env python
"""Extract slide content from a PPTX file as JSON.

Usage:
    python extract.py <path-to-pptx>

Output (stdout): UTF-8 JSON with slide size, layout name, title, text
paragraphs (with nesting level), tables, notes, and image counts.
"""

from __future__ import annotations

import json
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def _paragraph_text(paragraph) -> str:
    """Join the runs of one paragraph into its visible text."""
    return "".join(run.text for run in paragraph.runs).strip()


def _collect_shape(shape, blocks: list[dict], is_title: bool) -> None:
    """Append one shape's content to `blocks` (recursing into groups)."""
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            _collect_shape(child, blocks, False)
        return

    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        blocks.append({"type": "image"})
        return

    if getattr(shape, "has_table", False) and shape.has_table:
        rows = []
        for row in shape.table.rows:
            cells = [cell.text.replace("\n", " / ").strip() for cell in row.cells]
            rows.append(cells)
        blocks.append({"type": "table", "rows": rows})
        return

    if not shape.has_text_frame or is_title:
        return

    paragraphs = []
    for paragraph in shape.text_frame.paragraphs:
        text = _paragraph_text(paragraph)
        if text:
            paragraphs.append({"text": text, "level": paragraph.level or 0})
    if paragraphs:
        blocks.append({"type": "text", "paragraphs": paragraphs})


def _inferred_title(blocks: list[dict]) -> str:
    """Guess a slide title when the layout has no title placeholder."""
    for block in blocks:
        if block.get("type") != "text":
            continue
        paragraphs = block.get("paragraphs") or []
        if not paragraphs:
            continue
        first = paragraphs[0].get("text", "").strip()
        if 0 < len(first) <= 60:
            return first
        return ""
    return ""


def _notes_text(slide) -> str:
    """Return speaker notes for one slide, or an empty string."""
    try:
        if not slide.has_notes_slide:
            return ""
        frame = slide.notes_slide.notes_text_frame
        return frame.text.strip() if frame is not None else ""
    except Exception:  # notes are decoration; never fail extraction over them
        return ""


def extract(path: str) -> dict:
    """Read one presentation file and return its structured content."""
    presentation = Presentation(path)
    slides = []

    for index, slide in enumerate(presentation.slides, start=1):
        title_shape = slide.shapes.title
        title = ""
        if title_shape is not None and title_shape.has_text_frame:
            title = title_shape.text_frame.text.strip()

        blocks: list[dict] = []
        for shape in slide.shapes:
            _collect_shape(shape, blocks, shape is title_shape)

        title_source = "placeholder"
        if not title:
            title = _inferred_title(blocks)
            title_source = "inferred" if title else "none"

        slides.append(
            {
                "index": index,
                "layout": slide.slide_layout.name,
                "title": title,
                "titleSource": title_source,
                "blocks": blocks,
                "notes": _notes_text(slide),
            }
        )

    return {
        "file": path,
        "slideWidthIn": round(presentation.slide_width / 914400, 2),
        "slideHeightIn": round(presentation.slide_height / 914400, 2),
        "slideCount": len(slides),
        "slides": slides,
    }


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print(json.dumps({"error": "usage: python extract.py <path-to-pptx>"}), file=sys.stderr)
        return 2
    try:
        result = extract(argv[1])
    except Exception as error:  # report the failure as structured JSON
        print(json.dumps({"error": f"{type(error).__name__}: {error}"}, ensure_ascii=False))
        return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main(sys.argv))
