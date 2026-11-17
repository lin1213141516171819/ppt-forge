"""Expose stable template selectors and approximate text capacity."""
from __future__ import annotations

import argparse
import json
import math
import sys

EMU = 914400


def iter_shapes(shapes):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)
        else:
            yield shape


def text_metrics(shape):
    """Estimate line layout; font rendering still needs visual review."""
    frame = shape.text_frame
    width = max(0, shape.width - frame.margin_left - frame.margin_right) / EMU
    height = max(0, shape.height - frame.margin_top - frame.margin_bottom) / EMU
    sizes = [run.font.size.pt for p in frame.paragraphs for run in p.runs if run.font.size is not None]
    if not sizes:
        for element in shape._element.iter():
            if element.tag.endswith('}defRPr') and element.get('sz'):
                sizes.append(float(element.get('sz')) / 100.0)
                break
            if element.tag.endswith('}rPr') and element.get('sz'):
                sizes.append(float(element.get('sz')) / 100.0)
                break
    size = max(sizes) if sizes else 18
    line_capacity = max(1, width * 72 / size * 0.94)
    needed = 0.0
    for paragraph in frame.paragraphs:
        units = sum(0.55 if ord(ch) < 128 else 1 for ch in paragraph.text)
        lines = max(1, math.ceil(units / line_capacity)) if frame.word_wrap is not False else 1
        spacing = paragraph.line_spacing
        if isinstance(spacing, float):
            line_height = size * spacing / 72
        elif spacing is not None:
            line_height = spacing.pt / 72
        else:
            line_height = size * 1.2 / 72
        needed += lines * line_height
        needed += sum(value.pt / 72 for value in (paragraph.space_before, paragraph.space_after) if value is not None)
    return {"fontSizePt": size, "fontSizeEstimated": not bool(sizes),
            "capacityCjkApprox": max(0, int(line_capacity * height * 72 / (size * 1.2))),
            "requiredHeightInApprox": round(needed, 3), "availableHeightIn": round(height, 3)}


def inspect_deck(path):
    from pptx import Presentation
    prs = Presentation(path)
    slides = []
    for index, slide in enumerate(prs.slides, 1):
        objects = []
        for shape in iter_shapes(slide.shapes):
            item = {"id": shape.shape_id, "name": shape.name,
                    "coordinateSpace": "group-local" if shape._element.getparent().tag.endswith("}grpSp") else "slide",
                    "kind": "table" if shape.has_table else "text" if shape.has_text_frame else "picture" if hasattr(shape, "image") else "shape",
                    "boundsIn": {key: round(getattr(shape, key) / EMU, 3) for key in ("left", "top", "width", "height")},
                    "selector": {"id": shape.shape_id}}
            if shape.has_text_frame:
                item.update(text=shape.text_frame.text, **text_metrics(shape))
            if shape.has_table:
                item.update(rows=len(shape.table.rows), columns=len(shape.table.columns))
            objects.append(item)
        slides.append({"page": index, "slideId": slide.slide_id, "objects": objects})
    return {"source": path, "slideCount": len(slides), "slideWidthIn": prs.slide_width / EMU,
            "slideHeightIn": prs.slide_height / EMU, "slides": slides,
            "capacityNote": "Capacity is estimated; review rendered slides before delivery."}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    args = parser.parse_args()
    try:
        print(json.dumps(inspect_deck(args.path), ensure_ascii=False))
    except Exception as error:
        print(json.dumps({"error": str(error), "code": "INSPECTION_FAILED"}, ensure_ascii=False))
        return 1
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
