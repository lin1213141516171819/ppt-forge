"""Incrementally edit an existing deck without rebuilding unrelated pages."""
from __future__ import annotations

import argparse
import json
import os
import tempfile
from copy import deepcopy
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from template_engine import _iter_shapes, _set_block_text


def _shape(slide, selector):
    if isinstance(selector, str):
        selector = {"name": selector}
    matches = []
    for shape in _iter_shapes(slide.shapes):
        if "id" in selector and shape.shape_id == selector["id"]:
            matches.append(shape)
        elif "name" in selector and shape.name == selector["name"]:
            matches.append(shape)
    if len(matches) != 1:
        raise ValueError(f"selector {selector!r} matched {len(matches)} objects")
    return matches[0]


def _remove_slide(prs, index):
    slides = list(prs.slides)
    if not 1 <= index <= len(slides):
        raise ValueError(f"page {index} is outside the deck")
    slide = slides[index - 1]
    element = slide._element
    r_id = prs.part.rels.get_or_add(slide.part) if False else None
    for item in list(prs.slides._sldIdLst):
        if item.get("id") == str(slide.slide_id):
            r_id = item.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if r_id:
                prs.part.drop_rel(r_id)
            prs.slides._sldIdLst.remove(item)
            break


def _reorder(prs, order):
    slides = list(prs.slides)
    if sorted(order) != list(range(1, len(slides) + 1)):
        raise ValueError("pageOrder must contain every current 1-based page exactly once")
    nodes = list(prs.slides._sldIdLst)
    by_id = {int(node.get("id")): node for node in nodes}
    ordered = [by_id[slides[index - 1].slide_id] for index in order]
    sld_list = prs.slides._sldIdLst
    for node in list(sld_list):
        sld_list.remove(node)
    for node in ordered:
        sld_list.append(node)


def _titles(prs):
    values = []
    for slide in prs.slides:
        text = ""
        for shape in _iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                text = shape.text_frame.text.strip().splitlines()[0]
                break
        values.append(text)
    return values


def apply_edit(source, output, operations):
    prs = Presentation(source)
    for page in sorted({int(value) for value in operations.get("deletePages", [])}, reverse=True):
        _remove_slide(prs, page)

    for item in operations.get("addPages", []):
        after = int(item.get("after", len(prs.slides)))
        template_page = int(item.get("templatePage", 1))
        source_slide = list(prs.slides)[template_page - 1]
        from template_engine import duplicate_slide
        new_slide = duplicate_slide(prs, source_slide)
        spec = item.get("replace", {})
        for selector, value in spec.items():
            _set_block_text(_shape(new_slide, selector), str(value))
        nodes = list(prs.slides._sldIdLst)
        new_node = nodes[-1]
        sld_list = prs.slides._sldIdLst
        sld_list.remove(new_node)
        current = list(sld_list)
        if after < 0 or after > len(current):
            raise ValueError(f"addPages.after {after} is outside the deck")
        sld_list.insert(after, new_node)

    for item in operations.get("replace", []):
        page = int(item["page"])
        shape = _shape(list(prs.slides)[page - 1], item["selector"])
        if "text" in item:
            _set_block_text(shape, str(item["text"]))

    if "pageOrder" in operations:
        _reorder(prs, [int(value) for value in operations["pageOrder"]])

    toc = operations.get("syncToc")
    if toc:
        slide = list(prs.slides)[int(toc["page"]) - 1]
        titles = _titles(prs)
        items = toc.get("items") or [title for index, title in enumerate(titles, 1) if index != int(toc["page"]) and title]
        _set_block_text(_shape(slide, toc["selector"]), "\n".join(str(item) for item in items))

    numbers = operations.get("pageNumbers")
    if numbers:
        total = len(prs.slides)
        for page, slide in enumerate(prs.slides, 1):
            _set_block_text(_shape(slide, numbers["selector"]), str(numbers.get("format", "{page:02d}/{total:02d}")).format(page=page, total=total))

    directory = os.path.dirname(os.path.abspath(output))
    os.makedirs(directory, exist_ok=True)
    descriptor, staging = tempfile.mkstemp(prefix=".ppt-forge-edit-", suffix=".pptx", dir=directory)
    os.close(descriptor)
    try:
        prs.save(staging)
        os.replace(staging, output)
    finally:
        if os.path.exists(staging):
            os.unlink(staging)
    return {"output": os.path.abspath(output), "slideCount": len(prs.slides), "operation": "edit"}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("source")
    parser.add_argument("output")
    parser.add_argument("operations")
    args = parser.parse_args()
    try:
        with open(args.operations, encoding="utf-8") as handle:
            operations = json.load(handle)
        print(json.dumps(apply_edit(args.source, args.output, operations), ensure_ascii=False))
        return 0
    except Exception as error:
        print(json.dumps({"error": str(error), "code": "EDIT_FAILED"}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
