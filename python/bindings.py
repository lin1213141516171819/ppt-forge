"""Explicit field-to-shape bindings for template slides."""
from __future__ import annotations

import os

from inspection import iter_shapes
from validation import FIELDS


def select_shape(slide, selector):
    selector = {"name": selector} if isinstance(selector, str) else selector
    matches = [shape for shape in iter_shapes(slide.shapes)
               if (shape.shape_id == selector["id"] if "id" in selector else shape.name == selector["name"])]
    if len(matches) != 1:
        raise ValueError(f"Shape selector {selector!r} matched {len(matches)} objects; use inspect and a unique shape ID")
    return matches[0]


def field_value(spec, field):
    if field == "table":
        return [spec.get("headers", []), *spec.get("rows", [])]
    if field == "items" and spec.get("type") == "stats":
        return ["\n".join(str(item[key]) for key in ("value", "label", "note") if key in item and item[key] != "")
                for item in spec.get("items", [])]
    value = spec.get(field, "")
    if field in ("bullets", "leftBullets", "rightBullets"):
        return [item if isinstance(item, str) else item["text"] for item in value or []]
    return value


def _selector_ids(slide, selectors):
    ids = set()
    for selector in selectors or []:
        ids.add(select_shape(slide, selector).shape_id)
    return ids


def clear_unbound_text(slide, mapped_ids, preserve_selectors):
    """Clear template prose unless it is explicitly mapped or preserved."""
    preserve_ids = _selector_ids(slide, preserve_selectors)
    for shape in iter_shapes(slide.shapes):
        if getattr(shape, "has_table", False) and shape.has_table and shape.shape_id not in mapped_ids:
            for row in shape.table.rows:
                for cell in row.cells:
                    cell.text = ""
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        if shape.shape_id in mapped_ids or shape.shape_id in preserve_ids:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.text = ""


def plan_fields(slide, spec, fields):
    kind = spec.get("type", "content")
    for field in FIELDS[kind]:
        if field_value(spec, field) and field not in fields:
            # A table with no headers/rows is not meaningful content.
            if field == "table" and not spec.get("headers") and not spec.get("rows"):
                continue
            raise ValueError(f"Field {field!r} has content but no fieldMap binding for {kind}")
    used = set()
    plan = []
    for field, selectors in fields.items():
        targets = selectors if isinstance(selectors, list) else [selectors]
        shapes = [select_shape(slide, selector) for selector in targets]
        for shape in shapes:
            if shape.shape_id in used:
                raise ValueError(f"Shape {shape.shape_id} is bound more than once")
            used.add(shape.shape_id)
            if field == "table":
                if not shape.has_table:
                    raise ValueError(f"Field table requires a table shape, got {shape.name!r}")
                rows = field_value(spec, field)
                columns = max((len(row) for row in rows), default=0)
                if len(rows) > len(shape.table.rows) or columns > len(shape.table.columns):
                    raise ValueError(f"Table data exceeds template capacity ({len(shape.table.rows)} rows x {len(shape.table.columns)} columns)")
                for row_index, row in enumerate(rows):
                    for col_index, value in enumerate(row):
                        if value not in ("", None) and shape.table.cell(row_index, col_index).is_spanned:
                            raise ValueError("Table data targets a hidden merged cell; change the data or select an unmerged table")
            elif field in ("image", "images"):
                if not (hasattr(shape, "image") or shape.is_placeholder):
                    raise ValueError(f"Field {field} requires a picture or placeholder shape, got {shape.name!r}")
                if shape.width <= 0 or shape.height <= 0:
                    raise ValueError("Mapped picture slots must have positive width and height")
            elif not shape.has_text_frame:
                raise ValueError(f"Field {field} requires a text shape, got {shape.name!r}")
        value = field_value(spec, field)
        if field == "images":
            if len(value or []) > len(shapes):
                raise ValueError("Gallery has more images than mapped picture slots")
            paths = value or []
        elif field == "image":
            paths = [value] if value else []
        else:
            paths = []
        for path in paths:
            if not os.path.isfile(path):
                raise FileNotFoundError(f"Mapped image not found: {path}")
        plan.append((field, shapes, value))
    return plan


def replace_picture(slide, shape, path):
    """Replace only this picture instance, keeping position, ID and z-order."""
    if not path:
        shape._element.getparent().remove(shape._element)
        return
    image = slide.shapes.add_picture(path, shape.left, shape.top, shape.width, shape.height)
    image.rotation = shape.rotation
    image.name = shape.name
    width, height = image.image.size
    source_aspect = width / height
    frame_aspect = shape.width / shape.height
    if source_aspect > frame_aspect:
        image.crop_left = image.crop_right = (1 - frame_aspect / source_aspect) / 2
    else:
        image.crop_top = image.crop_bottom = (1 - source_aspect / frame_aspect) / 2
    image._element.nvPicPr.cNvPr.set("id", str(shape.shape_id))
    shape._element.getparent().replace(shape._element, image._element)


def rewrite_mapped_page(slide, spec, fields):
    from template_engine import _set_block_text, _shrink_to_fit
    plan = plan_fields(slide, spec, fields)
    mapped_ids = {shape.shape_id for _field, shapes, _value in plan for shape in shapes}
    clear_unbound_text(slide, mapped_ids, spec.get("preserveText", []))
    for field, shapes, value in plan:
        if field == "table":
            table = shapes[0].table
            for row_index, row in enumerate(table.rows):
                for column_index, cell in enumerate(row.cells):
                    text = str(value[row_index][column_index]) if row_index < len(value) and column_index < len(value[row_index]) else ""
                    _set_block_text(cell, text)
        elif field in ("image", "images"):
            values = value if isinstance(value, list) else [value]
            for index, shape in enumerate(shapes):
                replace_picture(slide, shape, values[index] if index < len(values) else "")
        else:
            entries = value if isinstance(value, list) else [str(value)]
            for index, shape in enumerate(shapes):
                values = entries[index:] if index == len(shapes) - 1 else entries[index:index + 1]
                text = "\n".join(str(item) for item in values)
                _set_block_text(shape, text)
                _shrink_to_fit(shape, text, 30)
                if field in ("bullets", "leftBullets", "rightBullets"):
                    original = spec.get(field, [])
                    selected = original[index:] if index == len(shapes) - 1 else original[index:index + 1]
                    levels = [item.get("level", 0) if isinstance(item, dict) else 0 for item in selected
                              for _ in (item["text"] if isinstance(item, dict) else item).split("\n")]
                    for paragraph, level in zip(shape.text_frame.paragraphs, levels):
                        paragraph.level = level
    notes = spec.get("notes")
    if notes:
        slide.notes_slide.notes_text_frame.text = notes
