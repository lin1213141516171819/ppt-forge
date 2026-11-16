"""Validate the complete outline before any output file is touched."""
from __future__ import annotations

import math
import re

FIELDS = {
    "cover": ("title", "subtitle", "tag", "image", "eyebrow", "badge"),
    "toc": ("title", "items"),
    "section": ("title", "subtitle", "number", "image"),
    "content": ("title", "bullets"),
    "two-column": ("title", "leftTitle", "leftBullets", "rightTitle", "rightBullets"),
    "table": ("title", "table"),
    "image": ("title", "image", "caption"),
    "image-text": ("title", "image", "caption", "bullets"),
    "gallery": ("title", "images", "captions"),
    "stats": ("title", "items"),
    "end": ("title", "subtitle"),
}
THEMES = ("navy", "teal", "wine", "graphite")
TEXT_FIELDS = ("title", "subtitle", "tag", "number", "image", "caption", "notes", "leftTitle", "rightTitle", "eyebrow", "badge")


class ValidationError(ValueError):
    code = "INVALID_OUTLINE"

    def __init__(self, issues):
        self.issues = issues
        super().__init__("; ".join(f"{item['path']}: {item['message']}" for item in issues))


def validate_outline(outline):
    """Return the input, or report all invalid field paths together."""
    issues = []

    def error(path, message):
        issues.append({"path": path, "message": message})

    def text_list(value, path, maximum=None, bullets=False):
        if not isinstance(value, list):
            error(path, "must be an array")
            return
        if maximum is not None and len(value) > maximum:
            error(path, f"supports at most {maximum} entries; split the content across slides")
        for index, item in enumerate(value):
            pos = f"{path}[{index}]"
            if isinstance(item, str):
                continue
            if bullets and isinstance(item, dict):
                if not isinstance(item.get("text"), str):
                    error(pos + ".text", "must be a string")
                level = item.get("level", 0)
                if type(level) is not int or not 0 <= level <= 8:
                    error(pos + ".level", "must be an integer between 0 and 8")
            else:
                error(pos, "must be a string or a {text, level} object" if bullets else "must be a string")

    def mapping(value, path, kind):
        if not isinstance(value, dict) or not value:
            error(path, "must be a non-empty field mapping")
            return
        for field, selectors in value.items():
            pos = f"{path}.{field}"
            if field not in FIELDS.get(kind, ()):
                error(pos, f"unknown field for {kind}")
                continue
            targets = selectors if isinstance(selectors, list) else [selectors]
            if not targets:
                error(pos, "must select at least one shape")
            if field in ("title", "subtitle", "tag", "number", "image", "caption", "leftTitle", "rightTitle", "table", "eyebrow", "badge") and len(targets) != 1:
                error(pos, "this field needs exactly one shape")
            for target in targets:
                if isinstance(target, str) and target.strip():
                    continue
                if not isinstance(target, dict) or set(target) not in ({"id"}, {"name"}):
                    error(pos, "selector must be a shape name, {name:string}, or {id:integer}")
                elif "id" in target and (type(target["id"]) is not int or target["id"] <= 0):
                    error(pos, "shape id must be a positive integer")
                elif "name" in target and (not isinstance(target["name"], str) or not target["name"].strip()):
                    error(pos, "shape name must be a non-empty string")

    if not isinstance(outline, dict):
        raise ValidationError([{"path": "outline", "message": "must be a JSON object"}])
    for field in ("title", "subtitle", "background"):
        if field in outline and not isinstance(outline[field], str):
            error(field, "must be a string")
    if "theme" in outline and outline["theme"] not in THEMES:
        error("theme", "must be navy, teal, wine, or graphite")
    if outline.get("qualityMode", "warn") not in ("warn", "strict"):
        error("qualityMode", "must be warn or strict")
    if "template" in outline and (not isinstance(outline["template"], str) or not outline["template"].strip()):
        error("template", "must be a non-empty .pptx file path")
    if "samplePages" in outline:
        if not isinstance(outline["samplePages"], dict):
            error("samplePages", "must be an object")
        else:
            for kind, page in outline["samplePages"].items():
                if kind not in FIELDS or type(page) is not int or page < 1:
                    error("samplePages." + kind, "must name a supported slide type and a positive 1-based page number")
    if "fieldMap" in outline:
        if not outline.get("template"):
            error("fieldMap", "requires a template")
        if not isinstance(outline["fieldMap"], dict):
            error("fieldMap", "must be an object keyed by slide type")
        else:
            for kind, fields in outline["fieldMap"].items():
                if kind not in FIELDS:
                    error("fieldMap." + kind, "unknown slide type")
                else:
                    mapping(fields, "fieldMap." + kind, kind)
    style = outline.get("style", {})
    if not isinstance(style, dict):
        error("style", "must be an object")
    else:
        for key in ("titleSize", "bodySize", "imageRadius"):
            value = style.get(key)
            if value is not None and (type(value) not in (int, float) or not math.isfinite(value) or value <= 0 and key != "imageRadius" or key == "imageRadius" and not 0 <= value <= 0.5):
                error("style." + key, "must be a positive finite number (imageRadius: 0 to 0.5)")
        for key in ("showPageNumber", "showFooter", "autoTitleSize"):
            if key in style and type(style[key]) is not bool:
                error("style." + key, "must be a boolean")
        for key, choices in (("titleAlign", ("left", "center")), ("imageFit", ("contain", "cover"))):
            if key in style and style[key] not in choices:
                error("style." + key, "must be one of " + ", ".join(choices))
        for key in ("footerText", "coverEyebrow", "coverBadge", "bulletChar"):
            if key in style and not isinstance(style[key], str):
                error("style." + key, "must be a string")
        palette = style.get("palette", {})
        if not isinstance(palette, dict):
            error("style.palette", "must be an object")
        else:
            for key, color in palette.items():
                if not isinstance(color, str) or not re.fullmatch(r"[0-9A-Fa-f]{6}", color):
                    error("style.palette." + key, "must be a six-digit RGB hex color")
    slides = outline.get("slides")
    if not isinstance(slides, list) or not slides:
        error("slides", "must be a non-empty array")
        slides = []
    for index, slide in enumerate(slides):
        path = f"slides[{index}]"
        if not isinstance(slide, dict):
            error(path, "must be an object")
            continue
        kind = slide.get("type", "content")
        if not isinstance(kind, str) or kind not in FIELDS:
            error(path + ".type", "unsupported slide type")
            continue
        allowed = {"type", "notes", "fieldMap", *FIELDS[kind]}
        if kind == "table":
            allowed = {"type", "notes", "fieldMap", "title", "headers", "rows"}
        if kind == "cover":
            allowed.update(("eyebrow", "badge"))
        if kind == "image-text":
            allowed.add("layout")
        for key in slide:
            if key not in allowed:
                error(path + "." + key, f"unknown field for {kind}")
        for field in TEXT_FIELDS:
            if field in slide and not isinstance(slide[field], str):
                error(path + "." + field, "must be a string")
        for field in ("bullets", "leftBullets", "rightBullets"):
            if field in slide:
                text_list(slide[field], path + "." + field, bullets=True)
        for field in ("images", "captions"):
            if field in slide:
                text_list(slide[field], path + "." + field, 4)
        if "layout" in slide and slide["layout"] not in ("left", "right"):
            error(path + ".layout", "must be left or right")
        if kind == "toc" and "items" in slide:
            text_list(slide["items"], path + ".items", 8 if not outline.get("template") else None)
        if kind == "stats":
            items = slide.get("items", [])
            if not isinstance(items, list):
                error(path + ".items", "must be an array")
            else:
                if not outline.get("template") and len(items) > 6:
                    error(path + ".items", "supports at most 6 entries")
                for pos, item in enumerate(items):
                    if not isinstance(item, dict):
                        error(f"{path}.items[{pos}]", "must be an object")
                        continue
                    for key in ("value", "label", "note"):
                        value = item.get(key, "")
                        if type(value) not in (str, int, float) or isinstance(value, float) and not math.isfinite(value):
                            error(f"{path}.items[{pos}].{key}", "must be a string or finite number")
        if kind == "table":
            if "headers" in slide:
                text_list(slide["headers"], path + ".headers")
            rows = slide.get("rows", [])
            if not isinstance(rows, list):
                error(path + ".rows", "must be an array of arrays")
            else:
                for pos, row in enumerate(rows):
                    if not isinstance(row, list) or any(type(cell) not in (str, int, float, bool, type(None)) or isinstance(cell, float) and not math.isfinite(cell) for cell in row):
                        error(f"{path}.rows[{pos}]", "must be an array of scalar cell values")
        if "fieldMap" in slide:
            if not outline.get("template"):
                error(path + ".fieldMap", "requires a template")
            mapping(slide["fieldMap"], path + ".fieldMap", kind)
    if issues:
        raise ValidationError(issues)
    return outline
