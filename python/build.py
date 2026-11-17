#!/usr/bin/env python
"""Build a clean 16:9 presentation from a structured JSON outline.

Usage:
    python build.py <outline.json|-> <output.pptx> [--theme navy]

Outline shape:
    {
      "title": "Deck title",
      "subtitle": "Deck subtitle",
      "theme": "navy",
      "slides": [
        {"type": "cover",       "title": ..., "subtitle": ..., "tag": ...},
        {"type": "toc",         "title": "目录", "items": ["...", "..."]},
        {"type": "section",     "number": "01", "title": ..., "subtitle": ...},
        {"type": "content",     "title": ..., "bullets": ["a", {"text": "b", "level": 1}], "notes": ...},
        {"type": "two-column",  "title": ..., "leftTitle": ..., "leftBullets": [...],
                                "rightTitle": ..., "rightBullets": [...]},
        {"type": "table",       "title": ..., "headers": [...], "rows": [[...]]},
        {"type": "end",         "title": "谢谢观看", "subtitle": ...}
      ]
    }

Optional style block (normally mapped from a learned style profile):
    {"style": {"titleAlign": "left"|"center", "showPageNumber": true,
               "showFooter": false, "footerText": "",
               "coverEyebrow": "", "coverBadge": "",
               "imageFit": "contain"|"cover", "imageRadius": 0.035,
               "autoTitleSize": true,
               "bulletChar": "•", "titleSize": 28, "bodySize": 16,
               "palette": {"dark": "...", "primary": "...", "accent": "..."}}}

Notes:
    - footerText / coverEyebrow / coverBadge default to empty, so no English
      template boilerplate is emitted unless you ask for it.
    - imageFit="cover" center-crops each image to its frame and rounds corners.
    - autoTitleSize shrinks long titles instead of letting them wrap onto body.

"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
import tempfile
from hashlib import md5

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from validation import validate_outline, ValidationError
from quality import analyze_deck, report_path, write_json

SLIDE_W = 13.333
SLIDE_H = 7.5
FONT = "Microsoft YaHei"

THEMES = {
    "navy": {
        "dark": "0B1220",
        "primary": "1D4ED8",
        "accent": "60A5FA",
        "ink": "0F172A",
        "body": "334155",
        "muted": "94A3B8",
        "line": "E2E8F0",
        "soft": "F1F5F9",
    },
    "teal": {
        "dark": "04211F",
        "primary": "0F766E",
        "accent": "2DD4BF",
        "ink": "0F172A",
        "body": "334155",
        "muted": "94A3B8",
        "line": "E2E8F0",
        "soft": "F0FDFA",
    },
    "wine": {
        "dark": "231014",
        "primary": "9F1239",
        "accent": "FB7185",
        "ink": "0F172A",
        "body": "334155",
        "muted": "94A3B8",
        "line": "E2E8F0",
        "soft": "FFF1F2",
    },
    "graphite": {
        "dark": "111113",
        "primary": "3F3F46",
        "accent": "A1A1AA",
        "ink": "0F172A",
        "body": "334155",
        "muted": "94A3B8",
        "line": "E2E8F0",
        "soft": "F4F4F5",
    },
}


DEFAULT_STYLE = {
    "titleAlign": "left",
    "showPageNumber": True,
    "showFooter": False,
    "footerText": "",
    "coverEyebrow": "",
    "coverBadge": "",
    "imageFit": "contain",
    "imageRadius": 0.035,
    "autoTitleSize": True,
    "bulletChar": "\u2022",
    "titleSize": None,
    "bodySize": None,
}

IMAGE_CACHE_DIR = os.path.join(tempfile.gettempdir(), "ppt-forge-imgcache")

def _style(theme: dict) -> dict:
    return theme.get("_style", DEFAULT_STYLE)


def rgb(hex_value: str) -> RGBColor:
    return RGBColor.from_string(hex_value)


def _set_font(run, size: float, color: str, bold: bool = False) -> None:
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = FONT
    run.font.color.rgb = rgb(color)
    rpr = run._r.get_or_add_rPr()
    east_asian = rpr.find(qn("a:ea"))
    if east_asian is None:
        east_asian = etree.SubElement(rpr, qn("a:ea"))
    east_asian.set("typeface", FONT)


def _set_bullet(paragraph, char: str, color: str | None = None) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    if color:
        bullet_color = etree.SubElement(ppr, qn("a:buClr"))
        srgb = etree.SubElement(bullet_color, qn("a:srgbClr"))
        srgb.set("val", color)
    bullet_font = etree.SubElement(ppr, qn("a:buFont"))
    bullet_font.set("typeface", "Arial")
    bullet_char = etree.SubElement(ppr, qn("a:buChar"))
    bullet_char.set("char", char)


def _indent(paragraph, marl_in: float, indent_in: float) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    ppr.set("marL", str(int(Inches(marl_in))))
    ppr.set("indent", str(int(Inches(indent_in))))


def _flatten(shape):
    """Drop the theme style reference so no inherited shadow or outline reaches the deck."""
    element = shape._element
    style = element.find(qn("p:style"))
    if style is not None:
        element.remove(style)
    return shape


def add_rect(slide, x, y, w, h, color, rounded: float | None = None):
    """Add a filled, borderless rectangle (or rounded rectangle)."""
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded is not None else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded is not None:
        try:
            shape.adjustments[0] = rounded
        except Exception:
            pass
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)
    shape.line.fill.background()
    return _flatten(shape)


def add_oval(slide, x, y, d, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)
    shape.line.fill.background()
    return _flatten(shape)


def add_text(
    slide,
    x,
    y,
    w,
    h,
    text: str,
    size: float,
    color: str,
    bold: bool = False,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    line_spacing: float | None = None,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    frame.margin_left = 0
    frame.margin_right = 0
    frame.margin_top = 0
    frame.margin_bottom = 0
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    if line_spacing is not None:
        paragraph.line_spacing = line_spacing
    run = paragraph.add_run()
    run.text = text
    _set_font(run, size, color, bold)
    return box


def paint_background(slide, color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def _paint_background(slide, deck, theme: dict) -> None:
    """Paint white, or stretch a template background image when provided."""
    background = deck.get("background") if isinstance(deck, dict) else None
    if isinstance(background, str) and background and os.path.isfile(background):
        aspect = _image_aspect(background)
        if aspect and aspect > 0:
            slide_aspect = SLIDE_W / SLIDE_H
            if aspect >= slide_aspect:
                draw_h = SLIDE_H
                draw_w = SLIDE_H * aspect
            else:
                draw_w = SLIDE_W
                draw_h = SLIDE_W / aspect
            slide.shapes.add_picture(
                background,
                Inches((SLIDE_W - draw_w) / 2),
                Inches((SLIDE_H - draw_h) / 2),
                Inches(draw_w),
                Inches(draw_h),
            )
            return
    paint_background(slide, "FFFFFF")


def _blank_slide(presentation):
    return presentation.slides.add_slide(presentation.slide_layouts[6])


# --------------------------------------------------------------------------- #
# Shared chrome
# --------------------------------------------------------------------------- #


def _chrome_top_bar(slide, theme: dict) -> None:
    """Legacy top strip; retained for callers that still opt in."""
    add_rect(slide, 0, 0, SLIDE_W, 0.07, theme["primary"])


def _chrome_footer(slide, theme: dict, page: int, total: int) -> None:
    style = _style(theme)
    footer_text = str(style.get("footerText") or "")
    if style.get("showFooter", False) and footer_text:
        add_text(slide, 0.75, 7.02, 4.0, 0.3, footer_text, 9, theme["muted"])
    if style.get("showPageNumber", True):
        add_text(
            slide,
            11.6,
            7.02,
            1.0,
            0.3,
            f"{page:02d} / {total:02d}",
            10,
            theme["muted"],
            align=PP_ALIGN.RIGHT,
        )


def _text_units(text: str) -> float:
    """Approximate rendered width in CJK character units (ASCII around 0.55)."""
    return sum(0.55 if ord(ch) < 128 else 1.0 for ch in text)


def _auto_title_size(title: str, base: float, width_in: float) -> float:
    """Shrink the title until it fits one line; never below 60% of the base size."""
    units = _text_units(title)
    if units <= 0 or base <= 0 or width_in <= 0:
        return base
    per_line = width_in / (base / 72.0)
    if units <= per_line * 0.94:
        return base
    return max(base * 0.60, base * per_line * 0.94 / units)


def _page_title(slide, theme: dict, title: str) -> None:
    style = _style(theme)
    align = PP_ALIGN.CENTER if style.get("titleAlign") == "center" else PP_ALIGN.LEFT
    title_size = style.get("titleSize") or 28
    if style.get("autoTitleSize", True):
        title_size = _auto_title_size(title, title_size, 11.83 if align == PP_ALIGN.CENTER else 11.1)
    if align == PP_ALIGN.CENTER:
        add_text(slide, 0.75, 0.55, 11.83, 0.9, title, title_size, theme["ink"], bold=True, align=PP_ALIGN.CENTER)
        add_rect(slide, (SLIDE_W - 0.9) / 2, 1.42, 0.9, 0.06, theme["primary"])
    else:
        add_rect(slide, 0.75, 0.62, 0.32, 0.32, theme["primary"])
        add_text(slide, 1.25, 0.52, 11.1, 0.75, title, title_size, theme["ink"], bold=True)
        add_rect(slide, 0.75, 1.22, 11.83, 0.02, theme["line"])


def _normalize_bullets(items) -> list[dict]:
    normalized = []
    for item in items or []:
        if isinstance(item, str):
            normalized.append({"text": item, "level": 0})
        elif isinstance(item, dict):
            normalized.append(
                {"text": str(item.get("text", "")), "level": int(item.get("level", 0) or 0)}
            )
    return [item for item in normalized if item["text"]]


def _fit_size(bullets: list[dict]) -> float:
    count = len(bullets)
    if count <= 4:
        return 19
    if count <= 6:
        return 17
    if count <= 9:
        return 15
    return 13.5


def _image_aspect(path: str) -> float | None:
    """Width / height of one image, or None when unreadable."""
    try:
        with Image.open(path) as image:
            width, height = image.size
            return width / height if height else None
    except Exception:
        return None


def _prepare_image(path: str, box_aspect: float, radius_ratio: float) -> str:
    """Center-crop to the box aspect and round the corners; returns a cached PNG path."""
    try:
        os.makedirs(IMAGE_CACHE_DIR, exist_ok=True)
        digest = md5(f"{os.path.abspath(path)}|{box_aspect!r}|{radius_ratio!r}".encode("utf-8"))
        with open(path, "rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        key = digest.hexdigest()
        out = os.path.join(IMAGE_CACHE_DIR, f"pf_{key}.png")
        if os.path.isfile(out):
            return out

        with Image.open(path) as source:
            image = source.convert("RGB")
            width, height = image.size
            if width <= 0 or height <= 0 or box_aspect <= 0:
                return path

            if width / height > box_aspect:
                target_w = max(1, int(height * box_aspect))
                left = (width - target_w) // 2
                image = image.crop((left, 0, left + target_w, height))
            elif width / height < box_aspect:
                target_h = max(1, int(width / box_aspect))
                top = (height - target_h) // 2
                image = image.crop((0, top, width, top + target_h))

            if image.width > 1600:
                scale = 1600 / image.width
                image = image.resize((1600, max(1, int(image.height * scale))), Image.LANCZOS)

            if radius_ratio > 0:
                w, h = image.size
                ss = 4
                mask_big = Image.new("L", (w * ss, h * ss), 0)
                ImageDraw.Draw(mask_big).rounded_rectangle(
                    [0, 0, w * ss - 1, h * ss - 1],
                    radius=max(1, int(w * ss * radius_ratio)),
                    fill=255,
                )
                image = image.convert("RGBA")
                image.putalpha(mask_big.resize((w, h), Image.LANCZOS))

            image.save(out)
            return out
    except Exception:
        return path


def add_picture_area(slide, x, y, w, h, path, theme: dict, caption: str = ""):
    """Fit an image inside the box, or draw a dashed placeholder when absent."""
    style = _style(theme)
    if path and os.path.isfile(path):
        if str(style.get("imageFit") or "") == "cover" and h > 0:
            prepared = _prepare_image(path, w / h, float(style.get("imageRadius") or 0.0))
            slide.shapes.add_picture(prepared, Inches(x), Inches(y), Inches(w), Inches(h))
            if caption:
                add_text(slide, x, y + h + 0.06, w, 0.3, caption, 10, theme["muted"], align=PP_ALIGN.CENTER)
            return None

        aspect = _image_aspect(path)
        if aspect and aspect > 0:
            box_aspect = w / h
            if aspect >= box_aspect:
                draw_w, draw_h = w, w / aspect
            else:
                draw_h, draw_w = h, h * aspect
            left = x + (w - draw_w) / 2
            top = y + (h - draw_h) / 2
            slide.shapes.add_picture(path, Inches(left), Inches(top), Inches(draw_w), Inches(draw_h))
            if caption:
                add_text(slide, x, y + h + 0.06, w, 0.3, caption, 10, theme["muted"], align=PP_ALIGN.CENTER)
            return None

    placeholder = add_rect(slide, x, y, w, h, theme["soft"], rounded=0.02)
    placeholder.line.color.rgb = rgb(theme["line"])
    placeholder.line.width = Pt(1.2)
    placeholder.line.dash_style = MSO_LINE.DASH
    label = caption or "图片占位"
    add_text(
        slide,
        x,
        y + h / 2 - 0.22,
        w,
        0.45,
        label,
        13,
        theme["muted"],
        align=PP_ALIGN.CENTER,
    )
    return placeholder


def _gallery_boxes(count: int) -> list[tuple[float, float, float, float]]:
    """Fixed grid positions for 1–4 images."""
    if count <= 1:
        return [(0.9, 1.9, 11.5, 4.75)]
    if count == 2:
        return [(0.75, 1.95, 5.75, 4.45), (6.83, 1.95, 5.75, 4.45)]
    if count == 3:
        return [
            (0.75, 1.95, 5.75, 2.12),
            (6.83, 1.95, 5.75, 2.12),
            ((SLIDE_W - 5.75) / 2, 4.3, 5.75, 2.1),
        ]
    return [
        (0.75, 1.95, 5.75, 2.12),
        (6.83, 1.95, 5.75, 2.12),
        (0.75, 4.3, 5.75, 2.1),
        (6.83, 4.3, 5.75, 2.1),
    ]


def _bullets_height(items: list[dict], width_in: float, size: float) -> float:
    """Rough rendered height (inches) of a bullet list, mirroring add_bullets()."""
    line_h = size / 72.0 * 1.18
    space_after = size / 72.0 * 0.55
    total = 0.0
    for item in items:
        level = item.get("level", 0)
        indent = 0.34 if level <= 0 else 0.82
        item_size = size if level <= 0 else max(size - 2.5, 11)
        avail = max(1.0, width_in - indent)
        per_line = max(1.0, avail / (item_size / 72.0) * 0.94)
        lines = max(1, int(math.ceil(_text_units(item.get("text", "")) / per_line)))
        total += lines * line_h + space_after
    return total


def add_bullets(slide, x, y, w, h, bullets: list[dict], theme: dict, base_size: float | None = None):
    """Render a bullet list into one text box; returns the created shape."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = MSO_ANCHOR.TOP
    frame.margin_left = 0
    frame.margin_right = 0
    frame.margin_top = 0
    frame.margin_bottom = 0

    size = base_size if base_size is not None else _fit_size(bullets)
    style = _style(theme)
    if style.get("bodySize") and not base_size:
        size = float(style["bodySize"])
    bullet_char = style.get("bulletChar") or "\u2022"
    first = True
    for item in bullets:
        paragraph = frame.paragraphs[0] if first else frame.add_paragraph()
        first = False
        level = item["level"]
        paragraph.space_after = Pt(size * 0.55)
        paragraph.line_spacing = 1.18
        if level <= 0:
            _indent(paragraph, 0.34, -0.26)
            _set_bullet(paragraph, bullet_char, theme["primary"])
            run = paragraph.add_run()
            run.text = item["text"]
            _set_font(run, size, theme["body"], bold=False)
        else:
            _indent(paragraph, 0.82, -0.24)
            _set_bullet(paragraph, "\u2013", theme["muted"])
            run = paragraph.add_run()
            run.text = item["text"]
            _set_font(run, max(size - 2.5, 11), theme["body"], bold=False)
    return box


# --------------------------------------------------------------------------- #
# Layouts
# --------------------------------------------------------------------------- #


def cover_slide(presentation, spec: dict, theme: dict, deck: dict):
    slide = _blank_slide(presentation)
    style = _style(theme)
    _paint_background(slide, deck, theme)

    add_picture_area(slide, 7.15, 0, 6.18, 5.95, str(spec.get("image") or ""), theme, "")

    title = str(spec.get("title") or deck.get("title") or "")
    subtitle = str(spec.get("subtitle") or deck.get("subtitle") or "")
    if title:
        if len(title) <= 8:
            title_size = 42
        elif len(title) <= 12:
            title_size = 34
        else:
            title_size = 30
        add_text(slide, 0.85, 1.75, 6.2, 2.0, title, title_size, theme["primary"], bold=True, line_spacing=1.12)
    eyebrow = str(spec.get("eyebrow") or style.get("coverEyebrow") or "")
    if eyebrow:
        add_text(slide, 0.88, 3.85, 6.0, 0.35, eyebrow, 12, theme["muted"], bold=True)
    if subtitle:
        add_text(slide, 0.88, 4.25, 6.0, 0.8, subtitle, 15, theme["body"], line_spacing=1.2)

    add_rect(slide, 0, 5.95, SLIDE_W, 1.55, theme["primary"])
    tag = str(spec.get("tag") or "")
    if tag:
        add_text(slide, 0.85, 6.42, 8.0, 0.45, tag, 13, "FFFFFF")
    badge = str(spec.get("badge") or style.get("coverBadge") or "")
    if badge:
        add_text(slide, 10.4, 6.42, 2.2, 0.45, badge, 11, "FFFFFF", bold=True, align=PP_ALIGN.RIGHT)

    add_rect(slide, 0, 0, 0.16, 5.95, theme["primary"])
    return slide


def toc_slide(presentation, spec: dict, theme: dict, deck: dict, page: int, total: int):
    slide = _blank_slide(presentation)
    _paint_background(slide, deck, theme)
    add_rect(slide, 0, 0, 0.22, SLIDE_H, theme["primary"])

    add_text(slide, 0.85, 0.6, 6.0, 0.8, str(spec.get("title") or "目录"), 32, theme["ink"], bold=True)
    add_rect(slide, 0.85, 1.35, 0.9, 0.07, theme["primary"])

    items = [str(item) for item in (spec.get("items") or []) if str(item).strip()]
    if items:
        start_y = 1.95
        gap = min(0.78, 4.7 / max(len(items[:8]), 1))
        for index, item in enumerate(items[:8]):
            top = start_y + index * gap
            add_text(slide, 0.85, top, 1.1, 0.55, f"{index + 1:02d}", 24, theme["primary"], bold=True)
            add_text(slide, 1.95, top + 0.04, 10.3, 0.5, item, 16, theme["ink"])
            add_rect(slide, 0.85, top + gap - 0.14, 11.4, 0.012, theme["line"])
    _chrome_footer(slide, theme, page, total)
    return slide


def section_slide(presentation, spec: dict, theme: dict, deck: dict):
    slide = _blank_slide(presentation)
    _paint_background(slide, deck, theme)
    add_rect(slide, 0, 0, 0.22, SLIDE_H, theme["primary"])
    add_picture_area(slide, 8.35, 0, 4.98, SLIDE_H, str(spec.get("image") or ""), theme, "")

    number = str(spec.get("number") or "")
    if number:
        add_text(slide, 0.8, 1.5, 4.5, 2.0, number, 100, theme["primary"], bold=True)
    add_rect(slide, 0.82, 3.85, 1.3, 0.07, theme["primary"])
    add_text(
        slide,
        0.82,
        4.1,
        7.0,
        1.1,
        str(spec.get("title") or ""),
        32,
        theme["ink"],
        bold=True,
        line_spacing=1.15,
    )
    subtitle = str(spec.get("subtitle") or "")
    if subtitle:
        add_text(slide, 0.82, 5.25, 6.8, 0.7, subtitle, 14, theme["muted"], line_spacing=1.3)
    return slide


def content_slide(presentation, spec: dict, theme: dict, deck: dict, page: int, total: int):
    slide = _blank_slide(presentation)
    _paint_background(slide, deck, theme)
    _page_title(slide, theme, str(spec.get("title") or ""))
    bullets = _normalize_bullets(spec.get("bullets"))
    if bullets:
        add_bullets(slide, 0.92, 1.55, 11.5, 5.1, bullets, theme)
    _chrome_footer(slide, theme, page, total)
    notes = str(spec.get("notes") or "")
    if notes:
        slide.notes_slide.notes_text_frame.text = notes
    return slide


def two_column_slide(presentation, spec: dict, theme: dict, deck: dict, page: int, total: int):
    slide = _blank_slide(presentation)
    _paint_background(slide, deck, theme)
    _page_title(slide, theme, str(spec.get("title") or ""))

    CARD_W = 5.72
    BODY_W = 4.95
    BASE = 15.0
    TOP, BOTTOM = 1.6, 6.45
    HEAD_GAP = 0.85          # 卡片顶 -> 正文顶
    TAIL_GAP = 0.35          # 正文底 -> 卡片底
    MIN_H = 1.6

    cards = [
        (0.75, spec.get("leftTitle"), spec.get("leftBullets")),
        (6.85, spec.get("rightTitle"), spec.get("rightBullets")),
    ]

    # 两张卡片共用一个高度：取内容较多的一侧，避免一高一矮
    need = 0.0
    for _, _, card_bullets in cards:
        items = _normalize_bullets(card_bullets)
        if items:
            need = max(need, _bullets_height(items, BODY_W, BASE))
    card_h = min(BOTTOM - TOP, max(MIN_H, need + HEAD_GAP + TAIL_GAP))
    top = TOP + ((BOTTOM - TOP) - card_h) / 2

    for x, card_title, card_bullets in cards:
        add_rect(slide, x, top, CARD_W, card_h, theme["soft"], rounded=0.04)
        add_text(slide, x + 0.4, top + 0.28, 5.0, 0.45, str(card_title or ""), 17, theme["primary"], bold=True)
        items = _normalize_bullets(card_bullets)
        if items:
            add_bullets(
                slide,
                x + 0.42,
                top + HEAD_GAP,
                BODY_W,
                card_h - HEAD_GAP - 0.1,
                items,
                theme,
                base_size=BASE,
            )
    _chrome_footer(slide, theme, page, total)
    return slide


def table_slide(presentation, spec: dict, theme: dict, deck: dict, page: int, total: int):
    slide = _blank_slide(presentation)
    _paint_background(slide, deck, theme)
    _page_title(slide, theme, str(spec.get("title") or ""))

    headers = [str(cell) for cell in (spec.get("headers") or [])]
    rows = [[str(cell) for cell in row] for row in (spec.get("rows") or [])]
    if not headers and not rows:
        return slide
    column_count = max(len(headers), max((len(row) for row in rows), default=0))
    if column_count == 0:
        return slide
    headers = (headers + [""] * column_count)[:column_count]
    row_data = [(row + [""] * column_count)[:column_count] for row in rows]

    table_shape = slide.shapes.add_table(
        len(row_data) + 1, column_count, Inches(0.75), Inches(1.95), Inches(11.83), Inches(0.62)
    )
    table = table_shape.table
    table.rows[0].height = Inches(0.55)
    column_width = Inches(11.83 / column_count)
    for column in table.columns:
        column.width = column_width

    for column_index, header in enumerate(headers):
        cell = table.cell(0, column_index)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(theme["primary"])
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.14)
        cell.margin_right = Inches(0.1)
        paragraph = cell.text_frame.paragraphs[0]
        if paragraph.runs:
            _set_font(paragraph.runs[0], 14, "FFFFFF", bold=True)

    for row_index, row in enumerate(row_data, start=1):
        table.rows[row_index].height = Inches(0.46)
        for column_index, value in enumerate(row):
            cell = table.cell(row_index, column_index)
            cell.text = value
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb("FFFFFF" if row_index % 2 == 1 else theme["soft"])
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.14)
            cell.margin_right = Inches(0.1)
            paragraph = cell.text_frame.paragraphs[0]
            if paragraph.runs:
                _set_font(paragraph.runs[0], 13, theme["body"])
    _chrome_footer(slide, theme, page, total)
    return slide


def image_slide(presentation, spec: dict, theme: dict, deck: dict, page: int, total: int):
    slide = _blank_slide(presentation)
    _paint_background(slide, deck, theme)
    _page_title(slide, theme, str(spec.get("title") or ""))
    add_picture_area(
        slide,
        0.9,
        1.55,
        11.5,
        5.1,
        str(spec.get("image") or ""),
        theme,
        str(spec.get("caption") or ""),
    )
    _chrome_footer(slide, theme, page, total)
    return slide


def image_text_slide(presentation, spec: dict, theme: dict, deck: dict, page: int, total: int):
    slide = _blank_slide(presentation)
    _paint_background(slide, deck, theme)
    _page_title(slide, theme, str(spec.get("title") or ""))
    bullets = _normalize_bullets(spec.get("bullets"))
    image_left = str(spec.get("layout") or "left") != "right"
    caption = str(spec.get("caption") or "")
    image_path = str(spec.get("image") or "")
    if image_left:
        add_picture_area(slide, 0.75, 1.55, 5.4, 5.05, image_path, theme, caption)
        if bullets:
            add_bullets(slide, 6.55, 1.55, 6.2, 5.05, bullets, theme)
    else:
        add_picture_area(slide, 7.15, 1.55, 5.4, 5.05, image_path, theme, caption)
        if bullets:
            add_bullets(slide, 0.9, 1.55, 5.95, 5.05, bullets, theme)
    _chrome_footer(slide, theme, page, total)
    return slide


def gallery_slide(presentation, spec: dict, theme: dict, deck: dict, page: int, total: int):
    slide = _blank_slide(presentation)
    _paint_background(slide, deck, theme)
    _page_title(slide, theme, str(spec.get("title") or ""))
    images = [str(value) for value in (spec.get("images") or []) if str(value).strip()]
    captions = spec.get("captions") or []
    count = max(len(images), len(captions))
    boxes = _gallery_boxes(count)
    for index, box in enumerate(boxes):
        image_path = images[index] if index < len(images) else ""
        caption = str(captions[index]) if index < len(captions) else ""
        add_picture_area(slide, box[0], box[1], box[2], box[3], image_path, theme, caption)
    _chrome_footer(slide, theme, page, total)
    return slide


def stats_slide(presentation, spec: dict, theme: dict, deck: dict, page: int, total: int):
    slide = _blank_slide(presentation)
    _paint_background(slide, deck, theme)
    _page_title(slide, theme, str(spec.get("title") or ""))

    items = [item for item in (spec.get("items") or []) if isinstance(item, dict)][:6]
    if items:
        count = len(items)
        gap = 0.3 if count <= 4 else 0.2
        card_width = (11.83 - gap * (count - 1)) / count
        base_value_size = 38 if count <= 4 else (30 if count == 5 else 25)
        label_size = 15 if count <= 4 else 13
        for index, item in enumerate(items):
            x = 0.75 + index * (card_width + gap)
            add_rect(slide, x, 1.75, card_width, 3.9, theme["soft"], rounded=0.04)
            value_text = str(item.get("value", ""))
            value_size = base_value_size
            units = _text_units(value_text)
            if units > 0 and units * (value_size / 72.0) > card_width * 0.92:
                value_size = max(14.0, card_width * 0.92 * 72.0 / units)
            add_text(
                slide,
                x,
                2.2,
                card_width,
                1.0,
                value_text,
                value_size,
                theme["primary"],
                bold=True,
                align=PP_ALIGN.CENTER,
            )
            add_text(
                slide,
                x,
                3.3,
                card_width,
                0.5,
                str(item.get("label", "")),
                label_size,
                theme["ink"],
                align=PP_ALIGN.CENTER,
            )
            note = str(item.get("note", ""))
            if note:
                add_text(
                    slide,
                    x + 0.2,
                    3.9,
                    card_width - 0.4,
                    1.3,
                    note,
                    11,
                    theme["muted"],
                    align=PP_ALIGN.CENTER,
                    line_spacing=1.3,
                )
    _chrome_footer(slide, theme, page, total)
    return slide


def end_slide(presentation, spec: dict, theme: dict, deck: dict):
    slide = _blank_slide(presentation)
    paint_background(slide, theme["dark"])
    add_oval(slide, -1.8, 4.6, 5.2, theme["primary"])
    add_text(
        slide,
        0,
        2.55,
        SLIDE_W,
        1.0,
        str(spec.get("title") or "谢谢观看"),
        38,
        "FFFFFF",
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    subtitle = str(spec.get("subtitle") or "")
    if subtitle:
        add_text(slide, 0, 3.75, SLIDE_W, 0.6, subtitle, 15, theme["muted"], align=PP_ALIGN.CENTER)
    add_rect(slide, SLIDE_W / 2 - 0.75, 4.55, 1.5, 0.075, theme["accent"])
    return slide


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #

RENDERERS = {
    "cover": lambda presentation, spec, theme, deck, number, total: cover_slide(
        presentation, spec, theme, deck
    ),
    "toc": lambda presentation, spec, theme, deck, number, total: toc_slide(
        presentation, spec, theme, deck, number, total
    ),
    "section": lambda presentation, spec, theme, deck, number, total: section_slide(
        presentation, spec, theme, deck
    ),
    "content": lambda presentation, spec, theme, deck, number, total: content_slide(
        presentation, spec, theme, deck, number, total
    ),
    "two-column": lambda presentation, spec, theme, deck, number, total: two_column_slide(
        presentation, spec, theme, deck, number, total
    ),
    "table": lambda presentation, spec, theme, deck, number, total: table_slide(
        presentation, spec, theme, deck, number, total
    ),
    "image": lambda presentation, spec, theme, deck, number, total: image_slide(
        presentation, spec, theme, deck, number, total
    ),
    "image-text": lambda presentation, spec, theme, deck, number, total: image_text_slide(
        presentation, spec, theme, deck, number, total
    ),
    "gallery": lambda presentation, spec, theme, deck, number, total: gallery_slide(
        presentation, spec, theme, deck, number, total
    ),
    "stats": lambda presentation, spec, theme, deck, number, total: stats_slide(
        presentation, spec, theme, deck, number, total
    ),
    "end": lambda presentation, spec, theme, deck, number, total: end_slide(
        presentation, spec, theme, deck
    ),
}


def _build(outline: dict, output_path: str) -> dict:
    """Render one outline into a .pptx file; returns a small summary."""
    template = outline.get("template")
    if template is not None:
        if not isinstance(template, str) or not template.strip():
            raise ValueError("template must be a non-empty .pptx file path")
        if not os.path.isfile(template):
            raise FileNotFoundError(f"template not found: {template}")
        from template_engine import build_from_template

        return build_from_template(outline, output_path, template)

    theme_name = str(outline.get("theme") or "navy")
    theme = dict(THEMES.get(theme_name, THEMES["navy"]))

    style = {**DEFAULT_STYLE, **(outline.get("style") or {})}
    palette = style.get("palette")
    if isinstance(palette, dict):
        for key in ("dark", "primary", "accent", "ink", "body", "muted", "line", "soft"):
            value = palette.get(key)
            if isinstance(value, str) and len(value) == 6:
                theme[key] = value.upper()
    if style.get("titleAlign") not in ("left", "center"):
        style["titleAlign"] = "left"
    theme["_style"] = style

    slides = [slide for slide in (outline.get("slides") or []) if isinstance(slide, dict)]
    if not slides or slides[0].get("type") != "cover":
        slides.insert(0, {"type": "cover"})
    if outline.get("syncToc"):
        titles = [str(slide.get("title") or "") for slide in slides]
        for index, slide in enumerate(slides):
            if slide.get("type") == "toc":
                slide = dict(slide)
                slide["items"] = [title for title in titles if title and title != str(slide.get("title") or "")]
                slides[index] = slide
                break

    presentation = Presentation()
    presentation.slide_width = Inches(SLIDE_W)
    presentation.slide_height = Inches(SLIDE_H)

    total = len(slides)
    for index, spec in enumerate(slides, start=1):
        kind = str(spec.get("type") or "content")
        renderer = RENDERERS.get(kind)
        if renderer is None:
            spec = {**spec, "type": "content"}
            renderer = RENDERERS["content"]
        renderer(presentation, spec, theme, outline, index, total)

    presentation.save(output_path)
    return {
        "output": output_path,
        "slideCount": total,
        "theme": theme_name,
    }


def build(outline: dict, output_path: str) -> dict:
    """Validate, generate aside, inspect, then publish a complete deck atomically."""
    validate_outline(outline)
    directory = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(directory, exist_ok=True)
    descriptor, staging = tempfile.mkstemp(prefix=".ppt-forge-", suffix=".pptx", dir=directory)
    os.close(descriptor)
    try:
        result = _build(outline, staging)
        report = analyze_deck(staging, outline.get("quality", {})) if outline.get("quality") else analyze_deck(staging)
        # Record missing local assets explicitly even when free layouts use
        # placeholder boxes. Mapped template images already fail before writing.
        paths = [(0, outline.get("background"))]
        offset = report["slideCount"] - len(outline["slides"])
        for page, slide in enumerate(outline["slides"], 1 + offset):
            paths.append((page, slide.get("image")))
            paths.extend((page, path) for path in slide.get("images", []))
        for page, path in paths:
            if path and not os.path.isfile(path):
                report["static"]["findings"].append({"severity": "warning", "code": "MISSING_IMAGE", "page": page,
                    "message": f"Image not found; generated output may contain a placeholder: {path}"})
        errors = [item for item in report["static"]["findings"] if item["severity"] == "error"]
        if errors or outline.get("qualityMode") == "strict" and report["static"]["findings"]:
            raise ValidationError([{"path": f"page {item.get('page', '?')}", "message": item["code"] + ": " + item["message"]}
                                   for item in (errors or report["static"]["findings"])])
        report["deck"] = os.path.abspath(output_path)
        os.replace(staging, output_path)
        write_json(report_path(output_path), report)
        return {**result, "output": output_path, "qualityReport": report_path(output_path), "quality": report}
    finally:
        if os.path.exists(staging):
            os.unlink(staging)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("outline", help="outline JSON path, or - for stdin")
    parser.add_argument("output", help="output .pptx path")
    parser.add_argument("--theme", choices=tuple(THEMES))
    try:
        args = parser.parse_args(argv[1:])
    except SystemExit as error:
        return int(error.code)

    outline_path, output_path = args.outline, args.output
    try:
        if outline_path == "-":
            outline = json.load(sys.stdin)
        else:
            with open(outline_path, "r", encoding="utf-8") as handle:
                outline = json.load(handle)
    except Exception as error:
        print(json.dumps({"error": f"cannot read outline: {type(error).__name__}: {error}"}, ensure_ascii=False))
        return 1

    try:
        if not isinstance(outline, dict):
            raise ValueError("outline must be a JSON object")
        if args.theme:
            outline["theme"] = args.theme
        result = build(outline, output_path)
    except Exception as error:
        print(json.dumps({"error": f"{type(error).__name__}: {error}",
            "code": getattr(error, "code", "BUILD_FAILED"),
            **({"issues": error.issues} if hasattr(error, "issues") else {})}, ensure_ascii=False))
        return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main(sys.argv))
