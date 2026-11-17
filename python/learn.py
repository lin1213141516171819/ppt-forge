#!/usr/bin/env python
"""Analyze many .pptx files and write a reusable style profile.

Usage:
    python learn.py <input-dir-or-file> <output-base> [--name "档案名"]

Writes <output-base>.json (machine) and <output-base>.md (readable).
Prints the profile JSON summary to stdout.
"""

from __future__ import annotations

import json
import os
import re
import sys
from collections import Counter
from datetime import datetime
from glob import glob

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

MAX_FILES = 500
NUMBERING_OPENER = re.compile(r"^(?:[一二三四五六七八九十]+[、.．]|\d+[.、．)）]|[（(][一二三四五六七八九十\d]+[)）])")
TITLE_SKIP = {"谢谢", "谢谢观看", "谢谢大家", "thank you", "q&a", "汇报完毕", "感谢聆听"}
LABEL_NOISE = {"presentation", "section", "report", "汇报", "目录", "contents", "content"}
PLACEHOLDER_MARKERS = (
    "点击添加",
    "点击输入",
    "请输入",
    "您的内容",
    "您的文本",
    "输入文字",
    "lorem ipsum",
    "your logo",
    "click to add",
    "add text",
    "示例文字",
    "样例文字",
)


def _is_placeholder(text: str) -> bool:
    """Detect template filler text that carries no style signal."""
    lowered = text.lower()
    return any(marker in lowered for marker in PLACEHOLDER_MARKERS)


def _is_toc_title(title: str) -> bool:
    """Match common table-of-contents labels, including misspellings."""
    lowered = title.lower().replace(" ", "")
    return "目录" in title or "contents" in lowered or "content" in lowered


def _is_ending_title(title: str) -> bool:
    """Match common closing labels."""
    lowered = title.lower()
    return any(word in lowered for word in ("谢谢", "thank", "感谢", "thanks", "q&a", "qa", "汇报完毕"))


def _is_likely_title(text: str) -> bool:
    """Reject decoration labels, numbers, and overlong paragraph-like lines."""
    stripped = text.strip()
    if not (4 <= len(stripped) <= 40):
        return False
    if re.fullmatch(r"[\d\s./-]+", stripped):
        return False
    if stripped.lower() in LABEL_NOISE:
        return False
    if _is_placeholder(stripped):
        return False
    if re.fullmatch(r"[A-Z][A-Z\s&]+", stripped):
        return False
    return True


def _rgb(shape) -> str | None:
    try:
        if shape.fill.type is not None and shape.fill.fore_color.type is not None:
            return str(shape.fill.fore_color.rgb)
    except Exception:
        return None
    return None


def _run_color(run) -> str | None:
    try:
        if run.font.color.type is not None:
            return str(run.font.color.rgb)
    except Exception:
        return None
    return None


def _iter_text_runs(shape):
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            yield from _iter_text_runs(child)
        return
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                yield paragraph, run


def _is_picture(shape) -> bool:
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return any(_is_picture(child) for child in shape.shapes)
    return shape.shape_type == MSO_SHAPE_TYPE.PICTURE


def _slide_stats(slide, index: int) -> dict:
    is_cover = index == 1
    title = ""
    try:
        if slide.shapes.title is not None and slide.shapes.title.has_text_frame:
            title = slide.shapes.title.text_frame.text.strip()
    except Exception:
        title = ""

    text_chars = 0
    bullets = 0
    openers: list[str] = []
    font_sizes: list[float] = []
    fonts: list[str] = []
    colors: list[str] = []
    title_aligns: list[str] = []
    pictures = 0
    tables = 0

    for shape in slide.shapes:
        if _is_picture(shape):
            pictures += 1
        if getattr(shape, "has_table", False) and shape.has_table:
            tables += 1
        fill = _rgb(shape)
        if fill and fill not in ("FFFFFF", "000000"):
            colors.append(fill)

        for paragraph, run in _iter_text_runs(shape):
            text = (run.text or "").strip()
            if not text:
                continue
            text_chars += len(text)
            size = run.font.size.pt if run.font.size is not None else None
            if size is not None:
                font_sizes.append(round(size, 1))
            if run.font.name:
                fonts.append(run.font.name)
            color = _run_color(run)
            if color and color not in ("000000", "FFFFFF"):
                colors.append(color)
            if not is_cover and paragraph.level == 0 and len(text) >= 6:
                bullets += 1
                head = text[:4]
                match = NUMBERING_OPENER.match(text)
                if match:
                    token = match.group(0)
                    if re.match(r"^[一二三四五六七八九十]", token):
                        openers.append(token)
                elif "：" in head or ":" in head:
                    prefix = head.split("：")[0].split(":")[0].strip()
                    if prefix:
                        openers.append(prefix + "：")
            if size is not None and size >= 20 and paragraph.level == 0:
                try:
                    alignment = paragraph.alignment
                    if alignment is not None:
                        align_name = getattr(alignment, "name", None) or str(alignment).split()[0]
                        title_aligns.append(str(align_name).upper().replace("PP_ALIGN.", ""))
                except Exception:
                    pass
            if not title and not is_cover and text_chars <= len(text):
                title = text

    notes = False
    try:
        notes = slide.has_notes_slide and bool(slide.notes_slide.notes_text_frame.text.strip())
    except Exception:
        notes = False

    return {
        "index": index,
        "isCover": is_cover,
        "title": title,
        "textChars": text_chars,
        "bullets": bullets,
        "openers": openers,
        "fontSizes": font_sizes,
        "fonts": fonts,
        "colors": colors,
        "pictures": pictures,
        "tables": tables,
        "hasNotes": notes,
        "aligns": title_aligns,
    }


def analyze(path: str) -> dict:
    presentation = Presentation(path)
    stats = {
        "file": path,
        "slideCount": len(presentation.slides),
        "widthIn": round(presentation.slide_width / 914400, 2),
        "heightIn": round(presentation.slide_height / 914400, 2),
        "slides": [],
    }
    for index, slide in enumerate(presentation.slides, start=1):
        stats["slides"].append(_slide_stats(slide, index))
    return stats


def _mode(values: list[float]) -> float | None:
    if not values:
        return None
    return Counter(values).most_common(1)[0][0]


def _pct(count: int, total: int) -> int:
    return round(count * 100 / total) if total else 0


def aggregate(analyses: list[dict]) -> dict:
    total_files = len(analyses)
    total_slides = sum(item["slideCount"] for item in analyses)
    if total_files == 0:
        return {"totalFiles": 0, "totalSlides": 0}

    title_sizes: list[float] = []
    body_sizes: list[float] = []
    fonts = Counter()
    colors = Counter()
    openers = Counter()
    all_bullets: list[int] = []
    all_chars: list[int] = []
    content_chars: list[int] = []
    content_bullets: list[int] = []
    align_counter = Counter()
    uses_toc = uses_ending = uses_notes = uses_pictures = uses_tables = 0
    coverless = 0
    title_lengths: list[int] = []
    widths = Counter()
    heights = Counter()

    for analysis in analyses:
        widths[analysis["widthIn"]] += 1
        heights[analysis["heightIn"]] += 1
        if analysis["slideCount"] <= 1:
            continue
        has_toc = False
        has_ending = False
        has_notes = False
        has_pictures = False
        has_tables = False
        first_slide = analysis["slides"][0]
        if first_slide["textChars"] < 40:
            coverless += 1
        last_slide = analysis["slides"][-1]
        if _is_ending_title(last_slide["title"]):
            has_ending = True

        for slide in analysis["slides"]:
            chars = slide["textChars"]
            bullets = slide["bullets"]
            if chars > 0:
                all_chars.append(chars)
                content_chars.append(chars)
            if bullets > 0:
                all_bullets.append(bullets)
                content_bullets.append(bullets)
            if slide["hasNotes"]:
                has_notes = True
            if slide["pictures"] > 0:
                has_pictures = True
            if slide["tables"] > 0:
                has_tables = True
            if slide["index"] > 1 and _is_likely_title(slide["title"]):
                title_lengths.append(len(slide["title"]))
            if _is_toc_title(slide["title"]):
                has_toc = True
            for size in slide["fontSizes"]:
                (title_sizes if size >= 20 else body_sizes).append(size)
            fonts.update(slide["fonts"])
            colors.update(slide["colors"])
            openers.update(slide["openers"])
            align_counter.update(slide["aligns"])

        uses_toc += 1 if has_toc else 0
        uses_ending += 1 if has_ending else 0
        uses_notes += 1 if has_notes else 0
        uses_pictures += 1 if has_pictures else 0
        uses_tables += 1 if has_tables else 0

    sample_titles = []
    for analysis in analyses[:12]:
        for slide in analysis["slides"][1:]:
            title = slide["title"].strip()
            if _is_likely_title(title) and title.lower() not in TITLE_SKIP:
                sample_titles.append(title)
                break

    return {
        "totalFiles": total_files,
        "totalSlides": total_slides,
        "avgSlides": round(total_slides / total_files, 1),
        "slideSize": {
            "widthIn": widths.most_common(1)[0][0] if widths else None,
            "heightIn": heights.most_common(1)[0][0] if heights else None,
        },
        "structure": {
            "avgBulletsPerSlide": round(sum(all_bullets) / len(all_bullets), 1) if all_bullets else None,
            "avgCharsPerSlide": round(sum(all_chars) / len(all_chars)) if all_chars else None,
            "usesTocPct": _pct(uses_toc, total_files),
            "usesEndingPct": _pct(uses_ending, total_files),
            "usesNotesPct": _pct(uses_notes, total_files),
            "usesPicturesPct": _pct(uses_pictures, total_files),
            "usesTablesPct": _pct(uses_tables, total_files),
            "sampleTitles": sample_titles[:10],
        },
        "typography": {
            "titleSizeMode": _mode(title_sizes),
            "bodySizeMode": _mode(body_sizes),
            "titleSizeRange": [min(title_sizes), max(title_sizes)] if title_sizes else None,
            "bodySizeRange": [min(body_sizes), max(body_sizes)] if body_sizes else None,
            "fonts": [name for name, _count in fonts.most_common(5)],
            "titleAlignMode": align_counter.most_common(1)[0][0] if align_counter else None,
        },
        "language": {
            "avgTitleLength": round(sum(title_lengths) / len(title_lengths), 1) if title_lengths else None,
            "commonOpeners": [opener for opener, _count in openers.most_common(8)],
            "numberingHeavy": NUMBERING_OPENER.match(openers.most_common(1)[0][0]).group(0) if openers and NUMBERING_OPENER.match(openers.most_common(1)[0][0]) else None,
        },
        "colors": [color for color, _count in colors.most_common(8)],
    }


def to_markdown(name: str, profile: dict) -> str:
    aggregate_data = profile["aggregate"]
    lines = [
        f"# PPT 风格档案：{name}",
        "",
        f"生成时间：{profile['generatedAt']}",
        f"样本：{aggregate_data.get('totalFiles', 0)} 份，共 {aggregate_data.get('totalSlides', 0)} 页，平均 {aggregate_data.get('avgSlides', 0)} 页/份",
        "",
        "## 结构习惯",
        "",
        f"- 有目录页：{aggregate_data.get('structure', {}).get('usesTocPct', 0)}%",
        f"- 有结尾/致谢页：{aggregate_data.get('structure', {}).get('usesEndingPct', 0)}%",
        f"- 使用演讲备注：{aggregate_data.get('structure', {}).get('usesNotesPct', 0)}%",
        f"- 使用图片：{aggregate_data.get('structure', {}).get('usesPicturesPct', 0)}%",
        f"- 使用表格：{aggregate_data.get('structure', {}).get('usesTablesPct', 0)}%",
        f"- 平均每条幻灯片要点数：{aggregate_data.get('structure', {}).get('avgBulletsPerSlide')}",
        f"- 平均每页字数：{aggregate_data.get('structure', {}).get('avgCharsPerSlide')}",
        "",
        "## 版式",
        "",
    ]
    typography = aggregate_data.get("typography", {})
    slide_size = aggregate_data.get("slideSize", {})
    lines.append(f"- 页面尺寸：{slide_size.get('widthIn')} × {slide_size.get('heightIn')} 英寸")
    if typography.get("titleSizeRange"):
        lines.append(f"- 标题字号：{typography['titleSizeRange'][0]}–{typography['titleSizeRange'][1]} pt（最常见 {typography.get('titleSizeMode')}）")
    if typography.get("bodySizeRange"):
        lines.append(f"- 正文字号：{typography['bodySizeRange'][0]}–{typography['bodySizeRange'][1]} pt（最常见 {typography.get('bodySizeMode')}）")
    if typography.get("fonts"):
        lines.append(f"- 常用字体：{'、'.join(typography['fonts'])}")
    if typography.get("titleAlignMode"):
        lines.append(f"- 标题对齐倾向：{typography['titleAlignMode']}")
    lines.extend(["", "## 语言", ""])
    language = aggregate_data.get("language", {})
    if language.get("avgTitleLength"):
        lines.append(f"- 标题平均长度：{language['avgTitleLength']} 字")
    if language.get("commonOpeners"):
        lines.append(f"- 常见要点开头：{'、'.join(language['commonOpeners'])}")
    lines.extend(["", "## 配色", ""])
    for color in aggregate_data.get("colors", []):
        lines.append(f"- #{color}")
    lines.extend(["", "## 样本清单", "", "| 文件 | 页数 | 标题 |", "| --- | --- | --- |"])
    for source in profile["sources"]:
        first_title = ""
        if source.get("title"):
            first_title = source["title"]
        lines.append(f"| {os.path.basename(source['file'])} | {source['slideCount']} | {first_title} |")
    if aggregate_data.get("structure", {}).get("sampleTitles"):
        lines.extend(["", "## 典型标题", ""])
        for title in aggregate_data["structure"]["sampleTitles"]:
            lines.append(f"- {title}")
    return "\n".join(lines) + "\n"


def collect_files(input_path: str) -> list[str]:
    if os.path.isfile(input_path):
        return [input_path] if input_path.lower().endswith(".pptx") else []
    pattern = os.path.join(input_path, "**", "*.pptx")
    files = [
        path
        for path in glob(pattern, recursive=True)
        if not os.path.basename(path).startswith("~$")
    ]
    return sorted(files)[:MAX_FILES]


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print(json.dumps({"error": "usage: python learn.py <input-dir-or-file> <output-base> [--name name]"}), file=sys.stderr)
        return 2

    input_path = argv[1]
    output_base = argv[2]
    name = "style-profile"
    if "--name" in argv:
        try:
            name = argv[argv.index("--name") + 1]
        except IndexError:
            pass

    if not os.path.exists(input_path):
        print(json.dumps({"error": f"input not found: {input_path}"}, ensure_ascii=False))
        return 1

    files = collect_files(input_path)
    if not files:
        print(json.dumps({"error": f"no .pptx files found under {input_path}"}, ensure_ascii=False))
        return 1

    analyses = []
    failures = []
    for path in files:
        try:
            analyses.append(analyze(path))
        except Exception as error:  # skip unreadable files rather than abort the batch
            failures.append({"file": path, "error": f"{type(error).__name__}: {error}"})

    if not analyses:
        print(json.dumps({"error": "no readable pptx files", "failures": failures}, ensure_ascii=False))
        return 1

    profile = {
        "name": name,
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "sources": [
            {
                "file": item["file"],
                "slideCount": item["slideCount"],
                "title": next((slide["title"] for slide in item["slides"] if slide["title"]), ""),
            }
            for item in analyses
        ],
        "aggregate": aggregate(analyses),
        "failures": failures,
    }

    os.makedirs(os.path.dirname(os.path.abspath(output_base)), exist_ok=True)
    with open(output_base + ".json", "w", encoding="utf-8") as handle:
        json.dump(profile, handle, ensure_ascii=False, indent=2)
    with open(output_base + ".md", "w", encoding="utf-8") as handle:
        handle.write(to_markdown(name, profile))

    summary = {
        "name": name,
        "jsonPath": output_base + ".json",
        "markdownPath": output_base + ".md",
        "files": len(analyses),
        "failed": len(failures),
        "totalSlides": profile["aggregate"]["totalSlides"],
        "aggregate": profile["aggregate"],
    }
    print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main(sys.argv))
