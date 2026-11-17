"""Static PPTX checks, render evidence and explicit visual-review records.

Static geometry checks are approximate. They never constitute visual review.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import posixpath
import sys
import tempfile
from types import SimpleNamespace
from urllib.parse import unquote
import xml.etree.ElementTree as ET
from zipfile import ZipFile, BadZipFile

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def file_hash(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def report_path(deck):
    return str(Path(deck).with_suffix(".quality.json"))


def write_json(path, value):
    directory = os.path.dirname(os.path.abspath(path))
    os.makedirs(directory, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(prefix=".ppt-forge-", suffix=".json", dir=directory)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2, allow_nan=False)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def integrity_findings(path):
    findings = []
    def error(code, message, part=None):
        findings.append({"severity": "error", "code": code, "message": message, **({"part": part} if part else {})})
    try:
        with ZipFile(path) as archive:
            names = archive.namelist()
            members = set(names)
            if len(names) != len(members):
                error("DUPLICATE_PART", "The archive contains duplicate part names")
            damaged = archive.testzip()
            if damaged:
                error("CORRUPT_PART", "ZIP checksum failed", damaged)
            for required in ("[Content_Types].xml", "ppt/presentation.xml", "_rels/.rels"):
                if required not in members:
                    error("MISSING_PART", "Required PPTX part is missing", required)
            for name in names:
                if not name.endswith((".xml", ".rels")):
                    continue
                try:
                    root = ET.fromstring(archive.read(name))
                except ET.ParseError as exc:
                    error("INVALID_XML", str(exc), name)
                    continue
                if name.endswith(".rels"):
                    base = posixpath.dirname(posixpath.dirname(name))
                    for relation in root.findall("{" + REL_NS + "}Relationship"):
                        if relation.get("TargetMode") == "External":
                            continue
                        target = unquote(relation.get("Target", "").split("#")[0])
                        resolved = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(base, target))
                        if resolved not in members:
                            error("BROKEN_RELATIONSHIP", f"Missing internal target: {resolved}", name)
                if name.startswith("ppt/slides/") and name.endswith(".xml"):
                    ids = [element.get("id") for element in root.iter("{" + P_NS + "}cNvPr")]
                    if len(ids) != len(set(ids)):
                        error("DUPLICATE_SHAPE_ID", "Slide contains duplicate shape IDs", name)
    except (BadZipFile, OSError) as exc:
        error("INVALID_PPTX", str(exc))
    return findings


def _text_bbox(shape, metrics, EMU):
    """Approximate used text area instead of treating the whole box as ink."""
    lines = shape.text_frame.text.splitlines() or [""]
    units = max(sum(0.55 if ord(ch) < 128 else 1 for ch in line) for line in lines)
    font = float(metrics.get("fontSizePt") or 18)
    used_width = min(shape.width, max(EMU * 0.02, int(units * font / 72 * EMU / 0.92)))
    used_height = min(shape.height, max(EMU * 0.02, int(metrics.get("requiredHeightInApprox", 0.02) * EMU)))
    return shape.left, shape.top, used_width, used_height


def _selector_matches(item, selectors):
    return any((isinstance(selector, str) and item.get("shapeName") == selector)
               or (isinstance(selector, dict) and (selector.get("id") == item.get("shapeId") or selector.get("name") == item.get("shapeName")))
               for selector in (selectors or []))


def analyze_deck(path, quality=None):
    quality = quality or {}
    findings = integrity_findings(path)
    count = 0
    if not any(item["severity"] == "error" for item in findings):
        from pptx import Presentation
        from inspection import iter_shapes, text_metrics, EMU
        prs = Presentation(path)
        count = len(prs.slides)
        if not count:
            findings.append({"severity": "error", "code": "EMPTY_DECK", "message": "Deck contains no slides"})
        for page, slide in enumerate(prs.slides, 1):
            texts = []
            text_boxes = {}
            for shape in iter_shapes(slide.shapes):
                if shape.has_table:
                    grouped = shape._element.getparent().tag.endswith("}grpSp")
                    if not grouped and (shape.left < 0 or shape.top < 0 or shape.left + shape.width > prs.slide_width or shape.top + shape.height > prs.slide_height):
                        findings.append({"severity": "warning", "code": "TABLE_OUTSIDE_SLIDE", "page": page,
                            "shapeId": shape.shape_id, "message": "Table extends outside the slide"})
                    for row_index, row in enumerate(shape.table.rows):
                        for col_index, cell in enumerate(row.cells):
                            if not cell.text.strip() or cell.is_spanned:
                                continue
                            cell_width = sum(shape.table.columns[i].width for i in range(col_index, min(len(shape.table.columns), col_index + cell.span_width)))
                            cell_height = sum(shape.table.rows[i].height for i in range(row_index, min(len(shape.table.rows), row_index + cell.span_height)))
                            metrics = text_metrics(SimpleNamespace(text_frame=cell.text_frame, width=cell_width, height=cell_height))
                            if metrics["requiredHeightInApprox"] > metrics["availableHeightIn"] * 1.08:
                                findings.append({"severity": "warning", "code": "TABLE_TEXT_CAPACITY", "page": page,
                                    "shapeId": shape.shape_id, "cell": [row_index, col_index], "message": "Estimated cell text exceeds its height; inspect the render", "metrics": metrics})
                if not shape.has_text_frame or not shape.text_frame.text.strip():
                    continue
                base = {"severity": "warning", "page": page, "shapeId": shape.shape_id, "shapeName": shape.name}
                # Group coordinates use a different coordinate space; skip bounds
                # and overlap comparisons for grouped objects rather than misreport.
                grouped = shape._element.getparent().tag.endswith("}grpSp")
                if not grouped:
                    if shape.left < -EMU * 0.02 or shape.top < -EMU * 0.02 or shape.left + shape.width > prs.slide_width + EMU * 0.02 or shape.top + shape.height > prs.slide_height + EMU * 0.02:
                        findings.append({**base, "code": "TEXT_OUTSIDE_SLIDE", "message": "Text box extends outside the slide"})
                    texts.append(shape)
                metrics = text_metrics(shape)
                if metrics["requiredHeightInApprox"] > metrics["availableHeightIn"] * 1.08:
                    findings.append({**base, "code": "TEXT_CAPACITY", "message": "Estimated text height exceeds the text box; inspect its render", "metrics": metrics})
                text_boxes[shape.shape_id] = _text_bbox(shape, metrics, EMU)
            for index, left in enumerate(texts):
                for right in texts[index + 1:]:
                    left_box = text_boxes.get(left.shape_id, (left.left, left.top, left.width, left.height))
                    right_box = text_boxes.get(right.shape_id, (right.left, right.top, right.width, right.height))
                    width = min(left_box[0] + left_box[2], right_box[0] + right_box[2]) - max(left_box[0], right_box[0])
                    height = min(left_box[1] + left_box[3], right_box[1] + right_box[3]) - max(left_box[1], right_box[1])
                    if width > EMU * 0.04 and height > EMU * 0.04:
                        if not _selector_matches({"shapeId": left.shape_id, "shapeName": left.name}, quality.get("ignoreOverlaps")) and not _selector_matches({"shapeId": right.shape_id, "shapeName": right.name}, quality.get("ignoreOverlaps")):
                            findings.append({"severity": "warning", "code": "TEXT_OVERLAP", "page": page,
                                "shapeIds": [left.shape_id, right.shape_id], "shapeNames": [left.name, right.name], "message": "Estimated rendered text overlaps; inspect the render"})
    report = {"formatVersion": 1, "deck": os.path.abspath(path), "deckSha256": file_hash(path), "slideCount": count,
              "static": {"findings": findings}, "render": {"status": "not-requested", "missingFonts": []},
              "visualReview": {"status": "pending", "findings": []},
              "contentReview": {"status": "pending", "findings": []}}
    return update_status(report)


def update_status(report):
    report.setdefault("contentReview", {"status": "pending", "findings": []})
    errors = any(item["severity"] == "error" for item in report["static"]["findings"])
    visual = report["visualReview"]["status"]
    content = report["contentReview"]["status"]
    report["status"] = "blocked" if errors or visual == "failed" or content == "failed" else "passed" if visual == "passed" and content == "passed" and report["render"]["status"] == "complete" else "needs-review"
    report["readyForDelivery"] = report["status"] == "passed"
    return report


def current_report(deck, require_existing=False):
    path = report_path(deck)
    try:
        with open(path, encoding="utf-8") as handle:
            report = json.load(handle)
        if report["deckSha256"] == file_hash(deck):
            return report
    except (OSError, ValueError, KeyError):
        pass
    if require_existing:
        raise ValueError("Review evidence is missing or stale; render the current deck again")
    return analyze_deck(deck)


def render_token(deck_hash, images):
    payload = json.dumps({"deckSha256": deck_hash, "images": images}, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def record_render(deck, evidence):
    report = current_report(deck)
    if evidence.get("deckSha256") != report["deckSha256"]:
        raise ValueError("Deck changed during rendering; render the current version again")
    report["visualReview"] = {"status": "pending", "findings": []}
    report["contentReview"] = {"status": "pending", "findings": []}
    images = []
    pages = set()
    for item in evidence.get("images", []):
        page = item.get("page")
        if type(page) is not int or page < 1 or page > report["slideCount"] or page in pages:
            raise ValueError("Render manifest contains invalid or duplicate page numbers")
        image = os.path.abspath(item["path"])
        images.append({"page": page, "path": image, "sha256": file_hash(image)})
        pages.add(page)
    images.sort(key=lambda item: item["page"])
    status = "failed" if evidence.get("error") else "complete" if len(pages) == report["slideCount"] and len(pages) > 0 else "partial"
    report["render"] = {"status": status, "deckSha256": report["deckSha256"], "images": images,
                         "missingFonts": evidence.get("missingFonts", []), **({"error": evidence["error"]} if evidence.get("error") else {})}
    report["reviewToken"] = render_token(report["deckSha256"], images)
    update_status(report)
    write_json(report_path(deck), report)
    return report


def record_review(deck, review):
    report = current_report(deck, require_existing=True)
    if report["render"]["status"] != "complete":
        raise ValueError("Render every page of the current deck before submitting a visual review")
    if not review.get("reviewToken") or review["reviewToken"] != report.get("reviewToken"):
        raise ValueError("Review token is missing or stale; use the token returned by the latest full render")
    for image in report["render"]["images"]:
        if file_hash(image["path"]) != image["sha256"]:
            raise ValueError("Preview changed after rendering; render again before review")
    findings = review.get("findings")
    if not isinstance(findings, list):
        raise ValueError("review.findings must be an array (use [] only after inspecting all pages)")
    for item in findings:
        if not isinstance(item, dict) or item.get("severity") not in ("warning", "error") or type(item.get("page")) is not int or not 1 <= item["page"] <= report["slideCount"] or not isinstance(item.get("message"), str) or not item["message"].strip():
            raise ValueError("Every finding needs a valid page, severity (warning/error), and non-empty message")
    review_type = str(review.get("reviewType") or "visual")
    if review_type not in ("visual", "content"):
        raise ValueError("reviewType must be visual or content")
    target = "contentReview" if review_type == "content" else "visualReview"
    report[target] = {"status": "failed" if any(item["severity"] == "error" for item in findings) else "passed",
                      "findings": findings, "reviewer": str(review.get("reviewer") or "host-agent")}
    update_status(report)
    write_json(report_path(deck), report)
    return report


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("action", choices=("check", "render", "review"))
    parser.add_argument("deck")
    parser.add_argument("payload", nargs="?", default="-")
    args = parser.parse_args()
    try:
        if args.action == "check":
            report = current_report(args.deck)
            if report["render"]["status"] == "complete":
                try:
                    valid = all(file_hash(item["path"]) == item["sha256"] for item in report["render"]["images"])
                except OSError:
                    valid = False
                if not valid:
                    report["render"]["status"] = "failed"
                    report["render"]["error"] = "Preview evidence changed; render again"
                    report["visualReview"] = {"status": "pending", "findings": []}
            update_status(report)
            write_json(report_path(args.deck), report)
        else:
            if args.payload == "-":
                payload = json.load(sys.stdin)
            else:
                with open(args.payload, encoding="utf-8") as handle:
                    payload = json.load(handle)
            if not isinstance(payload, dict):
                raise ValueError("payload must be an object")
            report = (record_render if args.action == "render" else record_review)(args.deck, payload)
        print(json.dumps({"reportPath": report_path(args.deck), **report}, ensure_ascii=False))
        return 0
    except Exception as error:
        print(json.dumps({"error": str(error), "code": "QUALITY_CHECK_FAILED"}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
