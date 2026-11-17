#!/usr/bin/env python
"""Extract reusable background and decoration art from sample decks.

Usage:
    python extract_assets.py <input-dir-or-file> <output-dir> [--name profile]

Writes <output-dir>/<name>/assets/... and <output-dir>/<name>/manifest.json.
"""

from __future__ import annotations

import hashlib
import io
import json
import os
import sys
from glob import glob

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

MAX_FILES = 200
MAX_BACKGROUNDS = 40
MAX_DECORATIONS = 40
MAX_BYTES = 6 * 1024 * 1024
MIN_BACKGROUND_WIDTH = 1000


def collect_files(input_path: str) -> list[str]:
    if os.path.isfile(input_path):
        return [input_path] if input_path.lower().endswith(".pptx") else []
    files = [
        path
        for path in glob(os.path.join(input_path, "**", "*.pptx"), recursive=True)
        if not os.path.basename(path).startswith("~$")
    ]
    return sorted(files)[:MAX_FILES]


def iter_pictures(slide):
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_pictures_in_group(shape)
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            yield shape


def iter_pictures_in_group(group):
    for child in group.shapes:
        if child.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_pictures_in_group(child)
        elif child.shape_type == MSO_SHAPE_TYPE.PICTURE:
            yield child


def classify(width: int, height: int) -> str:
    if height == 0:
        return "photo"
    aspect = width / height
    if 1.55 <= aspect <= 1.85 and width >= MIN_BACKGROUND_WIDTH:
        return "background"
    if width <= 600 and height <= 600:
        return "decoration"
    return "photo"


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print(json.dumps({"error": "usage: python extract_assets.py <input> <output-dir> [--name profile]"}), file=sys.stderr)
        return 2
    input_path = argv[1]
    output_root = argv[2]
    name = "assets"
    if "--name" in argv:
        try:
            name = argv[argv.index("--name") + 1]
        except IndexError:
            pass
    target = os.path.join(output_root, name)
    assets_dir = os.path.join(target, "assets")
    os.makedirs(assets_dir, exist_ok=True)

    files = collect_files(input_path)
    if not files:
        print(json.dumps({"error": f"no .pptx files under {input_path}"}, ensure_ascii=False))
        return 1

    seen: dict[str, dict] = {}
    skipped = 0
    for path in files:
        try:
            presentation = Presentation(path)
        except Exception:
            skipped += 1
            continue
        for slide in presentation.slides:
            for shape in iter_pictures(slide):
                try:
                    blob = shape.image.blob
                    if len(blob) > MAX_BYTES:
                        continue
                    with Image.open(io.BytesIO(blob)) as image:
                        width, height = image.size
                except Exception:
                    continue
                digest = hashlib.sha1(blob).hexdigest()
                if digest in seen:
                    seen[digest]["hits"] += 1
                    seen[digest]["source"].add(os.path.basename(path))
                    continue
                try:
                    ext = shape.image.ext.lower()
                except Exception:
                    # MPO and other container variants report no python-pptx format name
                    ext = "jpg"
                if ext not in ("png", "jpg", "jpeg", "webp", "gif", "bmp", "tif", "tiff"):
                    ext = "png"
                kind = classify(width, height)
                seen[digest] = {
                    "digest": digest,
                    "ext": ext,
                    "width": width,
                    "height": height,
                    "kind": kind,
                    "hits": 1,
                    "blob": blob,
                    "source": {os.path.basename(path)},
                }

    backgrounds = [item for item in seen.values() if item["kind"] == "background"]
    decorations = [item for item in seen.values() if item["kind"] == "decoration"]
    backgrounds.sort(key=lambda item: (-item["hits"], -(item["width"] * item["height"])))
    decorations.sort(key=lambda item: (-item["hits"], -(item["width"] * item["height"])))

    manifest = {"name": name, "sources": len(files), "backgrounds": [], "decorations": []}
    for kind, items, limit in (
        ("background", backgrounds, MAX_BACKGROUNDS),
        ("decoration", decorations, MAX_DECORATIONS),
    ):
        for index, item in enumerate(items[:limit]):
            filename = f"{kind}-{index + 1:02d}.{item['ext']}"
            file_path = os.path.join(assets_dir, filename)
            with open(file_path, "wb") as handle:
                handle.write(item["blob"])
            manifest[kind + "s"].append(
                {
                    "file": filename,
                    "width": item["width"],
                    "height": item["height"],
                    "hits": item["hits"],
                    "sources": sorted(item["source"])[:5],
                    "path": os.path.join(assets_dir, filename),
                }
            )

    manifest_path = os.path.join(target, "manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as handle:
        json.dump(manifest, handle, ensure_ascii=False, indent=2)

    print(
        json.dumps(
            {
                "name": name,
                "manifest": manifest_path,
                "files": len(files),
                "uniqueImages": len(seen),
                "backgrounds": len(manifest["backgrounds"]),
                "decorations": len(manifest["decorations"]),
                "skipped": skipped,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main(sys.argv))
