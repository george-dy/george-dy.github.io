#!/usr/bin/env python3
"""tools/compress-assets.py — Stage 2 of the bank build.

Reads question-bank.staging.json (images point at assets-raw/), compresses and
downscales every image into assets/ (browser-friendly sizes, aspect ratio
preserved), rewrites image srcs, and writes the final:
  question-bank.json   (canonical published bank)
  question-bank.js     (same data as `window.EXAM_BANK = {...}` so the app also
                        works when index.html is opened via file:// double-click)

No image content is altered beyond resize/re-encode; no images are dropped.
"""
import json
import os
import sys

from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, "assets-raw")
OUT = os.path.join(ROOT, "assets")
MAX_DIM = 1600          # max width/height in px
JPEG_QUALITY = 84
PNG_KEEP_BELOW = 350_000  # keep PNGs smaller than this as-is (charts w/ text)

import shutil
if os.path.isdir(OUT):
    shutil.rmtree(OUT)
os.makedirs(OUT, exist_ok=True)

mapping = {}   # raw name -> final name
stats = {"raw_bytes": 0, "out_bytes": 0, "converted": 0, "kept": 0}


def process(raw_name: str) -> str:
    if raw_name in mapping:
        return mapping[raw_name]
    src_path = os.path.join(RAW, raw_name)
    if not os.path.exists(src_path):
        print(f"  !! missing raw asset {raw_name}", file=sys.stderr)
        mapping[raw_name] = None
        return None
    raw_size = os.path.getsize(src_path)
    stats["raw_bytes"] += raw_size
    base, ext = os.path.splitext(raw_name)
    ext_l = ext.lower()

    try:
        im = Image.open(src_path)
        im.load()
    except Exception as e:
        print(f"  !! cannot open {raw_name}: {e}; copying as-is", file=sys.stderr)
        final_name = raw_name
        with open(src_path, "rb") as fh, open(os.path.join(OUT, final_name), "wb") as out:
            out.write(fh.read())
        stats["out_bytes"] += raw_size
        stats["kept"] += 1
        mapping[raw_name] = final_name
        return final_name

    has_alpha = (im.mode in ("RGBA", "LA")) or (im.mode == "P" and "transparency" in im.info)
    big = max(im.size) > MAX_DIM or raw_size > 150_000

    if not big:
        # small file: copy unchanged (keeps charts crisp)
        final_name = raw_name
        with open(src_path, "rb") as fh, open(os.path.join(OUT, final_name), "wb") as out:
            out.write(fh.read())
        stats["out_bytes"] += raw_size
        stats["kept"] += 1
        mapping[raw_name] = final_name
        return final_name

    if max(im.size) > MAX_DIM:
        im.thumbnail((MAX_DIM, MAX_DIM), Image.LANCZOS)

    def is_photographic(img):
        """Heuristic: many distinct colors -> photograph; few -> chart/text."""
        small = img.convert("RGB").resize((64, 64))
        return len(small.getcolors(maxcolors=4097) or []) > 1500

    if ext_l in (".png", ".bmp", ".tif", ".tiff") and not has_alpha:
        if is_photographic(im) or raw_size > 500_000:
            final_name = base + ".jpg"
            im.convert("RGB").save(os.path.join(OUT, final_name), "JPEG",
                                   quality=JPEG_QUALITY, optimize=True, progressive=True)
            stats["converted"] += 1
        else:
            final_name = raw_name
            im.save(os.path.join(OUT, final_name), optimize=True)
            stats["kept"] += 1
    elif ext_l in (".jpg", ".jpeg"):
        final_name = raw_name
        im.convert("RGB").save(os.path.join(OUT, final_name), "JPEG",
                               quality=JPEG_QUALITY, optimize=True, progressive=True)
        stats["converted"] += 1
    else:
        # keep format (alpha PNGs etc.), optimized
        final_name = raw_name
        im.save(os.path.join(OUT, final_name), optimize=True)
        stats["kept"] += 1

    out_size = os.path.getsize(os.path.join(OUT, final_name))
    # if re-encode didn't help, keep the original bytes
    if out_size >= raw_size and not has_alpha:
        with open(src_path, "rb") as fh, open(os.path.join(OUT, raw_name), "wb") as out:
            out.write(fh.read())
        if final_name != raw_name and os.path.exists(os.path.join(OUT, final_name)):
            os.remove(os.path.join(OUT, final_name))
        final_name = raw_name
        out_size = raw_size
    stats["out_bytes"] += out_size
    mapping[raw_name] = final_name
    return final_name


def main():
    staging = os.path.join(ROOT, "question-bank.staging.json")
    with open(staging, "r", encoding="utf-8") as fh:
        bank = json.load(fh)

    n_imgs = 0
    for q in bank["questions"]:
        for im in q.get("images") or []:
            src = im.get("src") or ""
            if src.startswith("assets-raw/"):
                raw_name = src.split("/", 1)[1]
                final_name = process(raw_name)
                if final_name:
                    im["src"] = "assets/" + final_name
                    n_imgs += 1
                else:
                    im["error"] = "asset-missing-after-compress"
    print(f"processed {len(mapping)} unique assets for {n_imgs} image references")
    print(f"raw total: {stats['raw_bytes']/1e6:.1f} MB -> out total: {stats['out_bytes']/1e6:.1f} MB "
          f"(converted {stats['converted']}, kept {stats['kept']})")

    final_json = os.path.join(ROOT, "question-bank.json")
    with open(final_json, "w", encoding="utf-8") as fh:
        json.dump(bank, fh, ensure_ascii=False, separators=(",", ":"))
    size = os.path.getsize(final_json)
    print(f"question-bank.json: {size/1e6:.2f} MB")

    final_js = os.path.join(ROOT, "question-bank.js")
    with open(final_js, "w", encoding="utf-8") as fh:
        fh.write("/* question-bank.js — mirror of question-bank.json for file:// fallback. */\n")
        fh.write("window.EXAM_BANK = ")
        json.dump(bank, fh, ensure_ascii=False, separators=(",", ":"))
        fh.write(";\n")
    print(f"question-bank.js: {os.path.getsize(final_js)/1e6:.2f} MB")

    # sanity: all referenced assets exist
    missing = []
    for q in bank["questions"]:
        for im in q.get("images") or []:
            src = im.get("src") or ""
            if src.startswith("assets/"):
                if not os.path.exists(os.path.join(ROOT, src)):
                    missing.append(src)
    if missing:
        print("MISSING ASSETS:", missing[:10], file=sys.stderr)
        sys.exit(1)
    print("all referenced assets exist ✓")


if __name__ == "__main__":
    main()
