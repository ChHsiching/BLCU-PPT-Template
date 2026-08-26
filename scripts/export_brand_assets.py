"""Export the manifest's brand-layer media into the web scaffold.

Usage:
    python scripts/export_brand_assets.py [--manifest PATH] [--extracted-dir DIR]
        [--web-public-brand DIR]

Reads templates/<id>/manifest.json brand_layer, collects every element's
media filename, and copies extracted/media/<name> to
assets/web-template/public/brand/<name> (the names are kept verbatim — the
manifest's media field is the single mapping; nothing is renamed here).
Idempotent; exits 2 when the manifest lacks a brand_layer or a media file is
missing from extracted/ (a half-shipped brand layer would fail the web QA
gate with 404 assets, so the export refuses to produce one silently).
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def brand_media(manifest: dict) -> list[str]:
    """Collect media filenames from brand_layer, preserving manifest order."""
    bl = manifest.get("brand_layer")
    if not isinstance(bl, dict):
        raise SystemExit("error: manifest has no brand_layer section")

    names: list[str] = []
    def collect(node) -> None:
        if isinstance(node, dict):
            media = node.get("media")
            if isinstance(media, str):
                if media not in names:
                    names.append(media)
            else:
                for value in node.values():
                    collect(value)
        elif isinstance(node, list):
            for item in node:
                collect(item)

    collect(bl.get("content"))
    collect(bl.get("cover"))
    if not names:
        raise SystemExit("error: brand_layer declares no media files")
    return names


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Copy brand-layer media from extracted/ into the web scaffold.")
    parser.add_argument("--manifest", type=Path,
                        default=REPO / "templates" / "blcu-report" / "manifest.json")
    parser.add_argument("--extracted-dir", type=Path,
                        default=REPO / "templates" / "blcu-report" / "extracted")
    parser.add_argument("--web-public-brand", type=Path,
                        default=REPO / "assets" / "web-template" / "public" / "brand")
    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="replace")

    try:
        manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        print(f"error: cannot read manifest {args.manifest}: {exc}", file=sys.stderr)
        return 2

    names = brand_media(manifest)
    args.web_public_brand.mkdir(parents=True, exist_ok=True)

    stale = {p.name for p in args.web_public_brand.iterdir() if p.is_file()} - set(names)
    for name in stale:  # a renamed media field must not leave orphans behind
        (args.web_public_brand / name).unlink()

    for name in names:
        src = args.extracted_dir / "media" / name
        if not src.is_file():
            print(f"error: brand media not found in extracted/: {src}", file=sys.stderr)
            return 2
        shutil.copyfile(src, args.web_public_brand / name)

    print(f"exported {len(names)} brand asset(s) -> {args.web_public_brand}")
    for name in names:
        print(f"  {name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
