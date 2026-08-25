"""Scaffold a self-contained web presentation project from a deck.json.

Usage:
    python scripts/scaffold_web.py <deck.json> -o <out-dir> [--templates-dir DIR] [--manifest PATH] [--force]

Copies the preset Vite+React scaffold (assets/web-template; no interactive
wizard) into <out-dir>, then injects: src/deck.json (image paths rewritten to
material/images/<basename>), src/manifest.json (the deck's template manifest —
regions and typography stay single-sourced from it, as in renderer-pptx), and
every deck-referenced image copied flat into public/material/images/ (two deck
images sharing a basename from different source files are refused — rename
one; the same file referenced twice is fine). The template's sample deck,
manifest and material images are wholesale-replaced, so a scaffolded project
never carries sample residue. The deck is validated first
(scripts/validate_deck.py, pure function); an invalid deck never produces a
project. The output directory must not exist (an empty directory is fine);
--force deletes an existing one first — the scaffold source and the templates
tree (gitignored .pptx original inside) can never be the target.

After scaffolding: cd <out-dir> && npm install && npm run dev.

Exit codes: 0 scaffolded, 1 deck failed validation, 2 usage / IO error.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import validate_deck as vd  # noqa: E402

REPO = Path(__file__).resolve().parent.parent
DEFAULT_SCAFFOLD = REPO / "assets" / "web-template"

# never copied out of the scaffold source (node/build state, OS noise)
COPY_IGNORE = shutil.ignore_patterns(
    "node_modules", "dist", ".vite", ".cache", "*.log", ".DS_Store", "Thumbs.db",
)


class ScaffoldError(Exception):
    """Unrecoverable scaffold condition (usage, IO, image conflicts)."""


@dataclass
class ScaffoldStats:
    pages: int = 0
    images: int = 0


def collect_images(deck: dict, deck_dir: Path) -> list[tuple[str, Path]]:
    """Return [(deck_path, resolved_source)] for every image block.

    Raises ScaffoldError when two different source files share a basename:
    web images are flattened into public/material/images/, and a silent
    overwrite would render one of them with the other's pixels. Names are
    compared casefold()ed — Windows/NTFS and macOS targets are
    case-insensitive, so Pic.png and pic.png collide there too.
    """
    seen: dict[str, Path] = {}
    out: list[tuple[str, Path]] = []
    for i, page in enumerate(deck["pages"]):
        for j, block in enumerate(page["blocks"]):
            if block.get("type") != "image":
                continue
            resolved = vd.resolve_image_path(block["path"], deck_dir).resolve()
            if not resolved.is_file():
                # validation catches this first; keep the guard for API callers
                raise ScaffoldError(
                    f"pages[{i}].blocks[{j}].path: image file not found: {block['path']}"
                )
            key = resolved.name.casefold()
            if key in seen and seen[key] != resolved:
                raise ScaffoldError(
                    f"two deck images share the basename '{resolved.name}' "
                    f"(case-insensitive): {seen[key]} and {resolved}; web images "
                    f"are flattened into public/material/images/ — rename one "
                    f"of them"
                )
            seen[key] = resolved
            out.append((block["path"], resolved))
    return out


def build_web_deck(deck: dict, images: list[tuple[str, Path]]) -> dict:
    """Deep-copy the deck with image paths rewritten to material/images/<name>."""
    mapping = {raw: f"material/images/{resolved.name}" for raw, resolved in images}
    web = json.loads(json.dumps(deck, ensure_ascii=False))
    for page in web["pages"]:
        for block in page["blocks"]:
            if block.get("type") == "image":
                block["path"] = mapping[block["path"]]
    return web


def _prepare_out_dir(out_dir: Path, scaffold_dir: Path, force: bool,
                     templates_dir: Path | None = None) -> None:
    out_dir = out_dir.resolve()
    # the scaffold source and the templates tree (the gitignored .pptx original
    # lives there, unrecoverable once deleted) can never be the target — not
    # even with --force, mirroring renderer-pptx's template-overwrite refusal
    protected = [scaffold_dir.resolve()]
    if templates_dir is not None:
        protected.append(Path(templates_dir).resolve())
    for root in protected:
        if out_dir == root or root.is_relative_to(out_dir) or out_dir.is_relative_to(root):
            raise ScaffoldError(
                f"output directory {out_dir} overlaps protected directory {root}; "
                f"refusing to touch it"
            )
    if out_dir.exists():
        if any(out_dir.iterdir()) and not force:
            raise ScaffoldError(
                f"output directory {out_dir} exists and is not empty; "
                f"pass --force to replace it"
            )
        if force and any(out_dir.iterdir()):
            shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)


def scaffold_deck(deck: dict, manifest_src: Path, deck_dir: Path,
                  out_dir: Path, scaffold_dir: Path = DEFAULT_SCAFFOLD,
                  templates_dir: Path | None = None,
                  force: bool = False) -> ScaffoldStats:
    """Scaffold a validated deck dict into a runnable web project at out_dir."""
    scaffold_dir = Path(scaffold_dir)
    if not scaffold_dir.is_dir():
        raise ScaffoldError(f"scaffold source not found: {scaffold_dir}")
    if not (scaffold_dir / "package.json").is_file():
        raise ScaffoldError(f"scaffold source lacks package.json: {scaffold_dir}")

    images = collect_images(deck, deck_dir)
    web_deck = build_web_deck(deck, images)

    _prepare_out_dir(Path(out_dir), scaffold_dir, force, templates_dir)
    shutil.copytree(scaffold_dir, out_dir, dirs_exist_ok=True, ignore=COPY_IGNORE)

    (out_dir / "src" / "deck.json").write_text(
        json.dumps(web_deck, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    shutil.copyfile(manifest_src, out_dir / "src" / "manifest.json")

    material = out_dir / "public" / "material"
    if material.exists():  # drop the sample material wholesale
        shutil.rmtree(material)
    if images:
        images_dir = material / "images"
        images_dir.mkdir(parents=True, exist_ok=True)
        for _, resolved in images:
            shutil.copyfile(resolved, images_dir / resolved.name)

    return ScaffoldStats(pages=len(deck["pages"]), images=len(images))


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Scaffold a web presentation project from a deck.json.")
    parser.add_argument("deck", help="path to deck.json")
    parser.add_argument("-o", "--out", required=True, help="output project directory")
    parser.add_argument("--templates-dir", type=Path,
                        default=Path(__file__).resolve().parent.parent / "templates",
                        help="directory holding <template-id>/ (default: <repo>/templates)")
    parser.add_argument("--manifest", type=Path,
                        help="explicit manifest path (overrides template resolution)")
    parser.add_argument("--scaffold", type=Path, default=DEFAULT_SCAFFOLD,
                        help="scaffold source directory (default: <repo>/assets/web-template)")
    parser.add_argument("--force", action="store_true",
                        help="replace an existing non-empty output directory")
    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="replace")

    deck_path = Path(args.deck)
    try:
        raw = deck_path.read_text(encoding="utf-8")
        deck = vd.parse_deck(raw)
    except (OSError, UnicodeError) as exc:
        print(f"error: cannot read deck: {exc}", file=sys.stderr)
        return 2
    except ValueError as exc:
        print(f"error: deck is not valid JSON: {exc}", file=sys.stderr)
        return 2
    deck_dir = deck_path.parent

    manifest_path, error_finding = vd.resolve_manifest_path(
        deck, args.templates_dir, args.manifest)
    if error_finding is not None or manifest_path is None:
        print(vd.build_report([error_finding] if error_finding else [], 0, "text"))
        return 1
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        print(f"error: cannot load manifest {manifest_path}: {exc}", file=sys.stderr)
        return 2

    findings = vd.validate_deck(deck, manifest, deck_dir)
    if findings:
        print(vd.build_report(findings, len(deck["pages"]), "text"))
        return 1

    try:
        stats = scaffold_deck(deck, manifest_path, deck_dir, Path(args.out),
                              scaffold_dir=args.scaffold,
                              templates_dir=args.templates_dir, force=args.force)
    except ScaffoldError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"error: cannot write project: {exc}", file=sys.stderr)
        return 2

    out = Path(args.out)
    print(f"scaffolded {stats.pages} page(s), {stats.images} image(s) -> {out}")
    print(f"next: cd {out} && npm install && npm run dev")
    return 0


if __name__ == "__main__":
    sys.exit(main())
