"""Validate a deck.json against a template manifest: schema shape, archetype
existence, capacity budgets, image path existence.

Usage:
    python scripts/validate_deck.py [deck] [--templates-dir DIR] [--manifest PATH] [--format text|json]

`deck` is a path to a deck.json file, or `-` (default) to read stdin. Image
paths in the deck resolve relative to the deck file's directory (the current
directory for stdin). The manifest resolves from the deck's `template` field as
<templates-dir>/<template>/manifest.json unless --manifest is given.

Exit codes: 0 valid, 1 validation findings, 2 usage / IO / parse error.
The report goes to stdout: one `<path>: [<code>] <message>` line per finding
plus a summary (text format), or a {"valid", "page_count", "findings"} object
(json format). Findings carry a JSON path into the deck so gates and humans can
locate every defect. The core check is the pure function validate_deck(deck,
manifest, image_root) -> list[Finding]; gates and renderers import it directly.
"""
from __future__ import annotations

import argparse
import json
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path

EPS = 1e-9

# ---------------------------------------------------------------------------
# char width (manifest budget_semantics: CJK 等宽计, 半角字符计 0.5)
# ---------------------------------------------------------------------------


def text_width(s: str) -> float:
    """CJK-width of a string: East Asian W/F chars count 1, everything else 0.5."""
    units = sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in s)
    return units / 2.0


def fmt_width(w: float) -> str:
    return f"{w:g}"


# ---------------------------------------------------------------------------
# schema tables
# ---------------------------------------------------------------------------

BLOCK_SPECS = {
    "title": {"required": ("text",), "optional": ()},
    "subhead": {"required": ("text",), "optional": ()},
    "text": {"required": ("text",), "optional": ()},
    "list": {"required": ("items",), "optional": ()},
    "formula": {"required": ("latex",), "optional": ()},
    "image": {"required": ("path",), "optional": ("caption",)},
}
BUDGET_KEYS = (
    "title_max_chars",
    "text_blocks_max",
    "text_block_max_chars",
    "text_total_max_chars",
    "subhead_max_chars",
    "list_items_max",
    "list_item_max_chars",
    "formulas_max",
    "images_min",
    "images_max",
    "caption_max_chars",
)

META_FIELDS = ("presenter", "date")


@dataclass(frozen=True)
class Finding:
    path: str
    code: str
    message: str


# ---------------------------------------------------------------------------
# field-level helpers
# ---------------------------------------------------------------------------


def type_name(value) -> str:
    return type(value).__name__


def check_known_keys(obj: dict, allowed: set[str], path: str, findings: list) -> None:
    for key in obj:
        if key not in allowed:
            findings.append(
                Finding(
                    f"{path}.{key}" if path else key,
                    "schema.unknown_field",
                    f"unknown field '{key}' (allowed: {', '.join(sorted(allowed))})",
                )
            )


def check_string(container: dict, key: str, path: str, findings: list, *, line: bool) -> bool:
    """Check a required string field; True when the field is usable."""
    value = container.get(key)
    if not isinstance(value, str):
        findings.append(
            Finding(f"{path}.{key}", "schema.type", f"expected string, got {type_name(value)}")
        )
        return False
    if not value.strip():
        findings.append(Finding(f"{path}.{key}", "schema.empty_field", "empty or whitespace-only string"))
        return False
    if line and any(ch in value for ch in "\n\r\t"):
        findings.append(
            Finding(
                f"{path}.{key}",
                "schema.control_char",
                "line fields cannot contain newlines or tabs; split into separate blocks/items",
            )
        )
        return False
    return True


# ---------------------------------------------------------------------------
# core validator (pure)
# ---------------------------------------------------------------------------


def validate_deck(deck, manifest: dict, image_root: Path) -> list:
    """Validate a parsed deck against a template manifest.

    Pure: no IO except image-path existence checks under image_root.
    Returns findings in deterministic traversal order.
    """
    findings: list[Finding] = []
    if not isinstance(deck, dict):
        return [Finding("", "schema.type", f"deck root must be an object, got {type_name(deck)}")]

    check_known_keys(deck, {"template", "meta", "pages"}, "", findings)
    if "template" not in deck:
        findings.append(Finding("template", "schema.missing_field", "missing required field 'template'"))
    else:
        check_string(deck, "template", "", findings, line=True)

    if "meta" in deck:
        meta = deck["meta"]
        if not isinstance(meta, dict):
            findings.append(Finding("meta", "schema.type", f"expected object, got {type_name(meta)}"))
        else:
            check_known_keys(meta, set(META_FIELDS), "meta", findings)
            for key in META_FIELDS:
                if key in meta:
                    check_string(meta, key, "meta", findings, line=True)

    pages = deck.get("pages")
    if "pages" not in deck:
        findings.append(Finding("pages", "schema.missing_field", "missing required field 'pages'"))
    elif not isinstance(pages, list):
        findings.append(Finding("pages", "schema.type", f"expected array, got {type_name(pages)}"))
    elif not pages:
        findings.append(Finding("pages", "schema.empty_field", "deck must contain at least one page"))
    else:
        for i, page in enumerate(pages):
            _validate_page(page, i, manifest, image_root, findings)
    return findings


def _validate_page(page, page_i: int, manifest: dict, image_root: Path, findings: list) -> None:
    ppath = f"pages[{page_i}]"
    if not isinstance(page, dict):
        findings.append(Finding(ppath, "schema.type", f"expected object, got {type_name(page)}"))
        return
    check_known_keys(page, {"archetype", "blocks"}, ppath, findings)

    archetype = None
    if "archetype" not in page:
        findings.append(Finding(f"{ppath}.archetype", "schema.missing_field", "missing required field 'archetype'"))
    elif not isinstance(page["archetype"], str):
        findings.append(
            Finding(f"{ppath}.archetype", "schema.type", f"expected string, got {type_name(page['archetype'])}")
        )
    else:
        archetype = page["archetype"]

    blocks = None
    if "blocks" not in page:
        findings.append(Finding(f"{ppath}.blocks", "schema.missing_field", "missing required field 'blocks'"))
    elif not isinstance(page["blocks"], list):
        findings.append(
            Finding(f"{ppath}.blocks", "schema.type", f"expected array, got {type_name(page['blocks'])}")
        )
    else:
        blocks = page["blocks"]

    # typed block records for budget checks (only shape-valid blocks)
    records = {"title": [], "subhead": [], "text": [], "list": [], "formula": [], "image": []}
    if blocks is not None:
        for j, block in enumerate(blocks):
            record = _validate_block(block, f"{ppath}.blocks[{j}]", findings)
            if record is not None:
                records[record[0]].append((f"{ppath}.blocks[{j}]", record[1]))

    if archetype is None or blocks is None:
        return

    archetypes = manifest.get("archetypes", {})
    arch = archetypes.get(archetype)
    if not isinstance(arch, dict):
        known = ", ".join(sorted(archetypes)) or "(none)"
        findings.append(
            Finding(f"{ppath}.archetype", "schema.unknown_archetype", f"unknown archetype '{archetype}'; manifest declares: {known}")
        )
        return

    budget = arch.get("budget")
    if not isinstance(budget, dict) or any(key not in budget for key in BUDGET_KEYS):
        missing = [key for key in BUDGET_KEYS if not isinstance(budget, dict) or key not in budget]
        for key in missing:
            findings.append(
                Finding(
                    f"manifest.archetypes.{archetype}.budget.{key}",
                    "manifest.missing_key",
                    f"manifest archetype '{archetype}' budget lacks '{key}'; cannot check budgets for this page",
                )
            )
        return

    _check_multiplicity(records, ppath, findings)
    _check_budgets(records, ppath, archetype, budget, findings)
    _check_images(records, ppath, image_root, findings)


def _validate_block(block, bpath: str, findings: list):
    """Shape-check one block; return (type, block) for budget checks, else None."""
    if not isinstance(block, dict):
        findings.append(Finding(bpath, "schema.type", f"expected object, got {type_name(block)}"))
        return None
    if "type" not in block:
        findings.append(Finding(f"{bpath}.type", "schema.missing_field", "missing required field 'type'"))
        return None
    btype = block["type"]
    if not isinstance(btype, str):
        findings.append(Finding(f"{bpath}.type", "schema.type", f"expected string, got {type_name(btype)}"))
        return None
    spec = BLOCK_SPECS.get(btype)
    if spec is None:
        allowed = ", ".join(sorted(BLOCK_SPECS))
        findings.append(Finding(f"{bpath}.type", "schema.unknown_block_type", f"unknown block type '{btype}' (allowed: {allowed})"))
        return None
    check_known_keys(block, {"type", *spec["required"], *spec["optional"]}, bpath, findings)

    for field in spec["required"]:
        if field not in block:
            findings.append(Finding(f"{bpath}.{field}", "schema.missing_field", f"missing required field '{field}'"))
            return None

    if btype == "list":
        items = block["items"]
        if not isinstance(items, list):
            findings.append(Finding(f"{bpath}.items", "schema.type", f"expected array, got {type_name(items)}"))
            return None
        for k, item in enumerate(items):
            if not isinstance(item, str):
                findings.append(Finding(f"{bpath}.items[{k}]", "schema.type", f"expected string, got {type_name(item)}"))
            elif not item.strip():
                findings.append(Finding(f"{bpath}.items[{k}]", "schema.empty_field", "empty or whitespace-only string"))
            elif any(ch in item for ch in "\n\r\t"):
                findings.append(Finding(f"{bpath}.items[{k}]", "schema.control_char", "list items must be single-line"))
        if any(not isinstance(item, str) for item in items):
            return None
        return "list", block

    # string-content blocks: title / subhead / text / formula / image path
    field = spec["required"][0]
    ok = check_string(block, field, bpath, findings, line=btype != "formula")
    if btype == "image":
        if "caption" in block:
            check_string(block, "caption", bpath, findings, line=True)
    return (btype, block) if ok else None


def _check_multiplicity(records: dict, ppath: str, findings: list) -> None:
    n_titles = len(records["title"])
    if n_titles == 0:
        findings.append(Finding(ppath, "schema.missing_title", "page requires exactly one title block"))
    elif n_titles > 1:
        findings.append(Finding(ppath, "schema.duplicate_title", f"only one title block allowed per page, found {n_titles}"))
    if len(records["subhead"]) > 1:
        findings.append(Finding(ppath, "schema.duplicate_subhead", f"only one subhead block allowed per page, found {len(records['subhead'])}"))
    if len(records["list"]) > 1:
        findings.append(Finding(ppath, "schema.duplicate_list", f"only one list block allowed per page, found {len(records['list'])}"))


def _check_budgets(records: dict, ppath: str, archetype: str, budget: dict, findings: list) -> None:
    for bpath, block in records["title"]:
        w = text_width(block["text"])
        limit = budget["title_max_chars"]
        if w > limit + EPS:
            findings.append(Finding(bpath, "budget.title_chars", f"title width {fmt_width(w)} > title_max_chars {limit} ({archetype})"))

    n_text = len(records["text"])
    if n_text > budget["text_blocks_max"]:
        if budget["text_blocks_max"] == 0:
            findings.append(Finding(ppath, "budget.text_blocks_count", f"text blocks not allowed on this archetype (text_blocks_max 0, {archetype})"))
        else:
            findings.append(Finding(ppath, "budget.text_blocks_count", f"{n_text} text blocks > text_blocks_max {budget['text_blocks_max']} ({archetype})"))
    for bpath, block in records["text"]:
        w = text_width(block["text"])
        limit = budget["text_block_max_chars"]
        if w > limit + EPS:
            findings.append(Finding(bpath, "budget.text_block_chars", f"text width {fmt_width(w)} > text_block_max_chars {limit} ({archetype})"))

    for bpath, block in records["subhead"]:
        w = text_width(block["text"])
        limit = budget["subhead_max_chars"]
        if limit == 0:
            findings.append(Finding(bpath, "budget.subhead_chars", f"subhead blocks not allowed on this archetype (subhead_max_chars 0)"))
        elif w > limit + EPS:
            findings.append(Finding(bpath, "budget.subhead_chars", f"subhead width {fmt_width(w)} > subhead_max_chars {limit} ({archetype})"))

    for bpath, block in records["list"]:
        limit = budget["list_items_max"]
        if limit == 0:
            findings.append(Finding(ppath, "budget.list_items_count", f"list blocks not allowed on this archetype (list_items_max 0, {archetype})"))
            continue
        if len(block["items"]) > limit:
            findings.append(Finding(bpath, "budget.list_items_count", f"{len(block['items'])} list items > list_items_max {limit} ({archetype})"))
        item_limit = budget["list_item_max_chars"]
        for k, item in enumerate(block["items"]):
            w = text_width(item)
            if w > item_limit + EPS:
                findings.append(Finding(f"{bpath}.items[{k}]", "budget.list_item_chars", f"list item width {fmt_width(w)} > list_item_max_chars {item_limit} ({archetype})"))

    if len(records["formula"]) > budget["formulas_max"]:
        findings.append(Finding(ppath, "budget.formulas_count", f"{len(records['formula'])} formula blocks > formulas_max {budget['formulas_max']} ({archetype})"))

    n_images = len(records["image"])
    if n_images < budget["images_min"]:
        findings.append(Finding(ppath, "budget.images_count", f"{n_images} images < images_min {budget['images_min']} ({archetype})"))
    elif n_images > budget["images_max"]:
        findings.append(Finding(ppath, "budget.images_count", f"{n_images} images > images_max {budget['images_max']} ({archetype})"))

    caption_limit = budget["caption_max_chars"]
    for bpath, block in records["image"]:
        caption = block.get("caption")
        if not isinstance(caption, str):
            continue
        if caption_limit == 0:
            findings.append(Finding(f"{bpath}.caption", "budget.caption_chars", f"captions not allowed on this archetype (caption_max_chars 0)"))
        else:
            w = text_width(caption)
            if w > caption_limit + EPS:
                findings.append(Finding(f"{bpath}.caption", "budget.caption_chars", f"caption width {fmt_width(w)} > caption_max_chars {caption_limit} ({archetype})"))

    total = sum(
        text_width(block["text"]) for _, block in records["subhead"] + records["text"]
    ) + sum(
        text_width(item) for _, block in records["list"] for item in block["items"]
    )
    if total > budget["text_total_max_chars"] + EPS:
        findings.append(
            Finding(
                ppath,
                "budget.text_total_chars",
                f"text_total {fmt_width(total)} > text_total_max_chars {budget['text_total_max_chars']} ({archetype}; subhead + text blocks + list items)",
            )
        )


def _check_images(records: dict, ppath: str, image_root: Path, findings: list) -> None:
    for bpath, block in records["image"]:
        raw = block["path"]
        p = Path(raw)
        resolved = p if p.is_absolute() else image_root / p
        if not resolved.is_file():
            findings.append(Finding(f"{bpath}.path", "image.path_missing", f"image file not found: {raw} (resolved: {resolved})"))


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


class DuplicateKeyError(ValueError):
    pass


def _reject_duplicate_keys(pairs):
    keys = [key for key, _ in pairs]
    if len(keys) != len(set(keys)):
        duplicate = next(key for key in keys if keys.count(key) > 1)
        raise DuplicateKeyError(f"duplicate object key '{duplicate}'")
    return dict(pairs)


def parse_deck(raw: str):
    return json.loads(raw, object_pairs_hook=_reject_duplicate_keys)


def resolve_manifest_path(deck, templates_dir: Path, explicit: Path | None):
    """Return (manifest_path or None, error_finding or None)."""
    if explicit is not None:
        return explicit, None
    template = deck.get("template") if isinstance(deck, dict) else None
    if not isinstance(template, str) or not template.strip():
        return None, Finding(
            "",
            "template.missing",
            "deck has no usable 'template' field; cannot locate a manifest (pass --manifest)",
        )
    candidate = templates_dir / template / "manifest.json"
    if not candidate.is_file():
        known = ", ".join(sorted(d.name for d in templates_dir.iterdir() if d.is_dir())) if templates_dir.is_dir() else "(none)"
        return None, Finding(
            "template",
            "template.unknown",
            f"template '{template}': manifest not found at {candidate} (templates dir has: {known})",
        )
    return candidate, None


def build_report(findings: list, page_count: int, fmt: str) -> str:
    if fmt == "json":
        report = {
            "valid": not findings,
            "page_count": page_count,
            "findings": [{"path": f.path, "code": f.code, "message": f.message} for f in findings],
        }
        return json.dumps(report, ensure_ascii=False, indent=2)
    lines = [f"{f.path}: [{f.code}] {f.message}" for f in findings]
    if findings:
        lines.append(f"INVALID: {len(findings)} finding(s)")
    else:
        lines.append(f"OK: {page_count} page(s), 0 findings")
    return "\n".join(lines)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Validate a deck.json against a template manifest.")
    parser.add_argument("deck", nargs="?", default="-", help="path to deck.json, or - for stdin (default)")
    parser.add_argument("--templates-dir", type=Path, default=Path(__file__).resolve().parent.parent / "templates",
                        help="directory holding <template-id>/manifest.json (default: <repo>/templates)")
    parser.add_argument("--manifest", type=Path, help="explicit manifest path (overrides template resolution)")
    parser.add_argument("--format", choices=("text", "json"), default="text", help="report format (default: text)")
    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="replace")

    try:
        if args.deck == "-":
            raw = sys.stdin.buffer.read().decode("utf-8")
            deck_dir = Path.cwd()
        else:
            deck_path = Path(args.deck)
            raw = deck_path.read_text(encoding="utf-8")
            deck_dir = deck_path.parent
    except (OSError, UnicodeError) as exc:
        print(f"error: cannot read deck as UTF-8: {exc}", file=sys.stderr)
        return 2

    try:
        deck = parse_deck(raw)
    except ValueError as exc:
        print(f"error: deck is not valid JSON: {exc}", file=sys.stderr)
        return 2

    manifest_path, error_finding = resolve_manifest_path(deck, args.templates_dir, args.manifest)
    if error_finding is not None or manifest_path is None:
        findings = [error_finding] if error_finding else []
        print(build_report(findings, 0, args.format))
        return 1

    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        print(f"error: cannot load manifest {manifest_path}: {exc}", file=sys.stderr)
        return 2

    findings = validate_deck(deck, manifest, deck_dir)
    pages = deck.get("pages") if isinstance(deck, dict) else None
    page_count = len(pages) if isinstance(pages, list) else 0
    print(build_report(findings, page_count, args.format))
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
