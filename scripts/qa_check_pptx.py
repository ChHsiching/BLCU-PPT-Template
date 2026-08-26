"""QA gate for a rendered .pptx deck (G2 machine checks, pptx side).

Usage:
    python scripts/qa_check_pptx.py <out.pptx> --deck <deck.json>
        [--templates-dir DIR] [--manifest PATH] [--com-screenshots DIR]
        [--format text|json]

The pptx is checked against the deck it claims to render (deck.json stays the
single source of truth). Checks, each a finding on failure:

  count     slide count == deck page count
  titles    per-page title alignment (title placeholder; agenda pages carry
            none, so their title block is matched against the rebuilt
            AgendaLabel box)
  residue   placeholder-text grep over every text run of every slide
            (TODO/FIXME/TBD/xxx/lorem/[insert/placeholder/待补充/…)
  budget    the deck itself revalidated against the manifest
            (scripts/validate_deck.py pure function) — catches decks edited
            after rendering or rendered through the API without validation
  structure python-pptx level: every shape inside the slide canvas;
            text-bearing shapes (pictures, the page-number placeholder and
            in-slot caption strips excepted) above safe_canvas
            .content_bottom_y; zero speaker notes (演讲稿 is a separate
            deliverable); pages referencing images carry at least that many
            embedded pictures (formula fallbacks may add more)

--com-screenshots DIR adds the optional PowerPoint COM spot check: every slide
exported as a PNG into DIR for visual overflow inspection. A missing COM
environment (not Windows, no PowerPoint, no pywin32) is never itself a
finding — a note goes to stderr and the programmatic checks above stand alone.

Exit codes: 0 all green, 1 findings, 2 usage / IO error.
The core is the pure check_pptx(prs, deck, manifest, image_root) -> list
of validate_deck.Finding; scripts/qa_check_web.py reuses it for the export
chain.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

sys.path.insert(0, str(Path(__file__).resolve().parent))
import validate_deck as vd  # noqa: E402

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

PICTURE_SHAPE_TYPE = 13  # MSO_SHAPE_TYPE.PICTURE, avoided as a hard import

# placeholder-residue patterns: (name, matcher). Regexes are for scripts and
# latin filler with word boundaries; CJK fillers match as substrings (CJK has
# no word boundaries to lean on). The list is deliberately conservative:
# bare "占位" / "待定" are excluded because they live in legitimate academic
# terms (可学习占位符 / 待定系数法) — only unambiguous filler compounds match.
_RESIDUE_REGEXES = [
    ("todo", re.compile(r"\btodo\b", re.IGNORECASE)),
    ("fixme", re.compile(r"\bfixme\b", re.IGNORECASE)),
    ("tbd", re.compile(r"\btbd\b", re.IGNORECASE)),
    ("xxx", re.compile(r"x{3,}", re.IGNORECASE)),
    ("lorem", re.compile(r"lorem", re.IGNORECASE)),
    ("insert-tag", re.compile(r"[\[<\{]insert\b", re.IGNORECASE)),
    ("placeholder", re.compile(r"\bplaceholder\b", re.IGNORECASE)),
    ("sample-text", re.compile(r"\bsample text\b", re.IGNORECASE)),
]
_RESIDUE_SUBSTRINGS = ("待补充", "待填写", "此处填写", "示例文本")

SNIPPET_RADIUS = 10


def find_residue(text: str) -> list[tuple[str, str]]:
    """Return [(pattern_name, snippet)] for every filler marker in text."""
    hits = []
    for name, rx in _RESIDUE_REGEXES:
        for m in rx.finditer(text):
            lo, hi = max(0, m.start() - SNIPPET_RADIUS), min(len(text), m.end() + SNIPPET_RADIUS)
            hits.append((name, text[lo:hi]))
    for needle in _RESIDUE_SUBSTRINGS:
        start = text.find(needle)
        if start != -1:
            lo, hi = max(0, start - SNIPPET_RADIUS), min(len(text), start + len(needle) + SNIPPET_RADIUS)
            hits.append((needle, text[lo:hi]))
    return hits


# ---------------------------------------------------------------------------
# slide introspection (mirrors the renderer's own serialization)
# ---------------------------------------------------------------------------


def ph_type(sp) -> str | None:
    """Placeholder type of a shape, None when not a placeholder."""
    ph = sp._element.find(f".//{{{NS_P}}}ph")
    if ph is None:
        return None
    return ph.get("type") or "body"


def shape_text(sp) -> str:
    """All visible text of a shape: a:t and m:t alike, concatenated."""
    parts = []
    for t in sp._element.iter(f"{{{NS_A}}}t"):
        parts.append(t.text or "")
    for t in sp._element.iter(f"{{{NS_M}}}t"):
        parts.append(t.text or "")
    return "".join(parts)


def slide_title_text(slide) -> str | None:
    """Title text of a rendered slide.

    Content pages carry a title placeholder; agenda pages have none — their
    title block lands in the renderer's rebuilt AgendaLabel box.
    """
    for sp in slide.shapes:
        if ph_type(sp) == "title":
            return sp.text_frame.text
    for sp in slide.shapes:
        if sp.name == "AgendaLabel":
            return sp.text_frame.text
    return None


def _deck_title(page) -> str | None:
    """Title text of a deck page, tolerating schema-broken pages (None).

    The tool's job is to judge exactly the decks that may have been edited
    after rendering, so traversal never raises on missing fields — the
    validator's schema findings report those, this just skips them.
    """
    if not isinstance(page, dict):
        return None
    blocks = page.get("blocks")
    if not isinstance(blocks, list):
        return None
    for block in blocks:
        if isinstance(block, dict) and block.get("type") == "title" \
                and isinstance(block.get("text"), str):
            return block["text"]
    return None


def _deck_image_count(page) -> int:
    if not isinstance(page, dict):
        return 0
    blocks = page.get("blocks")
    if not isinstance(blocks, list):
        return 0
    return sum(1 for b in blocks if isinstance(b, dict) and b.get("type") == "image")


# ---------------------------------------------------------------------------
# core checks (pure; reused by qa_check_web's export chain)
# ---------------------------------------------------------------------------


def check_pptx(prs, deck: dict, manifest: dict, image_root: Path) -> list:
    """Run every pptx-side check; return findings in slide order."""
    findings: list[vd.Finding] = []

    # budget: the deck is the single source of truth — revalidate it
    findings.extend(vd.validate_deck(deck, manifest, image_root))

    # count (deck pages may be anything; _deck_* helpers stay total)
    slides = list(prs.slides)
    pages = deck.get("pages") if isinstance(deck, dict) else None
    pages = pages if isinstance(pages, list) else []

    # count
    if len(slides) != len(pages):
        findings.append(vd.Finding(
            "", "pptx.page_count",
            f"pptx has {len(slides)} slide(s) but the deck declares {len(pages)} page(s)",
        ))

    # per-slide checks over the aligned prefix
    for i, slide in enumerate(slides):
        spath = f"slides[{i + 1}]"
        page = pages[i] if i < len(pages) else None

        # titles (skip pages whose deck side has no usable title block —
        # the validator's schema.missing_title finding already covers them)
        if page is not None:
            expected = _deck_title(page)
            actual = slide_title_text(slide)
            if expected is not None:
                if actual is None:
                    findings.append(vd.Finding(
                        spath, "pptx.title_missing",
                        f"no title found on slide {i + 1} (expected: {expected!r})",
                    ))
                elif expected != actual:
                    findings.append(vd.Finding(
                        spath, "pptx.title_mismatch",
                        f"title on slide {i + 1} is {actual!r} but the deck declares {expected!r}",
                    ))

        # residue + structure, shape by shape
        n_pictures = 0
        for sp in slide.shapes:
            is_picture = sp.shape_type == PICTURE_SHAPE_TYPE
            if is_picture:
                n_pictures += 1
            text = shape_text(sp)
            if text:
                for name, snippet in find_residue(text):
                    findings.append(vd.Finding(
                        f"{spath}.{sp.name}", "pptx.residue",
                        f"placeholder residue [{name}] on slide {i + 1}: …{snippet}…",
                    ))
            findings.extend(_check_geometry(sp, spath, manifest, is_picture))

        if page is not None:
            n_images = _deck_image_count(page)
            if n_pictures < n_images:
                findings.append(vd.Finding(
                    spath, "pptx.images_missing",
                    f"slide {i + 1} carries {n_pictures} picture(s) but the deck "
                    f"references {n_images} image block(s)",
                ))

        if slide.has_notes_slide:
            findings.append(vd.Finding(
                spath, "pptx.notes_slide",
                f"slide {i + 1} has speaker notes; 演讲稿 is a separate deliverable "
                f"and never lives in the pptx",
            ))

    return findings


def _check_geometry(sp, spath: str, manifest: dict, is_picture: bool) -> list:
    """Canvas bounds for every shape; safe content bottom for text shapes.

    Exceptions mirror manifest.safe_canvas: pictures only answer to the canvas,
    the page-number placeholder sits in its own reserved band, and caption
    strips overlay the internal bottom edge of their owning image slot.
    """
    findings: list[vd.Finding] = []
    if sp.left is None or sp.top is None or sp.width is None or sp.height is None:
        return findings
    tol = 0.02  # absorbs int-EMU rounding, nothing more
    canvas_w, canvas_h = manifest["slide_size"]["w"], manifest["slide_size"]["h"]
    left, top = Emu(sp.left).inches, Emu(sp.top).inches
    right = Emu(sp.left + sp.width).inches
    bottom = Emu(sp.top + sp.height).inches
    if left < -tol or top < -tol or right > canvas_w + tol or bottom > canvas_h + tol:
        findings.append(vd.Finding(
            f"{spath}.{sp.name}", "pptx.canvas_overflow",
            f"'{sp.name}' spans ({left:.2f}, {top:.2f})-({right:.2f}, {bottom:.2f}); "
            f"outside the {canvas_w}x{canvas_h} canvas",
        ))
        return findings

    if is_picture or ph_type(sp) == "sldNum" or sp.name.startswith("Caption"):
        return findings
    content_bottom = manifest.get("safe_canvas", {}).get("content_bottom_y")
    if content_bottom is not None and bottom > content_bottom + tol:
        findings.append(vd.Finding(
            f"{spath}.{sp.name}", "pptx.content_bottom",
            f"'{sp.name}' bottom edge {bottom:.2f} crosses safe_canvas "
            f"content_bottom_y {content_bottom}",
        ))
    return findings


# ---------------------------------------------------------------------------
# optional COM screenshot spot check (degrades to a stderr note)
# ---------------------------------------------------------------------------


def export_com_screenshots(pptx_path: Path, out_dir: Path) -> tuple[int, str | None]:
    """Export every slide as PNG via PowerPoint COM; (count, error)."""
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return 0, "pywin32 not installed; COM spot check unavailable"
    try:
        import pythoncom

        pythoncom.CoInitialize()
        import win32com.client

        # DispatchEx: a dedicated instance, so Quit never closes the user's own
        app = win32com.client.DispatchEx("PowerPoint.Application")
    except Exception as exc:
        return 0, f"cannot start PowerPoint COM: {exc}"
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        pres = app.Presentations.Open(
            str(pptx_path.resolve()), ReadOnly=True, Untitled=False, WithWindow=False)
        try:
            # slide count from the presentation itself: reusing a non-empty
            # DIR must not inflate the reported number with stale PNGs
            n_slides = pres.Slides.Count
            pres.Export(str(out_dir.resolve()), "PNG", 1280, 720)
        finally:
            pres.Close()
        return n_slides, None
    except Exception as exc:
        return 0, f"COM export failed: {exc}"
    finally:
        try:
            app.Quit()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_report(findings: list, slide_count: int, fmt: str, extra: dict | None = None) -> str:
    if fmt == "json":
        report = {
            "valid": not findings,
            "slide_count": slide_count,
            "findings": [{"path": f.path, "code": f.code, "message": f.message} for f in findings],
        }
        if extra:
            report.update(extra)
        return json.dumps(report, ensure_ascii=False, indent=2)
    lines = [f"{f.path}: [{f.code}] {f.message}" for f in findings]
    if findings:
        lines.append(f"INVALID: {len(findings)} finding(s)")
    else:
        lines.append(f"OK: {slide_count} slide(s), 0 findings")
    for key, value in (extra or {}).items():
        lines.append(f"{key}: {value}")
    return "\n".join(lines)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="QA-check a rendered .pptx against the deck.json it renders.")
    parser.add_argument("pptx", help="path to the rendered .pptx")
    parser.add_argument("--deck", required=True, help="path to the deck.json it renders")
    parser.add_argument("--templates-dir", type=Path,
                        default=Path(__file__).resolve().parent.parent / "templates",
                        help="directory holding <template-id>/ (default: <repo>/templates)")
    parser.add_argument("--manifest", type=Path,
                        help="explicit manifest path (overrides template resolution)")
    parser.add_argument("--com-screenshots", type=Path, metavar="DIR",
                        help="also export slide PNGs via PowerPoint COM into DIR")
    parser.add_argument("--format", choices=("text", "json"), default="text",
                        help="report format (default: text)")
    args = parser.parse_args(argv)
    # gate reports are consumed by programs (and asserted by tests) — keep
    # them UTF-8 regardless of the host locale / pipe encoding
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(encoding="utf-8", errors="replace")

    deck_path = Path(args.deck)
    try:
        deck = vd.parse_deck(deck_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as exc:
        print(f"error: cannot read deck: {exc}", file=sys.stderr)
        return 2

    manifest_path, error_finding = vd.resolve_manifest_path(
        deck, args.templates_dir, args.manifest)
    if error_finding is not None or manifest_path is None:
        print(build_report([error_finding] if error_finding else [], 0, args.format))
        return 1
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        print(f"error: cannot load manifest {manifest_path}: {exc}", file=sys.stderr)
        return 2

    pptx_path = Path(args.pptx)
    if not pptx_path.is_file():
        print(f"error: pptx not found: {pptx_path}", file=sys.stderr)
        return 2
    try:
        prs = Presentation(pptx_path)
    except Exception as exc:
        finding = vd.Finding("", "pptx.unreadable", f"cannot open with python-pptx: {exc}")
        print(build_report([finding], 0, args.format))
        return 1

    findings = check_pptx(prs, deck, manifest, deck_path.parent)

    extra = {}
    if args.com_screenshots is not None:
        count, error = export_com_screenshots(pptx_path, args.com_screenshots)
        if error is not None:
            print(f"note: {error}; programmatic checks stand alone", file=sys.stderr)
        else:
            extra["com_screenshots"] = f"{count} slide(s) -> {args.com_screenshots}"

    print(build_report(findings, len(prs.slides), args.format, extra))
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
