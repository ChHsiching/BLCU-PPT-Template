"""QA gate for a scaffolded web deck (G2 machine checks, web side) + the
export chain back to pptx.

Usage:
    python scripts/qa_check_web.py <web-dir> [--screenshots DIR]
        [--export-pptx OUT.pptx] [--templates-dir DIR] [--port N]
        [--startup-timeout S] [--format text|json]

Starts the project's dev server (npm run dev on a probed-free port,
--strictPort) and drives it with headless Chromium via Python Playwright.
Checks, each a finding on failure:

  pages    every deck page is reachable through its #N deep link and the
           progress counter shows i / N
  reveal   a page's list items all reveal under ArrowRight (steps advance)
  titles   the rendered title element matches the deck's title block
  formulas every formula block is rendered by KaTeX (.formula .katex present,
           zero .katex-error elements)
  images   every image block rendered an <img> that actually loaded
  console  zero console errors / uncaught page errors across the session

Screenshots: every page, fully revealed, captured to --screenshots DIR as
page-NN.png.

--export-pptx OUT closes the loop the web README promises: the project's own
src/deck.json is rendered through renderer-pptx (image root public/, template
original from --templates-dir) and QA-checked — the exported pptx must align
with the web content page by page (count + titles, via qa_check_pptx).

Requirements: node_modules installed in <web-dir> (npm install) and Python
playwright with Chromium. Exit codes: 0 all green, 1 findings, 2 usage /
environment error.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import qa_check_pptx as qp  # noqa: E402
import render_pptx as rp  # noqa: E402
import validate_deck as vd  # noqa: E402

VIEWPORT = {"width": 1280, "height": 720}
REVEAL_STEP_BUDGET = 60  # ArrowRight presses allowed to reveal one page


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _kill_tree(proc: subprocess.Popen) -> None:
    if proc.poll() is not None:
        return
    if os.name == "nt":
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                       capture_output=True)
    else:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()


class DevServer:
    """The scaffold's vite dev server on a probed-free port."""

    def __init__(self, web_dir: Path, port: int | None, timeout_s: float):
        self.url = f"http://127.0.0.1:{port if port is not None else free_port()}"
        self._web_dir = web_dir
        self._timeout_s = timeout_s
        self._proc: subprocess.Popen | None = None
        self._log_path = web_dir / ".qa-dev-server.log"

    def __enter__(self) -> "DevServer":
        npm = self._find_npm()
        with open(self._log_path, "w", encoding="utf-8", errors="replace") as log:
            self._proc = subprocess.Popen(
                [npm, "run", "dev", "--",
                 "--port", self.url.rsplit(":", 1)[-1], "--strictPort"],
                cwd=self._web_dir, stdout=log, stderr=subprocess.STDOUT)
        deadline = time.monotonic() + self._timeout_s
        try:
            while time.monotonic() < deadline:
                if self._proc.poll() is not None:
                    break  # server died (port clash, config error): fall through
                try:
                    with urllib.request.urlopen(self.url, timeout=2) as resp:
                        if resp.status == 200:
                            return self
                except OSError:
                    pass
                time.sleep(0.3)
        except BaseException:
            _kill_tree(self._proc)  # __exit__ never runs when __enter__ raises
            raise
        tail = self._log_tail()
        _kill_tree(self._proc)
        raise RuntimeError(
            f"dev server did not come up at {self.url} within "
            f"{self._timeout_s:.0f}s (npm log tail:{tail})")

    def _find_npm(self) -> str:
        npm = shutil.which("npm")
        if npm is None:
            raise RuntimeError("npm not found on PATH")
        return npm

    def _log_tail(self) -> str:
        try:
            return "\n" + "\n".join(self._log_path.read_text(
                encoding="utf-8", errors="replace").splitlines()[-15:])
        except OSError:
            return " (log unavailable)"

    def __exit__(self, *exc) -> None:
        if self._proc is not None:
            _kill_tree(self._proc)


# ---------------------------------------------------------------------------
# Playwright checks
# ---------------------------------------------------------------------------


def _page_blocks(page: dict, btype: str) -> list[dict]:
    return [b for b in page["blocks"] if b.get("type") == btype]


def _deck_title(page: dict) -> str | None:
    return next((b["text"] for b in page["blocks"] if b.get("type") == "title"), None)


def check_web(base_url: str, deck: dict, screenshots: Path | None) -> list:
    """Drive the running presentation; return findings (vd.Finding)."""
    from playwright.sync_api import sync_playwright

    findings: list[vd.Finding] = []
    pages = deck.get("pages") if isinstance(deck, dict) else None
    pages = pages if isinstance(pages, list) else []
    n_pages = len(pages)
    console_errors: list[str] = []

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch()
        except Exception as exc:  # missing browser, display, driver issue
            raise RuntimeError(f"cannot launch headless Chromium: {exc}") from exc
        try:
            page = browser.new_page(viewport=VIEWPORT)
            page.on("console", lambda m: console_errors.append(m.text)
                    if m.type == "error" else None)
            page.on("pageerror", lambda e: console_errors.append(f"uncaught: {e}"))

            page.goto(base_url, wait_until="domcontentloaded")
            try:
                page.wait_for_selector(".stage", timeout=20000)
            except Exception:
                # the app never mounted: a schema-broken deck crashing React, a
                # broken build… a mount failure is itself the finding (plus the
                # console evidence); per-page checks would only echo it
                findings.append(vd.Finding(
                    "", "web.mount",
                    f"the app did not mount ({base_url}): no .stage element "
                    f"within 20s; see the console errors on this report",
                ))
                for text in console_errors:
                    findings.append(vd.Finding("console", "web.console_error", text))
                return findings

            total_text = _norm(page.locator(".progress").inner_text())
            m = re.fullmatch(r"(\d+) / (\d+)", total_text)
            if m is None or int(m.group(2)) != n_pages:
                findings.append(vd.Finding(
                    "", "web.page_count",
                    f"progress counter shows {total_text!r}; the deck declares "
                    f"{n_pages} page(s)",
                ))

            for i in range(1, n_pages + 1):
                expected = f"{i} / {n_pages}"
                page.evaluate("n => { location.hash = '#' + n }", i)
                try:
                    page.wait_for_function(
                        "expected => (document.querySelector('.progress')?.innerText "
                        "|| '').replace(/\\s+/g, ' ').trim() === expected",
                        arg=expected, timeout=8000)
                except Exception:
                    findings.append(vd.Finding(
                        f"page {i}", "web.page_unreachable",
                        f"deep link #{i} never showed progress {expected!r}",
                    ))
                    continue

                # reveal every list item so the screenshot is the full page.
                # After each press confirm we are still on page i: a slow
                # final re-render must never let an extra press carry the
                # DOM checks onto page i+1.
                for _ in range(REVEAL_STEP_BUDGET):
                    if page.locator(".slide .is-hidden").count() == 0:
                        break
                    page.keyboard.press("ArrowRight")
                    page.wait_for_timeout(150)
                    if _norm(page.locator(".progress").inner_text()) != expected:
                        findings.append(vd.Finding(
                            f"page {i}", "web.reveal",
                            f"reveal overshot onto the next page at {expected!r}",
                        ))
                        break
                if _norm(page.locator(".progress").inner_text()) != expected:
                    continue  # DOM checks would judge the wrong page

                if page.locator(".slide .is-hidden").count() > 0:
                    findings.append(vd.Finding(
                        f"page {i}", "web.reveal",
                        "list items did not fully reveal under ArrowRight",
                    ))

                findings.extend(_check_page_dom(page, pages[i - 1], i))

                if screenshots is not None:
                    page.screenshot(path=str(screenshots / f"page-{i:02d}.png"))
        finally:
            browser.close()

    for text in console_errors:
        findings.append(vd.Finding("console", "web.console_error", text))
    return findings


def _check_page_dom(page, deck_page: dict, i: int) -> list:
    findings: list[vd.Finding] = []

    title_el = page.locator(".slide .cover-title, .slide .agenda-label, .slide .title-bar")
    expected_title = _deck_title(deck_page)
    if title_el.count() != 1:
        findings.append(vd.Finding(
            f"page {i}", "web.title_missing",
            f"expected exactly one title element, found {title_el.count()}",
        ))
    elif _norm(title_el.inner_text()) != _norm(expected_title or ""):
        findings.append(vd.Finding(
            f"page {i}", "web.title_mismatch",
            f"rendered title {_norm(title_el.inner_text())!r} != deck title "
            f"{expected_title!r}",
        ))

    formulas = _page_blocks(deck_page, "formula")
    n_formula_boxes = page.locator(".slide .formula").count()
    if n_formula_boxes != len(formulas):
        findings.append(vd.Finding(
            f"page {i}", "web.katex_missing",
            f"{len(formulas)} formula block(s) but {n_formula_boxes} rendered "
            f".formula element(s)",
        ))
    else:
        n_katex = page.locator(".slide .formula .katex").count()
        if n_katex != len(formulas):
            findings.append(vd.Finding(
                f"page {i}", "web.katex_missing",
                f"{len(formulas)} formula block(s) but only {n_katex} rendered "
                f"by KaTeX",
            ))
    n_errors = page.locator(".slide .katex-error").count()
    if n_errors:
        findings.append(vd.Finding(
            f"page {i}", "web.katex_error",
            f"{n_errors} KaTeX parse error element(s) on the page",
        ))

    images = _page_blocks(deck_page, "image")
    n_imgs = page.locator(".slide .image-slot img").count()
    if n_imgs != len(images):
        findings.append(vd.Finding(
            f"page {i}", "web.image_missing",
            f"{len(images)} image block(s) but {n_imgs} <img> rendered",
        ))
    # settle in-flight loads first: a cold first read (AV scan, loaded box)
    # must not read as broken. complete flips on load OR error; the filter
    # below then tells them apart via naturalWidth.
    try:
        page.wait_for_function(
            "() => [...document.querySelectorAll('.slide img')]"
            ".every(img => img.complete)", timeout=8000)
    except Exception:
        pass  # a hung request falls through to the broken check below
    broken = page.evaluate(
        "() => [...document.querySelectorAll('.slide .image-slot img')]"
        ".filter(img => !(img.complete && img.naturalWidth > 0))"
        ".map(img => img.getAttribute('src'))")
    for src in broken:
        findings.append(vd.Finding(
            f"page {i}", "web.image_broken",
            f"image did not load: {src}",
        ))
    return findings


# ---------------------------------------------------------------------------
# export chain: the web deck rendered through renderer-pptx, then pptx-checked
# ---------------------------------------------------------------------------


def check_export_chain(web_dir: Path, deck: dict, manifest: dict,
                       templates_dir: Path, out_pptx: Path | None) -> list:
    """Render src/deck.json via renderer-pptx and QA the result."""
    template_pptx = (templates_dir / deck.get("template", "") /
                     manifest.get("source_pptx", ""))
    if not template_pptx.is_file():
        raise RuntimeError(
            f"template original not found at {template_pptx}; pass --templates-dir")

    result = rp.render_deck(deck, manifest, template_pptx,
                            image_root=web_dir / "public")
    if out_pptx is not None:
        if os.path.normcase(str(Path(out_pptx).resolve())) == \
                os.path.normcase(str(template_pptx.resolve())):
            raise RuntimeError(f"refusing to overwrite the template original: {template_pptx}")
        result.presentation.save(out_pptx)

    findings = qp.check_pptx(result.presentation, deck, manifest, web_dir / "public")
    # re-scope the finding paths so the report names the export chain
    rescoped = [vd.Finding(f"export {f.path}".strip(), f.code, f.message)
                for f in findings]
    return rescoped


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_report(findings: list, page_count: int, fmt: str, extra: dict | None = None) -> str:
    return qp.build_report(findings, page_count, fmt, extra)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="QA-check a scaffolded web presentation with Playwright.")
    parser.add_argument("web_dir", help="scaffolded web project directory")
    parser.add_argument("--screenshots", type=Path, metavar="DIR",
                        help="capture every page as page-NN.png into DIR")
    parser.add_argument("--export-pptx", type=Path, metavar="OUT",
                        help="also export src/deck.json via renderer-pptx and QA it")
    parser.add_argument("--templates-dir", type=Path,
                        default=Path(__file__).resolve().parent.parent / "templates",
                        help="directory holding <template-id>/ (default: <repo>/templates)")
    parser.add_argument("--port", type=int,
                        help="dev server port (default: a probed-free port)")
    parser.add_argument("--startup-timeout", type=float, default=60.0,
                        help="seconds to wait for the dev server (default: 60)")
    parser.add_argument("--format", choices=("text", "json"), default="text",
                        help="report format (default: text)")
    args = parser.parse_args(argv)
    # gate reports are consumed by programs (and asserted by tests) — keep
    # them UTF-8 regardless of the host locale / pipe encoding
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(encoding="utf-8", errors="replace")

    # .resolve() expands Windows 8.3 short names (ADMINI~1) to the long form:
    # vite's transform pipeline fails to resolve src modules when the dev
    # server runs with a short-form cwd ("Pre-transform error: Failed to load
    # url /src/main.jsx"), so the app never mounts
    web_dir = Path(args.web_dir).resolve()
    deck_path = web_dir / "src" / "deck.json"
    manifest_path = web_dir / "src" / "manifest.json"
    for required in (web_dir / "package.json", deck_path, manifest_path):
        if not required.is_file():
            print(f"error: {required} not found; is this a scaffolded project?",
                  file=sys.stderr)
            return 2
    if not (web_dir / "node_modules").is_dir():
        print(f"error: {web_dir / 'node_modules'} missing; "
              f"run `npm install` in the project first", file=sys.stderr)
        return 2

    try:
        deck = vd.parse_deck(deck_path.read_text(encoding="utf-8"))
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as exc:
        print(f"error: cannot read the project's deck/manifest: {exc}", file=sys.stderr)
        return 2

    # a schema-broken deck crashes React at mount: report it as findings
    # before a browser (or the playwright package) is ever involved
    findings: list[vd.Finding] = vd.validate_deck(deck, manifest, web_dir / "public")
    if findings:
        print(build_report(findings, 0, args.format))
        return 1

    try:
        import playwright.sync_api  # noqa: F401
    except ImportError:
        print("error: python playwright is not installed (pip install playwright "
              "&& playwright install chromium)", file=sys.stderr)
        return 2

    if args.screenshots is not None:
        args.screenshots.mkdir(parents=True, exist_ok=True)

    try:
        with DevServer(web_dir, args.port, args.startup_timeout) as server:
            findings.extend(check_web(server.url, deck, args.screenshots))
    except RuntimeError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    extra = {}
    if args.export_pptx is not None:
        try:
            findings.extend(check_export_chain(
                web_dir, deck, manifest, args.templates_dir, args.export_pptx))
        except (RuntimeError, rp.RenderError) as exc:
            print(f"error: export chain failed: {exc}", file=sys.stderr)
            return 2
        except OSError as exc:
            print(f"error: cannot write output {args.export_pptx}: {exc}", file=sys.stderr)
            return 2
        extra["exported"] = str(args.export_pptx)
    if args.screenshots is not None:
        extra["screenshots"] = f"{len(deck['pages'])} page(s) -> {args.screenshots}"

    print(build_report(findings, len(deck["pages"]), args.format, extra))
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
