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
  styles   the typography.tokens role system at computed-style level: every
           [data-role] element's font-family/weight/color/size against its
           tokens role, rhythm elements' line pitch + 12pt paragraph space
           (single-line roles at the single pitch), the caption scrim, the
           emphasis run count/style (paired ** markers over body-flow blocks)
           and the hairline on white-backgrounded images
  fonts    Noto Sans SC actually served (document.fonts reports loaded faces
           for both weights after fonts.ready)
  brand    the master brand layer per manifest.brand_layer: band/logo counts,
           band geometry + color against the measured regions, every logo
           asset loaded, the right page number (none on cover/closing), plus
           a rendered-pixel spot-check of band color/position (element
           screenshot + PIL; without PIL the pixel check degrades to a note)
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


def check_web(base_url: str, deck: dict, screenshots: Path | None,
              manifest: dict | None = None, notes: list | None = None) -> list:
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
                findings.extend(_check_page_styles(page, pages[i - 1], manifest, i))

                # brand-layer checks need the entrance animation settled
                # (it translates the whole slide for 0.35s after a page change)
                page.wait_for_timeout(450)
                findings.extend(_check_page_brand(page, pages[i - 1], manifest, i,
                                                  notes or []))
                findings.extend(_check_brand_pixels(page, pages[i - 1], manifest,
                                                    i, notes or []))

                if screenshots is not None:
                    page.screenshot(path=str(screenshots / f"page-{i:02d}.png"))

            findings.extend(_check_fonts(page))
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
# style layer (#19): typography.tokens at computed-style level — the web
# counterpart of the pptx run-level style facts test_render_pptx asserts
# ---------------------------------------------------------------------------

STYLE_TOL_PX = 0.75  # computed px carry browser rounding (26.667 -> 26.67)
PT_TO_PX = 96 / 72  # the stage's 96dpi inch, same as src/lib/layout.js


def _css_px(value: str) -> float | None:
    m = re.match(r"([\d.]+)px", value or "")
    return float(m.group(1)) if m else None


def _css_rgb(value: str) -> tuple | None:
    got = [int(c) for c in re.findall(r"\d+", value or "")]
    return tuple(got[:3]) if len(got) >= 3 else None


def _hex_rgb(color: str) -> tuple:
    c = color.lstrip("#")
    return tuple(int(c[k:k + 2], 16) for k in (0, 2, 4))


def _expected_title_role(tokens: dict, deck_page: dict, text: str | None) -> str | None:
    """Independently resolve the title element's role: role_bindings, then the
    long-title downgrade over title_long.over_chars (CJK width)."""
    archetype = deck_page.get("archetype", "")
    bindings = tokens.get("role_bindings", {}).get(archetype, {})
    name = bindings.get("label") if archetype == "agenda" else bindings.get("title")
    long = tokens.get("roles", {}).get("title_long", {})
    if name == "title" and text is not None and \
            vd.text_width(text) > long.get("over_chars", float("inf")):
        name = "title_long"
    return name


def _body_flow_texts(deck_page: dict) -> list[str]:
    """Texts rendering as body flow — exactly the blocks the pptx renderer
    passes emphasis=True. Cover/closing subtitles, subheads and captions keep
    ** markers literal, so they contribute no expected emphasis runs."""
    archetype = deck_page.get("archetype", "")
    out: list[str] = []
    if archetype in ("text-formula", "text-image", "chart-focus"):
        out += [b.get("text", "") for b in deck_page["blocks"] if b.get("type") == "text"]
    if archetype in ("agenda", "text-image"):
        out += [item for b in deck_page["blocks"] if b.get("type") == "list"
                for item in b.get("items", [])]
    return out


def _check_page_styles(page, deck_page: dict, manifest: dict | None, i: int) -> list:
    findings: list[vd.Finding] = []
    if manifest is None:
        return findings
    tokens = manifest.get("typography", {}).get("tokens")
    if not tokens:
        findings.append(vd.Finding(
            f"page {i}", "web.style_missing",
            "manifest has no typography.tokens; the style layer cannot be "
            "checked or rendered",
        ))
        return findings
    roles = tokens.get("roles", {})
    weights = tokens.get("weights", {})
    spacing = tokens.get("spacing", {})
    weights_by_name = {v: k for k, v in weights.items()}

    state = page.evaluate(
        """() => {
          const slide = document.querySelector('.slide')
          const px = v => parseFloat(v) || 0
          const textTop = el => {
            // element-rect based: Range rects over anonymous flex text are
            // unreliable in Chrome; every measured box carries its text in a
            // real child element (span for single-line roles, p/div for flow)
            const first = el.firstElementChild
            return first
              ? first.getBoundingClientRect().top - el.getBoundingClientRect().top
              : null
          }
          const roleEls = [...slide.querySelectorAll('[data-role]')].map(el => {
            const cs = getComputedStyle(el)
            return {
              role: el.dataset.role, rhythm: el.dataset.rhythm === '1',
              family: cs.fontFamily, size: px(cs.fontSize),
              weight: cs.fontWeight, color: cs.color, lineHeight: cs.lineHeight,
              paraTop: el.dataset.rhythm === '1' && el.firstElementChild
                ? getComputedStyle(el.firstElementChild).marginTop : null,
              captionBg: el.dataset.role === 'caption' ? cs.backgroundColor : null,
              // vertical placement of the first text line inside the box —
              // guards the top-anchor model (v-inset, +12pt for body flow)
              textTop: ['subhead', 'caption'].includes(el.dataset.role)
                         || el.dataset.rhythm === '1' ? textTop(el) : null,
            }
          })
          const emph = [...slide.querySelectorAll('.emph')].map(el => ({
            color: getComputedStyle(el).color,
            weight: getComputedStyle(el).fontWeight,
            text: el.textContent,
          }))
          // the white-background hairline detector duplicated from
          // src/lib/hairline.js: the gate must verify the renderer applied
          // the outline, not take the renderer's word for it
          const hairlineImgs = [...slide.querySelectorAll('.image-fit img')].map(img => {
            let white = false
            try {
              const c = document.createElement('canvas')
              c.width = img.naturalWidth; c.height = img.naturalHeight
              const ctx = c.getContext('2d', { willReadFrequently: true })
              ctx.drawImage(img, 0, 0)
              const { width: w, height: h } = c
              const pts = [[0,0],[w-1,0],[0,h-1],[w-1,h-1],
                           [w>>1,0],[w>>1,h-1],[0,h>>1],[w-1,h>>1]]
              white = pts.every(([x,y]) => {
                const d = ctx.getImageData(x,y,1,1).data
                return Math.min(d[0], d[1], d[2]) >= 245
              })
            } catch { white = false }
            const cs = getComputedStyle(img)
            return { white, loaded: img.complete && img.naturalWidth > 0,
                     src: img.getAttribute('src'),
                     width: cs.outlineWidth, style: cs.outlineStyle,
                     color: cs.outlineColor }
          })
          const titleEl = slide.querySelector('.cover-title, .agenda-label, .title-bar')
          return { roleEls, emph, hairlineImgs,
                   titleRole: titleEl ? titleEl.dataset.role : null }
        }""")

    def style_finding(what, got, want):
        findings.append(vd.Finding(
            f"page {i}", "web.style",
            f"{what}: rendered {got!r}, tokens expect {want!r}",
        ))

    # the title element's role is re-resolved independently (bindings + the
    # long-title downgrade), so a mislabeled data-role cannot hide a defect
    expected_title = _expected_title_role(tokens, deck_page, _deck_title(deck_page))
    if expected_title is None:
        findings.append(vd.Finding(
            f"page {i}", "web.style",
            f"archetype {deck_page.get('archetype')!r} has no title role binding",
        ))
    elif state["titleRole"] != expected_title:
        style_finding("title data-role", state["titleRole"], expected_title)

    for el in state["roleEls"]:
        role = roles.get(el["role"])
        if role is None:
            findings.append(vd.Finding(
                f"page {i}", "web.style",
                f"element carries data-role={el['role']!r}, not a tokens role",
            ))
            continue
        where = f"[{el['role']}]"
        family_first = el["family"].split(",")[0].strip().strip('"')
        if family_first != tokens.get("face"):
            style_finding(f"{where} font-family", el["family"], tokens.get("face"))
        if abs(el["size"] - PT_TO_PX * role["size_pt"]) > STYLE_TOL_PX:
            style_finding(f"{where} font-size", f"{el['size']}px", f"{role['size_pt']}pt")
        if weights_by_name.get(int(el["weight"])) != role["weight"]:
            style_finding(f"{where} font-weight", el["weight"], role["weight"])
        if _css_rgb(el["color"]) != _hex_rgb(role["color"]):
            style_finding(f"{where} color", el["color"], role["color"])
        pitch = spacing.get("line_pitch_em" if el["rhythm"] else "single_pitch_em")
        line_height = _css_px(el["lineHeight"])
        if pitch and (line_height is None or
                      abs(line_height - el["size"] * pitch) > 1.0):
            style_finding(f"{where} line-height", el["lineHeight"],
                          f"{el['size'] * pitch:.1f}px ({pitch}em)")
        if el["rhythm"]:
            before = _css_px(el["paraTop"])
            if before is None or abs(before - PT_TO_PX * spacing["space_before_pt"]) > 0.5:
                style_finding(f"{where} paragraph space-before", el["paraTop"],
                              f"{spacing['space_before_pt']}pt")
        # vertical placement: the first text line sits at the v-inset
        # (top-anchored boxes), plus the 12pt paragraph space in body flow.
        # getBoundingClientRect scales with the stage transform, so this is
        # only valid unscaled because VIEWPORT equals the 1280x720 stage
        # (scale === 1); a different VIEWPORT needs a stage-size division
        # here like the brand checks' stageH normalization.
        if el["textTop"] is not None:
            inset_px = 1280 / manifest["slide_size"]["w"] * spacing["textbox_inset_v_in"]
            want = inset_px + (PT_TO_PX * spacing["space_before_pt"] if el["rhythm"] else 0)
            if abs(el["textTop"] - want) > 1.5:
                style_finding(f"{where} first-line offset", f"{el['textTop']:.1f}px",
                              f"{want:.1f}px")
        if el["role"] == "caption":
            scrim = role.get("scrim", {})
            nums = [float(c) for c in re.findall(r"[\d.]+", el["captionBg"] or "")]
            want_rgb = _hex_rgb(scrim.get("color", "#000000"))
            want_alpha = scrim.get("alpha_pct", 0) / 100
            if len(nums) < 4 or tuple(int(n) for n in nums[:3]) != want_rgb or \
                    abs(nums[3] - want_alpha) > 0.01:
                style_finding("caption scrim", el["captionBg"],
                              f"rgba{want_rgb + (want_alpha,)}")

    # emphasis: paired markers over body-flow blocks -> green bold runs
    emph_cfg = tokens.get("emphasis", {})
    marker = re.escape(emph_cfg.get("marker", "**"))
    expected_runs = sum(
        len(re.findall(rf"{marker}(.+?){marker}", t, re.S))
        for t in _body_flow_texts(deck_page))
    if len(state["emph"]) != expected_runs:
        findings.append(vd.Finding(
            f"page {i}", "web.style",
            f"{expected_runs} emphasized keyword(s) expected in body flow, "
            f"{len(state['emph'])} .emph element(s) rendered",
        ))
    for run in state["emph"]:
        if _css_rgb(run["color"]) != _hex_rgb(emph_cfg.get("color", "#548235")) or \
                weights_by_name.get(int(run["weight"])) != emph_cfg.get("weight"):
            style_finding(f"emphasis run {run['text']!r}",
                          f"{run['color']} {run['weight']}",
                          f"{emph_cfg.get('color')} {emph_cfg.get('weight')}")

    # hairline: white-backgrounded images carry the token outline
    hair = tokens.get("image", {}).get("hairline")
    if hair:
        want_w, want_c = PT_TO_PX * hair["width_pt"], _hex_rgb(hair["color"])
        for img in state["hairlineImgs"]:
            if not img["loaded"] or not img["white"]:
                continue
            if img["style"] != "solid" or \
                    abs(_css_px(img["width"]) - want_w) > STYLE_TOL_PX or \
                    _css_rgb(img["color"]) != want_c:
                style_finding(f"hairline on {img['src']}",
                              f"{img['style']} {img['width']} {img['color']}",
                              f"solid {want_w}px rgb{want_c}")
    return findings


def _check_fonts(page) -> list:
    """The bundled faces actually served: after fonts.ready, both Noto Sans SC
    weights report loaded (woff2 subsets load on demand via unicode-range, so
    'loaded' means the deck's glyphs pulled them)."""
    try:
        loaded = page.evaluate(
            """async () => {
              await document.fonts.ready
              return [...document.fonts].filter(f => f.status === 'loaded')
                .map(f => `${f.family.replace(/"/g, '')} ${f.weight}`)
            }""")
    except Exception as exc:
        return [vd.Finding("fonts", "web.fonts", f"document.fonts probe failed: {exc}")]
    findings = []
    for weight in (400, 700):
        if f"Noto Sans SC {weight}" not in loaded:
            findings.append(vd.Finding(
                "fonts", "web.fonts",
                f"Noto Sans SC {weight} did not load (loaded faces: "
                f"{sorted(set(loaded))})",
            ))
    return findings


# ---------------------------------------------------------------------------
# brand layer (issue #17): master bands/logos/page numbers, manifest-measured
# ---------------------------------------------------------------------------

GEO_TOL_IN = 0.05  # DOM-rect tolerance in slide inches (rounding + subpixel)


def _brand_expect(manifest: dict, archetype: str):
    """(expected logo count, [(region, name)] incl. the band first) or None
    when the manifest carries no brand_layer (a finding by itself)."""
    bl = manifest.get("brand_layer") if isinstance(manifest, dict) else None
    if not isinstance(bl, dict):
        return None
    if archetype in ("cover", "closing"):
        cover = bl["cover"]
        elements = [(cover["mid_band"], "mid_band")]
        elements += [(el, f"logos[{j}]") for j, el in enumerate(cover["logos"])]
        elements.append((cover["corner_logo_bar"], "corner_logo_bar"))
        return len(cover["logos"]) + 1, elements
    content = bl["content"]
    return 1, [(content["top_band"], "top_band"), (content["corner_logo"], "corner_logo")]


def _check_page_brand(page, deck_page: dict, manifest: dict | None, i: int,
                      notes: list) -> list:
    findings: list[vd.Finding] = []
    if manifest is None:
        return findings  # pre-brand manifest: pixel check below reports it once
    expect = _brand_expect(manifest, deck_page.get("archetype", ""))
    if expect is None:
        findings.append(vd.Finding(
            f"page {i}", "web.brand_missing",
            "manifest has no brand_layer section; the web brand layer "
            "(bands/logos/page numbers) cannot be checked or rendered",
        ))
        return findings
    n_logos, elements = expect

    state = page.evaluate(
        """() => {
          const slide = document.querySelector('.slide')
          const sr = document.querySelector('.stage').getBoundingClientRect()
          const rect = el => {
            const r = el.getBoundingClientRect()
            return {x: r.left - sr.left, y: r.top - sr.top, w: r.width, h: r.height}
          }
          return {
            stageH: sr.height,
            band: slide.querySelector('.brand-band')
              ? {...rect(slide.querySelector('.brand-band')),
                 color: getComputedStyle(slide.querySelector('.brand-band')).backgroundColor}
              : null,
            logos: [...slide.querySelectorAll('.brand-logo')].map(im => ({
              ...rect(im), loaded: im.complete && im.naturalWidth > 0,
              src: im.getAttribute('src')})),
            pageNumber: slide.querySelector('.page-number')?.innerText.trim() ?? null,
          }
        }""")

    if state["band"] is None:
        findings.append(vd.Finding(
            f"page {i}", "web.brand_missing",
            "no .brand-band rendered on this page",
        ))
        return findings
    if len(state["logos"]) != n_logos:
        findings.append(vd.Finding(
            f"page {i}", "web.brand_geometry",
            f"expected {n_logos} .brand-logo element(s) per manifest.brand_layer, "
            f"found {len(state['logos'])}",
        ))

    # geometry: rects are page px on the scaled stage — compare as slide inches
    stage_h = state["stageH"]
    to_in = lambda v: v / stage_h * 7.5

    def close(actual, region, name, what):
        for key in ("x", "y", "w", "h"):
            got = to_in(actual[key])
            want = region[key]
            if abs(got - want) > GEO_TOL_IN:
                findings.append(vd.Finding(
                    f"page {i}", "web.brand_geometry",
                    f"{name} {what} {key}: {got:.3f}in rendered vs "
                    f"{want:.3f}in measured",
                ))
                return

    band_region, _ = elements[0]
    close(state["band"], band_region, "band", "rect")
    for (region, name), logo in zip(elements[1:], state["logos"]):
        close(logo, region, name, "rect")
        if not logo["loaded"]:
            findings.append(vd.Finding(
                f"page {i}", "web.brand_asset_broken",
                f"brand logo {name} did not load: {logo['src']}",
            ))

    band_hex = manifest.get("typography", {}).get("tokens", {}).get(
        "colors", {}).get("band", "").lstrip("#")
    if band_hex:
        want = tuple(int(band_hex[k:k + 2], 16) for k in (0, 2, 4))
        got = [int(c) for c in re.findall(r"\d+", state["band"]["color"])]
        if len(got) == 3 and tuple(got) != want:
            findings.append(vd.Finding(
                f"page {i}", "web.brand_color",
                f"band computed color rgb{tuple(got)} != manifest "
                f"tokens.colors.band #{band_hex.upper()}",
            ))

    # page number: content pages show their 1-based index, cover/closing none
    if deck_page.get("archetype") in ("cover", "closing"):
        if state["pageNumber"] is not None:
            findings.append(vd.Finding(
                f"page {i}", "web.page_number",
                f"cover/closing must carry no page number, shows "
                f"{state['pageNumber']!r}",
            ))
    elif state["pageNumber"] != str(i):
        findings.append(vd.Finding(
            f"page {i}", "web.page_number",
            f"page number shows {state['pageNumber']!r}, expected {i!r}",
        ))
    return findings


def _check_brand_pixels(page, deck_page: dict, manifest: dict | None,
                        i: int, notes: list) -> list:
    """Pixel spot-check of band color/position in the rendered stage."""
    if manifest is None:
        return []
    expect = _brand_expect(manifest, deck_page.get("archetype", ""))
    if expect is None:
        return []
    band = expect[1][0][0]
    band_hex = manifest.get("typography", {}).get("tokens", {}).get(
        "colors", {}).get("band", "").lstrip("#")
    if not band_hex:
        return []
    want = tuple(int(band_hex[k:k + 2], 16) for k in (0, 2, 4))

    try:
        from PIL import Image  # noqa: F401
    except ImportError:
        note = ("pixel spot-check skipped: PIL unavailable "
                "(pip install pillow) — DOM geometry/color checks still ran")
        if note not in notes:
            notes.append(note)
        return []

    import io
    from PIL import Image
    img = Image.open(io.BytesIO(page.locator(".stage").screenshot())).convert("RGB")

    def px(fx, fy):
        return img.getpixel((min(img.width - 1, round(fx * img.width)),
                             min(img.height - 1, round(fy * img.height))))

    def is_band(c):
        return all(abs(a - b) <= 12 for a, b in zip(c, want))

    def in_(fy):  # pixel row -> slide inches
        return fy / img.height * 7.5

    # sample column x=0.03W: left of every text/logo region (title axis 0.74in,
    # corner logo 0.27in but only at the page bottom), so band pixels are pure
    col, findings = 0.03, []
    kind = "cover" if deck_page.get("archetype") in ("cover", "closing") else "content"
    if kind == "content":
        sample_y = (band["y"] + band["h"] / 2) / 7.5
        scan = range(0, int(0.3 * img.height))  # band top(0) + bottom edge live here
    else:
        sample_y = (band["y"] + band["h"] / 2) / 7.5
        scan = range(int(0.15 * img.height), int(0.75 * img.height))

    if not is_band(px(col, sample_y)):
        findings.append(vd.Finding(
            f"page {i}", "web.brand_pixel",
            f"band sample at ({col:.0%}W, y={band['y'] + band['h'] / 2:.2f}in) "
            f"is rgb{px(col, sample_y)}, expected #{band_hex.upper()}",
        ))
    rows = [y for y in scan if is_band(px(col, y / img.height))]
    if not rows:
        findings.append(vd.Finding(
            f"page {i}", "web.brand_pixel",
            f"no band-colored pixels found scanning column {col:.0%}W",
        ))
        return findings
    top, bottom = in_(rows[0]), in_(rows[-1])
    height = bottom - top + in_(1)
    if abs(top - band["y"]) > 0.06 or abs(height - band["h"]) > 0.06:
        findings.append(vd.Finding(
            f"page {i}", "web.brand_pixel",
            f"band renders at y={top:.2f}in h={height:.2f}in (scan column "
            f"{col:.0%}W); measured y={band['y']}in h={band['h']}in",
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

    notes: list[str] = []
    try:
        with DevServer(web_dir, args.port, args.startup_timeout) as server:
            findings.extend(check_web(server.url, deck, args.screenshots,
                                      manifest, notes))
    except RuntimeError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    extra = {}
    if notes:
        extra["notes"] = "; ".join(notes)
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
