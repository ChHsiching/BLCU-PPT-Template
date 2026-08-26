# BLCU-PPT-Template

A skill that generates PPT files or web-based PPT programs from the BLCU presentation template (see `templates/blcu-report/blcu-report.pptx`, kept local-only and gitignored), driven by user-supplied material and prompts. Designed to support multiple templates in the future.

## Agent skills

### Issue tracker

Issues are tracked in this repo's GitHub Issues via the `gh` CLI. See `docs/agents/issue-tracker.md`.

### Triage labels

Default five canonical triage labels, unchanged. See `docs/agents/triage-labels.md`.

### Domain docs

Single-context: `CONTEXT.md` + `docs/adr/` at the repo root. See `docs/agents/domain.md`.

## Environment notes

- Template original: `templates/blcu-report/blcu-report.pptx` (local-only, gitignored — never commit any `.pptx`; `.gitignore` has `templates/*/*.pptx` as the hard backstop).
- Python 3.11 available; `python-pptx` 1.0.2 installed; pytest 9.x installed — run tests with `python -m pytest tests/`. `scripts/validate_deck.py` is stdlib-only; `scripts/render_pptx.py` additionally uses `latex2mathml` + `matplotlib` (both installed) and `MML2OMML.XSL` (auto-located from the Office install, override with `BLCU_MML2OMML_XSL`; without it formulas fall back to rendered images). Font embedding (`scripts/embed_fonts.py`, manifest `fonts.embed` token) uses `fonttools` (installed; without it the render degrades to a declared-typeface + warning) and the full Noto Sans SC 400/700 TTFs + OFL license committed under `fonts/`. The fntdata payload must be spec-layout uncompressed EOT — PowerPoint ignores bare-TTF parts (empirically verified).
- PowerPoint is installed: COM automation available for render/screenshot QA; `MML2OMML.XSL` ships with the Office install (LaTeX→MathML→OMML chain).
- renderer-web: Node 24 + npm available. Preset scaffold at `assets/web-template/` (Vite + React + KaTeX, runs standalone via `npm install && npm run dev`); `scripts/scaffold_web.py <deck.json> -o <dir>` copies it into a presentation project. The npm-build e2e test needs npm on PATH and network for the first install. Web brand layer (bands/logos) renders from manifest `brand_layer` measured geometry; its assets live in the scaffold's `public/brand/`, kept in sync with the manifest's media fields by `scripts/export_brand_assets.py` (copy from `templates/<id>/extracted/media/` — rerun after editing brand_layer).
- QA gates (T6): `scripts/qa_check_pptx.py <pptx> --deck <deck.json>` (count/titles/residue/budget/structure; optional `--com-screenshots DIR` exports slide PNGs via PowerPoint COM, degrades to a note without COM) and `scripts/qa_check_web.py <web-dir>` (Playwright: pages/titles/KaTeX/images/brand-layer incl. band pixel spot-check/console; `--export-pptx OUT` also runs the web deck through renderer-pptx and aligns it). Python playwright 1.61 + Chromium are installed; the web QA needs `npm install` in the project first; the brand pixel spot-check additionally uses PIL (degrades to a note without it). Windows note baked into the script: web dirs are `Path.resolve()`d because vite breaks on 8.3 short-path cwds (`$TEMP` in Git Bash is short-form).
- End-to-end example: `examples/tracer-bullet/` (material → outline → deck.json → pptx + web + 演讲稿); regenerate with the commands in its README. `examples/tracer-bullet/out/` is gitignored.
- No LibreOffice on this machine.
- Shell is Git Bash but native Windows Python can't see Git Bash's `/tmp` — always pass absolute Windows paths when scripting.
- A rendered pptx open in PowerPoint holds a file lock (re-render fails with Permission denied) — `taskkill //IM POWERPNT.EXE //F` before re-rendering over an output you previously opened.
- Visual review via the Read tool routes through a CDN that **caches by file path**: re-checking an updated render under the same filename returns stale bytes. Always copy screenshots to a fresh filename before each vision check; adjudicate pixel-level disputes with PIL/numpy directly on the local file. Some localhost ports sit in Windows reserved ranges (listen fails with EACCES) — if a dev server hits EACCES, pick another port.
