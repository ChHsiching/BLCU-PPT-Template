# BLCU-PPT-Template

An [agent skill](SKILL.md) that turns raw material — a document, some images,
one sentence of intent — into presentation deliverables that faithfully carry
the BLCU report template: a PPTX file, an optional web slideshow project, and
a standalone speaker script. Built for graduate-lab group-meeting reports;
designed to support multiple templates over time.

[中文说明（README.zh-CN.md）](README.zh-CN.md)

## What you get

| Input | Output |
|---|---|
| `material/` — one markdown/docx document + images | `outline.md` → `deck.json` (single source of truth) |
| A prompt: "用它做一份组会汇报" | `out/<name>.pptx` — native-OMML formulas, embedded fonts |
| Optional: web output mode | `out/web/` — Vite+React slideshow, `npm run dev` to present |
| Always | `演讲稿.md` — speaker script as a separate document |

## How it works

The skill drives the agent through numbered steps, each closed by a **gate**
(zero-defect completion criterion): material ingest → outline (+ AI-flavor
check) → CP1 user confirmation → sample pages → CP2 acceptance → full deck +
machine QA + final review → delivery. Two renderers consume the same
`deck.json`: **renderer-pptx** (Clone & Fill on the template original) and
**renderer-web** (Vite + React). Machine gates re-check every deliverable:
budgets, placeholder residue, titles, geometry, typography tokens, brand
layer, and pixel-level band checks on the web side.

## Repository layout

```
SKILL.md                  the skill: contract, discipline, 8-step flow
CONTEXT.md                domain glossary (素材/页型/门/容量预算 …)
references/               outline format · deck.json schema · archetype
                          semantics · reviewer contract · AI-flavor patterns
scripts/                  render_pptx · scaffold_web · validate_deck ·
                          qa_check_pptx · qa_check_web · embed_fonts …
templates/blcu-report/    template derivatives: manifest.json (geometry,
                          budgets, typography tokens), extracted media
assets/web-template/      preset web-slideshow scaffold copied into projects
fonts/                    Noto Sans SC 400/700 TTFs + OFL license (embedded
                          subset into every pptx)
examples/qat-lsq-repro/   end-to-end example: material → outline → deck.json
                          → pptx + web + 演讲稿
tests/                    pytest suite for scripts and gates
```

## Requirements

- Python 3.11+: `python-pptx`, `latex2mathml`, `matplotlib`; the web QA gate
  additionally uses `playwright` (+ Chromium) and optionally PIL.
- Node 24 + npm: only for the optional web output.
- PowerPoint (Windows + COM), only for COM screenshot spot checks; the gates
  degrade gracefully without it.
- `templates/blcu-report/blcu-report.pptx` is **not in the repository**
  (gitignored by design): place your copy of the original template there
  manually. Everything else ships with the repo.

## Try it

```bash
# run the test suite
python -m pytest tests/

# regenerate the example's three deliverables (template original required)
python scripts/render_pptx.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/qat-report.pptx
python scripts/scaffold_web.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/web --force

# from here, invoke the skill itself with your own material
```

Using the skill means handing material and a prompt to the agent — the flow,
gates, and checkpoints live in [SKILL.md](SKILL.md).

## Design doctrine

- **模板 = 母版**: the template's masters are untouchable; everything outside
  them (typography, hierarchy, accent color, spacing rhythm) is self-designed
  and token-driven from `manifest.json`.
- Hierarchy is expressed through weight + color (one accent green), not font
  sizes; body text stays pure black for projection.
- Formulas stay native OMML (LaTeX → MathML → OMML); fonts are subset and
  embedded (<1 MB) so any machine opens the deck design-faithful.
- Every gate is evidence, not vibes: machine checks plus fresh-subagent
  reviews looped to zero findings.

## Status

v1 complete: skill pipeline (T1–T7) and the visual style round (S1–S6)
landed, end-to-end verified on the bundled example. See
[Issues](https://github.com/ChHsiching/BLCU-PPT-Template/issues) for history.
