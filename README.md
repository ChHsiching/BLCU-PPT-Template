<h1 align="center">BLCU-PPT-Template</h1>

<p align="center">
  <strong>素材 + 一句提示词，生成继承 BLCU 汇报模板的组会演示：PPTX、Web 放映、演讲稿</strong>
</p>

<p align="center"><b>English</b> · <a href="README.zh-CN.md">简体中文</a></p>

---

An [agent skill](SKILL.md) for graduate-lab group-meeting presentations. Give
the agent a document, some figures, and a prompt such as “make this into a
group-meeting report”; the skill runs outline planning, sample-page approval,
and gated rendering to produce deliverables built on the BLCU template
original.

The workflow is linear: material ingest, outline with capacity budgets,
outline confirmation (CP1), sample pages verified by machine checks and review
(CP2), then the full deck. Two renderers consume one intermediate file:
renderer-pptx clones slides from the template pptx at the XML level, and
renderer-web produces a Vite + React slideshow from the same deck.json. Each
deliverable passes scripted quality gates covering text budgets, placeholder
residue, geometry, run-level typography, the brand layer, and (for web) pixel
checks of the band; review gates loop fresh-subagent findings to zero.

## What makes the output faithful

- Slides are cloned from the template original, so the top band, logos, and
  page numbers come from the template's own master shapes instead of being
  redrawn approximations.
- All styling outside the masters lives in one token table
  ([`templates/blcu-report/manifest.json`](templates/blcu-report/manifest.json)):
  typeface, weights, sizes, colors, spacing. Both renderers and both QA
  scripts read these tokens.
- Formulas convert from LaTeX to native OMML and remain editable in
  PowerPoint. Noto Sans SC is subset to the characters used and embedded, so
  the file renders correctly on machines without the font installed.
- QA scripts exist as tools, not documentation: `validate_deck.py`,
  `qa_check_pptx.py`, and `qa_check_web.py` re-check every deliverable and
  exit nonzero on any finding.

## Install

```bash
# skills CLI, works with Claude Code, Cursor, Codex, and most agents
npx skills add ChHsiching/BLCU-PPT-Template

# or register as a local plugin directory for Claude Code
git clone https://github.com/ChHsiching/BLCU-PPT-Template.git
claude --plugin-dir BLCU-PPT-Template
```

Then provide material and a prompt; [SKILL.md](SKILL.md) defines what the
agent does next.

### Requirements

| Requirement | Used by | If missing |
|---|---|---|
| Python 3.11+ with python-pptx, latex2mathml, matplotlib | all rendering | hard stop |
| Template original at `templates/blcu-report/blcu-report.pptx` | rendering | hard stop (place manually; kept out of git by design) |
| playwright + Chromium, optionally PIL | web QA gate | gate downgrades to a note |
| Node 24 + npm | web output only | pptx-only mode unaffected |
| PowerPoint COM on Windows | screenshot spot checks | downgrades to a note |

## Deliverables

| File | Contents |
|---|---|
| `outline.md` | page plan and budget math, fixed after CP1 confirmation |
| `deck.json` | single source of truth consumed by both renderers |
| `out/<name>.pptx` | the deck: editable OMML formulas, embedded fonts |
| `out/web/` | Vite + React slideshow (`npm run dev`), exports back to pptx |
| `演讲稿.md` | standalone speaker script, kept out of pptx notes |

A complete worked example ships in
[`examples/qat-lsq-repro/`](examples/qat-lsq-repro/): an 11-page
quantization-aware-training repro with five native formulas, four figures,
and experiment numbers traced end-to-end from `material/make_figures.py`.

## Repository layout

```
SKILL.md                  contract, discipline, 8-step flow (the skill itself)
CONTEXT.md                domain glossary: 素材, 门, 容量预算, 页型 …
references/               outline format, deck schema, reviewer contract,
                          AI-flavor patterns
scripts/                  render_pptx, scaffold_web, validate_deck,
                          qa_check_pptx, qa_check_web, embed_fonts, …
templates/blcu-report/    manifest.json (geometry/budget/token source of
                          truth), extracted brand media [pptx original: local-only]
assets/web-template/      web-slideshow scaffold copied into projects
fonts/                    Noto Sans SC 400/700 TTFs with OFL license
examples/qat-lsq-repro/   end-to-end example
tests/                    pytest suite over scripts and gates
```

## Try it

```bash
python -m pytest tests/

python scripts/render_pptx.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/qat-report.pptx

python scripts/scaffold_web.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/web --force
```

## Status

v1 shipped: the pipeline (T1–T7) and the visual style round (S1–S6), verified
end-to-end on the bundled example. History and open work live in
[Issues](https://github.com/ChHsiching/BLCU-PPT-Template/issues).
