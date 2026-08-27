<h1 align="center">BLCU-PPT-Template</h1>

<p align="center">
  <strong>An agent skill that turns material and a one-line prompt into a
  group-meeting deck on the BLCU template: PPTX, web slideshow, speaker script.</strong>
</p>

<p align="center"><b>English</b> · <a href="README.zh-CN.md">简体中文</a></p>

---

Give your agent a document, some figures, and a prompt like “make this into a
group-meeting report”. The skill plans an outline, confirms it with you,
renders sample pages for approval, then produces the full deck on the BLCU
report template. Slides are cloned from the template pptx itself, so its brand
layer comes through exactly; formulas become editable PowerPoint math; fonts
are subset and embedded so the file renders correctly anywhere.

## Installation

```bash
# skills CLI — works with Claude Code, Cursor, Codex, and most agents
npx skills add ChHsiching/BLCU-PPT-Template

# or register as a local plugin directory for Claude Code
git clone https://github.com/ChHsiching/BLCU-PPT-Template.git
claude --plugin-dir BLCU-PPT-Template
```

Before rendering, place the template original at
`templates/blcu-report/blcu-report.pptx`. It is kept out of git on purpose;
everything else ships with the repo.

### Requirements

| Requirement | Used by | If missing |
|---|---|---|
| Python 3.11+ with python-pptx, latex2mathml, matplotlib | all rendering | hard stop |
| playwright + Chromium (optional PIL) | web QA gate | gate notes the skip |
| Node 24 + npm | web output only | pptx-only mode unaffected |
| PowerPoint COM on Windows | screenshot spot checks | downgrades to a note |

## Usage

Hand the agent material plus a prompt; [SKILL.md](SKILL.md) drives the rest:

```text
Material: docs + images (images named after their content)
Prompt:   用这份素材做一份组会汇报，pptx 和 web 都要
```

You confirm the outline (CP1), approve sample pages (CP2), receive the full
deck gated by scripted checks plus fresh-subagent review, then take delivery
(CP3). Every deliverable must pass its gates with zero findings.

## Deliverables

| File | Contents |
|---|---|
| `outline.md` | page plan and budgets, fixed after CP1 |
| `deck.json` | single source of truth both renderers consume |
| `out/<name>.pptx` | the deck: native OMML formulas, embedded fonts |
| `out/web/` | Vite + React slideshow (`npm run dev`) |
| `演讲稿.md` | standalone speaker script |

A complete worked example lives in
[`examples/qat-lsq-repro/`](examples/qat-lsq-repro/) — regenerate it:

```bash
python scripts/render_pptx.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/qat-report.pptx
python scripts/scaffold_web.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/web --force
```

## Development

```bash
python -m pytest tests/
```

Architecture, environment notes, and QA-gate details are documented in
[AGENTS.md](AGENTS.md); the domain glossary is [CONTEXT.md](CONTEXT.md).

[Issues](https://github.com/ChHsiching/BLCU-PPT-Template/issues) track history
and open work.
