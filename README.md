<h1 align="center">BLCU-PPT-Template</h1>

<p align="center">
  <strong>素材 + 一句提示词 → BLCU 模板保真的组会汇报：PPTX · Web 放映 · 演讲稿</strong>
</p>

<p align="center"><b>English</b> · <a href="README.zh-CN.md">简体中文</a></p>

---

An [agent skill](SKILL.md) for graduate-lab group-meeting decks. Feed it your
paper notes and a one-line prompt; it walks the agent through outline → gates
→ a deck that carries the real BLCU template — native OMML formulas, embedded
Noto Sans SC, the template's own masters untouched.

```
 素材 ingest      大纲 + AI 味门       CP1 你确认        样张先行
┌──────────┐   ┌───────────┐   ┌────────────┐   ┌────────────┐
│ doc / md │──▶│ outline.md│──▶│  决策回写    │──▶│ cover+最满页 │
│ images/  │   │ 预算核对   │   │            │   │ G1 双门     │
└──────────┘   └───────────┘   └────────────┘   └────────────┘
                                                      │ CP2 你验收
      三产物 ◀── 渲染 ◀── deck.json（单一真相源）◀─────┘
   ┌─────────┐  ┌─────────┐  ┌──────────┐
   │  PPTX   │  │ web 放映 │  │ 演讲稿.md │
   │ 公式可编辑│  │ npm dev │  │ 口头讲稿  │
   └─────────┘  └─────────┘  └──────────┘
        ▲ 每步过门：机器检查 + fresh-subagent 审查，循环至零缺陷
```

## Why it's different

- **The template is sacred.** Slides are cloned from the original pptx, so
  the green bands, logos, and page numbers are the template's *own* master
  shapes — not a look-alike redrawn from scratch.
- **Content form is designed, token-driven.** Everything outside the masters
  (weights, hierarchy, accent green `#548235`, spacing rhythm) lives in one
  machine-readable token table in [`templates/blcu-report/manifest.json`](templates/blcu-report/manifest.json);
  both renderers and both QA gates read the same tokens.
- **Gates are evidence, not vibes.** Every deliverable re-passes scripted
  checks (budgets, residue, geometry, typography at run level, brand layer,
  band-pixel spot checks) plus fresh-subagent reviews looped to zero findings.
- **Formulas stay editable.** LaTeX → MathML → OMML, no screenshots. Fonts
  are subset and embedded (<1 MB), so any machine opens it design-faithful.

## Install

### Use in your coding agent

```bash
# any agent supported by the skills CLI (Claude Code, Cursor, Codex, …)
npx skills add ChHsiching/BLCU-PPT-Template

# or Claude Code directly, as a local plugin directory
git clone https://github.com/ChHsiching/BLCU-PPT-Template.git
claude --plugin-dir BLCU-PPT-Template
```

Then just hand the agent material and say something like
「用这份素材做一份组会汇报」— [SKILL.md](SKILL.md) takes over.

### Prerequisites

| Requirement | Needed for | Missing → |
|---|---|---|
| Python 3.11+ with python-pptx, latex2mathml, matplotlib | everything | hard stop |
| Template original at `templates/blcu-report/blcu-report.pptx` | rendering | hard stop (place manually; kept out of git on purpose) |
| playwright + Chromium, optional PIL | web QA gate | that gate degrades to a note |
| Node 24 + npm | web output only | pptx-only mode still works |
| PowerPoint COM (Windows) | screenshot spot checks | degrade to a note |

## What comes out

| File | What it is |
|---|---|
| `outline.md` | page plan + budget math, fixed after you confirm at CP1 |
| `deck.json` | single source of truth both renderers consume |
| `out/<name>.pptx` | the deliverable: native formulas, embedded fonts |
| `out/web/` | Vite + React slideshow (`npm run dev`), exports back to pptx |
| `演讲稿.md` | standalone speaker script — never stuffed into pptx notes |

A complete worked example ships in
[`examples/qat-lsq-repro/`](examples/qat-lsq-repro/) — quantization-aware
training repro, 11 pages, five native formulas, four figures, numbers traced
end-to-end from `make_figures.py`.

## Repository layout

```
SKILL.md                  contract + discipline + 8-step flow (the skill)
CONTEXT.md                domain glossary: 素材/门/容量预算/页型 …
references/               outline format · deck schema · reviewer contract
scripts/                  render_pptx · scaffold_web · validate_deck ·
                          qa_check_pptx · qa_check_web · embed_fonts …
templates/blcu-report/    manifest.json (geometry/budget/tokens truth),
                          extracted brand media     [original pptx: local-only]
assets/web-template/      preset web-slideshow scaffold
fonts/                    Noto Sans SC 400/700 + OFL license
examples/qat-lsq-repro/   end-to-end example
tests/                    130-test suite over scripts & gates
```

## Try it

```bash
python -m pytest tests/                                  # 130 tests

python scripts/render_pptx.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/qat-report.pptx        # rebuild the demo pptx

python scripts/scaffold_web.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/web --force             # rebuild the web project
```

## Status

v1 shipped: pipeline (T1–T7) + visual style round (S1–S6), verified
end-to-end on the bundled example. Roadmap and history live in
[Issues](https://github.com/ChHsiching/BLCU-PPT-Template/issues).
