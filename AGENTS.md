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
- Python 3.11 available; `python-pptx` not installed yet (pip install when first needed).
- PowerPoint is installed: COM automation available for render/screenshot QA; `MML2OMML.XSL` ships with the Office install (LaTeX→MathML→OMML chain).
- No LibreOffice on this machine.
- Shell is Git Bash but native Windows Python can't see Git Bash's `/tmp` — always pass absolute Windows paths when scripting.
