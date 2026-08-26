"""Tests for the T6 QA gates: scripts/qa_check_pptx.py (fast, local) and
scripts/qa_check_web.py (e2e, needs npm + Playwright).

pptx-side contract under test: a correctly rendered fixture deck passes with
zero findings, while each injected defect class — placeholder residue,
over-budget text, page-count mismatch, title mismatch, off-canvas geometry,
crossed safe bottom — is caught by its own finding. The web-side contract:
a scaffolded project passes the Playwright checks (pages reachable, titles
aligned, KaTeX rendered, images loaded, no console errors), a broken image is
flagged, and --export-pptx produces a pptx that aligns with the web deck.
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = REPO / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import qa_check_pptx as qp  # noqa: E402
import render_pptx as rp  # noqa: E402

FIXTURE_DIR = REPO / "tests" / "fixtures"
FULL_DECK = FIXTURE_DIR / "deck.json"
PPTX_DECK = FIXTURE_DIR / "deck_pptx.json"
TEMPLATE_DIR = REPO / "templates" / "blcu-report"
TEMPLATE_PPTX = TEMPLATE_DIR / "blcu-report.pptx"
MANIFEST_PATH = TEMPLATE_DIR / "manifest.json"
RENDER_SCRIPT = SCRIPTS_DIR / "render_pptx.py"
QA_PPTX_SCRIPT = SCRIPTS_DIR / "qa_check_pptx.py"
QA_WEB_SCRIPT = SCRIPTS_DIR / "qa_check_web.py"
SCAFFOLD_SCRIPT = SCRIPTS_DIR / "scaffold_web.py"


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def absolutize_images(deck):
    """Point deck image paths at the fixture files (decks written outside
    fixtures/ would otherwise resolve images against their own directory)."""
    for page in deck["pages"]:
        for block in page["blocks"]:
            if block["type"] == "image":
                block["path"] = str((FIXTURE_DIR / block["path"]).resolve())
    return deck


def write_deck(deck, path):
    Path(path).write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
    return Path(path)


def render_cli(deck_path, out_path):
    return subprocess.run(
        [sys.executable, str(RENDER_SCRIPT), str(deck_path), "-o", str(out_path)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )


def qa_pptx(pptx, deck, *extra):
    return subprocess.run(
        [sys.executable, str(QA_PPTX_SCRIPT), str(pptx), "--deck", str(deck), *extra],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )


xsl = rp.find_mml2omml_xsl()
requires_math = pytest.mark.skipif(xsl is None or not rp.HAS_LATEX2MATHML,
                                   reason="MML2OMML.XSL or latex2mathml unavailable")


# ---------------------------------------------------------------------------
# green path
# ---------------------------------------------------------------------------

@requires_math
def test_cli_green_on_rendered_fixture(tmp_path):
    out = tmp_path / "out.pptx"
    proc = render_cli(FULL_DECK, out)
    assert proc.returncode == 0, proc.stderr + proc.stdout  # not vacuously green
    proc = qa_pptx(out, FULL_DECK)
    assert proc.returncode == 0, proc.stdout
    assert "OK: 7 slide(s), 0 findings" in proc.stdout


@requires_math
def test_cli_green_json_format(tmp_path):
    out = tmp_path / "out.pptx"
    assert render_cli(FULL_DECK, out).returncode == 0
    proc = qa_pptx(out, FULL_DECK, "--format", "json")
    assert proc.returncode == 0
    report = json.loads(proc.stdout)
    assert report == {"valid": True, "slide_count": 7, "findings": []}


# ---------------------------------------------------------------------------
# injected defects are each caught by their own finding
# ---------------------------------------------------------------------------

@requires_math
def test_catches_placeholder_residue(tmp_path):
    deck = absolutize_images(load(FULL_DECK))
    deck["pages"][3]["blocks"][1]["text"] = "神经机器翻译在低资源场景下性能受限，TODO：此处填写数据。"
    deck_path = write_deck(deck, tmp_path / "residue.json")
    out = tmp_path / "residue.pptx"
    assert render_cli(deck_path, out).returncode == 0  # schema-legal: render passes
    proc = qa_pptx(out, deck_path)
    assert proc.returncode == 1
    findings = [line for line in proc.stdout.splitlines() if "[pptx.residue]" in line]
    assert any("[todo]" in line for line in findings)
    assert any("[此处填写]" in line for line in findings)
    assert "slides[4]" in findings[0]  # slide numbers are 1-based


def test_catches_over_budget_deck(tmp_path):
    # rendered through the API, bypassing the CLI's validation gate: the QA
    # revalidation is what must flag it（超限量从 manifest 读取，预算重校准后随动）
    deck = absolutize_images(load(FULL_DECK))
    manifest = load(MANIFEST_PATH)
    cap = manifest["archetypes"][deck["pages"][2]["archetype"]]["budget"]["text_total_max_chars"]
    deck["pages"][2]["blocks"].append({"type": "text", "text": "超" * (cap + 1)})
    result = rp.render_deck(deck, manifest, TEMPLATE_PPTX, image_root=FIXTURE_DIR)
    out = tmp_path / "over.pptx"
    result.presentation.save(out)
    proc = qa_pptx(out, write_deck(deck, tmp_path / "over.json"))
    assert proc.returncode == 1
    assert "[budget.text_block_chars]" in proc.stdout
    assert "[budget.text_total_chars]" in proc.stdout


@requires_math
def test_catches_page_count_mismatch(tmp_path):
    out = tmp_path / "three.pptx"
    assert render_cli(PPTX_DECK, out).returncode == 0  # 3 pages
    proc = qa_pptx(out, FULL_DECK)  # checked against the 7-page deck
    assert proc.returncode == 1
    assert "[pptx.page_count]" in proc.stdout


@requires_math
def test_catches_title_mismatch(tmp_path):
    deck = load(PPTX_DECK)
    out = tmp_path / "three.pptx"
    assert render_cli(PPTX_DECK, out).returncode == 0
    deck["pages"][1]["blocks"][0]["text"] = "被篡改的标题"
    proc = qa_pptx(out, write_deck(deck, tmp_path / "retitled.json"))
    assert proc.returncode == 1
    assert "[pptx.title_mismatch]" in proc.stdout
    assert "被篡改的标题" in proc.stdout


@requires_math
def test_catches_canvas_and_safe_bottom_overflows(tmp_path):
    from pptx import Presentation
    from pptx.util import Inches

    out = tmp_path / "base.pptx"
    assert render_cli(FULL_DECK, out).returncode == 0
    prs = Presentation(out)
    agenda_list = next(sp for sp in prs.slides[1].shapes if sp.name == "AgendaList")
    agenda_list.top = Inches(7.0)  # bottom lands at 12.2in, off-canvas
    text_area = next(sp for sp in prs.slides[2].shapes if sp.name == "TextArea")
    text_area.top, text_area.height = Inches(4.32), Inches(3.0)  # bottom 7.32 > 6.9
    broken = tmp_path / "overflow.pptx"
    prs.save(broken)
    proc = qa_pptx(broken, FULL_DECK)
    assert proc.returncode == 1
    assert "[pptx.canvas_overflow]" in proc.stdout
    assert "[pptx.content_bottom]" in proc.stdout


def test_exit_2_on_usage_and_io_errors(tmp_path):
    proc = qa_pptx(tmp_path / "ghost.pptx", FULL_DECK)
    assert proc.returncode == 2
    assert "pptx not found" in proc.stderr
    proc = subprocess.run(
        [sys.executable, str(QA_PPTX_SCRIPT), str(tmp_path / "any.pptx"),
         "--deck", str(tmp_path / "no-deck.json")],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    assert proc.returncode == 2
    assert "Traceback" not in proc.stderr


# ---------------------------------------------------------------------------
# unit seam: the residue matcher
# ---------------------------------------------------------------------------

def test_find_residue_matches_fillers_without_false_positives():
    for text in ("TODO", "todo:", "FIXME now", "meeting TBD", "xxxx", "lorem ipsum",
                 "[insert name]", "<insert>here", "a placeholder b", "Sample Text",
                 "待补充", "此处填写内容", "示例文本"):
        assert qp.find_residue(text), text
    for clean in ("拓扑排序", "contrastive learning", "TODOor 的边界情况",
                  "maximization", "EXP（指数）", "插入 insert 图片", "西安 TBDX 公司",
                  "可学习占位符", "占位效应", "待定系数法", "时间待定"):
        # "TODOor" / "TBDX" / bare "insert" stay legal: the patterns need word
        # boundaries or a bracket tag; bare 占位/待定 stay legal because they
        # live in real academic terms (可学习占位符 / 待定系数法)
        assert not qp.find_residue(clean), clean


def test_qa_pptx_stays_total_on_schema_broken_deck(tmp_path):
    # a deck edited after rendering (page missing its blocks) must yield a
    # findings report, never a traceback — the tool exists to judge such decks
    out = tmp_path / "out.pptx"
    assert render_cli(PPTX_DECK, out).returncode == 0
    deck = {"template": "blcu-report", "pages": [{"archetype": "cover"}]}
    proc = qa_pptx(out, write_deck(deck, tmp_path / "broken.json"))
    assert proc.returncode == 1
    assert "Traceback" not in proc.stderr
    assert "schema" in proc.stdout


def test_web_qa_reports_invalid_deck_without_a_browser(tmp_path):
    # schema-broken project decks are caught by validation before any
    # dev server / browser starts
    web = tmp_path / "web"
    (web / "src").mkdir(parents=True)
    (web / "package.json").write_text("{}", encoding="utf-8")
    (web / "src" / "deck.json").write_text(
        json.dumps({"template": "blcu-report", "pages": [{"archetype": "cover"}]}),
        encoding="utf-8")
    shutil.copyfile(MANIFEST_PATH, web / "src" / "manifest.json")
    (web / "node_modules").mkdir()  # present, so the earlier env check passes
    proc = subprocess.run(
        [sys.executable, str(QA_WEB_SCRIPT), str(web)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    assert proc.returncode == 1
    assert "schema.missing_field" in proc.stdout
    assert "Traceback" not in proc.stderr


# ---------------------------------------------------------------------------
# e2e web QA (needs npm for the dev server + playwright with Chromium)
# ---------------------------------------------------------------------------

npm = shutil.which("npm")
try:
    import playwright.sync_api as _pw  # noqa: F401

    has_playwright = True
except ImportError:
    has_playwright = False
requires_web = pytest.mark.skipif(
    npm is None or not has_playwright,
    reason="npm or python playwright unavailable")


@requires_web
def test_web_qa_green_then_defect_then_export(tmp_path):
    web = tmp_path / "web"
    proc = subprocess.run(
        [sys.executable, str(SCAFFOLD_SCRIPT), str(FULL_DECK), "-o", str(web)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    assert proc.returncode == 0, proc.stderr + proc.stdout
    install = subprocess.run(
        [npm, "install", "--no-audit", "--no-fund", "--loglevel=error"],
        cwd=web, capture_output=True, text=True, encoding="utf-8",
        errors="replace", timeout=600,
    )
    assert install.returncode == 0, install.stderr + install.stdout

    shots = tmp_path / "shots"
    exported = tmp_path / "exported.pptx"
    proc = subprocess.run(
        [sys.executable, str(QA_WEB_SCRIPT), str(web),
         "--screenshots", str(shots), "--export-pptx", str(exported)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=300,
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "0 findings" in proc.stdout
    assert exported.is_file()
    assert sorted(p.name for p in shots.glob("page-*.png")) == [
        f"page-{i:02d}.png" for i in range(1, 8)]

    # defect phase: a missing deck image must fail the gate (caught by the
    # upfront deck validation, before any browser is involved)
    (web / "public" / "material" / "images" / "main-results-bar.png").unlink()
    proc = subprocess.run(
        [sys.executable, str(QA_WEB_SCRIPT), str(web)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=300,
    )
    assert proc.returncode == 1
    assert "image.path_missing" in proc.stdout


def test_web_qa_exit_2_without_node_modules(tmp_path):
    web = tmp_path / "bare"
    (web / "src").mkdir(parents=True)
    (web / "package.json").write_text("{}", encoding="utf-8")
    (web / "src" / "deck.json").write_text('{"template": "blcu-report", "pages": []}',
                                           encoding="utf-8")
    (web / "src" / "manifest.json").write_text("{}", encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, str(QA_WEB_SCRIPT), str(web)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    assert proc.returncode == 2
    assert "npm install" in proc.stderr
