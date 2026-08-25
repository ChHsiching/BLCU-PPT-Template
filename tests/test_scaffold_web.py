"""CLI contract tests for scripts/scaffold_web.py.

Contract under test: a deck.json scaffolds into a runnable web project — the
preset Vite+React template copied, src/deck.json injected with image paths
rewritten to material/images/<basename>, the deck's template manifest copied
to src/manifest.json, and every deck-referenced image present under
public/material/images/ with no sample residue. Invalid decks never produce a
project; basename collisions and unsafe targets are refused. The node build
test (skipped without npm) proves a scaffolded project actually compiles.
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

import scaffold_web as sw  # noqa: E402

FIXTURE_DIR = REPO / "tests" / "fixtures"
FULL_DECK = FIXTURE_DIR / "deck.json"
NO_IMAGE_DECK = FIXTURE_DIR / "deck_pptx.json"
TEMPLATE_DIR = REPO / "templates" / "blcu-report"
MANIFEST_PATH = TEMPLATE_DIR / "manifest.json"
SCAFFOLD_DIR = REPO / "assets" / "web-template"
SCRIPT = SCRIPTS_DIR / "scaffold_web.py"


def scaffold(tmp_path, deck=FULL_DECK, *extra):
    out = tmp_path / "web"
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), str(deck), "-o", str(out), *extra],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    return out, proc


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


# ---------------------------------------------------------------------------
# happy path: the T2 full-archetype fixture becomes a complete project
# ---------------------------------------------------------------------------

def test_cli_scaffolds_full_fixture(tmp_path):
    out, proc = scaffold(tmp_path)
    assert proc.returncode == 0, proc.stderr + proc.stdout
    assert "7 page(s), 3 image(s)" in proc.stdout
    assert "npm install" in proc.stdout

    # preset scaffold copied (spot files; node state excluded)
    for rel in ("package.json", "package-lock.json", "vite.config.js",
                "index.html", "src/main.jsx", "src/App.jsx", "README.md"):
        assert (out / rel).is_file(), rel
    assert not (out / "node_modules").exists()

    # injected deck: fixture content with image paths rewritten
    injected = load(out / "src" / "deck.json")
    fixture = load(FULL_DECK)
    assert injected["template"] == "blcu-report"
    assert len(injected["pages"]) == 7
    images = [b for p in injected["pages"] for b in p["blocks"] if b["type"] == "image"]
    assert [b["path"] for b in images] == [
        "material/images/framework-diagram.png",
        "material/images/attention-heatmap.png",
        "material/images/main-results-bar.png",
    ]
    # non-image content untouched
    assert injected["pages"][2]["blocks"][1]["latex"] == fixture["pages"][2]["blocks"][1]["latex"]

    # manifest copied verbatim from the deck's template
    assert (out / "src" / "manifest.json").read_bytes() == MANIFEST_PATH.read_bytes()

    # every referenced image present; sample material wholesale-replaced
    listed = sorted(p.name for p in (out / "public" / "material" / "images").iterdir())
    assert listed == ["attention-heatmap.png", "framework-diagram.png", "main-results-bar.png"]


def _text_only(page):
    """Rewrite a page as a minimal legal text-formula page (no subhead/list/images)."""
    title = next(b["text"] for b in page["blocks"] if b["type"] == "title")
    page["archetype"] = "text-formula"
    page["blocks"] = [{"type": "title", "text": title}, {"type": "text", "text": "略。"}]
    return page


def test_scaffolded_project_tree_has_no_sample_residue(tmp_path):
    # a deck with no images at all must not leave the sample material behind
    deck = load(FULL_DECK)
    for i, page in enumerate(deck["pages"]):
        if page["archetype"] in ("text-image", "chart-focus"):
            deck["pages"][i] = _text_only(page)
    deck_path = tmp_path / "noimg.json"
    deck_path.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
    out, proc = scaffold(tmp_path, deck_path)
    assert proc.returncode == 0, proc.stderr + proc.stdout
    assert not (out / "public" / "material").exists()


def test_cli_scaffolds_deck_without_images(tmp_path):
    out, proc = scaffold(tmp_path, NO_IMAGE_DECK)
    assert proc.returncode == 0, proc.stderr + proc.stdout
    assert "3 page(s), 0 image(s)" in proc.stdout


def test_image_referenced_twice_is_copied_once_worth(tmp_path):
    deck = absolutize_images(load(FULL_DECK))
    dup = next(b for p in deck["pages"] for b in p["blocks"]
               if b["type"] == "image" and "caption" in b)
    dup2 = {"type": "image", "path": dup["path"]}
    deck["pages"][4]["blocks"].append(dup2)  # same file twice: legal, 3 slots used
    deck_path = tmp_path / "dup.json"
    deck_path.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
    out, proc = scaffold(tmp_path, deck_path)
    assert proc.returncode == 0, proc.stderr + proc.stdout
    listed = sorted(p.name for p in (out / "public" / "material" / "images").iterdir())
    assert listed == ["attention-heatmap.png", "framework-diagram.png", "main-results-bar.png"]


# ---------------------------------------------------------------------------
# gate behavior: invalid decks, collisions, unsafe targets
# ---------------------------------------------------------------------------

def test_cli_refuses_invalid_deck(tmp_path):
    deck = load(FULL_DECK)
    deck["pages"][2]["blocks"][0]["text"] = "超" * 26  # > text-formula title_max_chars
    bad = tmp_path / "bad.json"
    bad.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
    out, proc = scaffold(tmp_path, bad)
    assert proc.returncode == 1
    assert "budget" in proc.stdout
    assert not out.exists()  # invalid decks never produce a project


def test_cli_refuses_missing_image(tmp_path):
    deck = load(FULL_DECK)
    deck["pages"][5]["blocks"][1]["path"] = "material/images/ghost.png"
    bad = tmp_path / "ghost.json"
    bad.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
    out, proc = scaffold(tmp_path, bad)
    assert proc.returncode == 1
    assert "image.path_missing" in proc.stdout
    assert not out.exists()


def test_basename_collision_is_refused(tmp_path):
    # two different files both named pic.png, both legal on one text-image page
    for name in ("a", "b"):
        d = tmp_path / name
        d.mkdir()
        (d / "pic.png").write_bytes(b"png-" + name.encode())
    deck = load(FULL_DECK)
    ti = deck["pages"][4]  # text-image: keep title/subhead/text/list, swap images
    ti["blocks"] = [b for b in ti["blocks"] if b["type"] != "image"]
    ti["blocks"].append({"type": "image", "path": "a/pic.png"})
    ti["blocks"].append({"type": "image", "path": "b/pic.png"})
    deck["pages"][5] = _text_only(deck["pages"][5])  # chart-focus needs its image
    deck_path = tmp_path / "collide.json"
    deck_path.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
    out, proc = scaffold(tmp_path, deck_path)
    assert proc.returncode == 2
    assert "basename 'pic.png'" in (proc.stderr + proc.stdout)
    assert not out.exists()


def test_basename_collision_is_case_insensitive(tmp_path):
    # Pic.png vs pic.png: same flattened path on Windows/macOS targets
    a_dir = tmp_path / "a"
    b_dir = tmp_path / "b"
    a_dir.mkdir()
    b_dir.mkdir()
    (a_dir / "Pic.png").write_bytes(b"png-a")
    (b_dir / "pic.png").write_bytes(b"png-b")
    deck = absolutize_images(load(FULL_DECK))
    ti = deck["pages"][4]
    ti["blocks"] = [b for b in ti["blocks"] if b["type"] != "image"]
    ti["blocks"].append({"type": "image", "path": str(a_dir / "Pic.png")})
    ti["blocks"].append({"type": "image", "path": str(b_dir / "pic.png")})
    deck["pages"][5] = _text_only(deck["pages"][5])
    deck_path = tmp_path / "case-collide.json"
    deck_path.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
    out, proc = scaffold(tmp_path, deck_path)
    assert proc.returncode == 2
    assert "case-insensitive" in (proc.stderr + proc.stdout)
    assert not out.exists()


def test_existing_output_needs_force(tmp_path):
    out, proc = scaffold(tmp_path)
    assert proc.returncode == 0
    (out / "stale.txt").write_text("stale", encoding="utf-8")
    _, again = scaffold(tmp_path)
    assert again.returncode == 2
    assert "--force" in again.stderr
    assert (out / "stale.txt").is_file()  # refused run left the tree untouched
    _, forced = scaffold(tmp_path, FULL_DECK, "--force")
    assert forced.returncode == 0, forced.stderr + forced.stdout
    assert not (out / "stale.txt").exists()  # --force replaced the project


def test_empty_existing_output_is_usable_without_force(tmp_path):
    out = tmp_path / "web"
    out.mkdir()
    _, proc = scaffold(tmp_path)
    assert proc.returncode == 0, proc.stderr + proc.stdout
    assert (out / "package.json").is_file()


def test_refuses_to_target_the_scaffold_source(tmp_path):
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), str(FULL_DECK), "-o", str(SCAFFOLD_DIR), "--force"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    assert proc.returncode == 2
    assert "protected" in proc.stderr
    # and a parent of it (deleting the parent would delete the source)
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), str(FULL_DECK), "-o", str(REPO / "assets"), "--force"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    assert proc.returncode == 2
    assert "protected" in proc.stderr
    assert SCAFFOLD_DIR.is_dir()  # nothing was deleted


def test_refuses_to_target_the_templates_tree(tmp_path):
    # the templates dir holds the gitignored, unrecoverable .pptx original
    templates = REPO / "templates"
    for target in (templates, templates / "blcu-report", templates / "blcu-report" / "extracted"):
        proc = subprocess.run(
            [sys.executable, str(SCRIPT), str(FULL_DECK), "-o", str(target), "--force"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        assert proc.returncode == 2, target
        assert "protected" in proc.stderr
    assert MANIFEST_PATH.is_file()
    assert TEMPLATE_DIR.is_dir()
    assert (TEMPLATE_DIR / "blcu-report.pptx").is_file()


def test_cli_reports_env_errors_as_exit_2(tmp_path):
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), str(FULL_DECK), "-o", str(tmp_path / "o"),
         "--manifest", str(tmp_path / "missing-manifest.json")],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    assert proc.returncode == 2
    assert "Traceback" not in proc.stderr
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), str(tmp_path / "no-such-deck.json"), "-o", str(tmp_path / "o")],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    assert proc.returncode == 2
    assert "cannot read deck" in proc.stderr


# ---------------------------------------------------------------------------
# unit seam: pure helpers
# ---------------------------------------------------------------------------

def test_build_web_deck_rewrites_only_image_paths():
    deck = {
        "template": "blcu-report",
        "pages": [{
            "archetype": "text-image",
            "blocks": [
                {"type": "title", "text": "结构总览"},
                {"type": "image", "path": "img/thing.png", "caption": "原始相对路径"},
            ],
        }],
    }
    images = [("img/thing.png", Path("C:/somewhere/thing.png"))]
    web = sw.build_web_deck(deck, images)
    assert web["pages"][0]["blocks"][1]["path"] == "material/images/thing.png"
    assert web["pages"][0]["blocks"][1]["caption"] == "原始相对路径"
    assert web["pages"][0]["blocks"][0]["text"] == "结构总览"
    # the input deck is not mutated
    assert deck["pages"][0]["blocks"][1]["path"] == "img/thing.png"


# ---------------------------------------------------------------------------
# end to end: the scaffolded project compiles (needs npm; network for install)
# ---------------------------------------------------------------------------

npm = shutil.which("npm")
requires_npm = pytest.mark.skipif(npm is None, reason="npm not available")


@requires_npm
def test_scaffolded_project_builds(tmp_path):
    out, proc = scaffold(tmp_path)
    assert proc.returncode == 0, proc.stderr + proc.stdout
    install = subprocess.run(
        [npm, "install", "--no-audit", "--no-fund", "--loglevel=error"],
        cwd=out, capture_output=True, text=True, encoding="utf-8",
        errors="replace", timeout=600,
    )
    assert install.returncode == 0, install.stderr + install.stdout
    build = subprocess.run(
        [npm, "run", "build"], cwd=out, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=300,
    )
    assert build.returncode == 0, build.stderr + build.stdout
    assert (out / "dist" / "index.html").is_file()
    built = out / "dist" / "assets"
    assert any(built.glob("*.js")) and any(built.glob("*.css"))
    # deck images ship with the build under material/
    for name in ("framework-diagram.png", "attention-heatmap.png", "main-results-bar.png"):
        assert (out / "dist" / "material" / "images" / name).is_file()
