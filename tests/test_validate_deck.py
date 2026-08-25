"""Seam-2 unit tests for scripts/validate_deck.py.

Covers the pure function (deck dict + manifest + image root -> findings) and the
CLI contract (path/stdin in, exit code + report out). The fixture deck.json at
tests/fixtures/ exercises every archetype, block type, formula, image, caption.
"""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = REPO / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import validate_deck as vd  # noqa: E402

FIXTURE_DIR = REPO / "tests" / "fixtures"
FIXTURE_DECK = FIXTURE_DIR / "deck.json"
MANIFEST_PATH = REPO / "templates" / "blcu-report" / "manifest.json"
SCRIPT = SCRIPTS_DIR / "validate_deck.py"


def load_fixture():
    deck = json.loads(FIXTURE_DECK.read_text(encoding="utf-8"))
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    return deck, manifest


def validate(deck, manifest=None, image_root=FIXTURE_DIR):
    if manifest is None:
        _, manifest = load_fixture()
    return vd.validate_deck(deck, manifest, image_root)


def codes(findings):
    return [f.code for f in findings]


class TestTextWidth:
    def test_halfwidth_counts_half(self):
        assert vd.text_width("abcd") == pytest.approx(2.0)

    def test_cjk_counts_one(self):
        assert vd.text_width("中文") == pytest.approx(2.0)

    def test_fullwidth_punctuation_counts_one(self):
        assert vd.text_width("：，。") == pytest.approx(3.0)

    def test_mixed(self):
        # 汇报人=3, ":"=0.5, 张三=2, " "=0.5, "2026-08-25"=5
        assert vd.text_width("汇报人:张三 2026-08-25") == pytest.approx(11.0)


class TestFixtureValid:
    def test_fixture_passes_with_zero_findings(self):
        deck, manifest = load_fixture()
        assert validate(deck, manifest) == []

    def test_fixture_covers_all_six_archetypes(self):
        deck, _ = load_fixture()
        archetypes = {p["archetype"] for p in deck["pages"]}
        assert archetypes == {"cover", "agenda", "text-formula", "text-image", "chart-focus", "closing"}


class TestSchemaShape:
    def test_missing_top_level_fields(self):
        deck, _ = load_fixture()
        for field in ("template", "pages"):
            broken = copy.deepcopy(deck)
            del broken[field]
            findings = validate(broken)
            assert any(f.path == field and f.code == "schema.missing_field" for f in findings), field

    def test_empty_pages_rejected(self):
        deck, _ = load_fixture()
        deck["pages"] = []
        findings = validate(deck)
        assert "schema.empty_field" in codes(findings)
        assert any(f.path == "pages" for f in findings)

    def test_missing_page_fields(self):
        deck, _ = load_fixture()
        for field in ("archetype", "blocks"):
            broken = copy.deepcopy(deck)
            del broken["pages"][0][field]
            findings = validate(broken)
            assert any(f.path == f"pages[0].{field}" and f.code == "schema.missing_field" for f in findings), field

    def test_missing_block_fields(self):
        cases = [
            (0, 0, "type"),
            (0, 0, "text"),    # title block
            (1, 1, "items"),   # list block
            (2, 1, "latex"),   # formula block
            (4, 4, "path"),    # image block
        ]
        for page_i, block_i, field in cases:
            deck, _ = load_fixture()
            del deck["pages"][page_i]["blocks"][block_i][field]
            findings = validate(deck)
            expected_path = f"pages[{page_i}].blocks[{block_i}].{field}"
            assert any(f.path == expected_path and f.code == "schema.missing_field" for f in findings), expected_path

    def test_wrong_types_rejected(self):
        deck, _ = load_fixture()
        mutations = [
            ("pages", "not-a-list"),
            ("pages", 3),
        ]
        for field, value in mutations:
            broken = copy.deepcopy(deck)
            broken[field] = value
            assert any(f.code == "schema.type" and f.path == field for f in validate(broken)), (field, value)

        broken = copy.deepcopy(deck)
        broken["pages"][0]["blocks"] = {"type": "title"}
        assert any(f.code == "schema.type" and f.path == "pages[0].blocks" for f in validate(broken))

        broken = copy.deepcopy(deck)
        broken["pages"][0]["blocks"][0]["text"] = 5
        assert any(f.code == "schema.type" and f.path == "pages[0].blocks[0].text" for f in validate(broken))

        broken = copy.deepcopy(deck)
        broken["pages"][1]["blocks"][1]["items"] = "五 总结"
        assert any(f.code == "schema.type" and f.path == "pages[1].blocks[1].items" for f in validate(broken))

        broken = copy.deepcopy(deck)
        broken["pages"][1]["blocks"][1]["items"] = [1, 2]
        assert any(f.code == "schema.type" and f.path == "pages[1].blocks[1].items[0]" for f in validate(broken))

    def test_empty_or_whitespace_strings_rejected(self):
        deck, _ = load_fixture()
        deck["pages"][0]["blocks"][0]["text"] = "   "
        findings = validate(deck)
        assert any(f.code == "schema.empty_field" and f.path == "pages[0].blocks[0].text" for f in findings)

        deck, _ = load_fixture()
        deck["pages"][1]["blocks"][1]["items"] = ["", "二 相关工作"]
        findings = validate(deck)
        assert any(f.code == "schema.empty_field" and f.path == "pages[1].blocks[1].items[0]" for f in findings)

        deck, _ = load_fixture()
        deck["pages"][2]["blocks"][1]["latex"] = "  "
        findings = validate(deck)
        assert any(f.code == "schema.empty_field" and f.path == "pages[2].blocks[1].latex" for f in findings)

    def test_control_chars_rejected_in_text_fields(self):
        deck, _ = load_fixture()
        deck["pages"][0]["blocks"][1]["text"] = "汇报人:张三\n2026-08-25"
        findings = validate(deck)
        assert any(f.code == "schema.control_char" and f.path == "pages[0].blocks[1].text" for f in findings)

    def test_unknown_fields_rejected(self):
        deck, _ = load_fixture()
        deck["pageCount"] = 7
        assert any(f.code == "schema.unknown_field" and f.path == "pageCount" for f in validate(deck))

        deck, _ = load_fixture()
        deck["meta"]["venue"] = "北京"
        assert any(f.code == "schema.unknown_field" and f.path == "meta.venue" for f in validate(deck))

        deck, _ = load_fixture()
        deck["pages"][0]["notes"] = "x"
        assert any(f.code == "schema.unknown_field" and f.path == "pages[0].notes" for f in validate(deck))

        deck, _ = load_fixture()
        deck["pages"][0]["blocks"][0]["tex"] = "y"  # typo for text
        assert any(f.code == "schema.unknown_field" and f.path == "pages[0].blocks[0].tex" for f in validate(deck))

    def test_unknown_archetype_rejected(self):
        deck, _ = load_fixture()
        deck["pages"][3]["archetype"] = "hero"
        findings = validate(deck)
        assert any(
            f.code == "schema.unknown_archetype" and f.path == "pages[3].archetype" for f in findings
        )
        # message points at the manifest and the known archetypes
        msg = next(f.message for f in findings if f.code == "schema.unknown_archetype")
        assert "manifest" in msg and "text-image" in msg

    def test_unknown_block_type_rejected(self):
        deck, _ = load_fixture()
        deck["pages"][0]["blocks"].append({"type": "quote", "text": "引言"})
        findings = validate(deck)
        assert any(f.code == "schema.unknown_block_type" for f in findings)

    def test_title_required_exactly_once(self):
        deck, _ = load_fixture()
        del deck["pages"][5]["blocks"][0]
        assert any(f.code == "schema.missing_title" and f.path == "pages[5]" for f in validate(deck))

        deck, _ = load_fixture()
        deck["pages"][5]["blocks"].append({"type": "title", "text": "第二标题"})
        assert any(f.code == "schema.duplicate_title" for f in validate(deck))

    def test_at_most_one_subhead_and_one_list(self):
        deck, _ = load_fixture()
        deck["pages"][4]["blocks"].append({"type": "subhead", "text": "第二小标"})
        assert any(f.code == "schema.duplicate_subhead" for f in validate(deck))

        deck, _ = load_fixture()
        deck["pages"][1]["blocks"].append({"type": "list", "items": ["六 展望"]})
        assert any(f.code == "schema.duplicate_list" for f in validate(deck))

    def test_meta_fields(self):
        deck, _ = load_fixture()
        deck["meta"] = "张三"
        assert any(f.code == "schema.type" and f.path == "meta" for f in validate(deck))

        deck, _ = load_fixture()
        deck["meta"]["presenter"] = 42
        assert any(f.code == "schema.type" and f.path == "meta.presenter" for f in validate(deck))

        deck, _ = load_fixture()
        del deck["meta"]
        assert validate(deck) == []


def over(page_i, block_i, field, value):
    """Load the fixture and overwrite one field."""
    deck, _ = load_fixture()
    deck["pages"][page_i]["blocks"][block_i][field] = value
    return deck


class TestBudgets:
    def test_title_over_chars(self):
        # cover title_max_chars = 16
        deck = over(0, 0, "text", "基于对比学习的神经机器翻译方法研究进展")
        findings = validate(deck)
        f = next(f for f in findings if f.code == "budget.title_chars")
        assert f.path == "pages[0].blocks[0]"
        assert "title_max_chars" in f.message and "16" in f.message

    def test_text_blocks_count(self):
        # text-formula text_blocks_max = 2 -> add a third text block
        deck, _ = load_fixture()
        deck["pages"][3]["blocks"].append({"type": "text", "text": "第三段补充说明文字。"})
        findings = validate(deck)
        assert any(f.code == "budget.text_blocks_count" and f.path == "pages[3]" for f in findings)

    def test_text_block_over_chars(self):
        # text-formula text_block_max_chars = 130
        deck = over(3, 1, "text", "数" * 131)
        findings = validate(deck)
        f = next(f for f in findings if f.code == "budget.text_block_chars")
        assert f.path == "pages[3].blocks[1]"
        assert "text_block_max_chars" in f.message and "130" in f.message

    def test_text_total_over_chars(self):
        # two blocks of 110 chars each pass per-block limit (130) but exceed total (150)
        deck, _ = load_fixture()
        deck["pages"][3]["blocks"][1]["text"] = "甲" * 110
        deck["pages"][3]["blocks"][2]["text"] = "乙" * 110
        findings = validate(deck)
        f = next(f for f in findings if f.code == "budget.text_total_chars")
        assert f.path == "pages[3]"
        assert "text_total_max_chars" in f.message and "150" in f.message

    def test_text_block_forbidden_on_agenda(self):
        # agenda text_blocks_max = 0
        deck, _ = load_fixture()
        deck["pages"][1]["blocks"].append({"type": "text", "text": "说明"})
        findings = validate(deck)
        assert any(f.code == "budget.text_blocks_count" and f.path == "pages[1]" for f in findings)

    def test_list_items_count(self):
        # agenda list_items_max = 6 -> 7 items (fixture has 5)
        deck, _ = load_fixture()
        deck["pages"][1]["blocks"][1]["items"].extend(["六 附录：证明细节", "七 复现细节"])
        findings = validate(deck)
        assert any(f.code == "budget.list_items_count" and f.path == "pages[1].blocks[1]" for f in findings)

    def test_list_item_over_chars(self):
        # agenda list_item_max_chars = 30
        deck = over(1, 1, "items", ["一 " + "长" * 30])
        findings = validate(deck)
        f = next(f for f in findings if f.code == "budget.list_item_chars")
        assert f.path == "pages[1].blocks[1].items[0]"
        assert "list_item_max_chars" in f.message

    def test_list_forbidden_on_text_formula(self):
        # text-formula list_items_max = 0
        deck, _ = load_fixture()
        deck["pages"][2]["blocks"].append({"type": "list", "items": ["要点一"]})
        findings = validate(deck)
        assert any(f.code == "budget.list_items_count" and f.path == "pages[2]" for f in findings)

    def test_formula_count(self):
        # text-formula formulas_max = 6 -> 7 formulas
        deck, _ = load_fixture()
        for i in range(5):
            deck["pages"][2]["blocks"].append({"type": "formula", "latex": f"x_{i} = i"})
        findings = validate(deck)
        assert any(f.code == "budget.formulas_count" and f.path == "pages[2]" for f in findings)

    def test_formula_forbidden_on_cover(self):
        # cover formulas_max = 0
        deck, _ = load_fixture()
        deck["pages"][0]["blocks"].append({"type": "formula", "latex": "E = mc^2"})
        findings = validate(deck)
        assert any(f.code == "budget.formulas_count" and f.path == "pages[0]" for f in findings)

    def test_images_over_max(self):
        # text-image images_max = 4 -> 5 images
        deck, _ = load_fixture()
        for i in range(3):
            deck["pages"][4]["blocks"].append({"type": "image", "path": "material/images/attention-heatmap.png"})
        findings = validate(deck)
        assert any(f.code == "budget.images_count" and f.path == "pages[4]" for f in findings)

    def test_images_below_min(self):
        # chart-focus images_min = 1 -> remove the image
        deck, _ = load_fixture()
        deck["pages"][5]["blocks"] = [b for b in deck["pages"][5]["blocks"] if b["type"] != "image"]
        findings = validate(deck)
        f = next(f for f in findings if f.code == "budget.images_count")
        assert f.path == "pages[5]"
        assert "images_min" in f.message

    def test_caption_over_chars(self):
        # text-image caption_max_chars = 30
        deck = over(4, 4, "caption", "图" * 31)
        findings = validate(deck)
        f = next(f for f in findings if f.code == "budget.caption_chars")
        assert f.path == "pages[4].blocks[4].caption"
        assert "caption_max_chars" in f.message

    def test_caption_forbidden_on_chart_focus(self):
        # chart-focus caption_max_chars = 0
        deck, _ = load_fixture()
        deck["pages"][5]["blocks"][1]["caption"] = "主结果柱状图"
        findings = validate(deck)
        assert any(f.code == "budget.caption_chars" and f.path == "pages[5].blocks[1].caption" for f in findings)

    def test_subhead_over_chars(self):
        # text-image subhead_max_chars = 14
        deck = over(4, 1, "text", "一二三四五六七八九十一二三四五")
        findings = validate(deck)
        f = next(f for f in findings if f.code == "budget.subhead_chars")
        assert f.path == "pages[4].blocks[1]"
        assert "subhead_max_chars" in f.message and "14" in f.message

    def test_subhead_forbidden_on_text_formula(self):
        # text-formula subhead_max_chars = 0
        deck, _ = load_fixture()
        deck["pages"][2]["blocks"].append({"type": "subhead", "text": "符号说明"})
        findings = validate(deck)
        assert any(f.code == "budget.subhead_chars" and f.path == "pages[2].blocks[4]" for f in findings)


class TestImages:
    def test_missing_image_path(self):
        deck = over(4, 4, "path", "material/images/missing-diagram.png")
        findings = validate(deck)
        f = next(f for f in findings if f.code == "image.path_missing")
        assert f.path == "pages[4].blocks[4].path"
        assert "material/images/missing-diagram.png" in f.message

    def test_absolute_path_checked_too(self):
        deck, _ = load_fixture()
        deck["pages"][5]["blocks"][1]["path"] = "C:/definitely/not/here.png"
        assert any(f.code == "image.path_missing" for f in validate(deck))


class TestManifestIntegrity:
    def test_missing_budget_key_reported(self):
        deck, manifest = load_fixture()
        del manifest["archetypes"]["cover"]["budget"]["title_max_chars"]
        findings = validate(deck, manifest)
        assert any(f.code == "manifest.missing_key" for f in findings)


class TestFindingsContract:
    def test_every_finding_is_locatable(self):
        deck, _ = load_fixture()
        deck["pages"][3]["archetype"] = "hero"
        deck["pages"][3]["blocks"].append({"type": "text", "text": "x" * 200})
        deck["pages"][4]["blocks"].append({"type": "image", "path": "nope.png"})
        findings = validate(deck)
        assert findings
        for f in findings:
            assert f.path and f.code and f.message
            assert f.code.count(".") == 1  # "<category>.<name>"


class TestCli:
    def run_cli(self, *args, stdin=None, cwd=None):
        return subprocess.run(
            [sys.executable, str(SCRIPT), *args],
            input=stdin,
            capture_output=True,
            text=True,
            encoding="utf-8",
            cwd=cwd,
        )

    def test_valid_deck_path_exit_0(self):
        result = self.run_cli(str(FIXTURE_DECK))
        assert result.returncode == 0
        assert "OK" in result.stdout
        assert "7" in result.stdout  # page count in summary

    def test_stdin_deck_with_cwd_image_root(self):
        payload = FIXTURE_DECK.read_text(encoding="utf-8")
        result = self.run_cli("-", stdin=payload, cwd=FIXTURE_DIR)
        assert result.returncode == 0

    def test_invalid_deck_exit_1_with_locatable_report(self):
        deck, _ = load_fixture()
        deck["pages"][0]["blocks"][0]["text"] = "基于对比学习的神经机器翻译方法研究进展"
        bad = FIXTURE_DIR.parent / "_bad_deck.json"
        bad.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
        try:
            result = self.run_cli(str(bad))
        finally:
            bad.unlink()
        assert result.returncode == 1
        assert "INVALID" in result.stdout
        assert "pages[0].blocks[0]" in result.stdout
        assert "budget.title_chars" in result.stdout

    def test_json_format(self):
        deck, _ = load_fixture()
        deck["pages"][5]["archetype"] = "hero"
        bad = FIXTURE_DIR.parent / "_bad_deck.json"
        bad.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
        try:
            result = self.run_cli("--format", "json", str(bad))
        finally:
            bad.unlink()
        assert result.returncode == 1
        report = json.loads(result.stdout)
        assert report["valid"] is False
        assert report["findings"]
        for f in report["findings"]:
            assert set(f) == {"path", "code", "message"}

    def test_json_format_valid(self):
        result = self.run_cli("--format", "json", str(FIXTURE_DECK))
        assert result.returncode == 0
        report = json.loads(result.stdout)
        assert report == {"valid": True, "page_count": 7, "findings": []}

    def test_bad_json_exit_2(self):
        bad = FIXTURE_DIR.parent / "_bad_deck.json"
        bad.write_text("{not json", encoding="utf-8")
        try:
            result = self.run_cli(str(bad))
        finally:
            bad.unlink()
        assert result.returncode == 2
        assert result.stderr.strip()

    def test_non_utf8_deck_exit_2(self):
        # a deck saved as ANSI/GBK must be a usage error, not a crash or exit 1
        bad = FIXTURE_DIR.parent / "_gbk_deck.json"
        bad.write_bytes('{"template": "blcu-report", "标题": "x"}'.encode("gbk"))
        try:
            result = self.run_cli(str(bad))
        finally:
            bad.unlink()
        assert result.returncode == 2
        assert result.stderr.strip()
        assert "Traceback" not in result.stderr

    def test_missing_file_exit_2(self):
        result = self.run_cli(str(FIXTURE_DIR / "nope.json"))
        assert result.returncode == 2
        assert result.stderr.strip()

    def test_unknown_template_exit_1(self):
        deck, _ = load_fixture()
        deck["template"] = "blcu-reprot"  # typo
        bad = FIXTURE_DIR.parent / "_bad_deck.json"
        bad.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
        try:
            result = self.run_cli(str(bad))
        finally:
            bad.unlink()
        assert result.returncode == 1
        assert "template" in result.stdout

    def test_manifest_override(self):
        result = self.run_cli("--manifest", str(MANIFEST_PATH), str(FIXTURE_DECK))
        assert result.returncode == 0
