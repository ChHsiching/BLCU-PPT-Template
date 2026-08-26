"""S1 font-embedding pipeline tests: scripts/embed_fonts.py + render_pptx wiring.

Contract under test: a deck rendered against the real manifest carries
p:embeddedFontLst + EOT fntdata parts for Noto Sans SC (regular + bold); the
EOT wrapper follows the layout PowerPoint itself serializes; each subset
covers every deck character, keeps the family name and weight flags, and
stays under 1MB per weight; without fonttools (or missing/corrupt font
files) the render still succeeds with the typeface declared and an explicit
warning surfaced through RenderStats.
"""
import io
import json
import re
import struct
import subprocess
import sys
import zipfile
from pathlib import Path

from fontTools.ttLib import TTFont
from lxml import etree

REPO = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = REPO / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import embed_fonts as efont  # noqa: E402
import render_pptx as rp  # noqa: E402

FIXTURE_DECK = REPO / "tests" / "fixtures" / "deck_pptx.json"
TEMPLATE_DIR = REPO / "templates" / "blcu-report"
TEMPLATE_PPTX = TEMPLATE_DIR / "blcu-report.pptx"
MANIFEST_PATH = TEMPLATE_DIR / "manifest.json"
SCRIPT = SCRIPTS_DIR / "render_pptx.py"
FONT_WEIGHT_MAX_BYTES = 1_000_000  # <1MB per weight (ticket S1)

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def load_manifest():
    return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))


def render_cli(tmp_path):
    out_path = tmp_path / "out.pptx"
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), str(FIXTURE_DECK), "-o", str(out_path)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    return out_path, proc


def parse_eot(data: bytes) -> dict:
    """Spec-layout EOT parser (the same layout PowerPoint serializes)."""
    r = {}
    (r["EOTSize"], r["FontDataSize"], r["Version"],
     r["Flags"]) = struct.unpack_from("<IIII", data, 0)
    r["charset"], r["italic"] = data[26], data[27]
    (r["weight"], r["fsType"], r["magic"]) = struct.unpack_from("<IHH", data, 28)
    off = 80  # after panose/charset/italic/weight/fsType/magic/ranges/reserved
    names = {}
    for key in ("Family", "Style", "Version", "Full", "RootString"):
        _, size = struct.unpack_from("<HH", data, off)
        off += 4
        names[key] = data[off:off + size].decode("utf-16-le", "replace")
        off += size
    r["names"] = names
    r["RootStringCheckSum"] = struct.unpack_from("<I", data, off)[0]
    off += 4 + 4  # RootStringCheckSum, EUDCCodePage
    off += 4 + 8  # Padding6+SignatureSize, EUDCFlags+EUDCFontSize
    r["font_data"] = data[off:]
    return r


def embedded_parts(out_path: Path) -> dict[str, bytes]:
    """{"face:weight": fntdata bytes} mapped from p:embeddedFontLst via rels."""
    with zipfile.ZipFile(out_path) as z:
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
        rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    rid_to_target = {rel.get("Id"): rel.get("Target") for rel in rels}
    parts = {}
    for font in pres.findall(f".//{{{NS_P}}}embeddedFont"):
        face = font.find(f"{{{NS_P}}}font").get("typeface")
        for child in font:
            tag = child.tag.split("}")[-1]
            if tag == "font":
                continue
            rid = child.get(f"{{{NS_R}}}id")
            parts[f"{face}:{tag}"] = "ppt/" + rid_to_target[rid]
    with zipfile.ZipFile(out_path) as z:
        return {k: z.read(v) for k, v in parts.items()}


# ---------------------------------------------------------------------------
# unit: character collection and the EOT wrapper
# ---------------------------------------------------------------------------

def test_collect_deck_characters_covers_every_rendered_field():
    deck = {
        "pages": [
            {"blocks": [
                {"type": "title", "text": "标题A"},
                {"type": "text", "text": "正文 b"},
                {"type": "list", "items": ["目录一", "目录二"]},
                {"type": "image", "path": "x.png", "caption": "图注。"},
                {"type": "formula", "latex": r"\alpha"},
            ]}
        ]
    }
    assert efont.collect_deck_characters(deck) == set("标题A正文b目录一二图注。- 0123456789")
    # latex is Cambria Math territory and never collected


def test_collect_deck_characters_keeps_emphasis_markers():
    # collection is raw: literal roles (caption/title) render ** verbatim,
    # so the subset must cover the marker character too
    deck = {"pages": [{"blocks": [
        {"type": "text", "text": "**关键**词正文"},
        {"type": "image", "path": "x.png", "caption": "**注**"},
    ]}]}
    assert efont.collect_deck_characters(deck) == set("**关键词正文注- 0123456789")


def test_wrap_eot_mirrors_powerpoint_layout():
    chars = set("嵌入字体测试ABC123")
    for weight, bold_bit in ((400, False), (700, True)):
        weight_name = "Regular" if weight == 400 else "Bold"
        ttf = efont.subset_font_bytes(
            REPO / "fonts" / f"NotoSansSC-{weight_name}.ttf", chars)
        eot = efont.wrap_eot(ttf, "Noto Sans SC", weight)
        r = parse_eot(eot)
        assert r["EOTSize"] == len(eot)
        assert r["Version"] == 0x00020002 and r["Flags"] == 0x1  # subset, no compression
        assert r["magic"] == 0x504C
        assert r["weight"] == weight
        assert r["names"]["Family"] == "Noto Sans SC\0"
        assert r["names"]["Style"] == f"{weight_name}\0"
        assert r["RootStringCheckSum"] == 0x50475342  # empty RootString
        assert len(r["font_data"]) == r["FontDataSize"]
        # the payload is a readable TTF that still answers to its family
        font = TTFont(io.BytesIO(r["font_data"]))
        assert str(font["name"].getName(1, 3, 1, 0x409)) == "Noto Sans SC"
        assert bool(font["OS/2"].fsSelection & 0x20) is bold_bit
        cmap = font.getBestCmap()
        assert all(ord(c) in cmap for c in chars)


# ---------------------------------------------------------------------------
# pipeline: the rendered fixture deck carries valid embedded subsets
# ---------------------------------------------------------------------------

def test_cli_embeds_fonts_into_fixture_deck(tmp_path):
    out_path, proc = render_cli(tmp_path)
    assert proc.returncode == 0, proc.stderr + proc.stdout
    assert re.search(r"embedded Noto Sans SC subset \(\d+ chars\): "
                     r"regular \d+KB, bold \d+KB", proc.stdout)
    with zipfile.ZipFile(out_path) as z:
        names = z.namelist()
        assert sorted(n for n in names if n.endswith(".fntdata")) == [
            "ppt/fonts/font1.fntdata", "ppt/fonts/font2.fntdata"]
        assert b'Extension="fntdata"' in z.read("[Content_Types].xml")
        pres = z.read("ppt/presentation.xml")
    assert b'<p:font typeface="Noto Sans SC"' in pres
    assert b'embedTrueTypeFonts="1"' in pres and b'saveSubsetFonts="1"' in pres
    # ECMA-376 CT_EmbeddedFontListEntry sequence: font, regular, bold, ...
    font_el = etree.fromstring(pres).find(f".//{{{NS_P}}}embeddedFont")
    assert [c.tag.split("}")[-1] for c in font_el] == ["font", "regular", "bold"]


def test_embedded_subsets_cover_deck_chars_under_1mb(tmp_path):
    deck = json.loads(FIXTURE_DECK.read_text(encoding="utf-8"))
    chars = efont.collect_deck_characters(deck)
    assert "*" in chars  # fixture carries **markers**; raw collection keeps them
    assert chars
    out_path, proc = render_cli(tmp_path)
    assert proc.returncode == 0, proc.stderr + proc.stdout
    parts = embedded_parts(out_path)
    assert set(parts) == {"Noto Sans SC:regular", "Noto Sans SC:bold"}
    for key, data in parts.items():
        assert len(data) < FONT_WEIGHT_MAX_BYTES, key
        r = parse_eot(data)
        font = TTFont(io.BytesIO(r["font_data"]))
        cmap = font.getBestCmap()
        missing = [c for c in chars if ord(c) not in cmap]
        assert missing == [], f"{key} misses deck characters"


def test_render_pipeline_stats_and_idempotent_reembed(tmp_path):
    deck = json.loads(FIXTURE_DECK.read_text(encoding="utf-8"))
    result = rp.render_deck(deck, load_manifest(), TEMPLATE_PPTX)
    if result.stats.fonts_embedded is None:  # environment without fonttools
        assert "fonttools" in (result.stats.font_warning or "")
        return
    assert result.stats.fonts_embedded["face"] == "Noto Sans SC"

    prs = result.presentation
    assert len(prs.part._element.findall(f".//{{{NS_P}}}embeddedFont")) == 1

    # embedding again on the same Presentation stays package-clean
    base = TEMPLATE_DIR
    again = efont.embed_fonts(
        prs, face="Noto Sans SC",
        weights={"regular": base / "../../fonts/NotoSansSC-Regular.ttf",
                 "bold": base / "../../fonts/NotoSansSC-Bold.ttf"},
        characters=efont.collect_deck_characters(deck), warn=lambda m: None)
    assert again is not None
    out_path = tmp_path / "twice.pptx"
    prs.save(out_path)
    with zipfile.ZipFile(out_path) as z:
        fntdata = [n for n in z.namelist() if n.endswith(".fntdata")]
        pres_xml = etree.fromstring(z.read("ppt/presentation.xml"))
    assert len(fntdata) == 2
    assert len(pres_xml.findall(f".//{{{NS_P}}}embeddedFont")) == 1


# ---------------------------------------------------------------------------
# degradation: embedding requested but impossible -> render survives + warning
# ---------------------------------------------------------------------------

def test_no_fonttools_degrades_to_warning(monkeypatch):
    monkeypatch.setattr(efont, "HAS_FONTTOOLS", False)
    deck = json.loads(FIXTURE_DECK.read_text(encoding="utf-8"))
    result = rp.render_deck(deck, load_manifest(), TEMPLATE_PPTX)
    assert result.stats.fonts_embedded is None
    assert "fonttools" in result.stats.font_warning
    prs = result.presentation
    assert prs.part._element.findall(f".//{{{NS_P}}}embeddedFont") == []
    assert not [r for r in prs.part.rels.values()
                if r.reltype == efont.FONT_REL_TYPE]
    # the render itself is intact: pages built, typefaces still declared
    assert len(prs.slides) == len(deck["pages"])


def test_missing_font_file_degrades_to_warning(tmp_path):
    manifest = load_manifest()
    manifest["fonts"]["embed"]["regular"] = "no/such/file.ttf"
    manifest["fonts"]["embed"]["bold"] = "gone.ttf"
    deck = json.loads(FIXTURE_DECK.read_text(encoding="utf-8"))
    result = rp.render_deck(deck, manifest, TEMPLATE_PPTX)
    assert result.stats.fonts_embedded is None
    assert "missing" in result.stats.font_warning
    prs = result.presentation
    assert prs.part._element.findall(f".//{{{NS_P}}}embeddedFont") == []
    out_path = tmp_path / "degraded.pptx"
    prs.save(out_path)  # degraded output is a normal, openable deck
    with zipfile.ZipFile(out_path) as z:
        assert not [n for n in z.namelist() if n.endswith(".fntdata")]


def test_corrupt_font_file_degrades_to_warning(tmp_path):
    bad = tmp_path / "bad.ttf"
    bad.write_bytes(b"not a font at all")
    manifest = load_manifest()
    manifest["fonts"]["embed"]["regular"] = str(bad)
    manifest["fonts"]["embed"]["bold"] = str(bad)
    deck = json.loads(FIXTURE_DECK.read_text(encoding="utf-8"))
    result = rp.render_deck(deck, manifest, TEMPLATE_PPTX)
    assert result.stats.fonts_embedded is None
    assert "font embedding failed" in (result.stats.font_warning or "")
    assert len(result.presentation.slides) == len(deck["pages"])


def test_embed_without_weights_degrades_to_warning():
    warnings = []
    out = efont.embed_fonts(object(), face="X", weights={}, characters=set("a"),
                            warn=warnings.append)
    assert out is None
    assert "no regular/bold weight" in warnings[0]


def test_manifest_without_fonts_token_embeds_nothing():
    deck = json.loads(FIXTURE_DECK.read_text(encoding="utf-8"))
    manifest = load_manifest()
    del manifest["fonts"]
    result = rp.render_deck(deck, manifest, TEMPLATE_PPTX)
    assert result.stats.fonts_embedded is None
    assert result.stats.font_warning is None
    prs = result.presentation
    assert not [r for r in prs.part.rels.values()
                if r.reltype == efont.FONT_REL_TYPE]
