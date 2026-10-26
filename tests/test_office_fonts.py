"""The faces PowerPoint would use on this machine: found in place, measured, drawn.

None of this needs Office.  The folders :mod:`pptx2svg.fonts.office` searches are pointed
at the OFL faces the font bundle ships -- read where they are, never copied -- or at
nothing at all, which is the state of every CI runner and must leave rendering exactly as
it was.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_svg
from pptx2svg.fonts import bundle_dir, office
from pptx2svg.png import available_backends, svg_to_png
from pptx2svg.text.metrics import METRICS

needs_bundle = pytest.mark.skipif(bundle_dir() is None, reason="pptx2svg-fonts is not importable")
needs_resvg = pytest.mark.skipif("resvg" not in available_backends(), reason="resvg-py is not installed")


@pytest.fixture
def no_office(monkeypatch, tmp_path):
    """A machine with no Office and no system fonts this module can see."""
    monkeypatch.setattr(office, "POWERPOINT_FONTS", tmp_path / "absent-bundle")
    monkeypatch.setattr(office, "OFFICE_CLOUD_FONTS", tmp_path / "absent-cache")
    monkeypatch.setattr(office, "system_font_dirs", lambda: ())


@pytest.fixture
def office_is_the_bundle(monkeypatch, tmp_path):
    """PowerPoint's folder is the font bundle's: Carlito, Arimo, ... "installed" there."""
    monkeypatch.setattr(office, "POWERPOINT_FONTS", bundle_dir())
    monkeypatch.setattr(office, "OFFICE_CLOUD_FONTS", tmp_path / "absent-cache")
    monkeypatch.setattr(office, "system_font_dirs", lambda: ())


TWO_FACES = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 100" width="400" height="100">'
    '<rect width="400" height="100" fill="#fff"/>'
    '<text x="0" y="0"><tspan x="10" y="40" font-size="28" font-family="Carlito, sans-serif">'
    'Handgloves</tspan><tspan x="10" y="85" font-size="28" font-family="Tinos, serif">'
    'Handgloves</tspan></text></svg>'
)


# -- Where nothing is installed, nothing changes ----------------------------------------


def test_with_no_folders_nothing_is_found(no_office):
    assert office.search_dirs() == ()
    assert office.find("Aptos") == ()
    assert office.layout_metrics(["Aptos Narrow", "Rockwell", "Gill Sans MT"]) == {}
    assert office.drawing_plan(TWO_FACES) is None


def test_cloud_folders_are_listed_only_where_they_exist(tmp_path):
    assert office.cloud_font_dirs(tmp_path / "missing") == ()
    (tmp_path / "Aptos Display").mkdir()
    (tmp_path / "stray.ttf").write_bytes(b"")
    assert office.cloud_font_dirs(tmp_path) == (tmp_path / "Aptos Display",)


def test_the_environment_switch_turns_the_lookup_off(office_is_the_bundle, monkeypatch):
    assert office.find("Carlito")
    monkeypatch.setenv("PPTX2SVG_OFFICE_FONTS", "0")
    assert office.find("Carlito") == ()
    assert office.drawing_plan(TWO_FACES) is None


@needs_bundle
@needs_resvg
def test_with_no_folders_the_png_is_todays(no_office):
    for skip in (True, False):
        assert svg_to_png(TWO_FACES, host_fonts=True, skip_system_fonts=skip) == svg_to_png(
            TWO_FACES, host_fonts=False, skip_system_fonts=skip
        )


def test_with_no_folders_the_svg_is_todays(no_office, authoring):
    assert convert_pptx_to_svg(authoring, ConvertOptions(host_fonts=True)) == convert_pptx_to_svg(
        authoring, ConvertOptions(host_fonts=False)
    )


def test_the_default_is_office_s_fonts_where_they_are_installed(monkeypatch, tmp_path):
    """Host faces take part wherever Office's fonts are installed, bundle or not -- most
    like PowerPoint -- and elsewhere exactly when the PNG is drawn from this machine's
    fonts.  ``host_fonts=False`` and ``PPTX2SVG_OFFICE_FONTS=0`` are the reproducible
    render."""
    from pptx2svg import _host_fonts, uses_host_fonts

    assert _host_fonts is uses_host_fonts

    monkeypatch.delenv("PPTX2SVG_OFFICE_FONTS", raising=False)
    assert _host_fonts(ConvertOptions(host_fonts=True)) is True
    assert _host_fonts(ConvertOptions(host_fonts=False)) is False
    # Office installed: on, with the bundle too.
    monkeypatch.setattr(office, "POWERPOINT_FONTS", tmp_path)
    monkeypatch.setattr("pptx2svg.fonts.bundle_dir", lambda: Path("/nonexistent"))
    assert office.available() and _host_fonts(ConvertOptions()) is True
    monkeypatch.setenv("PPTX2SVG_OFFICE_FONTS", "0")
    assert not office.available() and _host_fonts(ConvertOptions()) is False
    monkeypatch.delenv("PPTX2SVG_OFFICE_FONTS")
    # The cloud cache alone counts as Office's fonts.
    monkeypatch.setattr(office, "POWERPOINT_FONTS", tmp_path / "absent-bundle")
    monkeypatch.setattr(office, "OFFICE_CLOUD_FONTS", tmp_path)
    (tmp_path / "Aptos Display").mkdir()
    assert office.available()
    # No Office (CI): as before, the bundle decides.
    monkeypatch.setattr(office, "OFFICE_CLOUD_FONTS", tmp_path / "absent-cache")
    assert not office.available()
    monkeypatch.setattr("pptx2svg.fonts.bundle_dir", lambda: None)
    assert _host_fonts(ConvertOptions()) is True
    monkeypatch.setattr("pptx2svg.fonts.bundle_dir", lambda: Path("/nonexistent"))
    assert _host_fonts(ConvertOptions()) is False


# -- Finding a face ----------------------------------------------------------------------


@needs_bundle
def test_a_family_is_found_with_every_style_in_place(office_is_the_bundle):
    faces = office.find("carlito ")  # deck spellings are forgiving
    assert {(face.bold, face.italic) for face in faces} == {
        (False, False), (True, False), (False, True), (True, True)
    }
    assert {face.location for face in faces} == {"powerpoint"}
    assert all(Path(face.path).parent == bundle_dir() for face in faces)


@needs_bundle
def test_powerpoints_bundle_wins_over_the_system(monkeypatch, tmp_path):
    """The measured order: PowerPoint's own copy before macOS's (Rockwell, Symbol)."""
    linked = tmp_path / "bundle"
    linked.mkdir()
    _link(linked / "Carlito-Regular.ttf", bundle_dir() / "Carlito-Regular.ttf")
    monkeypatch.setattr(office, "POWERPOINT_FONTS", linked)
    monkeypatch.setattr(office, "OFFICE_CLOUD_FONTS", tmp_path / "absent")
    monkeypatch.setattr(office, "system_font_dirs", lambda: (bundle_dir(),))
    assert [face.location for face in office.find("Carlito")] == ["powerpoint"]
    # A family only the system has is still found there.
    assert {face.location for face in office.find("Tinos")} == {"system"}


@needs_bundle
def test_the_cloud_cache_comes_last(monkeypatch, tmp_path):
    cache = tmp_path / "CloudFonts"
    (cache / "Lato").mkdir(parents=True)
    _link(cache / "Lato" / "1234.ttf", bundle_dir() / "Lato-Regular.ttf")
    monkeypatch.setattr(office, "OFFICE_CLOUD_FONTS", cache)
    monkeypatch.setattr(office, "POWERPOINT_FONTS", tmp_path / "absent")
    monkeypatch.setattr(office, "system_font_dirs", lambda: ())
    assert [face.location for face in office.find("Lato")] == ["cloud"]
    monkeypatch.setattr(office, "system_font_dirs", lambda: (bundle_dir(),))
    assert {face.location for face in office.find("Lato")} == {"system"}


def test_an_unreadable_file_is_a_face_we_do_not_have(tmp_path):
    (tmp_path / "broken.ttf").write_bytes(b"\x00\x01\x00\x00garbage")
    (tmp_path / "empty.otf").write_bytes(b"")
    assert office._faces_in(tmp_path / "broken.ttf", "system") == []
    assert office._faces_in(tmp_path / "empty.otf", "system") == []


# -- Measuring with it -------------------------------------------------------------------


def test_own_tables_are_kept_and_the_rest_measured_from_the_face():
    assert office.has_own_table("Aptos")  # Aptos's own widths
    assert office.has_own_table("Calibri")  # Carlito: verified the same widths
    assert not office.has_own_table("Aptos Narrow")  # measured as Aptos
    assert not office.has_own_table("Aptos Light")  # reached by dropping "Light"
    assert not office.has_own_table("Rockwell")  # not in the tables at all


@needs_bundle
def test_an_installed_face_is_measured_from_its_own_file(office_is_the_bundle, monkeypatch):
    """A family the tables only guess at gets the installed file's advance widths."""
    real_find = office.find
    monkeypatch.setattr(
        office, "find", lambda family: real_find("Carlito" if family == "Corporate Sans" else family)
    )
    tables = office.layout_metrics(["Corporate Sans", "Calibri"])
    assert set(tables) == {"corporate sans"}  # Calibri keeps its verified table
    measured, carlito = tables["corporate sans"], METRICS["Carlito"]
    for char in "Hamburgefonstiv 0123":
        assert measured.widths[char] * carlito.units_per_em == carlito.widths[char] * measured.units_per_em
    assert measured.bold_widths  # the bold cut was found and read too


# -- Drawing with it ---------------------------------------------------------------------


@needs_bundle
def test_the_plan_names_the_files_for_each_stack(office_is_the_bundle):
    plan = office.drawing_plan(TWO_FACES)
    assert plan is not None and len(plan.passes) == 1
    names = {Path(path).name for path in plan.passes[0]}
    assert {"Carlito-Regular.ttf", "Tinos-Regular.ttf", "Carlito-Bold.ttf"} <= names
    assert not any(name.startswith("Arimo") for name in names)


@needs_bundle
def test_a_family_the_caller_supplies_is_left_to_the_caller(office_is_the_bundle):
    supplied = office.supplied_families([str(bundle_dir() / "Carlito-Regular.ttf")])
    assert "carlito" in supplied
    plan = office.drawing_plan(TWO_FACES, supplied=supplied)
    assert plan is not None
    assert not any("Carlito" in path for path in plan.passes[0])


def _link(link: Path, target: Path) -> None:
    """A symbolic link, not a copy: the face stays where it is installed."""
    try:
        link.symlink_to(target)
    except OSError:  # Windows without the privilege
        pytest.skip("cannot create a symbolic link here")


def _face(name: str, rasteriser: str, path: str) -> office.HostFace:
    return office.HostFace(path=path, number=0, location="powerpoint", families=frozenset({name}),
                           rasteriser_families=frozenset({rasteriser}), bold=False, italic=False,
                           css=(400, 5, "normal"))


def test_faces_filed_under_one_name_are_drawn_in_separate_passes(monkeypatch):
    """Aptos Display and Aptos are both "Aptos" to the rasteriser, at the same weight."""
    faces = {
        "aptos display": (_face("aptos display", "aptos", "/x/display.ttf"),),
        "aptos": (_face("aptos", "aptos", "/x/aptos.ttf"),),
        "aptos narrow": (_face("aptos narrow", "aptos narrow", "/x/narrow.ttf"),),
    }
    monkeypatch.setattr(office, "find", lambda family: faces.get(family.lower(), ()))
    monkeypatch.setattr(office, "_system_rasteriser_names", lambda: frozenset())
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">'
        "<text><tspan font-family=\"'Aptos Display', Aptos, Carlito, sans-serif\">T</tspan>"
        '<tspan font-family="Aptos, Carlito, sans-serif">b</tspan>'
        "<tspan font-family=\"'Aptos Narrow', Aptos, sans-serif\">n</tspan></text></svg>"
    )
    plan = office.drawing_plan(svg)
    assert plan.passes == [["/x/display.ttf", "/x/narrow.ttf"], ["/x/aptos.ttf"]]
    assert sorted(plan.stacks.values()) == [0, 0, 1]
    assert not plan.ahead_of_system


@needs_bundle
@needs_resvg
def test_drawing_in_two_passes_is_drawing_in_one(office_is_the_bundle):
    """The composite is the single render, pixel for pixel, clip paths included."""
    svg = TWO_FACES.replace(
        '<rect width="400"',
        '<defs><clipPath id="c"><rect width="380" height="100"/></clipPath></defs>'
        '<rect width="400"',
    ).replace("<text ", '<text clip-path="url(#c)" ')
    plan = office.drawing_plan(svg)
    files = plan.passes[0]
    carlito = [path for path in files if "Carlito" in path]
    tinos = [path for path in files if "Tinos" in path]
    one = office.DrawingPlan(passes=[carlito + tinos], stacks=plan.stacks, ahead_of_system=False)
    two = office.DrawingPlan(
        passes=[carlito, tinos],
        stacks={value: (1 if "Tinos" in value else 0) for value in plan.stacks},
        ahead_of_system=False,
    )
    from pptx2svg.png import _render_plan

    def draw(chosen, width=None):
        return _render_plan(svg, chosen, width=width, height=None, scale=None, background=None,
                            font_dirs=[], font_files=[], skip_system_fonts=True,
                            generic_families=None)

    from PIL import Image

    for width in (None, 800):
        single = Image.open(io.BytesIO(draw(one, width))).convert("RGBA")
        double = Image.open(io.BytesIO(draw(two, width))).convert("RGBA")
        assert single.size == double.size
        assert single.tobytes() == double.tobytes()


def test_a_later_pass_hides_everything_but_its_own_text():
    plan = office.DrawingPlan(passes=[["a"], ["b"]], stacks={"A, serif": 0, "B, serif": 1},
                              ahead_of_system=False)
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 20" width="10" height="20">'
           '<clipPath id="c"/><text><tspan font-family="A, serif">a</tspan>'
           '<tspan font-family="B, serif">b</tspan></text></svg>')
    first = office.pass_svg(svg, plan, 0, None)
    assert 'font-family="A, serif">' in first
    assert 'font-family="B, serif" visibility="hidden">' in first
    later = office.pass_svg(svg, plan, 1, b"\x89PNG")
    assert later.startswith('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 20" '
                            'width="10" height="20" visibility="hidden"><image x="0" y="0" '
                            'width="10" height="20"')
    assert 'font-family="A, serif" visibility="hidden">' in later
    assert 'font-family="B, serif" visibility="visible">' in later
    assert '<clipPath visibility="visible" id="c"/>' in later


def test_stack_names_unquote_and_unescape():
    assert office.stack_names("'Aptos Display', Aptos, &quot;Noto Sans JP&quot;, sans-serif") == [
        "Aptos Display", "Aptos", "Noto Sans JP", "sans-serif"
    ]


# -- A pass leaves out the texts it draws nothing of --------------------------------------

SEPARATE_TEXTS = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 200" width="400" height="200">'
    '<defs><filter id="f"><feGaussianBlur in="SourceAlpha" stdDeviation="2"/></filter>'
    '<clipPath id="c"><text x="0" y="190"><tspan font-family="Tinos, serif" font-size="40">'
    'clip</tspan></text></clipPath></defs>'
    '<rect width="400" height="200" fill="#fff"/>'
    '<text x="10" y="40"><tspan font-size="28" font-family="Carlito, sans-serif">Handgloves</tspan></text>'
    '<g transform="translate(0, 10)"><text x="10" y="80"><tspan font-size="28" '
    'font-family="Tinos, serif">Handgloves</tspan></text></g>'
    '<text x="200" y="40"><tspan font-size="20" font-family="Carlito, sans-serif">mixed </tspan>'
    '<tspan font-size="20" font-family="Tinos, serif">runs</tspan></text>'
    '<g filter="url(#f)"><text x="200" y="120"><tspan font-size="20" font-family="Tinos, serif">'
    'shadowed</tspan></text></g>'
    '<rect y="150" width="400" height="50" fill="#36c" clip-path="url(#c)"/>'
    '</svg>'
)


def test_a_pass_leaves_out_only_the_texts_it_draws_nothing_of():
    plan = office.DrawingPlan(passes=[["c"], ["t"]],
                              stacks={"Carlito, sans-serif": 0, "Tinos, serif": 1},
                              ahead_of_system=False)
    first = office.pass_svg(SEPARATE_TEXTS, plan, 0, None)
    # The Tinos-only text in the open goes; the mixed one, the filtered one (a filter may
    # paint its whole region) and the one in a clip path (drawn where it is used) stay.
    assert first.count("<text") == SEPARATE_TEXTS.count("<text") - 1
    assert 'y="80"' not in first
    later = office.pass_svg(SEPARATE_TEXTS, plan, 1, b"\x89PNG")
    assert later.count("<text") == SEPARATE_TEXTS.count("<text") - 1
    assert 'x="10" y="40"' not in later


@pytest.mark.parametrize("svg", [
    '<svg xmlns="http://www.w3.org/2000/svg"><!-- note --><text visibility="hidden">a</text></svg>',
    '<svg xmlns="http://www.w3.org/2000/svg"><text id="t" visibility="hidden">a</text><use href="#t"/></svg>',
    '<svg xmlns="http://www.w3.org/2000/svg"><text visibility="hidden">a</svg>',
])
def test_what_the_scanner_does_not_follow_is_left_as_it_is(svg):
    assert office.without_undrawn_text(svg) == svg


def test_a_text_whose_characters_are_all_hidden_is_left_out():
    svg = ('<svg xmlns="http://www.w3.org/2000/svg"><text visibility="hidden">a<tspan>b</tspan></text>'
           '<text visibility="hidden"><tspan visibility="visible">c</tspan></text>'
           '<text><tspan visibility="hidden">d</tspan></text><text/></svg>')
    assert office.without_undrawn_text(svg) == (
        '<svg xmlns="http://www.w3.org/2000/svg">'
        '<text visibility="hidden"><tspan visibility="visible">c</tspan></text><text/></svg>')


@needs_bundle
@needs_resvg
def test_leaving_undrawn_texts_out_draws_the_same_pixels(monkeypatch):
    files = sorted(str(path) for path in Path(bundle_dir()).iterdir())
    carlito = [path for path in files if "Carlito" in path]
    tinos = [path for path in files if "Tinos" in path]
    stacks = {"Carlito, sans-serif": 0, "Tinos, serif": 1}
    two = office.DrawingPlan(passes=[carlito, tinos], stacks=stacks, ahead_of_system=False)
    from pptx2svg.png import _render_plan

    def draw(chosen):
        return _render_plan(SEPARATE_TEXTS, chosen, width=800, height=None, scale=None,
                            background=None, font_dirs=[], font_files=[],
                            skip_system_fonts=True, generic_families=None)

    from PIL import Image

    def pixels(data):
        return Image.open(io.BytesIO(data)).convert("RGBA").tobytes()

    dropped = draw(two)
    from ooxml_common import raster

    monkeypatch.setattr(raster, "without_undrawn_text", lambda svg: svg)
    assert pixels(dropped) == pixels(draw(two))


# -- The caller's folder holding two faces filed under one name ---------------------------


def _named(source: Path, family: str, typographic: str) -> bytes:
    """An OFL face of the bundle under made-up names: name ID 1 ``family``, ID 16
    ``typographic`` -- how Aptos Display and Aptos are both "Aptos" to the rasteriser."""
    import struct

    from ooxml_common.fonts.sfnt import read_tables, write_sfnt

    records = [(1, family), (2, "Regular"), (4, f"{family} Regular"), (6, family.replace(" ", "")),
               (16, typographic), (17, "Regular")]
    strings = b""
    entries = b""
    for name_id, text in records:
        raw = text.encode("utf-16-be")
        entries += struct.pack(">HHHHHH", 3, 1, 0x409, name_id, len(raw), len(strings))
        strings += raw
    name = struct.pack(">HHH", 0, len(records), 6 + 12 * len(records)) + entries + strings
    tables = read_tables(source.read_bytes())
    tables[b"name"] = name
    return write_sfnt(list(tables.items()))


@pytest.fixture
def one_folder(tmp_path):
    """A caller's folder with "Body Face" and "Body Face Display", both filed under "Body
    Face" at the same weight, width and style (Carlito's and Tinos' outlines)."""
    folder = tmp_path / "fonts"
    folder.mkdir()
    (folder / "a-display.ttf").write_bytes(_named(bundle_dir() / "Tinos-Regular.ttf", "Body Face Display",
                                                  "Body Face"))
    (folder / "b-body.ttf").write_bytes(_named(bundle_dir() / "Carlito-Regular.ttf", "Body Face", "Body Face"))
    return folder


TITLE = ("<tspan x=\"10\" y=\"40\" font-size=\"28\" font-family=\"'Body Face Display', 'Body Face', "
         "sans-serif\">Handgloves</tspan>")
BODY = ("<tspan x=\"10\" y=\"85\" font-size=\"28\" font-family=\"'Body Face', sans-serif\">"
        "Handgloves</tspan>")


def _page(*spans: str) -> str:
    return ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 100" width="400" height="100">'
            f'<rect width="400" height="100" fill="#fff"/><text>{"".join(spans)}</text></svg>')


@needs_bundle
def test_two_faces_of_one_name_in_the_callers_folder_are_planned_apart(one_folder, no_office):
    own = office.supplied_faces((), [str(one_folder)])
    plan = office.drawing_plan(_page(TITLE, BODY), own=own, host=False)
    assert plan is not None and len(plan.passes) == 2
    display, body = str(one_folder / "a-display.ttf"), str(one_folder / "b-body.ttf")
    assert plan.passes == [[display], [body]]
    assert plan.withheld == [[body], [display]]
    # One family alone, its namesake unused: still drawn without the namesake.
    alone = office.drawing_plan(_page(TITLE), own=own, host=False)
    assert alone.passes == [[display]] and alone.withheld == [[body]]
    # Faces that share no name need no plan: the folder is handed over as it is.
    assert office.drawing_plan(TWO_FACES, own=office.supplied_faces((), [str(bundle_dir())]), host=False) is None


@needs_bundle
@needs_resvg
def test_each_face_of_the_callers_folder_draws_its_own_text(one_folder, no_office, tmp_path):
    """The title in the display face and the body in the body face -- each as it draws
    with its own face alone in the folder -- not both in whichever the rasteriser met
    first."""
    from PIL import Image

    def draw(svg: str, folder: Path) -> "Image.Image":
        return Image.open(io.BytesIO(svg_to_png(svg, font_dirs=[str(folder)], host_fonts=False))).convert("RGB")

    def only(name: str) -> Path:
        folder = tmp_path / name
        folder.mkdir()
        (folder / name).write_bytes((one_folder / name).read_bytes())
        return folder

    both = draw(_page(TITLE, BODY), one_folder)
    title = draw(_page(TITLE), only("a-display.ttf"))
    body = draw(_page(BODY), only("b-body.ttf"))
    top, bottom = (0, 0, 400, 50), (0, 50, 400, 100)
    assert both.crop(top).tobytes() == title.crop(top).tobytes()
    assert both.crop(bottom).tobytes() == body.crop(bottom).tobytes()
    assert title.crop(top).tobytes() != draw(_page(TITLE.replace("Display', ", "Xx', ")), one_folder).crop(top).tobytes()


def test_an_ascii_slide_is_not_scanned_for_emoji(monkeypatch):
    """A slide's pictures are base64: an SVG with no character past ASCII has no emoji, and
    is not walked character by character (megabytes, every slide)."""
    def unexpected():
        raise AssertionError("the emoji face was looked up")

    monkeypatch.setattr(office, "emoji_face", unexpected)
    svg = TWO_FACES.replace("</svg>", '<image href="data:image/png;base64,iVBORw0KGgo="/></svg>')
    assert office.with_emoji(svg) == (svg, [])
