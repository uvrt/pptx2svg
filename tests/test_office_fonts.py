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


def test_the_default_follows_the_bundle(monkeypatch):
    """Host faces take part exactly when the PNG is drawn from this machine's fonts."""
    from pptx2svg import _host_fonts

    assert _host_fonts(ConvertOptions(host_fonts=True)) is True
    assert _host_fonts(ConvertOptions(host_fonts=False)) is False
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
