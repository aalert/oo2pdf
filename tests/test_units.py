"""Unit tests for self-contained helpers (no documents, no Office needed)."""
import pytest

from oo2pdf.docx.layout import _numeric_picture
from oo2pdf.docx.math import math_italic
from oo2pdf.docx.props import Border
from oo2pdf.docx.reader import _cond_edges
from oo2pdf.docx.vml import css, length, parse_path
from oo2pdf.fonts import _to_unicode_cmap, _utf16_hex


# ---------------------------------------------------------------- fields
@pytest.mark.parametrize("value, picture, expected", [
    (6, "0#", "06"),        # PAGE \# 0# -> "06" (RedAndBlackReport)
    (12, "0#", "12"),
    (5, "00", "05"),
    (5, "#,##0", "5"),      # grouping comma without digits before it disappears
    (1234, "#,##0", "1,234"),
    (7, "Page 0", "Page 7"),
])
def test_numeric_picture(value, picture, expected):
    assert _numeric_picture(value, picture) == expected


# ---------------------------------------------------------------- equations
def test_math_italic_latin_and_greek():
    # Word italicises Latin and lowercase Greek letters in equations
    assert math_italic("x") == "\U0001d465"
    assert math_italic("α") == "\U0001d6fc"      # alpha -> mathematical italic alpha
    assert math_italic("Α") == "Α"          # capital Greek stays upright
    assert math_italic("h") == "ℎ"               # italic h lives in Letterlike Symbols


def test_math_bold_styles():
    assert math_italic("a", "b") == "\U0001d41a"
    assert math_italic("a", "bi") == "\U0001d482"


# ---------------------------------------------------------------- PDF text layer
def test_utf16_surrogates_for_astral_characters():
    assert _utf16_hex(0x41) == "0041"
    # U+1D44D (mathematical italic Z) must be a UTF-16 surrogate pair
    assert _utf16_hex(0x1D44D) == "D835DC4D"


def test_to_unicode_cmap_uses_surrogates():
    cmap = _to_unicode_cmap("F1", [0x41, 0x1D44D])
    assert "<00> <0041>" in cmap
    assert "<01> <D835DC4D>" in cmap


# ---------------------------------------------------------------- VML
def test_vml_path_empty_values_are_zero():
    # "m,1008" is 0,1008 and a trailing "2160," ends with 0
    pts, closed = parse_path("m,1008l2160,e")
    assert pts == [(0.0, 1008.0), (2160.0, 0.0)]
    assert not closed


def test_vml_path_curves_are_flattened_and_closed():
    pts, closed = parse_path("m0,0c0,100,100,100,100,0xe")
    assert closed
    assert pts[0] == (0.0, 0.0) and pts[-1] == (100.0, 0.0)
    assert len(pts) > 3                                  # the curve became several segments
    assert max(y for _, y in pts) > 50                   # ... that follow the bulge


def test_vml_css_and_lengths():
    st = css("position:absolute; width:72pt;HEIGHT:1in")
    assert st["position"] == "absolute"
    assert length(st["width"]) == pytest.approx(72)
    assert length(st["height"]) == pytest.approx(72)
    assert length("2.54cm") == pytest.approx(72, abs=0.01)


# ---------------------------------------------------------------- table styles
def _b(w=1.0):
    return Border("single", w, None, 0)


def test_row_condition_left_right_are_outer_edges():
    # a header row style: left/right borders only on the row's outer edges,
    # insideV (here: none) between its cells
    tcpr = {"tcBorders": {"left": _b(), "right": _b(), "insideV": Border("nil", 0, None, 0)}}
    first = _cond_edges("firstRow", tcpr, 0, 5, 0, 1, 3)["tcBorders"]
    middle = _cond_edges("firstRow", tcpr, 0, 5, 1, 1, 3)["tcBorders"]
    last = _cond_edges("firstRow", tcpr, 0, 5, 2, 1, 3)["tcBorders"]
    assert first["left"].visible and not first["right"].visible
    assert not middle["left"].visible and not middle["right"].visible
    assert not last["left"].visible and last["right"].visible


def test_column_condition_top_bottom_are_outer_edges():
    tcpr = {"tcBorders": {"top": _b(), "bottom": _b(), "insideH": _b(0.5)}}
    inner = _cond_edges("firstCol", tcpr, 2, 5, 0, 1, 3)["tcBorders"]
    assert inner["top"].width == 0.5 and inner["bottom"].width == 0.5


def test_condition_without_inside_border_leaves_table_border():
    tcpr = {"tcBorders": {"left": _b(2.0)}}
    out = _cond_edges("band1Horz", tcpr, 1, 5, 1, 1, 3)["tcBorders"]
    assert "left" not in out          # the table's insideV applies instead
