import shutil

import pytest

from sample_factory import TABLE1, TABLE2_BODY, find_font, make_sample
from tableocr.cli import main
from tableocr.export import to_delimited, write_xlsx
from tableocr.image_io import load_pages
from tableocr.ocr import OcrOptions, clean_text, ocr_table
from tableocr.table_detect import deskew, detect_tables

needs_font = pytest.mark.skipif(find_font() is None, reason="日本語フォントがありません")
needs_tesseract = pytest.mark.skipif(shutil.which("tesseract") is None, reason="tesseract がありません")

EXPECTED_TABLE2 = [
    ["支店", "売上", "", "備考"],
    ["", "4月", "5月", ""],
    *TABLE2_BODY,
]


@pytest.fixture(scope="module")
def sample(tmp_path_factory):
    path = tmp_path_factory.mktemp("img") / "sample.tif"
    make_sample(str(path))
    return str(path)


@pytest.fixture(scope="module")
def page_tables(sample):
    page = load_pages(sample)[0]
    gray, angle = deskew(page.gray)
    tables, masks = detect_tables(gray, page.dpi, page=page.index)
    return page, gray, angle, tables, masks


@needs_font
def test_multipage_load(sample):
    pages = load_pages(sample)
    assert len(pages) == 2
    assert pages[0].dpi == 300


@needs_font
def test_deskew_detects_rotation(page_tables):
    _, _, angle, _, _ = page_tables
    assert angle == pytest.approx(-0.6, abs=0.15)


@needs_font
def test_detects_both_tables(page_tables):
    _, _, _, tables, _ = page_tables
    assert [(t.n_rows, t.n_cols) for t in tables] == [(3, 3), (5, 4)]


@needs_font
def test_merged_cells(page_tables):
    _, _, _, tables, _ = page_tables
    spans = {(c.row, c.col): (c.row_span, c.col_span) for c in tables[1].cells}
    assert spans[(0, 0)] == (2, 1)  # 支店
    assert spans[(0, 1)] == (1, 2)  # 売上
    assert spans[(0, 3)] == (2, 1)  # 備考
    assert spans[(1, 1)] == (1, 1)
    assert len(tables[1].cells) == 4 * 5 - 3


@needs_font
def test_roi_limits_search(sample):
    page = load_pages(sample)[0]
    tables, _ = detect_tables(page.gray, page.dpi, roi=(100, 650, 1500, 700))
    assert len(tables) == 1
    assert tables[0].n_cols == 4


@needs_font
@needs_tesseract
def test_ocr_tables(page_tables):
    page, gray, _, tables, masks = page_tables
    options = OcrOptions(jobs=4)
    assert ocr_table(gray, masks, tables[0], page.dpi, options) == TABLE1
    assert ocr_table(gray, masks, tables[1], page.dpi, options) == EXPECTED_TABLE2


@needs_font
@needs_tesseract
def test_cli_keyword_and_xlsx(sample, tmp_path, capsys):
    out = tmp_path / "out.xlsx"
    assert main([sample, "--keyword", "支店別売上", "-o", str(out)]) == 0
    from openpyxl import load_workbook

    ws = load_workbook(out).active
    assert ws["A3"].value == "東京"
    assert {str(r) for r in ws.merged_cells.ranges} == {"A1:A2", "B1:C1", "D1:D2"}


@needs_font
def test_cli_list(sample, capsys):
    assert main([sample, "--list"]) == 0
    lines = capsys.readouterr().out.strip().splitlines()
    assert len(lines) == 2
    assert "5 行 x 4 列" in lines[1]


@needs_font
def test_cli_bad_table_number(sample):
    assert main([sample, "--table", "9"]) == 2


def test_clean_text():
    assert clean_text("名 古 屋\n") == "名古屋"
    assert clean_text("1, 200") == "1,200"
    assert clean_text("3 . 5 kg") == "3.5 kg"
    assert clean_text("Apple  Pie") == "Apple Pie"
    assert clean_text("一行目\n二行目") == "一行目 二行目"
    assert clean_text("一行目\n二行目", keep_newlines=True) == "一行目\n二行目"


def test_to_delimited_quotes_newlines():
    text = to_delimited([["a", "b\nc"], ["", "d"]])
    assert text == 'a\t"b\nc"\r\n\td\r\n'


def test_write_xlsx_keeps_text(tmp_path):
    path = tmp_path / "t.xlsx"
    write_xlsx(str(path), [["001", "x"]])
    from openpyxl import load_workbook

    assert load_workbook(path).active["A1"].value == "001"
