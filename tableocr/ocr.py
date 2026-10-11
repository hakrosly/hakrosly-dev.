"""セル単位で tesseract を呼び出して文字を読み取る。"""

from __future__ import annotations

import os
import re
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass

import cv2
import numpy as np
import pytesseract

from .table_detect import Cell, LineMasks, Table

_CJK = r"　-ヿ㐀-䶿一-鿿豈-﫿＀-￯"
_SPACE_BETWEEN_CJK = re.compile(rf"(?<=[{_CJK}])\s+(?=[{_CJK}])")
_SPACE_IN_NUMBER = re.compile(r"(?<=\d)\s*([,.])\s*(?=\d)")


@dataclass
class OcrOptions:
    lang: str = "jpn+eng"
    keep_newlines: bool = False
    jobs: int = 1
    extra_config: str = ""


def configure_tesseract(cmd: str | None) -> None:
    cmd = cmd or os.environ.get("TESSERACT_CMD")
    if cmd:
        pytesseract.pytesseract.tesseract_cmd = cmd


def clean_text(text: str, keep_newlines: bool = False) -> str:
    lines = []
    for line in text.splitlines():
        line = _SPACE_BETWEEN_CJK.sub("", line)
        line = _SPACE_IN_NUMBER.sub(r"\1", line)
        line = re.sub(r"[ \t]+", " ", line).strip()
        if line:
            lines.append(line)
    return ("\n" if keep_newlines else " ").join(lines)


def _remove_specks(ink: np.ndarray, dpi: int) -> np.ndarray:
    """小さな点ゴミと、セルの端に残った罫線のかけらを消す。"""
    n, labels, stats, _ = cv2.connectedComponentsWithStats(ink, connectivity=8)
    if n <= 1:
        return ink
    h, w = ink.shape
    min_area = max(6, int((dpi / 300) ** 2 * 20))
    sliver = max(3, dpi // 60)
    keep = np.zeros(n, dtype=bool)
    for i in range(1, n):
        x, y, cw, ch, area = stats[i]
        if area < min_area:
            continue
        on_edge = x == 0 or y == 0 or x + cw >= w or y + ch >= h
        if on_edge and (cw <= sliver or ch <= sliver):
            continue
        keep[i] = True
    return np.where(keep[labels], 255, 0).astype(np.uint8)


def _clean_ink(gray: np.ndarray, masks: LineMasks, y0: int, y1: int, x0: int, x1: int, dpi: int) -> np.ndarray:
    """範囲内の文字だけを取り出した 2 値画像 (文字 = 255) を返す。罫線とゴミは消す。"""
    ink = masks.binary[y0:y1, x0:x1].copy()
    grow = max(3, dpi // 100 * 2 + 1)
    lines = cv2.bitwise_or(masks.horizontal[y0:y1, x0:x1], masks.vertical[y0:y1, x0:x1])
    lines = cv2.dilate(lines, np.ones((grow, grow), np.uint8))
    ink[lines > 0] = 0
    return _remove_specks(ink, dpi)


def _text_bands(ink: np.ndarray, dpi: int) -> list[tuple[int, int]]:
    """文字の行 (帯) を上から順に (開始 y, 終了 y) で返す。字の中の小さな隙間は無視する。"""
    profile = (np.count_nonzero(ink, axis=1) > 0).astype(np.uint8).reshape(-1, 1)
    gap = max(3, dpi // 30)
    profile = cv2.morphologyEx(profile, cv2.MORPH_CLOSE, np.ones((gap, 1), np.uint8)).ravel()
    edges = np.diff(np.concatenate(([0], profile.astype(np.int8), [0])))
    return list(zip(np.flatnonzero(edges == 1), np.flatnonzero(edges == -1)))


def _to_ocr_image(gray: np.ndarray, ink: np.ndarray, bands: int) -> np.ndarray | None:
    """文字の部分を切り出し、tesseract が読みやすい白地黒文字の画像にする。"""
    if cv2.countNonZero(ink) == 0:
        return None
    ys, xs = np.nonzero(ink)
    by0, by1, bx0, bx1 = ys.min(), ys.max() + 1, xs.min(), xs.max() + 1
    ink = ink[by0:by1, bx0:bx1]

    # 文字の部分だけ元の濃淡を使い、それ以外は白にする
    keep = cv2.dilate(ink, np.ones((3, 3), np.uint8))
    crop = gray[by0:by1, bx0:bx1].copy()
    crop[keep == 0] = 255

    # 1 行の高さが 50px 程度のとき一番よく読めるので、それに合わせて拡大・縮小する
    line_h = crop.shape[0] / max(1, bands)
    scale = min(4.0, max(0.5, 50 / line_h))
    if abs(scale - 1.0) > 0.1:
        interp = cv2.INTER_CUBIC if scale > 1.0 else cv2.INTER_AREA
        crop = cv2.resize(crop, None, fx=scale, fy=scale, interpolation=interp)

    _, crop = cv2.threshold(crop, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    pad = 20
    return cv2.copyMakeBorder(crop, pad, pad, pad, pad, cv2.BORDER_CONSTANT, value=255)


def _prepare_cell(gray: np.ndarray, masks: LineMasks, cell: Cell, dpi: int) -> tuple[np.ndarray, int] | None:
    """セル画像から罫線を消し、文字部分だけを白地黒文字で切り出す。

    文字が無ければ None。戻り値の int は文字の行数の目安。
    """
    inset = max(2, dpi // 150)
    y0, y1 = cell.y0 + inset, cell.y1 - inset
    x0, x1 = cell.x0 + inset, cell.x1 - inset
    if y1 - y0 < 4 or x1 - x0 < 4:
        return None
    ink = _clean_ink(gray, masks, y0, y1, x0, x1, dpi)
    bands = max(1, len(_text_bands(ink, dpi)))
    image = _to_ocr_image(gray[y0:y1, x0:x1], ink, bands)
    return None if image is None else (image, bands)


def _read_image(image: np.ndarray, bands: int, options: OcrOptions) -> str:
    # 1 行なら行モード、複数行ならブロックモード。空振りしたら別モードで読み直す
    modes = (7, 13) if bands == 1 else (6, 11)
    text = ""
    for psm in modes:
        config = f"--psm {psm} --dpi 300 {options.extra_config}".strip()
        text = clean_text(pytesseract.image_to_string(image, lang=options.lang, config=config), options.keep_newlines)
        if text:
            break
    return text


def ocr_cell(gray: np.ndarray, masks: LineMasks, cell: Cell, dpi: int, options: OcrOptions) -> str:
    prepared = _prepare_cell(gray, masks, cell, dpi)
    if prepared is None:
        return ""
    image, bands = prepared
    return _read_image(image, bands, options)


def ocr_table(gray: np.ndarray, masks: LineMasks, table: Table, dpi: int, options: OcrOptions) -> list[list[str]]:
    """表の各セルを読み取り、行×列の文字列リストを返す。

    結合セルは左上のマスに文字を入れ、残りは空文字にする。
    """
    grid = [["" for _ in range(table.n_cols)] for _ in range(table.n_rows)]
    if options.jobs > 1:
        # 並列で動かすときは tesseract 内部のスレッドを 1 本に絞る
        os.environ.setdefault("OMP_THREAD_LIMIT", "1")
        with ThreadPoolExecutor(max_workers=options.jobs) as pool:
            texts = list(pool.map(lambda c: ocr_cell(gray, masks, c, dpi, options), table.cells))
    else:
        texts = [ocr_cell(gray, masks, c, dpi, options) for c in table.cells]
    for cell, text in zip(table.cells, texts):
        grid[cell.row][cell.col] = text
    return grid


def ocr_region_text(gray: np.ndarray, masks: LineMasks, table: Table, dpi: int, options: OcrOptions) -> str:
    """表と、そのすぐ上 (表題が書かれがちな所) をまとめて読む。

    キーワードで表を探すときに使う。罫線は消してから読む。
    """
    y0 = max(0, table.y - int(dpi * 0.6))
    y1, x0, x1 = table.y + table.h, table.x, table.x + table.w
    crop = gray[y0:y1, x0:x1].copy()
    lines = cv2.bitwise_or(masks.horizontal[y0:y1, x0:x1], masks.vertical[y0:y1, x0:x1])
    grow = max(3, dpi // 100 * 2 + 1)
    crop[cv2.dilate(lines, np.ones((grow, grow), np.uint8)) > 0] = 255
    text = pytesseract.image_to_string(crop, lang=options.lang, config="--psm 11")
    return re.sub(r"\s+", "", text)


def read_title(gray: np.ndarray, masks: LineMasks, table: Table, dpi: int, options: OcrOptions) -> str:
    """表の欄外左上 (表のすぐ上、左寄り) に書かれた表題を読む。

    表の上 0.8 インチ以内・表の左半分の範囲で、表に一番近い 1 行だけを読む。
    その範囲に別の表の罫線があれば、それより下だけを見る。
    """
    y1 = table.y - max(2, dpi // 100)
    y0 = max(0, table.y - int(dpi * 0.8))
    x0 = max(0, table.x - int(dpi * 0.3))
    x1 = min(gray.shape[1], table.x + max(table.w // 2, int(dpi * 1.5)))
    if y1 - y0 < 5:
        return ""
    above = np.flatnonzero(np.count_nonzero(masks.horizontal[y0:y1, x0:x1], axis=1))
    if above.size:
        y0 += int(above[-1]) + max(2, dpi // 100)
        if y1 - y0 < 5:
            return ""

    ink = _clean_ink(gray, masks, y0, y1, x0, x1, dpi)
    bands = _text_bands(ink, dpi)
    if not bands:
        return ""
    top, bottom = bands[-1]
    line_ink = np.zeros_like(ink)
    line_ink[top:bottom] = ink[top:bottom]
    image = _to_ocr_image(gray[y0:y1, x0:x1], line_ink, 1)
    if image is None:
        return ""
    return _read_image(image, 1, OcrOptions(lang=options.lang))


def normalize_title(text: str) -> str:
    """全角・半角の違いと空白を無くして比べやすくする。"""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text))


def title_matches(title: str, wanted: str) -> bool:
    return bool(wanted) and normalize_title(wanted) in normalize_title(title)
