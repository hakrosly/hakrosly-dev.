"""テスト用に「スキャンした帳票」風の TIFF を作る。"""

from __future__ import annotations

import os

import numpy as np
from PIL import Image, ImageDraw, ImageFont

FONT_CANDIDATES = [
    "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
    "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
    "C:/Windows/Fonts/msgothic.ttc",
    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
]

# 表 2 の中身 (1 行目の「売上」は 2 列分の結合セル、「備考」は 2 行分の結合セル)
TABLE2_HEADER = ["支店", "売上", "", "備考"]
TABLE2_SUB = ["", "4月", "5月", ""]
TABLE2_BODY = [
    ["東京", "1,200", "1,350", "好調"],
    ["大阪", "980", "1,020", ""],
    ["名古屋", "760", "805", "新店舗"],
]
TABLE1 = [
    ["No", "Item", "Qty"],
    ["1", "Apple", "10"],
    ["2", "Orange", "25"],
]


def find_font() -> str | None:
    for path in FONT_CANDIDATES:
        if os.path.exists(path):
            return path
    return None


def _grid(draw, x0, y0, col_w, row_h, n_rows, width=3):
    total_w = sum(col_w)
    for r in range(n_rows + 1):
        y = y0 + r * row_h
        draw.line([(x0, y), (x0 + total_w, y)], fill=0, width=width)
    x = x0
    for w in col_w + [0]:
        draw.line([(x, y0), (x, y0 + n_rows * row_h)], fill=0, width=width)
        x += w


def _text(draw, font, x, y, w, h, text):
    if not text:
        return
    left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
    draw.text((x + (w - (right - left)) / 2 - left, y + (h - (bottom - top)) / 2 - top), text, fill=0, font=font)


def make_sample(path: str, rotate: float = 0.6, noise: float = 0.002, seed: int = 0) -> None:
    font_path = find_font()
    if font_path is None:
        raise RuntimeError("日本語フォントが見つかりません")
    font = ImageFont.truetype(font_path, 34)
    title_font = ImageFont.truetype(font_path, 48)

    page = Image.new("L", (1800, 2000), 255)
    draw = ImageDraw.Draw(page)
    draw.text((150, 100), "月次報告書 (サンプル)", fill=0, font=title_font)
    draw.text((150, 190), "下の表は動作確認用のダミーです。", fill=0, font=font)

    # 表 1: 単純な 3x3 の表
    col_w, row_h = [150, 400, 200], 80
    x0, y0 = 150, 300
    _grid(draw, x0, y0, col_w, row_h, len(TABLE1))
    for r, row in enumerate(TABLE1):
        x = x0
        for c, value in enumerate(row):
            _text(draw, font, x, y0 + r * row_h, col_w[c], row_h, value)
            x += col_w[c]

    draw.text((150, 650), "支店別売上 (単位: 千円)", fill=0, font=font)

    # 表 2: 結合セルを含む表
    col_w, row_h = [300, 280, 280, 400], 90
    x0, y0 = 150, 720
    n_rows = 2 + len(TABLE2_BODY)
    xs = [x0]
    for w in col_w:
        xs.append(xs[-1] + w)
    ys = [y0 + r * row_h for r in range(n_rows + 1)]
    lw = 3
    # 外枠と通常の罫線
    for y in ys:
        draw.line([(xs[0], y), (xs[-1], y)], fill=0, width=lw)
    for x in xs:
        draw.line([(x, ys[0]), (x, ys[-1])], fill=0, width=lw)
    # 結合セル: 「売上」の下の区切り以外、1 行目と 2 行目の間の線を部分的に消す
    draw.line([(xs[0] + 3, ys[1]), (xs[1] - 3, ys[1])], fill=255, width=lw + 2)
    draw.line([(xs[3] + 3, ys[1]), (xs[4] - 3, ys[1])], fill=255, width=lw + 2)
    draw.line([(xs[2], ys[0] + 3), (xs[2], ys[1] - 3)], fill=255, width=lw + 2)
    _text(draw, font, xs[0], ys[0], col_w[0], row_h * 2, "支店")
    _text(draw, font, xs[1], ys[0], col_w[1] + col_w[2], row_h, "売上")
    _text(draw, font, xs[3], ys[0], col_w[3], row_h * 2, "備考")
    _text(draw, font, xs[1], ys[1], col_w[1], row_h, "4月")
    _text(draw, font, xs[2], ys[1], col_w[2], row_h, "5月")
    for r, row in enumerate(TABLE2_BODY):
        for c, value in enumerate(row):
            _text(draw, font, xs[c], ys[r + 2], col_w[c], row_h, value)

    draw.text((150, 1250), "以上", fill=0, font=font)

    if rotate:
        page = page.rotate(rotate, resample=Image.BICUBIC, expand=False, fillcolor=255)

    arr = np.array(page)
    rng = np.random.default_rng(seed)
    arr[rng.random(arr.shape) < noise] = 0
    page = Image.fromarray(arr).point(lambda v: 255 if v > 128 else 0).convert("1")

    # 2 ページ目は表なし (マルチページの確認用)
    blank = Image.new("1", page.size, 1)
    ImageDraw.Draw(blank).text((150, 150), "2ページ目", fill=0, font=font)
    page.save(path, compression="group4", dpi=(300, 300), save_all=True, append_images=[blank])


if __name__ == "__main__":
    import sys

    out = sys.argv[1] if len(sys.argv) > 1 else "sample.tif"
    make_sample(out)
    print(f"作成しました: {out}")
