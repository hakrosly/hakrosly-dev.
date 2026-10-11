"""コマンドライン入口。

例:
    python -m tableocr scan.tif --list --debug-dir debug
    python -m tableocr scan.tif --page 1 --table 2 -o result.xlsx
    python -m tableocr scan.tif --keyword 請求明細 --clipboard
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys

import cv2
import pytesseract

from . import __version__
from .export import copy_to_clipboard, to_delimited, write_output
from .image_io import load_pages
from .ocr import OcrOptions, configure_tesseract, ocr_region_text, ocr_table
from .table_detect import deskew, detect_tables, draw_tables


def _parse_roi(text: str) -> tuple[int, int, int, int]:
    try:
        x, y, w, h = (int(v) for v in text.split(","))
    except ValueError as exc:
        raise argparse.ArgumentTypeError("--roi は x,y,幅,高さ の形で指定してください (例: 100,200,1500,800)") from exc
    if w <= 0 or h <= 0:
        raise argparse.ArgumentTypeError("--roi の幅と高さは 1 以上にしてください")
    return x, y, w, h


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="tableocr",
        description="TIFF 画像の中の罫線付きの表を読み取り、Excel に貼れる形 (TSV / CSV / xlsx) で出力します。",
    )
    p.add_argument("input", help="TIFF ファイル (マルチページ可)")
    p.add_argument("--list", action="store_true", help="見つかった表の一覧を表示して終了する")
    p.add_argument("--page", type=int, help="対象のページ番号 (1 始まり)。省略時は全ページ")
    p.add_argument("--table", type=int, default=1, help="対象の表の番号 (--list で表示される番号)。既定は 1")
    p.add_argument("--keyword", help="この文字を含む表を選ぶ (--table より優先)")
    p.add_argument("--roi", type=_parse_roi, help="探す範囲を x,y,幅,高さ (ピクセル) で限定する")
    p.add_argument("-o", "--output", help="出力先 (.tsv / .csv / .xlsx)。省略時は画面に TSV を表示")
    p.add_argument("--clipboard", action="store_true", help="結果をクリップボードに入れる (Excel で Ctrl+V)")
    p.add_argument("--lang", default="jpn+eng", help="tesseract の言語 (既定: jpn+eng)")
    p.add_argument("--keep-newlines", action="store_true", help="セル内の改行を残す")
    p.add_argument("--no-deskew", action="store_true", help="傾き補正をしない")
    p.add_argument("--jobs", type=int, default=os.cpu_count() or 1, help="同時に読み取るセル数")
    p.add_argument("--tesseract-cmd", help="tesseract 本体の場所 (PATH に無い場合)")
    p.add_argument("--tess-config", default="", help="tesseract に渡す追加設定")
    p.add_argument("--debug-dir", help="検出結果の確認用画像を保存するフォルダ")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return p


def _log(msg: str) -> None:
    print(msg, file=sys.stderr)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return _run(args)
    except pytesseract.TesseractNotFoundError:
        _log("tesseract が見つかりません。インストールするか --tesseract-cmd で場所を指定してください。")
    except pytesseract.TesseractError as exc:
        _log(f"tesseract でエラーが起きました (言語データ {args.lang} が入っているか確認してください): {exc}")
    except (OSError, RuntimeError, ValueError) as exc:
        _log(f"エラー: {exc}")
    return 1


def _run(args: argparse.Namespace) -> int:
    configure_tesseract(args.tesseract_cmd)
    options = OcrOptions(
        lang=args.lang,
        keep_newlines=args.keep_newlines,
        jobs=max(1, args.jobs),
        extra_config=args.tess_config,
    )

    pages = load_pages(args.input)
    if args.page is not None:
        if not 1 <= args.page <= len(pages):
            _log(f"ページ {args.page} はありません (全 {len(pages)} ページ)")
            return 2
        pages = [pages[args.page - 1]]

    if args.debug_dir:
        os.makedirs(args.debug_dir, exist_ok=True)

    # 全対象ページの表を集める
    found = []  # (page, gray, masks, table)
    for page in pages:
        gray = page.gray
        if not args.no_deskew:
            gray, angle = deskew(gray)
            if angle:
                _log(f"ページ {page.index}: 傾きを {angle:.2f} 度補正しました")
        tables, masks = detect_tables(gray, page.dpi, page=page.index, roi=args.roi)
        numbers = list(range(len(found) + 1, len(found) + 1 + len(tables)))
        for table in tables:
            found.append((page, gray, masks, table))
        if args.debug_dir:
            base = os.path.join(args.debug_dir, f"page{page.index}")
            cv2.imwrite(f"{base}_tables.png", draw_tables(gray, tables, numbers))
            cv2.imwrite(f"{base}_lines.png", cv2.bitwise_or(masks.horizontal, masks.vertical))

    if not found:
        _log("罫線付きの表が見つかりませんでした。--debug-dir で検出の様子を確認してください。")
        return 1

    if args.list:
        for i, (page, _, _, t) in enumerate(found, start=1):
            print(f"#{i}  ページ {page.index}  位置 x={t.x} y={t.y} 幅={t.w} 高さ={t.h}  {t.n_rows} 行 x {t.n_cols} 列")
        return 0

    if args.keyword:
        target = args.keyword.replace(" ", "")
        chosen = None
        for i, (page, gray, masks, t) in enumerate(found, start=1):
            if target in ocr_region_text(gray, masks, t, page.dpi, options):
                chosen = i
                break
        if chosen is None:
            _log(f"「{args.keyword}」を含む表が見つかりませんでした")
            return 1
    else:
        chosen = args.table
        if not 1 <= chosen <= len(found):
            _log(f"表 #{chosen} はありません (見つかった表は {len(found)} 個)")
            return 2

    page, gray, masks, table = found[chosen - 1]
    _log(f"表 #{chosen} (ページ {page.index}, {table.n_rows} 行 x {table.n_cols} 列) を読み取ります")
    grid = ocr_table(gray, masks, table, page.dpi, options)
    tsv = to_delimited(grid)

    if args.output:
        write_output(args.output, grid, table)
        _log(f"保存しました: {args.output}")
    show = not args.output and not args.clipboard
    if args.clipboard:
        try:
            copy_to_clipboard(tsv)
            _log("クリップボードにコピーしました。Excel で貼り付けてください。")
        except (OSError, RuntimeError, subprocess.CalledProcessError) as exc:
            _log(f"クリップボードに入れられませんでした ({exc})。代わりに画面に表示します。")
            show = True
    if show:
        sys.stdout.write(tsv.replace("\r\n", "\n"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
