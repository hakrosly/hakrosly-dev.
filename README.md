# tableocr

TIFF 画像の中にある **罫線付きの表** を見つけて文字を読み取り、Excel にそのまま貼れる形で出力するツールです。

- 表と罫線の検出: OpenCV
- 文字の読み取り: Tesseract OCR (日本語 + 英語)
- 出力: TSV (Excel に貼り付け用) / CSV / xlsx / クリップボード

第一段階として「Excel に表として貼れるところまで」を目標にしています。

## できること

- マルチページの TIFF (FAX 形式の白黒圧縮も含む) を読み込む
- 少し傾いたスキャン画像を自動でまっすぐにする (±5 度まで)
- ページ内の表をすべて見つけて番号を付ける
- 番号・ページ・キーワード・範囲指定で「特定の表」を選ぶ
- 結合セルを判定する (文字は左上のマスに入る。xlsx 出力ではセル結合も再現)
- セル内の罫線のかけらや点ゴミを消してから読み取る

## 準備

### 1. Tesseract 本体を入れる

| OS | 方法 |
|---|---|
| Windows | [UB Mannheim 版インストーラ](https://github.com/UB-Mannheim/tesseract/wiki) を使う。途中の「Additional language data」で **Japanese** にチェックを入れる |
| macOS | `brew install tesseract tesseract-lang` |
| Ubuntu | `sudo apt install tesseract-ocr tesseract-ocr-jpn` |

`tesseract --list-langs` を実行して `jpn` が出れば OK です。

Windows で PATH が通っていない場合は、実行時に `--tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe"` を付けるか、環境変数 `TESSERACT_CMD` に同じパスを入れてください。

### 2. Python のライブラリを入れる

Python 3.10 以上が必要です。

```bash
pip install -r requirements.txt
```

## 使い方

### 手順 1: どんな表が見つかったか確認する

```bash
python -m tableocr scan.tif --list --debug-dir debug
```

```
#1  ページ 1  位置 x=146 y=296 幅=759 高さ=249  3 行 x 3 列
#2  ページ 1  位置 x=146 y=716 幅=1269 高さ=459  5 行 x 4 列
```

`debug` フォルダに次の画像ができます。

- `page1_tables.png`: 見つかった表に番号と枠を描いたもの (緑 = 普通のセル、橙 = 結合セル)
- `page1_lines.png`: 罫線として認識した線だけを描いたもの

### 手順 2: 表を読み取って Excel に貼る

いちばん簡単なのはクリップボードに入れる方法です。実行後に Excel で `Ctrl+V` するだけで表として貼れます。

```bash
python -m tableocr scan.tif --table 2 --clipboard
```

ファイルに保存することもできます。

```bash
python -m tableocr scan.tif --table 2 -o result.xlsx   # Excel ファイル (結合セルも再現)
python -m tableocr scan.tif --table 2 -o result.tsv    # メモ帳で開いて全選択→コピー→Excel に貼り付け
python -m tableocr scan.tif --table 2 -o result.csv
```

`-o` も `--clipboard` も付けなければ、画面に TSV (タブ区切り) で表示します。

### 特定の表の選び方

| 方法 | 例 | 向いている場面 |
|---|---|---|
| 番号 | `--table 2` | まず `--list` で番号を確認してから |
| ページ + 番号 | `--page 3 --table 1` | 3 ページ目の 1 つ目の表 (番号はそのページ内で数える) |
| キーワード | `--keyword 請求明細` | 表の中、または表のすぐ上 (表題) にある文字で探す |
| 範囲 | `--roi 100,650,1500,700` | 決まった様式の帳票で、表の位置がいつも同じとき (x,y,幅,高さ をピクセルで) |

### 主なオプション

| オプション | 意味 |
|---|---|
| `--lang` | 読み取る言語。既定は `jpn+eng`。英数字だけなら `eng` の方が正確なことがある |
| `--keep-newlines` | セル内の改行を残す (既定では空白でつなげる) |
| `--no-deskew` | 傾き補正をしない |
| `--jobs` | 同時に読み取るセル数 (既定は CPU の数) |
| `--tess-config` | tesseract への追加設定。例: `--tess-config "-c tessedit_char_whitelist=0123456789,"` |

## 精度を上げるコツ

- スキャンは **300dpi** 以上がおすすめです。200dpi だと数字の読み間違いが増えます。
- 白黒 (2 値) より、グレースケールでスキャンした方が良いことがあります。
- 読み間違いが多いときは、まず `--debug-dir` の画像で表と罫線が正しく取れているか確認してください。

## 今の制限

- **罫線のある表** が対象です。罫線が無い表 (空白だけで区切った表) は見つけられません。
- 手書き文字は読めません。
- Excel で貼り付けると、`001` のような数字は先頭の 0 が消えます。消したくない場合は `-o result.xlsx` を使ってください (文字列として保存します)。
- 読み取り結果は必ず目で確認してください。特に数字の `0/6/8/9`、`1/l/|` は間違いやすいです。

## テスト

```bash
pip install pytest
python -m pytest
```

テストではダミーの帳票画像 (`tests/sample_factory.py`) を作って、表の検出・結合セル・読み取り結果を確かめています。
自分で試すためのサンプル画像も作れます。

```bash
python tests/sample_factory.py sample.tif
python -m tableocr sample.tif --list --debug-dir debug
```

## ファイル構成

```
tableocr/
  image_io.py      TIFF の読み込み (マルチページ・DPI 取得)
  table_detect.py  傾き補正、罫線検出、表とセル (結合セル含む) の特定
  ocr.py           セルごとの前処理と tesseract による読み取り
  export.py        TSV / CSV / xlsx / クリップボードへの出力
  cli.py           コマンドライン
tests/
  sample_factory.py  テスト用の帳票画像を作る
  test_tableocr.py   テスト
```
