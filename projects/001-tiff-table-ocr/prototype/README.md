# tableocr (試作)

> **注意: これは試作です。正しい仕様は [`../spec.md`](../spec.md) です。**
> 試作は普通の書類向けに作ったもので、図面への対応 (表の探し方)、表題「構成表」(括弧なし)、余計な空白の補正、2 行に分かれた項目のつなぎ方は入っていません。
> 試作を動かすときは、このフォルダ (`prototype/`) で `python -m tableocr ...` や `python -m pytest` を実行してください。

TIFF 画像の中にある **罫線付きの表** を見つけて文字を読み取り、Excel にそのまま貼れる形で出力するツールです。

- 表と罫線の検出: OpenCV
- 文字の読み取り: Tesseract OCR (日本語 + 英語)
- 出力: TSV (Excel に貼り付け用) / CSV / xlsx / クリップボード

第一段階として「Excel に表として貼れるところまで」を目標にしています。

**外部通信は一切しません。** 画像の読み込みから出力まで、すべて手元の PC の中で完結します (詳しくは「外部通信について」)。

既定では、欄外左上に **「（構成表）」** と書かれた表を探して読み取ります。

## できること

- マルチページの TIFF (FAX 形式の白黒圧縮も含む) を読み込む
- 少し傾いたスキャン画像を自動でまっすぐにする (±5 度まで)
- ページ内の表をすべて見つけて番号を付ける
- 欄外左上の表題「（構成表）」で対象の表を自動で選ぶ (番号・ページ・キーワード・範囲指定でも選べる)
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

インターネットにつながらない PC に入れる場合は、つながる PC で部品をまとめて取ってきて持ち込みます。

```bash
# つながる PC で (対象 PC と同じ OS・同じ Python の版で行う)
pip download -r requirements.txt -d wheels

# wheels フォルダを持ち込んで、対象 PC で
pip install --no-index --find-links wheels -r requirements.txt
```

Tesseract もインストーラ (日本語データ込み) を持ち込めば、通信なしで入れられます。

## 使い方

### 手順 1: どんな表が見つかったか確認する

```bash
python -m tableocr scan.tif --list --debug-dir debug
```

```
#1  ページ 1  位置 x=146 y=296 幅=759 高さ=249  3 行 x 3 列  表題: (別表)
#2  ページ 1  位置 x=146 y=716 幅=1269 高さ=459  5 行 x 4 列  表題: (構成表)  ← 対象
```

「← 対象」が付いた表が、何も指定しないときに読み取られる表です。

`debug` フォルダに次の画像ができます。

- `page1_tables.png`: 見つかった表に番号と枠を描いたもの (緑 = 普通のセル、橙 = 結合セル)
- `page1_lines.png`: 罫線として認識した線だけを描いたもの

### 手順 2: 表を読み取って Excel に貼る

いちばん簡単なのはクリップボードに入れる方法です。実行後に Excel で `Ctrl+V` するだけで表として貼れます。

```bash
python -m tableocr scan.tif --clipboard
```

ファイルに保存することもできます。

```bash
python -m tableocr scan.tif -o result.xlsx   # Excel ファイル (結合セルも再現)
python -m tableocr scan.tif -o result.tsv    # メモ帳で開いて全選択→コピー→Excel に貼り付け
python -m tableocr scan.tif -o result.csv
```

`-o` も `--clipboard` も付けなければ、画面に TSV (タブ区切り) で表示します。

### 特定の表の選び方

何も指定しなければ「表題」で選びます。

| 方法 | 例 | 向いている場面 |
|---|---|---|
| 表題 (既定) | 指定なし、または `--title 構成表` | 欄外左上に「（構成表）」と書かれた表 |
| 番号 | `--table 2` | まず `--list` で番号を確認してから |
| ページ + 番号 | `--page 3 --table 1` | 3 ページ目の 1 つ目の表 (番号はそのページ内で数える) |
| キーワード | `--keyword 請求明細` | 表の中、または表のすぐ上 (表題) にある文字で探す |
| 範囲 | `--roi 100,650,1500,700` | 決まった様式の帳票で、表の位置がいつも同じとき (x,y,幅,高さ をピクセルで) |

優先順位は `--table` → `--keyword` → 表題 です。`--page` と `--roi` は、どの方法とも組み合わせられます。

#### 表題の見つけ方

- 表のすぐ上 (0.8 インチ以内) で、表の左半分の範囲にある文字を「欄外左上の表題」とみなします。
- その範囲に文字の行がいくつかあるときは、表に一番近い 1 行だけを読みます。本文中に「構成表」という言葉があっても選ばれません。
- 括弧は読み間違いが多いため、比べるのは **括弧の中の文字 (構成表)** だけです。全角・半角や空白の違いは無視します。
- 「（構成表）」の表が複数あるときは、一番先にある表を読み、残りの番号を画面に知らせます。

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

## 外部通信について

- 使う部品 (OpenCV、Tesseract、openpyxl など) はすべて手元で動くもので、通信を必要としません。
- 念のため、ツールは起動するとすぐにプロセス内の通信機能 (接続・名前解決) を封じます。万一どこかが通信しようとしても、エラーで止まります (`tableocr/offline.py`)。
- 起動する外部プログラムは `tesseract` だけです。クリップボードを使うときは OS 付属のコピー用コマンド (`clip` など) も起動します。
- テスト (`test_full_run_makes_no_network_access`) で、実行中に通信が一度も起きないことと、起動するのが tesseract だけであることを確かめています。

## 残課題

- **表題の例外表記の確認。** 今は「構成表」を含むかどうかだけで判定しています。実際の書類で次のような表記が無いか確認し、必要なら判定を直す。
  - 括弧なし、半角括弧、【構成表】など別の括弧
  - 「構成表（続き）」「構成表その2」など、続きや複数の表
  - 表題が左上以外 (中央、右上、表の中、表の下) にある書類
  - 「構成表」以外の言い方 (例: 構成一覧) をしている書類

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
