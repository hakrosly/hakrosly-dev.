# 開発構想書集

社内 Copilot (Microsoft 365 Copilot) に読ませるための開発構想書と、必要に応じて試作ツールをまとめたリポジトリです。

## 構想の一覧

| 番号 | 名前 | 状態 | 構想書 (Copilot に渡す URL) | 試作 |
|---|---|---|---|---|
| 001 | TIFF 図面の構成表読み取りツール | 仕様確定待ち (残課題あり) | [spec.md](projects/001-tiff-table-ocr/spec.md) / [raw](https://raw.githubusercontent.com/hakrosly/hakrosly-dev./ccr-77995653-gzbrvc/projects/001-tiff-table-ocr/spec.md) | あり (古い版) |

状態の種類: 構想中 / 仕様確定待ち / 仕様確定 / 社内で開発中 / 完了 / 中止

## Copilot への渡し方

- **構想書のファイル 1 つ (`spec.md`) だけ** を渡す。フォルダや、このリポジトリ全体は渡さない。
- URL は、上の表の **raw** の方 (文字だけのページ) を貼る。普通の GitHub の画面よりも余計な表示が無く、読み取りやすい。
- 構想書は 1 ファイルで完結するように書く (共通の決まりも中に書き写す)。他のファイルを読まなくても分かるようにするため。
- Copilot が URL を開けない場合は、`spec.md` をダウンロードして添付するか、中身を貼り付ける。

## フォルダの構成

```
README.md                  この一覧
common/constraints.md      全構想に共通の決まり (各構想書の 2 章へ書き写す元)
templates/spec_template.md 構想書のひな形
projects/
  001-tiff-table-ocr/
    README.md              概要・状態・経緯
    spec.md                構想書 (Copilot に渡すのはこれだけ)
    decisions.md           決まったこと・未決のことの記録
    prototype/             試作 (ある場合だけ)
```

## 新しい構想の追加手順

1. `projects/` に `番号-英語の短い名前` のフォルダを作る (例: `002-drawing-number-check`)。番号は 3 桁の連番。
2. `templates/spec_template.md` を `spec.md` として写し、中身を書く。
3. `common/constraints.md` の決まりを `spec.md` の 2 章に書き写す。
4. `README.md` と `decisions.md` を作る。
5. この一覧に 1 行足す。

## 決まりごと

- 本物の図面・書類などの社内データは入れない。テストには自分で作った画像を使う。
- 試作は `prototype/` に入れ、構想ごとに必要な部品の一覧 (`requirements.txt`) を持たせる。
- 試作と構想書が食い違うときは、構想書を正とする。試作の README の先頭にもそう書く。
