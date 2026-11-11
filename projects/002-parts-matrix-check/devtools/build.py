"""持ち込み用のファイルを作る。
   1. vba/ の点検 (lint_vba.py)
   2. vba_import/ : VBA の画面に取り込むための .bas (Shift-JIS、改行 CRLF、モジュール名つき)
   3. handover/照合ツール_ひな形.xlsx : シートを用意済みのブック + 使い方 + コード (シートに貼ってある)
   4. handover/結果の見本.xlsx       : テスト用のデータを照合した結果のブック
   3 と 4 は、LibreOffice (VBA 互換モード) で実際にマクロを動かして作る。
   使い方: python3 build.py"""
import os, sys, shutil, subprocess, copy, traceback
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import uno
import lo

here = os.path.dirname(os.path.abspath(__file__))
proj = os.path.dirname(here)
vba = os.path.join(proj, "vba")
imp = os.path.join(proj, "vba_import")
out = os.path.join(proj, "handover")
work = os.path.join(here, "work")
FONT = "游ゴシック"
MONO = "ＭＳ ゴシック"
TOOL_XLSX = "照合ツール_ひな形.xlsx"
SAMPLE_XLSX = "結果の見本.xlsx"

DRV = '''
Option Explicit
Function Drv_Setup() As String
    On Error GoTo EH
    SetupTool
    Drv_Setup = "OK"
    Exit Function
EH:
    Drv_Setup = "ERR " & Err.Number & " " & Err.Description
End Function
Function Drv_Test() As String
    On Error GoTo EH
    Drv_Test = RunSelfTest()
    Exit Function
EH:
    Drv_Test = "ERR " & Err.Number & " " & Err.Description & " / after " & TestProgress()
End Function
' テスト用の M3 を照合して、結果のブックを開いたままにする
Function Drv_Sample() As String
    Dim folder As String, sep As String, pBom As String, pMtx As String
    Dim wbB As Workbook, wbM As Workbook, outWb As Workbook, verdict As String, lay As Object, msg As String
    On Error GoTo EH
    sep = Application.PathSeparator
    folder = MakeTestFiles()
    pBom = folder & sep & "bom_test.xlsx"
    pMtx = folder & sep & "matrix_M3.xlsx"
    Set wbB = OpenReadOnly(pBom)
    Set wbM = OpenReadOnly(pMtx)
    msg = RunReadBom(wbB.Worksheets(1), TestBomConfig(True), pBom)
    Set lay = NewLayout("見本_ブラケット", wbM.Worksheets(1).Name, O_VERT, "B5:B11", "A5:A11", "C3:G4", K_MARK, OPT_YES, OPT_NOSPLIT, "")
    SaveLayout lay
    msg = RunReadMatrix(wbM.Worksheets(1), lay, pMtx)
    WriteTestMapToSheet "見本_ブラケット", MAP_EXCLUDE
    msg = RunCompare(wbM.Worksheets(1), False, outWb, verdict)
    Drv_Sample = "OK " & verdict
    Exit Function
EH:
    Drv_Sample = "ERR " & Err.Number & " " & Err.Description
End Function
'''

def step(msg):
    print("==", msg, flush=True)

def write_import_files():
    shutil.rmtree(imp, ignore_errors=True)
    os.makedirs(imp)
    for fn in sorted(os.listdir(vba)):
        if not fn.endswith(".bas"):
            continue
        name = fn[:-4]
        src = lo.read(os.path.join(vba, fn))
        text = f'Attribute VB_Name = "{name}"\n' + src
        data = text.replace("\r\n", "\n").replace("\n", "\r\n").encode("cp932")
        with open(os.path.join(imp, fn), "wb") as f:
            f.write(data)

def store_xlsx(doc, path):
    doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(path)), (lo.pv("FilterName", "Calc MS Excel 2007 XML"),))

def find_doc_with_sheet(ctx, sheet):
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    en = desktop.Components.createEnumeration()
    while en.hasMoreElements():
        d = en.nextElement()
        try:
            if d.Sheets.hasByName(sheet):
                return d
        except Exception:
            pass
    return None

def set_fonts(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    f = copy.copy(c.font)
                    f.name = FONT
                    c.font = f

def add_text_sheet(wb, title, lines, index=0, widths=(14, 120)):
    ws = wb.create_sheet(title, index)
    r = 0
    for kind, text in lines:
        r += 1
        if kind == "h1":
            c = ws.cell(r, 2, text); c.font = Font(name=FONT, bold=True, size=16)
        elif kind == "h2":
            c = ws.cell(r, 2, text); c.font = Font(name=FONT, bold=True, size=12)
            c.fill = PatternFill("solid", fgColor="DDEBF7"); ws.cell(r, 1).fill = PatternFill("solid", fgColor="DDEBF7")
        elif kind == "":
            pass
        else:
            ws.cell(r, 1, kind).font = Font(name=FONT)
            c = ws.cell(r, 2, text); c.font = Font(name=FONT); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = widths[0]
    ws.column_dimensions["B"].width = widths[1]
    return ws

def add_code_sheets(wb):
    names = []
    for fn in sorted(os.listdir(vba)):
        if not fn.endswith(".bas"):
            continue
        name = fn[:-4]
        ws = wb.create_sheet("コード_" + name)
        names.append(name)
        for i, line in enumerate(lo.read(os.path.join(vba, fn)).split("\n"), 1):
            c = ws.cell(i, 1)
            c.value = line
            c.data_type = "s"
            c.number_format = "@"
            c.font = Font(name=MONO, size=10)
        ws.column_dimensions["A"].width = 130
        c = ws.cell(1, 3, "← A 列を全部コピーして、VBA の標準モジュールに貼り付けます (シート「はじめに」の手順 3)")
        c.font = Font(name=FONT, bold=True, color="C00000")
    return names

def intro_lines(names, test_summary):
    L = []
    a = L.append
    a(("h1", "部品表とマトリックスの照合ツール  ひな形 (引き継ぎ用)"))
    a(("", ""))
    a(("h2", "このファイルは何か"))
    a(("・", "部品表と品番マトリックスが合っているかを確かめるツールの「土台」です。形が変わらない部分 (不動の部分) を、先に作ってあります。"))
    a(("・", "シートは用意済みです。VBA のコードは、まだブックに入っていません (マクロ付きのファイルは社外から持ち込みにくいため)。下の手順で入れます。"))
    a(("・", "マトリックスの細かい形、部品表の列の名前など、まだ決まっていない部分は、シートの設定と、社内の Copilot での手直しで合わせます。"))
    a(("・", "くわしい考え方と決まりは、構想書 (spec.md) にあります。Copilot には、構想書と、直したいモジュールのコードを渡します。"))
    a(("", ""))
    a(("h2", "最初にすること (1 回だけ)"))
    a(("1", "このブックを「名前を付けて保存」で、種類を「Excel マクロ有効ブック (*.xlsm)」にして保存します。保存先は、自分の PC の中のフォルダにします。"))
    a(("2", "Alt キーを押しながら F11 を押して、VBA の画面を開きます。"))
    a(("3", "コードを入れます。次のどちらかで行います。\n"
            "  (ア) vba_import フォルダの .bas ファイルを全部、VBA の画面の左側 (プロジェクトの一覧) へドラッグします。「ファイル」→「ファイルのインポート」でも入ります。\n"
            "  (イ) このブックの「コード_…」のシートを 1 枚ずつ開き、A 列を全部コピーします。VBA の画面で「挿入」→「標準モジュール」を選び、貼り付けます。これをシートの数 (" + str(len(names)) + " 枚) だけ繰り返します。\n"
            "       (イ) では、行の先頭の「'」が消えていないかを確かめてください。消えていたら (ア) を使ってください。"))
    a(("4", "VBA の画面で「デバッグ」→「VBAProject のコンパイル」を押します。何も表示されなければ合格です。エラーが出たら、その行と文を Copilot に見せて直します。"))
    a(("5", "Excel に戻り、Alt キーを押しながら F8 を押して、「SetupTool」を選んで実行します。シート「操作」にボタンができます。"))
    a(("6", "シート「操作」の「自己診断 (テスト) を実行」を押します。「合格 " + test_summary + "」と出れば、土台は正しく動いています。"))
    a(("7", "「コード_…」のシートは、コードを入れ終わったら消して構いません。"))
    a(("", ""))
    a(("h2", "ふだんの使い方"))
    a(("①", "「① 部品表を読む」を押して、部品表のファイルを選びます。列の名前は、シート「設定_部品表」に書きます。"))
    a(("②", "「② マトリックスを読む」を押して、マトリックスのファイルを選びます。初めての形のときは、品番・名称・仕様の見出しの場所をマウスで教えます (次からは覚えています)。"))
    a(("③", "シート「仕様対応表」で、マトリックスの見出しと、部品表の仕様コードの対応を確かめます。空いている所を書きます。比べない列は「対象外」と書きます。"))
    a(("④", "「③ 照合する」を押します。結果が新しいブックにまとまり、「照合結果」フォルダに保存されます。見本は「" + SAMPLE_XLSX + "」です。"))
    a(("", ""))
    a(("h2", "本当にチェックできているかを確かめる仕組み"))
    a(("・", "消し込み: マトリックスの写しに、見たセル全部へ色を付けます (緑 = 一致、赤 = 不一致、黄 = 判断できない、灰 = 対象外)。色の無いセルは、ツールが見ていないセルです。部品表も、全部の行に状態を付けます。"))
    a(("・", "検算: 読んだ数と状態の数が合うか、別の数え方と合うかを、毎回自動で確かめます。合わないときは、判定が「検算が合わない」になります。"))
    a(("・", "抜き打ち: 本物のデータのうえで、値を 1 か所ずつわざと逆にして、照合が気付くかを毎回確かめます (元のファイルは変えません)。"))
    a(("・", "自己診断: わざと間違いを入れた作り物のデータで、決まった数の食い違いが見つかるかを確かめます。"))
    a(("・", "抜き取り確認: ツールが「一致」としたセルを 10 個ほど選んで一覧にします。人が元の表と見比べて確かめます。"))
    a(("・", "おすすめ: 最初の数件は、今までどおり紙と赤ペンでも確かめて、ツールの結果と同じになるかを見てください。"))
    a(("", ""))
    a(("h2", "確かめてあること、まだのこと"))
    a(("済", "LibreOffice (Excel と似たソフト) の VBA 互換の機能でコードを動かし、自己診断が全部合格することを確かめました (" + test_summary + ")。"))
    a(("未", "本物の Excel では、まだ一度も動かしていません。手順 4 と 6 で確かめてください。"))
    a(("未", "ボタンを押してファイルを選ぶ所、マウスで範囲を選ぶ所 (モジュール modMain と、modSetup のボタン作り) は、画面が要るため、LibreOffice でも動かしていません。"))
    a(("未", "本物の部品表とマトリックスでは試していません。列の名前や形は、会社で合わせる必要があります (構想書 12.1 の確認事項)。"))
    a(("", ""))
    a(("h2", "モジュールの一覧 (どれが何をするか)"))
    desc = {
        "modConst": "名前や列の番号など、決まった値",
        "modDict": "辞書 (名前から値を引く入れ物) を作る",
        "modUtil": "小さな道具 (表の読み書き、控えの読み書きなど)",
        "modText": "文字のそろえ方 (全角と半角、ハイフン、印の判定)",
        "modLog": "読み取りの途中で気付いたこと (警告) をためる",
        "modFile": "ファイルを読み取り専用で開く",
        "modBom": "部品表を読む  ★ 部品表の形が想定と違えば、ここを直す",
        "modMatrix": "マトリックスを読む (結合セル、縦と横、空欄の引き継ぎ)  ★ 特別な形に対応するなら、ここを直す",
        "modSpecMap": "仕様の呼び名を合わせる (仕様対応表)",
        "modCompare": "照合の本体と、抜き打ちの確かめ",
        "modReport": "結果のブックを作る (消し込み、検算、抜き取り確認)",
        "modRun": "読み取り → 照合 → 結果 の流れ",
        "modMain": "ボタンから呼ばれる入口 (画面を出す所)  ★ 使い勝手を変えるなら、ここを直す",
        "modSetup": "シートとボタンを用意する",
        "modTestData": "テスト用の作り物のデータ",
        "modTest": "自己診断 (テスト)",
    }
    for n in names:
        a((n, desc.get(n, "")))
    return L

def main():
    step("1. 点検")
    r = subprocess.run([sys.executable, os.path.join(here, "lint_vba.py")], capture_output=True, text=True)
    print(r.stdout.strip().split("\n")[-1])
    if r.returncode != 0:
        print(r.stdout); sys.exit(1)
    step("2. 取り込み用の .bas")
    write_import_files()
    step("3. LibreOffice でマクロを動かす")
    shutil.rmtree(work, ignore_errors=True)
    os.makedirs(work); os.makedirs(out, exist_ok=True)
    os.chdir(work)
    wb = openpyxl.Workbook(); wb.active.title = "Sheet1"; wb.save("tool.xlsx")
    p, ctx = lo.start()
    test_summary = ""
    try:
        doc = lo.load(ctx, "tool.xlsx")
        lo.add_modules(doc, lo.tool_modules(vba, os.path.join(here, "lo_shim")) + [("modDrv", DRV, "std")])
        print("Setup:", lo.run(doc, "modDrv.Drv_Setup"))
        store_xlsx(doc, "tool_out.xlsx")
        res = lo.run(doc, "modDrv.Drv_Test")
        print("Test:", res)
        test_summary = res.replace("合格 ", "")
        nums = [int(x) for x in test_summary.replace("全", "").split("/")]
        if nums[0] != nums[1]:
            print("自己診断に不合格があるため、やめます"); sys.exit(1)
        print("Sample:", lo.run(doc, "modDrv.Drv_Sample"))
        rdoc = find_doc_with_sheet(ctx, "まとめ")
        store_xlsx(rdoc, "sample_out.xlsx")
    finally:
        lo.stop(p)

    step("4. ひな形のブックを仕上げる")
    wb = openpyxl.load_workbook("tool_out.xlsx")
    if "Sheet1" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["Sheet1"]
    set_fonts(wb)
    names = add_code_sheets(wb)
    add_text_sheet(wb, "はじめに", intro_lines(names, test_summary), 0)
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = (ws.title == "はじめに")
    wb.save(os.path.join(out, TOOL_XLSX))

    step("5. 結果の見本を仕上げる")
    wb = openpyxl.load_workbook("sample_out.xlsx")
    set_fonts(wb)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and work in c.value:
                    c.value = c.value.replace(work, "C:\\照合ツール").replace("/test/", "\\test\\")
    ws = wb["まとめ"]
    for row in ws.iter_rows(min_col=1, max_col=2):
        if row[0].value and str(row[0].value).startswith("実行した人"):
            row[1].value = "(見本)"
    c = ws.cell(1, 4, "これは、作り物のテスト用データ (matrix_M3) を照合した結果の見本です。")
    c.font = Font(name=FONT, bold=True, color="C00000")
    wb.active = 0
    wb.save(os.path.join(out, SAMPLE_XLSX))
    print("done:", os.listdir(out))

if __name__ == "__main__":
    main()
