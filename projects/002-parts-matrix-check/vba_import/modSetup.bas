Attribute VB_Name = "modSetup"
Option Explicit
' ============================================================
' modSetup : ツール用ブックのシートを用意する
'   足りないシートを作り、見出しと初期値を入れる。
'   すでに入っている設定は消さない。何度実行してもよい。
' ============================================================

Private Function EnsureSheet(ByVal sheetName As String, ByRef created As Boolean) As Worksheet
    Dim ws As Worksheet
    Set ws = SheetOrNothing(ThisWorkbook, sheetName)
    created = False
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        created = True
    End If
    Set EnsureSheet = ws
End Function

Private Sub PutRow(ByVal ws As Worksheet, ByVal r As Long, ByVal a As String, ByVal b As String, ByVal c As String)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).NumberFormat = "@"
    ws.Cells(r, 1).Value = a
    ws.Cells(r, 2).Value = b
    ws.Cells(r, 3).Value = c
End Sub

Public Sub SetupTool()
    Dim ws As Worksheet, created As Boolean, tbl As Variant, i As Long

    ' ---- 操作 ----
    Set ws = EnsureSheet(SH_MAIN, created)
    SetupMainSheet ws

    ' ---- 設定_部品表 ----
    Set ws = EnsureSheet(SH_CFG_BOM, created)
    WriteHeader ws, Array("項目", "値", "説明")
    If RowCount(ReadTable(ws, 2)) = 0 Then
        PutRow ws, 2, "シート名", "", "部品表のシートの名前。空なら 1 枚目のシート"
        PutRow ws, 3, "見出しの行", "1", "見出しがある行の番号"
        PutRow ws, 4, "品番の列", "品番", "見出しの文字で書く"
        PutRow ws, 5, "名称の列", "部品名称", "見出しの文字で書く。空なら名称は比べない"
        PutRow ws, 6, "仕様の列", "仕様コード", "見出しの文字で書く。複数あるときはカンマで並べる (値を「 / 」でつないで 1 つの仕様にする)"
        PutRow ws, 7, "仕様の区切り文字", "", "1 つのセルに仕様が複数入るときの区切り (例: , )。空なら分けない"
        PutRow ws, 8, "員数の列", "", "見出しの文字で書く。空なら員数は比べない"
        PutRow ws, 9, "絞り込みの列", "", "そのマトリックスが扱う部品だけに絞るための列。空なら絞らない (書きもれ N-1 は調べない)"
        PutRow ws, 10, "絞り込みの値", "", "Like の書き方 (例: AB123*)。* は何文字でも、? は 1 文字、# は数字 1 文字"
    End If

    ' ---- 設定_マトリックス ----
    Set ws = EnsureSheet(SH_CFG_MTX, created)
    WriteHeader ws, Array("設定名", "シート名", "品番の向き", "品番の範囲", "名称の範囲", "仕様の見出しの範囲", "値の種類", _
        "見出しの空欄の引き継ぎ", "1 セルに複数の品番", "品番の形", "最後に保存した日時", "メモ")
    ws.Cells(1, 14).Value = "書き方: 品番の向き = 縦 / 横、値の種類 = 印 / 員数、引き継ぎ = する / しない、複数の品番 = 分ける / 分けない、品番の形 = Like の書き方 (空でもよい)"

    ' ---- 設定_文字 ----
    Set ws = EnsureSheet(SH_CFG_TXT, created)
    WriteHeader ws, Array("種類", "文字", "説明")
    If RowCount(ReadTable(ws, 2)) = 0 Then
        tbl = DefaultTextRows()
        For i = LBound(tbl, 1) To UBound(tbl, 1)
            PutRow ws, i + 1, CStr(tbl(i, 1)), CStr(tbl(i, 2)), CStr(tbl(i, 3))
        Next i
    End If

    ' ---- 仕様対応表 ----
    Set ws = EnsureSheet(SH_MAP, created)
    WriteHeader ws, Array("設定名", "見出しの文字", "仕様コード", "対象外", "メモ")
    ws.Cells(1, 7).Value = "書き方: 見出し 1 つが仕様コード複数に当たるときは、行を増やして書く。比べない見出しは「対象外」の列に「対象外」と書く"

    ' ---- 読み取りの結果 ----
    Set ws = EnsureSheet(SH_RD_BOM, created)
    WriteHeader ws, Array("品番キー", "品番", "名称", "仕様コード (比べる形)", "仕様コード", "員数", "範囲の中", "元の行")
    Set ws = EnsureSheet(SH_RD_MTX, created)
    WriteHeader ws, Array("品番キー", "品番", "名称", "見出しの文字", "値", "員数", "セルに入っていた文字", _
        "品番のセル", "値のセル", "名称のセル", "見出しのセル")
    Set ws = EnsureSheet(SH_RD_WARN, created)
    WriteHeader ws, Array("出どころ", "記号", "品番", "見出しの文字", "セル", "元の文字", "説明")
    Set ws = EnsureSheet(SH_STATE, created)
    WriteHeader ws, Array("名前", "値")
    Set ws = EnsureSheet(SH_TEST, created)
    WriteHeader ws, Array("No", "内容", "期待", "実際", "合否")

    TextLoad
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_MAIN).Activate
    On Error GoTo 0
End Sub

Private Sub SetupMainSheet(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部品表とマトリックスの照合ツール (版 " & TOOL_VERSION & ")"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "上から順にボタンを押します。元のファイルは書き換えません。"
    ws.Range("F4").Value = "部品表のファイルを選ぶと、シート「" & SH_RD_BOM & "」に読み取った中身が出ます。列の名前はシート「" & SH_CFG_BOM & "」で決めます。"
    ws.Range("F6").Value = "マトリックスのファイルを選びます。初めての形のときは、品番・名称・仕様の見出しの場所をマウスで教えます。"
    ws.Range("F8").Value = "先にシート「" & SH_MAP & "」で、見出しと仕様コードの対応を確かめます。結果は新しいブックに出ます。"
    ws.Range("F10").Value = "保存してある形を、マウスで選び直します (行や列が増えたとき)。"
    ws.Range("F12").Value = "ツールが正しく動くかを、わざと間違いを入れたデータで確かめます。最初と、コードを直した後に実行します。"
    ws.Range("F14").Value = "操作の練習に使える、テスト用の部品表とマトリックスを作ります。"
    ws.Range("F16").Value = "シートを作り直します (設定は消えません)。"
    AddButton ws, 4, "① 部品表を読む", "Btn_ReadBom"
    AddButton ws, 6, "② マトリックスを読む", "Btn_ReadMatrix"
    AddButton ws, 8, "③ 照合する", "Btn_Compare"
    AddButton ws, 10, "形を教え直す", "Btn_TeachAgain"
    AddButton ws, 12, "自己診断 (テスト) を実行", "Btn_SelfTest"
    AddButton ws, 14, "テスト用のブックを作る", "Btn_MakeTestFiles"
    AddButton ws, 16, "シートを用意し直す", "Btn_Setup"
End Sub

' ボタンを置く。ボタンが作れない環境では、マクロの名前を書いておく (Alt + F8 から実行できる)
Private Sub AddButton(ByVal ws As Worksheet, ByVal r As Long, ByVal caption As String, ByVal macroName As String)
    Dim c As Range, b As Object
    Set c = ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
    ws.Cells(r, 1).Value = ""
    On Error Resume Next
    ws.Buttons("btn_" & macroName).Delete
    Err.Clear
    Set b = Nothing
    Set b = ws.Buttons.Add(c.Left, c.Top, c.Width, c.Height * 1.5)
    If Err.Number = 0 And Not b Is Nothing Then
        b.Caption = caption
        b.OnAction = macroName
        b.Name = "btn_" & macroName
        ws.Cells(r, 2).Value = ""
    Else
        ws.Cells(r, 2).Value = caption & " (マクロ名: " & macroName & ")"
    End If
    Err.Clear
    On Error GoTo 0
End Sub
