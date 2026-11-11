Option Explicit
' ============================================================
' modTestData : テスト用のデータを作る (本物のデータは使わない)
'   同じ中身のマトリックスを、形だけ変えて何通りも作る。
'   どの形でも同じ結果になることが、このツールの中心の確かめ。
'
'   わざと入れてある食い違い (どの形でも同じ):
'     AB123-0010 : 名称が違う                      → N-3
'     AB123-0020 : S4 の印が無い                   → N-4
'     AB123-0030 : 全角で書いてあるだけ            → 一致
'     AB123-0040 : S1 の印が余分                   → N-5
'     AB123-0060 : 部品表に無い品番                → N-2
'     AB123-0050 : マトリックスに書いていない      → N-1
' ============================================================

Private Function Maru() As String
    Maru = ChrW(&H25CB)
End Function

' 半角の英数字と記号を全角にする (テスト用)
Public Function ToWide(ByVal s As String) As String
    Dim i As Long, c As Long, t As String
    For i = 1 To Len(s)
        c = AscW(Mid(s, i, 1))
        If c >= &H21 And c <= &H7E Then
            t = t & ChrW(c + &HFEE0)
        Else
            t = t & Mid(s, i, 1)
        End If
    Next i
    ToWide = t
End Function

' ---- 部品表 ----
' 見出しは 品番 / 部品名称 / 仕様コード / 員数。データは 12 行
Public Sub FillTestBom(ByVal ws As Worksheet)
    Dim r As Long
    ws.Range("A1:D13").NumberFormat = "@"
    ws.Range("A1").Value = "品番"
    ws.Range("B1").Value = "部品名称"
    ws.Range("C1").Value = "仕様コード"
    ws.Range("D1").Value = "員数"
    r = 1
    BomRow ws, r, "AB123-0010", "ブラケットA", "S1", "1"
    BomRow ws, r, "AB123-0010", "ブラケットA", "S2", "1"
    BomRow ws, r, "AB123-0020", "ブラケットB", "S3", "1"
    BomRow ws, r, "AB123-0020", "ブラケットB", "S4", "1"
    BomRow ws, r, "AB123-0030", "ブラケットC", "S1", "2"
    BomRow ws, r, "AB123-0030", "ブラケットC", "S2", "2"
    BomRow ws, r, "AB123-0030", "ブラケットC", "S3", "2"
    BomRow ws, r, "AB123-0030", "ブラケットC", "S4", "2"
    BomRow ws, r, "AB123-0040", "ブラケットD", "S2", "1"
    BomRow ws, r, "AB123-0040", "ブラケットD", "S4", "1"
    BomRow ws, r, "AB123-0050", "ブラケットE", "S3", "1"
    BomRow ws, r, "ZZ999-0001", "カバー", "S1", "1"
End Sub

Private Sub BomRow(ByVal ws As Worksheet, ByRef r As Long, ByVal part As String, ByVal nm As String, ByVal spec As String, ByVal qty As String)
    r = r + 1
    ws.Cells(r, 1).Value = part
    ws.Cells(r, 2).Value = nm
    ws.Cells(r, 3).Value = spec
    ws.Cells(r, 4).Value = qty
End Sub

' 部品表の読み方 (テスト用)。withFilter = True なら、品番が AB123 で始まるものに絞る
Public Function TestBomConfig(ByVal withFilter As Boolean) As Object
    If withFilter Then
        Set TestBomConfig = NewBomConfig("", 1, "品番", "部品名称", "仕様コード", "", "員数", "品番", "AB123*")
    Else
        Set TestBomConfig = NewBomConfig("", 1, "品番", "部品名称", "仕様コード", "", "員数", "", "")
    End If
End Function

' ---- マトリックスの中身 ----
' 戻り値: (行, 6) = 名称, 品番, S1, S2, S3, S4
'   kind = "OK" : 食い違いを全部直したもの
'   kind = "M5" : 印の代わりに員数。AB123-0030 の S4 だけ 1 (部品表は 2) → N-6
'   それ以外    : 印。上に書いた食い違いが入っている
Public Function TestRows(ByVal kind As String) As Variant
    Dim a(1 To 5, 1 To 6) As Variant, m As String, two As String
    m = Maru()
    If kind = "OK" Then
        SetRow a, 1, "ブラケットA", "AB123-0010", m, m, "", ""
        SetRow a, 2, "ブラケットB", "AB123-0020", "", "", m, m
        SetRow a, 3, "ブラケット" & ChrW(&HFF23), ToWide("ab123-0030"), m, m, m, m
        SetRow a, 4, "ブラケットD", "AB123-0040", "", m, "", m
        SetRow a, 5, "ブラケットE", "AB123-0050", "", "", m, ""
    ElseIf kind = "M5" Then
        SetRow a, 1, "ブラケットA1", "AB123-0010", "1", "1", "", ""
        SetRow a, 2, "ブラケットB", "AB123-0020", "", "", "1", ""
        SetRow a, 3, "ブラケット" & ChrW(&HFF23), ToWide("ab123-0030"), "2", "2", "2", "1"
        SetRow a, 4, "ブラケットD", "AB123-0040", "1", "1", "", "1"
        SetRow a, 5, "ブラケットF", "AB123-0060", "1", "", "", ""
    Else
        SetRow a, 1, "ブラケットA1", "AB123-0010", m, m, "", ""
        SetRow a, 2, "ブラケットB", "AB123-0020", "", "", m, ""
        SetRow a, 3, "ブラケット" & ChrW(&HFF23), ToWide("ab123-0030"), m, m, m, m
        SetRow a, 4, "ブラケットD", "AB123-0040", m, m, "", m
        SetRow a, 5, "ブラケットF", "AB123-0060", m, "", "", ""
    End If
    TestRows = a
End Function

Private Sub SetRow(ByRef a() As Variant, ByVal r As Long, ByVal nm As String, ByVal part As String, _
        ByVal v1 As String, ByVal v2 As String, ByVal v3 As String, ByVal v4 As String)
    a(r, 1) = nm
    a(r, 2) = part
    a(r, 3) = v1
    a(r, 4) = v2
    a(r, 5) = v3
    a(r, 6) = v4
End Sub

Private Sub PutText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal s As String)
    If Len(s) = 0 Then Exit Sub
    ws.Cells(r, c).NumberFormat = "@"
    ws.Cells(r, c).Value = s
End Sub

' ---- マトリックスを 1 枚のシートに作る ----
'   M1 : 品番が縦。見出しは 2 段で、上の段 (STD / DX) は 2 列ずつ結合。上に題名の行が 2 行
'   M2 : M1 の縦と横を入れ替えた形 (品番が横)
'   M3 : M1 に加えて、値のセルの結合、途中の空行と小見出し、備考の列
'   M4 : M1 の上の段を結合せず、STD / 空 / DX / 空 と書いた形
'   M5 : M1 の印を員数にした形
'   OK : M1 の食い違いを全部直した形
Public Sub FillTestMatrix(ByVal ws As Worksheet, ByVal kind As String)
    Dim d As Variant, i As Long, j As Long, r As Long
    d = TestRows(kind)
    PutText ws, 1, 1, "ブラケット 品番マトリックス (テスト用 " & kind & ")"
    PutText ws, 2, 1, "これは作り物のデータです"
    If kind = "M2" Then
        ' 品番が横
        PutText ws, 3, 2, "名称"
        PutText ws, 4, 2, "品番"
        For i = 1 To 5
            PutText ws, 3, 2 + i, CStr(d(i, 1))
            PutText ws, 4, 2 + i, CStr(d(i, 2))
            For j = 1 To 4
                PutText ws, 4 + j, 2 + i, CStr(d(i, 2 + j))
            Next j
        Next i
        ws.Range("A5:A6").Merge
        PutText ws, 5, 1, "STD"
        ws.Range("A7:A8").Merge
        PutText ws, 7, 1, "DX"
        PutText ws, 5, 2, "2WD"
        PutText ws, 6, 2, "4WD"
        PutText ws, 7, 2, "2WD"
        PutText ws, 8, 2, "4WD"
        Exit Sub
    End If
    ' 品番が縦
    PutText ws, 4, 1, "名称"
    PutText ws, 4, 2, "品番"
    If kind = "M4" Then
        PutText ws, 3, 3, "STD"
        PutText ws, 3, 5, "DX"
    Else
        ws.Range("C3:D3").Merge
        PutText ws, 3, 3, "STD"
        ws.Range("E3:F3").Merge
        PutText ws, 3, 5, "DX"
    End If
    PutText ws, 4, 3, "2WD"
    PutText ws, 4, 4, "4WD"
    PutText ws, 4, 5, "2WD"
    PutText ws, 4, 6, "4WD"
    If kind = "M3" Then
        ' 5 行目 = 小見出し (品番は空)、8 行目 = 空行
        PutText ws, 5, 1, ChrW(&HFF1C) & "前側" & ChrW(&HFF1E)
        r = 5
        For i = 1 To 5
            r = r + 1
            If r = 8 Then r = 9
            PutText ws, r, 1, CStr(d(i, 1))
            PutText ws, r, 2, CStr(d(i, 2))
            If i = 1 Then
                ' AB123-0010 : S1 と S2 の印を 1 つに結合
                ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).Merge
                PutText ws, r, 3, CStr(d(i, 3))
            ElseIf i = 3 Then
                ' AB123-0030 : 4 つの印を 1 つに結合
                ws.Range(ws.Cells(r, 3), ws.Cells(r, 6)).Merge
                PutText ws, r, 3, CStr(d(i, 3))
            Else
                For j = 1 To 4
                    PutText ws, r, 2 + j, CStr(d(i, 2 + j))
                Next j
            End If
        Next i
        ws.Range("G3:G4").Merge
        PutText ws, 3, 7, "備考"
        PutText ws, 7, 7, "要確認"
    Else
        For i = 1 To 5
            PutText ws, 4 + i, 1, CStr(d(i, 1))
            PutText ws, 4 + i, 2, CStr(d(i, 2))
            For j = 1 To 4
                PutText ws, 4 + i, 2 + j, CStr(d(i, 2 + j))
            Next j
        Next i
    End If
End Sub

' その形の読み方 (設定)
Public Function TestLayout(ByVal kind As String, ByVal sheetName As String) As Object
    Select Case kind
        Case "M2"
            Set TestLayout = NewLayout("TEST_" & kind, sheetName, O_HORZ, "C4:G4", "C3:G3", "A5:B8", K_MARK, OPT_YES, OPT_NOSPLIT, "")
        Case "M3"
            Set TestLayout = NewLayout("TEST_" & kind, sheetName, O_VERT, "B5:B11", "A5:A11", "C3:G4", K_MARK, OPT_YES, OPT_NOSPLIT, "")
        Case "M5"
            Set TestLayout = NewLayout("TEST_" & kind, sheetName, O_VERT, "B5:B9", "A5:A9", "C3:F4", K_QTY, OPT_YES, OPT_NOSPLIT, "")
        Case Else
            Set TestLayout = NewLayout("TEST_" & kind, sheetName, O_VERT, "B5:B9", "A5:A9", "C3:F4", K_MARK, OPT_YES, OPT_NOSPLIT, "")
    End Select
End Function

' 仕様対応表 (テスト用)。biko = "" なら備考の行なし、"対象外" なら備考を対象外に、"未設定" なら備考を未設定のままにする
'   戻り値は modSpecMap.LoadSpecMap と同じ形 (行, 4)
Public Function TestMap(ByVal biko As String) As Variant
    Dim a() As Variant, n As Long
    n = 4
    If Len(biko) > 0 Then n = 5
    ReDim a(1 To n, 1 To 4)
    a(1, 1) = "STD / 2WD": a(1, 2) = "S1": a(1, 3) = False: a(1, 4) = "S1"
    a(2, 1) = "STD / 4WD": a(2, 2) = "S2": a(2, 3) = False: a(2, 4) = "S2"
    a(3, 1) = "DX / 2WD": a(3, 2) = "S3": a(3, 3) = False: a(3, 4) = "S3"
    a(4, 1) = "DX / 4WD": a(4, 2) = "S4": a(4, 3) = False: a(4, 4) = "S4"
    If n = 5 Then
        a(5, 1) = "備考": a(5, 2) = "": a(5, 3) = (biko = MAP_EXCLUDE): a(5, 4) = ""
    End If
    TestMap = a
End Function

' テスト用の仕様対応表を、シート「仕様対応表」に書く (その設定名の行を入れ替える)
Public Sub WriteTestMapToSheet(ByVal layName As String, ByVal biko As String)
    Dim ws As Worksheet, old As Variant, m As Variant, out() As Variant, n As Long, i As Long, j As Long, total As Long
    Set ws = ToolSheet(SH_MAP)
    old = ReadTable(ws, PC_N)
    m = TestMap(biko)
    total = RowCount(old) + RowCount(m)
    ReDim out(1 To total, 1 To PC_N)
    n = 0
    For i = 1 To RowCount(old)
        If Trim(ToText(old(i, PC_LAYOUT))) <> layName Then
            n = n + 1
            For j = 1 To PC_N
                out(n, j) = ToText(old(i, j))
            Next j
        End If
    Next i
    For i = 1 To RowCount(m)
        n = n + 1
        out(n, PC_LAYOUT) = layName
        out(n, PC_HEAD) = m(i, 1)
        out(n, PC_SPEC) = m(i, 4)
        If m(i, 3) = True Then out(n, PC_EXCL) = MAP_EXCLUDE Else out(n, PC_EXCL) = ""
        out(n, PC_MEMO) = "テスト用"
    Next i
    For i = n + 1 To total
        For j = 1 To PC_N
            out(i, j) = ""
        Next j
    Next i
    WriteTable ws, out, PC_N
End Sub

' ---- テスト用のブックを、ツール用ブックと同じ場所の test フォルダに作る ----
'   作ったフォルダの場所を返す
Public Function MakeTestFiles() As String
    Dim folder As String, kinds As Variant, i As Long, wb As Workbook
    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise ERR_TOOL, "MakeTestFiles", "先に、このツール用ブックを保存してください (保存した場所に test フォルダを作ります)。"
    End If
    folder = ThisWorkbook.Path & Application.PathSeparator & "test"
    EnsureFolder folder
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    FillTestBom wb.Worksheets(1)
    wb.SaveAs Filename:=folder & Application.PathSeparator & "bom_test.xlsx", FileFormat:=51
    wb.Close SaveChanges:=False
    kinds = Array("M1", "M2", "M3", "M4", "M5", "OK")
    For i = LBound(kinds) To UBound(kinds)
        Set wb = Workbooks.Add
        FillTestMatrix wb.Worksheets(1), CStr(kinds(i))
        wb.SaveAs Filename:=folder & Application.PathSeparator & "matrix_" & kinds(i) & ".xlsx", FileFormat:=51
        wb.Close SaveChanges:=False
    Next i
    Application.DisplayAlerts = True
    MakeTestFiles = folder
End Function
