Option Explicit
' ============================================================
' modUtil : どのモジュールからも使う小さな道具
' ============================================================

' 値が空かどうか (空のセル、長さ 0 の文字)
Public Function IsBlankV(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsBlankV = False
    ElseIf IsEmpty(v) Then
        IsBlankV = True
    ElseIf IsNull(v) Then
        IsBlankV = True
    ElseIf VarType(v) = vbString Then
        IsBlankV = (Len(v) = 0)
    Else
        IsBlankV = False
    End If
End Function

' セルの値を文字にする (エラー値と空は "")
Public Function ToText(ByVal v As Variant) As String
    If IsError(v) Then
        ToText = ""
    ElseIf IsEmpty(v) Then
        ToText = ""
    ElseIf IsNull(v) Then
        ToText = ""
    Else
        ToText = CStr(v)
    End If
End Function

' 見せるための文字にする (改行を取り、前後の空白を取る。全角と半角はそのまま)
Public Function CleanText(ByVal v As Variant) As String
    Dim t As String
    t = ToText(v)
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, ChrW(&H3000), " ")
    CleanText = Trim(t)
End Function

' 数値などが入ったセルの「表示されている文字」を返す
Public Function ShownText(ByVal c As Range, ByVal v As Variant) As String
    Dim t As String, i As Long, allSharp As Boolean
    On Error Resume Next
    If c.MergeCells Then
        t = c.MergeArea.Cells(1, 1).Text
    Else
        t = c.Text
    End If
    On Error GoTo 0
    allSharp = (Len(t) > 0)
    For i = 1 To Len(t)
        If Mid(t, i, 1) <> "#" Then allSharp = False
    Next i
    If Len(t) = 0 Or allSharp Then
        If IsNumeric(v) Then t = Format(v, "0") Else t = ToText(v)
    End If
    ShownText = t
End Function

Public Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function

Public Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

Public Function Z2(ByVal n As Long) As String
    Z2 = Right("0" & CStr(n), 2)
End Function

' 日時を「2026-11-11 16:05:09」の形にする
Public Function StampText(ByVal d As Date) As String
    StampText = CStr(Year(d)) & "-" & Z2(Month(d)) & "-" & Z2(Day(d)) & " " & _
                Z2(Hour(d)) & ":" & Z2(Minute(d)) & ":" & Z2(Second(d))
End Function

' 日時を「20261111_160509」の形にする (ファイルの名前用)
Public Function StampFile(ByVal d As Date) As String
    StampFile = CStr(Year(d)) & Z2(Month(d)) & Z2(Day(d)) & "_" & _
                Z2(Hour(d)) & Z2(Minute(d)) & Z2(Second(d))
End Function

' 2 次元の配列の行数 (空なら 0)
Public Function RowCount(ByVal a As Variant) As Long
    If IsEmpty(a) Then
        RowCount = 0
    ElseIf Not IsArray(a) Then
        RowCount = 0
    Else
        RowCount = UBound(a, 1) - LBound(a, 1) + 1
    End If
End Function

' (列, 行) の向きでためた入れ物を、(行, 列) の配列に直す
Public Function BufToRows(ByRef buf() As Variant, ByVal n As Long, ByVal nCol As Long) As Variant
    Dim out() As Variant, i As Long, j As Long
    If n <= 0 Then
        BufToRows = Empty
        Exit Function
    End If
    ReDim out(1 To n, 1 To nCol)
    For i = 1 To n
        For j = 1 To nCol
            out(i, j) = buf(j, i)
        Next j
    Next i
    BufToRows = out
End Function

' シートがあれば返す。無ければ Nothing
Public Function SheetOrNothing(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set SheetOrNothing = Nothing
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            Set SheetOrNothing = ws
            Exit Function
        End If
    Next ws
End Function

' ツール用ブックのシートを返す。無ければエラーにする
Public Function ToolSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = SheetOrNothing(ThisWorkbook, sheetName)
    If ws Is Nothing Then
        Err.Raise ERR_TOOL, "ToolSheet", "シート「" & sheetName & "」がありません。マクロ「SetupTool」を実行してください。"
    End If
    Set ToolSheet = ws
End Function

' シートの、見出し (1 行目) より下に入っている表を配列で返す。無ければ Empty
Public Function ReadTable(ByVal ws As Worksheet, ByVal nCol As Long) As Variant
    Dim ur As Range, lastRow As Long, v As Variant, r As Long, c As Long
    Dim out() As Variant, n As Long, hasAny As Boolean
    Set ur = ws.UsedRange
    lastRow = ur.Row + ur.Rows.Count - 1
    If lastRow < 2 Then
        ReadTable = Empty
        Exit Function
    End If
    v = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, nCol + 1)).Value
    ' 後ろの空の行は数えない
    n = 0
    For r = 1 To lastRow - 1
        hasAny = False
        For c = 1 To nCol
            If Not IsBlankV(v(r, c)) Then hasAny = True
        Next c
        If hasAny Then n = r
    Next r
    If n = 0 Then
        ReadTable = Empty
        Exit Function
    End If
    ReDim out(1 To n, 1 To nCol)
    For r = 1 To n
        For c = 1 To nCol
            out(r, c) = v(r, c)
        Next c
    Next r
    ReadTable = out
End Function

' シートの 2 行目から下を消して、配列を書く (文字として書くので、先頭の 0 などが変わらない)
Public Sub WriteTable(ByVal ws As Worksheet, ByVal a As Variant, ByVal nCol As Long)
    Dim ur As Range, lastRow As Long, n As Long
    Set ur = ws.UsedRange
    lastRow = ur.Row + ur.Rows.Count - 1
    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, MaxL(nCol, ur.Column + ur.Columns.Count - 1))).ClearContents
    End If
    n = RowCount(a)
    If n = 0 Then Exit Sub
    With ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, nCol))
        .NumberFormat = "@"
        .Value = a
    End With
End Sub

' シートの 1 行目に見出しを書く
Public Sub WriteHeader(ByVal ws As Worksheet, ByVal names As Variant)
    Dim i As Long, c As Long
    c = 0
    For i = LBound(names) To UBound(names)
        c = c + 1
        ws.Cells(1, c).Value = names(i)
        ws.Cells(1, c).Font.Bold = True
        ws.Cells(1, c).Interior.Color = COL_HEAD
    Next i
End Sub

' ---- 読取_記録 (前の段階から次の段階へ渡す、名前と値の控え) ----

Public Sub StateSet(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, v As Variant, n As Long, r As Long
    Set ws = ToolSheet(SH_STATE)
    v = ReadTable(ws, 2)
    n = RowCount(v)
    For r = 1 To n
        If ToText(v(r, 1)) = key Then
            ws.Cells(r + 1, 2).NumberFormat = "@"
            ws.Cells(r + 1, 2).Value = ToText(value)
            Exit Sub
        End If
    Next r
    ws.Cells(n + 2, 1).NumberFormat = "@"
    ws.Cells(n + 2, 2).NumberFormat = "@"
    ws.Cells(n + 2, 1).Value = key
    ws.Cells(n + 2, 2).Value = ToText(value)
End Sub

Public Function StateGet(ByVal key As String) As String
    Dim ws As Worksheet, v As Variant, r As Long
    StateGet = ""
    Set ws = ToolSheet(SH_STATE)
    v = ReadTable(ws, 2)
    For r = 1 To RowCount(v)
        If ToText(v(r, 1)) = key Then
            StateGet = ToText(v(r, 2))
            Exit Function
        End If
    Next r
End Function

Public Function StateNum(ByVal key As String) As Double
    Dim t As String
    t = StateGet(key)
    If IsNumeric(t) Then StateNum = CDbl(t) Else StateNum = 0
End Function

' 名前の先頭が prefix の控えを全部消す
Public Sub StateClear(ByVal prefix As String)
    Dim ws As Worksheet, v As Variant, n As Long, r As Long, m As Long
    Dim out() As Variant
    Set ws = ToolSheet(SH_STATE)
    v = ReadTable(ws, 2)
    n = RowCount(v)
    If n = 0 Then Exit Sub
    ReDim out(1 To n, 1 To 2)
    m = 0
    For r = 1 To n
        If Left(ToText(v(r, 1)), Len(prefix)) <> prefix Then
            m = m + 1
            out(m, 1) = v(r, 1)
            out(m, 2) = v(r, 2)
        End If
    Next r
    For r = m + 1 To n
        out(r, 1) = ""
        out(r, 2) = ""
    Next r
    WriteTable ws, out, 2
End Sub
