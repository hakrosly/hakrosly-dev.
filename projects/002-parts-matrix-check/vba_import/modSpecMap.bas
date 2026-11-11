Attribute VB_Name = "modSpecMap"
Option Explicit
' ============================================================
' modSpecMap : 仕様の呼び名を合わせる (仕様対応表)
'   マトリックスの「見出しの文字」と、部品表の「仕様コード」の対応を持つ。
'   1 つの見出しが複数の仕様コードに当たるときは、行を複数書く。
'   未設定の見出しが残っている間は、照合を始めない (黙って飛ばさないため)。
' ============================================================

' 読み取ったマトリックスの見出しを、仕様対応表に足す (その設定名でまだ無いものだけ)。
' 仕様コードが空の行は、部品表の仕様コードと同じ文字なら自動でうめる。
'   headings : 見出しの文字を改行で区切ったもの
'   bom      : 読取_部品表 の配列 (無ければ Empty。そのときは自動でうめない)
Public Sub SyncSpecMap(ByVal layName As String, ByVal headings As String, ByVal bom As Variant)
    Dim ws As Worksheet, old As Variant, out() As Variant
    Dim heads As Variant, have As Object, specs As Object
    Dim nOld As Long, n As Long, i As Long, j As Long, k As String, h As String

    Set ws = ToolSheet(SH_MAP)
    old = ReadTable(ws, PC_N)
    nOld = RowCount(old)
    If Len(headings) = 0 Then
        heads = Array()
    Else
        heads = Split(headings, vbLf)
    End If

    ' 部品表の仕様コード (そろえた形 → 元の文字)
    Set specs = NewDict()
    For i = 1 To RowCount(bom)
        k = ToText(bom(i, BC_SPECKEY))
        If Len(k) > 0 Then
            If Not specs.Exists(k) Then specs.Add k, ToText(bom(i, BC_SPEC))
        End If
    Next i

    ReDim out(1 To nOld + UBound(heads) - LBound(heads) + 2, 1 To PC_N)
    Set have = NewDict()
    n = 0
    For i = 1 To nOld
        n = n + 1
        For j = 1 To PC_N
            out(n, j) = ToText(old(i, j))
        Next j
        If Trim(out(n, PC_LAYOUT)) = layName Then
            If Not have.Exists(Trim(out(n, PC_HEAD))) Then have.Add Trim(out(n, PC_HEAD)), True
        End If
    Next i
    For i = LBound(heads) To UBound(heads)
        h = Trim(heads(i))
        If Len(h) > 0 Then
            If Not have.Exists(h) Then
                have.Add h, True
                n = n + 1
                out(n, PC_LAYOUT) = layName
                out(n, PC_HEAD) = h
                out(n, PC_SPEC) = ""
                out(n, PC_EXCL) = ""
                out(n, PC_MEMO) = ""
            End If
        End If
    Next i
    ' 自動でうめる
    For i = 1 To n
        If Trim(out(i, PC_LAYOUT)) = layName Then
            If Len(Trim(out(i, PC_SPEC))) = 0 And Len(Trim(out(i, PC_EXCL))) = 0 Then
                k = NormSpec(out(i, PC_HEAD))
                If specs.Exists(k) Then
                    out(i, PC_SPEC) = specs.Item(k)
                    out(i, PC_MEMO) = "自動でうめた"
                End If
            End If
        End If
    Next i
    For i = n + 1 To UBound(out, 1)
        For j = 1 To PC_N
            out(i, j) = ""
        Next j
    Next i
    WriteTable ws, out, PC_N
End Sub

' 仕様対応表から、その設定名の行を取り出す。
'   戻り値: (行, 4) の配列。1 = 見出しの文字、2 = 仕様コード (そろえた形)、
'           3 = 対象外なら True、4 = 仕様コード (元の文字)。無ければ Empty
Public Function LoadSpecMap(ByVal layName As String) As Variant
    Dim v As Variant, out() As Variant, n As Long, i As Long, m As Long
    v = ReadTable(ToolSheet(SH_MAP), PC_N)
    n = 0
    For i = 1 To RowCount(v)
        If Trim(ToText(v(i, PC_LAYOUT))) = layName And Len(Trim(ToText(v(i, PC_HEAD)))) > 0 Then n = n + 1
    Next i
    If n = 0 Then
        LoadSpecMap = Empty
        Exit Function
    End If
    ReDim out(1 To n, 1 To 4)
    m = 0
    For i = 1 To RowCount(v)
        If Trim(ToText(v(i, PC_LAYOUT))) = layName And Len(Trim(ToText(v(i, PC_HEAD)))) > 0 Then
            m = m + 1
            out(m, 1) = Trim(ToText(v(i, PC_HEAD)))
            out(m, 2) = NormSpec(ToText(v(i, PC_SPEC)))
            out(m, 3) = (Len(Trim(ToText(v(i, PC_EXCL)))) > 0)
            out(m, 4) = Trim(ToText(v(i, PC_SPEC)))
        End If
    Next i
    LoadSpecMap = out
End Function

' 対応が決まっていない見出しを、改行で区切って返す。
'   headings : マトリックスの見出しの文字を改行で区切ったもの
'   map      : LoadSpecMap の戻り値
Public Function UnmappedHeadings(ByVal headings As String, ByVal map As Variant) As String
    Dim heads As Variant, okSet As Object, i As Long, h As String, s As String
    Set okSet = NewDict()
    For i = 1 To RowCount(map)
        If map(i, 3) = True Or Len(map(i, 2)) > 0 Then
            If Not okSet.Exists(map(i, 1)) Then okSet.Add map(i, 1), True
        End If
    Next i
    s = ""
    If Len(headings) > 0 Then
        heads = Split(headings, vbLf)
        For i = LBound(heads) To UBound(heads)
            h = Trim(heads(i))
            If Len(h) > 0 Then
                If Not okSet.Exists(h) Then
                    If Len(s) > 0 Then s = s & vbLf
                    s = s & h
                End If
            End If
        Next i
    End If
    UnmappedHeadings = s
End Function
