Attribute VB_Name = "modLog"
Option Explicit
' ============================================================
' modLog : 読み取りの途中で気付いたこと (警告) をためる所
'   黙って捨てないための仕組み。読めなかったもの、飛ばしたものは
'   ここにためて、シート「読取_警告」と結果ブックに必ず出す。
' ============================================================

Private mLog() As Variant     ' (列, 行) の向きでためる
Private mLogN As Long
Private mLogCap As Long

Public Sub LogReset()
    mLogCap = 64
    ReDim mLog(1 To WC_N, 1 To mLogCap)
    mLogN = 0
End Sub

Public Sub LogAdd(ByVal src As String, ByVal code As String, ByVal part As String, ByVal head As String, _
                  ByVal cellAddr As String, ByVal raw As String, ByVal note As String)
    If mLogCap = 0 Then LogReset
    If mLogN >= mLogCap Then
        mLogCap = mLogCap * 2
        ReDim Preserve mLog(1 To WC_N, 1 To mLogCap)
    End If
    mLogN = mLogN + 1
    mLog(WC_SRC, mLogN) = src
    mLog(WC_CODE, mLogN) = code
    mLog(WC_PART, mLogN) = part
    mLog(WC_HEAD, mLogN) = head
    mLog(WC_CELL, mLogN) = cellAddr
    mLog(WC_RAW, mLogN) = raw
    mLog(WC_NOTE, mLogN) = note
End Sub

Public Function LogCount() As Long
    LogCount = mLogN
End Function

' ためた警告を (行, 列) の配列で返す。無ければ Empty
Public Function LogRows() As Variant
    If mLogCap = 0 Then LogReset
    LogRows = BufToRows(mLog, mLogN, WC_N)
End Function

' シート「読取_警告」の、出どころが src の行を入れ替える
Public Sub LogSave(ByVal src As String)
    Dim ws As Worksheet, old As Variant, cur As Variant
    Dim out() As Variant, n As Long, i As Long, j As Long, total As Long
    Set ws = ToolSheet(SH_RD_WARN)
    old = ReadTable(ws, WC_N)
    cur = LogRows()
    total = RowCount(old) + RowCount(cur)
    If total = 0 Then
        WriteTable ws, Empty, WC_N
        Exit Sub
    End If
    ReDim out(1 To total, 1 To WC_N)
    n = 0
    For i = 1 To RowCount(old)
        If ToText(old(i, WC_SRC)) <> src Then
            n = n + 1
            For j = 1 To WC_N
                out(n, j) = old(i, j)
            Next j
        End If
    Next i
    For i = 1 To RowCount(cur)
        If ToText(cur(i, WC_SRC)) = src Then
            n = n + 1
            For j = 1 To WC_N
                out(n, j) = cur(i, j)
            Next j
        End If
    Next i
    For i = n + 1 To total
        For j = 1 To WC_N
            out(i, j) = ""
        Next j
    Next i
    WriteTable ws, out, WC_N
End Sub
