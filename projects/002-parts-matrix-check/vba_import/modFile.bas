Attribute VB_Name = "modFile"
Option Explicit
' ============================================================
' modFile : ファイルを開く、名前を作る
'   元のファイルは読み取り専用で開く。保存しない。相手のマクロは動かさない。
' ============================================================

' すでに開いているブックを、場所と名前で探す。無ければ Nothing
Public Function FindOpenWorkbook(ByVal path As String) As Workbook
    Dim wb As Workbook
    Set FindOpenWorkbook = Nothing
    For Each wb In Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

' 読み取り専用で開く。すでに開いていればそれを返す。開けなかったら Nothing
Public Function OpenReadOnly(ByVal path As String) As Workbook
    Dim sec As Long, hasSec As Boolean, wb As Workbook
    Set wb = FindOpenWorkbook(path)
    If Not wb Is Nothing Then
        Set OpenReadOnly = wb
        Exit Function
    End If
    ' 開く相手のマクロを無効にする (3 = msoAutomationSecurityForceDisable)
    hasSec = False
    On Error Resume Next
    sec = Application.AutomationSecurity
    If Err.Number = 0 Then
        Application.AutomationSecurity = 3
        hasSec = (Err.Number = 0)
    End If
    Err.Clear
    Set wb = Nothing
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    Err.Clear
    If hasSec Then Application.AutomationSecurity = sec
    Err.Clear
    On Error GoTo 0
    Set OpenReadOnly = wb
End Function

' ファイルを選ぶ画面を出す。やめたら ""
Public Function PickFile(ByVal title As String) As String
    Dim v As Variant
    v = Application.GetOpenFilename("Excel のファイル (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", , title)
    If VarType(v) = vbBoolean Then
        PickFile = ""
    Else
        PickFile = CStr(v)
    End If
End Function

' ファイルの更新日時と大きさを文字にする (記録用)
Public Function FileStamp(ByVal path As String) As String
    Dim t As String
    t = ""
    On Error Resume Next
    t = StampText(FileDateTime(path)) & " / " & CStr(FileLen(path)) & " バイト"
    On Error GoTo 0
    FileStamp = t
End Function

' ファイルの名前に使えない文字を _ に変える
Public Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long, t As String
    t = s
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, bad(i), "_")
    Next i
    If Len(t) = 0 Then t = "noname"
    SafeFileName = t
End Function

' フォルダが無ければ作る
Public Sub EnsureFolder(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

' 場所と名前から、名前だけを取り出す
Public Function BaseName(ByVal path As String) As String
    Dim p As Long, q As Long
    p = InStrRev(path, "\")
    q = InStrRev(path, "/")
    If q > p Then p = q
    BaseName = Mid(path, p + 1)
End Function
