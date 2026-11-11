Option Explicit
' ============================================================
' modMain : ボタンから呼ばれる入口 (画面を出す所)
'   ここは Excel の画面を使うので、LibreOffice では試していない。
'   中身の処理は modRun にある。
' ============================================================

Private Const APP_TITLE As String = "部品表とマトリックスの照合"

' エラーを知らせる。Err の中身は、他の処理を呼ぶ前に引数で受け取る
Private Sub ShowError(ByVal errNo As Long, ByVal errText As String, ByVal errSource As String)
    Dim msg As String
    If errNo = ERR_TOOL Then
        msg = errText
    Else
        msg = "予想していないエラーが起きました。" & vbLf & "番号: " & errNo & vbLf & "内容: " & errText & vbLf & "場所: " & errSource
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox msg, vbExclamation, APP_TITLE
End Sub

' 範囲をマウスで選んでもらう。やめたら Nothing
Private Function AskRange(ByVal prompt As String) As Range
    Set AskRange = Nothing
    On Error Resume Next
    Set AskRange = Application.InputBox(prompt:=prompt, Title:=APP_TITLE, Type:=8)
    On Error GoTo 0
End Function

' シートを作り直す (最初に 1 回)
Public Sub Btn_Setup()
    On Error GoTo EH
    SetupTool
    MsgBox "シートを用意しました。", vbInformation, APP_TITLE
    Exit Sub
EH:
    ShowError Err.Number, Err.Description, Err.Source
End Sub

' ① 部品表を読む
Public Sub Btn_ReadBom()
    Dim path As String, wb As Workbook, ws As Worksheet, cfg As Object, msg As String
    On Error GoTo EH
    path = PickFile("部品表のファイルを選んでください")
    If Len(path) = 0 Then Exit Sub
    Set wb = OpenReadOnly(path)
    If wb Is Nothing Then Err.Raise ERR_TOOL, "Btn_ReadBom", "ファイルを開けませんでした。" & vbLf & path
    Set cfg = LoadBomConfig()
    If Len(cfg.Item("sheet")) = 0 Then
        Set ws = wb.Worksheets(1)
    Else
        Set ws = SheetOrNothing(wb, cfg.Item("sheet"))
        If ws Is Nothing Then Err.Raise ERR_TOOL, "Btn_ReadBom", "部品表のファイルに、シート「" & cfg.Item("sheet") & "」がありません。シート「" & SH_CFG_BOM & "」の「シート名」を確かめてください。"
    End If
    Application.ScreenUpdating = False
    msg = RunReadBom(ws, cfg, path)
    Application.ScreenUpdating = True
    ThisWorkbook.Activate
    ToolSheet(SH_RD_BOM).Activate
    MsgBox msg, vbInformation, APP_TITLE
    Exit Sub
EH:
    ShowError Err.Number, Err.Description, Err.Source
End Sub

' マトリックスの形を、マウスで教えてもらう。できた設定を返す。やめたら Nothing
Private Function TeachLayout(ByVal wb As Workbook, ByVal layName As String) As Object
    Dim r1 As Range, r2 As Range, r3 As Range, ws As Worksheet
    Dim orient As String, kind As String, nameAddr As String, lay As Object, msg As String
    Set TeachLayout = Nothing
    wb.Activate
    Set r1 = AskRange("① 品番が並んでいるセルを、マウスで選んでください。" & vbLf & "(見出しは含めないでください)")
    If r1 Is Nothing Then Exit Function
    Set ws = r1.Worksheet
    Set r2 = AskRange("② 部品名称が並んでいるセルを、マウスで選んでください。" & vbLf & "(品番と同じ長さで。名称を比べないときは「キャンセル」)")
    Set r3 = AskRange("③ 仕様の見出しのセルを、マウスで選んでください。" & vbLf & "(何段あっても全部。品番や名称の見出しは含めないでください)")
    If r3 Is Nothing Then Exit Function
    If r1.Rows.Count > r1.Columns.Count Then
        orient = O_VERT
    ElseIf r1.Rows.Count < r1.Columns.Count Then
        orient = O_HORZ
    ElseIf MsgBox("品番は、縦 (上から下) に並んでいますか?" & vbLf & "「いいえ」なら、横 (左から右) として読みます。", vbYesNo + vbQuestion, APP_TITLE) = vbYes Then
        orient = O_VERT
    Else
        orient = O_HORZ
    End If
    If MsgBox("値のセルに入っているのは、使う個数 (員数) ですか?" & vbLf & "「いいえ」なら、○ などの印として読みます。", vbYesNo + vbQuestion, APP_TITLE) = vbYes Then
        kind = K_QTY
    Else
        kind = K_MARK
    End If
    nameAddr = ""
    If Not r2 Is Nothing Then nameAddr = r2.Address(False, False)
    Set lay = NewLayout(layName, ws.Name, orient, r1.Address(False, False), nameAddr, r3.Address(False, False), _
                        kind, OPT_YES, OPT_NOSPLIT, "")
    msg = CheckLayout(ws, lay)
    If Len(msg) > 0 Then
        MsgBox msg & vbLf & vbLf & "設定は保存していません。もう一度やり直してください。", vbExclamation, APP_TITLE
        Exit Function
    End If
    SaveLayout lay
    Set TeachLayout = lay
End Function

' ② マトリックスを読む
Public Sub Btn_ReadMatrix()
    Dim path As String, wb As Workbook, ws As Worksheet, lay As Object
    Dim names As String, layName As String, msg As String
    On Error GoTo EH
    path = PickFile("マトリックスのファイルを選んでください")
    If Len(path) = 0 Then Exit Sub
    Set wb = OpenReadOnly(path)
    If wb Is Nothing Then Err.Raise ERR_TOOL, "Btn_ReadMatrix", "ファイルを開けませんでした。" & vbLf & path
    names = LayoutNames()
    If Len(names) = 0 Then names = "(まだありません)"
    layName = Trim(InputBox("使う設定の名前を入力してください。" & vbLf & _
        "初めての形のときは、新しい名前を付けてください (このあと、形をマウスで教えます)。" & vbLf & vbLf & _
        "今ある設定:" & vbLf & names, APP_TITLE, StateGet("mtx.layout")))
    If Len(layName) = 0 Then Exit Sub
    Set lay = LoadLayout(layName)
    Set ws = Nothing
    If Not lay Is Nothing Then
        Set ws = SheetOrNothing(wb, lay.Item("sheet"))
        If ws Is Nothing Then
            MsgBox "このファイルに、シート「" & lay.Item("sheet") & "」がありません。形を教え直してください。", vbExclamation, APP_TITLE
            Set lay = Nothing
        ElseIf LayoutMaybeGrown(ws, lay) Then
            If MsgBox("品番の範囲 (" & lay.Item("partRange") & ") のすぐ先に、まだデータがあるようです。" & vbLf & _
                "行や列が増えたのかもしれません。形を教え直しますか?", vbYesNo + vbQuestion, APP_TITLE) = vbYes Then
                Set lay = Nothing
            End If
        End If
    End If
    If lay Is Nothing Then
        Set lay = TeachLayout(wb, layName)
        If lay Is Nothing Then Exit Sub
        Set ws = SheetOrNothing(wb, lay.Item("sheet"))
    End If
    Application.ScreenUpdating = False
    msg = RunReadMatrix(ws, lay, path)
    Application.ScreenUpdating = True
    ThisWorkbook.Activate
    ToolSheet(SH_RD_MTX).Activate
    MsgBox msg, vbInformation, APP_TITLE
    Exit Sub
EH:
    ShowError Err.Number, Err.Description, Err.Source
End Sub

' 形を教え直す (今の設定を、マウスで選び直す)
Public Sub Btn_TeachAgain()
    Dim path As String, wb As Workbook, layName As String, lay As Object
    On Error GoTo EH
    layName = Trim(InputBox("教え直す設定の名前を入力してください。" & vbLf & vbLf & LayoutNames(), APP_TITLE, StateGet("mtx.layout")))
    If Len(layName) = 0 Then Exit Sub
    path = PickFile("マトリックスのファイルを選んでください")
    If Len(path) = 0 Then Exit Sub
    Set wb = OpenReadOnly(path)
    If wb Is Nothing Then Err.Raise ERR_TOOL, "Btn_TeachAgain", "ファイルを開けませんでした。" & vbLf & path
    Set lay = TeachLayout(wb, layName)
    ThisWorkbook.Activate
    If Not lay Is Nothing Then MsgBox "設定「" & layName & "」を保存しました。「② マトリックスを読む」で読み直してください。", vbInformation, APP_TITLE
    Exit Sub
EH:
    ShowError Err.Number, Err.Description, Err.Source
End Sub

' ③ 照合する
Public Sub Btn_Compare()
    Dim path As String, wb As Workbook, srcWs As Worksheet, outWb As Workbook, verdict As String, msg As String
    Dim errNo As Long, errText As String, errSource As String
    On Error GoTo EH
    ' マトリックスの元のシートを探す (色を付けた写しを作るため)
    Set srcWs = Nothing
    path = StateGet("mtx.path")
    If Len(path) > 0 Then
        Set wb = FindOpenWorkbook(path)
        If wb Is Nothing Then
            If Len(Dir(path)) > 0 Then Set wb = OpenReadOnly(path)
        End If
        If Not wb Is Nothing Then Set srcWs = SheetOrNothing(wb, StateGet("mtx.sheet"))
    End If
    Application.ScreenUpdating = False
    msg = RunCompare(srcWs, True, outWb, verdict)
    Application.ScreenUpdating = True
    If Not outWb Is Nothing Then outWb.Activate
    MsgBox msg, vbInformation, APP_TITLE
    Exit Sub
EH:
    errNo = Err.Number: errText = Err.Description: errSource = Err.Source
    On Error Resume Next
    If errNo = ERR_TOOL And InStr(errText, SH_MAP) > 0 Then
        Application.ScreenUpdating = True
        ThisWorkbook.Activate
        ToolSheet(SH_MAP).Activate
    End If
    On Error GoTo 0
    ShowError errNo, errText, errSource
End Sub

' 自己診断 (テスト) を実行する
Public Sub Btn_SelfTest()
    Dim msg As String
    On Error GoTo EH
    If MsgBox("自己診断を実行します。" & vbLf & "読み取り済みのデータは消えます (設定は消えません)。よろしいですか?", vbOKCancel + vbQuestion, APP_TITLE) <> vbOK Then Exit Sub
    Application.ScreenUpdating = False
    msg = RunSelfTest()
    Application.ScreenUpdating = True
    ThisWorkbook.Activate
    ToolSheet(SH_TEST).Activate
    MsgBox "自己診断の結果: " & msg & vbLf & "くわしくはシート「" & SH_TEST & "」を見てください。", vbInformation, APP_TITLE
    Exit Sub
EH:
    ShowError Err.Number, Err.Description, Err.Source
End Sub

' テスト用のブックを作る (操作の練習用)
Public Sub Btn_MakeTestFiles()
    Dim folder As String
    On Error GoTo EH
    Application.ScreenUpdating = False
    folder = MakeTestFiles()
    Application.ScreenUpdating = True
    ThisWorkbook.Activate
    MsgBox "テスト用のブックを作りました。" & vbLf & folder, vbInformation, APP_TITLE
    Exit Sub
EH:
    ShowError Err.Number, Err.Description, Err.Source
End Sub
