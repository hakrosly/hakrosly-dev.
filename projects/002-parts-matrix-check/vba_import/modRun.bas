Attribute VB_Name = "modRun"
Option Explicit
' ============================================================
' modRun : 読み取り → 照合 → 結果 の一連の流れ (画面は出さない)
'   段階の間は、ツール用ブックのシートで受け渡す。
'     読取_部品表 / 読取_マトリックス / 読取_警告 / 読取_記録 / 仕様対応表
'   画面 (ファイルを選ぶ、知らせる) は modMain が受け持つ。
' ============================================================

Private Function B01(ByVal b As Boolean) As String
    If b Then B01 = "1" Else B01 = "0"
End Function

' ① 部品表を読む。知らせる文を返す
Public Function RunReadBom(ByVal ws As Worksheet, ByVal cfg As Object, ByVal path As String) As String
    Dim info As Object, a As Variant
    TextLoad
    LogReset
    Set info = NewDict()
    a = ReadBomSheet(ws, cfg, info)
    WriteTable ToolSheet(SH_RD_BOM), a, BC_N
    LogSave SRC_BOM
    StateClear "bom."
    StateSet "bom.path", path
    StateSet "bom.sheet", ws.Name
    StateSet "bom.stamp", FileStamp(path)
    StateSet "bom.dataRows", info.Item("dataRows")
    StateSet "bom.readRows", info.Item("readRows")
    StateSet "bom.blankRows", info.Item("blankRows")
    StateSet "bom.errRows", info.Item("errRows")
    StateSet "bom.outRows", info.Item("outRows")
    StateSet "bom.countA", info.Item("countA")
    StateSet "bom.hasFilter", B01(info.Item("hasFilter"))
    StateSet "bom.hasName", B01(info.Item("hasName"))
    StateSet "bom.hasQty", B01(info.Item("hasQty"))
    RunReadBom = "部品表を読みました。" & vbLf & _
        "読んだ行: " & info.Item("readRows") & " 行 (空で飛ばした行: " & info.Item("blankRows") & " 行)" & vbLf & _
        "共通の形に直した行数: " & info.Item("outRows") & " 行" & vbLf & _
        "警告: " & LogCount() & " 件" & vbLf & vbLf & _
        "シート「" & SH_RD_BOM & "」で、正しく読めたかを確かめてください。"
End Function

' ② マトリックスを読む。知らせる文を返す
Public Function RunReadMatrix(ByVal ws As Worksheet, ByVal lay As Object, ByVal path As String) As String
    Dim info As Object, a As Variant, t As String
    TextLoad
    LogReset
    Set info = NewDict()
    a = ReadMatrixSheet(ws, lay, info)
    WriteTable ToolSheet(SH_RD_MTX), a, MC_N
    LogSave SRC_MTX
    StateClear "mtx."
    StateSet "mtx.path", path
    StateSet "mtx.sheet", ws.Name
    StateSet "mtx.layout", lay.Item("name")
    StateSet "mtx.stamp", FileStamp(path)
    StateSet "mtx.parts", info.Item("parts")
    StateSet "mtx.cols", info.Item("cols")
    StateSet "mtx.cells", info.Item("cells")
    StateSet "mtx.yes", info.Item("yes")
    StateSet "mtx.blankRows", info.Item("blankRows")
    StateSet "mtx.rawYes", info.Item("rawYes")
    StateSet "mtx.usedYes", info.Item("usedYes")
    StateSet "mtx.unreadCount", info.Item("unreadCount")
    StateSet "mtx.unread", info.Item("unread")
    StateSet "mtx.headings", info.Item("headings")
    StateSet "mtx.hasName", B01(info.Item("hasName"))
    StateSet "mtx.asQty", B01(info.Item("asQty"))
    ' 見出しを仕様対応表に足し、うめられる所は自動でうめる
    SyncSpecMap lay.Item("name"), info.Item("headings"), ReadTable(ToolSheet(SH_RD_BOM), BC_N)
    t = "マトリックスを読みました。" & vbLf & _
        "品番: " & info.Item("parts") & " 件 / 仕様の列: " & info.Item("cols") & " 列 / 「あり」: " & info.Item("yes") & " 件" & vbLf & _
        "警告: " & LogCount() & " 件" & vbLf
    If info.Item("unreadCount") > 0 Then
        t = t & vbLf & "注意: 読まれていない印があります (" & info.Item("unread") & ")。範囲の指定を確かめてください。" & vbLf
    End If
    RunReadMatrix = t & vbLf & "シート「" & SH_RD_MTX & "」で、正しく読めたかを確かめてください。" & vbLf & _
        "次に、シート「" & SH_MAP & "」で、見出しと仕様コードの対応を確かめてください。"
End Function

' ③ 照合する。結果のブックを作り、知らせる文を返す
'   srcWs  : マトリックスの元のシート (写しを作るため)。Nothing でもよい
'   doSave : 結果のブックを保存するか
'   outWb  : 作った結果のブック
'   verdict: 判定
Public Function RunCompare(ByVal srcWs As Worksheet, ByVal doSave As Boolean, ByRef outWb As Workbook, ByRef verdict As String) As String
    Dim bom As Variant, mtx As Variant, warn As Variant, map As Variant
    Dim layName As String, lay As Object, un As String
    Dim hasFilter As Boolean, cmpName As Boolean, cmpQty As Boolean
    Dim findings As Variant, bomCheck As Variant, cellCheck As Variant
    Dim res As Object, nTried As Long, nDetected As Long, detail As String, saved As String, t As String

    TextLoad
    bom = ReadTable(ToolSheet(SH_RD_BOM), BC_N)
    mtx = ReadTable(ToolSheet(SH_RD_MTX), MC_N)
    warn = ReadTable(ToolSheet(SH_RD_WARN), WC_N)
    If RowCount(bom) = 0 Then Err.Raise ERR_TOOL, "RunCompare", "部品表がまだ読まれていません。先に「① 部品表を読む」を実行してください。"
    If RowCount(mtx) = 0 Then Err.Raise ERR_TOOL, "RunCompare", "マトリックスがまだ読まれていません。先に「② マトリックスを読む」を実行してください。"
    layName = StateGet("mtx.layout")
    Set lay = LoadLayout(layName)
    If lay Is Nothing Then Err.Raise ERR_TOOL, "RunCompare", "設定「" & layName & "」がシート「" & SH_CFG_MTX & "」に見つかりません。"

    SyncSpecMap layName, StateGet("mtx.headings"), bom
    map = LoadSpecMap(layName)
    un = UnmappedHeadings(StateGet("mtx.headings"), map)
    If Len(un) > 0 Then
        Err.Raise ERR_TOOL, "RunCompare", "シート「" & SH_MAP & "」に、対応が決まっていない見出しがあります。" & vbLf & _
            "仕様コードを書くか、比べない列なら「対象外」の列に「" & MAP_EXCLUDE & "」と書いてください。" & vbLf & vbLf & un
    End If

    hasFilter = (StateGet("bom.hasFilter") = "1")
    cmpName = (StateGet("bom.hasName") = "1" And StateGet("mtx.hasName") = "1")
    cmpQty = (StateGet("bom.hasQty") = "1" And StateGet("mtx.asQty") = "1")
    Set res = NewDict()
    CompareCore bom, mtx, map, warn, hasFilter, cmpName, cmpQty, res, findings, bomCheck, cellCheck
    MutationCheck bom, mtx, map, warn, hasFilter, cmpName, cmpQty, res, nTried, nDetected, detail
    Set outWb = BuildResultBook(res, findings, bomCheck, cellCheck, bom, mtx, map, lay, srcWs, nTried, nDetected, detail, verdict)
    saved = ""
    If doSave Then saved = SaveResultBook(outWb, layName)
    t = "判定: " & verdict & vbLf & _
        "不一致: " & res.Item("NNg") & " 件 / 警告: " & res.Item("NWarn") & " 件" & vbLf
    If Len(ProofBadList()) > 0 Then
        t = t & vbLf & "検算で合わなかった項目: " & ProofBadList() & vbLf & _
            "読み取りの範囲や設定を確かめてください。くわしくは結果のシート「" & R_PROOF & "」にあります。" & vbLf
    End If
    If Len(saved) > 0 Then t = t & vbLf & "結果を保存しました:" & vbLf & saved
    RunCompare = t
End Function
