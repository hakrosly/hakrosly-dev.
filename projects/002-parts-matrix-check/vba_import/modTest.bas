Attribute VB_Name = "modTest"
Option Explicit
' ============================================================
' modTest : 自己診断 (テスト)
'   わざと間違いを入れたデータを照合して、決まった数だけ見つかることを確かめる。
'   「ツールが本当にチェックできているか」を、だれでもボタン 1 つで確かめられるようにする。
'   結果はシート「テスト結果」に出す。
' ============================================================

Private mTest() As Variant    ' (列, 行)
Private mTestN As Long
Private mTestPass As Long
Private mTestCap As Long
Private mF As Variant          ' 最後の照合の、食い違いの一覧
Private mB As Variant          ' 同じく、部品表の消し込み
Private mC As Variant          ' 同じく、マトリックスの消し込み

Private Sub TReset()
    mTestCap = 200
    ReDim mTest(1 To 5, 1 To mTestCap)
    mTestN = 0
    mTestPass = 0
End Sub

Private Sub Chk(ByVal id As String, ByVal title As String, ByVal expected As String, ByVal actual As String)
    If mTestCap = 0 Then TReset
    If mTestN >= mTestCap Then
        mTestCap = mTestCap * 2
        ReDim Preserve mTest(1 To 5, 1 To mTestCap)
    End If
    mTestN = mTestN + 1
    mTest(1, mTestN) = id
    mTest(2, mTestN) = title
    mTest(3, mTestN) = expected
    mTest(4, mTestN) = actual
    If expected = actual Then
        mTest(5, mTestN) = "合格"
        mTestPass = mTestPass + 1
    Else
        mTest(5, mTestN) = "不合格"
    End If
End Sub

' 最後に終わった確かめの番号 (途中で止まったとき、どこまで進んだかを知るため)
Public Function TestProgress() As String
    If mTestN = 0 Then
        TestProgress = "(まだ 1 つも終わっていない)"
    Else
        TestProgress = CStr(mTest(1, mTestN))
    End If
End Function

Private Function CodeCounts(ByVal res As Object) As String
    ' 例: "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1" (0 件の記号は書かない)
    Dim codes As Variant, i As Long, j As Long, n As Long, s As String
    codes = CodeList()
    For i = LBound(codes) To UBound(codes)
        n = 0
        For j = 1 To res.Item("NFind")
            If mF(j, FC_CODE) = codes(i) Then n = n + 1
        Next j
        If n > 0 Then
            If Len(s) > 0 Then s = s & " "
            s = s & codes(i) & "=" & n
        End If
    Next i
    If Len(s) = 0 Then s = "(なし)"
    CodeCounts = s
End Function

' 読み取った中身をつないだ文字 (形が違っても中身が同じかを比べるため)。skipHead の見出しは除く
Private Function Content(ByRef mtx As Variant, ByVal skipHead As String) As String
    Dim i As Long, s As String
    For i = 1 To RowCount(mtx)
        If mtx(i, MC_HEAD) <> skipHead Then
            s = s & mtx(i, MC_KEY) & "|" & mtx(i, MC_NAME) & "|" & mtx(i, MC_HEAD) & "|" & mtx(i, MC_VAL) & vbLf
        End If
    Next i
    Content = s
End Function

Private Function ReadM(ByVal ws As Worksheet, ByVal lay As Object, ByVal info As Object) As Variant
    LogReset
    ReadM = ReadMatrixSheet(ws, lay, info)
End Function

' 読んで照合するところまでを 1 回で行う (テスト用)
Private Sub ReadAndCompare(ByVal ws As Worksheet, ByVal lay As Object, ByRef bom As Variant, ByRef map As Variant, _
        ByVal hasFilter As Boolean, ByVal cmpQty As Boolean, ByVal res As Object)
    Dim info As Object, mtx As Variant, warn As Variant
    Set info = NewDict()
    mtx = ReadM(ws, lay, info)
    warn = LogRows()
    CompareCore bom, mtx, map, warn, hasFilter, True, cmpQty, res, mF, mB, mC
End Sub

Private Function NewSheet(ByVal wb As Workbook) As Worksheet
    Set NewSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
End Function

Private Function CellColor(ByVal ws As Worksheet, ByVal addr As String) As String
    CellColor = CStr(ws.Range(addr).Interior.Color)
End Function

' ============================================================
' 自己診断を全部実行する。「合格 n / 全 m」の形の文を返す
' ============================================================
Public Function RunSelfTest() As String
    Dim wb As Workbook, ws As Worksheet, wsB As Worksheet
    Dim info As Object, lay As Object
    Dim bom As Variant, bomNoFilter As Variant, mtx As Variant, m1 As Variant, warn As Variant, map As Variant
    Dim res As Object, baseText As String, n As Long, i As Long
    Dim nTried As Long, nDetected As Long, detail As String
    Dim t As String, summary As String, k7 As Boolean, k8 As Boolean, k9 As Boolean

    TReset
    TextDefaults
    Set res = NewDict()

    ' ---- 文字のそろえ方 ----
    Chk "T-01", "全角の品番を半角の大文字にそろえる", "AB123-0010", NormPart(ToWide("ab123-0010"))
    Chk "T-02", "前後の空白と小文字をそろえる", "AB123-0010", NormPart(" ab123-0010 ")
    Chk "T-03", "ハイフンに似た文字を - にそろえる", "AB123-0010", NormPart("AB123" & ChrW(&H2010) & "0010")
    SetIgnoreHyphen True
    Chk "T-04", "「ハイフンを無視する」のとき - を取る", "AB1230010", NormPart("AB123-0010")
    SetIgnoreHyphen False
    Chk "T-05", "セルの中の改行を取る", "AB123-0010", NormPart("AB123-" & vbLf & "0010")
    Chk "T-06", "長音は、品番ではハイフンに、名称ではそのまま", "AB-1/" & NormBase("カバー"), NormPart("AB" & ChrW(&H30FC) & "1") & "/" & NormName("カバー")

    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    Set wsB = wb.Worksheets(1)

    ' ---- 部品表の読み取り ----
    FillTestBom wsB
    LogReset
    Set info = NewDict()
    bom = ReadBomSheet(wsB, TestBomConfig(True), info)
    n = 0
    For i = 1 To RowCount(bom)
        If Len(ToText(bom(i, BC_INSCOPE))) > 0 Then n = n + 1
    Next i
    Chk "T-10", "部品表を読む (行数 / 範囲の中 / 別の数え方 / 警告)", "12/11/12/0", _
        RowCount(bom) & "/" & n & "/" & info.Item("countA") & "/" & LogCount()
    Set info = NewDict()
    bomNoFilter = ReadBomSheet(wsB, TestBomConfig(False), info)

    ' ---- マトリックスの読み取り ----
    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M1"
    Set info = NewDict()
    m1 = ReadM(ws, TestLayout("M1", ws.Name), info)
    baseText = Content(m1, "")
    Chk "T-11", "M1 を読む (品番 / 仕様の列 / あり / 印のセル: 別の数え方 = 使った数)", "5/4/11/11=11", _
        info.Item("parts") & "/" & info.Item("cols") & "/" & info.Item("yes") & "/" & info.Item("rawYes") & "=" & info.Item("usedYes")
    Chk "T-11b", "M1 の見出しの文字 (結合された上の段を読めている)", "STD / 2WD;STD / 4WD;DX / 2WD;DX / 4WD", _
        Replace(info.Item("headings"), vbLf, ";")
    ' M1 の照合
    map = TestMap("")
    warn = LogRows()
    CompareCore bom, m1, map, warn, True, True, False, res, mF, mB, mC
    Chk "T-20", "M1 の照合: 入れておいた 5 つの食い違いが見つかる", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1", CodeCounts(res)
    Chk "T-20b", "M1 の照合: 消し込みの数 (セル: 一致/不一致/判断できない/対象外、行: 同じ)", "17/3/0/0 9/2/0/1", _
        res.Item("MOk") & "/" & res.Item("MNg") & "/" & res.Item("MUnk") & "/" & res.Item("MOut") & " " & res.Item("BOk") & "/" & res.Item("BNg") & "/" & res.Item("BUnk") & "/" & res.Item("BOut")
    k7 = (CLng(res.Item("BPairsT")) = CLng(res.Item("BPairsOk")) + CLng(res.Item("BPairsUnk")) + CLng(res.Item("FN4")))
    k8 = (CLng(res.Item("MPairsYes")) = CLng(res.Item("MPairsOk")) + CLng(res.Item("FN5")))
    k9 = (CLng(res.Item("BPairsOk")) = CLng(res.Item("MPairsOk")))
    Chk "T-20c", "M1 の照合: 突き合わせの数が合う (K-7 / K-8 / K-9)", "True/True/True", k7 & "/" & k8 & "/" & k9
    ' 抜き打ち
    MutationCheck bom, m1, map, warn, True, True, False, res, nTried, nDetected, detail
    Chk "T-38", "抜き打ち: 値を逆にした所は、全部気付く", "10/10", nTried & "/" & nDetected
    Chk "T-38b", "抜き打ちのあと、データは元に戻っている", baseText, Content(m1, "")
    ' 絞り込みなし
    CompareCore bomNoFilter, m1, map, warn, False, True, False, res, mF, mB, mC
    Chk "T-26", "絞り込みが無いとき: N-1 は調べない", "N-2=1 N-3=1 N-4=1 N-5=1/True", CodeCounts(res) & "/" & res.Item("SkippedN1")

    ' M2: 品番が横
    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M2"
    Set info = NewDict()
    mtx = ReadM(ws, TestLayout("M2", ws.Name), info)
    Chk "T-12", "M2 (品番が横) を読む: 中身が M1 と同じ", baseText, Content(mtx, "")
    ReadAndCompare ws, TestLayout("M2", ws.Name), bom, map, True, False, res
    Chk "T-21", "M2 の照合: M1 と同じ結果", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1", CodeCounts(res)

    ' M3: 値の結合、空行、小見出し、備考の列
    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M3"
    Set info = NewDict()
    mtx = ReadM(ws, TestLayout("M3", ws.Name), info)
    Chk "T-13", "M3 (値の結合・空行・小見出し) を読む: 備考を除いた中身が M1 と同じ", baseText, Content(mtx, "備考")
    Chk "T-13b", "M3: 品番 / 仕様の列 / あり / 飛ばした行 / 印のセル (結合は 1 個と数える)", "5/5/11/2/7=7", _
        info.Item("parts") & "/" & info.Item("cols") & "/" & info.Item("yes") & "/" & info.Item("blankRows") & "/" & info.Item("rawYes") & "=" & info.Item("usedYes")
    Chk "T-23", "M3: 備考の対応が未設定なら、未設定として知らせる", "備考", UnmappedHeadings(info.Item("headings"), TestMap("未設定"))
    map = TestMap(MAP_EXCLUDE)
    ReadAndCompare ws, TestLayout("M3", ws.Name), bom, map, True, False, res
    Chk "T-22", "M3 の照合 (備考は対象外): M1 と同じ結果。備考の文字は警告にならない", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1/備考", _
        CodeCounts(res) & "/" & res.Item("HeadsExcluded")
    ' 読まれていない印に気付くか (品番が空の行に印を置く)
    ws.Range("C8").Value = ChrW(&H25CB)
    Set info = NewDict()
    mtx = ReadM(ws, TestLayout("M3", ws.Name), info)
    Chk "T-39", "検算: 品番が空の行に印があると、読まれていない印として気付く", "8/7/C8", _
        info.Item("rawYes") & "/" & info.Item("usedYes") & "/" & info.Item("unread")
    map = TestMap("")

    ' M4: 結合なしの空欄
    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M4"
    Set info = NewDict()
    mtx = ReadM(ws, TestLayout("M4", ws.Name), info)
    Chk "T-14", "M4 (結合せず空欄) を、引き継ぎ「する」で読む: 中身が M1 と同じ", baseText, Content(mtx, "")
    ReadAndCompare ws, TestLayout("M4", ws.Name), bom, map, True, False, res
    Chk "T-24", "M4 の照合: M1 と同じ結果", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1", CodeCounts(res)
    Set lay = NewLayout("TEST_M4", ws.Name, O_VERT, "B5:B9", "A5:A9", "C3:F4", K_MARK, OPT_NO, OPT_NOSPLIT, "")
    Set info = NewDict()
    mtx = ReadM(ws, lay, info)
    Chk "T-15", "M4 を、引き継ぎ「しない」で読む: 空欄は引き継がれない", "STD / 2WD;4WD;DX / 2WD", Replace(info.Item("headings"), vbLf, ";")

    ' M5: 員数
    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M5"
    ReadAndCompare ws, TestLayout("M5", ws.Name), bom, map, True, True, res
    Chk "T-25", "M5 (員数) の照合: 員数の違いが 1 件増える", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1 N-6=1", CodeCounts(res)

    ' OK: 食い違いなし
    Set ws = NewSheet(wb)
    FillTestMatrix ws, "OK"
    ReadAndCompare ws, TestLayout("OK", ws.Name), bom, map, True, False, res
    Chk "T-27", "食い違いを直したマトリックス: 不一致 0、警告 0", "(なし)/0/0", CodeCounts(res) & "/" & res.Item("NNg") & "/" & res.Item("NWarn")

    ' ---- 警告 (M1 を 1 か所だけ変える) ----
    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M1"
    ws.Range("A10:F10").NumberFormat = "@"
    ws.Range("A10").Value = "ブラケットA1"
    ws.Range("B10").Value = "AB123-0010"
    Set lay = NewLayout("TEST_W", ws.Name, O_VERT, "B5:B10", "A5:A10", "C3:F4", K_MARK, OPT_YES, OPT_NOSPLIT, "")
    ReadAndCompare ws, lay, bom, map, True, False, res
    Chk "T-30", "同じ品番の行が 2 つ: 品番の重複 (W-1)", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1 W-1=1", CodeCounts(res)

    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M1"
    ws.Range("C5").Value = ChrW(&H25B3)
    ReadAndCompare ws, TestLayout("M1", ws.Name), bom, map, True, False, res
    Chk "T-31", "印の一覧に無い値: 判断できない値 (W-3)。あり にも なし にもしない", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1 W-3=1/1", _
        CodeCounts(res) & "/" & res.Item("MUnk")

    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M1"
    ws.Range("B9").NumberFormat = "General"
    ws.Range("B9").Value = 1230060
    ReadAndCompare ws, TestLayout("M1", ws.Name), bom, map, True, False, res
    Chk "T-32", "品番が数値: 警告 (W-4)", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1 W-4=1", CodeCounts(res)

    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M1"
    ws.Columns(5).Hidden = True
    ReadAndCompare ws, TestLayout("M1", ws.Name), bom, map, True, False, res
    Chk "T-33", "非表示の列: 読んだうえで警告 (W-5)", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1 W-5=1", CodeCounts(res)

    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M1"
    ws.Range("B6").Font.Strikethrough = True
    ReadAndCompare ws, TestLayout("M1", ws.Name), bom, map, True, False, res
    Chk "T-34", "取り消し線のある品番: 読んだうえで警告 (W-8)", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1 W-8=1", CodeCounts(res)

    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M1"
    ws.Range("B9").Value = "メモ"
    Set lay = NewLayout("TEST_W", ws.Name, O_VERT, "B5:B9", "A5:A9", "C3:F4", K_MARK, OPT_YES, OPT_NOSPLIT, "AB###-####")
    ReadAndCompare ws, lay, bom, map, True, False, res
    Chk "T-35", "品番の形に合わないセル: 読まずに警告 (W-9)", "N-1=1 N-3=1 N-4=1 N-5=1 W-9=1", CodeCounts(res)

    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M1"
    ws.Range("A10:F10").NumberFormat = "@"
    ws.Range("A10").Value = "カバー"
    ws.Range("B10").Value = "ZZ999-0001"
    ws.Range("C10").Value = ChrW(&H25CB)
    Set lay = NewLayout("TEST_W", ws.Name, O_VERT, "B5:B10", "A5:A10", "C3:F4", K_MARK, OPT_YES, OPT_NOSPLIT, "")
    ReadAndCompare ws, lay, bom, map, True, False, res
    Chk "T-36", "絞り込みの外の品番: 警告 (W-10)。N-2 にはしない", "N-1=1 N-2=1 N-3=1 N-4=1 N-5=1 W-10=1", CodeCounts(res)

    Set ws = NewSheet(wb)
    FillTestMatrix ws, "M1"
    Set lay = NewLayout("TEST_W", ws.Name, O_VERT, "B5:B8", "A5:A8", "C3:F4", K_MARK, OPT_YES, OPT_NOSPLIT, "")
    ReadAndCompare ws, lay, bom, map, True, False, res
    Chk "T-37", "範囲が 1 行足りない: 続きがあると知らせる (W-11)", "N-1=1 N-3=1 N-4=1 N-5=1 W-11=1", CodeCounts(res)

    ' 範囲の決まり
    Set lay = NewLayout("TEST_W", ws.Name, O_HORZ, "B5:B9", "A5:A9", "C3:F4", K_MARK, OPT_YES, OPT_NOSPLIT, "")
    Chk "T-42", "向きを間違えた設定は、読む前に断る", "True", CStr(Len(CheckLayout(ws, lay)) > 0)

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    ' ---- 通しの確かめ (ファイルを作り、読み、照合し、結果のブックを作る) ----
    If Len(ThisWorkbook.Path) > 0 Then
        EndToEnd
    Else
        Chk "T-50", "通しの確かめ", "実行", "ツール用ブックが保存されていないため、実行できなかった"
    End If

    ' ---- 後片付け ----
    CleanupTestRows
    summary = "合格 " & mTestPass & " / 全 " & mTestN
    WriteTable ToolSheet(SH_TEST), BufToRows(mTest, mTestN, 5), 5
    For i = 1 To mTestN
        If mTest(5, i) = "合格" Then
            ToolSheet(SH_TEST).Cells(i + 1, 5).Interior.Color = COL_OK
        Else
            ToolSheet(SH_TEST).Cells(i + 1, 5).Interior.Color = COL_NG
        End If
    Next i
    StateSet "test.result", StampText(Now) & " " & summary & " (版 " & TOOL_VERSION & ")"
    TextLoad
    RunSelfTest = summary
End Function

' 通しの確かめ。ボタンを押したときと同じ道筋 (modRun) を通す
Private Sub EndToEnd()
    Dim folder As String, sep As String, pBom As String, pMtx As String
    Dim wbB As Workbook, wbM As Workbook, outWb As Workbook, verdict As String
    Dim stB As String, stM As String, lay As Object, ws As Worksheet, msg As String, t As String

    sep = Application.PathSeparator
    folder = MakeTestFiles()
    pBom = folder & sep & "bom_test.xlsx"
    pMtx = folder & sep & "matrix_M3.xlsx"
    stB = FileStamp(pBom)
    stM = FileStamp(pMtx)
    Set wbB = OpenReadOnly(pBom)
    Set wbM = OpenReadOnly(pMtx)
    If wbB Is Nothing Or wbM Is Nothing Then
        Chk "T-50", "通しの確かめ", "実行", "テスト用のブックを開けなかった"
        Exit Sub
    End If
    msg = RunReadBom(wbB.Worksheets(1), TestBomConfig(True), pBom)
    Set ws = wbM.Worksheets(1)
    Set lay = TestLayout("M3", ws.Name)
    SaveLayout lay
    msg = RunReadMatrix(ws, lay, pMtx)
    ' 対応が未設定のままなら、照合は始まらない
    t = "始まった"
    On Error Resume Next
    msg = RunCompare(ws, False, outWb, verdict)
    If Err.Number = ERR_TOOL Then t = "始まらなかった"
    Err.Clear
    On Error GoTo 0
    Chk "T-51", "通し: 見出しの対応が未設定のままでは、照合が始まらない", "始まらなかった", t
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    Set outWb = Nothing
    ' 対応を書いて照合する
    WriteTestMapToSheet lay.Item("name"), MAP_EXCLUDE
    msg = RunCompare(ws, False, outWb, verdict)
    Chk "T-50", "通し: 判定と、検算で合わなかった項目", "不一致あり / ", verdict & " / " & ProofBadList()
    Chk "T-52", "通し: 結果のブックのシートの数", "8", CStr(outWb.Worksheets.Count)
    Chk "T-53", "通し: 写しの色 (一致 = 緑 / 印が無い = 赤 / 印が余分 = 赤 / 部品表に無い品番 = 赤 / 名称が違う = 赤 / 備考 = 灰 / 見出し = 水色)", _
        COL_OK & "/" & COL_NG & "/" & COL_NG & "/" & COL_NG & "/" & COL_NG & "/" & COL_OUT & "/" & COL_HEAD, _
        CellColor(outWb.Worksheets(R_MTX), "C6") & "/" & CellColor(outWb.Worksheets(R_MTX), "F7") & "/" & _
        CellColor(outWb.Worksheets(R_MTX), "C10") & "/" & CellColor(outWb.Worksheets(R_MTX), "B11") & "/" & _
        CellColor(outWb.Worksheets(R_MTX), "A6") & "/" & CellColor(outWb.Worksheets(R_MTX), "G7") & "/" & _
        CellColor(outWb.Worksheets(R_MTX), "C4")
    Chk "T-54", "通し: 品番が空の行のセルには色が付かない (見ていないセルが分かる)", "True", _
        CStr(CellColor(outWb.Worksheets(R_MTX), "C8") <> CStr(COL_OK) And CellColor(outWb.Worksheets(R_MTX), "C8") <> CStr(COL_NG))
    outWb.Close SaveChanges:=False
    wbB.Close SaveChanges:=False
    wbM.Close SaveChanges:=False
    Chk "T-40", "通し: 元のファイルは変わっていない (更新日時と大きさ)", stB & " ; " & stM, FileStamp(pBom) & " ; " & FileStamp(pMtx)
End Sub

' テストが書き込んだ行を消す (設定名が TEST_ で始まる行、読み取りの結果)
Private Sub CleanupTestRows()
    RemoveTestRows ToolSheet(SH_CFG_MTX), 12
    RemoveTestRows ToolSheet(SH_MAP), PC_N
    WriteTable ToolSheet(SH_RD_BOM), Empty, BC_N
    WriteTable ToolSheet(SH_RD_MTX), Empty, MC_N
    WriteTable ToolSheet(SH_RD_WARN), Empty, WC_N
    StateClear "bom."
    StateClear "mtx."
End Sub

Private Sub RemoveTestRows(ByVal ws As Worksheet, ByVal nCol As Long)
    Dim v As Variant, out() As Variant, n As Long, i As Long, j As Long, total As Long
    v = ReadTable(ws, nCol)
    total = RowCount(v)
    If total = 0 Then Exit Sub
    ReDim out(1 To total, 1 To nCol)
    n = 0
    For i = 1 To total
        If Left(Trim(ToText(v(i, 1))), 5) <> "TEST_" Then
            n = n + 1
            For j = 1 To nCol
                out(n, j) = ToText(v(i, j))
            Next j
        End If
    Next i
    For i = n + 1 To total
        For j = 1 To nCol
            out(i, j) = ""
        Next j
    Next i
    WriteTable ws, out, nCol
End Sub
