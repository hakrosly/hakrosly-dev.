Option Explicit
' ============================================================
' modReport : 結果を 1 つの Excel のブックにまとめる
'   紙に赤ペンで印を付けていた作業の代わりになる「記録」を作る。
'     まとめ                  : 判定、いつ・何を・どの設定で比べたか、件数
'     食い違い一覧            : 不一致と警告
'     チェック済み_マトリックス: マトリックスの写しに色を付けたもの (見たセルは全部色が付く)
'     チェック済み_セル一覧    : マトリックスのセルごとの結果
'     チェック済み_部品表      : 部品表の全部の行に状態を付けたもの
'     検算                    : 数が合っているかの確かめ
'     抜き取り確認            : 人が目で確かめるための見本
'     設定の控え              : 使った設定
' ============================================================

Private mSum() As Variant     ' まとめの行をためる (列, 行)
Private mSumN As Long
Private mSumCap As Long
Private mProof() As Variant   ' 検算の行をためる (列, 行)
Private mProofN As Long
Private mProofBad As Long     ' 合わなかった項目の数

Private Sub SumReset()
    mSumCap = 200
    ReDim mSum(1 To 2, 1 To mSumCap)
    mSumN = 0
End Sub

Private Sub SumAdd(ByVal label As String, ByVal value As String)
    If mSumN >= mSumCap Then
        mSumCap = mSumCap * 2
        ReDim Preserve mSum(1 To 2, 1 To mSumCap)
    End If
    mSumN = mSumN + 1
    mSum(1, mSumN) = label
    mSum(2, mSumN) = value
End Sub

Private Sub ProofReset()
    ReDim mProof(1 To 6, 1 To 40)
    mProofN = 0
    mProofBad = 0
End Sub

' 検算の 1 行を足す。左と右が同じなら「合う」
Private Sub ProofAdd(ByVal code As String, ByVal title As String, ByVal leftV As String, ByVal rightV As String, ByVal note As String)
    Dim ok As Boolean
    mProofN = mProofN + 1
    ok = (leftV = rightV)
    mProof(1, mProofN) = code
    mProof(2, mProofN) = title
    mProof(3, mProofN) = leftV
    mProof(4, mProofN) = rightV
    If ok Then
        mProof(5, mProofN) = "合う"
    Else
        mProof(5, mProofN) = "合わない"
        mProofBad = mProofBad + 1
    End If
    mProof(6, mProofN) = note
End Sub

' 確かめられなかった項目を足す (合わない扱いにはしない)
Private Sub ProofSkip(ByVal code As String, ByVal title As String, ByVal note As String)
    mProofN = mProofN + 1
    mProof(1, mProofN) = code
    mProof(2, mProofN) = title
    mProof(3, mProofN) = ""
    mProof(4, mProofN) = ""
    mProof(5, mProofN) = "確かめていない"
    mProof(6, mProofN) = note
End Sub

Private Function OneLine(ByVal s As String) As String
    If Len(s) = 0 Then
        OneLine = "(なし)"
    Else
        OneLine = Replace(s, vbLf, ", ")
    End If
End Function

Private Function CountCode(ByRef res As TResult, ByVal code As String) As Long
    Dim i As Long, n As Long
    n = 0
    For i = 1 To res.NFind
        If res.Findings(i, FC_CODE) = code Then n = n + 1
    Next i
    CountCode = n
End Function

Private Function AddSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName
    Set AddSheet = ws
End Function

Private Function StatusColor(ByVal st As String) As Long
    Select Case st
        Case S_OK: StatusColor = COL_OK
        Case S_NG: StatusColor = COL_NG
        Case S_UNK: StatusColor = COL_UNK
        Case Else: StatusColor = COL_OUT
    End Select
End Function

' 範囲に色を付ける。結合セルにかかっていれば、結合範囲ごと付ける
Private Sub PaintAddr(ByVal ws As Worksheet, ByVal addr As String, ByVal color As Long)
    Dim rng As Range, c As Range
    If Len(addr) = 0 Then Exit Sub
    Set rng = ws.Range(addr)
    If RangeHasMerge(rng) Then
        For Each c In rng.Cells
            If c.MergeCells Then
                c.MergeArea.Interior.Color = color
            Else
                c.Interior.Color = color
            End If
        Next c
    Else
        rng.Interior.Color = color
    End If
End Sub

' ============================================================
' 検算の表を作り、判定を決める
' ============================================================
Private Function MakeProof(ByRef res As TResult, ByVal nTried As Long, ByVal nDetected As Long) As String
    Dim t As String, now1 As String
    ProofReset
    ' ---- マトリックス ----
    ProofAdd "K-1", "マトリックス: 品番の数 × 仕様の列の数 = 読んだ値のセルの数", _
        CStr(CLng(StateNum("mtx.parts")) * CLng(StateNum("mtx.cols"))), CStr(res.MCells), _
        "読んだ品番 " & StateGet("mtx.parts") & " 件 × 仕様の列 " & StateGet("mtx.cols") & " 列"
    ProofAdd "K-2", "マトリックス: 読んだ値のセルの数 = 一致 + 不一致 + 判断できない + 対象外", _
        CStr(res.MCells), CStr(res.MOk + res.MNg + res.MUnk + res.MOut), _
        "状態の付いていないセルが無いことの確かめ"
    t = "品番の範囲 × 見出しの範囲 を端から数えた「あり」の印のセルの数と、読み取りで使った「あり」のセルの数"
    If CLng(StateNum("mtx.unreadCount")) > 0 Then t = t & "。読まれていない印: " & StateGet("mtx.unread")
    ProofAdd "K-3", "マトリックス: 別の数え方で数えた「あり」の印 = 読み取りで使った「あり」の印", _
        StateGet("mtx.rawYes"), StateGet("mtx.usedYes"), t
    ' ---- 部品表 ----
    If StateNum("bom.countA") < 0 Then
        ProofSkip "K-4", "部品表: 品番の列の空でないセルの数 = 読んだ行 + エラーで読まなかった行", "別の数え方が使えなかった"
    Else
        ProofAdd "K-4", "部品表: 品番の列の空でないセルの数 = 読んだ行 + エラーで読まなかった行", _
            StateGet("bom.countA"), CStr(CLng(StateNum("bom.readRows")) + CLng(StateNum("bom.errRows"))), _
            "左は Excel の関数 (COUNTA) で数えた数"
    End If
    ProofAdd "K-5", "部品表: データの行数 = 読んだ行 + 空で飛ばした行 + エラーで読まなかった行", _
        StateGet("bom.dataRows"), CStr(CLng(StateNum("bom.readRows")) + CLng(StateNum("bom.blankRows")) + CLng(StateNum("bom.errRows"))), ""
    ProofAdd "K-6", "部品表: 共通の形に直した行数 = 一致 + 不一致 + 判断できない + 対象外", _
        CStr(res.BRows), CStr(res.BOk + res.BNg + res.BUnk + res.BOut), "状態の付いていない行が無いことの確かめ"
    ' ---- 両方の突き合わせ ----
    ProofAdd "K-7", "突き合わせ: 部品表の「品番 × 仕様」の組 = 一致の組 + 判断できない組 + N-4 の件数", _
        CStr(res.BPairsT), CStr(res.BPairsOk + res.BPairsUnk + res.FN4), _
        "左は部品表の側から、N-4 はマトリックスの側から数えた数"
    ProofAdd "K-8", "突き合わせ: マトリックスの「あり」の組 = 一致の組 + N-5 の件数", _
        CStr(res.MPairsYes), CStr(res.MPairsOk + res.FN5), "品番が部品表にあるものだけを数える"
    ProofAdd "K-9", "突き合わせ: 一致の組を、部品表の側から数えた数 = マトリックスの側から数えた数", _
        CStr(res.BPairsOk), CStr(res.MPairsOk), ""
    ' ---- 抜き打ち ----
    If nTried = 0 Then
        ProofSkip "K-10", "抜き打ち: わざと値を逆にした数 = 照合が気付いた数", "試せるセルが無かった"
    Else
        ProofAdd "K-10", "抜き打ち: わざと値を逆にした数 = 照合が気付いた数", CStr(nTried), CStr(nDetected), _
            "本物のデータのうえで、値を 1 か所ずつ逆にして照合し直した (元のファイルは変えていない)"
    End If
    ' ---- 読んだ後にファイルが変わっていないか ----
    If Len(StateGet("bom.path")) > 0 Then
        now1 = FileStamp(StateGet("bom.path"))
        If Len(now1) = 0 Then
            ProofSkip "K-11", "部品表のファイル: 読んだときの更新日時と大きさ = 今の更新日時と大きさ", "ファイルが見つからない"
        Else
            ProofAdd "K-11", "部品表のファイル: 読んだときの更新日時と大きさ = 今の更新日時と大きさ", StateGet("bom.stamp"), now1, _
                "合わないときは、読んだ後にファイルが変わっている。読み直すこと"
        End If
    End If
    If Len(StateGet("mtx.path")) > 0 Then
        now1 = FileStamp(StateGet("mtx.path"))
        If Len(now1) = 0 Then
            ProofSkip "K-12", "マトリックスのファイル: 読んだときの更新日時と大きさ = 今の更新日時と大きさ", "ファイルが見つからない"
        Else
            ProofAdd "K-12", "マトリックスのファイル: 読んだときの更新日時と大きさ = 今の更新日時と大きさ", StateGet("mtx.stamp"), now1, _
                "合わないときは、読んだ後にファイルが変わっている。読み直すこと"
        End If
    End If
    ' ---- 判定 ----
    If mProofBad > 0 Then
        MakeProof = "検算が合わない (この結果は使わないこと)"
    ElseIf res.NNg > 0 Then
        MakeProof = "不一致あり"
    ElseIf res.NWarn > 0 Then
        MakeProof = "一致 (警告あり)"
    Else
        MakeProof = "一致"
    End If
End Function

' ============================================================
' 結果のブックを作る。作ったブックを返す (保存はしない。保存は SaveResultBook)
'   srcWs : マトリックスの元のシート。Nothing なら写しは作らない
' ============================================================
Public Function BuildResultBook(ByRef res As TResult, ByRef bom As Variant, ByRef mtx As Variant, ByRef map As Variant, _
        ByVal lay As Object, ByVal srcWs As Worksheet, ByVal nTried As Long, ByVal nDetected As Long, _
        ByVal mutDetail As String, ByRef verdict As String) As Workbook
    Dim wb As Workbook, ws As Worksheet, wsM As Worksheet
    Dim i As Long, paintNote As String, hasCopy As Boolean

    verdict = MakeProof(res, nTried, nDetected)

    Set wb = Workbooks.Add
    Application.DisplayAlerts = False
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    wb.Worksheets(1).Name = R_SUM

    ' ---- 食い違い一覧 ----
    Set ws = AddSheet(wb, R_LIST)
    ' ---- マトリックスの写し ----
    paintNote = ""
    hasCopy = False
    Set wsM = Nothing
    If Not srcWs Is Nothing Then
        Set wsM = CopyMatrixSheet(srcWs, wb, paintNote)
        If Not wsM Is Nothing Then
            hasCopy = True
            If Not PaintMatrix(wsM, mtx, res) Then
                paintNote = "マトリックスの写しに色を付けられなかった (シートが保護されているなど)。「" & R_CELLS & "」を見ること"
            End If
        End If
    Else
        paintNote = "マトリックスのファイルが開かれていなかったため、写しを作っていない。「" & R_CELLS & "」を見ること"
    End If
    WriteFindings ws, res, hasCopy
    WriteCells AddSheet(wb, R_CELLS), mtx, res, hasCopy
    WriteBomCheck AddSheet(wb, R_BOM), bom, res
    WriteProof AddSheet(wb, R_PROOF), mutDetail
    WriteSample AddSheet(wb, R_SAMPLE), bom, mtx, map, res, hasCopy
    WriteConfig AddSheet(wb, R_CFG), lay
    WriteSummary wb.Worksheets(R_SUM), res, verdict, paintNote, nTried, nDetected
    wb.Worksheets(R_SUM).Activate
    Set BuildResultBook = wb
End Function

' 結果のブックを、ツール用ブックと同じ場所の「照合結果」フォルダに保存する。保存した場所を返す
Public Function SaveResultBook(ByVal wb As Workbook, ByVal layName As String) As String
    Dim folder As String, path As String
    SaveResultBook = ""
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    folder = ThisWorkbook.Path & Application.PathSeparator & RESULT_FOLDER
    EnsureFolder folder
    path = folder & Application.PathSeparator & "照合結果_" & SafeFileName(layName) & "_" & StampFile(Now) & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=51      ' 51 = マクロなしの Excel ブック (.xlsx)
    Application.DisplayAlerts = True
    SaveResultBook = path
End Function

' ---- マトリックスの写しを作る ----
Private Function CopyMatrixSheet(ByVal srcWs As Worksheet, ByVal wb As Workbook, ByRef note As String) As Worksheet
    Dim ws As Worksheet, f As Range, a As Range
    Set CopyMatrixSheet = Nothing
    On Error GoTo Fail
    srcWs.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Set ws = wb.Worksheets(wb.Worksheets.Count)
    ws.Name = R_MTX
    On Error Resume Next
    ws.Visible = True
    ws.Unprotect
    ' 式を値に直す (元のファイルへのリンクを残さないため)
    Set f = Nothing
    Set f = ws.UsedRange.SpecialCells(-4123)      ' -4123 = 式の入ったセル
    If Not f Is Nothing Then
        For Each a In f.Areas
            a.Value = a.Value
        Next a
    End If
    ' 条件付き書式は色の邪魔になるので、写しからは消す
    ws.Cells.FormatConditions.Delete
    Err.Clear
    On Error GoTo 0
    Set CopyMatrixSheet = ws
    Exit Function
Fail:
    note = "マトリックスの写しを作れなかった (" & Err.Description & ")。「" & R_CELLS & "」を見ること"
    Set CopyMatrixSheet = Nothing
End Function

' ---- 写しに色を付ける (赤ペンの代わり)。できたら True ----
'   読んだ見出し = 水色、対象外 = 灰、一致 = 緑、判断できない = 黄、不一致 = 赤
'   色の付いていないセルは、ツールが見ていないセル
Private Function PaintMatrix(ByVal ws As Worksheet, ByRef mtx As Variant, ByRef res As TResult) As Boolean
    Dim nM As Long, i As Long, pass As Long, st As String, want As String
    Dim seen As Object, addr As String, key As String, excl As Object, heads As Variant, k As Long
    PaintMatrix = False
    nM = RowCount(mtx)
    If nM = 0 Then
        PaintMatrix = True
        Exit Function
    End If
    On Error GoTo Fail
    Set excl = NewDict()
    If Len(res.HeadsExcluded) > 0 Then
        heads = Split(res.HeadsExcluded, vbLf)
        For k = LBound(heads) To UBound(heads)
            If Not excl.Exists(heads(k)) Then excl.Add heads(k), True
        Next k
    End If
    ' 見出し
    Set seen = NewDict()
    For i = 1 To nM
        addr = ToText(mtx(i, MC_HEADCELL))
        If Not seen.Exists(addr) Then
            seen.Add addr, True
            If excl.Exists(ToText(mtx(i, MC_HEAD))) Then
                PaintAddr ws, addr, COL_OUT
            Else
                PaintAddr ws, addr, COL_HEAD
            End If
        End If
    Next i
    ' 値のセル。悪い状態が上に残るように、対象外 → 一致 → 判断できない → 不一致 の順に塗る
    For pass = 1 To 4
        Select Case pass
            Case 1: want = S_OUT
            Case 2: want = S_OK
            Case 3: want = S_UNK
            Case Else: want = S_NG
        End Select
        For i = 1 To nM
            st = res.CellStatus(i)
            If st = want Then PaintAddr ws, ToText(mtx(i, MC_VALCELL)), StatusColor(st)
        Next i
    Next pass
    ' 品番と名称のセル
    Set seen = NewDict()
    For i = 1 To nM
        addr = ToText(mtx(i, MC_PARTCELL))
        key = ToText(mtx(i, MC_KEY))
        If Not seen.Exists(addr & vbTab & key) Then
            seen.Add addr & vbTab & key, True
            If InStr(res.PartNg, vbTab & key & vbTab) > 0 Then
                PaintAddr ws, addr, COL_NG
            ElseIf InStr(res.PartWarn, vbTab & key & vbTab) > 0 Then
                PaintAddr ws, addr, COL_UNK
            Else
                PaintAddr ws, addr, COL_OK
            End If
            If Len(ToText(mtx(i, MC_NAMECELL))) > 0 And Not res.SkippedN3 Then
                If InStr(res.NameNg, vbTab & key & vbTab) > 0 Then
                    PaintAddr ws, ToText(mtx(i, MC_NAMECELL)), COL_NG
                ElseIf InStr(res.PartNg, vbTab & key & vbTab) > 0 Then
                    PaintAddr ws, ToText(mtx(i, MC_NAMECELL)), COL_OUT
                Else
                    PaintAddr ws, ToText(mtx(i, MC_NAMECELL)), COL_OK
                End If
            End If
        End If
    Next i
    PaintMatrix = True
    Exit Function
Fail:
    PaintMatrix = False
End Function

' セルの場所を、写しのシートへ飛ぶリンクにする (範囲が 1 つのときだけ)
Private Sub LinkToCopy(ByVal c As Range, ByVal addr As String)
    If Len(addr) = 0 Or InStr(addr, ",") > 0 Then Exit Sub
    c.NumberFormat = "General"
    c.Formula = "=HYPERLINK(""#'" & R_MTX & "'!" & addr & """,""" & addr & """)"
End Sub

' ---- 食い違い一覧 ----
Private Sub WriteFindings(ByVal ws As Worksheet, ByRef res As TResult, ByVal hasCopy As Boolean)
    Dim out() As Variant, i As Long, j As Long
    WriteHeader ws, Array("No", "区分", "記号", "名前", "品番", "名称 (部品表)", "名称 (マトリックス)", "仕様コード", _
        "見出しの文字", "部品表の値", "マトリックスの値", "マトリックスのセル", "説明")
    If res.NFind = 0 Then
        ws.Cells(2, 1).Value = "食い違いも警告もありません。"
        Exit Sub
    End If
    ReDim out(1 To res.NFind, 1 To FC_N + 1)
    For i = 1 To res.NFind
        out(i, 1) = i
        For j = 1 To FC_N
            out(i, j + 1) = res.Findings(i, j)
        Next j
    Next i
    With ws.Range(ws.Cells(2, 1), ws.Cells(res.NFind + 1, FC_N + 1))
        .NumberFormat = "@"
        .Value = out
    End With
    For i = 1 To res.NFind
        If res.Findings(i, FC_CLASS) = CLASS_NG Then
            ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 4)).Interior.Color = COL_NG
        Else
            ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 4)).Interior.Color = COL_UNK
        End If
        If hasCopy Then LinkToCopy ws.Cells(i + 1, FC_CELL + 1), CStr(res.Findings(i, FC_CELL))
    Next i
    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(res.NFind + 1, FC_N + 1)).AutoFilter
    ws.Columns("A:M").AutoFit
    On Error GoTo 0
End Sub

' ---- マトリックスのセルごとの結果 ----
Private Sub WriteCells(ByVal ws As Worksheet, ByRef mtx As Variant, ByRef res As TResult, ByVal hasCopy As Boolean)
    Dim out() As Variant, n As Long, i As Long, a As Long, st As String
    WriteHeader ws, Array("No", "品番", "名称", "見出しの文字", "ツールの読み", "員数", "セルに入っていた文字", _
        "品番のセル", "値のセル", "状態", "説明")
    n = RowCount(mtx)
    If n = 0 Then Exit Sub
    ReDim out(1 To n, 1 To 11)
    For i = 1 To n
        out(i, 1) = i
        out(i, 2) = mtx(i, MC_PART)
        out(i, 3) = mtx(i, MC_NAME)
        out(i, 4) = mtx(i, MC_HEAD)
        out(i, 5) = mtx(i, MC_VAL)
        out(i, 6) = mtx(i, MC_QTY)
        out(i, 7) = mtx(i, MC_RAW)
        out(i, 8) = mtx(i, MC_PARTCELL)
        out(i, 9) = mtx(i, MC_VALCELL)
        out(i, 10) = res.CellStatus(i)
        out(i, 11) = res.CellNote(i)
    Next i
    With ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 11))
        .NumberFormat = "@"
        .Value = out
    End With
    ColorRuns ws, 10, n, res.CellStatus
    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 11)).AutoFilter
    ws.Columns("A:K").AutoFit
    On Error GoTo 0
End Sub

' ---- 部品表の消し込み (全部の行に状態を付ける) ----
Private Sub WriteBomCheck(ByVal ws As Worksheet, ByRef bom As Variant, ByRef res As TResult)
    Dim out() As Variant, n As Long, i As Long
    WriteHeader ws, Array("部品表の行", "品番", "名称", "仕様コード", "員数", "範囲の中", "状態", "説明")
    n = RowCount(bom)
    If n = 0 Then Exit Sub
    ReDim out(1 To n, 1 To 8)
    For i = 1 To n
        out(i, 1) = bom(i, BC_ROW)
        out(i, 2) = bom(i, BC_PART)
        out(i, 3) = bom(i, BC_NAME)
        out(i, 4) = bom(i, BC_SPEC)
        out(i, 5) = bom(i, BC_QTY)
        out(i, 6) = bom(i, BC_INSCOPE)
        out(i, 7) = res.BomStatus(i)
        out(i, 8) = res.BomNote(i)
    Next i
    With ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 8))
        .NumberFormat = "@"
        .Value = out
    End With
    ColorRuns ws, 7, n, res.BomStatus
    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 8)).AutoFilter
    ws.Columns("A:H").AutoFit
    On Error GoTo 0
End Sub

' 状態の列に色を付ける。同じ状態が続く所は、まとめて 1 回で塗る
Private Sub ColorRuns(ByVal ws As Worksheet, ByVal col As Long, ByVal n As Long, ByRef statuses As Variant)
    Dim i As Long, a As Long
    a = 1
    For i = 2 To n + 1
        If i > n Then
            ws.Range(ws.Cells(a + 1, col), ws.Cells(n + 1, col)).Interior.Color = StatusColor(statuses(a))
        ElseIf statuses(i) <> statuses(a) Then
            ws.Range(ws.Cells(a + 1, col), ws.Cells(i, col)).Interior.Color = StatusColor(statuses(a))
            a = i
        End If
    Next i
End Sub

' ---- 検算 ----
Private Sub WriteProof(ByVal ws As Worksheet, ByVal mutDetail As String)
    Dim out As Variant, i As Long, lines As Variant, r As Long
    WriteHeader ws, Array("記号", "確かめること", "左の数", "右の数", "結果", "説明")
    out = BufToRows(mProof, mProofN, 6)
    With ws.Range(ws.Cells(2, 1), ws.Cells(mProofN + 1, 6))
        .NumberFormat = "@"
        .Value = out
    End With
    For i = 1 To mProofN
        If mProof(5, i) = "合う" Then
            ws.Cells(i + 1, 5).Interior.Color = COL_OK
        ElseIf mProof(5, i) = "合わない" Then
            ws.Cells(i + 1, 5).Interior.Color = COL_NG
        Else
            ws.Cells(i + 1, 5).Interior.Color = COL_UNK
        End If
    Next i
    r = mProofN + 3
    ws.Cells(r, 1).Value = "抜き打ちで値を逆にした所"
    ws.Cells(r, 1).Font.Bold = True
    If Len(mutDetail) > 0 Then
        lines = Split(mutDetail, vbLf)
        For i = LBound(lines) To UBound(lines)
            r = r + 1
            ws.Cells(r, 2).NumberFormat = "@"
            ws.Cells(r, 2).Value = lines(i)
        Next i
    Else
        ws.Cells(r + 1, 2).Value = "(試せるセルが無かった)"
    End If
    On Error Resume Next
    ws.Columns("A:F").AutoFit
    On Error GoTo 0
End Sub

' ---- 抜き取り確認 (人が目で確かめるための見本) ----
'   ツールが「一致」としたセルから、「あり」と「なし」を半分ずつ、まんべんなく選ぶ
Private Sub WriteSample(ByVal ws As Worksheet, ByRef bom As Variant, ByRef mtx As Variant, ByRef map As Variant, _
        ByRef res As TResult, ByVal hasCopy As Boolean)
    Dim nM As Long, i As Long, j As Long, k As Long, pass As Long, want As String
    Dim cand() As Long, nCand As Long, nPick As Long, idx As Long, r As Long
    Dim headSpecs As Object, bomRows As Object, h As String, specs As Variant, pk As String, rowsTxt As String

    WriteHeader ws, Array("No", "品番", "見出しの文字", "仕様コード", "ツールの読み", "マトリックスのセル", _
        "部品表の行", "人の確認 (合っていれば ○)", "メモ")
    ws.Cells(1, 11).Value = "ツールが「一致」としたセルの見本です。元の表と見比べて、ツールの読みが正しいかを確かめてください。"
    nM = RowCount(mtx)
    If nM = 0 Then Exit Sub
    ' 見出し → 仕様コード (元の文字)、品番キー + 仕様 → 部品表の行
    Set headSpecs = NewDict()
    For i = 1 To RowCount(map)
        If map(i, 3) <> True And Len(map(i, 2)) > 0 Then
            If headSpecs.Exists(map(i, 1)) Then
                DictSet headSpecs, map(i, 1), headSpecs.Item(map(i, 1)) & vbTab & map(i, 2)
            Else
                headSpecs.Add map(i, 1), map(i, 2)
            End If
        End If
    Next i
    Set bomRows = NewDict()
    For j = 1 To RowCount(bom)
        pk = ToText(bom(j, BC_KEY)) & vbTab & ToText(bom(j, BC_SPECKEY))
        If bomRows.Exists(pk) Then
            DictSet bomRows, pk, bomRows.Item(pk) & ", " & ToText(bom(j, BC_ROW))
        Else
            bomRows.Add pk, ToText(bom(j, BC_ROW))
        End If
    Next j
    r = 1
    ReDim cand(1 To nM)
    For pass = 1 To 2
        If pass = 1 Then want = V_YES Else want = V_NO
        nCand = 0
        For i = 1 To nM
            If res.CellStatus(i) = S_OK And ToText(mtx(i, MC_VAL)) = want Then
                nCand = nCand + 1
                cand(nCand) = i
            End If
        Next i
        nPick = MinL(SAMPLE_COUNT \ 2, nCand)
        For k = 1 To nPick
            idx = cand(1 + Int((k - 1) * nCand / nPick))
            h = ToText(mtx(idx, MC_HEAD))
            rowsTxt = ""
            If headSpecs.Exists(h) Then
                specs = Split(headSpecs.Item(h), vbTab)
                For j = LBound(specs) To UBound(specs)
                    pk = ToText(mtx(idx, MC_KEY)) & vbTab & specs(j)
                    If Len(rowsTxt) > 0 Then rowsTxt = rowsTxt & " / "
                    If bomRows.Exists(pk) Then
                        rowsTxt = rowsTxt & bomRows.Item(pk) & " 行目"
                    Else
                        rowsTxt = rowsTxt & "登録なし"
                    End If
                Next j
            End If
            r = r + 1
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).NumberFormat = "@"
            ws.Cells(r, 1).Value = CStr(r - 1)
            ws.Cells(r, 2).Value = ToText(mtx(idx, MC_PART))
            ws.Cells(r, 3).Value = h
            If headSpecs.Exists(h) Then ws.Cells(r, 4).Value = Replace(headSpecs.Item(h), vbTab, ", ")
            ws.Cells(r, 5).Value = want
            ws.Cells(r, 6).Value = ToText(mtx(idx, MC_VALCELL))
            If hasCopy Then LinkToCopy ws.Cells(r, 6), ToText(mtx(idx, MC_VALCELL))
            ws.Cells(r, 7).Value = rowsTxt
        Next k
    Next pass
    On Error Resume Next
    ws.Columns("A:I").AutoFit
    On Error GoTo 0
End Sub

' ---- 設定の控え ----
Private Sub WriteConfig(ByVal ws As Worksheet, ByVal lay As Object)
    Dim r As Long, v As Variant, i As Long, j As Long, keys As Variant, names As Variant
    ws.Cells(1, 1).Value = "この照合で使った設定の控え"
    ws.Cells(1, 1).Font.Bold = True
    r = 3
    ws.Cells(r, 1).Value = "マトリックスの形": ws.Cells(r, 1).Font.Bold = True
    keys = Array("name", "sheet", "orient", "partRange", "nameRange", "headRange", "valueKind", "inherit", "split", "pattern")
    names = Array("設定名", "シート名", "品番の向き", "品番の範囲", "名称の範囲", "仕様の見出しの範囲", "値の種類", _
                  "見出しの空欄の引き継ぎ", "1 セルに複数の品番", "品番の形")
    For i = LBound(keys) To UBound(keys)
        r = r + 1
        ws.Cells(r, 1).Value = names(i)
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).Value = ToText(lay.Item(keys(i)))
    Next i
    r = r + 2
    ws.Cells(r, 1).Value = "部品表の読み方": ws.Cells(r, 1).Font.Bold = True
    r = CopyTable(ws, r, ReadTable(ToolSheet(SH_CFG_BOM), 2), 2)
    r = r + 2
    ws.Cells(r, 1).Value = "仕様対応表 (設定名 / 見出しの文字 / 仕様コード / 対象外 / メモ)": ws.Cells(r, 1).Font.Bold = True
    v = ReadTable(ToolSheet(SH_MAP), PC_N)
    For i = 1 To RowCount(v)
        If Trim(ToText(v(i, PC_LAYOUT))) = lay.Item("name") Then
            r = r + 1
            For j = 1 To PC_N
                ws.Cells(r, j).NumberFormat = "@"
                ws.Cells(r, j).Value = ToText(v(i, j))
            Next j
        End If
    Next i
    r = r + 2
    ws.Cells(r, 1).Value = "文字のそろえ方 (種類 / 文字)": ws.Cells(r, 1).Font.Bold = True
    r = CopyTable(ws, r, ReadTable(ToolSheet(SH_CFG_TXT), 2), 2)
    On Error Resume Next
    ws.Columns("A:E").AutoFit
    On Error GoTo 0
End Sub

Private Function CopyTable(ByVal ws As Worksheet, ByVal r As Long, ByVal v As Variant, ByVal nCol As Long) As Long
    Dim i As Long, j As Long
    For i = 1 To RowCount(v)
        r = r + 1
        For j = 1 To nCol
            ws.Cells(r, j).NumberFormat = "@"
            ws.Cells(r, j).Value = ToText(v(i, j))
        Next j
    Next i
    CopyTable = r
End Function

' ---- まとめ ----
Private Sub WriteSummary(ByVal ws As Worksheet, ByRef res As TResult, ByVal verdict As String, ByVal paintNote As String, _
        ByVal nTried As Long, ByVal nDetected As Long)
    Dim out As Variant, codes As Variant, i As Long, t As String, r As Long, c As Long, user As String

    SumReset
    SumAdd "判定", verdict
    SumAdd "", ""
    SumAdd "■ いつ・だれが・何で", ""
    SumAdd "実行した日時", StampText(Now)
    user = Environ("USERNAME")
    SumAdd "実行した人 (Windows のユーザー名)", user
    SumAdd "ツールの版", TOOL_VERSION
    t = StateGet("test.result")
    If Len(t) = 0 Then t = "まだ実行していない"
    SumAdd "自己診断 (テスト) の最後の結果", t
    SumAdd "", ""
    SumAdd "■ 部品表", ""
    SumAdd "ファイル", StateGet("bom.path")
    SumAdd "シート", StateGet("bom.sheet")
    SumAdd "読んだときの更新日時と大きさ", StateGet("bom.stamp")
    SumAdd "データの行数", StateGet("bom.dataRows")
    SumAdd "読んだ行", StateGet("bom.readRows")
    SumAdd "空で飛ばした行", StateGet("bom.blankRows")
    SumAdd "エラーで読まなかった行", StateGet("bom.errRows")
    SumAdd "共通の形に直した行数 (品番 × 仕様)", CStr(res.BRows)
    SumAdd "", ""
    SumAdd "■ マトリックス", ""
    SumAdd "ファイル", StateGet("mtx.path")
    SumAdd "シート", StateGet("mtx.sheet")
    SumAdd "設定名", StateGet("mtx.layout")
    SumAdd "読んだときの更新日時と大きさ", StateGet("mtx.stamp")
    SumAdd "読んだ品番の数", StateGet("mtx.parts")
    SumAdd "読んだ仕様の列の数", StateGet("mtx.cols")
    SumAdd "読んだ値のセルの数", CStr(res.MCells)
    SumAdd "そのうち「あり」の数", StateGet("mtx.yes")
    SumAdd "品番が空で飛ばした行 (列) の数", StateGet("mtx.blankRows")
    If Len(paintNote) > 0 Then SumAdd "注意", paintNote
    SumAdd "", ""
    SumAdd "■ 比べた範囲", ""
    SumAdd "対象の仕様", OneLine(res.SpecsT)
    SumAdd "対象外にした見出し", OneLine(res.HeadsExcluded)
    SumAdd "比べなかった部品表の仕様", OneLine(res.SpecsOutside)
    t = ""
    If res.SkippedN1 Then t = t & "N-1 (部品表の絞り込みが未設定のため、マトリックスの書きもれは調べていない) "
    If res.SkippedN3 Then t = t & "N-3 (名称の範囲か名称の列が未設定のため、名称は比べていない) "
    If res.SkippedN6 Then t = t & "N-6 (値の種類が印、または員数の列が未設定のため、員数は比べていない)"
    If Len(t) = 0 Then t = "(なし)"
    SumAdd "調べなかった項目", t
    SumAdd "", ""
    SumAdd "■ 件数", ""
    SumAdd "不一致の合計", CStr(res.NNg)
    SumAdd "警告の合計", CStr(res.NWarn)
    codes = CodeList()
    For i = LBound(codes) To UBound(codes)
        SumAdd codes(i) & " " & CodeTitle(codes(i)), CStr(CountCode(res, codes(i)))
    Next i
    SumAdd "", ""
    SumAdd "■ 消し込み (全部に状態が付いているか)", ""
    SumAdd "マトリックスのセル: 一致", CStr(res.MOk)
    SumAdd "マトリックスのセル: 不一致", CStr(res.MNg)
    SumAdd "マトリックスのセル: 判断できない", CStr(res.MUnk)
    SumAdd "マトリックスのセル: 対象外", CStr(res.MOut)
    SumAdd "部品表の行: 一致", CStr(res.BOk)
    SumAdd "部品表の行: 不一致", CStr(res.BNg)
    SumAdd "部品表の行: 判断できない", CStr(res.BUnk)
    SumAdd "部品表の行: 対象外", CStr(res.BOut)
    SumAdd "", ""
    SumAdd "■ 検算", ""
    If mProofBad = 0 Then
        SumAdd "検算の結果", "全部合った (シート「" & R_PROOF & "」)"
    Else
        SumAdd "検算の結果", "合わない項目が " & mProofBad & " 件ある (シート「" & R_PROOF & "」)"
    End If
    SumAdd "抜き打ち", "わざと値を逆にした " & nTried & " か所のうち、" & nDetected & " か所で照合が気付いた"
    SumAdd "", ""
    SumAdd "■ 色の意味 (シート「" & R_MTX & "」)", ""
    SumAdd "緑", "一致 (部品表と合っている)"
    SumAdd "赤", "不一致"
    SumAdd "黄", "判断できない、または警告"
    SumAdd "灰", "対象外 (比べていない)"
    SumAdd "水色", "読んだ見出し"
    SumAdd "色なし", "ツールが見ていないセル。表の中に色なしのセルがあれば、範囲の指定を確かめること"

    out = BufToRows(mSum, mSumN, 2)
    With ws.Range(ws.Cells(1, 1), ws.Cells(mSumN, 2))
        .NumberFormat = "@"
        .Value = out
    End With
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A1:B1").Font.Size = 14
    Select Case verdict
        Case "一致": c = COL_OK
        Case "一致 (警告あり)": c = COL_UNK
        Case Else: c = COL_NG
    End Select
    ws.Range("B1").Interior.Color = c
    For r = 1 To mSumN
        If Left(CStr(mSum(1, r)), 1) = ChrW(&H25A0) Then ws.Cells(r, 1).Font.Bold = True
        Select Case CStr(mSum(1, r))
            Case "緑": ws.Cells(r, 1).Interior.Color = COL_OK
            Case "赤": ws.Cells(r, 1).Interior.Color = COL_NG
            Case "黄": ws.Cells(r, 1).Interior.Color = COL_UNK
            Case "灰": ws.Cells(r, 1).Interior.Color = COL_OUT
            Case "水色": ws.Cells(r, 1).Interior.Color = COL_HEAD
        End Select
    Next r
    On Error Resume Next
    ws.Columns("A:B").AutoFit
    On Error GoTo 0
End Sub
