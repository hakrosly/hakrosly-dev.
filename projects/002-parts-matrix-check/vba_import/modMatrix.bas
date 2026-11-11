Attribute VB_Name = "modMatrix"
Option Explicit
' ============================================================
' modMatrix : マトリックスを読んで、共通の形に直す
'   形は設定 (品番の範囲、名称の範囲、仕様の見出しの範囲) で教えてもらう。
'   値のセルは「品番の範囲」と「見出しの範囲」が交わる所。
'   結合セルは、結合範囲の左上の値を、範囲の全部のセルの値とみなす。
'
'   中では「品番が並ぶ向き」を p、「仕様が並ぶ向き」を s と呼ぶ。
'     品番が縦のとき: p = 行、s = 列
'     品番が横のとき: p = 列、s = 行
'   こうしておくと、縦と横の違いは GV / GRaw / GCell などの小さな関数だけで済む。
' ============================================================

Private mWs As Worksheet
Private mVert As Boolean
Private mR1 As Long, mC1 As Long, mR2 As Long, mC2 As Long   ' 読み込んだ四角 (シートの行と列)
Private mRaw As Variant        ' セルそのものの値
Private mVal As Variant        ' 結合を反映した値
Private mAncR() As Long        ' 空のセルが結合されているとき、結合の左上の行 (結合でなければ 0)
Private mAncC() As Long        ' 同じく左上の列
Private mUsed() As Boolean     ' 「あり」として読み取りに使ったセル (検算用)
Private mUsedOut As Long       ' 四角の外にある結合の左上を使った数 (検算用)

' ---- 設定 (形) を 1 つにまとめた辞書を作る ----
'   name      : 設定名
'   sheet     : シート名
'   orient    : 品番の向き (縦 / 横)
'   partRange : 品番の範囲 (例 C6:C40)
'   nameRange : 名称の範囲 (空でもよい)
'   headRange : 仕様の見出しの範囲 (例 E3:P5)
'   valueKind : 値の種類 (印 / 員数)
'   inherit   : 見出しの空欄の引き継ぎ (する / しない)
'   split     : 1 セルに複数の品番 (分ける / 分けない)
'   pattern   : 品番の形 (Like の書き方。空でもよい)
Public Function NewLayout(ByVal layName As String, ByVal sheetName As String, ByVal orient As String, _
        ByVal partRange As String, ByVal nameRange As String, ByVal headRange As String, _
        ByVal valueKind As String, ByVal inherit As String, ByVal splitOpt As String, ByVal pattern As String) As Object
    Dim d As Object
    Set d = NewDict()
    d.Add "name", layName
    d.Add "sheet", sheetName
    d.Add "orient", orient
    d.Add "partRange", partRange
    d.Add "nameRange", nameRange
    d.Add "headRange", headRange
    d.Add "valueKind", valueKind
    d.Add "inherit", inherit
    d.Add "split", splitOpt
    d.Add "pattern", pattern
    Set NewLayout = d
End Function

' シート「設定_マトリックス」から、設定名の合う行を読む。無ければ Nothing
Public Function LoadLayout(ByVal layName As String) As Object
    Dim v As Variant, i As Long
    Set LoadLayout = Nothing
    v = ReadTable(ToolSheet(SH_CFG_MTX), 12)
    For i = 1 To RowCount(v)
        If Trim(ToText(v(i, 1))) = layName Then
            Set LoadLayout = NewLayout(layName, Trim(ToText(v(i, 2))), Trim(ToText(v(i, 3))), _
                Trim(ToText(v(i, 4))), Trim(ToText(v(i, 5))), Trim(ToText(v(i, 6))), _
                Trim(ToText(v(i, 7))), Trim(ToText(v(i, 8))), Trim(ToText(v(i, 9))), Trim(ToText(v(i, 10))))
            Exit Function
        End If
    Next i
End Function

' 設定名の一覧を、改行で区切った文字で返す
Public Function LayoutNames() As String
    Dim v As Variant, i As Long, s As String
    v = ReadTable(ToolSheet(SH_CFG_MTX), 12)
    For i = 1 To RowCount(v)
        If Len(Trim(ToText(v(i, 1)))) > 0 Then
            If Len(s) > 0 Then s = s & vbLf
            s = s & Trim(ToText(v(i, 1)))
        End If
    Next i
    LayoutNames = s
End Function

' 設定をシート「設定_マトリックス」に保存する (同じ設定名があれば書き換える)
Public Sub SaveLayout(ByVal lay As Object)
    Dim ws As Worksheet, v As Variant, i As Long, r As Long
    Set ws = ToolSheet(SH_CFG_MTX)
    v = ReadTable(ws, 12)
    r = RowCount(v) + 2
    For i = 1 To RowCount(v)
        If Trim(ToText(v(i, 1))) = lay.Item("name") Then r = i + 1
    Next i
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).NumberFormat = "@"
    ws.Cells(r, 1).Value = lay.Item("name")
    ws.Cells(r, 2).Value = lay.Item("sheet")
    ws.Cells(r, 3).Value = lay.Item("orient")
    ws.Cells(r, 4).Value = lay.Item("partRange")
    ws.Cells(r, 5).Value = lay.Item("nameRange")
    ws.Cells(r, 6).Value = lay.Item("headRange")
    ws.Cells(r, 7).Value = lay.Item("valueKind")
    ws.Cells(r, 8).Value = lay.Item("inherit")
    ws.Cells(r, 9).Value = lay.Item("split")
    ws.Cells(r, 10).Value = lay.Item("pattern")
    ws.Cells(r, 11).Value = StampText(Now)
End Sub

' ---- 範囲の扱い ----

' 範囲を、p の向きの端 (pA～pB) と s の向きの端 (sA～sB) に直す
Private Sub Extent(ByVal rng As Range, ByRef pA As Long, ByRef pB As Long, ByRef sA As Long, ByRef sB As Long)
    If mVert Then
        pA = rng.Row: pB = rng.Row + rng.Rows.Count - 1
        sA = rng.Column: sB = rng.Column + rng.Columns.Count - 1
    Else
        pA = rng.Column: pB = rng.Column + rng.Columns.Count - 1
        sA = rng.Row: sB = rng.Row + rng.Rows.Count - 1
    End If
End Sub

Private Function TryRange(ByVal ws As Worksheet, ByVal addr As String) As Range
    Set TryRange = Nothing
    If Len(addr) = 0 Then Exit Function
    On Error Resume Next
    Set TryRange = ws.Range(addr)
    On Error GoTo 0
End Function

Private Function Overlap(ByVal a1 As Long, ByVal a2 As Long, ByVal b1 As Long, ByVal b2 As Long) As Boolean
    Overlap = Not (a2 < b1 Or b2 < a1)
End Function

' 設定の範囲が決まりに合っているかを確かめる。良ければ ""、悪ければ理由の文を返す
Public Function CheckLayout(ByVal ws As Worksheet, ByVal lay As Object) As String
    Dim rPart As Range, rName As Range, rHead As Range
    Dim p1 As Long, p2 As Long, d1 As Long, d2 As Long
    Dim h1 As Long, h2 As Long, s1 As Long, s2 As Long
    Dim n1 As Long, n2 As Long, e1 As Long, e2 As Long
    Dim keepVert As Boolean
    CheckLayout = ""
    If lay.Item("orient") <> O_VERT And lay.Item("orient") <> O_HORZ Then
        CheckLayout = "品番の向きは「" & O_VERT & "」か「" & O_HORZ & "」にしてください。"
        Exit Function
    End If
    If lay.Item("valueKind") <> K_MARK And lay.Item("valueKind") <> K_QTY Then
        CheckLayout = "値の種類は「" & K_MARK & "」か「" & K_QTY & "」にしてください。"
        Exit Function
    End If
    Set rPart = TryRange(ws, lay.Item("partRange"))
    If rPart Is Nothing Then
        CheckLayout = "品番の範囲「" & lay.Item("partRange") & "」が正しくありません。"
        Exit Function
    End If
    Set rHead = TryRange(ws, lay.Item("headRange"))
    If rHead Is Nothing Then
        CheckLayout = "仕様の見出しの範囲「" & lay.Item("headRange") & "」が正しくありません。"
        Exit Function
    End If
    keepVert = mVert
    mVert = (lay.Item("orient") = O_VERT)
    Extent rPart, p1, p2, d1, d2
    Extent rHead, h1, h2, s1, s2
    If Overlap(p1, p2, h1, h2) Or Overlap(d1, d2, s1, s2) Then
        CheckLayout = "品番の範囲と仕様の見出しの範囲が、同じ行または同じ列にかかっています。" & _
            "品番の向きが「" & lay.Item("orient") & "」で合っているか、範囲が正しいかを確かめてください。"
    ElseIf Len(lay.Item("nameRange")) > 0 Then
        Set rName = TryRange(ws, lay.Item("nameRange"))
        If rName Is Nothing Then
            CheckLayout = "名称の範囲「" & lay.Item("nameRange") & "」が正しくありません。"
        Else
            Extent rName, n1, n2, e1, e2
            If n1 <> p1 Or n2 <> p2 Then
                CheckLayout = "名称の範囲は、品番の範囲と同じ長さで、同じ位置に並べてください。"
            ElseIf Overlap(e1, e2, s1, s2) Then
                CheckLayout = "名称の範囲が、仕様の見出しの範囲と重なっています。"
            End If
        End If
    End If
    mVert = keepVert
End Function

' 品番の範囲のすぐ先のセルが空でなければ True (行や列が増えたかもしれない)
Public Function LayoutMaybeGrown(ByVal ws As Worksheet, ByVal lay As Object) As Boolean
    Dim rPart As Range, c As Range, v As Variant
    LayoutMaybeGrown = False
    Set rPart = TryRange(ws, lay.Item("partRange"))
    If rPart Is Nothing Then Exit Function
    On Error GoTo Done
    If lay.Item("orient") = O_VERT Then
        Set c = ws.Cells(rPart.Row + rPart.Rows.Count, rPart.Column)
    Else
        Set c = ws.Cells(rPart.Row, rPart.Column + rPart.Columns.Count)
    End If
    If c.MergeCells Then
        v = c.MergeArea.Cells(1, 1).Value
    Else
        v = c.Value
    End If
    LayoutMaybeGrown = Not IsBlankV(v)
Done:
End Function

' ---- 四角を読み込む (結合をほどく) ----

' 範囲の中に結合セルがあるか。
' Range.MergeCells は、全部が結合なら True、結合なしなら False、混ざっていれば Null を返す
Public Function RangeHasMerge(ByVal rng As Range) As Boolean
    Dim m As Variant
    m = rng.MergeCells
    If IsNull(m) Then
        RangeHasMerge = True
    Else
        RangeHasMerge = CBool(m)
    End If
End Function

Private Sub LoadGrid(ByVal ws As Worksheet, ByVal r1 As Long, ByVal c1 As Long, ByVal r2 As Long, ByVal c2 As Long)
    Dim rng As Range, nr As Long, nc As Long, i As Long, j As Long
    Dim c As Range, ma As Range
    Set mWs = ws
    mR1 = r1: mC1 = c1: mR2 = r2: mC2 = c2
    nr = r2 - r1 + 1: nc = c2 - c1 + 1
    Set rng = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    If nr = 1 And nc = 1 Then
        ReDim mRaw(1 To 1, 1 To 1)
        ReDim mVal(1 To 1, 1 To 1)
        mRaw(1, 1) = rng.Value
        mVal(1, 1) = rng.Value
    Else
        ' 2 回読む (片方は元のまま残し、もう片方に結合を反映する)
        mRaw = rng.Value
        mVal = rng.Value
    End If
    ReDim mAncR(1 To nr, 1 To nc)
    ReDim mAncC(1 To nr, 1 To nc)
    ReDim mUsed(1 To nr, 1 To nc)
    mUsedOut = 0
    If RangeHasMerge(rng) Then
        ' 結合セルは左上にしか値が無い。空のセルだけを調べれば足りる
        For i = 1 To nr
            For j = 1 To nc
                If IsBlankV(mRaw(i, j)) Then
                    Set c = ws.Cells(r1 + i - 1, c1 + j - 1)
                    If c.MergeCells Then
                        Set ma = c.MergeArea
                        mAncR(i, j) = ma.Row
                        mAncC(i, j) = ma.Column
                        If ma.Row <> c.Row Or ma.Column <> c.Column Then
                            mVal(i, j) = ma.Cells(1, 1).Value
                        End If
                    End If
                End If
            Next j
        Next i
    End If
End Sub

' 結合を反映した値
Private Function GV(ByVal p As Long, ByVal s As Long) As Variant
    If mVert Then
        GV = mVal(p - mR1 + 1, s - mC1 + 1)
    Else
        GV = mVal(s - mR1 + 1, p - mC1 + 1)
    End If
End Function

' セルそのものの値
Private Function GRaw(ByVal p As Long, ByVal s As Long) As Variant
    If mVert Then
        GRaw = mRaw(p - mR1 + 1, s - mC1 + 1)
    Else
        GRaw = mRaw(s - mR1 + 1, p - mC1 + 1)
    End If
End Function

Private Function GCell(ByVal p As Long, ByVal s As Long) As Range
    If mVert Then
        Set GCell = mWs.Cells(p, s)
    Else
        Set GCell = mWs.Cells(s, p)
    End If
End Function

' 空で、結合されているセルか
Private Function IsMergedBlank(ByVal p As Long, ByVal s As Long) As Boolean
    If mVert Then
        IsMergedBlank = (mAncR(p - mR1 + 1, s - mC1 + 1) <> 0)
    Else
        IsMergedBlank = (mAncR(s - mR1 + 1, p - mC1 + 1) <> 0)
    End If
End Function

' 「あり」として使ったセルに印を付ける (結合なら左上のセルに付ける)
Private Sub MarkUsed(ByVal p As Long, ByVal s As Long)
    Dim i As Long, j As Long, ar As Long, ac As Long
    If mVert Then
        i = p - mR1 + 1: j = s - mC1 + 1
    Else
        i = s - mR1 + 1: j = p - mC1 + 1
    End If
    If mAncR(i, j) <> 0 Then
        ar = mAncR(i, j): ac = mAncC(i, j)
        If ar < mR1 Or ar > mR2 Or ac < mC1 Or ac > mC2 Then
            mUsedOut = mUsedOut + 1
            Exit Sub
        End If
        i = ar - mR1 + 1: j = ac - mC1 + 1
    End If
    mUsed(i, j) = True
End Sub

Private Function WasUsed(ByVal p As Long, ByVal s As Long) As Boolean
    If mVert Then
        WasUsed = mUsed(p - mR1 + 1, s - mC1 + 1)
    Else
        WasUsed = mUsed(s - mR1 + 1, p - mC1 + 1)
    End If
End Function

' p の向きの線 (品番が縦なら行) が非表示か
Private Function HiddenP(ByVal p As Long) As Boolean
    If mVert Then HiddenP = mWs.Rows(p).Hidden Else HiddenP = mWs.Columns(p).Hidden
End Function

' s の向きの線 (品番が縦なら列) が非表示か
Private Function HiddenS(ByVal s As Long) As Boolean
    If mVert Then HiddenS = mWs.Columns(s).Hidden Else HiddenS = mWs.Rows(s).Hidden
End Function

Private Function RangeAddr(ByVal pA As Long, ByVal sA As Long, ByVal pB As Long, ByVal sB As Long) As String
    RangeAddr = mWs.Range(GCell(pA, sA), GCell(pB, sB)).Address(False, False)
End Function

Private Function Rank(ByVal cls As String) As Long
    If cls = V_YES Then
        Rank = 2
    ElseIf cls = V_UNK Then
        Rank = 1
    Else
        Rank = 0
    End If
End Function

Private Function HasStrike(ByVal c As Range) As Boolean
    Dim st As Variant
    HasStrike = False
    On Error GoTo Done
    st = c.Font.Strikethrough
    If IsNull(st) Then
        HasStrike = True          ' 文字の一部だけに取り消し線がある
    ElseIf VarType(st) = vbBoolean Then
        HasStrike = CBool(st)
    End If
Done:
End Function

' ============================================================
' マトリックスのシートを読む。
'   戻り値: 読取_マトリックス と同じ並びの配列 (行, MC_N)。
'           値のセル 1 つにつき 1 行 (品番 × 仕様の列)。「なし」も書く。
'   info  : 読んだ数などを入れて返す辞書
'   警告は modLog にためる (呼ぶ前に LogReset しておくこと)
' ============================================================
Public Function ReadMatrixSheet(ByVal ws As Worksheet, ByVal lay As Object, ByVal info As Object) As Variant
    Dim msg As String
    Dim rPart As Range, rName As Range, rHead As Range, hasName As Boolean
    Dim p1 As Long, p2 As Long, d1 As Long, d2 As Long      ' 品番: 並ぶ向きの端、奥行きの端
    Dim h1 As Long, h2 As Long, s1 As Long, s2 As Long      ' 見出し: 段の端、並ぶ向きの端
    Dim n1 As Long, n2 As Long, e1 As Long, e2 As Long      ' 名称
    Dim g1 As Long, g2 As Long, t1 As Long, t2 As Long      ' 読み込む四角
    Dim nS As Long, nL As Long, j As Long, lv As Long, p As Long, s As Long, q As Long, k As Long
    Dim hv() As String, headText() As String, headCell() As String, colOk() As Boolean
    Dim inheritOn As Boolean, splitOn As Boolean, asQty As Boolean, pat As String
    Dim t As String, lab As String, prev As String, ok As Boolean, nCols As Long
    Dim c As Range, ma As Range, pe As Long, v As Variant
    Dim txt As String, piece As String, numFlag As Boolean, errFlag As Boolean
    Dim parts As Variant, key As String, nm As String, nmCell As String
    Dim eKey() As String, ePart() As String, eName() As String, eP1() As Long, eP2() As Long
    Dim ePartCell() As String, eNameCell() As String, nEnt As Long, capE As Long
    Dim seen As Object, nBlank As Long
    Dim out() As Variant, n As Long, e As Long
    Dim best As String, bestQty As Variant, bestRaw As String, cls As String, qty As Variant
    Dim nYes As Long, rawYes As Long, usedYes As Long, unread As String, nUnread As Long
    Dim heads As String, headSeen As Object, i As Long, hid As Boolean

    msg = CheckLayout(ws, lay)
    If Len(msg) > 0 Then Err.Raise ERR_TOOL, "ReadMatrixSheet", "設定「" & lay.Item("name") & "」: " & msg
    mVert = (lay.Item("orient") = O_VERT)
    Set rPart = ws.Range(lay.Item("partRange"))
    Set rHead = ws.Range(lay.Item("headRange"))
    hasName = (Len(lay.Item("nameRange")) > 0)
    Extent rPart, p1, p2, d1, d2
    Extent rHead, h1, h2, s1, s2
    g1 = MinL(p1, h1): g2 = MaxL(p2, h2)
    t1 = MinL(d1, s1): t2 = MaxL(d2, s2)
    If hasName Then
        Set rName = ws.Range(lay.Item("nameRange"))
        Extent rName, n1, n2, e1, e2
        t1 = MinL(t1, e1): t2 = MaxL(t2, e2)
    End If
    If mVert Then
        LoadGrid ws, g1, t1, g2, t2
    Else
        LoadGrid ws, t1, g1, t2, g2
    End If
    inheritOn = (lay.Item("inherit") = OPT_YES)
    splitOn = (lay.Item("split") = OPT_SPLIT)
    asQty = (lay.Item("valueKind") = K_QTY)
    pat = NormBase(lay.Item("pattern"))

    ' ---- 1. 仕様の列ごとに「見出しの文字」を作る ----
    nS = s2 - s1 + 1
    nL = h2 - h1 + 1
    ReDim hv(1 To nL, 1 To nS)
    ReDim headText(1 To nS)
    ReDim headCell(1 To nS)
    ReDim colOk(1 To nS)
    nCols = 0
    For j = 1 To nS
        s = s1 + j - 1
        For lv = 1 To nL
            p = h1 + lv - 1
            t = CleanText(GV(p, s))
            ' 空欄の引き継ぎ: 一番下の段以外で、空で、結合もされていないとき、
            ' 一つ上の段の値が隣と同じ間だけ、隣の値を引き継ぐ
            If Len(t) = 0 And inheritOn And lv < nL And j > 1 Then
                If Not IsMergedBlank(p, s) Then
                    ok = True
                    If lv > 1 Then ok = (hv(lv - 1, j) = hv(lv - 1, j - 1))
                    If ok Then t = hv(lv, j - 1)
                End If
            End If
            hv(lv, j) = t
        Next lv
        lab = "": prev = ""
        For lv = 1 To nL
            If Len(hv(lv, j)) > 0 Then
                If hv(lv, j) <> prev Then
                    If Len(lab) > 0 Then lab = lab & " / "
                    lab = lab & hv(lv, j)
                End If
                prev = hv(lv, j)
            End If
        Next lv
        headText(j) = lab
        headCell(j) = RangeAddr(h1, s, h2, s)
        If Len(lab) = 0 Then
            colOk(j) = False
            LogAdd SRC_MTX, "W-6", "", "", headCell(j), "", "見出しが全部空の列 (行) があり、読まなかった"
        Else
            colOk(j) = True
            nCols = nCols + 1
            If HiddenS(s) Then
                LogAdd SRC_MTX, "W-5", "", lab, headCell(j), "", "非表示になっている仕様の列 (行) を読んだ"
            End If
        End If
    Next j

    ' ---- 2. 品番を先頭から順に読む ----
    capE = 64
    ReDim eKey(1 To capE)
    ReDim ePart(1 To capE)
    ReDim eName(1 To capE)
    ReDim eP1(1 To capE)
    ReDim eP2(1 To capE)
    ReDim ePartCell(1 To capE)
    ReDim eNameCell(1 To capE)
    nEnt = 0
    Set seen = NewDict()
    p = p1
    Do While p <= p2
        Set c = GCell(p, d1)
        pe = p
        If c.MergeCells Then
            ' 品番が結合されているときは、結合範囲で 1 つの品番として読む
            Set ma = c.MergeArea
            If mVert Then
                pe = ma.Row + ma.Rows.Count - 1
            Else
                pe = ma.Column + ma.Columns.Count - 1
            End If
            If pe > p2 Then pe = p2
            If pe < p Then pe = p
        End If
        txt = "": numFlag = False: errFlag = False
        For s = d1 To d2
            v = GV(p, s)
            If IsError(v) Then
                errFlag = True
            ElseIf Not IsBlankV(v) Then
                If VarType(v) = vbString Then
                    piece = CStr(v)
                Else
                    numFlag = True
                    piece = ShownText(GCell(p, s), v)
                End If
                txt = txt & piece
            End If
        Next s
        If errFlag Then
            LogAdd SRC_MTX, "W-4", "", "", RangeAddr(p, d1, pe, d2), "", "品番のセルがエラー値のため、読まなかった"
        ElseIf Len(NormBase(txt)) = 0 Then
            nBlank = nBlank + (pe - p + 1)
        Else
            If splitOn Then
                parts = Split(Replace(txt, vbCr, ""), vbLf)
            Else
                parts = Array(txt)
            End If
            nm = "": nmCell = ""
            If hasName Then
                For s = e1 To e2
                    t = CleanText(GV(p, s))
                    If Len(t) > 0 Then
                        If Len(nm) > 0 Then nm = nm & " "
                        nm = nm & t
                    End If
                Next s
                nmCell = RangeAddr(p, e1, pe, e2)
            End If
            For k = LBound(parts) To UBound(parts)
                piece = CleanText(parts(k))
                key = NormPart(piece)
                If Len(key) = 0 Then
                    ' 空の切れはしは飛ばす
                ElseIf Len(pat) > 0 And Not (key Like pat) Then
                    LogAdd SRC_MTX, "W-9", piece, "", RangeAddr(p, d1, pe, d2), piece, "「品番の形」に合わないため、品番として読まなかった"
                Else
                    If nEnt >= capE Then
                        capE = capE * 2
                        ReDim Preserve eKey(1 To capE)
                        ReDim Preserve ePart(1 To capE)
                        ReDim Preserve eName(1 To capE)
                        ReDim Preserve eP1(1 To capE)
                        ReDim Preserve eP2(1 To capE)
                        ReDim Preserve ePartCell(1 To capE)
                        ReDim Preserve eNameCell(1 To capE)
                    End If
                    nEnt = nEnt + 1
                    eKey(nEnt) = key
                    ePart(nEnt) = piece
                    eName(nEnt) = nm
                    eP1(nEnt) = p
                    eP2(nEnt) = pe
                    ePartCell(nEnt) = RangeAddr(p, d1, pe, d2)
                    eNameCell(nEnt) = nmCell
                    If numFlag Then
                        LogAdd SRC_MTX, "W-4", piece, "", ePartCell(nEnt), piece, "品番が文字ではなく数値や日付として入っている (先頭の 0 や桁が変わっているおそれ)"
                    End If
                    If seen.Exists(key) Then
                        LogAdd SRC_MTX, "W-1", piece, "", ePartCell(nEnt), piece, "マトリックスに同じ品番が 2 回以上ある (前は " & seen.Item(key) & ")。どちらかで「あり」なら「あり」として扱った"
                    Else
                        seen.Add key, ePartCell(nEnt)
                    End If
                    hid = False
                    For q = p To pe
                        If HiddenP(q) Then hid = True
                    Next q
                    If hid Then
                        LogAdd SRC_MTX, "W-5", piece, "", ePartCell(nEnt), "", "非表示になっている品番の行 (列) を読んだ"
                    End If
                    If HasStrike(GCell(p, d1)) Then
                        LogAdd SRC_MTX, "W-8", piece, "", ePartCell(nEnt), piece, "取り消し線のある品番を読んだ"
                    End If
                End If
            Next k
        End If
        p = pe + 1
    Loop

    ' ---- 3. 品番ごとに、仕様の列のセルを読む ----
    nYes = 0
    n = 0
    If nEnt > 0 And nCols > 0 Then
        ReDim out(1 To nEnt * nCols, 1 To MC_N)
        For e = 1 To nEnt
            For j = 1 To nS
                If colOk(j) Then
                    s = s1 + j - 1
                    best = V_NO: bestQty = "": bestRaw = ""
                    For q = eP1(e) To eP2(e)
                        v = GV(q, s)
                        cls = Classify(v, asQty, qty)
                        If cls = V_YES Then MarkUsed q, s
                        If Rank(cls) > Rank(best) Then
                            best = cls
                            bestQty = qty
                            bestRaw = CleanText(v)
                        ElseIf Len(bestRaw) = 0 Then
                            bestRaw = CleanText(v)
                        End If
                    Next q
                    n = n + 1
                    out(n, MC_KEY) = eKey(e)
                    out(n, MC_PART) = ePart(e)
                    out(n, MC_NAME) = eName(e)
                    out(n, MC_HEAD) = headText(j)
                    out(n, MC_VAL) = best
                    out(n, MC_QTY) = bestQty
                    out(n, MC_RAW) = bestRaw
                    out(n, MC_PARTCELL) = ePartCell(e)
                    out(n, MC_VALCELL) = RangeAddr(eP1(e), s, eP2(e), s)
                    out(n, MC_NAMECELL) = eNameCell(e)
                    out(n, MC_HEADCELL) = headCell(j)
                    If best = V_YES Then
                        nYes = nYes + 1
                    ElseIf best = V_UNK Then
                        LogAdd SRC_MTX, "W-3", ePart(e), headText(j), CStr(out(n, MC_VALCELL)), bestRaw, "印の一覧に無い値が入っている。「あり」にも「なし」にも数えていない"
                    End If
                End If
            Next j
        Next e
    End If

    ' ---- 4. 検算用: 値の四角を、読み取りとは別の道筋で数える ----
    '   品番の範囲 × 見出しの範囲 の全部のセルを端から見て、「あり」の印が入ったセルを数える。
    '   読み取りで使ったセルの数と合わなければ、読まれていない印がある。
    rawYes = 0: unread = "": nUnread = 0
    For p = p1 To p2
        For s = s1 To s2
            cls = Classify(GRaw(p, s), asQty, qty)
            If cls = V_YES Then
                rawYes = rawYes + 1
                If Not WasUsed(p, s) Then
                    nUnread = nUnread + 1
                    If nUnread <= 30 Then
                        If Len(unread) > 0 Then unread = unread & ", "
                        unread = unread & GCell(p, s).Address(False, False)
                    End If
                End If
            End If
        Next s
    Next p
    usedYes = mUsedOut
    For i = 1 To mR2 - mR1 + 1
        For j = 1 To mC2 - mC1 + 1
            If mUsed(i, j) Then usedYes = usedYes + 1
        Next j
    Next i

    ' ---- 5. 範囲の先に続きがありそうなら知らせる ----
    If LayoutMaybeGrown(ws, lay) Then
        LogAdd SRC_MTX, "W-11", "", "", lay.Item("partRange"), "", "品番の範囲のすぐ先のセルが空でない。行 (列) が増えていないか、範囲を確かめること"
    End If

    ' ---- 見出しの文字の一覧 (重なりを除く) ----
    heads = ""
    Set headSeen = NewDict()
    For j = 1 To nS
        If colOk(j) Then
            If Not headSeen.Exists(headText(j)) Then
                headSeen.Add headText(j), True
                If Len(heads) > 0 Then heads = heads & vbLf
                heads = heads & headText(j)
            End If
        End If
    Next j

    DictSet info, "parts", nEnt
    DictSet info, "cols", nCols
    DictSet info, "cells", n
    DictSet info, "yes", nYes
    DictSet info, "blankRows", nBlank
    DictSet info, "rawYes", rawYes
    DictSet info, "usedYes", usedYes
    DictSet info, "unreadCount", nUnread
    DictSet info, "unread", unread
    DictSet info, "headings", heads
    DictSet info, "hasName", hasName
    DictSet info, "asQty", asQty
    If n = 0 Then
        ReadMatrixSheet = Empty
    Else
        ReadMatrixSheet = out
    End If
End Function
