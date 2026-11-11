Option Explicit
' ============================================================
' modBom : 部品表を読んで、共通の形 (1 行 = 品番 1 つ × 仕様 1 つ) に直す
'   列の場所はコードに書かない。見出しの文字で探す (設定_部品表)。
' ============================================================

' 設定を 1 つにまとめた辞書を作る
'   sheet      : シート名 (空なら 1 枚目)
'   headRow    : 見出しの行
'   colPart    : 品番の列 (見出しの文字)
'   colName    : 名称の列 (空でもよい)
'   colSpec    : 仕様の列 (複数ならカンマで並べる)
'   specSep    : 1 つのセルに仕様が複数入るときの区切り文字 (空なら分けない)
'   colQty     : 員数の列 (空でもよい)
'   colFilter  : 絞り込みの列 (空でもよい)
'   filterLike : 絞り込みの値 (Like の書き方)
Public Function NewBomConfig(ByVal sheetName As String, ByVal headRow As Long, ByVal colPart As String, _
        ByVal colName As String, ByVal colSpec As String, ByVal specSep As String, ByVal colQty As String, _
        ByVal colFilter As String, ByVal filterLike As String) As Object
    Dim d As Object
    Set d = NewDict()
    d.Add "sheet", sheetName
    d.Add "headRow", headRow
    d.Add "colPart", colPart
    d.Add "colName", colName
    d.Add "colSpec", colSpec
    d.Add "specSep", specSep
    d.Add "colQty", colQty
    d.Add "colFilter", colFilter
    d.Add "filterLike", filterLike
    Set NewBomConfig = d
End Function

' シート「設定_部品表」(1 列目 = 項目、2 列目 = 値) から設定を読む
Public Function LoadBomConfig() As Object
    Dim v As Variant, i As Long, k As String, sv As String
    Dim sheetName As String, headRow As Long, colPart As String, colName As String
    Dim colSpec As String, specSep As String, colQty As String, colFilter As String, filterLike As String
    headRow = 1
    v = ReadTable(ToolSheet(SH_CFG_BOM), 2)
    For i = 1 To RowCount(v)
        k = Trim(ToText(v(i, 1)))
        sv = ToText(v(i, 2))
        Select Case k
            Case "シート名": sheetName = Trim(sv)
            Case "見出しの行": If IsNumeric(sv) Then headRow = CLng(sv)
            Case "品番の列": colPart = Trim(sv)
            Case "名称の列": colName = Trim(sv)
            Case "仕様の列": colSpec = Trim(sv)
            Case "仕様の区切り文字": specSep = sv
            Case "員数の列": colQty = Trim(sv)
            Case "絞り込みの列": colFilter = Trim(sv)
            Case "絞り込みの値": filterLike = Trim(sv)
        End Select
    Next i
    Set LoadBomConfig = NewBomConfig(sheetName, headRow, colPart, colName, colSpec, specSep, colQty, colFilter, filterLike)
End Function

' 見出しの行から、名前の合う列を探す (全角と半角、空白の違いは無視)。無ければ 0
Private Function FindCol(ByRef v As Variant, ByVal hr As Long, ByVal nc As Long, ByVal title As String) As Long
    Dim c As Long, want As String
    FindCol = 0
    want = NormBase(title)
    If Len(want) = 0 Then Exit Function
    For c = 1 To nc
        If NormBase(ToText(v(hr, c))) = want Then
            FindCol = c
            Exit Function
        End If
    Next c
End Function

Private Function NeedCol(ByRef v As Variant, ByVal hr As Long, ByVal nc As Long, ByVal title As String, ByVal what As String) As Long
    Dim c As Long
    c = FindCol(v, hr, nc, title)
    If c = 0 Then
        Err.Raise ERR_TOOL, "ReadBomSheet", "部品表に「" & title & "」という見出しが見つかりません。シート「" & _
            SH_CFG_BOM & "」の「" & what & "」を確かめてください。"
    End If
    NeedCol = c
End Function

' 部品表のシートを読む。
'   戻り値: 読取_部品表 と同じ並びの配列 (行, BC_N)。1 行も無ければ Empty
'   info  : 読んだ行数などを入れて返す辞書
'   警告は modLog にためる (呼ぶ前に LogReset しておくこと)
Public Function ReadBomSheet(ByVal ws As Worksheet, ByVal cfg As Object, ByVal info As Object) As Variant
    Dim ur As Range, v As Variant, r0 As Long, c0 As Long, nr As Long, nc As Long, hr As Long
    Dim cPart As Long, cName As Long, cQty As Long, cFilter As Long
    Dim specNames As Variant, cSpec() As Long, nSpecCol As Long
    Dim sep As String, pat As String, hasFilter As Boolean
    Dim i As Long, j As Long, k As Long
    Dim pv As Variant, partTxt As String, key As String, nm As String, qty As Variant, inScope As Boolean
    Dim pieces() As String, nPiece As Long, t As String, tmp As Variant, fv As String
    Dim buf() As Variant, cap As Long, n As Long
    Dim seen As Object, dk As String
    Dim nRead As Long, nBlank As Long, nErr As Long, cntA As Long
    Dim addr As String

    Set ur = ws.UsedRange
    r0 = ur.Row: c0 = ur.Column
    nr = ur.Rows.Count: nc = ur.Columns.Count
    If nr = 1 And nc = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = ur.Value
    Else
        v = ur.Value
    End If
    hr = CLng(cfg.Item("headRow")) - r0 + 1
    If hr < 1 Or hr > nr Then
        Err.Raise ERR_TOOL, "ReadBomSheet", "部品表の見出しの行 (" & cfg.Item("headRow") & " 行目) にデータがありません。シート「" & _
            SH_CFG_BOM & "」の「見出しの行」を確かめてください。"
    End If

    ' ---- 列を探す ----
    cPart = NeedCol(v, hr, nc, cfg.Item("colPart"), "品番の列")
    cName = 0
    If Len(cfg.Item("colName")) > 0 Then cName = NeedCol(v, hr, nc, cfg.Item("colName"), "名称の列")
    cQty = 0
    If Len(cfg.Item("colQty")) > 0 Then cQty = NeedCol(v, hr, nc, cfg.Item("colQty"), "員数の列")
    cFilter = 0
    hasFilter = (Len(cfg.Item("colFilter")) > 0 And Len(cfg.Item("filterLike")) > 0)
    If hasFilter Then cFilter = NeedCol(v, hr, nc, cfg.Item("colFilter"), "絞り込みの列")
    pat = NormBase(cfg.Item("filterLike"))
    specNames = Split(Replace(cfg.Item("colSpec"), ChrW(&H3001), ","), ",")
    nSpecCol = 0
    ReDim cSpec(1 To UBound(specNames) - LBound(specNames) + 2)
    For k = LBound(specNames) To UBound(specNames)
        If Len(Trim(specNames(k))) > 0 Then
            nSpecCol = nSpecCol + 1
            cSpec(nSpecCol) = NeedCol(v, hr, nc, Trim(specNames(k)), "仕様の列")
        End If
    Next k
    If nSpecCol = 0 Then
        Err.Raise ERR_TOOL, "ReadBomSheet", "シート「" & SH_CFG_BOM & "」の「仕様の列」が空です。"
    End If
    sep = cfg.Item("specSep")

    ' ---- 1 行ずつ読む ----
    cap = 256
    ReDim buf(1 To BC_N, 1 To cap)
    n = 0
    Set seen = NewDict()
    For i = hr + 1 To nr
        pv = v(i, cPart)
        addr = ws.Cells(r0 + i - 1, c0 + cPart - 1).Address(False, False)
        If IsError(pv) Then
            nErr = nErr + 1
            LogAdd SRC_BOM, "W-4", "", "", addr, "", "品番のセルがエラー値のため、この行を読まなかった"
        ElseIf Len(NormBase(ToText(pv))) = 0 Then
            nBlank = nBlank + 1
        Else
            nRead = nRead + 1
            If VarType(pv) = vbString Then
                partTxt = CStr(pv)
            Else
                partTxt = ShownText(ws.Cells(r0 + i - 1, c0 + cPart - 1), pv)
                LogAdd SRC_BOM, "W-4", partTxt, "", addr, ToText(pv), "品番が文字ではなく数値や日付として入っている (先頭の 0 や桁が変わっているおそれ)"
            End If
            key = NormPart(partTxt)
            nm = ""
            If cName > 0 Then nm = CleanText(v(i, cName))
            qty = ""
            If cQty > 0 Then
                t = NormBase(ToText(v(i, cQty)))
                If Len(t) > 0 And IsNumeric(t) Then qty = CDbl(t)
            End If
            inScope = False
            If hasFilter Then
                If cFilter = cPart Then fv = key Else fv = NormBase(ToText(v(i, cFilter)))
                inScope = (fv Like pat)
            End If
            ' 仕様を取り出す
            nPiece = 0
            ReDim pieces(1 To 1)
            If nSpecCol = 1 And Len(sep) > 0 Then
                tmp = Split(ToText(v(i, cSpec(1))), sep)
                If UBound(tmp) >= LBound(tmp) Then
                    ReDim pieces(1 To UBound(tmp) - LBound(tmp) + 1)
                    For k = LBound(tmp) To UBound(tmp)
                        t = CleanText(tmp(k))
                        If Len(t) > 0 Then
                            nPiece = nPiece + 1
                            pieces(nPiece) = t
                        End If
                    Next k
                End If
            Else
                t = ""
                For k = 1 To nSpecCol
                    If Len(CleanText(v(i, cSpec(k)))) > 0 Then
                        If Len(t) > 0 Then t = t & " / "
                        t = t & CleanText(v(i, cSpec(k)))
                    End If
                Next k
                If Len(t) > 0 Then
                    nPiece = 1
                    pieces(1) = t
                End If
            End If
            If nPiece = 0 Then
                nPiece = 1
                pieces(1) = ""
            End If
            For k = 1 To nPiece
                If n >= cap Then
                    cap = cap * 2
                    ReDim Preserve buf(1 To BC_N, 1 To cap)
                End If
                n = n + 1
                buf(BC_KEY, n) = key
                buf(BC_PART, n) = partTxt
                buf(BC_NAME, n) = nm
                buf(BC_SPECKEY, n) = NormSpec(pieces(k))
                buf(BC_SPEC, n) = pieces(k)
                buf(BC_QTY, n) = qty
                If inScope Then buf(BC_INSCOPE, n) = ChrW(&H25CB) Else buf(BC_INSCOPE, n) = ""
                buf(BC_ROW, n) = r0 + i - 1
                If Len(pieces(k)) > 0 Then
                    dk = key & vbTab & NormSpec(pieces(k))
                    If seen.Exists(dk) Then
                        LogAdd SRC_BOM, "W-7", partTxt, pieces(k), addr, "", "部品表に同じ「品番 × 仕様」が 2 行以上ある (" & seen.Item(dk) & " 行目と " & (r0 + i - 1) & " 行目)"
                    Else
                        seen.Add dk, r0 + i - 1
                    End If
                End If
            Next k
        End If
    Next i

    ' ---- 別の数え方 (検算用): 品番の列で、空でないセルの数 ----
    cntA = -1
    If nr > hr Then
        On Error Resume Next
        cntA = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r0 + hr, c0 + cPart - 1), ws.Cells(r0 + nr - 1, c0 + cPart - 1)))
        If Err.Number <> 0 Then cntA = -1
        Err.Clear
        On Error GoTo 0
    Else
        cntA = 0
    End If

    DictSet info, "dataRows", nr - hr
    DictSet info, "readRows", nRead
    DictSet info, "blankRows", nBlank
    DictSet info, "errRows", nErr
    DictSet info, "outRows", n
    DictSet info, "countA", cntA
    DictSet info, "hasFilter", hasFilter
    DictSet info, "hasName", (cName > 0)
    DictSet info, "hasQty", (cQty > 0)
    ReadBomSheet = BufToRows(buf, n, BC_N)
End Function
