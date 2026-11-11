Option Explicit
' ============================================================
' modCompare : 照合 (共通の形に直した 2 つの表を比べる)
'   入力は配列だけ。シートやファイルには触らない。
'   そのため、同じ入力なら必ず同じ結果になり、テストもしやすい。
'
'   結果は 3 つ出す。
'     1. 食い違いの一覧
'     2. 消し込み: 部品表の全部の行と、マトリックスの全部のセルに付けた状態
'        (紙に赤ペンで印を付けるのと同じ。印の無いものが残らないことを確かめる)
'     3. 検算のための数
' ============================================================

' ---- 結果の返し方 ----
'   一覧は 3 つの引数で返す。
'     findings  : 食い違いの一覧 (行, FC_N)。無ければ Empty
'     bomCheck  : 部品表の行ごとの消し込み (行, 2)。1 = 状態、2 = 説明
'     cellCheck : マトリックスのセルごとの消し込み (行, 2)。1 = 状態、2 = 説明
'   数や文字は辞書 (res) で返す。名前と中身は次のとおり。
'   NFind / NNg / NWarn : 一覧の行数 / 不一致の件数 / 警告の件数
'   Signature     : 不一致の内容をつないだ文字 (結果が変わったかを見るのに使う)
'   PartNg / PartWarn / NameNg : 部品表に無い品番 (N-2) / 警告のある品番 (W-10) / 名称が違う品番 (N-3)。
'                                品番キーをタブで囲んで並べた文字
'   検算のための数:
'   MCells, MOk, MNg, MUnk, MOut : マトリックスのセルの数と、状態ごとの数
'   BRows, BOk, BNg, BUnk, BOut  : 部品表の行の数 (品番 × 仕様) と、状態ごとの数
'   BPairsT   : 部品表の「品番 × 仕様」の組 (対象の仕様で、品番がマトリックスにあるもの。重なりは除く)
'   BPairsOk  : そのうち、マトリックスも「あり」だった組 (部品表の側から数えた)
'   BPairsUnk : そのうち、マトリックスが「判断できない」だった組
'   MPairsYes : マトリックスの「あり」の組 (品番が部品表にあるもの。重なりは除く)
'   MPairsOk  : そのうち、部品表にもあった組 (マトリックスの側から数えた)
'   FN4 / FN5 : 一覧に出した N-4 / N-5 の件数
'   まとめに書くこと:
'   SpecsT / SpecsOutside / HeadsExcluded : 対象の仕様 / 比べなかった部品表の仕様 / 対象外にした見出し (改行区切り)
'   SkippedN1 / SkippedN3 / SkippedN6     : その項目を調べなかったら True

Private mF() As Variant      ' 一覧をためる入れ物 (列, 行)
Private mFOrd() As Long      ' 並べる順番
Private mFN As Long
Private mFCap As Long

Public Function CodeTitle(ByVal code As String) As String
    Select Case code
        Case "N-1": CodeTitle = "マトリックスに品番が無い"
        Case "N-2": CodeTitle = "部品表に品番が無い"
        Case "N-3": CodeTitle = "名称が違う"
        Case "N-4": CodeTitle = "マトリックスに印が無い"
        Case "N-5": CodeTitle = "部品表に登録が無い"
        Case "N-6": CodeTitle = "員数が違う"
        Case "W-1": CodeTitle = "品番の重複"
        Case "W-2": CodeTitle = "仕様の重なり"
        Case "W-3": CodeTitle = "判断できない値"
        Case "W-4": CodeTitle = "品番が数値やエラー"
        Case "W-5": CodeTitle = "非表示"
        Case "W-6": CodeTitle = "見出しが空"
        Case "W-7": CodeTitle = "部品表の重複"
        Case "W-8": CodeTitle = "取り消し線"
        Case "W-9": CodeTitle = "品番の形に合わない"
        Case "W-10": CodeTitle = "絞り込みの外"
        Case "W-11": CodeTitle = "範囲の先に続きがある"
        Case Else: CodeTitle = code
    End Select
End Function

' 一覧に並べる順番 (不一致が先。その中は記号の順)
Public Function CodeOrder(ByVal code As String) As Long
    Dim n As Long
    n = 0
    If IsNumeric(Mid(code, 3)) Then n = CLng(Mid(code, 3))
    If Left(code, 1) = "N" Then
        CodeOrder = n
    Else
        CodeOrder = 20 + n
    End If
End Function

Public Function CodeList() As Variant
    CodeList = Array("N-1", "N-2", "N-3", "N-4", "N-5", "N-6", "W-1", "W-2", "W-3", "W-4", "W-5", _
                     "W-6", "W-7", "W-8", "W-9", "W-10", "W-11")
End Function

Private Sub FReset()
    mFCap = 64
    ReDim mF(1 To FC_N, 1 To mFCap)
    ReDim mFOrd(1 To mFCap)
    mFN = 0
End Sub

Private Sub FAdd(ByVal code As String, ByVal part As String, ByVal nameB As String, ByVal nameM As String, _
        ByVal spec As String, ByVal head As String, ByVal valB As String, ByVal valM As String, _
        ByVal cellAddr As String, ByVal note As String)
    If mFN >= mFCap Then
        mFCap = mFCap * 2
        ReDim Preserve mF(1 To FC_N, 1 To mFCap)
        ReDim Preserve mFOrd(1 To mFCap)
    End If
    mFN = mFN + 1
    If Left(code, 1) = "N" Then mF(FC_CLASS, mFN) = CLASS_NG Else mF(FC_CLASS, mFN) = CLASS_WARN
    mF(FC_CODE, mFN) = code
    mF(FC_TITLE, mFN) = CodeTitle(code)
    mF(FC_PART, mFN) = part
    mF(FC_NAME_B, mFN) = nameB
    mF(FC_NAME_M, mFN) = nameM
    mF(FC_SPEC, mFN) = spec
    mF(FC_HEAD, mFN) = head
    mF(FC_VAL_B, mFN) = valB
    mF(FC_VAL_M, mFN) = valM
    mF(FC_CELL, mFN) = cellAddr
    mF(FC_NOTE, mFN) = note
    mFOrd(mFN) = CodeOrder(code)
End Sub

Private Function WithQty(ByVal word As String, ByVal qty As String) As String
    If Len(qty) > 0 Then
        WithQty = word & " (員数 " & qty & ")"
    Else
        WithQty = word
    End If
End Function

Private Function Worse(ByVal a As String, ByVal b As String) As String
    ' 状態の悪い方を返す (不一致 > 判断できない > 一致)
    If a = S_NG Or b = S_NG Then
        Worse = S_NG
    ElseIf a = S_UNK Or b = S_UNK Then
        Worse = S_UNK
    Else
        Worse = S_OK
    End If
End Function

' ============================================================
' 照合の本体
'   bom       : 読取_部品表 の配列 (行, BC_N)。無ければ Empty
'   mtx       : 読取_マトリックス の配列 (行, MC_N)。無ければ Empty
'   map       : 仕様対応表 (行, 4)。modSpecMap.LoadSpecMap の戻り値
'   warn      : 読取_警告 の配列 (行, WC_N)。無ければ Empty
'   hasFilter : 部品表の絞り込みが設定されているか (されていなければ N-1 を調べない)
'   cmpName   : 名称を比べるか
'   cmpQty    : 員数を比べるか
'   res       : 数や文字の結果を入れて返す辞書
'   findings / bomCheck / cellCheck : 一覧の結果 (上の説明のとおり)
' ============================================================
Public Sub CompareCore(ByRef bom As Variant, ByRef mtx As Variant, ByRef map As Variant, ByRef warn As Variant, _
        ByVal hasFilter As Boolean, ByVal cmpName As Boolean, ByVal cmpQty As Boolean, ByVal res As Object, _
        ByRef findings As Variant, ByRef bomCheck As Variant, ByRef cellCheck As Variant)
    Dim nB As Long, nM As Long, nP As Long, nW As Long
    Dim i As Long, j As Long, k As Long, ord As Long
    Dim headSpecs As Object, headExcl As Object, specText As Object, tSpec As Object
    Dim bomPart As Object, bomIn As Object, bomPair As Object
    Dim mtxPart As Object, pairYes As Object, pairUnk As Object, headCols As Object
    Dim done As Object, n1 As Object, outSpecs As Object
    Dim h As String, sk As String, key As String, pk As String, mv As String, st As String, note As String
    Dim specs As Variant, keys As Variant
    Dim bRow As Long, bq As String, mq As String, nameB As String
    Dim bChk() As Variant, cChk() As Variant
    Dim out() As Variant, sig As String, first As Boolean, t As String
    Dim rsNFind As Long, rsNNg As Long, rsNWarn As Long, rsMCells As Long, rsMOk As Long, rsMNg As Long, rsMUnk As Long, rsMOut As Long, rsBRows As Long, rsBOk As Long
    Dim rsBNg As Long, rsBUnk As Long, rsBOut As Long, rsBPairsT As Long, rsBPairsOk As Long, rsBPairsUnk As Long, rsMPairsYes As Long, rsMPairsOk As Long, rsFN4 As Long, rsFN5 As Long
    Dim rsSignature As String, rsPartNg As String, rsPartWarn As String, rsNameNg As String, rsSpecsT As String, rsSpecsOutside As String, rsHeadsExcluded As String
    Dim rsSkippedN1 As Boolean, rsSkippedN3 As Boolean, rsSkippedN6 As Boolean

    nB = RowCount(bom): nM = RowCount(mtx): nP = RowCount(map): nW = RowCount(warn)
    FReset

    ' ---- 1. 仕様対応表を引きやすい形にする ----
    Set headSpecs = NewDict()     ' 見出しの文字 → 仕様コード (複数ならタブ区切り)
    Set headExcl = NewDict()      ' 対象外の見出し
    Set specText = NewDict()      ' 仕様コード (そろえた形) → 元の文字
    Set tSpec = NewDict()         ' 対象の仕様 → それに当たる見出し (改行区切り)
    For i = 1 To nP
        If map(i, 3) = True Then
            If Not headExcl.Exists(map(i, 1)) Then headExcl.Add map(i, 1), True
        End If
    Next i
    For i = 1 To nP
        h = map(i, 1)
        sk = map(i, 2)
        If Not headExcl.Exists(h) And Len(sk) > 0 Then
            If headSpecs.Exists(h) Then
                If InStr(vbTab & headSpecs.Item(h) & vbTab, vbTab & sk & vbTab) = 0 Then
                    DictSet headSpecs, h, headSpecs.Item(h) & vbTab & sk
                End If
            Else
                headSpecs.Add h, sk
            End If
            If Not specText.Exists(sk) Then specText.Add sk, map(i, 4)
            If tSpec.Exists(sk) Then
                If InStr(vbLf & tSpec.Item(sk) & vbLf, vbLf & h & vbLf) = 0 Then
                    DictSet tSpec, sk, tSpec.Item(sk) & vbLf & h
                End If
            Else
                tSpec.Add sk, h
            End If
        End If
    Next i

    ' ---- 2. 部品表を引きやすい形にする ----
    Set bomPart = NewDict()       ' 品番キー → 最初の行
    Set bomIn = NewDict()         ' 絞り込みの中にある品番キー
    Set bomPair = NewDict()       ' 品番キー + 仕様 → 最初の行
    Set outSpecs = NewDict()      ' 対象の仕様に入らなかった仕様
    For j = 1 To nB
        key = ToText(bom(j, BC_KEY))
        sk = ToText(bom(j, BC_SPECKEY))
        If Not bomPart.Exists(key) Then bomPart.Add key, j
        If Len(ToText(bom(j, BC_INSCOPE))) > 0 Then
            If Not bomIn.Exists(key) Then bomIn.Add key, True
        End If
        If Len(sk) > 0 Then
            pk = key & vbTab & sk
            If Not bomPair.Exists(pk) Then bomPair.Add pk, j
            If Not tSpec.Exists(sk) Then
                If Not outSpecs.Exists(sk) Then outSpecs.Add sk, ToText(bom(j, BC_SPEC))
            End If
        End If
    Next j

    ' ---- 3. マトリックスを引きやすい形にする ----
    Set mtxPart = NewDict()       ' 品番キー → 最初の行
    Set pairYes = NewDict()       ' 品番キー + 仕様 → 「あり」のセルの行
    Set pairUnk = NewDict()       ' 品番キー + 仕様 → 「判断できない」のセルの行
    Set headCols = NewDict()      ' 見出しの文字 → その見出しのセル (複数ならタブ区切り)
    For i = 1 To nM
        key = ToText(mtx(i, MC_KEY))
        h = ToText(mtx(i, MC_HEAD))
        mv = ToText(mtx(i, MC_VAL))
        If Not mtxPart.Exists(key) Then mtxPart.Add key, i
        t = ToText(mtx(i, MC_HEADCELL))
        If headCols.Exists(h) Then
            If InStr(vbTab & headCols.Item(h) & vbTab, vbTab & t & vbTab) = 0 Then
                DictSet headCols, h, headCols.Item(h) & vbTab & t
            End If
        Else
            headCols.Add h, t
        End If
        If Not headExcl.Exists(h) Then
            If headSpecs.Exists(h) Then
                specs = Split(headSpecs.Item(h), vbTab)
                For k = LBound(specs) To UBound(specs)
                    pk = key & vbTab & specs(k)
                    If mv = V_YES Then
                        If Not pairYes.Exists(pk) Then pairYes.Add pk, i
                    ElseIf mv = V_UNK Then
                        If Not pairUnk.Exists(pk) Then pairUnk.Add pk, i
                    End If
                Next k
            End If
        End If
    Next i

    ' ---- 4. マトリックスのセルを 1 つずつ確かめる (マトリックス側の消し込み) ----
    rsMCells = nM: rsMOk = 0: rsMNg = 0: rsMUnk = 0: rsMOut = 0
    rsFN4 = 0: rsFN5 = 0
    Set done = NewDict()
    If nM > 0 Then
        ReDim cChk(1 To nM, 1 To 2)
    End If
    For i = 1 To nM
        key = ToText(mtx(i, MC_KEY))
        h = ToText(mtx(i, MC_HEAD))
        mv = ToText(mtx(i, MC_VAL))
        note = ""
        If headExcl.Exists(h) Then
            st = S_OUT
            note = "対象外にした見出し"
        ElseIf Not headSpecs.Exists(h) Then
            st = S_UNK
            note = "見出しの対応が決まっていない"
        ElseIf mv = V_UNK Then
            st = S_UNK
            note = "印の一覧に無い値"
        Else
            st = S_OK
            specs = Split(headSpecs.Item(h), vbTab)
            For k = LBound(specs) To UBound(specs)
                sk = specs(k)
                pk = key & vbTab & sk
                If mv = V_YES Then
                    If bomPair.Exists(pk) Then
                        bRow = bomPair.Item(pk)
                        bq = ToText(bom(bRow, BC_QTY))
                        mq = ToText(mtx(i, MC_QTY))
                        If cmpQty And IsNumeric(bq) And IsNumeric(mq) Then
                            If CDbl(bq) <> CDbl(mq) Then
                                st = S_NG
                                note = "員数が違う (N-6)"
                                If Not done.Exists("N-6" & vbTab & pk) Then
                                    done.Add "N-6" & vbTab & pk, True
                                    FAdd "N-6", ToText(mtx(i, MC_PART)), ToText(bom(bRow, BC_NAME)), ToText(mtx(i, MC_NAME)), _
                                        specText.Item(sk), h, "員数 " & bq, "員数 " & mq, ToText(mtx(i, MC_VALCELL)), _
                                        "員数が、部品表は " & bq & "、マトリックスは " & mq
                                End If
                            End If
                        End If
                    Else
                        st = S_NG
                        If bomPart.Exists(key) Then
                            note = "部品表に登録が無い (N-5)"
                            If Not done.Exists("N-5" & vbTab & pk) Then
                                done.Add "N-5" & vbTab & pk, True
                                rsFN5 = rsFN5 + 1
                                bRow = bomPart.Item(key)
                                FAdd "N-5", ToText(mtx(i, MC_PART)), ToText(bom(bRow, BC_NAME)), ToText(mtx(i, MC_NAME)), _
                                    specText.Item(sk), h, V_NO, WithQty(V_YES, ToText(mtx(i, MC_QTY))), ToText(mtx(i, MC_VALCELL)), _
                                    "マトリックスでは「あり」だが、部品表ではこの仕様で使っていない"
                            End If
                        Else
                            note = "部品表に品番が無い (N-2)"
                        End If
                    End If
                Else
                    ' マトリックスは「なし」
                    If bomPair.Exists(pk) Then
                        If pairYes.Exists(pk) Then
                            ' 別のセル (同じ仕様に当たる別の列、または重複した品番) で「あり」
                            If Len(note) = 0 Then note = "別のセルで「あり」"
                        ElseIf pairUnk.Exists(pk) Then
                            st = Worse(st, S_UNK)
                            If Len(note) = 0 Then note = "別のセルが「判断できない」"
                        Else
                            st = S_NG
                            note = "マトリックスに印が無い (N-4)"
                            If Not done.Exists("N-4" & vbTab & pk) Then
                                done.Add "N-4" & vbTab & pk, True
                                rsFN4 = rsFN4 + 1
                                bRow = bomPair.Item(pk)
                                FAdd "N-4", ToText(mtx(i, MC_PART)), ToText(bom(bRow, BC_NAME)), ToText(mtx(i, MC_NAME)), _
                                    specText.Item(sk), h, WithQty(V_YES, ToText(bom(bRow, BC_QTY))), V_NO, ToText(mtx(i, MC_VALCELL)), _
                                    "部品表ではこの仕様で使うが、マトリックスに印が無い"
                            End If
                        End If
                    End If
                End If
            Next k
        End If
        cChk(i, 1) = st
        cChk(i, 2) = note
        Select Case st
            Case S_OK: rsMOk = rsMOk + 1
            Case S_NG: rsMNg = rsMNg + 1
            Case S_UNK: rsMUnk = rsMUnk + 1
            Case Else: rsMOut = rsMOut + 1
        End Select
    Next i

    ' ---- 5. 部品表の行を 1 つずつ確かめる (部品表側の消し込み) ----
    rsBRows = nB: rsBOk = 0: rsBNg = 0: rsBUnk = 0: rsBOut = 0
    rsBPairsT = 0: rsBPairsOk = 0: rsBPairsUnk = 0
    Set n1 = NewDict()            ' N-1 の品番キー → 使う仕様 (元の文字を並べたもの)
    If nB > 0 Then
        ReDim bChk(1 To nB, 1 To 2)
    End If
    For j = 1 To nB
        key = ToText(bom(j, BC_KEY))
        sk = ToText(bom(j, BC_SPECKEY))
        note = ""
        If Len(sk) = 0 Then
            st = S_OUT
            note = "仕様が空"
        ElseIf Not tSpec.Exists(sk) Then
            st = S_OUT
            note = "マトリックスに無い仕様"
        ElseIf Not mtxPart.Exists(key) Then
            If Not hasFilter Then
                st = S_OUT
                note = "マトリックスに無い品番 (絞り込みが未設定のため、書きもれかどうかは調べていない)"
            ElseIf Len(ToText(bom(j, BC_INSCOPE))) > 0 Then
                st = S_NG
                note = "マトリックスに品番が無い (N-1)"
                If n1.Exists(key) Then
                    DictSet n1, key, n1.Item(key) & ", " & ToText(bom(j, BC_SPEC))
                Else
                    n1.Add key, ToText(bom(j, BC_SPEC))
                End If
            Else
                st = S_OUT
                note = "絞り込みの外"
            End If
        Else
            pk = key & vbTab & sk
            first = (bomPair.Item(pk) = j)
            If first Then rsBPairsT = rsBPairsT + 1
            If pairYes.Exists(pk) Then
                st = S_OK
                If first Then rsBPairsOk = rsBPairsOk + 1
                bq = ToText(bom(j, BC_QTY))
                mq = ToText(mtx(pairYes.Item(pk), MC_QTY))
                If cmpQty And IsNumeric(bq) And IsNumeric(mq) Then
                    If CDbl(bq) <> CDbl(mq) Then
                        st = S_NG
                        note = "員数が違う (N-6)"
                    End If
                End If
            ElseIf pairUnk.Exists(pk) Then
                st = S_UNK
                note = "マトリックスの値が判断できない"
                If first Then rsBPairsUnk = rsBPairsUnk + 1
            Else
                st = S_NG
                note = "マトリックスに印が無い (N-4)"
            End If
        End If
        bChk(j, 1) = st
        bChk(j, 2) = note
        Select Case st
            Case S_OK: rsBOk = rsBOk + 1
            Case S_NG: rsBNg = rsBNg + 1
            Case S_UNK: rsBUnk = rsBUnk + 1
            Case Else: rsBOut = rsBOut + 1
        End Select
    Next j

    ' ---- 6. 品番ごとの食い違い ----
    ' N-1: 部品表にあるのに、マトリックスに品番が無い
    If n1.Count > 0 Then
        keys = n1.Keys
        For k = LBound(keys) To UBound(keys)
            key = keys(k)
            bRow = bomPart.Item(key)
            FAdd "N-1", ToText(bom(bRow, BC_PART)), ToText(bom(bRow, BC_NAME)), "", n1.Item(key), "", V_YES, "品番が無い", "", _
                "部品表では対象の仕様で使っているが、マトリックスに品番が無い (部品表の " & ToText(bom(bRow, BC_ROW)) & " 行目など)"
        Next k
    End If
    ' N-2 / W-10 / N-3
    rsPartNg = vbTab: rsPartWarn = vbTab: rsNameNg = vbTab
    If mtxPart.Count > 0 Then
        keys = mtxPart.Keys
        For k = LBound(keys) To UBound(keys)
            key = keys(k)
            i = mtxPart.Item(key)
            If Not bomPart.Exists(key) Then
                rsPartNg = rsPartNg & key & vbTab
                FAdd "N-2", ToText(mtx(i, MC_PART)), "", ToText(mtx(i, MC_NAME)), "", "", "品番が無い", "品番がある", ToText(mtx(i, MC_PARTCELL)), _
                    "マトリックスにある品番が、部品表に無い"
            Else
                bRow = bomPart.Item(key)
                nameB = ToText(bom(bRow, BC_NAME))
                If hasFilter And Not bomIn.Exists(key) Then
                    rsPartWarn = rsPartWarn & key & vbTab
                    FAdd "W-10", ToText(mtx(i, MC_PART)), nameB, ToText(mtx(i, MC_NAME)), "", "", "", "", ToText(mtx(i, MC_PARTCELL)), _
                        "マトリックスの品番が、部品表の絞り込みの外にある (絞り込みの値が合っているか確かめること)"
                End If
                If cmpName Then
                    If NormName(nameB) <> NormName(ToText(mtx(i, MC_NAME))) Then
                        rsNameNg = rsNameNg & key & vbTab
                        FAdd "N-3", ToText(mtx(i, MC_PART)), nameB, ToText(mtx(i, MC_NAME)), "", "", nameB, ToText(mtx(i, MC_NAME)), ToText(mtx(i, MC_NAMECELL)), _
                            "名称が、部品表とマトリックスで違う"
                    End If
                End If
            End If
        Next k
    End If

    ' ---- 7. 警告 ----
    ' W-2: 複数の見出しが同じ仕様コードに当たっている / 同じ見出しの列が複数ある
    If tSpec.Count > 0 Then
        keys = tSpec.Keys
        For k = LBound(keys) To UBound(keys)
            If InStr(tSpec.Item(keys(k)), vbLf) > 0 Then
                FAdd "W-2", "", "", "", specText.Item(keys(k)), Replace(tSpec.Item(keys(k)), vbLf, " ; "), "", "", "", _
                    "複数の見出しが同じ仕様コードに当たっている。どれかで「あり」なら「あり」として扱った"
            End If
        Next k
    End If
    If headCols.Count > 0 Then
        keys = headCols.Keys
        For k = LBound(keys) To UBound(keys)
            If InStr(headCols.Item(keys(k)), vbTab) > 0 And Not headExcl.Exists(keys(k)) Then
                FAdd "W-2", "", "", "", "", keys(k), "", "", Replace(headCols.Item(keys(k)), vbTab, ", "), _
                    "同じ見出しの文字の列が複数ある。どれかで「あり」なら「あり」として扱った"
            End If
        Next k
    End If
    ' 読み取りのときにためた警告 (対象外にした見出しの「判断できない値」は出さない)
    For i = 1 To nW
        t = ToText(warn(i, WC_CODE))
        If Not (t = "W-3" And headExcl.Exists(ToText(warn(i, WC_HEAD)))) Then
            FAdd t, ToText(warn(i, WC_PART)), "", "", "", ToText(warn(i, WC_HEAD)), "", ToText(warn(i, WC_RAW)), _
                IIf(ToText(warn(i, WC_SRC)) = SRC_MTX, ToText(warn(i, WC_CELL)), ""), _
                IIf(ToText(warn(i, WC_SRC)) = SRC_BOM, "部品表の " & ToText(warn(i, WC_CELL)) & ": ", "") & ToText(warn(i, WC_NOTE))
        End If
    Next i

    ' ---- 8. 検算のための数 (マトリックスの側から数え直す) ----
    rsMPairsYes = 0: rsMPairsOk = 0
    If pairYes.Count > 0 Then
        keys = pairYes.Keys
        For k = LBound(keys) To UBound(keys)
            pk = keys(k)
            key = Left(pk, InStr(pk, vbTab) - 1)
            If bomPart.Exists(key) Then
                rsMPairsYes = rsMPairsYes + 1
                If bomPair.Exists(pk) Then rsMPairsOk = rsMPairsOk + 1
            End If
        Next k
    End If

    ' ---- 9. 一覧を、区分 → 記号 → 表での並び順 に並べて返す ----
    rsNFind = mFN: rsNNg = 0: rsNWarn = 0
    sig = ""
    If mFN = 0 Then
        findings = Empty
    Else
        ReDim out(1 To mFN, 1 To FC_N)
        k = 0
        For ord = 1 To 40
            For i = 1 To mFN
                If mFOrd(i) = ord Then
                    k = k + 1
                    For j = 1 To FC_N
                        out(k, j) = mF(j, i)
                    Next j
                    If mF(FC_CLASS, i) = CLASS_NG Then
                        rsNNg = rsNNg + 1
                        sig = sig & mF(FC_CODE, i) & "|" & NormPart(CStr(mF(FC_PART, i))) & "|" & mF(FC_SPEC, i) & "|" & mF(FC_CELL, i) & vbLf
                    Else
                        rsNWarn = rsNWarn + 1
                    End If
                End If
            Next i
        Next ord
        findings = out
    End If
    rsSignature = sig
    If nB > 0 Then
        bomCheck = bChk
    Else
        bomCheck = Empty
    End If
    If nM > 0 Then
        cellCheck = cChk
    Else
        cellCheck = Empty
    End If

    ' ---- 10. まとめに書くこと ----
    rsSpecsT = JoinItems(specText)
    rsSpecsOutside = JoinItems(outSpecs)
    t = ""
    If headExcl.Count > 0 Then
        keys = headExcl.Keys
        For k = LBound(keys) To UBound(keys)
            If headCols.Exists(keys(k)) Then
                If Len(t) > 0 Then t = t & vbLf
                t = t & keys(k)
            End If
        Next k
    End If
    rsHeadsExcluded = t
    rsSkippedN1 = Not hasFilter
    rsSkippedN3 = Not cmpName
    rsSkippedN6 = Not cmpQty

    ' ---- 11. 結果を辞書に入れて返す ----
    DictSet res, "NFind", rsNFind
    DictSet res, "NNg", rsNNg
    DictSet res, "NWarn", rsNWarn
    DictSet res, "MCells", rsMCells
    DictSet res, "MOk", rsMOk
    DictSet res, "MNg", rsMNg
    DictSet res, "MUnk", rsMUnk
    DictSet res, "MOut", rsMOut
    DictSet res, "BRows", rsBRows
    DictSet res, "BOk", rsBOk
    DictSet res, "BNg", rsBNg
    DictSet res, "BUnk", rsBUnk
    DictSet res, "BOut", rsBOut
    DictSet res, "BPairsT", rsBPairsT
    DictSet res, "BPairsOk", rsBPairsOk
    DictSet res, "BPairsUnk", rsBPairsUnk
    DictSet res, "MPairsYes", rsMPairsYes
    DictSet res, "MPairsOk", rsMPairsOk
    DictSet res, "FN4", rsFN4
    DictSet res, "FN5", rsFN5
    DictSet res, "Signature", rsSignature
    DictSet res, "PartNg", rsPartNg
    DictSet res, "PartWarn", rsPartWarn
    DictSet res, "NameNg", rsNameNg
    DictSet res, "SpecsT", rsSpecsT
    DictSet res, "SpecsOutside", rsSpecsOutside
    DictSet res, "HeadsExcluded", rsHeadsExcluded
    DictSet res, "SkippedN1", rsSkippedN1
    DictSet res, "SkippedN3", rsSkippedN3
    DictSet res, "SkippedN6", rsSkippedN6
End Sub

Private Function JoinItems(ByVal d As Object) As String
    Dim keys As Variant, k As Long, s As String
    s = ""
    If d.Count > 0 Then
        keys = d.Keys
        For k = LBound(keys) To UBound(keys)
            If Len(s) > 0 Then s = s & vbLf
            s = s & d.Item(keys(k))
        Next k
    End If
    JoinItems = s
End Function

' ============================================================
' 抜き打ちの確かめ
'   本物のデータのうえで、マトリックスの値を 1 か所だけわざと逆にして照合し直し、
'   結果が変わる (= 照合が気付く) ことを確かめる。これを何か所かで繰り返す。
'   終わったら値は元に戻す。
'     nTried    : 試した数
'     nDetected : 気付いた数 (nTried と同じなら合格)
'     detail    : 試した所の一覧 (改行区切り)
' ============================================================
Public Sub MutationCheck(ByRef bom As Variant, ByRef mtx As Variant, ByRef map As Variant, ByRef warn As Variant, _
        ByVal hasFilter As Boolean, ByVal cmpName As Boolean, ByVal cmpQty As Boolean, ByVal baseRes As Object, _
        ByRef nTried As Long, ByRef nDetected As Long, ByRef detail As String)
    Dim nM As Long, nP As Long, i As Long, k As Long, j As Long
    Dim headOne As Object, headMany As Object, headExcl As Object, cellCount As Object, bomParts As Object
    Dim h As String, pk As String, oldVal As String, newVal As String
    Dim cand() As Long, nCand As Long, nPick As Long, idx As Long
    Dim res2 As Object, hit As Boolean, baseSig As String
    Dim f2 As Variant, b2 As Variant, c2 As Variant

    nTried = 0: nDetected = 0: detail = ""
    nM = RowCount(mtx): nP = RowCount(map)
    If nM = 0 Or RowCount(bom) = 0 Then Exit Sub
    baseSig = baseRes.Item("Signature")
    Set res2 = NewDict()

    ' 仕様コードが 1 つだけに決まる見出しを集める
    Set headOne = NewDict(): Set headMany = NewDict(): Set headExcl = NewDict()
    For i = 1 To nP
        If map(i, 3) = True Then
            If Not headExcl.Exists(map(i, 1)) Then headExcl.Add map(i, 1), True
        End If
    Next i
    For i = 1 To nP
        h = map(i, 1)
        If Not headExcl.Exists(h) And Len(map(i, 2)) > 0 Then
            If headOne.Exists(h) Then
                If headOne.Item(h) <> map(i, 2) Then
                    If Not headMany.Exists(h) Then headMany.Add h, True
                End If
            Else
                headOne.Add h, map(i, 2)
            End If
        End If
    Next i
    Set bomParts = NewDict()
    For j = 1 To RowCount(bom)
        If Not bomParts.Exists(ToText(bom(j, BC_KEY))) Then bomParts.Add ToText(bom(j, BC_KEY)), True
    Next j
    ' 同じ「品番 × 仕様」に当たるセルの数
    Set cellCount = NewDict()
    For i = 1 To nM
        h = ToText(mtx(i, MC_HEAD))
        If headOne.Exists(h) And Not headMany.Exists(h) Then
            pk = ToText(mtx(i, MC_KEY)) & vbTab & headOne.Item(h)
            If cellCount.Exists(pk) Then
                DictSet cellCount, pk, cellCount.Item(pk) + 1
            Else
                cellCount.Add pk, 1
            End If
        End If
    Next i
    ' 試せるセル: 仕様が 1 つに決まり、その組に当たるセルが 1 つだけで、品番が部品表にあり、値が あり か なし
    ReDim cand(1 To nM)
    nCand = 0
    For i = 1 To nM
        h = ToText(mtx(i, MC_HEAD))
        If headOne.Exists(h) And Not headMany.Exists(h) Then
            pk = ToText(mtx(i, MC_KEY)) & vbTab & headOne.Item(h)
            If cellCount.Item(pk) = 1 And bomParts.Exists(ToText(mtx(i, MC_KEY))) Then
                If ToText(mtx(i, MC_VAL)) = V_YES Or ToText(mtx(i, MC_VAL)) = V_NO Then
                    nCand = nCand + 1
                    cand(nCand) = i
                End If
            End If
        End If
    Next i
    If nCand = 0 Then Exit Sub
    nPick = MinL(MUTATION_COUNT, nCand)
    For k = 1 To nPick
        idx = cand(1 + Int((k - 1) * nCand / nPick))
        oldVal = ToText(mtx(idx, MC_VAL))
        If oldVal = V_YES Then newVal = V_NO Else newVal = V_YES
        mtx(idx, MC_VAL) = newVal
        CompareCore bom, mtx, map, warn, hasFilter, cmpName, cmpQty, res2, f2, b2, c2
        hit = (res2.Item("Signature") <> baseSig)
        mtx(idx, MC_VAL) = oldVal
        nTried = nTried + 1
        If hit Then nDetected = nDetected + 1
        If Len(detail) > 0 Then detail = detail & vbLf
        detail = detail & ToText(mtx(idx, MC_PART)) & " / " & ToText(mtx(idx, MC_HEAD)) & " (" & ToText(mtx(idx, MC_VALCELL)) & "): " & _
            oldVal & " → " & newVal & " に変えた → " & IIf(hit, "気付いた", "気付かなかった")
    Next k
End Sub
