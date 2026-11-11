Attribute VB_Name = "modText"
Option Explicit
' ============================================================
' modText : 文字のそろえ方 (比べるための下ごしらえ)
'   記号の一覧はシート「設定_文字」に置く。コードの中に記号を書くときは
'   ChrW(番号) で書く (VBA の編集画面が保存できない記号があるため)。
' ============================================================

Private mReady As Boolean
Private mHyphens As String          ' ハイフンとみなす文字を並べたもの
Private mYes As Object              ' あり を表す印 (そろえた形)
Private mNo As Object               ' なし を表す印 (そろえた形)
Private mIgnoreHyphen As Boolean    ' ハイフンを無視するか

' ---- 初期値 (設定_文字 の最初の中身。テストもこれを使う) ----
Public Function DefaultTextRows() As Variant
    Dim a(1 To 23, 1 To 3) As Variant, n As Long
    n = 0
    AddRow a, n, T_IGNORE, OPT_NO, "「する」にすると、品番のハイフンを全部取ってから比べる"
    AddRow a, n, T_HYPHEN, ChrW(&H2010), "ハイフン"
    AddRow a, n, T_HYPHEN, ChrW(&H2011), "改行しないハイフン"
    AddRow a, n, T_HYPHEN, ChrW(&H2012), "数字幅のダッシュ"
    AddRow a, n, T_HYPHEN, ChrW(&H2013), "短いダッシュ"
    AddRow a, n, T_HYPHEN, ChrW(&H2014), "長いダッシュ"
    AddRow a, n, T_HYPHEN, ChrW(&H2015), "横線"
    AddRow a, n, T_HYPHEN, ChrW(&H2212), "マイナス記号"
    AddRow a, n, T_HYPHEN, ChrW(&HFF0D), "全角のハイフン"
    AddRow a, n, T_HYPHEN, ChrW(&H30FC), "長音 (品番の中だけ、ハイフンとみなす)"
    AddRow a, n, T_HYPHEN, ChrW(&HFF70), "半角の長音"
    AddRow a, n, T_YES, ChrW(&H25CB), "白丸"
    AddRow a, n, T_YES, ChrW(&H25EF), "大きい丸"
    AddRow a, n, T_YES, ChrW(&H3007), "漢数字のゼロ (丸の代わりに使われる)"
    AddRow a, n, T_YES, ChrW(&H25CF), "黒丸"
    AddRow a, n, T_YES, ChrW(&H25CE), "二重丸"
    AddRow a, n, T_YES, ChrW(&H30EC), "カタカナのレ"
    AddRow a, n, T_YES, ChrW(&H2713), "チェックの印"
    AddRow a, n, T_YES, ChrW(&H2714), "太いチェックの印"
    AddRow a, n, T_NO, "-", "ハイフン (ハイフンとみなす文字も同じ扱い)"
    AddRow a, n, T_NO, ChrW(&HD7), "かける (バツ)"
    AddRow a, n, T_NO, ChrW(&H2715), "バツの印"
    AddRow a, n, T_NO, ChrW(&H2573), "斜めのバツ"
    DefaultTextRows = a
End Function

Private Sub AddRow(ByRef a() As Variant, ByRef n As Long, ByVal kind As String, ByVal ch As String, ByVal memo As String)
    n = n + 1
    a(n, 1) = kind
    a(n, 2) = ch
    a(n, 3) = memo
End Sub

' 初期値を使う
Public Sub TextDefaults()
    TextApply DefaultTextRows()
End Sub

' シート「設定_文字」から読む
Public Sub TextLoad()
    Dim v As Variant
    v = ReadTable(ToolSheet(SH_CFG_TXT), 2)
    If RowCount(v) = 0 Then
        TextDefaults
    Else
        TextApply v
    End If
End Sub

' 種類と文字の表 (1 列目 = 種類、2 列目 = 文字) を取り込む
Public Sub TextApply(ByVal tbl As Variant)
    Dim i As Long, kind As String, ch As String, k As String
    mHyphens = ""
    Set mYes = NewDict()
    Set mNo = NewDict()
    mIgnoreHyphen = False
    ' 先にハイフンを集める (印をそろえるのに使うため)
    For i = LBound(tbl, 1) To UBound(tbl, 1)
        kind = Trim(ToText(tbl(i, 1)))
        ch = ToText(tbl(i, 2))
        If kind = T_HYPHEN And Len(ch) > 0 Then
            mHyphens = mHyphens & ch
            k = NormBase(ch)
            If Len(k) > 0 And k <> "-" Then mHyphens = mHyphens & k
        ElseIf kind = T_IGNORE Then
            mIgnoreHyphen = (Trim(ch) = OPT_YES)
        End If
    Next i
    mReady = True
    For i = LBound(tbl, 1) To UBound(tbl, 1)
        kind = Trim(ToText(tbl(i, 1)))
        ch = ToText(tbl(i, 2))
        If Len(ch) > 0 Then
            k = NormMark(ch)
            If kind = T_YES Then
                If Len(k) > 0 And Not mYes.Exists(k) Then mYes.Add k, True
            ElseIf kind = T_NO Then
                If Len(k) > 0 And Not mNo.Exists(k) Then mNo.Add k, True
            End If
        End If
    Next i
End Sub

Public Sub SetIgnoreHyphen(ByVal onOff As Boolean)
    EnsureReady
    mIgnoreHyphen = onOff
End Sub

Private Sub EnsureReady()
    If Not mReady Then TextDefaults
End Sub

' ---- そろえ方 ----

' 共通: 改行を取る → 全角の英数字と記号を半角に → 大文字に → 空白を取る
Public Function NormBase(ByVal s As String) As String
    Dim t As String
    t = s
    If Len(t) = 0 Then
        NormBase = ""
        Exit Function
    End If
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = StrConv(t, vbNarrow, 1041)
    t = UCase(t)
    t = Replace(t, " ", "")
    t = Replace(t, ChrW(&H3000), "")
    t = Replace(t, ChrW(160), "")
    NormBase = t
End Function

Private Function HyphenToAscii(ByVal t As String) As String
    Dim i As Long
    For i = 1 To Len(mHyphens)
        t = Replace(t, Mid(mHyphens, i, 1), "-")
    Next i
    HyphenToAscii = t
End Function

' 品番: 共通の処理 → ハイフンとみなす文字を - に → (設定により) - を取る
Public Function NormPart(ByVal s As String) As String
    Dim t As String
    EnsureReady
    t = HyphenToAscii(NormBase(s))
    If mIgnoreHyphen Then t = Replace(t, "-", "")
    NormPart = t
End Function

' 名称: 共通の処理だけ (長音を壊さないため、ハイフンの処理はしない)
Public Function NormName(ByVal s As String) As String
    NormName = NormBase(s)
End Function

' 仕様の文字: 共通の処理だけ
Public Function NormSpec(ByVal s As String) As String
    NormSpec = NormBase(s)
End Function

' 印: 共通の処理 → ハイフンとみなす文字を - に
Public Function NormMark(ByVal s As String) As String
    EnsureReady
    NormMark = HyphenToAscii(NormBase(s))
End Function

' 値のセルを「あり / なし / 判断できない」に分ける。員数が分かれば qty に入れる
Public Function Classify(ByVal v As Variant, ByVal asQty As Boolean, ByRef qty As Variant) As String
    Dim t As String, n As Double
    EnsureReady
    qty = ""
    If IsError(v) Then
        Classify = V_UNK
        Exit Function
    End If
    t = NormMark(ToText(v))
    If Len(t) = 0 Then
        Classify = V_NO
    ElseIf mNo.Exists(t) Then
        Classify = V_NO
    ElseIf mYes.Exists(t) Then
        Classify = V_YES
    ElseIf asQty And IsNumeric(t) Then
        n = CDbl(t)
        If n > 0 Then
            Classify = V_YES
            qty = n
        ElseIf n = 0 Then
            Classify = V_NO
        Else
            Classify = V_UNK
        End If
    Else
        Classify = V_UNK
    End If
End Function
