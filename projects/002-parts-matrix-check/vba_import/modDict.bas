Attribute VB_Name = "modDict"
Option Explicit
' ============================================================
' modDict : 辞書 (名前から値を引く入れ物) を作る所
'   Windows に最初から入っている Scripting.Dictionary を使う。
'   参照設定は足さない。大文字と小文字、全角と半角は別の物として扱う。
'   辞書は Exists / Add / Item(読む) / Keys / Count だけを使い、
'   書き換えは必ず DictSet を通す。
' ============================================================

Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

' 値を入れる (無ければ足し、あれば書き換える)
Public Sub DictSet(ByVal d As Object, ByVal k As String, ByVal v As Variant)
    d.Item(k) = v
End Sub
