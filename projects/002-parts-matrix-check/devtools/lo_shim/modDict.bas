Option Explicit
' LibreOffice で試すときだけ使う modDict (Excel では vba/modDict.bas を使う)
Public Function NewDict() As Object
    Set NewDict = New clsDict
End Function

Public Sub DictSet(ByVal d As Object, ByVal k As String, ByVal v As Variant)
    d.SetItem k, v
End Sub
