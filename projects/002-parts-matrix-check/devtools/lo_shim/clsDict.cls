Option Explicit
' LibreOffice で試すときだけ使う、辞書の代わり (Excel では使わない)。
' Linux には Scripting.Dictionary が無いため。
Private mK() As String
Private mV() As Variant
Private mN As Long

Private Sub Class_Initialize()
    mN = 0
    ReDim mK(0 To 15)
    ReDim mV(0 To 15)
End Sub

Private Function Find(ByVal k As String) As Long
    Dim i As Long
    Find = -1
    For i = 0 To mN - 1
        If StrComp(mK(i), k, vbBinaryCompare) = 0 Then
            Find = i
            Exit Function
        End If
    Next i
End Function

Public Function Exists(ByVal k As String) As Boolean
    Exists = (Find(k) >= 0)
End Function

Public Sub Add(ByVal k As String, ByVal v As Variant)
    If mN > UBound(mK) Then
        ReDim Preserve mK(0 To mN * 2)
        ReDim Preserve mV(0 To mN * 2)
    End If
    mK(mN) = k
    mV(mN) = v
    mN = mN + 1
End Sub

Public Property Get Item(ByVal k As String) As Variant
    Dim i As Long
    i = Find(k)
    If i >= 0 Then Item = mV(i)
End Property

Public Sub SetItem(ByVal k As String, ByVal v As Variant)
    Dim i As Long
    i = Find(k)
    If i >= 0 Then
        mV(i) = v
    Else
        Add k, v
    End If
End Sub

Public Property Get Count() As Long
    Count = mN
End Property

Public Function Keys() As Variant
    Dim a() As Variant, i As Long
    If mN = 0 Then
        Keys = Array()
        Exit Function
    End If
    ReDim a(0 To mN - 1)
    For i = 0 To mN - 1
        a(i) = mK(i)
    Next i
    Keys = a
End Function
