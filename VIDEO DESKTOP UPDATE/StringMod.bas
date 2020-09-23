Attribute VB_Name = "StringMod"
Public Function GetFileName(Path As String, _
                            ByVal Extension As Boolean) As String
Dim anuku As Long
    GetFileName = getRight(Path, 1)
    If Not Extension Then
        anuku = InStrRev(GetFileName, ".")
        If anuku <> 0 Then
            GetFileName = left$(GetFileName, anuku - 1)
        End If
    End If
End Function
Private Function getRight(Key As String, _
                          ByVal length As Long) As String
Dim anumu As Long
Dim I     As Long
    anumu = Len(Key)
    For I = 1 To length
        anumu = InStrRev(Key, "\", anumu - 1)
        If anumu = 0 Then
            Exit For
        End If
    Next I
    getRight = right$(Key, Len(Key) - anumu)
End Function

