Attribute VB_Name = "split_cell_by_language"
Public Function arabic(R As Range) As String
Dim L As Long, LLen As Long, v As String
arabic = ""
v = R.Value
LLen = Len(v)
For L = LLen To 1 Step -1
    If AscW(Mid(v, L, 1)) > 1000 Then
        arabic = Mid(v, 1, L)
        Exit Function
    End If
Next
End Function

Public Function english(R As Range)
Dim L As Long, LLen As Long, v As String
english = ""
v = R.Value
LLen = Len(v)
For L = LLen To 1 Step -1
    If AscW(Mid(v, L, 1)) > 1000 Then
        english = Mid(v, L + 1, 9999)
        Exit Function
    End If
Next
End Function
