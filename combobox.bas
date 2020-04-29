Attribute VB_Name = "Module1"
Private Sub T4_Change()
    If Me.T4.Value <> "" Then
        Dim Sh As Worksheet
        Set Sh = ThisWorkbook.Sheets("Database_newACT")
        Dim i As Integer
        
        If IsError(Application.Match(Me.T4.Value, Sh.Range("M:M"), 0)) = False Then
            i = Application.Match(Me.T4.Value, Sh.Range("M:M"), 0)
            
            Me.T5.Value = Sh.Range("O" & i).Value
            Me.T6.Value = Sh.Range("P" & i).Value
            
        End If
        
    
    End If
    
End Sub
