Attribute VB_Name = "Finish_NEW_format"
Sub Finish_Save_NEW()



    Dim username As String
    username = Application.username
        
    '-------------- Get Username -----------------'
    Dim getusername As String
    getusername = Replace(LCase(Application.username), " ", ".")
    '---------------------------------------------'
        
    
    '-------------- Save As ----------------------'
    Dim FileName1 As String
    Dim FileName2 As String
    
    '---------------KKS     ------------------------
    Dim kks As String
    With Selection.Find
        .Text = "aku.1227."
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    CommandBars("Navigation").Visible = False
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=5, Extend:=wdExtend
    kks = Selection.Text
    
    '---------------END-----------------------------
    
    '---------------project     ------------------------
    Dim project As String
    With Selection.Find
        .Text = "aku.1227."
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    CommandBars("Navigation").Visible = False
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=17, Extend:=wdExtend
    project = Replace(Selection.Text, ".0", "")
    
    '---------------END-----------------------------
    
    
    
    '--------------Act no--------------------------
    Dim act_no As String
    
    Selection.EndKey Unit:=wdStory
    Selection.MoveUp Unit:=wdParagraph, Count:=1
    Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    act_no = Selection.Text
    
    If InStr(1, act_no, "/", vbBinaryCompare) <> 0 Then
        act_no = Replace(act_no, "/", "-")
    End If
    
    '---------------END-----------------------------
    
    Dim Path1 As String
    Dim Path2 As String
    Application.DisplayAlerts = False
    
    If Len(Dir("C:\Users\" & getusername & "\OneDrive\Electrical - Shared\ELEKTRIK\ACTS", vbDirectory)) = 0 Then
       MkDir "C:\Users\" & getusername & "\OneDrive\Electrical - Shared\ELEKTRIK\ACTS"
    End If
    
    If Len(Dir("C:\Users\" & getusername & "\OneDrive\Electrical - Shared\ELEKTRIK\ACTS\" & kks, vbDirectory)) = 0 Then
       MkDir "C:\Users\" & getusername & "\OneDrive\Electrical - Shared\ELEKTRIK\ACTS\" & kks
    End If

    If Len(Dir("C:\Users\" & getusername & "\OneDrive\Electrical - Shared\ELEKTRIK\ACTS\" & kks & "\" & project, vbDirectory)) = 0 Then
       MkDir "C:\Users\" & getusername & "\OneDrive\Electrical - Shared\ELEKTRIK\ACTS\" & kks & "\" & project
    End If

    Path1 = "C:\Users\"
    
    Path2 = "\OneDrive\Electrical - Shared\ELEKTRIK\ACTS\" & kks & "\" & "\" & project & "\"
 
    
    ActiveDocument.SaveAs FileName:=Path1 & getusername & Path2 & act_no & ".docx", _
        FileFormat:=wdFormatDocumentDefault
    Application.DisplayAlerts = True

    '---------------END Save As-------------------'

    
End Sub









