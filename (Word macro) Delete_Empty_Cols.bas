Attribute VB_Name = "Module1"
Option Explicit

Public Sub DeleteEmptyCols()
    Dim oTable As Table, oCol As Integer, oRows As Integer
    Dim iMergeCount() As Integer, dCellData() As Double
    Dim MyCell As Cell
    Dim iCurrentRow As Integer, iRowCounter As Integer

    'Watching this happen will slow things down considerably
    Application.ScreenUpdating = False

    oRows = 4
    
    If VBA.Len(ActiveDocument.Tables(2).Cell(3, 3)) = 2 Then
    
        ActiveDocument.Tables(2).Cell(3, 3).Select
        Selection.Columns.Delete
        oRows = oRows - 1
        ActiveDocument.Tables(2).Cell(2, 1).Range.Delete        'clear contents of second row
        ActiveDocument.Tables(2).Cell(2, 2).Range.Delete
        ActiveDocument.Tables(2).Cell(2, 3).Range.Delete
        ActiveDocument.Tables(2).Cell(2, 4).Range.Delete
        ActiveDocument.Tables(2).Cell(2, 5).Range.Delete
        ActiveDocument.Tables(3).Cell(2, 1).Range.Delete
        ActiveDocument.Tables(3).Cell(2, 2).Range.Delete
        ActiveDocument.Tables(3).Cell(2, 3).Range.Delete
        ActiveDocument.Tables(3).Cell(2, 4).Range.Delete
        ActiveDocument.Tables(3).Cell(2, 5).Range.Delete
        
        
        ActiveDocument.Tables(2).Cell(2, 1).Range.InsertAfter "1"       ' insert numeration for the case if only column 3 is deleted
        ActiveDocument.Tables(2).Cell(2, 2).Range.InsertAfter "2"
        ActiveDocument.Tables(2).Cell(2, 3).Range.InsertAfter "3"
        ActiveDocument.Tables(2).Cell(2, 4).Range.InsertAfter "4"
        ActiveDocument.Tables(2).Cell(2, 5).Range.InsertAfter "5"
        ActiveDocument.Tables(3).Cell(2, 1).Range.InsertAfter "6"
        ActiveDocument.Tables(3).Cell(2, 2).Range.InsertAfter "7"
        ActiveDocument.Tables(3).Cell(2, 3).Range.InsertAfter "8"
        ActiveDocument.Tables(3).Cell(2, 4).Range.InsertAfter "9"
        ActiveDocument.Tables(3).Cell(2, 5).Range.InsertAfter "10"
        
        ActiveDocument.Tables(2).Columns(4).Width = InchesToPoints(2.5)
        
    
    End If
    
    If VBA.Len(ActiveDocument.Tables(2).Cell(3, oRows)) = 2 Then
    
        ActiveDocument.Tables(2).Cell(3, oRows).Select
        Selection.Columns.Delete
        ActiveDocument.Tables(2).Cell(2, 1).Range.Delete                        'clear contents of second row
        ActiveDocument.Tables(2).Cell(2, 2).Range.Delete
        ActiveDocument.Tables(3).Cell(2, 1).Range.Delete
        ActiveDocument.Tables(3).Cell(2, 2).Range.Delete
        ActiveDocument.Tables(3).Cell(2, 3).Range.Delete
        ActiveDocument.Tables(3).Cell(2, 4).Range.Delete
        ActiveDocument.Tables(3).Cell(2, 5).Range.Delete
        ActiveDocument.Tables(2).Cell(2, 1).Range.InsertAfter "1"               ' insert numeration based on condition
        ActiveDocument.Tables(2).Cell(2, 2).Range.InsertAfter "2"
        
        If oRows = 3 Then
            ActiveDocument.Tables(2).Cell(2, 3).Range.Delete                    'clear contents of second row
            ActiveDocument.Tables(2).Cell(2, 4).Range.Delete
            ActiveDocument.Tables(2).Cell(2, 3).Range.InsertAfter "3"           ' insert numeration based on condition
            ActiveDocument.Tables(2).Cell(2, 4).Range.InsertAfter "4"
            ActiveDocument.Tables(3).Cell(2, 1).Range.InsertAfter "5"
            ActiveDocument.Tables(3).Cell(2, 2).Range.InsertAfter "6"
            ActiveDocument.Tables(3).Cell(2, 3).Range.InsertAfter "7"
            ActiveDocument.Tables(3).Cell(2, 4).Range.InsertAfter "8"
            ActiveDocument.Tables(3).Cell(2, 5).Range.InsertAfter "9"
            
            ActiveDocument.Tables(2).Columns(4).Width = InchesToPoints(2.5)
            
            Else
            ActiveDocument.Tables(2).Cell(2, 3).Range.Delete                    'clear contents of second row
            ActiveDocument.Tables(2).Cell(2, 4).Range.Delete
            ActiveDocument.Tables(2).Cell(2, 5).Range.Delete
            ActiveDocument.Tables(2).Cell(2, 3).Range.InsertAfter "3"           ' insert numeration based on condition
            ActiveDocument.Tables(2).Cell(2, 4).Range.InsertAfter "4"
            ActiveDocument.Tables(2).Cell(2, 5).Range.InsertAfter "5"
            ActiveDocument.Tables(3).Cell(2, 1).Range.InsertAfter "6"
            ActiveDocument.Tables(3).Cell(2, 2).Range.InsertAfter "7"
            ActiveDocument.Tables(3).Cell(2, 3).Range.InsertAfter "8"
            ActiveDocument.Tables(3).Cell(2, 4).Range.InsertAfter "9"
            ActiveDocument.Tables(3).Cell(2, 5).Range.InsertAfter "10"
        
            ActiveDocument.Tables(2).Columns(5).Width = InchesToPoints(2.45)
            
        End If
        
            
    End If
    
    


    Application.ScreenUpdating = True
    
End Sub


