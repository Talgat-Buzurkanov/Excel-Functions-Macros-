Attribute VB_Name = "Delete_Empty_Rows"
Option Explicit

Public Sub DeleteEmptyRows()
    Dim oTable As Table, oCol As Integer, oRows As Integer
    Dim iMergeCount() As Integer, dCellData() As Double
    Dim MyCell As Cell
    Dim iCurrentRow As Integer, iRowCounter As Integer

    'Watching this happen will slow things down considerably
    Application.ScreenUpdating = False

    ' Specify which table you want to work on.
    For Each oTable In ActiveDocument.Tables
        'We need to store the number of columns to determine if there are any merges
        oCol = oTable.Columns.Count

        ReDim dCellData(1 To oTable.Rows.Count, 1 To 3)
        'The first column will count the number of columns in the row if this doesn't match the table columns then we have merged cells
        'The second column will count the vertical spans which tells us if a vertically merged cell begins in this row
        'The third column will count the characters of all the text entries in the row.  If it equals zero it's empty.

        iCurrentRow = 0: iRowCounter = 0
        For Each MyCell In oTable.Range.Cells
            'The Information property only works if you select the cell. Bummer.
            MyCell.Select

            'Increment the counter if necessary and set the current row
            If MyCell.RowIndex <> iCurrentRow Then
                iRowCounter = iRowCounter + 1
                iCurrentRow = MyCell.RowIndex
            End If

            'Check column index count
            If MyCell.ColumnIndex > VBA.Val(dCellData(iRowCounter, 1)) Then dCellData(iRowCounter, 1) = MyCell.ColumnIndex

            'Check the start of vertically merged cells here
            dCellData(iRowCounter, 2) = dCellData(iRowCounter, 2) + (Selection.Information(wdEndOfRangeRowNumber) - Selection.Information(wdStartOfRangeRowNumber)) + 1

            'Add up the length of any text in the cell
            dCellData(iRowCounter, 3) = dCellData(iRowCounter, 3) + VBA.Len(Selection.Text) - 3 '(subtract one for the table and one for cursor(?))

            'Just put this in so you can see in the immediate window how Word handles all these variables
            Debug.Print "Row: " & MyCell.RowIndex & ", Column: " & MyCell.ColumnIndex & ", Rowspan = " & _
                (Selection.Information(wdEndOfRangeRowNumber) - _
                Selection.Information(wdStartOfRangeRowNumber)) + 1
        Next MyCell

        'Now we have all the information we need about the table and can start deleting some rows
        For oRows = oTable.Rows.Count To 1 Step -1
            'Check if there is no text, no merges at all and no start of a vertical merge
            If dCellData(oRows, 3) < 0 And dCellData(oRows, 1) = oCol And dCellData(oRows, 2) = oCol Then
                'Delete the row (we know it's totally unmerged so we can select the first column without issue
                oTable.Cell(oRows, 1).Select
                Selection.Rows.Delete
            End If
        Next oRows
    Next oTable

    Application.ScreenUpdating = True
End Sub
