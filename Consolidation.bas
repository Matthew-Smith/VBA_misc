Sub consolidate()
    Dim ws As Worksheet
    Dim lo As Excel.ListObject
    Dim allDataTable As Excel.ListObject
    Set allDataTable = ActiveSheet.ListObjects("AllData") ' Set the table for consolidation
    
    If allDataTable.DataBodyRange.Rows.Count > 1 Then
        Call clearAllData   ' Clear the table for use if it is empty
    End If
    
    For Each ws In Worksheets
        If ws.Name <> "All" And ws.Name <> "Summary" And ws.Name <> "Summary2" Then
            ' Go through each table in the worksheet (should only be the one)
            For Each lo In ws.ListObjects
                        
                ' Insert Rows to the consolidated table to fit the Data,
                ' Start at 2 because we keep an extra row when clearing the old data
                Dim X As Integer
                For X = 1 To lo.DataBodyRange.Rows.Count
                    allDataTable.DataBodyRange.Rows(1).EntireRow.Insert
                Next X
                
                lo.DataBodyRange.Copy ' Copy the Data from the current worksheet
                
                ' Paste the selected table from the other worksheet to the All Data table
                allDataTable.DataBodyRange.Offset(, 1).PasteSpecial Paste:=xlPasteAll, _
                    Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
                ' Move selection back to the Division (1st) column
                Selection.Offset(0, -1).Resize(, 1).Select
                Selection.Value = ws.Name ' Insert the worksheet name as the Division
            Next lo
        End If
    Next ws
    
    ' Delete the last row that we kept in the data clearing
    allDataTable.DataBodyRange.Rows(allDataTable.DataBodyRange.Rows.Count).Delete
    
    Application.CutCopyMode = False ' remove copy data from clipboard
    ThisWorkbook.RefreshAll ' Update the Pivot table and the slicers
    Cells(4, 1).Select ' Move selector to the first cell in the
    
End Sub

Sub clearAllData() ' This will clear the Data in the AllData table
    Dim loSource As Excel.ListObject
    Set loSource = ActiveSheet.ListObjects("AllData") ' Use the AllData Table
    
    'loSource.DataBodyRange.ClearContents
        
    ' Clear Data from the Table
    With loSource
        .Range.AutoFilter ' select full table
        ' Deselect the last Row so we don't kill the table
        .DataBodyRange.Resize(.DataBodyRange.Rows.Count - 1, _
            .DataBodyRange.Columns.Count).Rows.Delete
        .DataBodyRange.ClearContents ' clear the data in remaining row
    End With
End Sub
