Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' merge()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub merge()
    Dim sqlString As String
    Dim thisFileLocation As String
    Dim connectionString As String
    Dim recordSet As ADODB.recordSet
    Dim connection As ADODB.connection
    Dim ws As Worksheet
    
    ' Set ws = importCSV() ' import the CSV to the Temp worksheet
    ' On Error GoTo ExitSub
    
    ' The SQL query string
    sqlString = "SELECT * FROM [sccmssystems$] " & _
                "UNION ALL " & _
                "SELECT * FROM [SMS$];"
    
    thisFileLocation = ActiveWorkbook.FullName ' Get the file name to create the ADO connection
    Set connection = New ADODB.connection
    
    ' Set up the ADO Connection
    With connection
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Extended Properties").Value = "Excel 8.0"
        .Open thisFileLocation
    End With

    ' Opens a Recordset (basically a temporary SQL table) and performs the query
    Set recordSet = New ADODB.recordSet
    recordSet.Open sqlString, connection, 3, 3
    
    ' Copy the Data out of the recordset to the Temp2 worksheet
    Set ws = getWorksheet("Merged")
    Call setHeaders(ws)
    ws.Cells(2, 1).CopyFromRecordset recordSet
    
    connection.Close ' Close the ADO connection
    
    
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=IF(COUNTIF([Name],RC[-11])>1,""Duplicate"","""")"
        
    ' Delete the Temporary worksheet without displaying a warning message
    ' Application.DisplayAlerts = False
    '     Sheets("Temp").Select
    '     ActiveWindow.SelectedSheets.Delete
    '     Sheets("Temp2").Select
    '     ActiveWindow.SelectedSheets.Delete
    ' Application.DisplayAlerts = True
    
    Sheets("Merged").Activate ' Activate the Comparison worksheet
ExitSub:
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' setHeaders(ws: worksheet to set headers on)
' Helper Function which sets the headers of
' the worksheet to the predefined values
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function setHeaders(ws As Worksheet)
    ws.Cells(1, 1).Value = "Name"
    ws.Cells(1, 2).Value = "ResourceType"
    ws.Cells(1, 3).Value = "Domain"
    ws.Cells(1, 4).Value = "SiteCode"
    ws.Cells(1, 5).Value = "Client"
    ws.Cells(1, 6).Value = "Approved"
    ws.Cells(1, 7).Value = "Assigned"
    ws.Cells(1, 8).Value = "Blocked"
    ws.Cells(1, 9).Value = "ClientType"
    ws.Cells(1, 10).Value = "Obsolete"
    ws.Cells(1, 11).Value = "Active"
    ws.Cells(1, 12).Value = "Duplicate"
    
    ws.ListObjects.Add(xlSrcRange, Range("$A$1:$L$1"), , xlYes).name = "MergedTable"
    Range("MergedTable[#All]").Select
    ws.ListObjects("MergedTable").TableStyle = "TableStyleMedium2"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' importCSV(): the worksheet where the CSV has been imported
' Opens a dialog window to select a CSV then imports the CSV to a Temporary
' worksheet. Returns the reference to the worksheet where
' the CSV was imported
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function importCSV() As Worksheet
    Dim fileLocation As String
    Dim ws As Worksheet
    
    fileLocation = openDialog() ' open the dialog window to select a CSV
    If fileLocation = "" Then ' if no file was selected return from the function
        Exit Function
    End If
    
    Set ws = getWorksheet("Temp") ' get reference to the Temporary worksheet
    ws.Activate
    ws.Cells.Clear
        
    With ws.QueryTables.Add(connection:="TEXT;" & fileLocation, Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Call setHeaders(ws) ' set the headers to match what we need
    Set importCSV = ws ' return the worksheet
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' getWorksheet(name:Name of the worksheet to retrieve)
'             :The requested worksheet
' Helper function, will return a reference to the worksheet with the
' passed name. If the worksheet does not exist it will create it.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function getWorksheet(name As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In Worksheets ' loop through every worksheet
        If name = ws.name Then
            Set getWorksheet = ws ' return it if it exists
            Exit Function
        End If
    Next ws
    Set ws = Sheets.Add ' create it if it doesn't exist
    ws.name = name
    Set getWorksheet = ws
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' openDialog(): selected file location as a string
' Opens an Open dialog window and returns the
' selected file location as a string
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function openDialog() As String
    Dim fd As Office.FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

   With fd

      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add csv.
      .Filters.Clear
      .Filters.Add "Comma-separated Values", "*.csv"
      .Filters.Add "All Files", "*.*"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
        openDialog = .SelectedItems(1) ' set the return value of the function

      End If
   End With
End Function
