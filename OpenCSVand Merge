Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' copyListed()
' selects all the rows which have a value "Listed" in the "E" column
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub copyListed()
    Dim lastRow As Long, i As Long
    Dim cell As Range, selectRange As Range

    ' Change to Device Summary sheet
    With Sheets("Device Summary")
        lastRow = .Range("E" & .Rows.Count).End(xlUp).Row ' Select from F2 to last device listed
                
        ' go over each cell (filtered with specialcells)
        For Each cell In Range("E2", "E" & lastRow).SpecialCells(xlCellTypeConstants)
            If (InStr(cell.Value, "Listed") > 0) Then ' Check if the cell contains the string "Listed"
                If selectRange Is Nothing Then
                    Set selectRange = cell
                Else
                    cell = cell.Resize(, 5) ' resize the selection for the table size
                    Set selectRange = Union(cell, selectRange) ' union the new cell with the current selection
                End If
            End If
        Next cell
        
        selectRange.Select
        Selection.Copy ' Now that they are all selected, copy the text
    End With
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' difference()
' This subroutine will ask the user to open a CSV file, it will then
' perform an SQL MINUS between the CSV and the Hardware listed in the
' Original worksheet. This results in a list of computers which have not
' been recorded in the Original worksheet. The result is stored in the
' Comparison Worksheet
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub difference()
    Dim sqlString As String
    Dim thisFileLocation As String
    Dim connectionString As String
    Dim recordSet As ADODB.recordSet
    Dim connection As ADODB.connection
    Dim ws As Worksheet
    
    Set ws = importCSV() ' import the CSV to the Temp worksheet
    On Error GoTo ExitSub
    
    ' The SQL query string
    ' The Query is basically an SQL MINUS because ADO does not implement it
    sqlString = "SELECT DISTINCT * FROM [Temp$] " & _
                "LEFT OUTER JOIN [Original$] " & _
                "ON [Temp$].Model = [Original$].Model " & _
                "WHERE [Original$].Model IS NULL;"
    
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
    Set ws = getWorksheet("Temp2")
    Call setHeaders(ws)
    ws.Cells(2, 1).CopyFromRecordset recordSet
        
    connection.Close ' Close the ADO connection
    
    ' Remove the Windows XP versions from the Table
    ' This also puts the information into the Comparison Worksheet
    Call removeXP
    
    ' Delete the Temporary worksheet without displaying a warning message
    Application.DisplayAlerts = False
        Sheets("Temp").Select
        ActiveWindow.SelectedSheets.Delete
        Sheets("Temp2").Select
        ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
    
    Sheets("Comparison").Activate ' Activate the Comparison worksheet
ExitSub:
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' removeXP()
' Helper Function removes all the Windows XP computers from the Temp2 list
' and places the result into a worksheet named Comparison
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function removeXP()
    Dim sqlString As String
    Dim thisFileLocation As String
    Dim connectionString As String
    Dim recordSet As ADODB.recordSet
    Dim connection As ADODB.connection
    Dim ws As Worksheet
    
    On Error GoTo ExitSub ' exit if there is an error
    
    ' The SQL Query just selects the cases where the OS is windows 7
    sqlString = "SELECT DISTINCT * FROM [Temp2$] " & _
                "WHERE ([Temp2$].OS = 'Microsoft Windows 7 Enterprise' " & _
                "OR [Temp2$].OS = 'Microsoft Windows 7 Entreprise' " & _
                "OR [Temp2$].OS = 'Microsoft Windowsá7 Entreprise');"
    
    ' Set up the ADO Connection
    thisFileLocation = ActiveWorkbook.FullName
    Set connection = New ADODB.connection
    With connection
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Extended Properties").Value = "Excel 8.0"
        .Open thisFileLocation
    End With

    ' set up the Recordset and perform the query
    Set recordSet = New ADODB.recordSet
    recordSet.Open sqlString, connection, 3, 3
    
    ' Paste the recordset into the comparison worksheet
    Set ws = getWorksheet("Comparison")
    ws.Cells.Clear
    Call setHeaders(ws)
    ws.Cells(2, 1).CopyFromRecordset recordSet
    
    
    Sheets("Comparison").Activate
    connection.Close
ExitSub:
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' setHeaders(ws: worksheet to set headers on)
' Helper Function which sets the headers of
' the worksheet to the predefined values
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function setHeaders(ws As Worksheet)
    ws.Cells(1, 1).Value = "OS"
    ws.Cells(1, 2).Value = "Manufacturer"
    ws.Cells(1, 3).Value = "Model"
    ws.Cells(1, 4).Value = "Site"
    ws.Cells(1, 5).Value = "64Bit"
    ws.Cells(1, 6).Value = "Number"
    ws.Cells(1, 7).Value = "NetBios"
    ws.Cells(1, 8).Value = "Contact"
    ws.Cells(1, 9).Value = "Status"
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
Function getWorksheet(name As String) As Worksheet
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
