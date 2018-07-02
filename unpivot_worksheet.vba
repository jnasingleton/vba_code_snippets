Public Sub unpivot_worksheet(ByVal iOffsetColumnsCount As Integer, ByVal iIDColumnsCount As Integer, ByVal iValueColumnsCount As Integer, _
                             ByVal wb As Workbook, Optional ByVal ws_source_name As String = "", Optional ByVal ws_dest_name As String = "")
'This function converts a worksheet from a table format (having numerous value columns) into a database format (having a single value column)
'Assume ws_source has a header row, with data rows starting on row 2
'Assume at least one of ws_source_name or ws_dest_name is provided as a parameter (else both resulting worksheets are set to the same worksheet)

'ws_source is the worksheet to be unpivoted
Dim ws_source As Worksheet
'Set ws_source = Sheets("Sheet1")
If ws_source_name = "" Then
    Set ws_source = Application.ActiveSheet
Else
    If SheetExists(ws_source_name) Then
        Set ws_source = wb.Sheets(ws_source_name)
    Else
        MsgBox ("The source worksheet specified does not exist!")
        Exit Sub
    End If
End If

'ws_dest is the worksheet the unpivoted data should be exported to.
Dim ws_dest As Worksheet
'Set ws_dest = Sheets("Sheet1")
If ws_dest_name = "" Then
    Set ws_dest = Application.ActiveSheet
Else
    If SheetExists(ws_dest_name) Then
        Set ws_dest = wb.Sheets(ws_dest_name)
        'Clear the formats and contents
        ws_dest.Cells.ClearFormats
        ws_dest.Cells.ClearContents
    Else
        Set ws_dest = ThisWorkbook.Sheets.Add()
        ws_dest.Name = ws_dest_name
    End If
End If

'iOffset is the number of columns that the first ID column is offset, from the left.
'Dim iOffsetColumnsCount As Integer
'iOffsetColumnsCount = 1

'iIDColumnsCount are the number of columns that store ID information (ie. are not transposed)
Dim iIDColumnsIndex As Integer
'Dim iIDColumnsCount As Integer
'iIDColumnsCount = 1

'iValueColumnsCount are the number of columns that store value information (ie. are to be transposed into a single value column)
Dim iValueColumnsIndex As Integer
'Dim iValueColumnsCount As Integer
'iValueColumnsCount = 8

'sColumnIndex is the index column of ws_source (having a complete/full column)
Dim sColumnIndex As String
sColumnIndex = Column_Letter(1 + iOffsetColumnsCount)

'Set rows range for the source worksheet
'Assume a full column on column sColumnIndex
Dim iRowSourceStart As Long
Dim iRowSourceEnd As Long
Dim iRowSource As Long
iRowSourceStart = 2
iRowSourceEnd = ws_source.Cells(ws_source.Rows.Count, sColumnIndex).End(xlUp).Row

'Set rows for the dest worksheet
Dim iRowDestStart As Long
Dim iRowDest As Long
iRowDestStart = 2
iRowDest = iRowDestStart

'Determine where ID and Value columns start
Dim iIDColumnsStart As Integer
Dim iValueColumnsStart As Integer
iIDColumnsStart = iOffsetColumnsCount + 1
iValueColumnsStart = iOffsetColumnsCount + iIDColumnsCount + 1

'Build ws_dest header
ws_source.Range("A1:" & Column_Letter(iValueColumnsStart - 1) & 1).Copy _
    Destination:=ws_dest.Range("A1")
ws_dest.Range(Column_Letter(iValueColumnsStart) & "1") = "value_header"
ws_dest.Range(Column_Letter(iValueColumnsStart + 1) & "1") = "value"

For iRowSource = iRowSourceStart To iRowSourceEnd
    For iValueColumnsIndex = 1 To iValueColumnsCount

        'Copy over ID column values
        For iIDColumnsIndex = 1 To iIDColumnsCount
            ws_dest.Cells(iRowDest, iOffsetColumnsCount + iIDColumnsIndex) = _
                ws_source.Cells(iRowSource, iIDColumnsStart + iIDColumnsIndex - 1)
        Next

        iIDColumnLast = iOffsetColumnsCount + iIDColumnsCount

        'Copy over value column header value
        ws_dest.Cells(iRowDest, iIDColumnLast + 1) = _
            ws_source.Cells(1, iValueColumnsStart + iValueColumnsIndex - 1)

        'Copy over value column value
        ws_dest.Cells(iRowDest, iIDColumnLast + 2) = _
            ws_source.Cells(iRowSource, iValueColumnsStart + iValueColumnsIndex - 1)

        iRowDest = iRowDest + 1

    Next
Next

End Sub