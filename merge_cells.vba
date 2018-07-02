Public Sub UnmergeWorksheetCells(Optional ByVal ws = Nothing)

	If ws Is Nothing Then
		Set ws = Application.ActiveSheet
	End If

	Dim rCell As Range, rJoinedCells As Range
	For Each rCell In ws.UsedRange
		If rCell.MergeCells Then
			Set rJoinedCells = rCell.MergeArea
			rCell.MergeCells = False
			rJoinedCells.Value = rCell.Value
		End If
	Next

End Sub

Public Sub UnmergeWorkbookCells(Optional ByVal wb = Nothing)

	If wb Is Nothing Then
		Set wb = Application.ActiveWorkbook
	End If

	Dim ws As Worksheet
	For Each ws In wb.Worksheets
		Call(UnmergeWorksheetCells(ws))
	Next

End Sub

Public Sub UnmergeAllWorkbookCells()

	Dim wb As Workbook
	For Each wb In Workbooks
		Call(UnmergeWorkbookCells(wb))
	Next

End Sub