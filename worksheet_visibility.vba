Public Sub ChangeWBSheetsVisibility(Optional ByVal wb As Workbook = Nothing, Optional ByVal visibliity_type As Integer = xlSheetVisible)
                                    
    If wb Is Nothing Then
        Set wb = Application.ActiveWorkbook
    End If
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Call ChangeSheetVisibility(ws, visibliity_type)
    Next ws
    
End Sub

Public Sub ChangeSheetVisibility(Optional ByVal ws As Worksheet = Nothing, Optional ByVal visibliity_type As Integer = xlSheetVisible)

    If ws Is Nothing Then
        Set ws = Application.ActiveSheet
    End If
    
    ws.Visible = visibliity_type
    
End Sub