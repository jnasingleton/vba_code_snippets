Public Sub UnprotectAll(Optional ByVal wb as Workbook = Nothing, Optional ByVal pw as String = "DEFAULT_PASSWORD")
    
    If wb Is Nothing Then
        Set wb = Application.ActiveWorkbook
    End If

    For Each ws In wb.Worksheets
        ws.Unprotect Password:=pw
    Next
End Sub

Public Sub ProtectAll(Optional ByVal wb as Workbook = Nothing, Optional ByVal pw as String = "DEFAULT_PASSWORD")
    'You can add in additional parameters for the .Protect function call
    'UserInterfaceOnly = True has been included by default as it protects the UI but not macros

    If wb Is Nothing Then
        Set wb = Application.ActiveWorkbook
    End If

    For Each ws In wb.Worksheets
        ws.Protect Password:=pw, UserInterfaceOnly:=True
    Next

End Sub