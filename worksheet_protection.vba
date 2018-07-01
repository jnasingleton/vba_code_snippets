Public Sub UnprotectAll(Optional ByVal wb as Workbook = ActiveWorkbook, Optional ByVal pw as String = "DEFAULT_PASSWORD")
    For Each ws In wb.Worksheets
        ws.Unprotect Password:=sPassword
    Next
End Sub

Public Sub ProtectAll(Optional ByVal wb as Workbook = ActiveWorkbook, Optional ByVal pw as String = "DEFAULT_PASSWORD")
    'You can add in additional parameters for the .Protect function call
    'UserInterfaceOnly = True has been included by default as it protects the UI but not macros
    For Each ws In wb.Worksheets
        ws.Protect Password:=sPassword, UserInterfaceOnly:=True
    Next
End Sub