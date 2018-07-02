Public Function Column_Letter(iColumn As Integer) As String
    'Source: https://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
    'Credit: https://stackoverflow.com/users/641067/brettdj
    Dim vArray
    vArray = Split(Cells(1, iColumn).Address(True, False), "$")
    Column_Letter = vArray(0)
End Function


Public Function SheetExists(sheetToFind As String) As Boolean
    'Source: https://stackoverflow.com/questions/6040164/excel-vba-if-worksheetwsname-exists
    'Credit: https://stackoverflow.com/users/571433/dante-is-not-a-geek
    SheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            SheetExists = True
            Exit Function
        End If
    Next Sheet
End Function