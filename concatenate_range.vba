Public Function ConcatenateRange(ByVal cell_range As Range, Optional ByVal seperator As String = ",", Optional ByVal include_blanks = False) As String
	
	Dim cellvalues_array As Variant
	cellvalues_array = cell_range.Value

	Dim newString As String
	newString = ""

	Dim i As Long, j As Long
	For i = 1 To UBound(cellvalues_array, 1)
	    For j = 1 To UBound(cellvalues_array, 2)
	        
	        cellvalue = cellvalues_array(i, j)
	        string_length = Len(cellvalue)

	        If (include_blanks and string_length == 0) Or (string_length > 0) Then
	            If newString = "" Then
	            	newString = cellvalue
	            Else 
	            	newString = newString & (seperator & cellvalue)
	            End If
	        End If

	    Next
	Next

	ConcatenateRange = newString

End Function