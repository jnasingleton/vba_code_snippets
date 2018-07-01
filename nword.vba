Public Function nWord(ByVal sWord1 As String, ByVal sWord2 As String, Optional ByVal N As Long = 1, Optional ByVal bReverse As Boolean = False, Optional ByVal bToClean As Boolean = True) As String
' This function returns the n-th word of a provided string, split by another specified string
' Options for reverse search and cleaning are included

    nWord = ""

    Dim sWord1Temp As String
    Dim sWord2Temp As String
    If bReverse Then
        sWord1Temp = StrReverse(sWord1)
        sWord2Temp = StrReverse(sWord2)
    Else
        sWord1Temp = sWord1
        sWord2Temp = sWord2
    End If
    
    Dim sArray1() As String
    If InStr(sWord1Temp, sWord2Temp) Then
        sArray1 = Split(sWord1Temp, sWord2Temp)
        'This is guaranteed to be True because of the above InStr check
        If UBound(sArray1) >= N - 1 Then
            nWord = sArray1(N - 1)
            If bReverse Then
                nWord = StrReverse(nWord)
            End If
            If bToClean Then
                nWord = Trim$(Application.WorksheetFunction.Clean(nWord))
            End If
        End If
    Else
    	'Default to returning the entire word if no match found
    	nWord = sWord1
    End If
    
End Function

Public Function nWordRange(ByVal r As Range, ByVal sWord2 As String, Optional ByVal N As Long = 1, Optional ByVal bReverse As Boolean = False) As String
    
    nWordRange = nWord(r.Value, sWord2, N, bReverse)

End Function