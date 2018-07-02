Public Function CountIn(ByVal sWord1 As String, ByVal sWord2 As String)
' This function returns the number of instances that sWord2 appears in sWord1

    sFindWord = sWord2
    sFindWordLength = Len(sFindWord)

    sReplaceWord = ""

    CountIn = Len(sWord1) - Len(Replace(sWord1, sFindWord, sReplaceWord))
    CountIn = CountIn / sFindWordLength
    
End Function

Public Function CountInRange(ByVal r As Range, ByVal sWord2 As String) As String 
    
    nWordRange = CountIn(r.Value, sWord2)

End Function