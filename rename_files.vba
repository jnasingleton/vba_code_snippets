Public Sub doRenameFile(ByVal sFolder As String, ByVal sFile As String, ByVal sFileExtOld As String, ByVal sFileExtNew as String, Optional ByVal bDeleteOriginal As Boolean = False)

Dim FileExt As String
sFileExt = Right(sFile, Len(sFile) - InStrRev(sFile, "."))

Dim FileBase As String
sFileBase = Replace(sFile, "." & sFileExt, "")

If sFileExt = sFileExtOld Then
     
    Dim sFileFull As String
    sFileFull = sFolder & sFile
    
    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(sFileFull)

    With wb
        .SaveAs sFolder & sFileBase & "." & sFileExtNew
        .Close
    End With
    
    If bDeleteOriginal Then
        Kill sFileFull
    End If
    
End If

End Sub


Public Sub doRenameFiles(Optional ByVal bDeleteOriginal As Boolean = False)

Dim sFolder As String
Dim sFile As String
sFolder = ThisWorkbook.Path & "\"
sFile = Dir(sFolder & "*.*")

Dim sFileExtOld As String
Dim sFileExtNew As String
sFileExtOld = "xlsb"
sFileExtNew = "xls"

'Dim bDeleteOriginal As Boolean
'bDeleteOriginal = True

'Exclude this current Excel workbook
Do While sFile <> "" And sFile <> ThisWorkbook.Name
    Call doRenameFile(sFolder, sFile, sFileExtOld, sFileExtNew, bDeleteOriginal)
    sFile = Dir
Loop

End Sub