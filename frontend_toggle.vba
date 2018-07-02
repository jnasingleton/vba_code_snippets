Public Sub enable_frontend()
	With Application
	    .DisplayAlerts = True
	    .ScreenUpdating = True
	    .EnableEvents = True
	End With
End Sub

Public Sub disable_frontend()
	With Application
	    .DisplayAlerts = False
	    .ScreenUpdating = False
	    .EnableEvents = False
	End With
End Sub

