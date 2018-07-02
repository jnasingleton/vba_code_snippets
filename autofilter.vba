Public Sub TurnAutoFilterOn(Optional ByVal autofilter_ws As Worksheet = Nothing, _
					 Optional ByVal autofilter_rangestring As String = "A1", _
					 Optional ByVal replace_autofilter As Boolean = False)

	If autofilter_ws Is Nothing Then
        Set autofilter_ws = Application.ActiveSheet
    End If

	If not autofilter_ws.AutoFilterMode
		autofilter_ws.Range(autofilter_rangestring).AutoFilter
	ElseIf autofilter_ws.AutoFilterMode and replace_autofilter Then
		'Autofilter enabled, and we are replacing the autofilter
		autofilter_ws.AutoFilterMode = False
		autofilter_ws.Range(autofilter_rangestring).AutoFilter
	End If
	
End Sub

Public Sub TurnAutoFilterOff(Optional ByVal autofilter_ws As Worksheet = Nothing)

	If autofilter_ws Is Nothing Then
        Set autofilter_ws = Application.ActiveSheet
    End If

	autofilter_ws.AutoFilterMode = False
	
End Sub