Public Sub enable_UI()

'Show all normal functionality
ActiveWindow.DisplayGridlines = True
ActiveWindow.DisplayHeadings = True
ActiveWindow.DisplayWorkbookTabs = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Application.DisplayStatusBar = True
Application.DisplayFormulaBar = True
    
End Sub

Public Sub disable_UI()

'Do Not Show all normal functionality
ActiveWindow.DisplayGridlines = False
ActiveWindow.DisplayHeadings = False
ActiveWindow.DisplayWorkbookTabs = False
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
Application.DisplayStatusBar = False
Application.DisplayFormulaBar = False
    
End Sub