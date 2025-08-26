Attribute VB_Name = "TestFunctions"
Sub TestFunction()

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    wb.Sheets(Array("Test 2", "Test 3")).Copy
    'Call RefreshRibbon(Tag:="*")

    'If ActiveWindow.SplitRow = 0 And ActiveWindow.SplitColumn = 0 Then
    '    MsgBox "No freeze Pane Found"
    'Else
    '    MsgBox Cells(ActiveWindow.SplitRow + 1, ActiveWindow.SplitColumn + 1).Address
    'End If
    
    
    'FormCFManager.Show
    'FormTimeTest.Show
    'CreateValueFixedWorkbookCopy
    
    'MsgBox TypeName(Selection)
    'UpdateCheck True
    'Dim rand As Integer
    'rand = WorksheetFunction.RandBetween(1, 1000)
    'text = DownloadTextFile("http://legalov.ru/accenture/version.txt?a=" & rand)
    'If text <> "" Then
    '    version = CDec(text)
    '    MsgBox text
    'Else
    '    MsgBox "Can't connect to the server to check version", vbOKOnly, "Connection error"
    'End If
    
    'If visibleTag Then
    '    visibleTag = False
    '    imageTag = "HappyFace"
    'Else
    '    visibleTag = True
    '    imageTag = "FormulaMoreFunctionsMenu"
    'End If

End Sub







