Attribute VB_Name = "ViewManagement"
Sub GoToCorner()
    
    Dim aSheet As Worksheet
    currentSheet = ActiveSheet.Name
    
    For Each aSheet In Worksheets
         If (aSheet.visible = True) Then
            aSheet.Activate
            If ActiveWindow.SplitRow <> 0 Or ActiveWindow.SplitColumn <> 0 Then
                aSheet.Cells(ActiveWindow.SplitRow + 1, ActiveWindow.SplitColumn + 1).Select
            End If
            Range("A1").Select
         End If
    Next
    
    Sheets(currentSheet).Select
    MsgBox "Done! Don't forget to save your file before sending"

End Sub

Sub SwitchZoom()

    If ActiveWindow.Zoom <> 100 Then
        ActiveWindow.Zoom = 100
    Else
        ActiveWindow.Zoom = 70
    End If

End Sub

Sub HideInterface()
    
    Dim wsh As Worksheet
    Dim wshOrig As Worksheet
    
    If ActiveWindow.DisplayHeadings Then
        Set wshOrig = ActiveSheet
        Application.ScreenUpdating = False
        For Each wsh In ActiveWorkbook.Worksheets
            wsh.Activate
            ActiveWindow.DisplayHeadings = False
        Next wsh
        ActiveWindow.DisplayWorkbookTabs = False
        wshOrig.Activate
        Application.ScreenUpdating = True
    Else
        Set wshOrig = ActiveSheet
        Application.ScreenUpdating = False
        For Each wsh In ActiveWorkbook.Worksheets
            wsh.Activate
            ActiveWindow.DisplayHeadings = True
        Next wsh
        ActiveWindow.DisplayWorkbookTabs = True
        wshOrig.Activate
        Application.ScreenUpdating = True
    End If

End Sub

Sub ZoomToWidth()

    Dim lastCellAddress As String, zoomRange As Range, initialSelection As Range
    Set initialSelection = Selection
    lastCellAddress = FindLastCell(ActiveSheet.Name)
    Set zoomRange = ActiveSheet.Range(Cells(1, 1), Cells(1, Range(lastCellAddress).Column))
    Application.GoTo zoomRange
    ActiveWindow.Zoom = True
    initialSelection.Select

End Sub
