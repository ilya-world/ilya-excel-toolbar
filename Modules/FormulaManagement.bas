Attribute VB_Name = "FormulaManagement"
Sub AutoFillDown()
    
    Dim cellsRange As Range
    Dim masterColumn As Long, LastRow As Long
    Set cellsRange = Selection
    masterColumn = 0
    
    If cellsRange.Rows.Count > 1 Then Exit Sub
    If cellsRange.Cells(1, 1).Column = 1 Then Exit Sub
    
    For i = cellsRange.Cells(1, 1).Column To 2 Step -1
        If Not IsEmpty(ActiveSheet.Cells(cellsRange.Cells(1, 1).Row, i - 1)) Then
            masterColumn = i - 1
            LastRow = cellsRange.Cells(1, 1).Row
            For J = LastRow To Rows.Count
                If IsEmpty(ActiveSheet.Cells(J, masterColumn)) Then
                    LastRow = J - 1
                    If LastRow = cellsRange.Cells(1, 1).Row Then Exit Sub
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    If masterColumn = 0 Then Exit Sub
    
    For Each cellSelected In cellsRange
        cellSelected.Copy
        ActiveSheet.Range(Cells(cellSelected.Row + 1, cellSelected.Column), Cells(LastRow, cellSelected.Column)).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Next cellSelected
    
    Application.CutCopyMode = False
    
    ActiveSheet.Range(Cells(cellsRange.Cells(1, 1).Row, cellsRange.Cells(1, 1).Column), Cells(LastRow, cellsRange.Cells(1, 1).Column + cellsRange.Count - 1)).Select

End Sub


Sub AutoFillRight()
    
    Dim cellsRange As Range
    Dim masterRow As Long, LastColumn As Long
    Set cellsRange = Selection
    masterRow = 0
    
    If cellsRange.Columns.Count > 1 Then Exit Sub
    If cellsRange.Cells(1, 1).Row = 1 Then Exit Sub
    
    For i = cellsRange.Cells(1, 1).Row To 2 Step -1
        If Not IsEmpty(ActiveSheet.Cells(i - 1, cellsRange.Cells(1, 1).Column)) Then
            masterRow = i - 1
            LastColumn = cellsRange.Cells(1, 1).Column
            For J = LastColumn To Columns.Count
                If IsEmpty(ActiveSheet.Cells(masterRow, J)) Then
                    LastColumn = J - 1
                    If LastColumn = cellsRange.Cells(1, 1).Column Then Exit Sub
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    If masterRow = 0 Then Exit Sub
    
    For Each cellSelected In cellsRange
        cellSelected.Copy
        ActiveSheet.Range(Cells(cellSelected.Row, cellSelected.Column + 1), Cells(cellSelected.Row, LastColumn)).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Next cellSelected
    
    Application.CutCopyMode = False
    
    ActiveSheet.Range(Cells(cellsRange.Cells(1, 1).Row, cellsRange.Cells(1, 1).Column), Cells(cellsRange.Cells(1, 1).Row + cellsRange.Count - 1, LastColumn)).Select

End Sub

Sub AutoConcatenate()
    
    Dim formulaText As String, i As Integer
    
    If Not IsEmpty(ActiveSheet.Cells(Selection.Rows(1).Row, Selection.Columns(1).Column + Selection.Columns.Count)) Then
        answer = MsgBox("There are some data in the cell next to selection. Do you want to overwrite it?", vbYesNo, "Overwrite")
        If answer = vbNo Then Exit Sub
    End If
    formulaText = "=CONCATENATE("
    i = 0
    For Each cellSelected In Selection
        i = i + 1
        formulaText = formulaText & cellSelected.Address
        If i <> Selection.Cells.Count Then formulaText = formulaText & ", "
    Next cellSelected
    formulaText = formulaText & ")"
    ActiveSheet.Cells(Selection.Rows(1).Row, Selection.Columns(1).Column + Selection.Columns.Count).Value = formulaText
    
End Sub

Sub CopyFormula()
    
    Dim sourceRng As Range, destinationRange As Range
    On Error GoTo LabelSkipCopy
    Set sourceRng = Application.InputBox("Select a range to copy", "Select Range", Type:=8)
    Set destinationRange = Application.InputBox("Select a range to start pasting", "Select Range", Type:=8)
    On Error GoTo 0
    destinationRange.Formula = sourceRng.Formula
LabelSkipCopy:
    
End Sub
