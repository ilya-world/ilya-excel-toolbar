Attribute VB_Name = "FormattingManagement"
Sub ApplyCustomFormat(rowFormat As Integer)

    If TypeName(Selection) = "Range" Then
        Dim rangeSelected As Range
        Set rangeSelected = Selection
        rangeSelected.NumberFormat = ThisWorkbook.Worksheets("Formats").Cells(rowFormat + 1, 5).Value
    End If
    
End Sub

Sub CustomFormatTable()
        
    Dim selCells As Range, tableRows As Long, tableCols As Long
    Set selCells = Selection
    If selCells.Count = 1 Then
        tableRows = selCells.Row
        tableCols = selCells.Column
        For i = selCells.Row To selCells.Row + 10000
            'MsgBox ActiveSheet.Cells(i, selCells.Column).Formula
            If ActiveSheet.Cells(i, selCells.Column).Formula = "" Then Exit For
            tableRows = i
            On Error Resume Next
            
            'For k = selCells.Column To selCells.Column + 1000
            '    If ActiveSheet.Cells(i, k).Formula = "" Then Exit For
            '    If k > tableCols Then tableCols = k
            'Next k
        Next i
        tableCols = Range(Cells(selCells.Row, 1), Cells(tableRows, 1)).EntireRow.Find(What:="*", _
                            After:=Range(Cells(selCells.Row, 1), Cells(tableRows, 1)).EntireRow.Cells(1), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
        On Error GoTo 0
        Set selCells = ActiveSheet.Range(selCells, ActiveSheet.Cells(tableRows, tableCols))
        selCells.Select
        'Exit Sub
    End If
    With selCells
        .Interior.color = RGB(255, 255, 255)
        .Columns(1).Interior.color = RGB(240, 240, 240)
        .Rows(1).Interior.color = RGB(6, 107, 176)
        .Rows(1).Font.color = RGB(255, 255, 255)
        .Rows(1).Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .NumberFormat = "#,##0.00"
        If .Rows.Count > 2 Then
            For i = 2 To selCells.Rows.Count
                If .Cells(i, 1).Font.Bold Then
                    .Rows(i).Interior.color = RGB(220, 220, 220)
                    .Rows(i).Font.Bold = True
                    '.Rows(i).SpecialCells(xlCellTypeBlanks).Formula = "=SUBTOTAL(
                End If
            Next i
        End If
    End With


End Sub

