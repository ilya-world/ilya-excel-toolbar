Attribute VB_Name = "SmallFunctions"
Option Private Module

Sub ExtractLinks()

    If Selection.Columns.Count > 1 Then
        MsgBox "Please select only one column", vbOKOnly, "Error"
        Exit Sub
    End If
    
    If WorksheetFunction.CountA(Selection.Offset(0, 1)) > 0 Then
        answer = MsgBox("There are some data in the next column. Do you want to overwrite it?", vbYesNo, "Overwrite")
        If answer = vbNo Then Exit Sub
    End If
    For i = 1 To Selection.Rows.Count
        If Selection.Rows(i).Hyperlinks.Count > 0 Then
            ActiveSheet.Cells(Selection.Rows(i).Row, Selection.Rows(i).Column + 1) = Selection.Rows(i).Hyperlinks(1).Address
        End If
    Next

End Sub

Sub ProtectSheets()

    Dim currentCell As Range, MergedRange As Range, inputColor As Integer, protectPassword As String
    
    inputColor = 36
    
    If ActiveWorkbook.ActiveSheet.ProtectContents = False Then
        If MsgBox("Do you want to lock this workbook?", vbYesNo, "Lock workbook") = vbYes Then
            protectPassword = InputBox("Please enter a password (optional)", "Protection password")
            For i = 1 To ActiveWorkbook.Worksheets.Count
                Worksheets(i).Unprotect (protectPassword)
                Worksheets(i).Cells.Locked = True
                For Each currentCell In Range(Worksheets(i).Cells(1, 1), Worksheets(i).Cells.SpecialCells(xlCellTypeLastCell)).Cells
                    If currentCell.Interior.ColorIndex = inputColor Then
                        If currentCell.MergeCells = False Then
                            currentCell.Locked = False
                        Else
                            Set MergedRange = currentCell.MergeArea
                            MergedRange.Locked = False
                        End If
                    End If
                Next currentCell
                Worksheets(i).Protect (protectPassword)
            Next
        End If
    Else
        protectPassword = InputBox("Please enter a password if it extists", "Unprotect workbook")
        For i = 1 To ActiveWorkbook.Worksheets.Count
            Worksheets(i).Unprotect (protectPassword)
        Next
    End If

End Sub

Sub UnlockDocument()
    'Removes workbook protection
    'The code source is http://www.mrexcel.com/archive2/32200/36793.htm
    Dim AllClear As String, Mess As String
    Dim PWord1 As String
    Dim ShTag As Boolean, WinTag As Boolean
    Dim w1 As Worksheet, w2 As Worksheet
    Dim i As Integer, J As Integer, k As Integer, l As Integer
    Dim m As Integer, n As Integer, i1 As Integer, i2 As Integer
    Dim i3 As Integer, i4 As Integer, i5 As Integer, i6 As Integer
    
    Application.ScreenUpdating = False
    
    'Check to see if protected first...
    With ActiveWorkbook
        WinTag = .ProtectStructure Or .ProtectWindows
    End With
    
    ShTag = False
    
    For Each w1 In Worksheets
            ShTag = ShTag Or w1.ProtectContents
    Next w1
    
    If Not ShTag And Not WinTag Then
        Application.StatusBar = "No worksheet or workbook protection was found in the document."
        Application.StatusBar = False
        Exit Sub
    End If
    
    Application.StatusBar = "Document is being unlocked...please wait"
    
    If Not WinTag Then
        'There were no passwords on workbook, only sheets
    Else
      On Error Resume Next
      Do      'dummy do loop
        For i = 65 To 66: For J = 65 To 66: For k = 65 To 66
        For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
        For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
        For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
        With ActiveWorkbook
          .Unprotect Chr(i) & Chr(J) & Chr(k) & _
             Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
             Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
          If .ProtectStructure = False And _
          .ProtectWindows = False Then
              PWord1 = Chr(i) & Chr(J) & Chr(k) & Chr(l) & _
                Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
              Application.StatusBar = "Workbook successfully unlocked..."
              Exit Do  'Bypass all for...nexts
          End If
        End With
        Next: Next: Next: Next: Next: Next
        Next: Next: Next: Next: Next: Next
      Loop Until True
      On Error GoTo 0
    End If
    
    If WinTag And Not ShTag Then
      'Only workbook protected, so exit now
      Application.StatusBar = False
      Exit Sub
    End If
    
    Application.StatusBar = "Worksheet(s) are now being unlocked...please wait"
    
    On Error Resume Next
    For Each w1 In Worksheets
      'Attempt clearance with PWord1
      w1.Unprotect PWord1
    Next w1
    On Error GoTo 0
    ShTag = False
    For Each w1 In Worksheets
      'Checks for all clear ShTag triggered to 1 if not.
      ShTag = ShTag Or w1.ProtectContents
    Next w1
    If Not ShTag Then
      Application.StatusBar = "Worksheet successfully unlocked..."
      Application.StatusBar = False
      Exit Sub
    End If
    For Each w1 In Worksheets
      With w1
        If .ProtectContents Then
          On Error Resume Next
          Do      'Dummy do loop
            For i = 65 To 66: For J = 65 To 66: For k = 65 To 66
            For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
            For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
            For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
            .Unprotect Chr(i) & Chr(J) & Chr(k) & _
              Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
              Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
            If Not .ProtectContents Then
              PWord1 = Chr(i) & Chr(J) & Chr(k) & Chr(l) & _
                Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
              Application.StatusBar = "Worksheet successfully unlocked..."
              'leverage finding Pword by trying on other sheets
              For Each w2 In Worksheets
                w2.Unprotect PWord1
              Next w2
              Exit Do  'Bypass all for...nexts
            End If
            Next: Next: Next: Next: Next: Next
            Next: Next: Next: Next: Next: Next
          Loop Until True
          On Error GoTo 0
        End If
      End With
    Next w1
    Application.StatusBar = False
    Application.ScreenUpdating = True

End Sub

Sub BorderHorizontal()
    With Selection
        .Borders.LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub

Sub BorderVertical()
    With Selection
        .Borders.LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
End Sub

Sub CopyToSheets()
    FormCopyToSheets.Show
End Sub

Sub SwapCells()

    If TypeName(Selection) = "Range" Then
        Dim rangeSelected As Range, rangeBuf As Variant
        Set rangeSelected = Selection
        If rangeSelected.Areas.Count < 2 Then
            MsgBox "You must select two equal size ranges. Use Ctrl + Mouse Click to select more than one range at a time.", vbOKOnly + vbCritical
            Exit Sub
        End If
        If rangeSelected.Areas.Count > 2 Then
            MsgBox "You must select exactly TWO equal size ranges.", vbOKOnly + vbCritical
            Exit Sub
        End If
        If (rangeSelected.Areas(1).Rows.Count <> rangeSelected.Areas(2).Rows.Count) Or (rangeSelected.Areas(1).Columns.Count <> rangeSelected.Areas(2).Columns.Count) Then
            MsgBox "Both ranges should have equal amount of rows and columns", vbOKOnly + vbCritical
            Exit Sub
        End If
        Dim c As Range, c2 As Range, cRelRow As Long, cRelCol As Long, isCArray As Boolean
        For Each c In rangeSelected.Areas(1).Cells
            cRelRow = c.Row - rangeSelected.Areas(1).Row
            cRelCol = c.Column - rangeSelected.Areas(1).Column
            Set c2 = Cells(rangeSelected.Areas(2).Row + cRelRow, rangeSelected.Areas(2).Column + cRelCol)
            If c.HasArray Then
                rangeBuf = c.FormulaArray
                isCArray = True
            Else
                rangeBuf = c.Formula
                isCArray = False
            End If
            If c2.HasArray Then
                c.FormulaArray = c2.FormulaArray
            Else
                c.Formula = c2.Formula
            End If
            If isCArray Then
                c2.FormulaArray = rangeBuf
            Else
                c2.Formula = rangeBuf
            End If
        Next c
    End If

End Sub

Function GetUnicodeForSymbol(symbolName As String) As Long
    Select Case symbolName
        'Harvey balls
        Case "btnInsertSymbolHB0": GetUnicodeForSymbol = 9675
        Case "btnInsertSymbolHB1": GetUnicodeForSymbol = 9684
        Case "btnInsertSymbolHB2": GetUnicodeForSymbol = 9681
        Case "btnInsertSymbolHB3": GetUnicodeForSymbol = 9685
        Case "btnInsertSymbolHB4": GetUnicodeForSymbol = 9679
        'Currencies
        Case "btnInsertSymbolEuro": GetUnicodeForSymbol = 8364
        Case "btnInsertSymbolPound": GetUnicodeForSymbol = 163
        Case "btnInsertSymbolRuble": GetUnicodeForSymbol = 8381
        Case "btnInsertSymbolRupee": GetUnicodeForSymbol = 8377
        Case "btnInsertSymbolYen": GetUnicodeForSymbol = 165
        'Arrows
        Case "btnInsertSymbolArrowUp": GetUnicodeForSymbol = 8593
        Case "btnInsertSymbolArrowDown": GetUnicodeForSymbol = 8595
        Case "btnInsertSymbolArrowLeft": GetUnicodeForSymbol = 8592
        Case "btnInsertSymbolArrowRight": GetUnicodeForSymbol = 8594
        Case "btnInsertSymbolArrowIncrease": GetUnicodeForSymbol = 9650
        Case "btnInsertSymbolArrowDecrease": GetUnicodeForSymbol = 9660
        'Ticks
        Case "btnInsertSymbolTick": GetUnicodeForSymbol = 10003
        Case "btnInsertSymbolCross": GetUnicodeForSymbol = 10060
        'Mood
        Case "btnInsertSymbolHappy": GetUnicodeForSymbol = 9786
        Case "btnInsertSymbolNeutral": GetUnicodeForSymbol = 9787
        Case "btnInsertSymbolSad": GetUnicodeForSymbol = 9785
        'Math
        Case "btnInsertSymbolPlusMinus": GetUnicodeForSymbol = 177
        Case "btnInsertSymbolDivision": GetUnicodeForSymbol = 247
        Case "btnInsertSymbolMultiplication": GetUnicodeForSymbol = 215
        'User defined
        Case "btnInsertSymbolUser1": GetUnicodeForSymbol = ThisWorkbook.Sheets("Symbols").Cells(2, 3).Value
        Case "btnInsertSymbolUser2": GetUnicodeForSymbol = ThisWorkbook.Sheets("Symbols").Cells(3, 3).Value
        Case "btnInsertSymbolUser3": GetUnicodeForSymbol = ThisWorkbook.Sheets("Symbols").Cells(4, 3).Value
        Case "btnInsertSymbolUser4": GetUnicodeForSymbol = ThisWorkbook.Sheets("Symbols").Cells(5, 3).Value
        
        Case Else: GetUnicodeForSymbol = 0
    End Select
End Function

Sub InsertSymbol(controlName As String)
    
    Dim selCell As Range
    If Selection.Cells.Count >= 1 Then
        For Each selCell In Selection.Cells
            If selCell.HasArray Then
                selCell.FormulaArray = selCell.FormulaArray & "&""" & WorksheetFunction.Unichar(GetUnicodeForSymbol(controlName)) & """"
            ElseIf selCell.HasFormula Then
                selCell.Formula = selCell.Formula & "&""" & WorksheetFunction.Unichar(GetUnicodeForSymbol(controlName)) & """"
            Else
                selCell.Value = selCell.Value & WorksheetFunction.Unichar(GetUnicodeForSymbol(controlName))
            End If
        Next selCell
    Else
        MsgBox "No cells are selected", vbCritical
    End If
    
End Sub

Sub CaseSentence()
    
    Dim sents As Long, oCell As Range, sentences() As String, selCells As Range
'    Set selCells = Selection.Cells
    For Each oCell In Intersect(Selection, Selection.SpecialCells(xlCellTypeConstants, 2))
        sentences = Split(oCell, ".")
        For sents = 0 To UBound(sentences)
          sentences(sents) = " " & UCase(Left(LTrim(sentences(sents)), 1)) & _
                         LCase(Mid(LTrim(sentences(sents)), 2))
        Next
        oCell.Value = Trim(Join(sentences, "."))
    Next

End Sub

Sub CaseLower()
    
    For Each oCell In Intersect(Selection, Selection.SpecialCells(xlCellTypeConstants, 2))
        oCell.Value = LCase(oCell.Value)
        
    Next

End Sub

Sub CaseUpper()

    For Each oCell In Intersect(Selection, Selection.SpecialCells(xlCellTypeConstants, 2))
        oCell.Value = UCase(oCell.Value)
    Next

End Sub

Sub CaseCapitalize()

    For Each oCell In Intersect(Selection, Selection.SpecialCells(xlCellTypeConstants, 2))
        oCell.Value = StrConv(oCell.Value, vbProperCase)
    Next

End Sub

Sub CaseToogle()

    For Each oCell.Value In Intersect(Selection, Selection.SpecialCells(xlCellTypeConstants, 2))
        
    Next

End Sub
