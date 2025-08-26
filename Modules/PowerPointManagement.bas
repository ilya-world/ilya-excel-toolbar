Attribute VB_Name = "PowerPointManagement"
Public pptSheet As Worksheet
Public oPPTApp As PowerPoint.Application
Public oPPTShape As Object
Public oPPTPres As PowerPoint.Presentation
Public pres As PowerPoint.Presentation
Public resultCol As Integer, valueCol As Integer, nameCol As Integer, prevCol As Integer, partCol As Integer, tempCol As Integer, totalPartCol As Integer, bgColorCol As Integer, fontColorCol As Integer
Public skipRows As Integer
Public totalRules As Integer
Public strPresPath As String
Public sourceCellRange As Range
Public shapeRangeName As String
Public selectedTableColumn As Integer, selectedTableRow As Integer
Public bgColorGroupRange As Range, fontColorGroupRange As Range

Sub InitializeColumnNumbers()

    nameCol = 1
    valueCol = nameCol + 1
    prevCol = valueCol + 1
    bgColorCol = prevCol + 1
    fontColorCol = bgColorCol + 1
    partCol = fontColorCol + 1
    totalPartCol = partCol + 1
    resultCol = totalPartCol + 1
    tempCol = resultCol + 1
    Set pptSheet = ActiveWorkbook.Worksheets(ThisWorkbook.Sheets("General").Range("B5").Value)

End Sub

Sub ConnectToPowerPoint()

    InitializeColumnNumbers
    
    skipRows = 4
    totalRules = WorksheetFunction.CountA(pptSheet.Columns(1)) - skipRows + 1
    strPresPath = pptSheet.Range("B1")
    Set oPPTApp = New PowerPoint.Application

    For Each pres In oPPTApp.Presentations
       If pres.FullName = strPresPath Then
          ' found it!
          Set oPPTPres = pres
          Exit For
       End If
    Next
    If oPPTPres Is Nothing Then
        Set oPPTPres = oPPTApp.Presentations.Open(strPresPath)
    End If
    
End Sub

Sub DisconnectFromPowerPoint()
    Set pres = Nothing
    Set oPPTShape = Nothing
    Set oPPTFile = Nothing
    Set oPPTPres = Nothing
    Set oPPTApp = Nothing
End Sub

Sub SelectThinkCell()

    Dim sourceRng As Range, objectName As String
    
    Set sourceRng = SelectRangeOfCells
    If Not (sourceRng Is Nothing) Then
        objectName = InputBox("Enter the name of your think-cell chart (assign name through think-cell chart text box)", "Chart name")
        If objectName = "" Then Exit Sub
        pptSheet.Cells(Selection.Row, nameCol).Value = "[ThinkCell]" & objectName
        pptSheet.Cells(Selection.Row, valueCol).Value = "[Range]|" & sourceRng.Worksheet.Name & "|" & sourceRng.Address
    End If

End Sub

Sub SelectPPTChart()

    Dim sourceRng As Range, objectName As String
    
    Set sourceRng = SelectRangeOfCells
    If Not (sourceRng Is Nothing) Then
        objectName = InputBox("Enter the name of your chart (assign name through selection pane)", "Chart name")
        If objectName = "" Then Exit Sub
        pptSheet.Cells(Selection.Row, nameCol).Value = "[Chart]" & objectName
        pptSheet.Cells(Selection.Row, valueCol).Value = "[Range]|" & sourceRng.Worksheet.Name & "|" & sourceRng.Address
    End If

End Sub

Function SelectRangeOfCells(Optional inputMessage As String) As Range
    
    If IsMissing(inputMessage) Then inputMessage = "Select a range to use in PowerPoint"
    
    Set SelectRangeOfCells = Nothing
    If ActiveSheet.Name = ThisWorkbook.Sheets("General").Range("B5").Value And Selection.Count = 1 Then
        Dim sourceRng As Range
        On Error GoTo LabelSkipRangeSelect
        Set sourceRng = Application.InputBox(inputMessage, "Select Range", Type:=8)
        On Error GoTo 0
        Set SelectRangeOfCells = sourceRng
        InitializeColumnNumbers
    End If
LabelSkipRangeSelect:


End Function

Sub SelectTableCell()

    Dim sourceRng As Range, objectName As String
    
    Set sourceRng = SelectRangeOfCells
    If Not (sourceRng Is Nothing) Then
        objectName = InputBox("Enter the name of your table (assign name through selection pane)", "Table name")
        If objectName = "" Then Exit Sub
        Set fontColorGroupRange = Nothing
        Set bgColorGroupRange = Nothing
        FormSelectTableCell.Show
        If selectedTableColumn <> 0 And selectedTableRow <> 0 Then
            pptSheet.Cells(Selection.Row, nameCol).Value = "[Table]" & objectName & "|" & selectedTableRow & "|" & selectedTableColumn
            pptSheet.Cells(Selection.Row, valueCol).Value = "[Range]|" & sourceRng.Worksheet.Name & "|" & sourceRng.Address
            If Not bgColorGroupRange Is Nothing Then
                pptSheet.Cells(Selection.Row, bgColorCol).Value = "[Range]|" & bgColorGroupRange.Worksheet.Name & "|" & bgColorGroupRange.Address
            End If
            If Not fontColorGroupRange Is Nothing Then
                pptSheet.Cells(Selection.Row, fontColorCol).Value = "[Range]|" & fontColorGroupRange.Worksheet.Name & "|" & fontColorGroupRange.Address
            End If
        End If
    End If

End Sub

Sub SelectShapeGroup()
    
    If ActiveSheet.Name = ThisWorkbook.Sheets("General").Range("B5").Value And Selection.Count = 1 Then
    
        Dim sourceRng As Range
        
        On Error GoTo LabelSkipRangeSelect
        Set sourceRng = Application.InputBox("Select a range to use in PowerPoint", "Select Range", Type:=8)
        On Error GoTo 0
        Set sourceCellRange = sourceRng
        Set fontColorGroupRange = Nothing
        Set bgColorGroupRange = Nothing
        FormSelectShapeGroup.Show
        pptSheet.Cells(Selection.Row, nameCol).Value = "[ShapeGroup]" & shapeRangeName
        pptSheet.Cells(Selection.Row, valueCol).Value = "[Range]|" & sourceCellRange.Worksheet.Name & "|" & sourceCellRange.Address
        If Not bgColorGroupRange Is Nothing Then
            pptSheet.Cells(Selection.Row, bgColorCol).Value = "[Range]|" & bgColorGroupRange.Worksheet.Name & "|" & bgColorGroupRange.Address
        End If
        If Not fontColorGroupRange Is Nothing Then
            pptSheet.Cells(Selection.Row, fontColorCol).Value = "[Range]|" & fontColorGroupRange.Worksheet.Name & "|" & fontColorGroupRange.Address
        End If
    End If
    
LabelSkipRangeSelect:

End Sub

Sub SelectPPTShapePart()
    FormSelectParts.Show
End Sub

Sub ReselectPPTFile()
    FormSelectPPT.Show
End Sub

Function GetRGBFromText(inputString As String) As Long

    Dim splitString() As String
    GetRGBFromText = -1
    On Error GoTo NoRGB
    splitString = Split(inputString, "|", 3)
    GetRGBFromText = RGB(splitString(0), splitString(1), splitString(2))
    On Error GoTo 0
NoRGB:

End Function

Sub UpdatePPTShapes()
    
    If WorksheetExists(ThisWorkbook.Sheets("General").Range("B5").Value) Then
        
        Dim currentSlide As PowerPoint.Slide
        Dim tcaddin As Object
        Dim chartSheet As Worksheet
        Dim bgColorRange As Range, fontColorRange As Range
        Dim slideNum As Integer
        Dim newValue As String
        Dim skipChange As Boolean, changeValue As Boolean
        Dim ruleName As String, ruleNameX As String, patternName As String
        Dim splitLine() As String, splitBgColor() As String, splitFontColor() As String
        Dim severalCells As Boolean
        Dim cycleCount As Integer, cycleNum As Integer
        Dim valueRange As Range
        Dim genericFormat As Boolean, ruleValue As Range
        Dim tableStartRow As Integer, tableStartCol As Integer
        Dim bgColorCell As Range, fontColorCell As Range
        Dim bgColorRGB As Long, fontColorRGB As Long
        
        ConnectToPowerPoint
        
        If totalRules > 0 Then
            For ruleNum = 1 To totalRules
                pptSheet.Cells(skipRows + ruleNum, resultCol).Value = 0
            Next ruleNum
            Dim bar As Progressbar
            Set bar = New Progressbar
            bar.createLoadingBar
            bar.createString
            bar.createtimeDuration
            bar.setParameters oPPTPres.Slides.Count
            bar.Start
            For slideNum = 1 To oPPTPres.Slides.Count
                Set currentSlide = oPPTPres.Slides(slideNum)
                bar.Update slideNum, "Slides"
                For ruleNum = 1 To totalRules
                    ruleName = pptSheet.Cells(skipRows + ruleNum, nameCol).Value
                    patternName = ""
                    If Left(ruleName, 1) = "[" Then
                        patternName = Split(ruleName, "]", 2)(0)
                        'MsgBox patternName
                        'For k = LBound(splitBracket) To UBound(splitBracket)
                        ruleNameX = Split(ruleName, "]", 2)(1)
                        If patternName = "[Table" Then
                            tableStartRow = Split(ruleNameX, "|", 3)(1)
                            tableStartCol = Split(ruleNameX, "|", 3)(2)
                            ruleNameX = Split(ruleNameX, "|", 3)(0)
                        End If
                        splitLine = Split(pptSheet.Cells(skipRows + ruleNum, valueCol).Value, "|")
                        Set valueRange = ActiveWorkbook.Sheets(splitLine(1)).Range(splitLine(2))
                        If pptSheet.Cells(skipRows + ruleNum, bgColorCol).Value <> "" Then
                            splitBgColor = Split(pptSheet.Cells(skipRows + ruleNum, bgColorCol).Value, "|")
                            Set bgColorRange = ActiveWorkbook.Sheets(splitBgColor(1)).Range(splitBgColor(2))
                        Else
                            Set bgColorRange = Nothing
                        End If
                        If pptSheet.Cells(skipRows + ruleNum, fontColorCol).Value <> "" Then
                            splitFontColor = Split(pptSheet.Cells(skipRows + ruleNum, fontColorCol).Value, "|")
                            Set fontColorRange = ActiveWorkbook.Sheets(splitFontColor(1)).Range(splitFontColor(2))
                        Else
                            Set fontColorRange = Nothing
                        End If
                        severalCells = True
                        If patternName = "[Chart" Or patternName = "[ThinkCell" Then
                            cycleCount = 1
                        Else
                            cycleCount = valueRange.Count
                        End If
                        
                        'Exit Sub
                    Else
                        severalCells = False
                        cycleCount = 1
                    End If
                    
                    For cycleNum = 1 To cycleCount
                        If severalCells Then
                            ruleName = Replace(ruleNameX, "[x]", cycleNum)
                        End If
                        On Error GoTo CheckIsFalse
                        Select Case patternName
                        Case "[Table":
                            Set oPPTShape = currentSlide.Shapes(ruleName).Table.Cell(valueRange(cycleNum).Row - valueRange.Row + tableStartRow, valueRange(cycleNum).Column - valueRange.Column + tableStartCol).Shape
                        Case "[ThinkCell":
                            Set tcaddin = Application.COMAddIns("thinkcell.addin").Object
                        Case Else
                            Set oPPTShape = currentSlide.Shapes(ruleName)
                        End Select
                        On Error GoTo 0
                        'MsgBox "Found shape " & oPPTShape.Name & " on slide " & slideNum & ". New value: " & pptSheet.Cells(skipRows + ruleNum, 2).Value
                        If patternName = "[Chart" Then
                            Set chartSheet = oPPTShape.Chart.ChartData.Workbook.Worksheets(1)
                            valueRange.Copy
                            chartSheet.Range("A1").PasteSpecial xlPasteValues
                            Application.CutCopyMode = False
                            pptSheet.Cells(skipRows + ruleNum, tempCol).Value = slideNum
                            pptSheet.Cells(skipRows + ruleNum, resultCol).Value = 1
                            GoTo NextRule
                        End If
                        If patternName = "[ThinkCell" Then
                            If pptSheet.Cells(skipRows + ruleNum, tempCol).Value <> -2 Then
                                On Error GoTo CheckIsFalse
                                Call tcaddin.UpdateChart(oPPTPres, ruleName, valueRange, False)
                                On Error GoTo 0
                                pptSheet.Cells(skipRows + ruleNum, tempCol).Value = -2
                                pptSheet.Cells(skipRows + ruleNum, resultCol).Value = 0
                            End If
                            GoTo NextRule
                        End If
                        If severalCells Then
                            Set ruleValue = valueRange(cycleNum)
                            If bgColorRange Is Nothing Then
                                Set bgColorCell = Nothing
                            Else
                                Set bgColorCell = bgColorRange(cycleNum)
                            End If
                            If fontColorRange Is Nothing Then
                                Set fontColorCell = Nothing
                            Else
                                Set fontColorCell = fontColorRange(cycleNum)
                            End If
                        Else
                            Set ruleValue = pptSheet.Cells(skipRows + ruleNum, valueCol)
                            If pptSheet.Cells(skipRows + ruleNum, bgColorCol).Value <> "" Then
                                Set bgColorCell = pptSheet.Cells(skipRows + ruleNum, bgColorCol)
                            Else
                                Set bgColorCell = Nothing
                            End If
                            If pptSheet.Cells(skipRows + ruleNum, fontColorCol).Value <> "" Then
                                Set fontColorCell = pptSheet.Cells(skipRows + ruleNum, fontColorCol)
                            Else
                                Set fontColorCell = Nothing
                            End If
                        End If
                        changeValue = True
                        If patternName <> "[Table" Then
                            If oPPTShape.Type = msoGroup Then '6
                                changeValue = False
                            End If
                        End If
                        If changeValue Then
                            If pptSheet.Cells(skipRows + ruleNum, partCol).Value <> "" Then
                                If pptSheet.Cells(skipRows + ruleNum, totalPartCol).Value <> oPPTShape.TextFrame2.TextRange.Runs.Count Then
                                    pptSheet.Cells(skipRows + ruleNum, tempCol).Value = -3
                                    pptSheet.Cells(skipRows + ruleNum, resultCol).Interior.color = RGB(255, 255, 0)
                                    GoTo NextRule
                                End If
                                pptSheet.Cells(skipRows + ruleNum, resultCol).Interior.color = RGB(255, 255, 255)
                                pptSheet.Cells(skipRows + ruleNum, prevCol).Value = oPPTShape.TextFrame2.TextRange.Runs(pptSheet.Cells(skipRows + ruleNum, partCol).Value, 1).text
                            Else
                                pptSheet.Cells(skipRows + ruleNum, prevCol).Value = oPPTShape.TextFrame.TextRange.text
                            End If
                        End If
                        If pptSheet.Cells(skipRows + ruleNum, valueCol).NumberFormat = "General" Then
                            newValue = ruleValue.Value
                        Else
                            newValue = Format(ruleValue.Value, ruleValue.NumberFormat)
                        End If
                        skipChange = False
                        If pptSheet.Cells(skipRows + ruleNum, prevCol).Value = newValue Then
                            If pptSheet.Cells(skipRows + ruleNum, tempCol).Value = "" Then
                                pptSheet.Cells(skipRows + ruleNum, tempCol).Value = -1
                                skipChange = True
                            End If
                        End If
                        If (Not fontColorCell Is Nothing) Or (Not bgColorCell Is Nothing) Then skipChange = False
                        If Not skipChange Then
                            If changeValue Then
                                If pptSheet.Cells(skipRows + ruleNum, partCol).Value <> "" Then
                                    If ruleValue.NumberFormat = "General" Then
                                        oPPTShape.TextFrame2.TextRange.Runs(pptSheet.Cells(skipRows + ruleNum, partCol).Value, 1).text = ruleValue.Value
                                    Else
                                        oPPTShape.TextFrame2.TextRange.Runs(pptSheet.Cells(skipRows + ruleNum, partCol).Value, 1).text = Format(ruleValue.Value, ruleValue.NumberFormat)
                                    End If
                                Else
                                    If ruleValue.NumberFormat = "General" Then
                                        oPPTShape.TextFrame.TextRange.text = ruleValue.Value
                                    Else
                                        oPPTShape.TextFrame.TextRange.text = Format(ruleValue.Value, ruleValue.NumberFormat)
                                    End If
                                End If
                            End If
                            If Not bgColorCell Is Nothing Then
                                bgColorRGB = GetRGBFromText(bgColorCell.Value)
                                If bgColorRGB <> -1 Then
                                    oPPTShape.Fill.ForeColor.RGB = bgColorRGB
                                End If
                            End If
                            If Not fontColorCell Is Nothing Then
                                fontColorRGB = GetRGBFromText(fontColorCell.Value)
                                If fontColorRGB <> -1 Then
                                    oPPTShape.TextFrame.TextRange.Font.color = fontColorRGB
                                End If
                            End If
                            pptSheet.Cells(skipRows + ruleNum, resultCol).Value = pptSheet.Cells(skipRows + ruleNum, resultCol).Value + 1
                            pptSheet.Cells(skipRows + ruleNum, tempCol).Value = slideNum
                        End If
                        GoTo NextRule
CheckIsFalse:
                        Resume NextRule
NextRule:
                    Next cycleNum
    '                For Each oPPTShape In currentSlide.Shapes
    '                    If pptSheet.Cells(skipRows + ruleNum, 1).Value = oPPTShape.Name Then
    '                        MsgBox "Found shape " & oPPTShape.Name & " on slide " & slideNum & ". New value: " & pptSheet.Cells(skipRows + ruleNum, 2).Value
                            
                            'oPPTShape.TextFrame.TextRange.text = Format(pptSheet.Cells(skipRows + ruleNum, 2).Value, pptSheet.Cells(skipRows + ruleNum, 2).NumberFormat)
    '                    End If
    '                Next oPPTShape
                Next ruleNum
            Next slideNum
            For ruleNum = 1 To totalRules
                If pptSheet.Cells(skipRows + ruleNum, resultCol).Value = 0 Then
                    If pptSheet.Cells(skipRows + ruleNum, tempCol).Value = -1 Then
                        pptSheet.Cells(skipRows + ruleNum, resultCol).Value = "Value is the same and wasn't changed"
                    ElseIf pptSheet.Cells(skipRows + ruleNum, tempCol).Value = -2 Then
                        pptSheet.Cells(skipRows + ruleNum, resultCol).Value = "ThinkCell chart was updated"
                    ElseIf pptSheet.Cells(skipRows + ruleNum, tempCol).Value = -3 Then
                        pptSheet.Cells(skipRows + ruleNum, resultCol).Value = "Total # of parts was changed. Please, reselect part"
                    Else
                        pptSheet.Cells(skipRows + ruleNum, resultCol).Value = "Name was not found"
                    End If
                ElseIf pptSheet.Cells(skipRows + ruleNum, resultCol).Value = 1 Then
                    pptSheet.Cells(skipRows + ruleNum, resultCol).Value = "Updated " & pptSheet.Cells(skipRows + ruleNum, resultCol).Value & " shape on slide " & pptSheet.Cells(skipRows + ruleNum, tempCol).Value
                Else
                    pptSheet.Cells(skipRows + ruleNum, resultCol).Value = "Updated " & pptSheet.Cells(skipRows + ruleNum, resultCol).Value & " shapes"
                End If
                pptSheet.Cells(skipRows + ruleNum, tempCol).Value = ""
                pptSheet.Cells(2, 2).Value = Now
            Next ruleNum
        End If
        bar.exitBar
        Set bar = Nothing
        DisconnectFromPowerPoint
        Set currentSlide = Nothing
    Else
        FormSelectPPT.Show
    End If

End Sub




