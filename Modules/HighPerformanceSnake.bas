Attribute VB_Name = "HighPerformanceSnake"
Const goLeft As String = "L3"
Const goRight As String = "N3"
Const goUp As String = "M2"
Const goDown As String = "M4"
Dim snakeMoves As Boolean, win As Boolean
Dim direction As String
Dim foodCell As Integer, foodR As Integer, foodC As Integer

Sub OpenSnake()
    Dim isRun As Boolean
    isRun = False
    
    Application.ScreenUpdating = False
    OpenDataWorkbook
    Dim snakeSheet As Worksheet
    Set snakeSheet = Workbooks("AccentureToolbarUserData.xlsx").Worksheets("SnakeData")
    If ThisWorkbook.Sheets("SnakeData").Range("N11").Value = "New" Then
        If snakeSheet.Range("A1") = "" Or IsEmpty(snakeSheet.Range("A1")) Then
            ThisWorkbook.Sheets("SnakeData").Range("AA:AB").Copy destination:=snakeSheet.Range("A:B")
        Else
            snakeSheet.Range("A:B").Copy destination:=ThisWorkbook.Sheets("SnakeData").Range("AA:AB")
        End If
        ThisWorkbook.Sheets("SnakeData").Range("N11").Value = "Not new"
    Else
        'If snakeSheet.Range("A1") = "" Or IsEmpty(snakeSheet.Range("A1")) Then
        ThisWorkbook.Sheets("SnakeData").Range("AA:AB").Copy destination:=snakeSheet.Range("A:B")
        ThisWorkbook.Sheets("SnakeData").Range("N11").Value = "Not new"
        'End If
    End If
    CloseDataWorkbook
    Application.ScreenUpdating = True
    
    If WorksheetExists("HighPerformanceSnake") Then
        If ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("B10").Value = "Go" Then
            ActiveWorkbook.Worksheets("HighPerformanceSnake").Select
            ActiveSheet.Range("M3").Select
            isRun = True
            StartTimer 1, "MoveSnake()", 30000
        Else
            Application.DisplayAlerts = False
            ActiveWorkbook.Sheets("HighPerformanceSnake").Delete
            Application.DisplayAlerts = True
        End If
    End If
    If Not isRun Then
        ThisWorkbook.Sheets("HighPerformanceSnakeTempl").Copy After:=ActiveWorkbook.Sheets(Sheets.Count)
        ActiveSheet.Name = "HighPerformanceSnake"
        ZoomToWidth
        ActiveSheet.Range("M3").Select
        snakeMoves = False
        StartSnake
    End If
End Sub

Sub StartSnake()
    
    win = False
    ThisWorkbook.Sheets("SnakeData").Range("B2:I9").Value = ""
    ThisWorkbook.Sheets("SnakeData").Range("J24:J120").Value = ""
    ThisWorkbook.Sheets("SnakeData").Range("E5").Value = 1
    ThisWorkbook.Sheets("SnakeData").Range("S3").Value = 0
    ThisWorkbook.Sheets("SnakeData").Range("S2").Value = 3
    ThisWorkbook.Sheets("SnakeData").Range("S10").Value = 0
    For i = 1 To 8
        ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("AG" & (i + 1)).Value = ThisWorkbook.Sheets("SnakeData").Range("AP" & (i + 2)).Value
        ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("AL" & (i + 1)).Value = ThisWorkbook.Sheets("SnakeData").Range("AQ" & (i + 2)).Value
    Next
    ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(4 + 1, 4 + 1).Interior.color = RGB(255, 153, 0)
    ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(4 + 1, 4 + 1).Font.color = RGB(255, 255, 255)
    ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(4 + 1, 4 + 1).Value = ">"
    CreateFood
    ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(ThisWorkbook.Sheets("SnakeData").Range("S19") + 1, ThisWorkbook.Sheets("SnakeData").Range("S20") + 1).Value = ">"
    ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("B10").Value = "Go"
    StartTimer 1, "MoveSnake()", 30000

End Sub

Sub EndSnake()

    StopTimer

End Sub

Sub MoveSnake()
    
    win = False
    'Select direction
    If Not Intersect(ActiveCell, ActiveWorkbook.Worksheets("HighPerformanceSnake").Range(goLeft)) Is Nothing Then
        With ActiveWorkbook.Sheets("HighPerformanceSnake").Range("M3")
            .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
        If direction <> "right" Then direction = "left"
        snakeMoves = True
    End If
    
    If Not Intersect(ActiveCell, ActiveWorkbook.Worksheets("HighPerformanceSnake").Range(goRight)) Is Nothing Then
        With ActiveWorkbook.Sheets("HighPerformanceSnake").Range("M3")
            .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
        If direction <> "left" Then direction = "right"
        snakeMoves = True
    End If
    
    If Not Intersect(ActiveCell, ActiveWorkbook.Worksheets("HighPerformanceSnake").Range(goUp)) Is Nothing Then
        With ActiveWorkbook.Sheets("HighPerformanceSnake").Range("M3")
            .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
        If direction <> "down" Then direction = "up"
        snakeMoves = True
    End If
    
    If Not Intersect(ActiveCell, ActiveWorkbook.Worksheets("HighPerformanceSnake").Range(goDown)) Is Nothing Then
        With ActiveWorkbook.Sheets("HighPerformanceSnake").Range("M3")
            .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
        If direction <> "up" Then direction = "down"
        snakeMoves = True
    End If
    
    
    If snakeMoves Then
        With ThisWorkbook.Worksheets("SnakeData")
            'Change "Do you know" section
            Dim selectedDYK As Integer, rowDYK As Integer, textDYK As String
            .Range("S10").Value = .Range("S10").Value + 1
            If .Range("S10").Value Mod 20 = 0 Then
                If .Range("S21") = 0 Then .Range("J24:J120").Value = ""
                selectedDYK = WorksheetFunction.RandBetween(1, WorksheetFunction.Max(.Range("S21")))
                rowDYK = WorksheetFunction.Match(selectedDYK, .Range("M:M"), 0)
                .Range("J" & rowDYK).Value = 1
                textDYK = .Range("B" & rowDYK).Value
                ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("L7").Value = textDYK
            End If
            
            'Increase Tail
            For i = 1 To 8 'Row
                For J = 1 To 8 'Column
                    If .Cells(i + 1, J + 1).Value <> "" Then
                        .Cells(i + 1, J + 1).Value = .Cells(i + 1, J + 1).Value + 1
                    End If
                Next
            Next
            Dim newR As Integer, newC As Integer
            
            'Remove the last tail part
            Dim removedR As Integer, removedC As Integer
            removedR = 0
            removedC = 0
            If .Cells(.Range("S8").Value + 1, .Range("S9").Value + 1).Value > .Range("S2").Value Then
                removedR = .Range("S8").Value
                removedC = .Range("S9").Value
                .Cells(.Range("S8") + 1, .Range("S9") + 1).Value = ""
            End If
            
            ' Put the head
            If direction = "right" Then
                If .Range("S5") = 8 Then
                    Lose
                    Exit Sub
                End If
                newR = .Range("S4")
                newC = .Range("S5") + 1
            End If
            If direction = "left" Then
                If .Range("S5") = 1 Then
                    Lose
                    Exit Sub
                End If
                newR = .Range("S4")
                newC = .Range("S5") - 1
            End If
            If direction = "up" Then
                If .Range("S4") = 1 Then
                    Lose
                    Exit Sub
                End If
                newR = .Range("S4") - 1
                newC = .Range("S5")
            End If
            If direction = "down" Then
                If .Range("S4") = 8 Then
                    Lose
                    Exit Sub
                End If
                newR = .Range("S4") + 1
                newC = .Range("S5")
            End If
            If .Cells(newR + 1, newC + 1) <> "" Then
                Lose
                Exit Sub
            End If
            .Cells(newR + 1, newC + 1) = 1
            
            'Consume and create food
            Dim foodCreated As Boolean
            foodCreated = False
            If .Range("S19").Value = newR And .Range("S20").Value = newC Then
                .Range("S2").Value = .Range("S2").Value + 1 'Increase length
                If .Range("S2").Value = 64 Then 'Win condition
                    win = True
                    ActiveWorkbook.Sheets("HighPerformanceSnake").Range("B2:I9").Interior.color = RGB(255, 153, 0)
                    ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(newR + 1, newC + 1).Font.color = RGB(255, 255, 255)
                    ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(newR + 1, newC + 1).Value = ">"
                    ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(.Range("S4") + 1, .Range("S5") + 1).Value = ""
                    Lose
                    Exit Sub
                End If
                CreateFood
                foodCreated = True
                .Range("S3").Value = .Range("S3").Value + 21 'Increase score
            End If
            
            
            'Update the field
            ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(newR + 1, newC + 1).Interior.color = RGB(255, 153, 0)
            ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(newR + 1, newC + 1).Font.color = RGB(255, 255, 255)
            ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(newR + 1, newC + 1).Value = ">"
            If removedR <> 0 And (newR <> removedR Or newC <> removedC) Then
                ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(removedR + 1, removedC + 1).Interior.color = RGB(255, 255, 255)
                ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(removedR + 1, removedC + 1).Font.color = RGB(0, 0, 0)
            End If
            ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(.Range("S4") + 1, .Range("S5") + 1).Value = ""
            If foodCreated Then ActiveWorkbook.Sheets("HighPerformanceSnake").Cells(.Range("S19") + 1, .Range("S20") + 1).Value = ">"
            
            'Update position
            Select Case .Range("S2").Value
                Case 3 To 4: ActiveWorkbook.Sheets("HighPerformanceSnake").Range("Q11").Value = "Intern"
                Case 5 To 14: ActiveWorkbook.Sheets("HighPerformanceSnake").Range("Q11").Value = "Analyst"
                Case 15 To 19: ActiveWorkbook.Sheets("HighPerformanceSnake").Range("Q11").Value = "Consultant"
                Case 20 To 29: ActiveWorkbook.Sheets("HighPerformanceSnake").Range("Q11").Value = "Manager"
                Case 30 To 39: ActiveWorkbook.Sheets("HighPerformanceSnake").Range("Q11").Value = "Senior Manager"
                Case 40 To 49: ActiveWorkbook.Sheets("HighPerformanceSnake").Range("Q11").Value = "Managing Director"
                Case 50 To 59: ActiveWorkbook.Sheets("HighPerformanceSnake").Range("Q11").Value = "Senior Country Managing Director"
                Case 60 To 64: ActiveWorkbook.Sheets("HighPerformanceSnake").Range("Q11").Value = "HOLY SHIT! Pierre Nanterme!"
            End Select
            
            'Update score
            .Range("S3").Value = .Range("S3").Value - 1 'Minus move point
            If .Range("S3").Value < 0 Then .Range("S3").Value = 0
            ActiveWorkbook.Sheets("HighPerformanceSnake").Range("F11").Value = .Range("S3").Value
            
        End With
        
        
    End If
    
    
    
End Sub

Sub Lose()
    If win Then
        ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("L7").Value = "You won!!! Sharing a screenshot of this message can make significant impact on your career advancement. Push the secret button again to replay"
    Else
        ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("L7").Value = "You are Delivered. Push the secret button again to replay"
    End If
    ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("B10").Value = "Stop"
    EndSnake
    With ThisWorkbook.Sheets("SnakeData")
        .Range("AB" & (.Range("AI1") + 1)).Value = .Range("S3").Value
        .Range("AA" & (.Range("AI1") + 1)).Value = Environ("USERNAME")
        For i = 1 To 8
            ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("AG" & (i + 1)).Value = .Range("AP" & (i + 2)).Value
            ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("AL" & (i + 1)).Value = .Range("AQ" & (i + 2)).Value
        Next
        If .Range("AI2") <= 8 Then
            ActiveWorkbook.Worksheets("HighPerformanceSnake").Range("AG" & (.Range("AI2") + 1) & ":AL" & (.Range("AI2") + 1)).Interior.color = RGB(255, 221, 0)
        End If
        If Not win Then
            ActiveWorkbook.Worksheets("HighPerformanceSnake").Cells(.Range("S4") + 1, .Range("S5") + 1).Interior.color = RGB(217, 13, 57)
        End If
    End With
    direction = ""
    
    
End Sub

Sub CreateFood()
    
    foodCell = WorksheetFunction.RandBetween(1, ThisWorkbook.Sheets("SnakeData").Range("S13").Value)
    ThisWorkbook.Sheets("SnakeData").Range("S15").Value = foodCell
    foodR = ThisWorkbook.Sheets("SnakeData").Range("S16").Value
    foodC = ThisWorkbook.Sheets("SnakeData").Range("S17").Value
    ThisWorkbook.Sheets("SnakeData").Range("S19").Value = foodR
    ThisWorkbook.Sheets("SnakeData").Range("S20").Value = foodC
    
End Sub

Function WorksheetExists(WSName As String) As Boolean
    On Error Resume Next
    WorksheetExists = ActiveWorkbook.Worksheets(WSName).Name = WSName
    On Error GoTo 0
End Function
