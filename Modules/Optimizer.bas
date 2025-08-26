Attribute VB_Name = "Optimizer"
Option Explicit

' Acknowledgement for the microtimer procedures used here to
' thanks to Charles Wheeler - http://www.decisionmodels.com/

Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Sub FinalizeWorkbook()
    
    Dim calcSheet As Worksheet, rSum As Range, rKey As Range, r As Range
    Set calcSheet = Worksheets("Calculation Times")
    Set rKey = calcSheet.Range(calcSheet.Cells(2, 2).Address, calcSheet.Range(FindLastCell(calcSheet.Name)))
    'MsgBox calcSheet.Range(calcSheet.Cells(1, 1).Address, calcSheet.Range(FindLastCell(calcSheet.Name))).Address
    With calcSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=rKey, _
            SortOn:=xlSortOnValues, Order:=xlDescending, _
                DataOption:=xlSortNormal
    
            .SetRange calcSheet.Range(calcSheet.Cells(1, 1).Address, calcSheet.Range(FindLastCell(calcSheet.Name)))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
    End With
    For Each r In rKey.Cells
        r.Offset(0, 1).Formula = "=" & r.Address & "/SUM(B:B)"
        r.Offset(0, 1).NumberFormat = "0.00%"
    Next r
    calcSheet.Columns.AutoFit
    calcSheet.Cells.Interior.color = RGB(255, 255, 255)
    calcSheet.UsedRange.Borders.LineStyle = xlContinuous
    calcSheet.Range("A1:C1").Font.Bold = True
    
End Sub


Sub PrepareWorkbook()
    
    Dim calcSheet As Worksheet, where As Range
    
    If WorksheetExists("Calculation Times") Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Sheets("Calculation Times").Delete
        Application.DisplayAlerts = True
    End If
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(Sheets.Count)
    ActiveSheet.Name = "Calculation Times"
    Set calcSheet = Worksheets("Calculation Times")
    Set where = calcSheet.[A1]
    calcSheet.Cells.ClearContents
    calcSheet.Cells(1, 1).Value = "Address"
    calcSheet.Cells(1, 2).Value = "Time"
    

End Sub

Function timeSheet(ws As Worksheet, routput As Range) As Range
    Dim ro As Range, col As Integer
    Dim c As Range, ct As Range, rt As Range, u As Range
    
    col = 1
    
    ws.Activate
    Set u = ws.UsedRange
    Set ct = u.Resize(1)
    Set ro = routput

    For Each c In ct.Columns
        'Set ro = ro.Offset(1)
        Set rt = c.Resize(u.Rows.Count)
        rt.Select
        ro.Offset(col, 0).Value = rt.Worksheet.Name & "!" & rt.Address
        ro.Offset(col, 1) = shortCalcTimer(rt, False)
        col = col + 1
    Next c
    Set timeSheet = ro

End Function

Sub timeallsheets()
    Call timeloopSheets
End Sub

Sub timeloopSheets(Optional wsingle As Worksheet)
    
    Dim ws As Worksheet, ro As Range, rAll As Range
    Dim rKey As Range, r As Range, rSum As Range
    Const where = "ExecutionTimes!a1"
    
    Set ro = Range(where)
    ro.Worksheet.Cells.ClearContents
    Set rAll = ro
    'headers
    rAll.Cells(1, 1).Value = "address"
    rAll.Cells(1, 2).Value = "time"
    
    If wsingle Is Nothing Then
    ' all sheets
        For Each ws In Worksheets
            Set ro = timeSheet(ws, ro)
        Next ws
    Else
    ' or just a single one
        Set ro = timeSheet(wsingle, ro)
    End If
    
    'now sort results, if there are any
    
    If ro.Row > rAll.Row Then
        Set rAll = rAll.Resize(ro.Row - rAll.Row + 1, 2)
        Set rKey = rAll.Offset(1, 1).Resize(rAll.Rows.Count - 1, 1)
        ' sort highest to lowest execution time
        With rAll.Worksheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=rKey, _
            SortOn:=xlSortOnValues, Order:=xlDescending, _
                DataOption:=xlSortNormal
    
            .SetRange rAll
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        ' sum times
        Set rSum = rAll.Cells(1, 3)
        rSum.Formula = "=sum(" & rKey.Address & ")"
        ' %ages formulas
        For Each r In rKey.Cells
            r.Offset(, 1).Formula = "=" & r.Address & "/" & rSum.Address
            r.Offset(, 1).NumberFormat = "0.00%"
        Next r
        
    End If
    rAll.Worksheet.Activate

End Sub

Function shortCalcTimer(rt As Range, Optional bReport As Boolean = True) As Double
    Dim dTime As Double
    Dim sCalcType As String
    Dim lCalcSave As Long
    Dim bIterSave As Boolean
    '
    On Error GoTo Errhandl


    ' Save calculation settings.
    lCalcSave = Application.Calculation
    bIterSave = Application.Iteration
    If Application.Calculation <> xlCalculationManual Then
        Application.Calculation = xlCalculationManual
    End If

    ' Switch off iteration.
    If Application.Iteration <> False Then
        Application.Iteration = False
    End If

' Get start time.
    dTime = MicroTimer
    If Val(Application.version) >= 12 Then
        rt.CalculateRowMajorOrder
    Else
        rt.Calculate
    End If


' Calc duration.
    sCalcType = "Calculate " & CStr(rt.Count) & _
        " Cell(s) in Selected Range: " & rt.Address
    dTime = MicroTimer - dTime
    On Error GoTo 0

    dTime = Round(dTime, 5)
    If bReport Then
        MsgBox sCalcType & " " & CStr(dTime) & " Seconds"
    End If

    shortCalcTimer = dTime
Finish:

    ' Restore calculation settings.
    If Application.Calculation <> lCalcSave Then
         Application.Calculation = lCalcSave
    End If
    If Application.Iteration <> bIterSave Then
         Application.Calculation = bIterSave
    End If
    Exit Function
Errhandl:
    On Error GoTo 0
    MsgBox "Unable to Calculate " & sCalcType, _
        vbOKOnly + vbCritical, "CalcTimer"
    GoTo Finish
End Function
'
Function MicroTimer() As Double
'

' Returns seconds.
'
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    '
    MicroTimer = 0

' Get frequency.
    If cyFrequency = 0 Then getFrequency cyFrequency

' Get ticks.
    getTickCount cyTicks1

' Seconds
    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
End Function

