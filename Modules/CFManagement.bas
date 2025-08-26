Attribute VB_Name = "CFManagement"
Function GetCFList() As Variant
    
    Dim CFList() As Variant
    Dim totalCF As Integer, curSheet As Worksheet, lastCell As String, currentCF As Integer
    Dim rCell As Range
    Dim allCF As Collection
    Dim cfSel As Variant
    Dim cfTestFC As FormatCondition, cfTestCSC As ColorScaleCriteria, cfTestISC As IconSetCondition
    Dim cfTestTop As Top10
    
    totalCF = 0
    currentCF = 1
    For k = 1 To Application.Sheets.Count
        Set curSheet = Application.Sheets(k)
        lastCell = FindLastCell(curSheet.Name)
        If lastCell <> "" Then
            totalCF = totalCF + curSheet.Range("A1:" & lastCell).FormatConditions.Count
        End If
    Next k
    ReDim CFList(1 To totalCF, 1 To 12)
    For k = 1 To Application.Sheets.Count
        Set curSheet = Application.Sheets(k)
        lastCell = FindLastCell(curSheet.Name)
        If lastCell <> "" Then
            For i = 1 To curSheet.Range("A1:" & lastCell).FormatConditions.Count
                Set cfSel = curSheet.Range("A1:" & lastCell).FormatConditions(i)
                CFList(currentCF, 1) = FCTypeFromIndex(cfSel.Type)
                CFList(currentCF, 2) = curSheet.Name
                CFList(currentCF, 3) = cfSel.AppliesTo.Address
                On Error Resume Next
                CFList(currentCF, 4) = cfSel.Formula1
                CFList(currentCF, 5) = cfSel.Formula2
                CFList(currentCF, 6) = cfSel.ColorScaleCriteria(1).Value
                CFList(currentCF, 7) = cfSel.ColorScaleCriteria(2).Value
                CFList(currentCF, 8) = cfSel.IconCriteria(1).Value
                CFList(currentCF, 9) = cfSel.IconCriteria(2).Value
                CFList(currentCF, 10) = cfSel.IconCriteria(3).Value
                CFList(currentCF, 11) = cfSel.IconCriteria(4).Value
                CFList(currentCF, 12) = cfSel.IconCriteria(5).Value
                On Error GoTo 0
                currentCF = currentCF + 1
            Next i
        End If
    Next k
    
    'PrintArray CFList, ActiveWorkbook.Worksheets("CFs").[A2]
    GetCFList = CFList
    
End Function

Function FCTypeFromIndex(lIndex As Long) As String
   
    Select Case lIndex
        Case 12: FCTypeFromIndex = "Above Average"
        Case 10: FCTypeFromIndex = "Blanks"
        Case 1: FCTypeFromIndex = "Cell Value"
        Case 3: FCTypeFromIndex = "Color Scale"
        Case 4: FCTypeFromIndex = "DataBar"
        Case 16: FCTypeFromIndex = "Errors"
        Case 2: FCTypeFromIndex = "Expression"
        Case 6: FCTypeFromIndex = "Icon Sets"
        Case 14: FCTypeFromIndex = "No Blanks"
        Case 17: FCTypeFromIndex = "No Errors"
        Case 9: FCTypeFromIndex = "Text"
        Case 11: FCTypeFromIndex = "Time Period"
        Case 5: FCTypeFromIndex = "Top 10?"
        Case 8: FCTypeFromIndex = "Unique Values"
        Case Else: FCTypeFromIndex = "Unknown"
    End Select
       
End Function

Sub PrintArray(ByVal Data As Variant, Cl As Range)
    i1 = UBound(Data, 1)
    i2 = UBound(Data, 2)
    For i = 1 To i1
        For k = 1 To i2
            If Left(Data(i, k), 1) = "=" Then Data(i, k) = "'" & Data(i, k)
        Next k
    Next i
    Cl.Resize(i1, i2) = Data
End Sub
