VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormRandomGenerator 
   Caption         =   "Random generator"
   ClientHeight    =   5670
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4704
   OleObjectBlob   =   "FormRandomGenerator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormRandomGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnNormalGenerate_Click()
    
    Dim cycleCell As Range, generatedNum As Double, median As Double, sigma As Double, digits As Integer, cellsCount As Long
    
    If IsNumeric(tbNormalMedian.text) And IsNumeric(tbNormalSigma.text) And IsNumeric(tbNormalDigits.text) Then
        median = tbNormalMedian.text
        sigma = tbNormalSigma.text
        digits = tbNormalDigits.text
        Me.Hide
        Dim bar As Progressbar
        Set bar = New Progressbar
        bar.createLoadingBar
        bar.createString
        bar.createtimeDuration
        bar.setParameters Selection.Cells.Count
        bar.Start
        cellsCount = 0
        For Each cycleCell In Selection
            '=SQRT(-2*LN($A2))*COS(2*PI()*$B2)*$G$5+$G$4
            generatedNum = Sqr(-2 * WorksheetFunction.Ln(Rnd())) * Cos(2 * WorksheetFunction.Pi * Rnd()) * sigma + median
            generatedNum = Round(generatedNum * (10 ^ digits), 0) / (10 ^ digits)
            cycleCell.Value = generatedNum
            cellsCount = cellsCount + 1
            bar.Update cellsCount, "Cells"
        Next cycleCell
        bar.exitBar
        Set bar = Nothing
        'Me.Show
        Unload Me
    Else
        MsgBox "All text fields must contain numbers", vbOKOnly + vbCritical, "Error"
    End If
    
End Sub

Private Sub btnUniformGenerate_Click()
    
    Dim cycleCell As Range, generatedNum As Double, numFrom As Double, numTo As Double, digits As Integer, cellsCount As Long
    If IsNumeric(tbUniformFrom.text) And IsNumeric(tbUniformTo.text) And IsNumeric(tbUniformDigits.text) Then
        numFrom = tbUniformFrom.text
        numTo = tbUniformTo.text
        digits = tbUniformDigits.text
        Me.Hide
        Dim bar As Progressbar
        Set bar = New Progressbar
        bar.createLoadingBar
        bar.createString
        bar.createtimeDuration
        bar.setParameters Selection.Cells.Count
        bar.Start
        cellsCount = 0
        For Each cycleCell In Selection
            generatedNum = WorksheetFunction.RandBetween(numFrom * (10 ^ digits), numTo * (10 ^ digits)) / (10 ^ digits)
            cycleCell.Value = generatedNum
            cellsCount = cellsCount + 1
            bar.Update cellsCount, "Cells"
        Next cycleCell
        bar.exitBar
        Set bar = Nothing
        'Me.Show
        Unload Me
    Else
        MsgBox "All text fields must contain numbers", vbOKOnly + vbCritical, "Error"
    End If
    
End Sub

Private Sub tbNormalDigits_Change()
    UpdateNormalDistrLabels
End Sub

Private Sub tbNormalMedian_Change()
    UpdateNormalDistrLabels
End Sub

Private Sub tbNormalSigma_Change()
    UpdateNormalDistrLabels
End Sub

Private Sub UserForm_Click()
    
End Sub

Private Sub UserForm_Initialize()
    
    If TypeName(Selection) = "Range" Then
        lSelectedCells.Caption = "You selected " & Selection.Count & " cells"
        UpdateNormalDistrLabels
    Else
        MsgBox "You must select a range", vbOKOnly + vbCritical, "Error"
        Unload Me
    End If
    
End Sub

Private Sub UpdateNormalDistrLabels()

    Dim median As Double, sigma As Double, digits As Integer, cellsCount As Long
    If IsNumeric(tbNormalMedian.text) And IsNumeric(tbNormalSigma.text) And IsNumeric(tbNormalDigits.text) Then
        median = tbNormalMedian.text
        sigma = tbNormalSigma.text
        digits = tbNormalDigits.text
        cellsCount = Selection.Count
        lGet999.Caption = CalculateLabelText(0.999, median, sigma, digits, cellsCount)
        lGet99.Caption = CalculateLabelText(0.99, median, sigma, digits, cellsCount)
        lGet95.Caption = CalculateLabelText(0.95, median, sigma, digits, cellsCount)
        lGet90.Caption = CalculateLabelText(0.9, median, sigma, digits, cellsCount)
    End If

End Sub

Private Function CalculateLabelText(prob As Double, median As Double, sigma As Double, digits As Integer, cellsCount As Long) As String

    Dim lowestNum As Double, highestNum As Double, percentText As String
    '=NORM.INV((1-G14)/2;G4;G5)
    lowestNum = WorksheetFunction.Norm_Inv((1 - prob) / 2, median, sigma)
    highestNum = WorksheetFunction.Norm_Inv(1 - ((1 - prob) / 2), median, sigma)
    lowestNum = Round(lowestNum * (10 ^ digits), 0) / (10 ^ digits)
    highestNum = Round(highestNum * (10 ^ digits), 0) / (10 ^ digits)
    Select Case prob
        Case 0.999: percentText = "99,9%"
        Case 0.99: percentText = "99%"
        Case 0.95: percentText = "95%"
        Case 0.9: percentText = "90%"
        Case 0.7: percentText = "70%"
    End Select
    CalculateLabelText = percentText & " of results (" & Round(cellsCount * prob, 0) & " cells) are expected to be between " & lowestNum & " and " & highestNum & " (dif: " & (highestNum - lowestNum) & ")"

End Function

