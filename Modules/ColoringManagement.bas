Attribute VB_Name = "ColoringManagement"
Sub SetTheColor(Red As Integer, Green As Integer, Blue As Integer)

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = RGB(Red, Green, Blue)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Sub nofill()

    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Sub SetTextColor(Red As Integer, Green As Integer, Blue As Integer)
    
    With Selection.Font
        .color = RGB(Red, Green, Blue)
    End With
    
End Sub
