VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
End Sub

Private Sub Workbook_Open()
    
    myTag = "show"
    myImage = "Risks"
    visibleTag = True
    If ThisWorkbook.Sheets("General").Range("B2").Value = 1 Then
        updateVisibleTag = True
    Else
        updateVisibleTag = False
    End If
    'Copy cusom formats from user file on first launch
    If ThisWorkbook.Sheets("General").Range("B4").Value = 1 Then
        CopyValuesFromDataWorkbook "Formats"
        CopyValuesFromDataWorkbook "Symbols"
        CopyValuesFromDataWorkbook "CustomSettings"
        ThisWorkbook.Sheets("General").Range("B4").Value = 0
    End If
End Sub
