VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormTimeTest 
   Caption         =   "Calculation time"
   ClientHeight    =   5685
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   2820
   OleObjectBlob   =   "FormTimeTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormTimeTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    
    Dim selSheet As Worksheet, outputRange As Range, calcSheet As Worksheet
    
    
    PrepareWorkbook
    Set calcSheet = Worksheets("Calculation Times")
    For i = 0 To lbSheets.ListCount - 1
        If lbSheets.Selected(i) Then
            Set selSheet = Worksheets(lbSheets.List(i))
            timeSheet selSheet, calcSheet.Cells(calcSheet.Range(FindLastCell(calcSheet.Name)).Row, 1)
        End If
    Next i
    FinalizeWorkbook
    calcSheet.Select
    Range("A1").Select
    FormTimeTest.Hide

End Sub

Private Sub btnCancel_Click()
    FormTimeTest.Hide
End Sub



Private Sub cbApplyToAll_Click()
    For i = 0 To lbSheets.ListCount - 1
        lbSheets.Selected(i) = cbApplyToAll.Value
    Next
End Sub

Private Sub UserForm_Initialize()
    lbSheets.List = GetSheetList()
End Sub
