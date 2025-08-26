VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddSheetsToTemplates 
   Caption         =   "Add to templates"
   ClientHeight    =   3570
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   2412
   OleObjectBlob   =   "AddSheetsToTemplates.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddSheetsToTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    AddSheetsToTemplates.Hide
End Sub

Private Sub btnOK_Click()
    Dim currentBook As Workbook
    Set currentBook = ActiveWorkbook
    Application.ScreenUpdating = False
    OpenDataWorkbook
    For i = 1 To lbSheets.ListCount
        If lbSheets.Selected(i - 1) Then
            CopySheetToDataWorkbook currentBook, lbSheets.List(i - 1)
        End If
    Next
    CloseDataWorkbook
    Application.ScreenUpdating = True
    AddSheetsToTemplates.Hide
End Sub

Private Sub UserForm_Click()
    
End Sub

Private Sub UserForm_Initialize()
    lbSheets.List = GetSheetList()
End Sub
