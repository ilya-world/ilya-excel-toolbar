VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSelectPPT 
   Caption         =   "Select PowerPoint presentation"
   ClientHeight    =   1875
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   9432.001
   OleObjectBlob   =   "FormSelectPPT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSelectPPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbOpenedPresentations_Change()
    lFullPath.Caption = "Full path: " & cbOpenedPresentations.Value
End Sub

Private Sub cmdBrowse_Click()
    If Not WorksheetExists(ThisWorkbook.Sheets("General").Range("B5").Value) Then
        ThisWorkbook.Sheets(ThisWorkbook.Sheets("General").Range("B5").Value).Copy After:=ActiveSheet
    End If
    With Application.FileDialog(msoFileDialogOpen)
        .title = "Select PowerPoint presentation to link"
        .AllowMultiSelect = False
        .Show
        ActiveWorkbook.Sheets(ThisWorkbook.Sheets("General").Range("B5").Value).Range("B1").Value = .SelectedItems(1)
    End With
    Unload FormSelectPPT
End Sub

Private Sub cmdCancel_Click()
    Unload FormSelectPPT
End Sub

Private Sub cmdOK_Click()
    If Not WorksheetExists(ThisWorkbook.Sheets("General").Range("B5").Value) Then
        ThisWorkbook.Sheets(ThisWorkbook.Sheets("General").Range("B5").Value).Copy After:=ActiveSheet
    End If
    ActiveWorkbook.Sheets(ThisWorkbook.Sheets("General").Range("B5").Value).Range("B1").Value = cbOpenedPresentations.Value
    Unload FormSelectPPT
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Set oPPTApp = New PowerPoint.Application

    For Each pres In oPPTApp.Presentations
       cbOpenedPresentations.AddItem pres.FullName
    Next
    
    Set oPPTApp = Nothing
    Set pres = Nothing
End Sub
