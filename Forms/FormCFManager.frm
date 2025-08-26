VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCFManager 
   Caption         =   "Manage Conditional Formating"
   ClientHeight    =   5670
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   13140
   OleObjectBlob   =   "FormCFManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCFManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnRefresh_Click()
    RefreshCFList
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    RefreshCFList
End Sub

Sub RefreshCFList()
    listCF.List = GetCFList
End Sub
