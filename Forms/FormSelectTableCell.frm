VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSelectTableCell 
   Caption         =   "Select table cells to update"
   ClientHeight    =   4110
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   2544
   OleObjectBlob   =   "FormSelectTableCell.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSelectTableCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    selectedTableColumn = 0
    selectedTableRow = 0
    FormSelectTableCell.Hide
End Sub

Private Sub btnOK_Click()
    selectedTableColumn = cbRow.Value
    selectedTableRow = cbRow.Value
    FormSelectTableCell.Hide
End Sub

Private Sub cmdSelectBgColor_Click()
    Set bgColorGroupRange = SelectRangeOfCells("Select a range with RGB values to use as background color")
End Sub

Private Sub cmdSelectFontColor_Click()
    Set fontColorGroupRange = SelectRangeOfCells("Select a range with RGB values to use as font color")
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    For i = 1 To 20
        cbRow.AddItem i
        cbColumn.AddItem i
    Next i
    cbRow.Value = 1
    cbColumn.Value = 1
    
End Sub
