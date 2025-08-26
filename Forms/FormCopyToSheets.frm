VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCopyToSheets 
   Caption         =   "Copy to sheets"
   ClientHeight    =   3765
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   2340
   OleObjectBlob   =   "FormCopyToSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCopyToSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    FormCopyToSheets.Hide
End Sub

Private Sub btnOK_Click()
    
    Dim selShape As Shape, selRange As Range, pastedShape As Shape, pastedShapeRange As ShapeRange
    Dim cycleSheet As Worksheet, selArea As Range
    For i = 0 To lbSheets.ListCount - 1
        If lbSheets.Selected(i) Then
            Set cycleSheet = ActiveWorkbook.Worksheets(lbSheets.List(i))
            If TypeName(Selection) <> "Range" Then
                For Each selShape In Selection.ShapeRange
                    selShape.Copy
                    cycleSheet.Paste
                    Set pastedShape = cycleSheet.Shapes(cycleSheet.Shapes.Count)
                    pastedShape.Top = selShape.Top
                    pastedShape.Left = selShape.Left
                Next selShape
            Else
                For Each selArea In Selection.Areas
                    selArea.Copy destination:=cycleSheet.Range(selArea.Address)
                Next selArea
            End If
        End If
    Next i
    
    FormCopyToSheets.Hide
    
End Sub

Private Sub cbApplyToAll_Click()
    For i = 0 To lbSheets.ListCount - 1
        lbSheets.Selected(i) = cbApplyToAll.Value
    Next
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()
    lbSheets.List = GetSheetList()
End Sub
