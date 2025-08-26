VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertUserSheet 
   Caption         =   "Insert sheets"
   ClientHeight    =   4320
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   2424
   OleObjectBlob   =   "InsertUserSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InsertUserSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    InsertUserSheet.Hide
End Sub

Private Sub btnDelete_Click()
    
    Application.ScreenUpdating = False
    OpenDataWorkbook
    For i = 1 To lbSheets.ListCount
        If lbSheets.Selected(i - 1) Then
            DeleteSheetFromDataWorkbook (lbSheets.List(i - 1))
        End If
    Next
    CloseDataWorkbook
    Application.ScreenUpdating = True
    RefreshTemlateList
    
End Sub

Private Sub btnInsert_Click()
    
    Dim currentBook As Workbook
    Set currentBook = ActiveWorkbook
    Application.ScreenUpdating = False
    OpenDataWorkbook
    For i = 1 To lbSheets.ListCount
        If lbSheets.Selected(i - 1) Then
            'ThisWorkbook.Sheets(tempName).Copy After:=ActiveSheet
            'ActiveSheet.Name = lbSheets.List(i - 1)
            CopySheetFromDataWorkbook currentBook, lbSheets.List(i - 1)
        End If
    Next
    CloseDataWorkbook
    Application.ScreenUpdating = True
    InsertUserSheet.Hide
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    RefreshTemlateList
End Sub

Sub RefreshTemlateList()
    Application.ScreenUpdating = False
    OpenDataWorkbook
    If Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("F3").Value > 0 Then
        Dim allSheets() As Variant, realName As String, counter As Integer
        counter = 1
        
        ReDim allSheets(1 To Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("F3").Value)
        For i = 1 To Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("F1").Value
            realName = Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("B" & i).Value
            If realName <> "" Then
                allSheets(counter) = realName
                counter = counter + 1
            End If
        Next
        lbSheets.List = allSheets
    Else
        lbSheets.Clear
    End If
    CloseDataWorkbook
    Application.ScreenUpdating = True
End Sub
