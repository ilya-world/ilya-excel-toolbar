VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormUpdate 
   Caption         =   "Update"
   ClientHeight    =   3645
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4704
   OleObjectBlob   =   "FormUpdate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDownload_Click()
    OpenWebLink "http://legalov.ru/accenture/AccentureExcelToolbar_LastVersion_Installer.exe"
End Sub

Private Sub btnLater_Click()
    FormUpdate.Hide
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    
    Dim historyText As String, rand As Integer, historyArray As Variant, totalLines As Integer, currentVersionString As String
    rand = WorksheetFunction.RandBetween(1, 1000)
    historyText = DownloadTextFile("http://legalov.ru/accenture/history.txt?a=" & rand) 'to avoid cache
    historyArray = Split(historyText, Chr(10))
    totalLines = UBound(historyArray)
    historyText = ""
    For i = 0 To totalLines
        If i = 1 Then
            currentVersionString = Replace(historyArray(i), Chr(10), "")
        End If
        If Not IsNumeric(historyArray(i)) Then
            historyText = historyText & Chr(10) & historyArray(i)
        Else
            If CDec(historyArray(i)) <= ThisWorkbook.Sheets("General").Range("B1").Value Then
                Exit For
            End If
        End If
    Next
    LabelNewVersion.Caption = "New " & currentVersionString & "is available!"
    BoxChangelog.text = historyText
    BoxChangelog.SetFocus
    BoxChangelog.CurLine = 0
    
End Sub
