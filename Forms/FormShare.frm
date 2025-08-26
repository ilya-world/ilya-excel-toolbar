VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormShare 
   Caption         =   "Share"
   ClientHeight    =   2940
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4392
   OleObjectBlob   =   "FormShare.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    FormShare.Hide
End Sub

Private Sub btnEmail_Click()
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = ""
        .CC = ""
        .BCC = ""
        .Subject = "Check this Excel add-in"
        .Body = "I found Excel add-in ""Accenture Excel Toolbar"" that could be interesting to you. It's like QPT, but for Excel. Here is the link: " & vbNewLine & "http://blog.accenture.com/ilya_legalov/accenture-excel-toolbar/"
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub


Private Sub tbLink_Change()

End Sub

Private Sub tbLink_Enter()
    tbLink.SelStart = 0
    tbLink.SelLength = Len(tbLink.text)
End Sub
