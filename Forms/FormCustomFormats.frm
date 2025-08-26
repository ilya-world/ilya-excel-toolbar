VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCustomFormats 
   Caption         =   "Toolbar settings"
   ClientHeight    =   4980
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   7200
   OleObjectBlob   =   "FormCustomFormats.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCustomFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnApply_Click()
    Dim formatsSheet As Worksheet, imageText As String
    Set formatsSheet = ThisWorkbook.Sheets("Formats")
    For i = 1 To lbFormats.ListCount
        If lbFormats.Selected(i - 1) Then
            If opIconFromList Then
                imageText = ConvertIconTextToImage(cbIcon.Value)
            Else
                If Len(tbIcon.Value) = 1 Then
                    If tbIcon.Value >= "0" And tbIcon.Value <= "9" Then
                        imageText = "_" & tbIcon.Value
                    Else
                        tbIcon.Value = UCase(tbIcon.Value)
                        If tbIcon.Value >= "A" And tbIcon.Value <= "Z" Then
                            imageText = tbIcon.Value
                        Else
                            MsgBox "You must enter digit (0-9) or letter (A-Z) inside the field", vbCritical + vbOKOnly, "Error"
                            Exit For
                        End If
                    End If
                Else
                    MsgBox "You must enter a symbol inside the icon field or choose icon from the list", vbCritical + vbOKOnly, "Error"
                    Exit For
                End If
            End If
            formatsSheet.Cells(i + 1, 2) = tbTitle.Value
            formatsSheet.Cells(i + 1, 3) = tbDescription.Value
            formatsSheet.Cells(i + 1, 4) = imageText
            formatsSheet.Cells(i + 1, 5) = tbFormatString.Value
            lChanges.visible = True
            RefreshRibbon "*"
            CopyValuesToDataWorkbook "Formats"
            Exit For
        End If
    Next i
End Sub

Private Sub btnApplyOther_Click()
    ThisWorkbook.Sheets("CustomSettings").Range("B1").Value = tbFileName.Value
    CopyValuesToDataWorkbook "CustomSettings"
End Sub

Private Sub btnApplySymbols_Click()
    
    ThisWorkbook.Sheets("Symbols").Cells(2, 3).Value = tbSymbolCode1.text
    ThisWorkbook.Sheets("Symbols").Cells(3, 3).Value = tbSymbolCode2.text
    ThisWorkbook.Sheets("Symbols").Cells(4, 3).Value = tbSymbolCode3.text
    ThisWorkbook.Sheets("Symbols").Cells(5, 3).Value = tbSymbolCode4.text
    
    ThisWorkbook.Sheets("Symbols").Cells(2, 2).Value = tbSymbolDescription1.text
    ThisWorkbook.Sheets("Symbols").Cells(3, 2).Value = tbSymbolDescription2.text
    ThisWorkbook.Sheets("Symbols").Cells(4, 2).Value = tbSymbolDescription3.text
    ThisWorkbook.Sheets("Symbols").Cells(5, 2).Value = tbSymbolDescription4.text
    
    lSymbolesChanged.visible = True
    RefreshRibbon "*"
    CopyValuesToDataWorkbook "Symbols"
    
End Sub

Private Sub btnClose_Click()
    lChanges.visible = False
    lSymbolesChanged.visible = False
    FormCustomFormats.Hide
End Sub

Private Sub cbIcon_Change()
    opIconFromList.Value = True
End Sub

Private Sub lbFormats_Click()
    
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Sheets("Formats")
    lChanges.visible = False
    For i = 1 To lbFormats.ListCount
        If lbFormats.Selected(i - 1) Then
            tbFormatString.text = formatsSheet.Cells(i + 1, 5)
            tbTitle.text = formatsSheet.Cells(i + 1, 2)
            tbDescription.text = formatsSheet.Cells(i + 1, 3)
            'MsgBox Len(formatsSheet.Cells(i + 1, 4))
            If Len(formatsSheet.Cells(i + 1, 4)) = 1 Then
                cbIcon.Value = ""
                tbIcon.text = formatsSheet.Cells(i + 1, 4)
                opIconFromList.Value = False
                opIconFromLetter.Value = True
            ElseIf Len(formatsSheet.Cells(i + 1, 4)) = 2 Then
                cbIcon.Value = ""
                tbIcon.text = Right(formatsSheet.Cells(i + 1, 4), 1)
                opIconFromList.Value = False
                opIconFromLetter.Value = True
            Else
                tbIcon.text = ""
                cbIcon.Value = ConvertImageToIconText(formatsSheet.Cells(i + 1, 4))
                opIconFromLetter.Value = False
                opIconFromList.Value = True
            End If
            Exit For
        End If
    Next i
    
End Sub

Private Sub lInfo_Click()
    OpenWebLink "https://support.office.microsoft.com/en-us/article/Create-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4?CorrelationId=7f4fac7f-3e2c-440c-b6e2-363ba58c7e16&ui=en-US&rs=en-US&ad=US"
End Sub

Private Sub lUnicodeInfo_Click()
    OpenWebLink "http://unicode-table.com/en"
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub tbIcon_Change()
    opIconFromLetter.Value = True
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    RefreshButtonsList
    RefreshSymbolsList
    cbIcon.AddItem ""
    cbIcon.AddItem "$"
    cbIcon.AddItem "%"
    cbIcon.AddItem ","
    cbIcon.AddItem "abcd"
    cbIcon.AddItem "Lightning icon"
    cbIcon.AddItem "Calendar icon"
    cbIcon.AddItem "Money icon"
    tbFileName.Value = ThisWorkbook.Sheets("CustomSettings").Range("B1").Value
End Sub

Sub RefreshButtonsList()
    
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Sheets("Formats")
    For i = 1 To 6
        lbFormats.AddItem formatsSheet.Cells(i + 1, 2).Value
    Next i

End Sub

Sub RefreshSymbolsList()

    tbSymbolCode1.text = ThisWorkbook.Sheets("Symbols").Cells(2, 3).Value
    tbSymbolCode2.text = ThisWorkbook.Sheets("Symbols").Cells(3, 3).Value
    tbSymbolCode3.text = ThisWorkbook.Sheets("Symbols").Cells(4, 3).Value
    tbSymbolCode4.text = ThisWorkbook.Sheets("Symbols").Cells(5, 3).Value
    
    tbSymbolDescription1.text = ThisWorkbook.Sheets("Symbols").Cells(2, 2).Value
    tbSymbolDescription2.text = ThisWorkbook.Sheets("Symbols").Cells(3, 2).Value
    tbSymbolDescription3.text = ThisWorkbook.Sheets("Symbols").Cells(4, 2).Value
    tbSymbolDescription4.text = ThisWorkbook.Sheets("Symbols").Cells(5, 2).Value

End Sub

Function ConvertIconTextToImage(iconText As String) As String
    
    Select Case iconText
        Case Is = "$": ConvertIconTextToImage = "ApplyCurrencyFormat"
        Case Is = "%": ConvertIconTextToImage = "ApplyPercentageFormat"
        Case Is = ",": ConvertIconTextToImage = "ApplyCommaFormat"
        Case Is = "abcd": ConvertIconTextToImage = "AsianLayoutCombineCharacters"
        Case Is = "Lightning icon": ConvertIconTextToImage = "AnimationTriggerAddMenu"
        Case Is = "Calendar icon": ConvertIconTextToImage = "StatusDate"
        Case Is = "Money icon": ConvertIconTextToImage = "DataTypeEuro"
        Case Else: ConvertIconTextToImage = "AnimationTriggerAddMenu"
    End Select
    
End Function

Function ConvertImageToIconText(imageText As String) As String

    Select Case imageText
        Case Is = "ApplyCurrencyFormat": ConvertImageToIconText = "$"
        Case Is = "ApplyPercentageFormat": ConvertImageToIconText = "%"
        Case Is = "ApplyCommaFormat": ConvertImageToIconText = ","
        Case Is = "AsianLayoutCombineCharactersabcd": ConvertImageToIconText = "abcd"
        Case Is = "AnimationTriggerAddMenu": ConvertImageToIconText = "Lightning icon"
        Case Is = "StatusDate": ConvertImageToIconText = "Calendar icon"
        Case Is = "DataTypeEuro": ConvertImageToIconText = "Money icon"
        Case Else: ConvertImageToIconText = ""
    End Select

End Function
