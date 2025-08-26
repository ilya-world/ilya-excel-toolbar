VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSelectParts 
   Caption         =   "Select parts to replace"
   ClientHeight    =   4305
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6852
   OleObjectBlob   =   "FormSelectParts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSelectParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
    Unload FormSelectParts
End Sub

Private Sub btnOK_Click()
    For i = 1 To lbParts.ListCount
        If lbParts.Selected(i - 1) Then
            ActiveWorkbook.Sheets(ThisWorkbook.Sheets("General").Range("B5").Value).Cells(Selection.Row, partCol).Value = lbParts.List(i - 1)
            ActiveWorkbook.Sheets(ThisWorkbook.Sheets("General").Range("B5").Value).Cells(Selection.Row, totalPartCol).Value = lbParts.ListCount
        End If
    Next i
    Unload FormSelectParts
End Sub

Private Sub UserForm_Click()

End Sub


Sub RefreshPartsList()
    
    If WorksheetExists(ThisWorkbook.Sheets("General").Range("B5").Value) And Selection.Count = 1 Then
    
        Dim currentRule As Integer
        Dim currentSlide As PowerPoint.Slide
        Dim oPPTPres As PowerPoint.Presentation
        Dim slideNum As Integer
        Dim foundShape As Boolean
        
        ConnectToPowerPoint
        
        foundShape = False
        
        If Selection.Row > skipRows And Selection.Row <= totalRules + skipRows Then
            currentRule = Selection.Row - skipRows
            Set oPPTApp = New PowerPoint.Application
            For Each pres In oPPTApp.Presentations
               If pres.FullName = strPresPath Then
                  ' found it!
                  Set oPPTPres = pres
                  Exit For
               End If
            Next
            If oPPTPres Is Nothing Then
                Set oPPTPres = oPPTApp.Presentations.Open(strPresPath)
            End If
            
            For slideNum = 1 To oPPTPres.Slides.Count
                Set currentSlide = oPPTPres.Slides(slideNum)
                On Error GoTo CheckIsFalsePart
                Set oPPTShape = currentSlide.Shapes(pptSheet.Cells(Selection.Row, 1).Value)
                On Error GoTo 0
                foundShape = True
                With oPPTShape.TextFrame2.TextRange
                    Dim PartsArray() As Variant, i As Integer
                    ReDim PartsArray(1 To .Runs.Count, 1 To 3)
                    For i = 1 To .Runs.Count
                        PartsArray(i, 1) = i
                        PartsArray(i, 2) = .Runs(i).text
                        PartsArray(i, 3) = "None"
                        If Left(.Runs(i).text, 1) = " " Then
                            PartsArray(i, 3) = "Beginning"
                        End If
                        If Right(.Runs(i).text, 1) = " " Then
                            If Left(.Runs(i).text, 1) = " " Then
                                PartsArray(i, 3) = "Beginning&End"
                            Else
                                PartsArray(i, 3) = "End"
                            End If
                        End If
                    Next i
                End With
                lbParts.List = PartsArray
                GoTo NextSlide
CheckIsFalsePart:
                Resume NextSlide

NextSlide:
            Next slideNum
            
            DisconnectFromPowerPoint
            Set currentSlide = Nothing
            If Not foundShape Then
                MsgBox "Shape with name '" & pptSheet.Cells(Selection.Row, 1).Value & "' was not found in PowerPoint", vbCritical + vbOKOnly, "Shape not found"
                
            End If
        End If
    End If

End Sub

Private Sub UserForm_Initialize()
    RefreshPartsList
End Sub
