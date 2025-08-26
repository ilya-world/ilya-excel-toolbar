VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetManager 
   Caption         =   "Sheet Manager"
   ClientHeight    =   6060
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8628.001
   OleObjectBlob   =   "SheetManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SheetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCopyToClipboard_Click()
    Dim sheetsString As String
    sheetsString = ""
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    For i = 1 To SheetList.ListCount
        If SheetList.Selected(i - 1) Then
            sheetsString = sheetsString & SheetList.List(i - 1) & vbNewLine
        End If
    Next
    MSForms_DataObject.SetText sheetsString
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

Private Sub btnHideSheets_Click()
    ChangeSheetVisibility "hide"
End Sub

Private Sub btnMoveDown_Click()
    
    Dim movedName As String, modifier As Integer
    movedName = ""
    modifier = 0
    For i = SheetList.ListCount To 1 Step -1
        If SheetList.Selected(i - 1) Then
            If i < SheetList.ListCount Then
                If Not Application.Sheets(i + 1).Name = movedName Then
                    Application.Sheets(i).Move After:=Application.Sheets(i + 1)
                    modifier = 1
                End If
            End If
            movedName = Application.Sheets(i + modifier).Name
        End If
    Next
    RefreshSheetList
    
End Sub

Private Sub btnMoveUp_Click()
    
    Dim movedName As String, modifier As Integer
    movedName = ""
    modifier = 0
    For i = 1 To SheetList.ListCount
        If SheetList.Selected(i - 1) Then
            If i > 1 Then
                If Not Application.Sheets(i - 1).Name = movedName Then
                    Application.Sheets(i).Move Before:=Application.Sheets(i - 1)
                    modifier = -1
                End If
            End If
            movedName = Application.Sheets(i + modifier).Name
        End If
    Next
    RefreshSheetList

End Sub

Private Sub btnShowSheets_Click()
    ChangeSheetVisibility "show"
End Sub

Private Sub btnVBAHideSheets_Click()
    ChangeSheetVisibility "vba_hide"
End Sub

Private Sub cbApplyToAll_Click()
    For i = 0 To SheetList.ListCount - 1
        SheetList.Selected(i) = cbApplyToAll.Value
    Next
End Sub

Private Sub cbColors_Change()
    If cbColors.Value <> "" Then
        For i = 1 To SheetList.ListCount
            If SheetList.Selected(i - 1) Then
                If GetIndexFromColor(cbColors.Value) = 0 Then
                    Application.Sheets(i).Tab.ColorIndex = xlNone
                Else
                    Application.Sheets(i).Tab.ColorIndex = GetIndexFromColor(cbColors.Value)
                End If
            End If
        Next
        RefreshSheetList
        cbColors.Value = ""
    End If
End Sub

Private Sub cmdCustomColor_Click()
    
    Dim colorR As Integer, colorG As Integer, colorB As Integer
    colorR = Val(tbColorR.Value)
    colorG = Val(tbColorG.Value)
    colorB = Val(tbColorB.Value)
    If colorR > 255 Or colorG > 255 Or colorB > 255 Or colorR < 0 Or colorG < 0 Or colorB < 0 Then
        MsgBox "Please enter numbers between 0 and 255", vbOKOnly, "Input error"
        Exit Sub
    End If
    For i = 1 To SheetList.ListCount
        If SheetList.Selected(i - 1) Then
            Application.Sheets(i).Tab.color = RGB(colorR, colorG, colorB)
        End If
    Next
    RefreshSheetList
    
End Sub

Private Sub UserForm_Click()
    'RefreshTabList
End Sub

Sub RefreshSheetList()

    Dim SheetArray() As Variant, i As Integer, colorName As String
    ReDim SheetArray(1 To Application.Sheets.Count, 1 To 4)
    For i = 1 To Application.Sheets.Count
        SheetArray(i, 1) = Application.Sheets(i).Name
        Select Case CStr(Application.Sheets(i).visible)
            Case Is = "-1": SheetArray(i, 2) = "Visible"
            Case Is = "0": SheetArray(i, 2) = "Hidden"
            Case Is = "2": SheetArray(i, 2) = "VBA Hidden"
            Case Else: SheetArray(i, 2) = "Unknown"
        End Select
        colorName = GetColorFromIndex(Application.Sheets(i).Tab.ColorIndex)
        If colorName = "Custom" Then
            If Application.Sheets(i).Tab.ColorIndex = xlNone Then
                colorName = "No color"
            Else
                colorName = getSheetRGB(i)
            End If
        End If
        SheetArray(i, 3) = colorName
        For J = 1 To SheetList.ListCount
            If SheetList.List(J - 1, 0) = SheetArray(i, 1) Then
                SheetArray(i, 4) = SheetList.Selected(J - 1)
            End If
        Next
    Next
    
    SheetList.List = SheetArray
    
    For i = 1 To Application.Sheets.Count
        If SheetArray(i, 4) Then SheetList.Selected(i - 1) = True
    Next
    
End Sub

Private Sub UserForm_Initialize()
    RefreshSheetList
    cbColors.Clear
    With cbColors
        .AddItem "No color"
        .AddItem "Red"
        .AddItem "Orange"
        .AddItem "Yellow"
        .AddItem "Green"
        .AddItem "Lime"
        .AddItem "Blue"
        .AddItem "Pink"
        .AddItem "Gold"
        .AddItem "White"
        .AddItem "Black"
    End With
End Sub

Function ChangeSheetVisibility(action As String)

    For i = 1 To SheetList.ListCount
        If SheetList.Selected(i - 1) Then
            Select Case action
                Case Is = "hide": Application.Sheets(i).visible = False
                Case Is = "show": Application.Sheets(i).visible = True
                Case Is = "vba_hide": Application.Sheets(i).visible = xlVeryHidden
            End Select
        End If
    Next
    RefreshSheetList
    
End Function
