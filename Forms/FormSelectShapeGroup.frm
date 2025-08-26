VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSelectShapeGroup 
   Caption         =   "Select shape group"
   ClientHeight    =   5955
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   3372
   OleObjectBlob   =   "FormSelectShapeGroup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSelectShapeGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnNameRangeEntered_Click()
    shapeRangeName = tbExistingNames.Value
    Unload FormSelectShapeGroup
End Sub

Private Sub btnPowerPointObjectsSelected_Click()
    
    Dim slideNum As Integer
    Dim skipChange As Boolean
    Dim selectedShapes As Variant, selectedNum As Integer, notSorted As Boolean
    Dim curCell As Range
    
    shapeRangeName = ""
    
    ConnectToPowerPoint
    
    selectedNum = oPPTApp.ActiveWindow.Selection.ShapeRange.Count
    
    If selectedNum <> sourceCellRange.Count Then
        MsgBox "You must select " & sourceCellRange.Count & " shapes in PowerPoint to proceed. Currently selected: " & selectedNum
        Exit Sub
    End If
    If tbNewName.text = "" Then
        MsgBox "You must enter name for shapes in PowerPoint"
        Exit Sub
    End If
    
    i = 1
    ReDim selectedShapes(3, selectedNum)
    For Each oPPTShape In oPPTApp.ActiveWindow.Selection.ShapeRange
        selectedShapes(1, i) = oPPTShape.Top
        selectedShapes(2, i) = oPPTShape.Left
        selectedShapes(3, i) = oPPTShape.Name
        'oPPTShape.Name = "MegaShape" & i
        i = i + 1
    Next oPPTShape
    Dim bufTop As Integer, bufLeft As Integer, bufName As String, allShapes As String
    'Sort shapes
    Do
        notSorted = False
        For i = 1 To selectedNum - 1
            allShapes = ""
            For k = 1 To selectedNum
                allShapes = allShapes & selectedShapes(3, k) & " "
            Next k
            allShapes = i & ": " & allShapes
            'MsgBox allShapes
            If selectedShapes(1, i) > selectedShapes(1, i + 1) Then
                bufTop = selectedShapes(1, i)
                bufLeft = selectedShapes(2, i)
                bufName = selectedShapes(3, i)
                selectedShapes(1, i) = selectedShapes(1, i + 1)
                selectedShapes(2, i) = selectedShapes(2, i + 1)
                selectedShapes(3, i) = selectedShapes(3, i + 1)
                selectedShapes(1, i + 1) = bufTop
                selectedShapes(2, i + 1) = bufLeft
                selectedShapes(3, i + 1) = bufName
                notSorted = True
            ElseIf selectedShapes(1, i) = selectedShapes(1, i + 1) Then
                If selectedShapes(2, i) > selectedShapes(2, i + 1) Then
                    bufTop = selectedShapes(1, i)
                    bufLeft = selectedShapes(2, i)
                    bufName = selectedShapes(3, i)
                    selectedShapes(1, i) = selectedShapes(1, i + 1)
                    selectedShapes(2, i) = selectedShapes(2, i + 1)
                    selectedShapes(3, i) = selectedShapes(3, i + 1)
                    selectedShapes(1, i + 1) = bufTop
                    selectedShapes(2, i + 1) = bufLeft
                    selectedShapes(3, i + 1) = bufName
                    notSorted = True
                End If
            End If
        Next i
    Loop Until Not notSorted
    For i = 1 To selectedNum
        'MsgBox selectedShapes(3, i) & ": " & selectedShapes(1, i) & ", " & selectedShapes(2, i)
        oPPTApp.ActiveWindow.Selection.ShapeRange(selectedShapes(3, i)).Name = tbNewName.text & "_" & i
    Next i
    
    shapeRangeName = tbNewName.text & "_[x]"
    'For i = 1 To 9
    '    oPPTApp.ActiveWindow.Selection.ShapeRange(tbNewName.text & "_" & i).TextFrame.TextRange.text = i
    'Next i
    DisconnectFromPowerPoint
    Unload FormSelectShapeGroup
    
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

    labelSelectedCells.Caption = "You selected " & sourceCellRange.Count & " cells in Excel. Now you can either:"

End Sub

