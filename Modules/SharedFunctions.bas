Attribute VB_Name = "SharedFunctions"
Function getSheetRGB(sheetNum As Integer) As String
    Dim c As Long
    Dim r As Long
    Dim G As Long
    Dim B As Long

    c = Application.Sheets(sheetNum).Tab.color
    r = c Mod 256
    G = c \ 256 Mod 256
    B = c \ 65536 Mod 256
    getSheetRGB = "R=" & r & ", G=" & G & ", B=" & B
End Function

Function GetColorFromIndex(index As Integer) As String
    
    Dim strColor As String
    Select Case index
        Case 1
        strColor = "Black"
        iIndexNum = 1
        Case 53
        strColor = "Brown"
        iIndexNum = 53
        Case 52
        strColor = "Olive Green"
        iIndexNum = 52
        Case 51
        strColor = "Dark Green"
        iIndexNum = 51
        Case 49
        strColor = "Dark Teal"
        iIndexNum = 49
        Case 11
        strColor = "Dark Blue"
        iIndexNum = 11
        Case 55
        strColor = "Indigo"
        iIndexNum = 55
        Case 56
        strColor = "Gray-80%"
        iIndexNum = 56
        Case 9
        strColor = "Dark Red"
        iIndexNum = 9
        Case 46
        strColor = "Orange"
        iIndexNum = 46
        Case 12
        strColor = "Dark Yellow"
        iIndexNum = 12
        Case 10
        strColor = "Green"
        iIndexNum = 10
        Case 14
        strColor = "Teal"
        iIndexNum = 14
        Case 5
        strColor = "Blue"
        iIndexNum = 5
        Case 47
        strColor = "Blue-Gray"
        iIndexNum = 47
        Case 16
        strColor = "Gray-50%"
        iIndexNum = 16
        Case 3
        strColor = "Red"
        iIndexNum = 3
        Case 45
        strColor = "Light Orange"
        iIndexNum = 45
        Case 43
        strColor = "Lime"
        iIndexNum = 43
        Case 50
        strColor = "Sea Green"
        iIndexNum = 50
        Case 42
        strColor = "Aqua"
        iIndexNum = 42
        Case 41
        strColor = "Light Blue"
        iIndexNum = 41
        Case 13
        strColor = "Violet"
        iIndexNum = 13
        Case 48
        strColor = "Gray-40%"
        iIndexNum = 48
        Case 7
        strColor = "Pink"
        iIndexNum = 7
        Case 44
        strColor = "Gold"
        iIndexNum = 44
        Case 6
        strColor = "Yellow"
        iIndexNum = 6
        Case 4
        strColor = "Bright Green"
        iIndexNum = 4
        Case 8
        strColor = "Turqoise"
        iIndexNum = 8
        Case 33
        strColor = "Sky Blue"
        iIndexNum = 33
        Case 54
        strColor = "Plum"
        iIndexNum = 54
        Case 15
        strColor = "Gray-25%"
        iIndexNum = 15
        Case 38
        strColor = "Rose"
        iIndexNum = 38
        Case 40
        strColor = "Tan"
        iIndexNum = 40
        Case 36
        strColor = "Light Yellow"
        iIndexNum = 36
        Case 35
        strColor = "Light Green"
        iIndexNum = 35
        Case 34
        strColor = "Light Turqoise"
        iIndexNum = 34
        Case 37
        strColor = "Pale Blue"
        iIndexNum = 37
        Case 39
        strColor = "Lavendar"
        iIndexNum = 39
        Case 2
        strColor = "White"
        iIndexNum = 2
        Case Else
        strColor = "Custom"
    End Select

    GetColorFromIndex = strColor

End Function

Function GetIndexFromColor(color As String) As Integer

    Dim indexResult As Integer
    
    Select Case color
        Case "White": indexResult = 2
        Case "Red": indexResult = 3
        Case "Orange": indexResult = 46
        Case "Yellow": indexResult = 6
        Case "Green": indexResult = 10
        Case "Lime": indexResult = 43
        Case "Blue": indexResult = 5
        Case "Pink": indexResult = 7
        Case "Gold": indexResult = 44
        Case "White": indexResult = 2
        Case "Black": indexResult = 1
        Case Else: indexResult = 0
    End Select
    GetIndexFromColor = indexResult
    
End Function

Sub OpenWebLink(link)
    ActiveWorkbook.FollowHyperlink Address:=link, NewWindow:=True
End Sub

Function FindLastCell(sheetName As String) As String

    Dim LastColumn As Integer
    Dim LastRow As Long
    Dim lastCell As Range

    If WorksheetFunction.CountA(Worksheets(sheetName).Cells) > 0 Then
        'Search for any entry, by searching backwards by Rows.
        LastRow = Worksheets(sheetName).Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        'Search for any entry, by searching backwards by Columns.
        LastColumn = Worksheets(sheetName).Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        FindLastCell = Cells(LastRow, LastColumn).Address
    Else
        FindLastCell = ""
    End If
    
End Function

Function FindLastCellRef(sheetRef As Worksheet) As Range

    Dim LastColumn As Integer
    Dim LastRow As Long
    Dim lastCell As Range

    If WorksheetFunction.CountA(sheetRef.Cells) > 0 Then
    'Search for any entry, by searching backwards by Rows.
    LastRow = sheetRef.Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    'Search for any entry, by searching backwards by Columns.
    LastColumn = sheetRef.Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Set FindLastCellRef = Cells(LastRow, LastColumn)
    Else
        Set FindLastCellRef = Nothing
    End If

End Function
