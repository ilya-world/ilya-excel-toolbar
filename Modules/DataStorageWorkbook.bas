Attribute VB_Name = "DataStorageWorkbook"
Sub CreateDataWorkbook()
    
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\AccentureToolbarUserData.xlsx"
    ThisWorkbook.Sheets("UserSheets").Copy Before:=Workbooks("AccentureToolbarUserData.xlsx").Sheets(1)
    Dim snakeSheet As Worksheet
    Workbooks("AccentureToolbarUserData.xlsx").Worksheets.Add
    Workbooks("AccentureToolbarUserData.xlsx").Worksheets(1).Name = "SnakeData"
    'Workbooks("AccentureToolbarUserData.xlsx").Worksheets("SnakeData").Visible = xlVeryHidden
    
    ActiveWorkbook.Close savechanges:=True
    
End Sub

Sub OpenDataWorkbook()
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    If Dir(ThisWorkbook.Path & "\AccentureToolbarUserData.xlsx") <> "" Then
        Workbooks.Open Filename:=ThisWorkbook.Path & "\AccentureToolbarUserData.xlsx"
        'Workbooks.Open ThisWorkbook.Path & "\AccentureToolbarUserData.xlsx"
    Else
        CreateDataWorkbook
        Workbooks.Open Filename:=ThisWorkbook.Path & "\AccentureToolbarUserData.xlsx"
    End If
    ThisWorkbook.Activate
End Sub

Sub CloseDataWorkbook()
    If Not Workbooks("AccentureToolbarUserData.xlsx") Is Nothing Then Workbooks("AccentureToolbarUserData.xlsx").Close savechanges:=True
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
End Sub

Sub CopySheetToDataWorkbook(activeBook As Workbook, sheetName As String)
    Dim newName As String, listRow As Integer
    listRow = Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("F2").Value
    newName = "UserSheet" & listRow
    activeBook.Sheets(sheetName).Copy Before:=Workbooks("AccentureToolbarUserData.xlsx").Sheets(1)
    Workbooks("AccentureToolbarUserData.xlsx").Sheets(1).Name = newName
    'Application.DisplayAlerts = False
    For Each rCell In Workbooks("AccentureToolbarUserData.xlsx").Sheets(1).UsedRange
        rCell.Replace "[*]", ""
    Next rCell
    'Application.DisplayAlerts = True
    Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("C" & listRow).Value = newName
    Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("B" & listRow).Value = sheetName
End Sub

Sub CopySheetFromDataWorkbook(activeBook As Workbook, sheetName As String)
    Dim tempName As String
    tempName = WorksheetFunction.index(Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("C:C"), WorksheetFunction.Match(sheetName, Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("B:B"), 0))
    Workbooks("AccentureToolbarUserData.xlsx").Sheets(tempName).Copy After:=activeBook.ActiveSheet
    'Application.DisplayAlerts = False
    For Each rCell In ActiveSheet.UsedRange
        rCell.Replace What:="[*]", Replacement:=""
    Next rCell
    'Application.DisplayAlerts = True
    ActiveSheet.Name = sheetName
End Sub

Sub DeleteSheetFromDataWorkbook(sheetName As String)
    tempRow = WorksheetFunction.Match(sheetName, Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("B:B"), 0)
    tempName = WorksheetFunction.index(Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("C:C"), tempRow)
    Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("B" & tempRow).Clear
    Workbooks("AccentureToolbarUserData.xlsx").Sheets("UserSheets").Range("C" & tempRow).Clear
    'ThisWorkbook.Sheets(tempName).Delete
    'Application.DisplayAlerts = False
    Workbooks("AccentureToolbarUserData.xlsx").Sheets(tempName).Delete
    'Application.DisplayAlerts = True
End Sub

Sub CopyValuesToDataWorkbook(sheetName As String)
    Application.ScreenUpdating = False
    OpenDataWorkbook
    Dim dataBook As Workbook
    Set dataBook = Workbooks("AccentureToolbarUserData.xlsx")
    If Not WorksheetExistsInDataWorkbook(sheetName) Then
        ThisWorkbook.Sheets(sheetName).Copy Before:=dataBook.Sheets(1)
    Else
        ThisWorkbook.Sheets(sheetName).Range("A:F").Copy destination:=dataBook.Sheets(sheetName).Range("A:F")
    End If
    CloseDataWorkbook
    Application.ScreenUpdating = True
End Sub

Sub CopyValuesFromDataWorkbook(sheetName As String)
    Application.ScreenUpdating = False
    OpenDataWorkbook
    Dim dataBook As Workbook
    Set dataBook = Workbooks("AccentureToolbarUserData.xlsx")
    If WorksheetExistsInDataWorkbook(sheetName) Then
        dataBook.Sheets(sheetName).Range("A:F").Copy destination:=ThisWorkbook.Sheets(sheetName).Range("A:F")
    End If
    CloseDataWorkbook
    Application.ScreenUpdating = True
End Sub


Function WorksheetExistsInDataWorkbook(WSName As String) As Boolean
    On Error Resume Next
    WorksheetExistsInDataWorkbook = Workbooks("AccentureToolbarUserData.xlsx").Worksheets(WSName).Name = WSName
    On Error GoTo 0
End Function

