Attribute VB_Name = "SheetManagement"
Sub ShowSheetManager()
    SheetManager.Show
End Sub

Sub InsertSheet(sheetName As String)

    If WorksheetExists(sheetName) Then
        MsgBox "Sheet '" & sheetName & "' already exists in the workbook", vbOKOnly, "Error"
        Exit Sub
    End If
    Select Case sheetName
        Case Is = "Title": ThisWorkbook.Sheets(sheetName).Copy Before:=ActiveWorkbook.Sheets(1)
        Case Is = "Version"
            If WorksheetExists("Title") Then
                ThisWorkbook.Sheets(sheetName).Copy After:=ActiveWorkbook.Sheets("Title")
            Else
                ThisWorkbook.Sheets(sheetName).Copy Before:=ActiveWorkbook.Sheets(1)
            End If
        Case Is = "ToC"
            If WorksheetExists("Version") Then
                ThisWorkbook.Sheets(sheetName).Copy After:=ActiveWorkbook.Sheets("Version")
            Else
                If WorksheetExists("Title") Then
                    ThisWorkbook.Sheets(sheetName).Copy After:=ActiveWorkbook.Sheets("Title")
                Else
                    ThisWorkbook.Sheets(sheetName).Copy Before:=ActiveWorkbook.Sheets(1)
                End If
            End If
            Dim rowNum As Integer
            rowNum = 3
            For Each sht In ActiveWorkbook.Sheets
                If sht.visible And sht.index > ActiveWorkbook.Sheets("Toc").index Then
                    ActiveWorkbook.Sheets("Toc").Cells(rowNum, 2).Value = sht.Name
                    ActiveWorkbook.Sheets("Toc").Hyperlinks.Add Cells(rowNum, 2), "", "'" & sht.Name & "'!A1"
                    ActiveWorkbook.Sheets("Toc").Cells(rowNum, 2).Font.Name = "Arial"
                    ActiveWorkbook.Sheets("Toc").Cells(rowNum, 2).Font.Size = 14
                    rowNum = rowNum + 1
                End If
            Next sht
        Case Else
            ThisWorkbook.Sheets(sheetName).Copy After:=ActiveSheet
    End Select
    
End Sub

Function GetSheetList() As Variant
    
    Dim allSheets() As Variant
    ReDim allSheets(1 To Application.Sheets.Count)
    For i = 1 To Application.Sheets.Count
        allSheets(i) = Application.Sheets(i).Name
    Next
    GetSheetList = allSheets
    
End Function

Sub ShowAddSheet()
    AddSheetsToTemplates.Show
End Sub

Sub ShowInsertSheet()
    InsertUserSheet.Show
End Sub

Sub CreateValueFixedWorkbookCopy()

    Dim newBook As Workbook, activeBook As Workbook, newName As String, oldFile As String
    Set activeBook = ActiveWorkbook
    oldFile = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    newName = Application.GetSaveAsFilename
    If newName = "False" Then Exit Sub
    newName = newName & "xlsx"
    Application.DisplayAlerts = False
    ActiveWorkbook.CheckCompatibility = False
    ActiveWorkbook.SaveAs Filename:=newName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True
    For Each sh In ActiveWorkbook.Worksheets
        If sh.visible = True Then
            sh.Activate
            sh.Cells.Copy
            sh.Range("A1").PasteSpecial Paste:=xlValues
            sh.Range("A1").Select
        End If
    Next sh
    Application.CutCopyMode = False
    Workbooks.Open oldFile
    MsgBox "Workbook without formulas was saved and opened", vbOKOnly, "Save complete"
    
End Sub

Sub HideUnusedColumnsAndRowsInSelected()
    
    Dim sht As Worksheet
    If ActiveSheet.Cells(ActiveSheet.Rows.Count, ActiveSheet.Columns.Count).EntireColumn.Hidden Then
        For Each sht In ActiveWindow.SelectedSheets
            ShowUnusedColumnsAndRowsOnSheet sht
        Next sht
    Else
        For Each sht In ActiveWindow.SelectedSheets
            HideUnusedColumnsAndRowsOnSheet sht
        Next sht
    End If
    
End Sub

Sub HideUnusedColumnsAndRowsOnSheet(sheetRef As Worksheet)

    Dim rangeToHide As Range, lastCell As Range
    Set lastCell = FindLastCellRef(sheetRef) 'sheetRef.Range(FindLastCell(sheetRef.Name))
    Set rangeToHide = lastCell.Offset(1, 1).Resize(sheetRef.Rows.Count - lastCell.Row, sheetRef.Columns.Count - lastCell.Column)
    rangeToHide.Columns.Hidden = True
    rangeToHide.Rows.Hidden = True
    
End Sub

Sub ShowUnusedColumnsAndRowsOnSheet(sheetRef As Worksheet)

    Dim rangeToShow As Range, lastCell As Range
    Set lastCell = FindLastCellRef(sheetRef) 'sheetRef.Range(FindLastCell(sheetRef.Name))
    Set rangeToShow = lastCell.Offset(1, 1).Resize(sheetRef.Rows.Count - lastCell.Row, sheetRef.Columns.Count - lastCell.Column)
    rangeToShow.Columns.Hidden = False
    rangeToShow.Rows.Hidden = False

End Sub

Sub DisplaySheetsWindow()
    If ActiveWorkbook.Sheets.Count > 16 Then
        Application.CommandBars("Workbook Tabs").Controls("More Sheets...").Execute
    Else
        Application.CommandBars("Workbook Tabs").ShowPopup
    End If
End Sub


