Attribute VB_Name = "MailManagement"
Sub SendRange()
    SendSaveAction "Range", "Send"
End Sub

Sub SaveRange()
    SendSaveAction "Range", "Save"
End Sub

Sub SendSheets()
    SendSaveAction "Sheets", "Send"
End Sub

Sub SaveSheets()
    SendSaveAction "Sheets", "Save"
End Sub

Sub BodyRange()
    SendSaveAction "Range", "Body"
End Sub

Sub SendSaveAction(objectType As String, actionType As String)
'Some credits: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    Dim Source As Range
    Dim Dest As Workbook
    Dim wb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim currentSheet As Worksheet
    Dim rangeToHide As Range
    Dim forCell As Range, insertedCells As Range
    Dim newName As String
    Dim sheetsArray As Variant
    
    Set currentSheet = ActiveSheet
    Set Source = Nothing
    
    If objectType = "Range" Then
        On Error Resume Next
        Set Source = Selection.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        If Source Is Nothing Then
            MsgBox "The source is not a range or the sheet is protected, please correct and try again.", vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
    
        If ActiveWindow.SelectedSheets.Count > 1 Or _
           Selection.Cells.Count = 1 Or _
           Selection.Areas.Count > 1 Then
            MsgBox "An Error occurred :" & vbNewLine & vbNewLine & _
                   "You have more than one sheet selected." & vbNewLine & _
                   "You only selected one cell." & vbNewLine & _
                   "You selected more than one area." & vbNewLine & vbNewLine & _
                   "Please correct and try again.", vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If
    
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With
        If actionType <> "Body" Then
            Set wb = ActiveWorkbook
            Set Dest = Workbooks.Add(xlWBATWorksheet)
            
            Source.Copy
            With Dest.Sheets(1)
                .Name = currentSheet.Name
                .Cells(1).PasteSpecial Paste:=8 'width
                .Cells(1).PasteSpecial Paste:=xlPasteAll
                '.Cells(1).PasteSpecial Paste:=xlPasteValues
                '.Cells(1).PasteSpecial Paste:=xlPasteFormats
                Set insertedCells = Selection
                Set rangeToHide = insertedCells.Cells(insertedCells.Rows.Count + 1, insertedCells.Columns.Count + 1).Resize(.Rows.Count - insertedCells.Rows.Count, .Columns.Count - insertedCells.Columns.Count)
                rangeToHide.Columns.Hidden = True
                rangeToHide.Rows.Hidden = True
                .Cells(1).Select
                Application.CutCopyMode = False
            End With
            RemoveExternalLinksFromCells insertedCells, Source
        End If
    End If
    
    If objectType = "Sheets" Then
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With
        Set wb = ActiveWorkbook
        ReDim sheetsArray(1 To ActiveWindow.SelectedSheets.Count)
        For i = 1 To ActiveWindow.SelectedSheets.Count
            sheetsArray(i) = ActiveWindow.SelectedSheets(i).Name
        Next
        wb.Sheets(sheetsArray).Copy ' After:=Dest.Sheets(Dest.Sheets.Count)
        Set Dest = ActiveWorkbook
        For i = 1 To Dest.Sheets.Count
            Set insertedCells = Dest.Sheets(i).Range(Cells(1, 1).Address, FindLastCellRef(Dest.Sheets(i)).Address)
            Set Source = wb.Sheets(Dest.Sheets(i).Name).Range(Cells(1, 1).Address, FindLastCellRef(wb.Sheets(Dest.Sheets(i).Name)).Address)
            'MsgBox insertedCells.Count - Source.Count
            RemoveExternalLinksFromCells insertedCells, Source
        Next
        
    End If
    
    If actionType <> "Body" Then
        TempFilePath = Environ$("temp") & "\"
        TempFileName = wb.Name & " - " & currentSheet.Name
        'TempFileName = "Selection of " & wb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")

        If Val(Application.version) < 12 Then
            'You use Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'You use Excel 2007-2016
            FileExtStr = ".xlsx": FileFormatNum = 51
        End If
    End If
    If actionType = "Save" Then
        newName = Application.GetSaveAsFilename(wb.Name)
        If newName = "False" Then
            Dest.Close False
            With Application
                .ScreenUpdating = True
                .EnableEvents = True
            End With
            Exit Sub
        End If
        newName = newName & "xlsx"
        Dest.CheckCompatibility = False
        Dest.SaveAs Filename:=newName, FileFormat:=FileFormatNum, CreateBackup:=False
    End If
    If actionType = "Send" Then
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
    
        With Dest
            .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
            On Error Resume Next
            With OutMail
                .To = ""
                .CC = ""
                .BCC = ""
                .Subject = wb.Name & " - " & currentSheet.Name
                .Body = ""
                .Attachments.Add Dest.FullName
                .Display
            End With
            On Error GoTo 0
            .Close savechanges:=False
        End With
    
        Kill TempFilePath & TempFileName & FileExtStr
    
        Set OutMail = Nothing
        Set OutApp = Nothing
    End If
    
    If actionType = "Body" Then
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        On Error Resume Next
        With OutMail
            .To = ""
            .CC = ""
            .BCC = ""
            .Subject = wb.Name & " - " & currentSheet.Name
            .HTMLBody = RangetoHTML(Source)
            .Display
        End With
        On Error GoTo 0
        Set OutMail = Nothing
        Set OutApp = Nothing
    End If
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

Sub RemoveExternalLinksFromCells(insertedCells As Range, Source As Range)
    Dim forCell As Range
    For Each forCell In insertedCells
        'MsgBox "Cell: " & forCell.Row & " " & forCell.Column & vbNewLine & IsError(Evaluate(insertedCells.Cells(forCell.Row, forCell.Column).Formula)) & vbNewLine & insertedCells.Cells(forCell.Row, forCell.Column).Formula
        If IsError(Evaluate(insertedCells.Cells(forCell.Row, forCell.Column).Formula)) Then
            forCell.Value = Source.Cells(forCell.Row, forCell.Column).Value
        End If
        If insertedCells.Cells(forCell.Row, forCell.Column).Value <> Source.Cells(forCell.Row, forCell.Column).Value Then
            forCell.Value = Source.Cells(forCell.Row, forCell.Column).Value
        End If
        If CBool(InStr(1, forCell.Formula, "[")) Then
            forCell.Value = forCell.Value
        End If
    Next forCell
End Sub

Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")
'Close TempWB
    TempWB.Close savechanges:=False
'Delete the htm file we used in this function
    Kill TempFile
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function


