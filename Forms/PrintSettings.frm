VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrintSettings 
   Caption         =   "Print Settings"
   ClientHeight    =   6570
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8880.001
   OleObjectBlob   =   "PrintSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PrintSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    PrintSettings.Hide
End Sub

Private Sub btnOK_Click()
    Dim isFirst As Boolean, firstSheetName As String
    isFirst = True
    For i = 0 To lbSheets.ListCount - 1
        If lbSheets.Selected(i) Then
            If isFirst Then
                firstSheetName = lbSheets.List(i)
                isFirst = False
            End If
            ChangePrintSettings lbSheets.List(i), firstSheetName
        End If
    Next
    PrintSettings.Hide
End Sub

Sub ChangePrintSettings(sheetName As String, firstSheetName As String)
    
    With Worksheets(sheetName).PageSetup
        If cbPrintArea.Value Then
            Dim lastCell As String
            lastCell = FindLastCell(sheetName)
            If lastCell <> "" Then .PrintArea = "A1:" & lastCell
        End If
        
        If opPrintOnOne.Value Then
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        ElseIf opPrintOnSeveral.Value Then
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End If
        
        If opMarginsZero.Value Then
            .LeftMargin = 0
            .RightMargin = 0
            .TopMargin = 0
            .BottomMargin = 0
            .HeaderMargin = 0
            .FooterMargin = 0
        ElseIf opMarginsSmall.Value Then
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .HeaderMargin = Application.InchesToPoints(0.5)
            .FooterMargin = Application.InchesToPoints(0.5)
        End If
        
        If opFormatA3.Value Then
            .PaperSize = xlPaperA3
        ElseIf opFormatA4.Value Then
            .PaperSize = xlPaperA4
        End If
        
        If opOrientationAuto.Value Then
            If .PrintArea <> "" Then
                If Range(.PrintArea).Height > Range(.PrintArea).Width Then
                    .Orientation = xlPortrait
                Else
                    .Orientation = xlLandscape
                End If
            End If
        ElseIf opOrientationLandscape.Value Then
            .Orientation = xlLandscape
        ElseIf opOrientationPortrait.Value Then
            .Orientation = xlPortrait
        End If
        
        If opHeaderFooterFirst.Value Then
            If firstSheetName <> sheetName Then
                .LeftHeader = Sheets(firstSheetName).PageSetup.LeftHeader
                .CenterHeader = Sheets(firstSheetName).PageSetup.CenterHeader
                .RightHeader = Sheets(firstSheetName).PageSetup.RightHeader
                .LeftFooter = Sheets(firstSheetName).PageSetup.LeftFooter
                .CenterFooter = Sheets(firstSheetName).PageSetup.CenterFooter
                .RightFooter = Sheets(firstSheetName).PageSetup.RightFooter
            End If
        ElseIf opHeaderFooterNone.Value Then
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
        End If
        
        If opCenterBoth Then
            .CenterVertically = True
            .CenterHorizontally = True
        ElseIf opCenterNone Then
            .CenterVertically = False
            .CenterHorizontally = False
        End If
        
        .PrintComments = xlPrintNoComments
        .BlackAndWhite = False
        
    End With
    
End Sub

Private Sub cbApplyToAll_Click()
    For i = 0 To lbSheets.ListCount - 1
        lbSheets.Selected(i) = cbApplyToAll.Value
    Next
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    lbSheets.List = GetSheetList()
    For i = 0 To lbSheets.ListCount - 1
        If lbSheets.List(i) = ActiveWorkbook.ActiveSheet.Name Then lbSheets.Selected(i) = True
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

End Sub
