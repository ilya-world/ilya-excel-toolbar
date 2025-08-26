Attribute VB_Name = "RibbonManagement"
Option Explicit

Dim accentureRibbon As IRibbonUI
Public myTag As String
Public imageTag As String
Public visibleTag As Boolean, updateVisibleTag As Boolean
Public controlNum As Integer, controlType As String

#If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef Source As Any, ByVal length As Long)
#Else
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef Source As Any, ByVal length As Long)
#End If

Sub Ribbon_Onload(ribbon As IRibbonUI)
    
    On Error GoTo SkipWorkbookTest
    If ActiveWorkbook.Name = "Chart in Microsoft PowerPoint" Then
        ThisWorkbook.Close savechanges:=False
    End If
    On Error GoTo 0
SkipWorkbookTest:
    Set accentureRibbon = ribbon
    ThisWorkbook.Sheets("General").Range("B3").Value = ObjPtr(ribbon)
    If ThisWorkbook.Sheets("General").Range("B2").Value <> 1 Then
        UpdateCheck False
    End If
    
End Sub

Sub OnProdRibbonButton(control As IRibbonControl)
    
    Dim sh As Worksheet
    Set sh = ActiveSheet
    If Not sh Is Nothing Then
    
    Select Case control.ID
    ' Color
        Case Is = "btnColorBlue0": SetTheColor 6, 107, 176
        Case Is = "btnColorBlue1": SetTheColor 76, 155, 220
        Case Is = "btnColorBlue2": SetTheColor 118, 178, 228
        Case Is = "btnColorBlue3": SetTheColor 165, 204, 237
        Case Is = "btnColorGrey1": SetTheColor 240, 240, 240
        Case Is = "btnColorGrey2": SetTheColor 220, 220, 220
        Case Is = "btnColorGrey3": SetTheColor 191, 191, 191
        Case Is = "btnColorGrey4": SetTheColor 147, 147, 147
        Case Is = "btnColorGrey5": SetTheColor 105, 105, 105
        Case Is = "btnColorGrey6": SetTheColor 60, 60, 60
        Case Is = "btnColorGreen3": SetTheColor 218, 240, 168
        Case Is = "btnColorGreen2": SetTheColor 175, 224, 110
        Case Is = "btnColorGreen1": SetTheColor 125, 185, 53
        Case Is = "btnColorGreen0": SetTheColor 96, 139, 45
        Case Is = "btnColorOrange3": SetTheColor 243, 207, 116
        Case Is = "btnColorOrange2": SetTheColor 239, 182, 67
        Case Is = "btnColorOrange1": SetTheColor 241, 137, 23
        Case Is = "btnColorOrange0": SetTheColor 210, 99, 8
        Case Is = "btnColorWhite": SetTheColor 255, 255, 255
        Case Is = "btnColorBlack": SetTheColor 0, 0, 0
        Case Is = "btnColorPurple1": SetTheColor 186, 156, 197
        Case Is = "btnColorSignalRed": SetTheColor 255, 77, 62
        Case Is = "btnColorNoFill": nofill
        Case Is = "btnColorYellowInput": SetTheColor 255, 255, 153
        Case Is = "btnColorTLRed": SetTheColor 255, 154, 5
        Case Is = "btnColorTLYellow": SetTheColor 255, 221, 0
        Case Is = "btnColorTLGreen": SetTheColor 136, 221, 0
        
        Case Is = "btnFontColorRed": SetTextColor 255, 0, 0
        Case Is = "btnFontColorBlack": SetTextColor 0, 0, 0
        Case Is = "btnFontColorWhite": SetTextColor 255, 255, 255
    
    ' Functions
        Case Is = "btnGoToCorner": GoToCorner
        Case Is = "btnAutoZoom": SwitchZoom
        Case Is = "btnAutoFillDown": AutoFillDown
        Case Is = "btnAutoFillRight": AutoFillRight
        Case Is = "btnSheetManager": ShowSheetManager
        Case Is = "btnPrintSettings": OpenPrintSettings
        Case Is = "btnExtractLinks": ExtractLinks
        Case Is = "btnProtect": ProtectSheets
        Case Is = "btnUnprotect": UnlockDocument
        Case Is = "btnCopyFormula": CopyFormula
        Case Is = "btnCopyToSheets": CopyToSheets
        Case Is = "btnHideInterface": HideInterface
        Case Is = "btnConcatenate": AutoConcatenate
        Case Is = "btnBorderHorizontal": BorderHorizontal
        Case Is = "btnBorderVertical": BorderVertical
        Case Is = "btnSecret": OpenSnake
        Case Is = "btnTestFunction": TestFunction
        Case Is = "btnCalculationTime": FormTimeTest.Show
        Case Is = "btnValuesWorkbook": CreateValueFixedWorkbookCopy
        Case Is = "btnZoomToWidth": ZoomToWidth
        Case Is = "btnSwapCells": SwapCells
        Case Is = "btnHideUnused": HideUnusedColumnsAndRowsInSelected
        Case Is = "btnDisplaySheetsWindow": DisplaySheetsWindow
        Case Is = "btnFillRandom": FormRandomGenerator.Show
    'Case
        Case Is = "btnCaseSentence": CaseSentence
        Case Is = "btnCaseLower": CaseLower
        Case Is = "btnCaseUpper": CaseUpper
        Case Is = "btnCaseEachWord": CaseCapitalize
        Case Is = "btnCaseToogle": CaseToogle
        
    'Save/send selection/sheets
        Case Is = "btnSendRange": SendRange
        Case Is = "btnSaveRange": SaveRange
        Case Is = "btnSaveSheets": SaveSheets
        Case Is = "btnSendSheets": SendSheets
        Case Is = "btnBodyRange": BodyRange
        Case Is = "btnFormatTable": CustomFormatTable
        
    'PPT Updater
        Case Is = "btnUpdatePPT": UpdatePPTShapes
        Case Is = "btnSelectShapePart": SelectPPTShapePart
        Case Is = "btnSelectShapeGroup": SelectShapeGroup
        Case Is = "btnSelectTableCell": SelectTableCell
        Case Is = "btnSelectPPTChart": SelectPPTChart
        Case Is = "btnSelectThinkCell": SelectThinkCell
        Case Is = "btnReselectPPTFile": ReselectPPTFile
        
    'Custom number formatting
        Case Is = "btnFormatSettings": FormCustomFormats.Show
        Case Is = "btnFormatApply1": ApplyCustomFormat 1
        Case Is = "btnFormatApply2": ApplyCustomFormat 2
        Case Is = "btnFormatApply3": ApplyCustomFormat 3
        Case Is = "btnFormatApply4": ApplyCustomFormat 4
        Case Is = "btnFormatApply5": ApplyCustomFormat 5
        Case Is = "btnFormatApply6": ApplyCustomFormat 6
        
    ' Help & Support
        Case Is = "btnToolbarSettings": FormCustomFormats.Show
        Case Is = "btnFeedback": OpenWebLink "http://accentureexceltoolbar.idea.informer.com/"
        Case Is = "btnCoP": OpenWebLink "http://kxsites.accenture.com/groups/ExcelCOP"
        Case Is = "btnExcelIntroduction": OpenWebLink "https://support.accenture.com/technology/mycomputer/Documents/Excel%202010%20An%20Introcduction%20Workshop%20Presentation/Excel%202010%20-%20An%20Introduction.pptx"
        Case Is = "btnAddInGuide": OpenWebLink "http://legalov.ru/accenture/GuideToAccentureExcelToolbar.pptx"
        Case Is = "btnAbout": About.Show
        Case Is = "btnShare": FormShare.Show
        Case Is = "btnUpdateCheck": UpdateCheck True
        Case Is = "btnUpdateAvailable": UpdateCheck True
        'Case Is = "btnUpdateCheck": CopySheetToDataWorkbook
        
    ' Manual macros
        Case Is = "btnMacroEdit": ShowMacroEdit
        Case Is = "btnMacroRun": RunMacro
    
    ' Inserts
        Case Is = "btnInsertTitle": InsertSheet "Title"
        Case Is = "btnInsertToC": InsertSheet "ToC"
        Case Is = "btnInsertSection": InsertSheet "Section"
        Case Is = "btnInsertManagement": InsertSheet "Version"
        Case Is = "btnAddUserSheet": ShowAddSheet
        Case Is = "btnInsertUserSheet": ShowInsertSheet
    
    'Insert symbol
        Case Is = "btnInsertSymbolHB0": InsertSymbol control.ID
        Case Is = "btnInsertSymbolHB1": InsertSymbol control.ID
        Case Is = "btnInsertSymbolHB2": InsertSymbol control.ID
        Case Is = "btnInsertSymbolHB3": InsertSymbol control.ID
        Case Is = "btnInsertSymbolHB4": InsertSymbol control.ID
        
        Case Is = "btnInsertSymbolEuro": InsertSymbol control.ID
        Case Is = "btnInsertSymbolPound": InsertSymbol control.ID
        Case Is = "btnInsertSymbolRuble": InsertSymbol control.ID
        Case Is = "btnInsertSymbolRupee": InsertSymbol control.ID
        Case Is = "btnInsertSymbolYen": InsertSymbol control.ID
        
        Case Is = "btnInsertSymbolArrowUp": InsertSymbol control.ID
        Case Is = "btnInsertSymbolArrowDown": InsertSymbol control.ID
        Case Is = "btnInsertSymbolArrowLeft": InsertSymbol control.ID
        Case Is = "btnInsertSymbolArrowRight": InsertSymbol control.ID
        Case Is = "btnInsertSymbolArrowIncrease": InsertSymbol control.ID
        Case Is = "btnInsertSymbolArrowDecrease": InsertSymbol control.ID
        
        Case Is = "btnInsertSymbolTick": InsertSymbol control.ID
        Case Is = "btnInsertSymbolCross": InsertSymbol control.ID
        
        Case Is = "btnInsertSymbolHappy": InsertSymbol control.ID
        Case Is = "btnInsertSymbolSad": InsertSymbol control.ID
        
        Case Is = "btnInsertSymbolPlusMinus": InsertSymbol control.ID
        Case Is = "btnInsertSymbolDivision": InsertSymbol control.ID
        Case Is = "btnInsertSymbolMultiplication": InsertSymbol control.ID
        
        Case Is = "btnInsertSymbolUser1": InsertSymbol control.ID
        Case Is = "btnInsertSymbolUser2": InsertSymbol control.ID
        Case Is = "btnInsertSymbolUser3": InsertSymbol control.ID
        Case Is = "btnInsertSymbolUser4": InsertSymbol control.ID
        
        Case Else
            MsgBox "Button """ & control.ID & """ has no events", vbInformation
    End Select
    
    Else
    MsgBox "You need to open a workbook to use the toolbar", vbOKOnly + vbCritical, "Error"
    End If
   
End Sub

Sub RefreshRibbon(Tag As String)
    myTag = Tag
    If accentureRibbon Is Nothing Then
        Set accentureRibbon = GetRibbon(ThisWorkbook.Sheets("General").Range("B3").Value)
        accentureRibbon.Invalidate
    Else
        accentureRibbon.Invalidate
    End If
End Sub

Sub GetScreentip(control As IRibbonControl, ByRef setScreentip)
    
    controlNum = 0
    controlType = ""
    If Len(control.ID) > 0 Then
        controlType = Left(control.ID, Len(control.ID) - 1)
    End If
    If controlType = "btnFormatApply" Then
        controlNum = Val(Right(control.ID, 1))
        setScreentip = ThisWorkbook.Worksheets("Formats").Cells(controlNum + 1, 2)
    End If
    'Select Case control.ID
    'Case "btnFormatApply1": setScreentip = "123"
    'End Select
    'Call RefreshRibbon(control.ID)
End Sub

Sub GetLabel(control As IRibbonControl, ByRef setLabel)
    
    Dim unicodeNum As Long, symbolDescription As String
    unicodeNum = GetUnicodeForSymbol(control.ID)
    If unicodeNum > 0 Then
        Select Case control.ID
            Case "btnInsertSymbolUser1": symbolDescription = ThisWorkbook.Sheets("Symbols").Cells(2, 2).Value
            Case "btnInsertSymbolUser2": symbolDescription = ThisWorkbook.Sheets("Symbols").Cells(3, 2).Value
            Case "btnInsertSymbolUser3": symbolDescription = ThisWorkbook.Sheets("Symbols").Cells(4, 2).Value
            Case "btnInsertSymbolUser4": symbolDescription = ThisWorkbook.Sheets("Symbols").Cells(5, 2).Value
            Case Else: symbolDescription = ""
        End Select
        setLabel = WorksheetFunction.Unichar(unicodeNum) & " " & symbolDescription
    Else
        setLabel = "Unidentified"
    End If
    Call RefreshRibbon(control.ID)

End Sub

Sub GetSupertip(control As IRibbonControl, ByRef setSupertip)
    
    controlNum = 0
    controlType = ""
    If Len(control.ID) > 0 Then
        controlType = Left(control.ID, Len(control.ID) - 1)
    End If
    If controlType = "btnFormatApply" Then
        controlNum = Val(Right(control.ID, 1))
        setSupertip = ThisWorkbook.Worksheets("Formats").Cells(controlNum + 1, 3)
    End If
    'Select Case control.ID
    'Case "btnFormatApply1": setSupertip = "MegaTest 2"
    'End Select
    'Call RefreshRibbon(control.ID)
End Sub

Sub GetVisible(control As IRibbonControl, ByRef visible)
    Select Case control.ID
    Case "btnUpdateAvailable": visible = updateVisibleTag
    Case "btnTestFunction": visible = ThisWorkbook.Worksheets("General").Range("J2").Value
    End Select
    'Call RefreshRibbon(control.ID)
End Sub

Sub GetImage(control As IRibbonControl, ByRef setImage)
    
    controlNum = 0
    controlType = ""
    If Len(control.ID) > 0 Then
        controlType = Left(control.ID, Len(control.ID) - 1)
    End If
    If controlType = "btnFormatApply" Then
        controlNum = Val(Right(control.ID, 1))
        setImage = ThisWorkbook.Worksheets("Formats").Cells(controlNum + 1, 4)
    End If
    'imageTag = ""
    'Select Case control.ID
    'Case "btnFormatApply1": setImage = "_6"
    'End Select
    'Call RefreshRibbon(control.ID)
    'If imageTag <> "" Then
    '    setImage = imageTag
    '    Call RefreshRibbon(control.ID)
    'End If
End Sub

#If VBA7 Then
Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If
        Dim objRibbon As Object
        CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
        Set GetRibbon = objRibbon
        Set objRibbon = Nothing
End Function




