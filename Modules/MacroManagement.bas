Attribute VB_Name = "MacroManagement"
Sub ShowMacroEdit()

    MacroEdit.Show

End Sub

Sub RunMacro()

    Dim LineNum As Long, profile As String, macroWorkbook As Excel.Workbook
    'profile = Environ("userprofile") & "\AppData\Roaming\Microsoft\AddIns\"
    profile = Environ("userprofile") & "\Documents\"
     'create a new module in the current workbook, enter the code, run and remove the new module
     
    Set macroWorkbook = Workbooks.Open(profile & "MacroRun.xlsm")
        
    Set VBComp = macroWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    VBComp.Name = "NewModule"
     'add the code lines
    Set VBCodeMod = macroWorkbook.VBProject.VBComponents("NewModule").CodeModule
    With VBCodeMod
        LineNum = .CountOfLines + 1
        .InsertLines LineNum, _
        "Sub MyNewProcedure()" & Chr(13) & MacroEdit.tbMacroCode.Value & Chr(13) & "End Sub"
    End With
     
     'run the new module
    Application.Run "MyNewProcedure"
     'remove the new module
    ThisWorkbook.VBProject.VBComponents.Remove VBComp

End Sub
