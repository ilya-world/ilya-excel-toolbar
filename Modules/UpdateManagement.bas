Attribute VB_Name = "UpdateManagement"
Sub UpdateCheck(manualCheck As Boolean)
    
    Dim text As String, version As Integer, rand As Integer
    
'    rand = WorksheetFunction.RandBetween(1, 1000)
'    text = DownloadTextFile("http://legalov.ru/accenture/version.txt?a=" & rand) 'to avoid cache
'    If text <> "" Then
'        version = CDec(text)
'        'version = ThisWorkbook.Sheets("General").Range("J1").Value
'        If version > ThisWorkbook.Sheets("General").Range("B1").Value Then
'            ThisWorkbook.Sheets("General").Range("B2").Value = 1
'            updateVisibleTag = True
'            Call RefreshRibbon(Tag:="btnUpdateAvailable")
'            If manualCheck Then
'                FormUpdate.Show
'            End If
'            'check = MsgBox("A new version of the Accenture Excel Toolbar is available. Would you like to download it?", vbYesNo, "Update available")
'            '        If check = vbNo Then
'            '            Exit Sub
'            '        Else
'            '            FormUpdate.Show
'            '            OpenWebLink "http://legalov.ru/accenture/toolbar_last_version.zip"
'            '            'profile = Environ("userprofile") & "\AppData\Roaming\Microsoft\AddIns\"
'            '            'Workbooks.Add (profile & "ExcelToolbarUpdater.xlsm")
'            '            'objExcel.Application.Run "'" & profile & "ExcelToolbarUpdater.xlsm!Update"
'            '        End If
'
'        Else
'            If manualCheck Then
'                MsgBox "You are using the last version!", vbOKOnly, "Updater"
'            End If
'        End If
'    Else
'        If manualCheck Then
'            MsgBox "Can't connect to the server to check version", vbOKOnly, "Connection error"
'        End If
'    End If
    
End Sub

Public Function DownloadTextFile(url As String) As String
    On Error GoTo Errhandl
    Dim oHTTP As WinHttp.WinHttpRequest
    Set oHTTP = New WinHttp.WinHttpRequest
    oHTTP.SetTimeouts 3000, 3000, 3000, 3000
    oHTTP.Open Method:="GET", url:=url, async:=False
    oHTTP.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    oHTTP.SetRequestHeader "Content-Type", "multipart/form-data; "
    oHTTP.Option(WinHttpRequestOption_EnableRedirects) = True
    oHTTP.Send

    Dim success As Boolean
    success = oHTTP.WaitForResponse()
    If Not success Then
        Debug.Print "DOWNLOAD FAILED!"
        Exit Function
    End If

    Dim responseText As String
    responseText = oHTTP.responseText

    Set oHTTP = Nothing

    DownloadTextFile = responseText
    Exit Function
Errhandl:
    DownloadTextFile = ""
End Function
