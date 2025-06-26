Attribute VB_Name = "Func_Subs"

' add your procedures here
Public tbHelperSettings As clsSettings
Public tbHelperClass As clstBHelper
Public fso As FileSystemObject
Public chgLogs As New colChangeLogItems
Public githubReleasesURL As String = "https://github.com/twinbasic/twinbasic/releases"

Public Function askForFolder(dlgCaption As String, Optional currentFolder As String) As String
        
    ' open the select folder form and return the selection (if any)
    Dim frmSelFolder As New frmSelectFolder
    
    With frmSelFolder
        If fso.FolderExists(currentFolder) Then .selectedFolder = currentFolder
        .Caption = dlgCaption
        .Show(vbModal)
    End With
    
    Return frmSelFolder.selectedFolder
        
End Function

Public Function ChangeLogItemColor(changeLogType As String) As Long
    Dim itemColor As Long
    
    Select Case UCase(Trim(changeLogType))
        Case "IMPORTANT"
            itemColor = vbBlue
        Case "KNOWN ISSUE"
            itemColor = vbBlack
        Case "TIP"
            itemColor = RGB(22, 83, 126) ' blueish
        Case "WARNING"
            itemColor = RGB(153, 0, 0)   ' dark red
        Case "FIXED"
            itemColor = RGB(56, 118, 29) ' green
        Case "ADDED"
            itemColor = RGB(75, 0, 130)  ' indigo
        Case "UPDATED"
            itemColor = RGB(0, 128, 0)   ' other green 
        Case Else
            itemColor = vbBlack
    End Select
    
    Return itemColor
    
End Function

Public Function GetCurrentTBVersion(tBFolder As String) As String
        
    ' attempt to find the version number of twinBasic in use
    Dim fileWithVersionInfo As String = tBFolder & "ide\build.js"
    Dim versionIndicator As String = "BETA"
    Dim fileContents As String
    Dim tempString As String
    
    If Not fso.FileExists(fileWithVersionInfo) Then
        GetCurrentTBVersion = "Not found"
        Exit Function
    End If
        
    ' open the file designated as the one with the version number
    fileContents = fsoFileRead(fileWithVersionInfo)
    
    ' parse the text for the version number
    tempString = Mid(fileContents, InStr(fileContents, versionIndicator))
    GetCurrentTBVersion = Mid(tempString, Len(versionIndicator) + 1, 4)
    
    tbHelperClass.InstalledtBVersion = GetCurrentTBVersion
        
End Function

Public Function FillLogHistoryGrid(Optional ViewDate As String = "") As Boolean
    
    ' open the log text file and display it in the flexgrid
    Dim logContents As New colHistoryLogItems
    Dim logItem As clsHistoryLogItem
    Dim itemColor As Long
    Dim colNum As Integer
    
    If Not logContents.LoadLog Then
        MsgBox("There was an issue reading the log file", vbExclamation, "View log")
        FillLogHistoryGrid = False
        Exit Function
    End If
    
    Form1.ShowStatusMessage("Display history log" & IIf(Len(ViewDate) = 0, "", " for " & ViewDate))
    
    frmViewLog.flgLog.Rows = 1
    For Each logItem In logContents
        With frmViewLog.flgLog
            If Len(ViewDate) = 0 Or logItem.LogDate = ViewDate Or ViewDate = "Show All" Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            
                If logItem.LogCLI.tBVersion = 0 Then
                    .TextMatrix(.Row, 0) = logItem.LogDateTime
                    .TextMatrix(.Row, 3) = logItem.LogMessage
                Else
                    .TextMatrix(.Row, 0) = logItem.LogDateTime
                    .TextMatrix(.Row, 1) = logItem.LogCLI.tBVersion
                    .TextMatrix(.Row, 2) = logItem.LogCLI.Type
                    .TextMatrix(.Row, 3) = logItem.LogCLI.Notes
                End If
                
                ' color the change log of the tB version that was installed
                ' during the logging process
                itemColor = ChangeLogItemColor(logItem.LogCLI.Type)
                For colNum = 0 To .Cols - 1
                    .Col = colNum
                    .CellForeColor = itemColor
                Next
            End If
            
        End With
        DoEvents()
        
    Next
    
    ' tell the grid to resize the row for the wordwrap function on 4th column
    frmViewLog.flgLog.AutoSize(3, frmViewLog.flgLog.Rows - 1, FlexAutoSizeModeRowHeight)
    
    If Len(ViewDate) = 0 Then
        ' fill the dropdown with the unique dates from the log file
        ' just during the first time through this
        Dim logDate As String
        
        With frmViewLog.cboLogDate
            .Clear()
            .AddItem("Show All")
            For Each logDate In logContents.HistoryLogDates
                .AddItem(logDate)
            Next
            .ListIndex = 0
        End With
    End If
    
    FillLogHistoryGrid = True
End Function

Public Sub FilltBChangeLog(Optional fortBVersion As String = "")
    
    ' open the log text file and display it in the flexgrid
    Dim clItem As clsChangeLogItem
    Dim itemColor As Long
    Dim colNum As Integer
    
    If chgLogs.tBVersionGap > 1 Then
        Dim gridTitleCaption As String
        
        ' make the caption seem more correct
        If chgLogs.tBVersionGap = 2 Then
            gridTitleCaption = CStr(tbHelperClass.InstalledtBVersion + 1) & " and " & chgLogs.LatestVersion
        Else
            gridTitleCaption = CStr(tbHelperClass.InstalledtBVersion + 1) & " thru " & chgLogs.LatestVersion
        End If
        
        Form1.lblChangeLogTitle.Caption = "Change logs for " & gridTitleCaption
    ElseIf chgLogs.tBVersionGap = 1 Then
        Form1.lblChangeLogTitle.Caption = "Change Log for " & chgLogs.LatestVersion
    Else
        Form1.lblChangeLogTitle.Caption = "Change Log"
    End If
    
    With Form1.flgLog
        .Rows = 1
        For Each clItem In chgLogs
            
                If Len(fortBVersion) = 0 Or clItem.tBVersion = Val(fortBVersion) Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                
                    .TextMatrix(.Row, 0) = clItem.tBVersion
                    .TextMatrix(.Row, 1) = clItem.Type
                    .TextMatrix(.Row, 2) = clItem.Notes
                    
                    ' color the change log of the tB version that was installed
                    ' during the logging process
                    itemColor = ChangeLogItemColor(clItem.Type)
                    For colNum = 0 To .Cols - 1
                        .Col = colNum
                        .CellForeColor = itemColor
                    Next
                End If
            
            DoEvents()
        Next
        
        ' tell the grid to resize the row for the wordwrap function on 3nd column
        .AutoSize(2, .Rows - 1, FlexAutoSizeModeRowHeight)
        
    End With
End Sub

Public Function fsoFileRead(filePath As String) As String
    
    If Not fso.FileExists(filePath) Then Return "Failed fsoFileRead"
    
    On Error GoTo readError
    
    Dim fso As New Scripting.FileSystemObject
        Dim fileToRead As TextStream
        
        Set fileToRead = fso.OpenTextFile(filePath, ForReading)
            fsoFileRead = fileToRead.ReadAll()
readError:
        If fsoFileRead = vbNullString Then
            MsgBox("Unable to read " & filePath, vbExclamation, "FileRead")
        End If
        fileToRead.Close()
    Set fso = Nothing
    
End Function

Private Function GettBParentFolder() As String
        
    Dim idx As Integer
    Dim slashCount As Integer
        
    ' loop backwards until the second \ is found - which will indicate where
    ' the parent folder for twinBASIC is
    For idx = Len(tbHelperSettings.twinBASICFolder) To 1 Step -1
        If Mid(tbHelperSettings.twinBASICFolder, idx, 1) = "\" Then slashCount += 1
        If slashCount = 2 Then Exit For
    Next
        
    ' truncate the value in the textbox holding the install folder, to get the parent folder
    GettBParentFolder = Left(tbHelperSettings.twinBASICFolder, idx)
        
End Function

Public Sub InstallTwinBasic(zipLocation As String)
        
    ' go through the steps of deleting the current files and unziping the new files
    ' to the folder that has been desgniated
        
    ' delete current files & recreate the folder
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim RetVal As Long
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = tbHelperSettings.twinBASICFolder
        .fFlags = FOF_ALLOWUNDO
    End With
    RetVal = SHFileOperation(SHFileOp)
        
    'unzip to the twinBasic folder
    With New cZipArchive
        .OpenArchive zipLocation
        .Extract tbHelperSettings.twinBASICFolder
    End With
    ' ************************** this asks for admin rights, the complete zip isn't decompressed 2-24-25
    ' timing perhaps?
        
    ' check to make sure the twinBASIC folder exists after attempted installation
    If fso.FolderExists(tbHelperSettings.twinBASICFolder) Then
        Form1.ShowStatusMessage("twinBASIC from " & zipLocation & " has been extracted and is ready to use.")
        MsgBox("twinBASIC from " & zipLocation & " has been extracted and is ready to use.", vbInformation, "Completed")
    Else
        MsgBox("There was a problem recreating " & tbHelperSettings.twinBASICFolder & ". The parent folder and the zip file will be opened so that you can finish the process.", vbCritical, "Unable to complete")
            
        ShellExecute(0, "open", zipLocation, vbNullString, vbNullString, 1) ' open the zipfile for the user
        ShellExecute(0, "open", GettBParentFolder, vbNullString, vbNullString, 1) ' open the folder where twinBASIC is supposed to be installed.
            
        MsgBox("Going forward, you can open this utility as administrator to avoid this extra step.")
            
    End If
        
End Sub

Public Function IsCodeRunningInTheIDE() As Boolean
    
    Dim strFileName As String
    Dim lngCount As Long

    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = Left(strFileName, lngCount)
    
    IsCodeRunningInTheIDE = Not InStr(UCase(strFileName), "TWINBASIC_WIN32") = 0
     
End Function

Public Function IsProcessRunning(ByVal ProcessName As String) As Boolean
    
    Dim objWMI As Object, colProcesses As Variant, objProcess As Variant

    ' Get the WMI service object
    Set objWMI = GetObject("winmgmts:\\")

    ' Query for processes
    Set colProcesses = objWMI.ExecQuery("Select * From Win32_Process Where Name='" & ProcessName & "'")

    ' Check if any processes matching the name were found
    If colProcesses.Count > 0 Then
        IsProcessRunning = True
    Else
        IsProcessRunning = False
    End If

    ' Clean up objects
    Set objProcess = Nothing
    Set colProcesses = Nothing
    Set objWMI = Nothing
    
End Function

Public Sub LoadSettingsIntoForm()
    
    'Set the controls on the form with their settings values
    With Form1
        .txtDownloadTo.Text = tbHelperSettings.DownloadFolder
        .txttBLocation.Text = tbHelperSettings.twinBASICFolder
        Select Case tbHelperSettings.PostDownloadAction
            Case 1
                .optOpenFolder.Value = True
            Case 2
                .optOpenZip.Value = True
            Case 3
                .optInstallTB.Value = True
        End Select
        .chkLookForUpdateOnLaunch.Value = tbHelperSettings.CheckForNewVersionOnLoad
        .chkStarttwinBASIC.Value = tbHelperSettings.StarttwinBASICAfterUpdate
        .chkLog.Value = tbHelperSettings.LogActivity
        .chkSaveSettings.Value = tbHelperSettings.SaveSettingsOnExit
    End With
    
End Sub

Public Sub LogGridClick(logGrid As VBFlexGrid)
    
    ' override the forecolor for the selected row with the color for the type of record
    Dim selectCLType As String
    Dim typeColNum As Integer
    
    ' the girds that use this have the type col in different places, find the proper col
    For typeColNum = 0 To logGrid.Cols - 1
        If logGrid.TextMatrix(0, typeColNum) = "Type" Then Exit For
    Next
    
    selectCLType = logGrid.TextMatrix(logGrid.Row, typeColNum)
        
    logGrid.ForeColorSel = ChangeLogItemColor(selectCLType)
    
End Sub

Public Sub WriteToLogFile()
        
    ' write the contents of the displayed log to the log history file
    Dim logFile As TextStream
    Dim logFileName As String = App.Path & "\log.txt"
    Dim logIndex As Integer
    Dim tbVersionInstalled As Boolean = False
    
    Dim lb As ListBox = Form1.lbStatus
    Dim grd As VBFlexGrid = Form1.flgLog
    
    Set logFile = fso.OpenTextFile(logFileName, ForAppending, True)
        For logIndex = 0 To lb.ListCount - 1
            logFile.WriteLine(lb.List(logIndex))
            If Not tbVersionInstalled Then tbVersionInstalled = InStr(lb.List(logIndex), "Post download") > 1 ' if the user at least downloaded the zip
        Next logIndex
        
        ' write the change log(s) for the version downloaded, plus the previous versions inbetween 
        ' the installed and the latest available installed
        If tbVersionInstalled Then
            
            ' these force the var to use a fixed length
            Dim tBVersion As String * 4
            Dim clType As String * 11
            Dim clText As String * 150
            
            For logIndex = 1 To grd.Rows - 1
                tBVersion = grd.TextMatrix(logIndex, 0)
                clType = grd.TextMatrix(logIndex, 1)
                clText = grd.TextMatrix(logIndex, 2)
                
                logFile.WriteLine(Format(Now, "MM/dd/yy hh:mm:ss AM/PM: ") & tBVersion & " - " & clType & ": " & clText)
            Next logIndex
        End If
    logFile.Close()
    Set logFile = Nothing
        
End Sub