Attribute VB_Name = "modtBHelper"
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
    
Public FO_DELETE As Long = &H3
Public FOF_ALLOWUNDO As Long = &H40

Public Const FO_COPY = &H2
Public Const FOF_SILENT = &H4
Public Const FOF_NOCONFIRMATION = &H10

Public Declare Function SHFO_UnZip Lib "shell32" Alias "SHFileOperationW" (ByVal lpFileOp As Long) As Long

Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public BINDF_GETNEWESTVERSION As Long = &H10
    
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public SW_HIDE As Integer = 0

Public Sub ConfigureCustomButton(theButton As ucCustomButton, buttonCaption As String, bkColor As OLE_COLOR, frColor As OLE_COLOR, _
    pngImagePath As String, iconSize As Integer, startEnabled As Boolean, boldFont As Boolean, _
    Optional borderColor As OLE_COLOR = 0, Optional borderWidth As Integer = 0)
    
    'WriteToDebugLogFile "       ConfigureCustomButton identifier " & IIf(buttonCaption = "", pngImagePath, buttonCaption)
    
    With theButton
        .Caption = buttonCaption
        .BackColor = bkColor
        .ForeColor = frColor
        If borderWidth > 0 Then
            .BorderColor = borderColor
            .BorderWidth = borderWidth
        End If
        .FontSize = 11
        .BorderRadius = 3
        .FontBold = boldFont
        .PngIconPath = pngImagePath
        .IconSize = iconSize
        .IconSpacing = 8
        .Enabled = startEnabled
    End With
    
End Sub

Public Sub WriteToDebugLogFile(logFileLine As String)
    
    Dim logFileName As String = App.Path & "\debug_log.txt"
    Static debugLogFile As TextStream
    'Dim debugLogFile As TextStream
    
    ' it this was not used then no reason to close it
    If logFileLine = "CLOSE" And debugLogFile Is Nothing Then Exit Sub

    ' open this once during app run
    On Error GoTo errorHandler
    If debugLogFile Is Nothing Then
        Set debugLogFile = fso.OpenTextFile(logFileName, ForAppending, True)
    End If
    
    'Set debugLogFile = fso.OpenTextFile(logFileName, ForAppending, True)
    
    debugLogFile.WriteLine(Format(Now, "mm/dd/yy hh:MM:ss") & ": " & logFileLine)
    
    ' only close it if this is received - which should only be during form1 unload
    If logFileLine = "CLOSE" Then debugLogFile.Close()
    'debugLogFile.Close()
    
errorHandler:
    ' just skip if there is an error 
    If Err.Number <> 0 Then
        Debug.Print "WriteToDebugLogFile error " & Err.Description & " " & Err.Source
    End If
End Sub

Public Function PixelsToTwips(pixels As Long) As Long
    PixelsToTwips = pixels * Screen.TwipsPerPixelY
End Function

' add your procedures here
Public tbHelperSettings As clsSettings
Public tbHelperClass As clstBHelper
Public fso As FileSystemObject
Public chgLogs As New colChangeLogItems
Public githubReleasesURL As String = "https://github.com/twinbasic/twinbasic/releases"
Public activityLog As ucActivityLog

Public Function GetCurrentTBVersion(tBFolder As String) As String

    'WriteToDebugLogFile("GetCurrentTBVersion " & tBFolder)
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
    'WriteToDebugLogFile("Exit GetCurrentTBVersion")
    
End Function

Public Function fsoFileRead(filePath As String) As String
    
    If Not fso.FileExists(filePath) Then Return "Failed fsoFileRead"
    
    On Error GoTo readError
    
    Dim fso As New Scripting.FileSystemObject
        Dim fileToRead As TextStream
        
        Set fileToRead = fso.OpenTextFile(filePath, ForReading)
            fsoFileRead = fileToRead.ReadAll()
readError:
        If fsoFileRead = vbNullString Then
            MsgBox("Unable to read " & filePath, "error", "FileRead")
        End If
        fileToRead.Close()
    Set fso = Nothing
    
End Function

Public Function GettBParentFolder() As String
        
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

Public Function InstallTwinBasic(tBZipFile As String) As Boolean
        
    ' go through the steps of deleting the current files and unziping the new files
    ' to the folder that has been desgniated
    'WriteToDebugLogFile("           InstallTwinBasic " & tBZipFile & " - start")
    Dim result As Boolean
    
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
        .OpenArchive tBZipFile
        .Extract tbHelperSettings.twinBASICFolder
    End With
    
    'DoEvents()
    ' ************************** this asks for admin rights, the complete zip isn't decompressed 2-24-25
    ' timing perhaps?
        
    ' check to make sure the twinBASIC folder exists after attempted installation
    result = fso.FolderExists(tbHelperSettings.twinBASICFolder)
    
    WriteToDebugLogFile("           InstallTwinBasic - end")
    InstallTwinBasic = result
End Function

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

Public Sub UpdateActivityLog(statMessage As String, Optional updatePreviousStatus As Boolean = False)
    
    WriteToDebugLogFile("In ShowStatusMessage " & statMessage)
    ' write the message to the listbox on the form
    If updatePreviousStatus Then
        activityLog.AddEntry "", statMessage, True
    Else
        activityLog.AddEntry Format(Now, "MM/dd/yy hh:mm:ss AM/PM: "), statMessage
    End If

    WriteToDebugLogFile("Out ShowStatusMessage ")
End Sub

Private Sub CenterPanel(pnlToCenter As Frame, Optional inObject As Object)
    
    ' default the object to center the panel in to Form1 if
    ' none is supplied
    If inObject Is Nothing Then
        Set inObject = Form1
    End If
        
    Dim x As Long, y As Long
    x = (inObject.Width - pnlToCenter.Width) \ 2
    y = (inObject.Height - pnlToCenter.Height) \ 2
    pnlToCenter.Left = x
    pnlToCenter.Top = y
        
End Sub

Public currentPanelTop As Long
Public currentPanelLeft As Long

Public Sub ShowPanelView(innerPanel As Frame, Optional radius As Long = 10)
    
    ' this will be called to display a hidden panel / frames 
    ' use the parent "frame panel" as the drop shadow for the "frame form"
    Dim parentPanel As Frame
    If TypeOf innerPanel Is Frame Then
        Set parentPanel = innerPanel.Container
    End If
    
    currentPanelTop = parentPanel.Top
    currentPanelLeft = parentPanel.Left
    
    ApplyRoundedRegion parentPanel, radius
    parentPanel.BackColor = RGB(180, 180, 180)  ' light gray shadow
    
    innerPanel.BackColor = RGB(240, 240, 240)   ' lighter background
    
    ' ensure the size creates a border around the inner panel
    innerPanel.Left = 100
    innerPanel.Top = 120
        
    parentPanel.Width = innerPanel.Width + 90
    parentPanel.Height = innerPanel.Height + 100
    
    ApplyRoundedRegion innerPanel, 12
    
    CenterPanel parentPanel
    CenterPanel innerPanel, parentPanel
    
    parentPanel.Visible = True
    parentPanel.ZOrder 0

    Form1.isAPanelDisplayed = True
End Sub

Public Sub HidePanelView(parentPanel As Frame)
    
    ' hide the panel and put it back where it was
    Form1.isAPanelDisplayed = False
    parentPanel.Visible = False
        
    ' put the frame back off the screen
    parentPanel.Top = currentPanelTop
    parentPanel.Left = currentPanelLeft
        
    FlushRedraws()
    
End Sub