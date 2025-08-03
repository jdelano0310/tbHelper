Attribute VB_Name = "modtBHelper"
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long

Private Const WM_MOUSEWHEEL As Long = &H20A

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const GWL_WNDPROC As Long = -4

Private CallbackOwner As Object
Public OriginalCanvasProc As Long

Public ucDictionary As New Scripting.Dictionary   ' my own dictionary to hold window handle to object (user controls) 

Public Sub RegisterScrollableCanvas(ByVal hWnd As Long, ByVal ownerCtrl As Object)
    Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf Canvas_WindowProc)
    ucDictionary.Add hWnd, ownerCtrl
End Sub

Public Const ROW_ALT_COLOR = &HF8F8F8
Public Const CUST_BTN_BCOLOR = &HA2640C


Private Sub UpdateScrollOwnership()
    
    ' which control needs to scroll?
    Dim pt As POINTAPI
    GetCursorPos pt

    Dim hOver As Long
    hOver = WindowFromPoint(pt.X, pt.Y)

    ' the which user control is trying to scroll 
    If ucDictionary.Exists(hOver) Then
        Set CallbackOwner = ucDictionary(hOver)
    End If
    
End Sub

Public Function Canvas_WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    ' this must reside in a bas file and not in the user control for it to be found, it seems
    If uMsg = WM_MOUSEWHEEL Then
        
        ' where is the mouse
        UpdateScrollOwnership
        
        ' Use a callback interface or global reference to your control
        If Not CallbackOwner Is Nothing Then
            CallByName CallbackOwner, "HandleMouseScroll", vbMethod, wParam
        End If

        Exit Function
    End If

    Canvas_WindowProc = CallWindowProc(OriginalCanvasProc, hwnd, uMsg, wParam, lParam)
End Function

Public Sub WriteToDebugFile(logFileLine As String)
    
    Dim logFileName As String = App.Path & "\debug_log.txt"
    Dim fso As New FileSystemObject
    Dim debugLogFile As TextStream = fso.OpenTextFile(logFileName, ForAppending, True)
    
    debugLogFile.WriteLine(Format(Now, "mm/dd/yy hh:MM:ss") & ": " & logFileLine)
    debugLogFile.Close()
    
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
        'Form1.ShowStatusMessage("twinBASIC from " & zipLocation & " has been extracted and is ready to use.")
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

