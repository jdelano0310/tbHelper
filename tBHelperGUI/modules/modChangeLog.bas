Attribute VB_Name = "modChangeLog"
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long

Private Const WM_MOUSEWHEEL As Long = &H20A

Public CallbackOwner As Object
Public OriginalCanvasProc As Long

Public Function Canvas_WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    ' this must reside in a bas file and not in the user control for it to be found, it seems
    If uMsg = WM_MOUSEWHEEL Then
        ' Use a callback interface or global reference to your control
        If Not CallbackOwner Is Nothing Then
            CallByName CallbackOwner, "HandleMouseScroll", vbMethod, wParam
        End If
        'Canvas_WindowProc = 0
        Exit Function
    End If

    Canvas_WindowProc = CallWindowProc(OriginalCanvasProc, hwnd, uMsg, wParam, lParam)
End Function

Public Sub WriteToLogFile(logFileLine As String)
    
    Dim logFileName As String = App.Path & "\debug_log.txt"
    Dim fso As New FileSystemObject
    Dim debugLogFile As TextStream = fso.OpenTextFile(logFileName, ForAppending, True)
    
    debugLogFile.WriteLine(Format(Now, "mm/dd/yy hh:MM:ss") & ": " & logFileLine)
    debugLogFile.Close()
    
End Sub

Public Function PixelsToTwips(pixels As Long) As Long
    PixelsToTwips = pixels * Screen.TwipsPerPixelY
End Function
