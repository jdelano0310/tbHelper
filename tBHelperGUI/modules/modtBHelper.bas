Attribute VB_Name = "modtBHelper"
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long

Private Const WM_MOUSEWHEEL As Long = &H20A

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
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
